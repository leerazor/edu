"""Validate the independent balance package after PowerPoint rendering."""

from __future__ import annotations

import copy
import csv
import json
import re

from build_balance import (
    CHOICE_STATES, DOCS, OUT, PACKAGE, ROOT, SOURCE, alternative_lines,
    check_source, prepare_film,
)
from production_common import (
    compact, cue_text, markdown_scene_blocks, measure, name,
    prompter_scene_blocks, prompter_turns, scene_content_hash, timecode,
)
from validate_story_pair import Validation, check_render, sha256


def check_documents(film, validation):
    check = validation.check
    check_source(film)
    check(True, "Source timeline, speakers, cue anchors, and reading budget")
    timing = json.loads((OUT / "timing-report.json").read_text(encoding="utf-8"))
    rows = [measure(scene, film) for scene in film["scenes"]]
    check(timing["scenes"] == rows, "Timing estimates match current source")
    check(timing["version"] == film["version"] and timing["date"] == film["date"], "Timing version and date")
    fixed_seconds = sum(item["duration"] for item in film["fixed_segments"])
    check(timing["main_allocated_seconds"] == 275 and timing["fixed_seconds"] == fixed_seconds == 25,
          "275 studio seconds and 25 inserted video seconds")
    for rate in film["reading_rates"]:
        total = round(sum(row["estimates_seconds"][rate] for row in rows), 2)
        check(timing["main_estimates_seconds"][rate] == total and
              timing["whole_film_estimates_seconds"][rate] == round(total + fixed_seconds, 2),
              f"{rate}: studio and whole estimates")
    script = (DOCS / "02-script.md").read_text(encoding="utf-8")
    board = (DOCS / "01-storyboard.md").read_text(encoding="utf-8")
    prompt = (OUT / "teleprompter.txt").read_text(encoding="utf-8")
    script_blocks = markdown_scene_blocks(script)
    board_blocks = markdown_scene_blocks(board)
    prompt_blocks = prompter_scene_blocks(prompt)
    expected_ids = {scene["id"] for scene in film["scenes"]}
    check(set(script_blocks) == set(board_blocks) == set(prompt_blocks) == expected_ids,
          "Script, storyboard, and prompter cover all current scenes")
    facts = (ROOT / "docs/03-facts-and-assets.md").read_text(encoding="utf-8")
    registered = set(re.findall(r"(?m)^\|\s*(IMG\d{2})\s*\|", facts))
    for filename in ["01-storyboard.md", "02-script.md"]:
        value = (DOCS / filename).read_text(encoding="utf-8")
        check("이 문서에서 해야 할 일" in value, f"{filename}: reader action")
        check(film["display_name"] in value, f"{filename}: current format name")
        check(set(re.findall(r"\bIMG\d{2}\b", value)) <= registered, f"{filename}: known assets")
    with (OUT / "rehearsal.csv").open(encoding="utf-8-sig", newline="") as handle:
        rehearsal = list(csv.DictReader(handle))
    check([row["scene"] for row in rehearsal] == [scene["id"] for scene in film["scenes"]],
          "Rehearsal rows match scene order")
    for scene, row, estimate in zip(film["scenes"], rehearsal, rows, strict=True):
        sid, pid = scene["id"], scene["slide"]
        check(set(scene["assets"]) <= registered, f"{sid}: known source assets")
        check(set(re.findall(r"P\d{2}", scene["cue"]["effect"])) == {pid}, f"{sid}: main cue targets its slide")
        check(scene["camera"] in board_blocks[sid] and scene["hold"] in board_blocks[sid],
              f"{sid}: storyboard camera and hold")
        check(scene["cue"]["anchor"] in board_blocks[sid], f"{sid}: storyboard cue")
        for turn in scene["turns"]:
            expected = f"{name(turn['speaker'], film)}: {cue_text(turn, scene, film)}"
            check(compact(expected) in compact(script_blocks[sid].replace("**", "")), f"{sid}: script speech and cues")
        for slot in scene.get("asset_slots", []):
            check(slot["id"] in registered, f"{sid}: known editable photo slot")
            check(slot["label"].replace("\n", " / ") in script_blocks[sid] and slot["detail"] in board_blocks[sid],
                  f"{sid}: photo instructions in script and board")
        actual_slide, prompt_block = prompt_blocks[sid]
        expected_turns = [(name(turn["speaker"], film), compact(cue_text(turn, scene, film, phonetic=True)))
                          for turn in scene["turns"]]
        check(actual_slide == pid and [(who, compact(value)) for who, value in prompter_turns(prompt_block)] == expected_turns,
              f"{sid}: exact prompter speech with pronunciations")
        check(row["content_hash"] == scene_content_hash(scene, film), f"{sid}: rehearsal hash")
        check(row["PPT"] == pid and row["start"] == timecode(scene["start"]) and
              row["end"] == timecode(scene["start"] + scene["duration"]) and
              float(row["allocated_seconds"]) == scene["duration"], f"{sid}: rehearsal timing")
        check(int(row["syllables"]) == estimate["spoken_syllables"] and
              float(row["estimated_baseline_seconds"]) == estimate["estimates_seconds"]["baseline"], f"{sid}: rehearsal estimate")
        if row["measured_seconds"].strip():
            check(float(row["measured_seconds"]) > 0, f"{sid}: positive measurement")
    for fixed in film["fixed_segments"]:
        start, end = timecode(fixed["start"]), timecode(fixed["start"] + fixed["duration"])
        check(f"{start}–{end}" in board and f"{start}–{end}" in script and f"{start}–{end}" in prompt,
              f"{fixed['id']}: fixed video interval in all documents")
    check(script.index("## SCENE #1") < script.index("00:25–00:40", script.index("## SCENE #1")) < script.index("## SCENE #2"),
          "AI video cue sits between studio opening and first question")
    validation.metrics.update({"studio_seconds": 275, "whole_seconds": film["total_seconds"],
        "spoken_syllables": timing["total_spoken_syllables"], "estimates": timing["whole_film_estimates_seconds"],
        "measured_scenes": sum(bool(row["measured_seconds"].strip()) for row in rehearsal)})


def check_deck(film, deck, validation):
    check = validation.check
    check(len(deck.slides) == len(film["scenes"]) + len(film["appendix"]), "Ten main slides and three appendices")
    for index, slide in enumerate(deck.slides):
        main = index < len(film["scenes"])
        check(slide._element.get("show", "1") == ("1" if main else "0"), f"Slide {index + 1}: visibility")
        notes = slide.notes_slide.notes_text_frame.text
        check(bool(notes.strip()), f"Slide {index + 1}: production notes")
        if not main:
            continue
        scene = film["scenes"][index]
        pid = scene["slide"]
        check(scene["camera"] in notes and scene["hold"] in notes, f"{pid}: camera and hold in notes")
        for turn in scene["turns"]:
            check(cue_text(turn, scene, film) in notes, f"{pid}: exact speech and cue in notes")
        for cue in scene.get("secondary_cues", []):
            check(all(cue[key] in notes for key in ["anchor", "label", "effect"]), f"{pid}: secondary cue notes")
        shapes = {shape.name: shape for shape in slide.shapes}
        for slot in scene.get("asset_slots", []):
            frame, label = shapes.get(f"{pid}-asset-frame-{slot['id']}"), shapes.get(f"{pid}-asset-label-{slot['id']}")
            check(frame is not None and label is not None and label.text == slot["label"], f"{pid}: editable slot {slot['id']}")
            check(all(slot[key] in notes for key in ["id", "label", "detail"]), f"{pid}: photo notes {slot['id']}")
        visible = compact("\n".join(shape.text for shape in slide.shapes if shape.has_text_frame))
        visual = scene["visual"]
        for key in ["eyebrow", "headline", "body", "footer", "bottom", "left", "right", "detail", "supporting_line", "supporting"]:
            if visual.get(key):
                check(compact(visual[key]) in visible, f"{pid}: visible source {key}")
        for item in visual.get("items", []):
            for key in ["label", "value", "detail"]:
                if item.get(key):
                    check(compact(item[key]) in visible, f"{pid}: visible item {key}")


def check_choices(film, validation):
    check = validation.check
    rounds = film["balance"]["rounds"]
    choice_ids = [f"{item['id']}-{state}" for item in rounds for state, _ in CHOICE_STATES]
    deck = check_render(OUT / "choice-states.pptx", OUT / "choice-states.pdf", OUT / "choice-preview", choice_ids, validation, choice=True)
    script = (DOCS / "02-script.md").read_text(encoding="utf-8")
    alternate_timings = {}
    for question in rounds:
        alternate = copy.deepcopy(film)
        scenes = {scene["id"]: scene for scene in alternate["scenes"]}
        patches = question.get("alternative_turns", [])
        check(isinstance(patches, list) and bool(patches), f"{question['id']}: alternative response patches")
        if not isinstance(patches, list):
            continue
        touched = set()
        for patch in patches:
            scene = scenes[patch["scene"]]
            original = scene["turns"][patch["turn_index"]]
            check(patch["speaker"] == original["speaker"], f"{question['id']}: patch preserves intended speaker")
            scene["turns"][patch["turn_index"]] = {"speaker": patch["speaker"], "text": patch["text"]}
            touched.add(scene["id"])
            check(patch["text"] in script, f"{question['id']}: alternative response documented")
        reveal = scenes[question["reveal_scene"]]
        raw = " ".join(turn["text"] for turn in reveal["turns"])
        check(raw.count(question["alternative_reveal_anchor"]) == 1, f"{question['id']}: alternate reveal anchor unique")
        alternate_timings[question["id"]] = {}
        for sid in sorted(touched):
            estimate = measure(scenes[sid], alternate)
            check(estimate["estimates_seconds"]["slow"] <= scenes[sid]["duration"], f"{question['id']} / {sid}: alternate fits scene")
            alternate_timings[question["id"]][sid] = estimate["estimates_seconds"]
    for index, slide in enumerate(deck.slides):
        question = rounds[index // 3]
        state, selected = CHOICE_STATES[index % 3]
        notes = slide.notes_slide.notes_text_frame.text
        visible = "\n".join(shape.text for shape in slide.shapes if shape.has_text_frame)
        check(f"{question['id']} / {state}" in notes, f"{choice_ids[index]}: notes identify state")
        expected_anchor = (question["alternative_reveal_anchor"]
                           if selected and selected != question["selected"] else question["reveal_anchor"])
        check(f"이 상태의 공개 발화: {expected_anchor}" in notes, f"{choice_ids[index]}: matching reveal anchor")
        check(all(choice["name"] in visible for choice in question["choices"]), f"{choice_ids[index]}: both choices visible")
        check("추가 장면이나 시간이 아니다" in notes, f"{choice_ids[index]}: no extra timeline allocation")
        for choice in question["choices"]:
            for line in alternative_lines(question, film, choice["id"]):
                check(line in notes, f"{choice_ids[index]}: branch response notes")
        if selected:
            selected_name = next(choice["name"] for choice in question["choices"] if choice["id"] == selected)
            check(f"지금의 선택  {selected} · {selected_name}" in visible, f"{choice_ids[index]}: selected choice disclosed")
        else:
            check("지금의 선택" not in visible, f"{choice_ids[index]}: unrevealed before choice")
    for index, question in enumerate(rounds):
        objects = [b"".join(shape._element.xml.encode("utf-8") for shape in deck.slides[index * 3 + offset].shapes)
                   for offset in range(3)]
        check(len(set(objects)) == 3, f"{question['id']}: PRE and both alternatives differ visibly")
    validation.metrics["alternative_scene_estimates"] = alternate_timings
    validation.metrics["independent_questions"] = len(rounds)
    validation.metrics["choice_states"] = len(choice_ids)


def check_facts(film, validation):
    check = validation.check
    check(film["official_expansion"] == "Advanced Infrastructure for Artificial Intelligence", "Official AI 4 AI expansion")
    check(set(film["presenters"].values()) == {"류경동 팀장", "이관우 교육담당"}, "Confirmed presenters")
    check("승자" in film["balance"]["disclaimer"] and "실제 선택" in film["balance"]["disclaimer"],
          "Independent choices distinguish proposal from real choice")
    for scene in film["scenes"]:
        if scene["visual"]["layout"] == "experience":
            check({slot["id"] for slot in scene["asset_slots"]} == {"IMG16", "IMG17"}, "Personal experience uses its two actual photo slots")
    schedule = film.get("schedule", {})
    check(schedule.get("as_of") == "2026-10-11", "Fact status dated October 11")
    for key, date in {"outsystems": "2026-10-12", "submission": "2026-10-14", "tech_x_final": "2026-10-16", "filming": "2026-10-19"}.items():
        check(schedule.get(key) == {"date": date, "status": "scheduled"}, f"{key}: date and scheduled status")


def main():
    validation = Validation(PACKAGE)
    film = validation.section("Source JSON", lambda: prepare_film(json.loads(SOURCE.read_text(encoding="utf-8"))))
    if film is not None:
        validation.section("Confirmed facts", lambda: check_facts(film, validation))
        validation.section("Time and documents", lambda: check_documents(film, validation))
        page_ids = [scene["slide"] for scene in film["scenes"]] + [item["id"] for item in film["appendix"]]
        deck = validation.section("PowerPoint render", lambda: check_render(
            OUT / f"learning-fair-2026-{PACKAGE}.pptx", OUT / f"learning-fair-2026-{PACKAGE}.pdf",
            OUT / "preview", page_ids, validation))
        if deck is not None:
            validation.section("Deck source synchronization", lambda: check_deck(film, deck, validation))
        validation.section("Choice branches and rendered states", lambda: check_choices(film, validation))
    paths = [SOURCE, DOCS / "01-storyboard.md", DOCS / "02-script.md", OUT / "teleprompter.txt", OUT / "timing-report.json",
             OUT / f"learning-fair-2026-{PACKAGE}.pptx", OUT / f"learning-fair-2026-{PACKAGE}.pdf", OUT / "choice-states.pptx", OUT / "choice-states.pdf"]
    report = {"status": "FAIL" if validation.errors else "PASS", "passed_checks": len(validation.passed),
              "errors": validation.errors, "metrics": validation.metrics,
              "artifact_sha256": {str(path.relative_to(ROOT)): sha256(path) for path in paths if path.is_file()},
              "limits": ["Speech durations are estimates, not presenter rehearsal measurements.",
                         "Internal photo slots await user-supplied originals and company visual assets.",
                         "Contact sheets require visual inspection; no actual PowerPoint animation is embedded.",
                         "Alternative replies are timed proposals and must be synchronized after actual selection."]}
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "validation-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))
    if validation.errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
