"""Shared timing, production documents, source checks, and text parsers.

This module writes only the paths explicitly passed by the pair builder.
"""

from __future__ import annotations

import csv
import hashlib
import json
import re

def timecode(seconds: int) -> str:
    return f"{seconds // 60:02d}:{seconds % 60:02d}"


def spoken(text: str, film: dict) -> str:
    for word, reading in sorted(film["pronunciations"].items(), key=lambda pair: -len(pair[0])):
        text = text.replace(word, reading)
    return text


def measure(scene: dict, film: dict) -> dict:
    raw = " ".join(turn["text"] for turn in scene["turns"])
    read = spoken(raw, film)
    # Hangul blocks are counted after explicit pronunciation expansion. Punctuation
    # is handled by the per-scene pause allowance; these are planning estimates.
    count = len(re.findall(r"[가-힣]", read))
    assert not re.search(r"[A-Za-z0-9]", read), f"Unmapped pronunciation: {scene['id']}"
    times = {key: round(count / rate + scene["pause_seconds"], 2)
             for key, rate in film["reading_rates"].items()}
    return {
        "id": scene["id"], "slide": scene["slide"], "raw_characters": len(raw),
        "spoken_syllables": count, "pause_seconds": scene["pause_seconds"],
        "allocated_seconds": scene["duration"], "estimates_seconds": times,
        "baseline_visual_hold_seconds": round(scene["duration"] - times["baseline"], 2),
    }


def name(role: str, film: dict) -> str:
    return film["presenters"].get(role, role)


def cue_text(turn: dict, scene: dict, film: dict, phonetic: bool = False) -> str:
    value = turn["text"]
    cues = [(scene["cue"]["anchor"], f"{scene['slide']} 등장")]
    cues += [(cue["anchor"], cue["label"]) for cue in scene.get("secondary_cues", [])]
    for anchor, label in cues:
        if anchor in value:
            value = value.replace(anchor, f"[{label}] {anchor}", 1)
    if not phonetic:
        return value
    # Keep production cues unchanged; phonetic expansion applies only to speech.
    return "".join(piece if piece.startswith("[") else spoken(piece, film)
                   for piece in re.split(r"(\[[^\]]*\])", value))


def add_notes(slide, scene, film):
    lines = [
        f"{scene['slide']} / {scene['id']} / 최종 {timecode(scene['start'])}–{timecode(scene['start'] + scene['duration'])}",
        f"원본 버전: {film['version']} / {film['date']}",
        "자료 큐는 발화 기준. 타임코드는 최종 편집 목표이며 실측 발화 시점이 아니다.",
        f"자료 등장: {scene['cue']['anchor']}", scene["cue"]["effect"],
        f"촬영: {scene['camera']}",
        f"원본: {', '.join(scene['assets']) or '편집 가능한 키워드·도식'}",
        "스크립트:",
    ]
    lines += [f"{name(turn['speaker'], film)}: {cue_text(turn, scene, film)}"
              for turn in scene["turns"]]
    lines += [f"추가 큐: {cue['anchor']} / {cue['label']} / {cue['effect']}"
              for cue in scene.get("secondary_cues", [])]
    lines += [scene["hold"], film.get("source_basis", ""), "PPTX에는 실제 애니메이션이 없다. 효과와 자료 전환은 영상팀 편집 지시다.",
              "내부 행사 사진·화면 원본은 현재 미수급이다. 범용 외부 사진으로 실제 행사를 대체하지 않는다.",
              "출처: 사용자 브리프·추가 확인. 사실과 원본 ID 기준: docs/03-facts-and-assets.md."]
    slide.notes_slide.notes_text_frame.text = "\n".join(lines)


def scene_content_hash(scene, film):
    """A measurement belongs to this speech and cue configuration only."""
    payload = {key: scene.get(key) for key in
               ["id", "slide", "turns", "cue", "secondary_cues", "duration", "pause_seconds"]}
    payload["pronunciations"] = film["pronunciations"]
    payload["presenters"] = film["presenters"]
    value = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def cell(value):
    return str(value).replace("|", "\\|").replace("\n", "<br>")


def make_documents(film, out, docs, source_path):
    label = film.get("display_name", "흐름형" if film["variant"] == "story" else "월드컵형")
    rows = [measure(scene, film) for scene in film["scenes"]]
    report = {
        "basis": "Planning estimate, not a rehearsal measurement",
        "version": film["version"], "date": film["date"],
        "reading_rates_syllables_per_second": film["reading_rates"],
        "main_allocated_seconds": sum(s["duration"] for s in film["scenes"]),
        "fixed_seconds": sum(s["duration"] for s in film["fixed_segments"]),
        "scenes": rows,
        "total_spoken_syllables": sum(r["spoken_syllables"] for r in rows),
        "main_estimates_seconds": {key: round(sum(r["estimates_seconds"][key] for r in rows), 2)
                                   for key in film["reading_rates"]},
    }
    report["whole_film_estimates_seconds"] = {
        key: round(value + report["fixed_seconds"], 2)
        for key, value in report["main_estimates_seconds"].items()
    }
    report["iap_allocated_seconds"] = sum(
        scene["duration"] for scene in film["scenes"]
        if re.search(r"(?<![A-Za-z])IAP(?![A-Za-z])|Infra Agent Platform",
                     " ".join(turn["text"] for turn in scene["turns"]))
    )
    (out / "timing-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    heading = f"**{film['version']} · {film['date']}** · {film['title']} — {film['subtitle']}"
    board = [f"# 초안 기반 {label} 영상 콘티", "",
             "> 이 문서에서 해야 할 일: 장면별 자료 등장 발화·화면·촬영·홀드를 확인하고, 같은 발화 큐에 맞춰 편집한다.", "", heading, "",
             film.get("timeline_summary", "전체 300초 = 회사 인트로 5초 + AI 팀 소개 15초 + 스튜디오 본편 275초 + 회사 아웃트로 5초."), "",
             "| 최종 타임코드 | 길이 | 장면 | 자료 | 화면 목적 |", "| --- | ---: | --- | --- | --- |"]
    timeline = sorted(film["fixed_segments"] + film["scenes"], key=lambda scene: scene["start"])
    for scene in timeline:
        tc = f"{timecode(scene['start'])}–{timecode(scene['start'] + scene['duration'])}"
        board.append(f"| {tc} | {scene['duration']}초 | {scene['id']} | {scene.get('slide', '별도 영상')} | {cell(scene.get('title', scene.get('name', '')))} |")
    script = [f"# 초안 기반 {label} 영상 촬영 대본", "",
              f"> 이 문서에서 해야 할 일: 두 출연자가 발화와 자료 큐를 리허설하고, 장면별 실측을 deliverables/{out.name}/rehearsal.csv에 기록한다. 대괄호와 촬영 지시는 읽지 않는다.", "", heading, "",
              "출연: " + " · ".join(name(role, film) for role in film["presenters"] if role != "전체"), "",
              film.get("timeline_note", "회사 인트로·AI 팀 소개·아웃트로에는 스튜디오 발화가 없다. 본편은 00:20에 시작한다. 자료 전환·화자 교대·시선·호흡·홀드는 장면 배정 시간에 포함한다."), ""]
    prompt = [f"{film['title']} — {film['subtitle']}", f"{film['version']} / {film['date']}",
              "발음 표기 대본 / 대괄호는 읽지 않는 자료 큐", ""]
    board += ["", "## 장면별 편집 지시", ""]
    for scene, row in zip(film["scenes"], rows, strict=True):
        if film.get("display_name"):
            previous_start = next((s["start"] for s in reversed(film["scenes"])
                                   if s["start"] < scene["start"]), -1)
            for fixed in film["fixed_segments"]:
                if previous_start < fixed["start"] < scene["start"]:
                    stamp = f"{timecode(fixed['start'])}–{timecode(fixed['start'] + fixed['duration'])}"
                    instruction = f"{fixed['name']} · {stamp} · {fixed['duration']}초 · 스튜디오 발화 없음"
                    script += [f"### {fixed['id']} · 별도 영상", "", instruction, ""]
                    prompt += [f"[별도 영상 / {instruction}]", ""]
        tc = f"{timecode(scene['start'])}–{timecode(scene['start'] + scene['duration'])}"
        speaker = " → ".join(name(role, film) for role in scene["speaker"].split(" → "))
        board += [f"### {scene['id']} · {scene['title']}", "",
                  f"- 자료: {scene['slide']} / {int(scene['slide'][1:])}쪽.",
                  f"- 등장 발화: **‘{scene['cue']['anchor']}’**.",
                  f"- 화면·효과: {scene['cue']['effect']}", f"- 촬영: {scene['camera']}",
                  f"- 원본: {', '.join(scene['assets']) or '키워드·도식'}.",
                  f"- 편집·홀드: {scene['hold']}", ""]
        script += [f"## SCENE #{int(scene['id'][1:])} · {scene['id']} / {scene['slide']}", "",
                   f"- **장면:** {scene['chapter']} · {tc} · {scene['duration']}초",
                   f"- **발화자:** {speaker}", f"- **자료:** PPTX {int(scene['slide'][1:])}쪽 ({scene['slide']})",
                   f"- **발화 추정:** {row['spoken_syllables']}음절 · 기준 {row['estimates_seconds']['baseline']:.1f}초 · 느림 {row['estimates_seconds']['slow']:.1f}초. 호흡 포함.", "",
                   "**스크립트**", ""]
        prompt += [f"[{scene['id']} / {scene['slide']} / {tc} / {scene['duration']}초]"]
        for turn in scene["turns"]:
            script += [f"**{name(turn['speaker'], film)}:** {cue_text(turn, scene, film)}", ""]
            prompt += [f"<{name(turn['speaker'], film)}>", cue_text(turn, scene, film, phonetic=True), ""]
        script += [f"- **장면 효과:** {scene['cue']['effect']}", f"- **촬영 지시:** {scene['camera']}",
                   f"- **참고 원본:** {', '.join(scene['assets']) or '키워드·도식'}",
                   f"- **편집·확인:** {scene['hold']}", ""]
        for cue in scene.get("secondary_cues", []):
            instruction = f"‘{cue['anchor']}’ → [{cue['label']}] {cue['effect']}"
            board += [f"- 추가 자료 큐: {instruction}", ""]
            script += [f"- **추가 자료 큐:** {instruction}", ""]
    if film.get("display_name"):
        last_end = film["scenes"][-1]["start"] + film["scenes"][-1]["duration"]
        for fixed in film["fixed_segments"]:
            if fixed["start"] >= last_end:
                stamp = f"{timecode(fixed['start'])}–{timecode(fixed['start'] + fixed['duration'])}"
                instruction = f"{fixed['name']} · {stamp} · {fixed['duration']}초 · 스튜디오 발화 없음"
                script += [f"### {fixed['id']} · 별도 영상", "", instruction, ""]
                prompt += [f"[별도 영상 / {instruction}]", ""]
    shared = ["## 제작과 확인", "", film.get("source_basis", ""), "",
              "현재 내부 행사 사진·화면 원본은 미수급이다. 원본이 없으면 JSON에 정의된 키워드·도식과 스튜디오 화면으로 제작한다. 범용 외부 사진을 실제 사내 시설·행사 증거로 사용하지 않는다. 원본 수집과 사실 확인은 docs/03-facts-and-assets.md를 따른다.", "",
              "PPTX에는 실제 애니메이션이 없다. 자료 전환과 모션은 영상팀 편집 지시이며 각 장면 시간 안에서 구현한다. 숨김 부록 A01~A03은 본편 사용 제외지만 검토 PDF·PNG에는 포함된다.", "",
              "프로그램 사이의 연결은 교육문화의 관점을 설명한다. Job Fair가 Tech X를 발생시켰거나 세미나가 IAP 개발의 직접 원인이었다는 인과를 표현하지 않는다. 10월 세미나 상태는 개최 확인 후 갱신한다.", "",
              f"`{source_path}`를 수정한 뒤 `scripts/build_story_pair.py`로 다시 생성한다. story-v02와 worldcup-v02의 원본·산출물을 각각 유지하며, 사용자가 실제 자료를 삽입한 별도 PPTX는 덮어쓰지 않는다. 전달 지시는 `docs/{out.name}/04-handoff.md`를 따른다.", ""]
    if "production_notes" in film:
        shared = ["## 제작과 확인", "", film.get("source_basis", ""), ""]
        for note in film["production_notes"]:
            shared += [note, ""]
        shared += [f"`{source_path}`를 수정한 뒤 `{film['build_command']}`로 다시 생성한다. 전달 지시는 `docs/{out.name}/04-handoff.md`를 따른다.", ""]
    board += shared
    rates = " / ".join(f"{key} {value}" for key, value in film["reading_rates"].items())
    script += ["## 발화량과 리허설", "",
               f"총 {report['total_spoken_syllables']}음절. 본편 추정: 느림 {report['main_estimates_seconds']['slow']:.1f}초 / 기준 {report['main_estimates_seconds']['baseline']:.1f}초 / 빠름 {report['main_estimates_seconds']['fast']:.1f}초. 고정 영상 포함 느린 추정 {report['whole_film_estimates_seconds']['slow']:.1f}초.", "",
               f"속도 가정은 초당 음절 수 {rates}다. 영문·숫자를 발음으로 치환하며 화자명·대괄호 큐는 제외한다. 장면별 pause_seconds는 호흡 예산이다. 이 값은 제작 계획 추정이며 실제 리허설 속도나 실측을 대체하지 않는다.", "",
               film.get("rehearsal_note", "리허설 CSV의 measured_seconds에 장면별 실측을 기록한다. 대사·자료 큐·화자·발음·배정 시간이 같을 때만 이전 실측과 메모를 보존한다. 내용 해시가 바뀌면 해당 실측 칸을 비워 다시 측정한다. IAP는 20초 이내의 짧은 대표 사례다. 초과하면 기능 해설을 늘리지 않고 중복 문구를 줄인다."), "",
               "10분 촬영 발화 상한은 사용 가능한 한 테이크와 고정 영상 기준이다. NG·준비·추가 테이크의 총 촬영 시간은 별도다.", ""] + shared
    (docs / "01-storyboard.md").write_text("\n".join(board).rstrip() + "\n", encoding="utf-8")
    (docs / "02-script.md").write_text("\n".join(script).rstrip() + "\n", encoding="utf-8")
    (out / "teleprompter.txt").write_text("\n".join(prompt).rstrip() + "\n", encoding="utf-8")
    rehearsal = out / "rehearsal.csv"
    saved = {}
    if rehearsal.is_file():
        with rehearsal.open(encoding="utf-8-sig", newline="") as handle:
            saved = {row["scene"]: row for row in csv.DictReader(handle)}
    with rehearsal.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, lineterminator="\n")
        writer.writerow(["scene", "PPT", "start", "end", "allocated_seconds", "speaker", "syllables", "estimated_baseline_seconds", "measured_seconds", "notes", "content_hash"])
        for scene, row in zip(film["scenes"], rows, strict=True):
            fingerprint = scene_content_hash(scene, film)
            old = saved.get(scene["id"], {})
            if old.get("content_hash") != fingerprint:
                old = {}
            writer.writerow([scene["id"], scene["slide"], timecode(scene["start"]),
                             timecode(scene["start"] + scene["duration"]), scene["duration"], scene["speaker"],
                             row["spoken_syllables"], row["estimates_seconds"]["baseline"],
                             old.get("measured_seconds", ""), old.get("notes", ""), fingerprint])
    return report


def check_source(film):
    assert film["total_seconds"] == 300
    assert [s["id"] for s in film["scenes"]] == [f"S{i:02d}" for i in range(1, 11)]
    assert [s["slide"] for s in film["scenes"]] == [f"P{i:02d}" for i in range(1, 11)]
    assert [s["id"] for s in film["appendix"]] == [f"A{i:02d}" for i in range(1, 4)]
    assert sum(s["duration"] for s in film["scenes"]) == 275
    if film["variant"] in {"story", "worldcup"}:
        assert sorted((s["start"], s["duration"]) for s in film["fixed_segments"]) == [(0, 5), (5, 15), (295, 5)]
    assert set(film["reading_rates"]) == {"slow", "baseline", "fast"}
    assert 0 < film["reading_rates"]["slow"] <= film["reading_rates"]["baseline"] <= film["reading_rates"]["fast"]
    previous = 0
    for segment in sorted(film["fixed_segments"] + film["scenes"], key=lambda item: item["start"]):
        assert segment["start"] == previous and segment["duration"] > 0, segment["id"]
        previous += segment["duration"]
    assert previous == film["total_seconds"]
    assert [s["start"] for s in film["scenes"]] == sorted(s["start"] for s in film["scenes"])
    for scene in film["scenes"]:
        assert scene["pause_seconds"] >= 0, scene["id"]
        raw = " ".join(turn["text"] for turn in scene["turns"])
        for turn in scene["turns"]:
            assert turn["speaker"] in film["presenters"], scene["id"]
        for cue in [scene["cue"], *scene.get("secondary_cues", [])]:
            assert cue["anchor"] and raw.count(cue["anchor"]) == 1, scene["id"]
        assert measure(scene, film)["estimates_seconds"]["slow"] <= scene["duration"], scene["id"]


def compact(value: str) -> str:
    return re.sub(r"\s+", "", value)


def markdown_scene_blocks(value):
    headings = list(re.finditer(r"(?m)^#{2,3}[^\n]*\b(S\d{2})\b[^\n]*$", value))
    return {match.group(1): value[match.end():headings[i + 1].start() if i + 1 < len(headings) else len(value)]
            for i, match in enumerate(headings)}


def prompter_scene_blocks(value):
    headings = list(re.finditer(r"(?m)^\[(S\d{2})\s*/\s*(P\d{2})\s*/[^\n]*\]\s*$", value))
    return {match.group(1): (match.group(2), value[match.end():headings[i + 1].start() if i + 1 < len(headings) else len(value)])
            for i, match in enumerate(headings)}


def prompter_turns(value):
    value = re.sub(r"(?m)^\[별도 영상 /[^\n]*\]\s*$", "", value)
    headings = list(re.finditer(r"(?m)^<([^>\n]+)>\s*$", value))
    return [(match.group(1), value[match.end():headings[i + 1].start() if i + 1 < len(headings) else len(value)].strip())
            for i, match in enumerate(headings)]
