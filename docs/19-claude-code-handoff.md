# Claude Code 인계 기록

> 이 문서에서 해야 할 일: [Claude Code 지침](../CLAUDE.md)과 [AGENTS.md](../AGENTS.md)를 먼저 읽고, 아래 최신 파일과 검증 상태를 확인한다. 사용자의 다음 요청에 맞춰 남은 제작 준비를 이어간다. 이전 제작본을 새로 덮어쓰지 않는다.

작성 시각: **2026-10-11 09:30:31 UTC**. 작업 폴더: `/home/kwl/projects/edu`. 브랜치: `main`. 원격: `https://github.com/leerazor/edu`. 최신 제작 패키지 기준 커밋은 **`dea5ccc690218262ea0c2f635ec1aae5437dfaae`**다. 이 값은 인계 문서를 추가하기 전 제작 상태이며, 현재 커밋은 `git log -1 --oneline`으로 확인한다.

## 완료한 것과 현재 판단

2026 Learning Fair 최종 5분 영상을 위한 자료·대본·콘티를 제작했다. 기본본·월드컵 초안·공동 해설 다섯 개정본을 보존하고 모두 비교했다. 사용자는 **원래 `deliverables/learning-fair-2026.pptx`가 가장 좋다**고 평가했다. 이후 그 초안의 시각 구조와 이야기 흐름을 바탕으로, 사용자가 지정한 새 순서의 **흐름형 story-v0.2 / 월드컵형 worldcup-v0.2**를 별도로 완성했다.

현재 추천은 흐름형이다. 월드컵은 친숙한 선택 형식이지만 상호 보완적인 문화를 승자·탈락자로 나누는 한계가 남는다. **최종 제출 형식은 사용자 미선택 상태**다. 두 안의 PPT·PDF·콘티·대본·프롬프터·시간표·검증·미리보기와 월드컵 선택 공개 화면까지 저장돼 있다. 영상 촬영·편집과 최종 영상 제작은 아직 진행 전이다.

최신 순서는 **팀 역할과 AI 4 AI 비전·문화 → 외부 전문가 세미나와 녹화 공유 → Tech X와 이관우의 실제 GPU 참여 경험 → 짧은 IAP → Job Fair·DCIS의 업무 이해 → 전체 문화와 미션 회수**다. IAP는 두 안 모두 **18초**이며 개인 GPU 과제와 다른 프로젝트다. 오래된 버전의 35·40·73초를 최신 제작에 적용하지 않는다.

## 바로 열 파일

| 용도 | 흐름형 · 현재 추천 | 월드컵형 · 비교 대안 |
| --- | --- | --- |
| 원본 | [story-v02-film.json](../content/story-v02-film.json) | [worldcup-v02-film.json](../content/worldcup-v02-film.json) |
| 편집 PPTX | [흐름형 PPT](../deliverables/story-v02/learning-fair-2026-story-v02.pptx) | [월드컵 PPT](../deliverables/worldcup-v02/learning-fair-2026-worldcup-v02.pptx) |
| 실제 PowerPoint PDF | [흐름형 PDF](../deliverables/story-v02/learning-fair-2026-story-v02.pdf) | [월드컵 PDF](../deliverables/worldcup-v02/learning-fair-2026-worldcup-v02.pdf) |
| 대본 / 콘티 | [대본](story-v02/02-script.md) / [콘티](story-v02/01-storyboard.md) | [대본](worldcup-v02/02-script.md) / [콘티](worldcup-v02/01-storyboard.md) |
| 비판적 검토 | [흐름 검토](story-v02/03-story-review.md) | [선택·흐름 검토](worldcup-v02/03-story-review.md) |
| 자료 삽입·촬영 큐 | [전달 안내](story-v02/04-handoff.md) | [전달·공개 안내](worldcup-v02/04-handoff.md) |
| 전체 화면 | [미리보기](../deliverables/story-v02/preview/contact-sheet.jpg) | [미리보기](../deliverables/worldcup-v02/preview/contact-sheet.jpg) |
| 리허설 입력 | [CSV](../deliverables/story-v02/rehearsal.csv) | [CSV](../deliverables/worldcup-v02/rehearsal.csv) |

월드컵 선택 전·후 6장은 [choice-states.pptx](../deliverables/worldcup-v02/choice-states.pptx)와 [PDF](../deliverables/worldcup-v02/choice-states.pdf)다. 기준은 세 경기 모두 ‘다른 조직에도 권하고 싶은 교류 방식은?’이며 제안 승자는 Tech X·DCIS·Tech X다. 실제 출연자의 선택은 아직 확인되지 않았다. 공개 화면은 본편 시간 안에서 교체한다.

전체 경로는 [버전 목록](14-version-index.md), 초안을 기준으로 돌아온 이유는 [기존 7종 비교](18-full-version-review.md)에 있다. `docs/15-*`와 `16-*`의 방송·유튜브 검토, `17-*`의 당시 추천은 연구 기록이다. `docs/worldcup/06-storyline-revision-proposal.md`는 예전 미적용 제안이며 최신 생성본 원본이 아니다.

## 확인된 검증과 한계

| 확인 항목 | 흐름형 | 월드컵형 |
| --- | --- | --- |
| 저장된 검증 보고서 | [PASS · 1,471개](../deliverables/story-v02/validation-report.json) | [PASS · 1,928개](../deliverables/worldcup-v02/validation-report.json) |
| 본편 / 고정 영상 / 전체 배정 | 275 / 25 / 300초 | 275 / 25 / 300초 |
| 발음 환산 대본 | 917음절 | 960음절 |
| 느린 발화·호흡·고정 영상 추정 | 262.34초 | 271.56초 |
| 실측된 리허설 장면 | 0 | 0 |
| 실제 PowerPoint 렌더 | 본편·숨김 부록 13장 PDF·PNG | 13장과 선택 전후 6장 PDF·PNG |
| 이전 제작 파일 보존 | 224개 | 224개 |

제작 단계에서 실제 PowerPoint 렌더와 미리보기를 검토했고, 줄바꿈·선택 화면 식별자 등을 고쳤다. **이번 인계에서는 보고서의 기록을 읽고, 보고서에 기록된 원본·산출물 해시 7개와 9개가 현재 파일과 모두 일치함을 확인했다.** 문서 인계만을 위해 PPT를 다시 생성하거나 렌더하지 않았다. PASS는 내부 사진 수급·회사 KV 반영·출연자 실제 낭독·최종 영상 완료를 뜻하지 않는다.

## 남은 작업과 필요한 입력

1. **최종 형식과 출연자 말투 확인:** 사용자의 다음 피드백을 최신 두 안에 적용한다. 흐름형 추천을 사용자 확정으로 기록하지 않는다. 월드컵을 택하면 제안 선택과 이유를 출연자가 받아들일 수 있는지 함께 확인한다.
2. **사용자가 붙일 실제 자료:** 개인 사진 **IMG16·IMG17은 아직 없다.** 흐름형 P05 / 월드컵형 P04에 넣는다. 대표 세미나 포스터·녹화 플랫폼·Tech X 모집 공고·Job Fair·DCIS·IAP 실제 화면의 요청 내용은 각 안의 전달 안내에 있다. 사진 삽입은 별도 제출용 PPTX 사본에 하고, 그 사본의 PDF를 다시 내보낸다.
3. **행사 진행 사실 갱신:** 10월 11일 기준 OutSystems 12일 예정, 제출 14일, Tech X 발표 16일 예정, 촬영 19일이다. 개최·발표·수상 결과는 사용자 확인 후 반영한다. 제출 이후 변경은 장면별 대사·자막·자료 큐와 함께 기록한다.
4. **리허설:** CSV의 `measured_seconds`에 장면별 실제 낭독 시간을 기록하고 자료를 볼 호흡·전환까지 점검한다. 대사·큐·발음·시간이 바뀌면 이전 실측의 내용 해시가 맞는지 확인한다. 추정 잔여 시간은 영상팀과 사진·도식·인물 화면의 유지 시간으로 배분한다.
5. **영상팀 전달:** 회사 KV·폰트·로고·오프닝·클로징 원본은 미수급이다. 별도 AI 팀 소개 15초는 [제작 브리프](07-ai-intro-brief.md)만 있으며 영상 파일은 없다. 사용자가 요청하는 제작 범위에 맞춰 자료 묶음·촬영본·편집 큐를 전달 준비한다. PPT 모션은 편집 지시이고 현재 애니메이션 구현은 없다.

새로운 내부 성과 수치·행사 완료·사진 내용을 추정해서 채우지 않는다. 사용자에게서 추가 입력이 오기 전에도 최신 PPT와 대본의 자연스러운 발화·전환 이유·결말 회수는 검토할 수 있다. 새 피드백이 없다면 이미 완성한 두 패키지를 이유 없이 재작성하지 않는다.

## 작업 재개와 Git

먼저 `git status --short --branch`와 `git log -1 --oneline`을 확인한다. 인계 시작 시 사용자 작업 여부를 알 수 없는 미추적 **`.serena/`**만 있었다. 이 폴더는 커밋 대상에서 제외하며 삭제하지 않는다. 진행 중인 빌드·렌더 프로세스나 후속 자동 실행은 없다.

수정할 원본과 출력 경로를 정한 뒤 생성·렌더·검증 순서로 작업한다. 구체적인 명령과 보존 원칙은 [CLAUDE.md](../CLAUDE.md)에 있다. 대사·콘티만 손으로 수정하지 않는다. 이후 새 내용 개정은 별도 버전으로 보존한다. 사용자는 작업 산출물·Markdown의 **main 커밋·푸시**를 명시했으며, 강제 푸시는 허용되지 않는다. 이번 인계 변경은 지침 문서와 안내 링크이며 제작 파일은 그대로다.

다음 세션에 아래를 복사해 시작할 수 있다.

```text
CLAUDE.md, AGENTS.md, docs/19-claude-code-handoff.md를 읽고 Learning Fair 작업을 이어가줘.
먼저 최신 story-v02와 worldcup-v02의 PPT·대본·콘티·검토 기록을 확인하고,
내 다음 요청을 그 기준으로 반영해줘. 초안과 모든 이전 제작본을 보존하고,
IAP는 대표 사례로 짧게 유지하며, 미확인 행사 결과와 성과 수치는 만들지 마.
아이디어 제시에서 끝내지 말고 수정한 자료와 대본을 비판적으로 검토한 뒤,
검증한 변경사항을 main에 커밋·푸시해줘.
```
