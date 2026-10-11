# 제작본 보존과 파일 찾기

> 이 문서에서 해야 할 일: 검토할 버전을 먼저 선택한 뒤 같은 행의 PPT·대본·콘티를 함께 연다. 기존 제작본은 보존하고, 후속 개정본은 별도 경로로 추가한다.

**2026-10-11 보존 기준:** 커밋 `86007d110858f7cf9f75c6b026f969495daed5b9`에 저장된 기본본과 월드컵 추가본. 스토리라인 검토를 시작하기 전에 GitHub `main`과 이 커밋이 일치함을 확인했다.

| 버전 | PPT·PDF | 대본·콘티 | 상태 |
| --- | --- | --- | --- |
| **최신 공동 해설 editorial-v0.2** | [PPTX](../deliverables/editorial-v02/learning-fair-2026-editorial-v02.pptx) · [PDF](../deliverables/editorial-v02/learning-fair-2026-editorial-v02.pdf) | [대본](editorial-v02/02-script.md) · [콘티](editorial-v02/01-storyboard.md) | 2026-10-11 추가 사실 반영. 이관우 참여 경험·팀장 관점·모집 방식·자료 삽입 9곳. [검토·제출 일정](editorial-v02/03-review-and-schedule.md). 리허설·실제 자료 삽입 전 |
| **공동 해설 editorial-v0.1** | [PPTX](../deliverables/editorial/learning-fair-2026-editorial.pptx) · [PDF](../deliverables/editorial/learning-fair-2026-editorial.pdf) | [대본](editorial/02-script.md) · [콘티](editorial/01-storyboard.md) | 2026-10-11 별도 생성. ‘그 일, 같이 해볼까요?’ / [비판적 검토 기록](editorial/03-editorial-review.md). 실제 낭독·내부 원본 교체 전 |
| 기존 기본본 v0.1 | [PPTX](../deliverables/learning-fair-2026.pptx) · [PDF](../deliverables/learning-fair-2026.pdf) | [대본](06-script.md) · [콘티](05-storyboard.md) | 2026-10-09 생성된 원본 보존. 후속 서사 검토가 자동 반영된 파일이 아님 |
| 월드컵 추가본 v0.1 | [PPTX](../deliverables/worldcup/learning-fair-2026-worldcup.pptx) · [PDF](../deliverables/worldcup/learning-fair-2026-worldcup.pdf) | [대본](worldcup/02-script.md) · [콘티](worldcup/01-storyboard.md) | 2026-10-10 제작본 보존. 4강·결승과 Tech X 촬영용 선택안 |
| 월드컵 v0.2 서사 제안 | 새 PPT 미생성 | [흐름 검토](worldcup/05-storyline-review.md) · [장면·연결 대사 제안](worldcup/06-storyline-revision-proposal.md) | 2026-10-11 검토 문서. 전체 촬영 대본·콘티·발화 검증 완료본이 아님 |

월드컵 결승의 [선택 전·후 PPTX](../deliverables/worldcup/choice-states.pptx)와 PDF·PNG도 v0.1 자료로 보존한다. 각 버전의 프롬프터·리허설 표·시간 보고서·검수 보고서·미리보기도 해당 `deliverables/` 경로에 함께 남긴다.

공동 해설 v02는 `676aeca6e110887042ab9bf912234bb6cebefa9e`에서 추적하던 기존 제작 파일 104개의 보존을 [별도 기준](../deliverables/editorial-v02/preservation-baseline.json)과 [검증 결과](../deliverables/editorial-v02/validation-report.json)로 확인한다. 기본본·월드컵본·공동 해설 v01을 덮어쓰지 않는다.

## 후속 제작 시 지킬 일

- `content/film.json`과 `content/worldcup-film.json`을 기존 두 제작본의 기준으로 유지한다.
- 새로운 월드컵 제작본을 구현할 때는 예를 들어 `content/worldcup-v02-film.json`, `docs/worldcup-v02/`, `deliverables/worldcup-v02/`처럼 구분된 경로를 사용한다. 이 경로들은 현재 파일이 있다는 뜻이 아니라 다음 제작본의 권장 저장 위치다.
- 기존 빌드 스크립트를 그대로 실행해 보존할 파일을 덮어쓰지 않는다. 새로운 출력 경로와 원본을 먼저 분리한 뒤 생성한다.
- 같은 자료 번호 P01~P10도 버전마다 내용이 다르므로 PPT·대본·콘티를 서로 섞지 않는다.
- 스토리라인 제안의 시간 배정표를 실제 발화 검수 결과로 인용하지 않는다. 새 전체 대사와 장표가 만들어진 뒤 별도 검증한다.
- 사용자 요청에 따라 새 검토 문서와 후속 제작 산출물도 `main`에 커밋·푸시하며 기존 이력을 보존한다.

기존 파일의 내용 보존 여부는 [보존 확인 기록](../deliverables/review/version-preservation-2026-10-11.json)에 남긴다. 이 기록은 기준 커밋과 작업 파일의 바이트 일치를 확인한 것이며, 새 영상이나 새 PPT의 품질 검증 결과가 아니다.
