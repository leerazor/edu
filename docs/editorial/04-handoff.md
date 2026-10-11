# 공동 해설 개정본 전달

> 이 문서에서 해야 할 일: PPT·대본·콘티를 같은 버전으로 열고, 출연자 리허설과 실제 원본 교체 후 영상팀에 전달한다. 기본본·월드컵본과 장면 ID를 섞지 않는다.

## 열어볼 파일

- [PPTX 편집 원본](../../deliverables/editorial/learning-fair-2026-editorial.pptx)
- [PowerPoint 검토 PDF](../../deliverables/editorial/learning-fair-2026-editorial.pdf)
- [콘티](01-storyboard.md) / [대본](02-script.md)
- [발음 표기 프롬프터](../../deliverables/editorial/teleprompter.txt)
- [실측 입력 CSV](../../deliverables/editorial/rehearsal.csv)
- [비판적 검토·변경 기록](03-editorial-review.md)
- [검증 결과](../../deliverables/editorial/validation-report.json)

본편 10장, 숨김 제작 부록 3장이다. 검토 PDF·PNG는 부록까지 13장 포함한다. IAP는 P06/S06의 35초 한 장면이다. 총 300초 중 스튜디오 본편은 275초다.

## 원본을 받으면 바꿀 화면

| 우선순위 | 화면 | 필요한 원본 | 지금의 대체 화면 |
| --- | --- | --- | --- |
| 1 | P02 | Job Fair 업무 소개 판넬 또는 부스 사진 IMG04·05 | 업무 소개 방식과 11개 부서 |
| 1 | P04·05 | Tech X 과제 목록·모집/선발 공지·참여 집계 IMG07 | 참여 방식과 6개·44명 |
| 1 | P06 | 실제 로그 알림 기능·소개·실습 IMG08·09·10 | 확인된 기능·소개·실습 사실 |
| 1 | P08 | 실제 교육 플랫폼 녹화 목록 IMG13 | 세미나·녹화·플랫폼 공유 순서 |
| 2 | P03·07 | DCIS 대표 자료 IMG06, AI 4 AI 포스터·행사 IMG02·03 | 실제 운영 주기와 대표 주제 |

실제 자료에서는 해당 발화에 필요한 부분만 확대한다. 자료를 교체해도 새로운 성과 수치나 후기 문구를 임의로 추가하지 않는다. 공개 가능한 범위와 가림 처리는 기존 사실·원본 목록의 기준을 따른다.

## 촬영·편집

첫 개인 화면에 류경동 / Data Center팀 팀장, 이관우 / Data Center팀 교육담당 자막을 표시한다. 의견 대사는 두 출연자가 자신의 표현으로 조정한다. 처음 듣는 척하는 질문이나 확인되지 않은 현장 회고는 추가하지 않는다.

PPTX는 정적 자료다. 대괄호가 들어간 발화 위치에서 장표를 표시하고, P06·P08의 추가 강조 큐는 대본과 발표자 노트를 따른다. 모든 효과 시간은 장면 안에 포함한다. 15초 AI 팀 소개는 기존 별도 영상 구간이며 이번 패키지에서 새로 제작하지 않았다.

발화 보고서는 초당 4.2/4.6/5.0음절과 장면별 호흡을 가정한 추정이다. 리허설 CSV는 아직 실측이 없다. 읽기가 빠르다고 문장을 더 채우지 말고, 인물의 호흡과 자료를 보는 시간을 함께 확인한다. 너무 긴 빈 홀드는 줄이고 대본·장면 배정을 동시에 수정한다.

## 재생성과 검증

```bash
.venv/bin/python scripts/build_editorial_package.py
bash scripts/render_editorial_slides.sh
.venv/bin/python scripts/validate_editorial_package.py
```

단일 원본은 `content/editorial-film.json`이다. 생성된 PPT·대본·콘티를 각각 손으로 수정하지 않는다. 원본의 대사·큐·발음·배정이 바뀌면 해당 장면의 이전 실측은 비워지므로 다시 측정한다. 실제 PowerPoint 렌더 후 컨택트시트를 직접 확인한다.

촬영 직전에는 참여 수치 최신성, 출연자 의견, 회사 KV, 10월 OutSystems 개최 여부를 확인한다. 마지막 사용자 확인 시점에는 개최 전이었으며 이번 본편에는 완료 주장이나 완료 횟수를 넣지 않았다.
