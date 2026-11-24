# 번들 스크립트

실행은 `python -I <script>` / `powershell -ExecutionPolicy Bypass -File <script>`. 각 스크립트 머리말에 자세한 사용법이 있다.

필요: Python 3.11+, `python-pptx`, `lxml`, `Pillow`. `.ps1`은 Windows + PowerPoint(Windows에 기본으로 있는 PowerShell로 실행된다).

| 스크립트 | 용도 | 쓰기 |
|---|---|---|
| `deck_profile.py <deck> [slides]` | 템플릿·덱의 레이아웃과 슬라이드 목록, 슬라이드별 도형 상세(위치, 색, 글꼴) | 읽기 전용 |
| `storyboard.py make <storyline.md> <out.pptx> [--template T] [--layout L]` / `read <deck> <storyline.md>` / `map <deck>` | 뼈대 덱 만들기 / 사용자가 고친 뼈대 덱을 `storyline.md`로 읽기 / 1번 장 흐름도 다시 그리기 | make는 새 파일, read는 `storyline.md` 갱신(.bak 보관), map은 덱을 그 자리에서 고침 |
| `storyline_check.py <storyline.md> [--json out.json]` | 스토리라인 파일 읽기: 슬라이드 표, 장별 분량, 발표 시간 대비 장수, 비어 있는 장, 없는 자료, `?` 목록. `--json`은 조립용 데이터 | 읽기 전용 |
| `deck_check.py <deck> [--max-chars N] [--min-pt N] [--text]` | 내용 점검: 남은 표시, 빈 자리, 글 분량, 작은 글자, 그림·표 번호, 인용과 참고문헌 대응 | 읽기 전용 |
| `ppt_export.ps1 <deck> <dir> [-Slides "9,12"] [-Width 1600]` | PowerPoint로 슬라이드를 PNG로 렌더링(`sNN.png`) | 읽기 전용 |
| `contact_sheet.py <render_dir> <out_dir> [--cols 3] [--rows 3] [--width 1800]` | 렌더링한 슬라이드를 여러 장씩 한 그림으로 모음(검수 첫 훑어보기용) | 새 파일 |
| `progress.py init <storyline.md> <progress.json>` / `set <json> <장> <상태> [메모]` / `stage <json> <0-8>` / `show <json>` | 장별 작업 상태와 현재 단계를 파일에 기록하고 보여 줌 | `progress.json` 갱신 |
| `copy_slides.py <src> <s1,s2> <dst> <after>` / `move <dst> <from> <to>` | 덱 사이 슬라이드 복사, 덱 안에서 이동 | 덱을 그 자리에서 고침 |
| `para_swap.py extract <src> <dict.json> hangul\|latin` / `build <src> <dict.json> <out>` | 단락 단위 문구 추출·치환(영문판 등). 서식 유지 | build는 새 파일 |
| `fit_boxes.py <before> <after>` | 문구가 길어져 넘치는 텍스트 상자를 넓힘 | 덱을 그 자리에서 고침 |

그 자리에서 고치는 스크립트는 사용자가 직접 고친 덱에 쓰기 전에 백업한다.

## 알려진 제약

- `storyboard.py`
  - 흐름도(1번 장)의 상자는 보기 전용이다. 거기서 옮기거나 고친 것은 읽지 않는다(문서 제목과
    `키: 값` 줄은 읽는다).
  - 사용자가 새로 만든 슬라이드는 맨 위 글 상자를 제목으로 읽는다. 표 안의 글과 그림 속 글은 읽지 못한다.
  - `make`는 덱을 새로 만든다. 사용자가 고친 뼈대 덱 위에 다시 돌리면 손본 배치가 사라진다.
  - 덱이 PowerPoint에 열려 있으면 `map`의 저장이 실패한다.

- `deck_check.py`
  - 캡션은 단락이 "그림 N", "Figure N", "Fig. N", "표 N", "Table N"으로 시작할 때만 인식한다
    (앞에 `<`, `[`, `(`가 붙어도 됨). 템플릿의 캡션 표기가 다르면 스크립트 위쪽의 `FIG`, `TAB`을 고친다.
  - 인용은 번호식 `[1]`, `[2, 3]`, `[4-6]`만 대조한다. 저자-연도식은 눈으로 대조한다.
  - 참고문헌 슬라이드는 제목에 "참고문헌", "References" 등이 있어야 인식한다.
  - 그림 파일 안의 글자, 차트 안의 글자는 읽지 못한다.
- `para_swap.py build`는 1번 슬라이드(표지)를 바꾸지 않는다. 표와 차트 안의 글이 빠질 수 있으니
  `deck_check.py --text`로 남은 원문을 확인한다.
- `fit_boxes.py`는 슬라이드 너비를 960pt(16:9, 13.33in)로 가정하고 글자 폭을 추정한다. 4:3 템플릿이면
  `SLIDE_W`를 고친다. 결과는 렌더링으로 확인한다.
- `ppt_export.ps1`은 임시 폴더의 복사본을 열기 때문에 원본이 PowerPoint에 열려 있어도 동작하지만,
  저장하지 않은 변경은 반영되지 않는다. PowerPoint가 없는 환경에서는
  `soffice --headless --convert-to pdf <deck>` 뒤 PDF를 그림으로 바꿔 본다.
- 프로젝트에 맞게 고친 스크립트는 `work/`에 사본을 두고 쓴다(스킬 폴더의 원본을 덮어쓰지 않는다).
