# cms-pptx-builder

자료와 디자인 지침을 바탕으로 수업용 PPTX의 제작 계획, 발표자 노트, 구조 검증과 시각 검수를 돕는 Codex 스킬입니다.

## 설치

이 폴더를 Codex 스킬 디렉터리의 `cms-pptx-builder` 위치에 두고 Codex를 다시 시작합니다. 운영체제별 기본 위치는 다음과 같습니다.

```text
Windows: %USERPROFILE%\.codex\skills\cms-pptx-builder
macOS/Linux: ~/.codex/skills/cms-pptx-builder
CODEX_HOME을 설정한 경우: $CODEX_HOME/skills/cms-pptx-builder
```

별도 패키지 설치 없이 문서 지침을 사용할 수 있지만, PPTX를 만들려면 현재 환경에 하나 이상의 로컬 제작 도구가 필요합니다. 구조 검증에는 제공되는 `scripts/validate_pptx.py` 또는 동등한 검사 방법이 필요합니다. 시각 검수에는 PowerPoint, LibreOffice 계열 또는 동등한 로컬 렌더러가 필요합니다.

## 사용법

Codex에서 `$cms-pptx-builder`를 호출하고 다음 입력을 제공합니다.

- 대상 자료의 정확한 경로
- 주제 또는 주제 선정을 위임한다는 명시적 지시
- PPTX 출력 경로
- 화면 비율, 글꼴, 디자인과 발표자 노트 요구
- 웹·외부 자료 사용 허용 여부

제작 전에 `scripts/preflight.py`를 실행하거나 [references/environment-preflight.md](references/environment-preflight.md)의 수동 점검을 수행합니다. 제작 후 `scripts/validate_pptx.py`를 실행하거나 동등 검증을 하고, 가능한 경우 전체 슬라이드를 렌더링해 시각 검수합니다.

## 호환 모드

`presentations:Presentations`를 우선 사용하지만, 현재 환경에 없으면 호환 모드로 가능한 로컬 제작·검증·렌더링 도구를 선택합니다. 기준 렌더러, 실제 사용 렌더러, 미검증 엔진과 대체 글꼴을 결과에 기록합니다. 권장 도구가 없다는 이유만으로 중단하지 않지만, 제작 도구 전무·PPTX 손상·구조 검증 불능·필수 발표자 노트 누락은 `FAIL`로 중단합니다.

자세한 규칙은 [references/compatibility-mode.md](references/compatibility-mode.md)와 [references/validation-criteria.md](references/validation-criteria.md)를 참고합니다.

## 제한사항

- 입력값과 주제를 임의로 만들지 않습니다.
- `01_자료` 등 원본은 읽기 전용으로 취급합니다.
- Pretendard가 없으면 실제 설치된 한글 글꼴로 대체하고 `확인 필요`를 보고합니다.
- 구조 검증은 시각적 품질을 보장하지 않습니다. 렌더링하지 못한 항목을 자동으로 통과시키지 않습니다.
- 렌더러별 줄바꿈·글꼴·도형 차이는 사람의 최종 확인이 필요할 수 있습니다.

## 라이선스

MIT License. 자세한 내용은 [LICENSE](LICENSE)를 확인하십시오.
