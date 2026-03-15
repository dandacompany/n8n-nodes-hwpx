# n8n-nodes-hwpx

[![NPM Version](https://img.shields.io/npm/v/n8n-nodes-hwpx?style=flat-square)](https://www.npmjs.com/package/n8n-nodes-hwpx)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)
[![N8N Community Node](https://img.shields.io/badge/n8n-community--node-blue.svg?style=flat-square)](https://n8n.io)

한글(아래아한글) HWPX 및 HWP 문서를 n8n 워크플로우에서 직접 읽기, 생성, 편집, 변환할 수 있는 커뮤니티 노드입니다. 외부 API 없이 로컬에서 모든 처리가 수행됩니다.

## 주요 특징

- **완전 로컬 처리**: 외부 서비스나 API 키 없이 n8n 인스턴스 내에서 모든 문서 작업 수행
- **AI Agent 연동**: `usableAsTool: true` 설정으로 AI Agent 노드에서 자연어로 문서 조작 가능
- **Zero Dependencies**: esbuild 번들링으로 추가 의존성 설치 없이 바로 사용
- **30개 작업**: 문서 관리 6개 + 콘텐츠 편집 24개 작업 지원

## 기능 목록

### 문서 (Document) 리소스

| 작업 | 설명 |
|------|------|
| **생성** | 새 HWPX 문서 생성 (일반 텍스트, Markdown, Structured JSON 입력 지원) |
| **읽기** | HWPX 문서에서 텍스트, 메타데이터, 이미지 목록 추출 |
| **유효성 검사** | HWPX 문서의 구조적 유효성 검증 |
| **HTML 변환** | HWPX 문서를 HTML로 변환 (이미지, 표, 스타일 렌더링 옵션) |
| **HWP 변환** | 구형 HWP 파일을 HWPX 포맷으로 변환 |
| **템플릿 채우기** | 템플릿 HWPX의 플레이스홀더를 일괄/순차 치환 |

### 콘텐츠 (Content) 리소스

| 작업 | 설명 |
|------|------|
| **텍스트 치환** | 문서 내 텍스트 찾기/바꾸기 (ZIP 레벨 치환) |
| **순차 텍스트 치환** | 동일 플레이스홀더의 각 출현을 서로 다른 값으로 순서대로 치환 |
| **구조 추출** | 문서의 텍스트 콘텐츠 및 구조 추출 |
| **텍스트 목록** | 문서 내 모든 텍스트 요소 나열 (플레이스홀더 발견에 유용) |
| **페이지 설정** | 용지 크기, 방향, 여백 변경 |
| **이미지 삽입** | HWPX 문서에 이미지 삽입 |
| **표 추가** | JSON 데이터로 표 추가 |
| **머리글/바닥글** | 머리글 및 바닥글 설정 |
| **셀 병합** | 표 내 셀 병합 |
| **수식 삽입** | 수학 수식 삽입 |
| **단 레이아웃** | 다단 레이아웃 설정 |
| **단 나누기** | 단 나누기 삽입 |
| **시험지 머리글** | 시험지 형식 머리글 서식 지정 |
| **탭 정지** | 탭 정지 위치 추가 |
| **차트 생성** | SVG 기반 차트 생성 (꺾은선, 막대, 원형, 좌표평면) |
| **메모 추가** | 문단에 코멘트/메모 삽입 |
| **도형 추가** | 선, 사각형, 타원 도형 객체 삽입 |
| **변경 추적** | 변경 추적 읽기/활성화/비활성화 |
| **스타일로 치환** | 특정 스타일(굵게, 기울임, 색상 등) 기반 텍스트 선별 치환 |
| **각주/미주 추가** | 각주 또는 미주 삽입 |
| **북마크 추가** | 북마크 삽입 |
| **워터마크 추가** | 텍스트 워터마크 삽입 (크기, 색상, 각도, 투명도 조절) |
| **문서 보호** | 문서 보호 상태 확인/설정/해제 |
| **Markdown 변환** | HWPX 문서를 Markdown 텍스트로 변환 |

## 설치

### n8n UI에서 설치 (권장)

1. **설정 > 커뮤니티 노드**으로 이동
2. **설치** 클릭
3. `n8n-nodes-hwpx` 입력
4. **설치** 클릭 후 n8n 재시작

### npm으로 설치

```bash
cd ~/.n8n/nodes
npm install n8n-nodes-hwpx
```

설치 후 n8n을 재시작하면 노드 패널에서 **HWPX** 노드를 사용할 수 있습니다.

## 사용 예시

### 1. 문서 생성 및 텍스트 치환

```
[Manual Trigger] → [HWPX: 생성] → [HWPX: 텍스트 치환] → [결과]
```

- HWPX 생성 노드로 템플릿 문서 생성
- 텍스트 치환 노드로 `{{이름}}`, `{{날짜}}` 등 플레이스홀더 치환
- 완성된 문서를 이메일 첨부 또는 파일로 저장

### 2. HWP → HWPX → HTML 변환 파이프라인

```
[Read Binary File] → [HWPX: HWP 변환] → [HWPX: HTML 변환] → [HTML 출력]
```

- 구형 HWP 파일을 HWPX로 변환
- HWPX를 HTML로 변환하여 웹에서 표시

### 3. AI Agent와 문서 작업

```
[Chat Trigger] → [AI Agent (Tool: HWPX)] → [응답]
```

- `usableAsTool: true` 설정으로 AI Agent가 자연어로 HWPX 문서 생성/편집
- "시험지 양식으로 3문단짜리 문서를 만들어줘" 같은 요청 처리

### 4. 워터마크 + 문서 보호

```
[HWPX 입력] → [HWPX: 워터마크 추가] → [HWPX: 문서 보호] → [저장]
```

- "CONFIDENTIAL" 워터마크를 45도 각도로 삽입
- 읽기 전용 보호 설정으로 문서 잠금

## 기술 스택

- **TypeScript** + **esbuild** 번들링 (617KB 단일 파일, 외부 의존성 없음)
- **JSZip** - HWPX(ZIP) 아카이브 조작
- **@ssabrojs/hwpxjs** - HWPX 문서 파싱 및 텍스트/HTML 추출
- **OWPML** (KS X 6101) - 한국 개방형 워드프로세서 마크업 언어 표준 준수

## 호환성

- n8n v1.0+ (n8nNodesApiVersion: 1)
- Node.js 18+

## 라이선스

[MIT](LICENSE.md)

---

## Dante Labs

**Developed and maintained by Dante Labs**

- **Homepage**: [dante-labs.com](https://dante-labs.com)
- **YouTube**: [@dante-labs](https://youtube.com/@dante-labs)
- **Discord**: [Dante Labs Community](https://discord.com/invite/rXyy5e9ujs)
- **Email**: dante@dante-labs.com

### Support

If you find this project helpful, consider supporting the development!

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-Support-yellow?style=for-the-badge&logo=buy-me-a-coffee)](https://buymeacoffee.com/dante.labs)

**☕ Buy Me a Coffee**: https://buymeacoffee.com/dante.labs
