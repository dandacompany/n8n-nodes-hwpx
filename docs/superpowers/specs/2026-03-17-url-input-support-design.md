# URL 입력 지원 설계

## 목적
HWPX 노드의 모든 파일 입력(HWPX/HWP, 이미지)에 URL 직접 입력 옵션을 추가하여
AI Agent tool로 사용할 때 HTTP Request 노드 없이 URL만으로 문서를 처리할 수 있게 한다.

## 범위
- 입력: 바이너리 / URL 선택 지원 (document 5개 + content 24개 operation)
- 이미지: insertImage, replaceImage의 이미지 입력도 URL 지원
- 출력: 기존 바이너리 방식 유지 (변경 없음)

## 구현
1. `shared/inputHelper.ts` — 공통 파라미터 배열 + `resolveInputBuffer()` 헬퍼
2. `resources/document/index.ts` — inputSource 파라미터로 교체
3. `resources/content/index.ts` — 동일 적용
4. `shared/documentOps.ts` — resolveInputBuffer() 호출로 교체
5. `shared/contentOps.ts` — 동일 적용

## 호환성
- `inputSource` default = `'binary'` → 기존 워크플로우 영향 없음
- `ctx.helpers.httpRequest` 사용 → 외부 의존성 없음
