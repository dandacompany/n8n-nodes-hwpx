import type { INodeProperties } from 'n8n-workflow';

export const contentOperations: INodeProperties[] = [
	{
		displayName: '작업',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['content'],
			},
		},
		options: [
			{
				name: '텍스트 치환',
				value: 'replaceText',
				description: 'HWPX 문서 텍스트 치환 (ZIP 레벨)',
				action: '문서 텍스트 치환',
			},
			{
				name: '순차 텍스트 치환',
				value: 'replaceTextSequential',
				description: '플레이스홀더 각 출현을 순서대로 다른 값으로 치환',
				action: '순차 텍스트 치환',
			},
			{
				name: '구조 추출',
				value: 'extractStructure',
				description: 'HWPX 문서 텍스트 내용 및 구조 추출',
				action: '문서 구조 추출',
			},
			{
				name: '텍스트 목록',
				value: 'listTexts',
				description: 'HWPX 문서 내 모든 텍스트 요소 나열 (플레이스홀더 탐색용)',
				action: '텍스트 목록 조회',
			},
			{
				name: '페이지 설정',
				value: 'pageSetup',
				description: 'HWPX 문서 페이지 크기, 방향, 여백 수정',
				action: '페이지 설정',
			},
			{
				name: '이미지 삽입',
				value: 'insertImage',
				description: 'HWPX 문서에 이미지 삽입',
				action: '이미지 삽입',
			},
			{
				name: '이미지 교체',
				value: 'replaceImage',
				description: 'HWPX 문서의 기존 이미지를 새 이미지로 교체',
				action: '이미지 교체',
			},
			{
				name: '표 추가',
				value: 'addTable',
				description: 'JSON 데이터로 HWPX 문서에 표 추가',
				action: '표 추가',
			},
			{
				name: '머리글/바닥글 설정',
				value: 'setHeaderFooter',
				description: 'HWPX 문서 머리글/바닥글 텍스트 설정',
				action: '머리글/바닥글 설정',
			},
			{
				name: '셀 병합',
				value: 'mergeCells',
				description: 'HWPX 문서 내 표 셀 병합',
				action: '셀 병합',
			},
			{
				name: '수식 삽입',
				value: 'insertEquation',
				description: '한컴 수식 스크립트 구문으로 수학 수식 삽입 (LaTeX 아님)',
				action: '수식 삽입',
			},
			{
				name: '단 레이아웃',
				value: 'setColumnLayout',
				description: '다단 레이아웃 설정 (2단 신문형 등)',
				action: '단 레이아웃 설정',
			},
			{
				name: '단 나누기',
				value: 'insertColumnBreak',
				description: '다음 단으로 전환하는 단 나누기 삽입',
				action: '단 나누기 삽입',
			},
			{
				name: '시험지 머리글 서식',
				value: 'formatExamHeader',
				description: '한국 표준 시험지 머리글 추가 (학년도, 학년, 과목, 교시)',
				action: '시험지 머리글 서식',
			},
			{
				name: '탭 정지 추가',
				value: 'addTabStops',
				description: '수평 텍스트 정렬용 사용자 정의 탭 정지 추가',
				action: '탭 정지 추가',
			},
			{
				name: '차트 생성',
				value: 'generateChart',
				description: 'SVG 차트/그래프 생성 (함수 그래프, 삼각형, 원, 사각형)',
				action: '차트 생성',
			},
			{
				name: '메모 추가',
				value: 'addMemo',
				description: 'HWPX 문서에 메모/주석 추가',
				action: '메모 추가',
			},
			{
				name: '도형 추가',
				value: 'addShape',
				description: 'HWPX 문서에 도형 삽입 (선, 사각형, 타원)',
				action: '도형 추가',
			},
			{
				name: '변경 추적',
				value: 'trackChanges',
				description: 'HWPX 문서 변경 추적 읽기/활성화/비활성화',
				action: '변경 추적 관리',
			},
			{
				name: '스타일로 치환',
				value: 'replaceByStyle',
				description: '스타일(굵게, 기울임, 밑줄, 색상, 글꼴 크기) 기반 텍스트 치환',
				action: '스타일 기반 텍스트 치환',
			},
			{
				name: '각주/미주 추가',
				value: 'addNote',
				description: 'HWPX 문서에 각주 또는 미주 삽입',
				action: '각주/미주 추가',
			},
			{
				name: '북마크 추가',
				value: 'addBookmark',
				description: '상호 참조용 북마크 마커 삽입',
				action: '북마크 추가',
			},
			{
				name: '워터마크 추가',
				value: 'addWatermark',
				description: '문서에 텍스트 워터마크 오버레이 추가',
				action: '워터마크 추가',
			},
			{
				name: '문서 보호',
				value: 'setPassword',
				description: '읽기 전용 보호 설정/해제/상태 확인',
				action: '문서 보호 관리',
			},
			{
				name: 'Markdown 변환',
				value: 'toMarkdown',
				description: 'HWPX 문서를 Markdown 형식으로 변환',
				action: 'Markdown 변환',
			},
		],
		default: 'replaceText',
	},
];

export const contentFields: INodeProperties[] = [
	// ----------------------------------
	//         content:replaceText
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceText'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '치환 목록',
		name: 'replacements',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		default: {},
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceText'],
			},
		},
		description: '적용할 텍스트 치환 목록',
		options: [
			{
				name: 'pairs',
				displayName: '치환 쌍',
				values: [
					{
						displayName: '찾기',
						name: 'find',
						type: 'string',
						default: '',
						description: '문서에서 검색할 텍스트',
					},
					{
						displayName: '치환할 텍스트',
						name: 'replace',
						type: 'string',
						default: '',
						description: '치환할 텍스트 내용',
					},
				],
			},
		],
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceText'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceText'],
			},
		},
		options: [
			{
				displayName: '대상 파일',
				name: 'targetFiles',
				type: 'options',
				default: 'contentsXml',
				options: [
					{
						name: 'Contents XML만',
						value: 'contentsXml',
						description: 'Contents/*.xml 파일에서만 텍스트를 치환합니다 (권장)',
					},
					{
						name: '모든 XML 파일',
						value: 'allXml',
						description: '아카이브 내 모든 XML 파일에서 텍스트를 치환합니다',
					},
				],
				description: 'HWPX 아카이브 내에서 치환을 적용할 파일',
			},
		],
	},

	// ----------------------------------
	//         content:replaceTextSequential
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceTextSequential'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '검색 텍스트',
		name: 'findText',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceTextSequential'],
			},
		},
		description: '검색할 텍스트. 각 출현이 다음 값으로 순차적으로 치환됩니다.',
	},
	{
		displayName: '치환 값 (한 줄에 하나씩)',
		name: 'replaceValues',
		type: 'string',
		typeOptions: {
			rows: 8,
		},
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceTextSequential'],
			},
		},
		description:
			'한 줄에 하나의 치환 값을 입력합니다. 1번째 출현은 1번째 줄로, 2번째 출현은 2번째 줄로 치환됩니다.',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceTextSequential'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceTextSequential'],
			},
		},
		options: [
			{
				displayName: '대상 파일',
				name: 'targetFiles',
				type: 'options',
				default: 'contentsXml',
				options: [
					{
						name: 'Contents XML만',
						value: 'contentsXml',
						description: 'Contents/*.xml 파일에서만 텍스트를 치환합니다 (권장)',
					},
					{
						name: '모든 XML 파일',
						value: 'allXml',
						description: '아카이브 내 모든 XML 파일에서 텍스트를 치환합니다',
					},
				],
				description: 'HWPX 아카이브 내에서 치환을 적용할 파일',
			},
		],
	},

	// ----------------------------------
	//         content:extractStructure
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['extractStructure'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['extractStructure'],
			},
		},
		options: [
			{
				displayName: '빈 텍스트 포함',
				name: 'includeEmpty',
				type: 'boolean',
				default: false,
				description: '출력에 빈 텍스트 요소를 포함할지 여부',
			},
			{
				displayName: '섹션 필터',
				name: 'sectionFilter',
				type: 'string',
				default: '',
				description:
					'특정 섹션 파일로 필터링합니다 (예: "section0"으로 section0.xml만 추출)',
			},
		],
	},

	// ----------------------------------
	//         content:listTexts
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['listTexts'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['listTexts'],
			},
		},
		options: [
			{
				displayName: '빈 텍스트 포함',
				name: 'includeEmpty',
				type: 'boolean',
				default: false,
				description: '빈 텍스트 요소를 포함할지 여부',
			},
			{
				displayName: '섹션 필터',
				name: 'sectionFilter',
				type: 'string',
				default: '',
				description:
					'특정 파일로 필터링합니다 (예: "section0"으로 section0.xml만 검색)',
			},
			{
				displayName: '중복 제거',
				name: 'deduplicate',
				type: 'boolean',
				default: false,
				description: '출력 목록에서 중복 텍스트 항목을 제거할지 여부',
			},
		],
	},

	// ----------------------------------
	//         content:pageSetup
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['pageSetup'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '용지 크기',
		name: 'paperSize',
		type: 'options',
		default: 'A4',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['pageSetup'],
			},
		},
		options: [
			{ name: 'A4 (210×297mm)', value: 'A4' },
			{ name: 'A3 (297×420mm)', value: 'A3' },
			{ name: 'B5 (176×250mm)', value: 'B5' },
			{ name: 'Letter (216×279mm)', value: 'Letter' },
			{ name: 'Legal (216×356mm)', value: 'Legal' },
			{ name: '사용자 정의', value: 'custom' },
		],
		description: '문서의 용지 크기',
	},
	{
		displayName: '방향',
		name: 'orientation',
		type: 'options',
		default: 'portrait',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['pageSetup'],
			},
		},
		options: [
			{ name: '세로', value: 'portrait' },
			{ name: '가로', value: 'landscape' },
		],
		description: '페이지 방향',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['pageSetup'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['pageSetup'],
			},
		},
		options: [
			{
				displayName: '위쪽 여백 (Mm)',
				name: 'marginTop',
				type: 'number',
				default: 20,
				description: '위쪽 여백 (밀리미터)',
			},
			{
				displayName: '아래쪽 여백 (Mm)',
				name: 'marginBottom',
				type: 'number',
				default: 15,
				description: '아래쪽 여백 (밀리미터)',
			},
			{
				displayName: '왼쪽 여백 (Mm)',
				name: 'marginLeft',
				type: 'number',
				default: 30,
				description: '왼쪽 여백 (밀리미터)',
			},
			{
				displayName: '오른쪽 여백 (Mm)',
				name: 'marginRight',
				type: 'number',
				default: 30,
				description: '오른쪽 여백 (밀리미터)',
			},
			{
				displayName: '머리글 여백 (Mm)',
				name: 'marginHeader',
				type: 'number',
				default: 15,
				description: '머리글 여백 (밀리미터)',
			},
			{
				displayName: '바닥글 여백 (Mm)',
				name: 'marginFooter',
				type: 'number',
				default: 15,
				description: '바닥글 여백 (밀리미터)',
			},
			{
				displayName: '사용자 정의 너비 (Mm)',
				name: 'customWidth',
				type: 'number',
				default: 210,
				description: '사용자 정의 용지 너비 (밀리미터, 용지 크기가 사용자 정의일 때만 사용)',
			},
			{
				displayName: '사용자 정의 높이 (Mm)',
				name: 'customHeight',
				type: 'number',
				default: 297,
				description: '사용자 정의 용지 높이 (밀리미터, 용지 크기가 사용자 정의일 때만 사용)',
			},
		],
	},

	// ----------------------------------
	//         content:insertImage
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertImage'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '이미지 바이너리 속성',
		name: 'imageBinaryPropertyName',
		type: 'string',
		default: 'image',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertImage'],
			},
		},
		description: '이미지 파일(PNG, JPG, GIF, BMP)이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertImage'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertImage'],
			},
		},
		options: [
			{
				displayName: '위치',
				name: 'position',
				type: 'options',
				default: 'append',
				options: [
					{
						name: '끝에 추가 (문서 끝)',
						value: 'append',
					},
					{
						name: '앞에 추가 (문서 시작)',
						value: 'prepend',
					},
				],
				description: '문서에서 이미지를 삽입할 위치',
			},
			{
				displayName: '너비 (Mm)',
				name: 'widthMm',
				type: 'number',
				default: 100,
				description: '이미지 표시 너비 (밀리미터)',
			},
			{
				displayName: '높이 (Mm)',
				name: 'heightMm',
				type: 'number',
				default: 75,
				description: '이미지 표시 높이 (밀리미터)',
			},
		],
	},

	// ----------------------------------
	//         content:replaceImage
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceImage'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '이미지 바이너리 속성',
		name: 'imageBinaryPropertyName',
		type: 'string',
		default: 'image',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceImage'],
			},
		},
		description: '교체할 새 이미지 파일(PNG, JPG, GIF, BMP)이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '대상 이미지',
		name: 'targetImage',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceImage'],
			},
		},
		description: '교체할 이미지 식별자 (파일명, ID, 또는 순번). 비워두면 첫 번째 이미지를 교체합니다.',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceImage'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceImage'],
			},
		},
		options: [
			{
				displayName: '너비 (Mm)',
				name: 'widthMm',
				type: 'number',
				default: 0,
				description: '새 이미지 표시 너비 (밀리미터, 0이면 기존 크기 유지)',
			},
			{
				displayName: '높이 (Mm)',
				name: 'heightMm',
				type: 'number',
				default: 0,
				description: '새 이미지 표시 높이 (밀리미터, 0이면 기존 크기 유지)',
			},
		],
	},

	// ----------------------------------
	//         content:addTable
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTable'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '표 데이터 (JSON)',
		name: 'tableData',
		type: 'string',
		typeOptions: {
			rows: 8,
		},
		default: '[["Header1","Header2"],["Data1","Data2"]]',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTable'],
			},
		},
		description: '배열의 배열 형태의 JSON 표 데이터. 예: [["이름","나이"],["Alice","30"],["Bob","25"]].',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTable'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTable'],
			},
		},
		options: [
			{
				displayName: '위치',
				name: 'position',
				type: 'options',
				default: 'append',
				options: [
					{
						name: '끝에 추가 (문서 끝)',
						value: 'append',
					},
					{
						name: '앞에 추가 (문서 시작)',
						value: 'prepend',
					},
				],
				description: '문서에서 표를 삽입할 위치',
			},
			{
				displayName: '셀 너비 (Mm)',
				name: 'cellWidthMm',
				type: 'number',
				default: 30,
				description: '각 표 셀의 너비 (밀리미터)',
			},
			{
				displayName: '셀 높이 (Mm)',
				name: 'cellHeightMm',
				type: 'number',
				default: 10,
				description: '각 표 셀의 높이 (밀리미터)',
			},
			{
				displayName: '테두리 채우기 ID 참조',
				name: 'borderFillIDRef',
				type: 'number',
				default: 3,
				description: 'Header.xml에 정의된 테두리/채우기 스타일의 참조 ID (기본값 3은 표준 테두리)',
			},
		],
	},

	// ----------------------------------
	//         content:setHeaderFooter
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setHeaderFooter'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '머리글 텍스트',
		name: 'headerText',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setHeaderFooter'],
			},
		},
		description: '페이지 머리글에 표시할 텍스트 (비워두면 건너뜀)',
	},
	{
		displayName: '바닥글 텍스트',
		name: 'footerText',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setHeaderFooter'],
			},
		},
		description: '페이지 바닥글에 표시할 텍스트 (비워두면 건너뜀)',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setHeaderFooter'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setHeaderFooter'],
			},
		},
		options: [
			{
				displayName: '적용 페이지',
				name: 'applyPageType',
				type: 'options',
				default: 'BOTH',
				options: [
					{ name: '모든 페이지', value: 'BOTH' },
					{ name: '짝수 페이지만', value: 'EVEN' },
					{ name: '홀수 페이지만', value: 'ODD' },
				],
				description: '머리글/바닥글을 적용할 페이지',
			},
		],
	},

	// ----------------------------------
	//         content:mergeCells
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['mergeCells'],
			},
		},
		description: '표가 포함된 HWPX 파일의 바이너리 속성 이름',
	},
	{
		displayName: '시작 행',
		name: 'startRow',
		type: 'number',
		default: 0,
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['mergeCells'],
			},
		},
		description: '병합 범위의 시작 행 인덱스 (0부터 시작)',
	},
	{
		displayName: '시작 열',
		name: 'startCol',
		type: 'number',
		default: 0,
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['mergeCells'],
			},
		},
		description: '병합 범위의 시작 열 인덱스 (0부터 시작)',
	},
	{
		displayName: '끝 행',
		name: 'endRow',
		type: 'number',
		default: 1,
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['mergeCells'],
			},
		},
		description: '병합 범위의 끝 행 인덱스 (0부터 시작, 포함)',
	},
	{
		displayName: '끝 열',
		name: 'endCol',
		type: 'number',
		default: 1,
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['mergeCells'],
			},
		},
		description: '병합 범위의 끝 열 인덱스 (0부터 시작, 포함)',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['mergeCells'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['mergeCells'],
			},
		},
		options: [
			{
				displayName: '표 인덱스',
				name: 'tableIndex',
				type: 'number',
				default: 0,
				description:
					'수정할 표의 인덱스 (0부터 시작, 문서에 여러 표가 있는 경우)',
			},
		],
	},

	// ----------------------------------
	//         content:insertEquation
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertEquation'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '수식 스크립트',
		name: 'equationScript',
		type: 'string',
		typeOptions: {
			rows: 3,
		},
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertEquation'],
			},
		},
		description: '한컴 수식 스크립트 (LaTeX 아님). 예: "x^2 + 2x + 1", "{-b +- sqrt {b^2 - 4ac}} over {2a}", "sum _{k=1} ^{n} k".',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertEquation'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertEquation'],
			},
		},
		options: [
			{
				displayName: '위치',
				name: 'position',
				type: 'options',
				default: 'append',
				options: [
					{ name: '끝에 추가 (문서 끝)', value: 'append' },
					{ name: '앞에 추가 (문서 시작)', value: 'prepend' },
				],
				description: '문서에서 수식을 삽입할 위치',
			},
			{
				displayName: '기본 단위',
				name: 'baseUnit',
				type: 'number',
				default: 1000,
				description: 'HWPUNIT 단위의 글꼴 크기 (1000 = 10pt, 1200 = 12pt)',
			},
			{
				displayName: '텍스트 색상',
				name: 'textColor',
				type: 'color',
				default: '#000000',
				description: '수식 텍스트 색상 (16진수 형식)',
			},
			{
				displayName: '접두 텍스트',
				name: 'prefixText',
				type: 'string',
				default: '',
				description: '수식 앞에 표시할 텍스트 (예: "방정식 ")',
			},
			{
				displayName: '접미 텍스트',
				name: 'suffixText',
				type: 'string',
				default: '',
				description: '수식 뒤에 표시할 텍스트 (예: " 의 해를 구하라.")',
			},
		],
	},

	// ----------------------------------
	//         content:setColumnLayout
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setColumnLayout'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '단 수',
		name: 'columnCount',
		type: 'number',
		default: 2,
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setColumnLayout'],
			},
		},
		description: '단 수 (1 = 1단, 2 = 2단 등)',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setColumnLayout'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setColumnLayout'],
			},
		},
		options: [
			{
				displayName: '단 유형',
				name: 'columnType',
				type: 'options',
				default: 'NEWSPAPER',
				options: [
					{
						name: '신문형 (왼쪽에서 오른쪽)',
						value: 'NEWSPAPER',
						description: '왼쪽에서 오른쪽으로 단을 채운 후 다음 페이지로',
					},
					{
						name: '균형 신문형',
						value: 'BALANCED_NEWSPAPER',
						description: '단 전체에 내용을 균등하게 배분',
					},
					{
						name: '병렬',
						value: 'PARALLEL',
						description: '독립적인 병렬 단',
					},
				],
				description: '단 사이에서 내용이 흐르는 방식',
			},
			{
				displayName: '간격 (Mm)',
				name: 'gapMm',
				type: 'number',
				default: 8,
				description: '단 사이 간격 (밀리미터)',
			},
			{
				displayName: '동일한 단 크기',
				name: 'sameSize',
				type: 'boolean',
				default: true,
				description: '모든 단의 너비를 동일하게 할지 여부',
			},
		],
	},

	// ----------------------------------
	//         content:insertColumnBreak
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertColumnBreak'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertColumnBreak'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['insertColumnBreak'],
			},
		},
		options: [
			{
				displayName: '문단 인덱스 뒤에',
				name: 'afterParagraphIndex',
				type: 'number',
				default: -1,
				description:
					'이 문단 인덱스 뒤에 단 나누기를 삽입합니다 (0부터 시작). -1은 끝에 추가.',
			},
		],
	},

	// ----------------------------------
	//         content:formatExamHeader
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['formatExamHeader'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '학년도',
		name: 'year',
		type: 'number',
		default: 2025,
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['formatExamHeader'],
			},
		},
		description: '학년도 (예: 2025)',
	},
	{
		displayName: '시험 월',
		name: 'month',
		type: 'number',
		default: 3,
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['formatExamHeader'],
			},
		},
		description: '시험 월 (예: 3, 6, 9, 11)',
	},
	{
		displayName: '학년',
		name: 'grade',
		type: 'options',
		default: '고1',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['formatExamHeader'],
			},
		},
		options: [
			{ name: '중1', value: '중1' },
			{ name: '중2', value: '중2' },
			{ name: '중3', value: '중3' },
			{ name: '고1', value: '고1' },
			{ name: '고2', value: '고2' },
			{ name: '고3', value: '고3' },
		],
		description: '학생 학년',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['formatExamHeader'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['formatExamHeader'],
			},
		},
		options: [
			{
				displayName: '과목 영역',
				name: 'subjectArea',
				type: 'string',
				default: '수학',
				description: '과목 영역 이름 (예: 수학, 영어, 국어)',
			},
			{
				displayName: '교시',
				name: 'session',
				type: 'number',
				default: 2,
				description: '시험 교시 번호',
			},
			{
				displayName: '문제 유형 라벨',
				name: 'questionTypeLabel',
				type: 'string',
				default: '5지선다형',
				description: '머리글에 표시되는 문제 유형 라벨',
			},
			{
				displayName: '사용자 정의 시험 제목',
				name: 'examTitle',
				type: 'string',
				default: '',
				description: '자동 생성된 제목을 대체합니다. 비워두면 기본 형식 사용: "2025학년도 3월 고1 전국연합학력평가".',
			},
		],
	},

	// ----------------------------------
	//         content:addTabStops
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTabStops'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '탭 위치 (Mm)',
		name: 'tabPositions',
		type: 'string',
		default: '16,32,48,64',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTabStops'],
			},
		},
		description:
			'쉼표로 구분된 탭 정지 위치 (밀리미터) (예: 5지선다 시험 레이아웃의 경우 "16,32,48,64")',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTabStops'],
			},
		},
		description: '수정된 HWPX 파일을 저장할 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addTabStops'],
			},
		},
		options: [
			{
				displayName: '탭 유형',
				name: 'tabType',
				type: 'options',
				default: 'LEFT',
				options: [
					{ name: '왼쪽', value: 'LEFT' },
					{ name: '가운데', value: 'CENTER' },
					{ name: '오른쪽', value: 'RIGHT' },
					{ name: '소수점', value: 'DECIMAL' },
				],
				description: '탭 정지의 정렬 유형',
			},
			{
				displayName: '채움선',
				name: 'leader',
				type: 'options',
				default: 'NONE',
				options: [
					{ name: '없음', value: 'NONE' },
					{ name: '실선', value: 'SOLID' },
					{ name: '대시', value: 'DASH' },
					{ name: '점선', value: 'DOT' },
				],
				description: '탭 정지 앞의 채움선 문자 스타일',
			},
		],
	},

	// ----------------------------------
	//         content:generateChart
	// ----------------------------------
	{
		displayName: '차트 유형',
		name: 'chartType',
		type: 'options',
		default: 'function',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['generateChart'],
			},
		},
		options: [
			{
				name: '함수 그래프',
				value: 'function',
				description: '좌표 평면에 수학 함수를 그립니다',
			},
			{
				name: '삼각형',
				value: 'triangle',
				description: '꼭짓점 라벨과 측정값이 있는 삼각형을 그립니다',
			},
			{
				name: '원',
				value: 'circle',
				description: '중심, 점, 현, 접선이 있는 원을 그립니다',
			},
			{
				name: '사각형',
				value: 'quadrilateral',
				description: '라벨과 선택적 대각선이 있는 사각형을 그립니다',
			},
			{
				name: '좌표 평면',
				value: 'coordinate',
				description: '사용자 정의 점과 선이 있는 좌표 평면을 그립니다',
			},
		],
		description: '생성할 차트 또는 도형의 유형',
	},
	{
		displayName: '차트 사양 (JSON)',
		name: 'chartSpec',
		type: 'string',
		typeOptions: {
			rows: 10,
		},
		default: '{\n  "expression": "x^2",\n  "xMin": -5,\n  "xMax": 5,\n  "yMin": -2,\n  "yMax": 25\n}',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['generateChart'],
			},
		},
		description:
			'차트의 JSON 사양. 함수 그래프: {expression, xMin, xMax, yMin, yMax}. 삼각형: {vertices, labels}. 원: {center, radius, pointsOnCircle}. 전체 옵션은 문서를 참조하세요.',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['generateChart'],
			},
		},
		options: [
			{
				displayName: '너비 (Px)',
				name: 'width',
				type: 'number',
				default: 400,
				description: 'SVG 이미지 너비 (픽셀)',
			},
			{
				displayName: '높이 (Px)',
				name: 'height',
				type: 'number',
				default: 400,
				description: 'SVG 이미지 높이 (픽셀)',
			},
			{
				displayName: '배경색',
				name: 'backgroundColor',
				type: 'color',
				default: '#ffffff',
				description: '배경색 (16진수 형식)',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'chart',
				description: '생성된 SVG를 저장할 바이너리 속성의 이름',
			},
		],
	},

	// ----------------------------------
	//         content:addMemo
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addMemo'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '메모 텍스트',
		name: 'memoText',
		type: 'string',
		typeOptions: { rows: 3 },
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addMemo'],
			},
		},
		description: '추가할 메모/주석 텍스트',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addMemo'],
			},
		},
		options: [
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'memo.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
			{
				displayName: '대상 문단 인덱스',
				name: 'targetParagraphIndex',
				type: 'number',
				default: 0,
				description: '메모를 첨부할 문단의 인덱스 (0부터 시작)',
			},
		],
	},

	// ----------------------------------
	//         content:addShape
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addShape'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '도형 유형',
		name: 'shapeType',
		type: 'options',
		default: 'rectangle',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addShape'],
			},
		},
		options: [
			{ name: '타원', value: 'ellipse', description: '타원/원을 그립니다' },
			{ name: '선', value: 'line', description: '직선을 그립니다' },
			{ name: '사각형', value: 'rectangle', description: '사각형을 그립니다' },
		],
		description: '삽입할 도형 유형',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addShape'],
			},
		},
		options: [
			{
				displayName: '끝 X (HWPUNIT, 선만)',
				name: 'endX',
				type: 'number',
				default: 14173,
				description: 'HWPUNIT 단위의 선 도형 끝 X 좌표',
			},
			{
				displayName: '끝 Y (HWPUNIT, 선만)',
				name: 'endY',
				type: 'number',
				default: 0,
				description: 'HWPUNIT 단위의 선 도형 끝 Y 좌표',
			},
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'shape.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '채우기 색상',
				name: 'fillColor',
				type: 'color',
				default: '',
				description: '채우기 색상 (16진수 형식, 비워두면 채우기 없음)',
			},
			{
				displayName: '높이 (Mm)',
				name: 'heightMm',
				type: 'number',
				default: 25,
				description: '도형 높이 (밀리미터)',
			},
			{
				displayName: '선 색상',
				name: 'lineColor',
				type: 'color',
				default: '#000000',
				description: '선/테두리 색상 (16진수 형식)',
			},
			{
				displayName: '선 굵기 (Mm)',
				name: 'lineWidth',
				type: 'number',
				default: 0.4,
				description: '선 두께 (mm 단위)',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
			{
				displayName: '시작 X (HWPUNIT, 선만)',
				name: 'startX',
				type: 'number',
				default: 0,
				description: '직선 도형의 시작 X 좌표 (HWPUNIT 단위)',
			},
			{
				displayName: '시작 Y (HWPUNIT, 직선 전용)',
				name: 'startY',
				type: 'number',
				default: 0,
				description: '직선 도형의 시작 Y 좌표 (HWPUNIT 단위)',
			},
			{
				displayName: '대상 문단 인덱스',
				name: 'targetParagraphIndex',
				type: 'number',
				default: 0,
				description: '도형을 삽입할 문단의 인덱스 (0부터 시작)',
			},
			{
				displayName: '너비 (Mm)',
				name: 'widthMm',
				type: 'number',
				default: 50,
				description: '도형 너비 (mm 단위)',
			},
		],
	},

	// ----------------------------------
	//         content:trackChanges
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['trackChanges'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '동작',
		name: 'trackChangeAction',
		type: 'options',
		default: 'read',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['trackChanges'],
			},
		},
		options: [
			{ name: '비활성화', value: 'disable', description: '변경 내용 추적을 비활성화합니다' },
			{ name: '활성화', value: 'enable', description: '변경 내용 추적을 활성화합니다' },
			{ name: '읽기', value: 'read', description: '기존 변경 내용 추적을 읽습니다' },
		],
		description: '변경 내용 추적에 대해 수행할 작업',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['trackChanges'],
			},
		},
		options: [
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'tracked.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
		],
	},

	// ----------------------------------
	//         content:replaceByStyle
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceByStyle'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '치환할 텍스트',
		name: 'replaceWith',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceByStyle'],
			},
		},
		description: '일치하는 스타일의 텍스트를 대체할 텍스트',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['replaceByStyle'],
			},
		},
		options: [
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'styled.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '굵게 필터',
				name: 'filterBold',
				type: 'boolean',
				default: false,
				description: '굵은 텍스트만 대체할지 여부',
			},
			{
				displayName: '색상 필터',
				name: 'filterColor',
				type: 'color',
				default: '',
				description: '이 색상의 텍스트만 대체합니다 (예: #FF0000)',
			},
			{
				displayName: '글꼴 크기 필터',
				name: 'filterFontSize',
				type: 'number',
				default: 0,
				description: '이 글꼴 크기(pt)의 텍스트만 대체합니다 (0 = 모두)',
			},
			{
				displayName: '기울임 필터',
				name: 'filterItalic',
				type: 'boolean',
				default: false,
				description: '기울임꼴 텍스트만 대체할지 여부',
			},
			{
				displayName: '밑줄 필터',
				name: 'filterUnderline',
				type: 'boolean',
				default: false,
				description: '밑줄 텍스트만 대체할지 여부',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
		],
	},

	// ----------------------------------
	//         content:addNote
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addNote'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '주석 유형',
		name: 'noteType',
		type: 'options',
		default: 'footnote',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addNote'],
			},
		},
		options: [
			{ name: '미주', value: 'endnote', description: '미주를 추가합니다 (문서 끝에 표시)' },
			{ name: '각주', value: 'footnote', description: '각주를 추가합니다 (페이지 하단에 표시)' },
		],
		description: '삽입할 주석의 유형',
	},
	{
		displayName: '주석 텍스트',
		name: 'noteText',
		type: 'string',
		typeOptions: { rows: 3 },
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addNote'],
			},
		},
		description: '각주 또는 미주 텍스트',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addNote'],
			},
		},
		options: [
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'noted.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
			{
				displayName: '대상 문단 인덱스',
				name: 'targetParagraphIndex',
				type: 'number',
				default: 0,
				description: '주석을 첨부할 문단의 인덱스 (0부터 시작)',
			},
		],
	},

	// ----------------------------------
	//         content:addBookmark
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addBookmark'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '책갈피 이름',
		name: 'bookmarkName',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addBookmark'],
			},
		},
		description: '상호 참조를 위한 책갈피 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addBookmark'],
			},
		},
		options: [
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'bookmarked.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
			{
				displayName: '대상 문단 인덱스',
				name: 'targetParagraphIndex',
				type: 'number',
				default: 0,
				description: '책갈피를 삽입할 문단의 인덱스 (0부터 시작)',
			},
		],
	},

	// ----------------------------------
	//         content:addWatermark
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addWatermark'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '워터마크 텍스트',
		name: 'watermarkText',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addWatermark'],
			},
		},
		description: '워터마크로 표시할 텍스트',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['addWatermark'],
			},
		},
		options: [
			{
				displayName: '각도 (도)',
				name: 'angleDeg',
				type: 'number',
				default: -45,
				description: '회전 각도 (음수 = 반시계 방향)',
			},
			{
				displayName: '색상',
				name: 'color',
				type: 'color',
				default: '#CCCCCC',
				description: '워터마크 텍스트 색상',
			},
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'watermarked.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '글꼴 크기 (Pt)',
				name: 'fontSizePt',
				type: 'number',
				default: 48,
				description: '글꼴 크기 (pt 단위)',
			},
			{
				displayName: '불투명도 (%)',
				name: 'opacity',
				type: 'number',
				default: 30,
				description: '워터마크 불투명도 백분율 (0-100)',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
		],
	},

	// ----------------------------------
	//         content:setPassword
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setPassword'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '동작',
		name: 'passwordAction',
		type: 'options',
		default: 'checkProtection',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setPassword'],
			},
		},
		options: [
			{
				name: '보호 확인',
				value: 'checkProtection',
				description: '문서가 보호되어 있는지 확인합니다',
			},
			{
				name: '보호 해제',
				value: 'removeProtection',
				description: '읽기 전용 보호를 해제합니다',
			},
			{
				name: '읽기 전용 설정',
				value: 'setReadOnly',
				description: '문서를 읽기 전용으로 설정합니다',
			},
		],
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['setPassword'],
			},
		},
		options: [
			{
				displayName: '파일 이름',
				name: 'fileName',
				type: 'string',
				default: 'protected.hwpx',
				description: '출력 파일 이름',
			},
			{
				displayName: '출력 바이너리 속성',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: '출력 파일의 바이너리 속성 이름',
			},
		],
	},

	// ----------------------------------
	//         content:toMarkdown
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['toMarkdown'],
			},
		},
		description: 'HWPX 파일이 포함된 바이너리 속성의 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['content'],
				operation: ['toMarkdown'],
			},
		},
		options: [
			{
				displayName: '표 포함',
				name: 'includeTables',
				type: 'boolean',
				default: true,
				description: '표를 Markdown 표 형식으로 변환할지 여부',
			},
			{
				displayName: '출력 형식',
				name: 'outputFormat',
				type: 'options',
				default: 'text',
				options: [
					{
						name: '바이너리 파일',
						value: 'binary',
						description: '.md 바이너리 파일로 출력',
					},
					{ name: 'JSON 텍스트', value: 'text', description: 'JSON 텍스트 필드로 출력' },
				],
				description: 'Markdown 콘텐츠를 반환하는 방식',
			},
		],
	},
];
