import type { INodeProperties } from 'n8n-workflow';

export const documentOperations: INodeProperties[] = [
	{
		displayName: '작업',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['document'],
			},
		},
		options: [
			{
				name: 'Create',
				value: 'create',
				description: '새 HWPX 문서를 생성합니다',
				action: '문서 생성',
			},
			{
				name: 'Read',
				value: 'read',
				description: 'HWPX 문서에서 텍스트/메타데이터를 읽고 추출합니다',
				action: '문서 읽기',
			},
			{
				name: 'Validate',
				value: 'validate',
				description: 'HWPX 문서의 구조를 검증합니다',
				action: '문서 검증',
			},
			{
				name: 'To HTML',
				value: 'toHtml',
				description: 'HWPX 문서를 HTML로 변환합니다',
				action: '문서를 HTML로 변환',
			},
			{
				name: 'Convert HWP',
				value: 'convertHwp',
				description: 'HWP 파일을 HWPX 형식으로 변환합니다',
				action: 'HWP를 HWPX로 변환',
			},
			{
				name: 'Fill Template',
				value: 'fillTemplate',
				description:
					'템플릿 HWPX에 일괄 및 순차 텍스트 치환을 적용합니다',
				action: '템플릿 문서 채우기',
			},
		],
		default: 'read',
	},
];

export const documentFields: INodeProperties[] = [
	// ----------------------------------
	//         document:create
	// ----------------------------------
	{
		displayName: '파일 이름',
		name: 'fileName',
		type: 'string',
		default: 'document.hwpx',
		required: true,
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['create'],
			},
		},
		description: '새 HWPX 문서의 파일 이름',
	},
	{
		displayName: '초기 텍스트',
		name: 'initialText',
		type: 'string',
		typeOptions: {
			rows: 5,
		},
		default: '',
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['create'],
			},
		},
		description:
			'문서 본문 내용. 입력 형식 옵션에 따라 해석됩니다: 일반 텍스트(기본값), Markdown (**bold**, *italic*, # heading), 또는 Structured JSON.',
	},
	{
		displayName: '바이너리 속성',
		name: 'binaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['create'],
			},
		},
		description: '생성된 HWPX 파일을 저장할 바이너리 속성 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['create'],
			},
		},
		options: [
			{
				displayName: '입력 형식',
				name: 'inputFormat',
				type: 'options',
				default: 'plainText',
				options: [
					{
						name: 'Plain Text',
						value: 'plainText',
						description: '서식 없는 일반 텍스트',
					},
					{
						name: 'Markdown',
						value: 'markdown',
						description:
							'Markdown 문법: **bold**, *italic*, ~~strikeout~~, # heading',
					},
					{
						name: 'Structured JSON',
						value: 'structuredJson',
						description: 'Bold, italic, underline, fontSize, textColor를 완전히 제어할 수 있는 paragraphs 배열 JSON',
					},
				],
				description: '초기 텍스트 내용을 해석하는 방식',
			},
		],
	},

	// ----------------------------------
	//         document:read
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['read'],
			},
		},
		description: '읽을 HWPX 파일이 포함된 바이너리 속성 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['read'],
			},
		},
		options: [
			{
				displayName: '메타데이터 포함',
				name: 'includeMetadata',
				type: 'boolean',
				default: true,
				description: '문서 메타데이터(제목, 작성자 등)를 포함할지 여부',
			},
			{
				displayName: '파일 목록 포함',
				name: 'includeFileList',
				type: 'boolean',
				default: false,
				description: 'HWPX 아카이브 내부의 파일 목록을 포함할지 여부',
			},
			{
				displayName: '이미지 포함',
				name: 'includeImages',
				type: 'boolean',
				default: false,
				description: '문서 내 이미지 목록을 포함할지 여부',
			},
		],
	},

	// ----------------------------------
	//         document:validate
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['validate'],
			},
		},
		description: '검증할 HWPX 파일이 포함된 바이너리 속성 이름',
	},

	// ----------------------------------
	//         document:toHtml
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['toHtml'],
			},
		},
		description: '변환할 HWPX 파일이 포함된 바이너리 속성 이름',
	},
	{
		displayName: '옵션',
		name: 'options',
		type: 'collection',
		placeholder: '옵션 추가',
		default: {},
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['toHtml'],
			},
		},
		options: [
			{
				displayName: '이미지 렌더링',
				name: 'renderImages',
				type: 'boolean',
				default: true,
				description: 'HTML 출력에 이미지를 포함할지 여부',
			},
			{
				displayName: '표 렌더링',
				name: 'renderTables',
				type: 'boolean',
				default: true,
				description: 'HTML 출력에 표를 포함할지 여부',
			},
			{
				displayName: '스타일 렌더링',
				name: 'renderStyles',
				type: 'boolean',
				default: true,
				description: '텍스트 스타일(bold, italic, 색상 등)을 적용할지 여부',
			},
			{
				displayName: '이미지 임베드',
				name: 'embedImages',
				type: 'boolean',
				default: false,
				description: '이미지를 Base64 데이터 URL로 임베드할지 여부',
			},
			{
				displayName: '문단 태그',
				name: 'paragraphTag',
				type: 'string',
				default: 'p',
				description: '문단에 사용할 HTML 태그 (예: "p", "div")',
			},
		],
	},

	// ----------------------------------
	//         document:convertHwp
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['convertHwp'],
			},
		},
		description: '변환할 HWP 파일이 포함된 바이너리 속성 이름',
	},
	{
		displayName: '출력 파일 이름',
		name: 'fileName',
		type: 'string',
		default: 'converted.hwpx',
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['convertHwp'],
			},
		},
		description: '변환된 HWPX 출력 파일 이름',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['convertHwp'],
			},
		},
		description: '변환된 HWPX 파일을 저장할 바이너리 속성 이름',
	},

	// ----------------------------------
	//         document:fillTemplate
	// ----------------------------------
	{
		displayName: '입력 바이너리 속성',
		name: 'inputBinaryPropertyName',
		type: 'string',
		default: 'data',
		required: true,
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['fillTemplate'],
			},
		},
		description: '템플릿 HWPX 파일이 포함된 바이너리 속성 이름',
	},
	{
		displayName: '일괄 치환',
		name: 'replacements',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		default: {},
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['fillTemplate'],
			},
		},
		description: '각 검색 텍스트의 모든 항목을 치환 값으로 대체합니다',
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
						description: '검색할 텍스트 (모든 항목이 치환됩니다)',
					},
					{
						displayName: '바꿀 내용',
						name: 'replace',
						type: 'string',
						default: '',
						description: '대체할 텍스트',
					},
				],
			},
		],
	},
	{
		displayName: '순차 치환',
		name: 'sequentialReplacements',
		type: 'fixedCollection',
		typeOptions: {
			multipleValues: true,
		},
		default: {},
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['fillTemplate'],
			},
		},
		description:
			'플레이스홀더의 각 항목을 순차적인 값으로 대체합니다 (1번째 항목은 1번째 값, 2번째 항목은 2번째 값 등)',
		options: [
			{
				name: 'groups',
				displayName: '순차 그룹',
				values: [
					{
						displayName: '찾기',
						name: 'find',
						type: 'string',
						default: '',
						description: '순차적으로 치환할 플레이스홀더 텍스트',
					},
					{
						displayName: '치환 값 (한 줄에 하나씩)',
						name: 'values',
						type: 'string',
						typeOptions: {
							rows: 5,
						},
						default: '',
						description:
							'한 줄에 하나의 값. 1번째 항목은 1번째 줄로, 2번째 항목은 2번째 줄로 대체됩니다.',
					},
				],
			},
		],
	},
	{
		displayName: '출력 파일 이름',
		name: 'fileName',
		type: 'string',
		default: 'filled.hwpx',
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['fillTemplate'],
			},
		},
		description: '채워진 출력 문서의 파일 이름',
	},
	{
		displayName: '출력 바이너리 속성',
		name: 'outputBinaryPropertyName',
		type: 'string',
		default: 'data',
		displayOptions: {
			show: {
				resource: ['document'],
				operation: ['fillTemplate'],
			},
		},
		description: '채워진 HWPX 파일을 저장할 바이너리 속성 이름',
	},
];
