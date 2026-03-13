import { configWithoutCloudSupport as config } from '@n8n/node-cli/eslint';

export default [
	...config,
	{
		rules: {
			// 한글 UI 텍스트와 호환되지 않는 영문 기준 린트 규칙 비활성화
			'n8n-nodes-base/node-param-description-boolean-without-whether': 'off',
			'n8n-nodes-base/node-param-collection-type-unsorted-items': 'off',
			'n8n-nodes-base/node-param-options-type-unsorted-items': 'off',
		},
	},
];
