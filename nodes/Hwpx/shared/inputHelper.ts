import type { IExecuteFunctions, INodeExecutionData, INodeProperties } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

/**
 * Resolve input buffer from either binary property or URL.
 * Supports the inputSource parameter pattern for all operations.
 */
export async function resolveInputBuffer(
	ctx: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
	sourceParamName = 'inputSource',
	binaryParamName = 'inputBinaryPropertyName',
	urlParamName = 'fileUrl',
): Promise<Buffer> {
	const source = ctx.getNodeParameter(sourceParamName, itemIndex, 'binary') as string;

	if (source === 'url') {
		const url = ctx.getNodeParameter(urlParamName, itemIndex) as string;
		if (!url) {
			throw new NodeOperationError(ctx.getNode(), 'URL이 비어 있습니다', { itemIndex });
		}
		const response = await ctx.helpers.httpRequest({
			method: 'GET',
			url,
			encoding: 'arraybuffer',
		});
		return Buffer.from(response as ArrayBuffer);
	}

	// Binary mode (default)
	const binaryProp = ctx.getNodeParameter(binaryParamName, itemIndex) as string;
	const binaryData = item.binary?.[binaryProp];
	if (!binaryData) {
		throw new NodeOperationError(
			ctx.getNode(),
			`No binary data found in property "${binaryProp}"`,
			{ itemIndex },
		);
	}
	return ctx.helpers.getBinaryDataBuffer(itemIndex, binaryProp);
}

/**
 * Get the original file name from binary data or URL.
 * Falls back to the provided default name.
 */
export function getOriginalFileName(
	ctx: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
	defaultName = 'modified.hwpx',
	sourceParamName = 'inputSource',
	binaryParamName = 'inputBinaryPropertyName',
	urlParamName = 'fileUrl',
): string {
	const source = ctx.getNodeParameter(sourceParamName, itemIndex, 'binary') as string;
	if (source === 'url') {
		const url = ctx.getNodeParameter(urlParamName, itemIndex, '') as string;
		if (url) {
			const match = url.match(/\/([^/?#]+)(?:[?#]|$)/);
			if (match?.[1]) return decodeURIComponent(match[1]);
		}
		return defaultName;
	}
	const binaryProp = ctx.getNodeParameter(binaryParamName, itemIndex, 'data') as string;
	return item.binary?.[binaryProp]?.fileName ?? defaultName;
}

/**
 * Get MIME type for image input from binary data or URL extension.
 */
export function getImageMimeType(
	ctx: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): string {
	const source = ctx.getNodeParameter('imageInputSource', itemIndex, 'binary') as string;
	if (source === 'url') {
		const url = ctx.getNodeParameter('imageUrl', itemIndex, '') as string;
		const ext = url.match(/\.(\w+)(?:[?#]|$)/)?.[1]?.toLowerCase();
		const mimeMap: Record<string, string> = {
			png: 'image/png',
			jpg: 'image/jpeg',
			jpeg: 'image/jpeg',
			gif: 'image/gif',
			bmp: 'image/bmp',
			webp: 'image/webp',
			svg: 'image/svg+xml',
		};
		return mimeMap[ext ?? ''] ?? 'image/png';
	}
	const binaryProp = ctx.getNodeParameter('imageBinaryPropertyName', itemIndex, 'image') as string;
	return item.binary?.[binaryProp]?.mimeType ?? 'image/png';
}

/**
 * Generate file input parameters (inputSource + inputBinaryPropertyName + fileUrl)
 * for the given resource and operations.
 */
export function makeFileInputParams(
	resource: string,
	operations: string[],
	description = 'HWPX 파일이 포함된 바이너리 속성 이름',
): INodeProperties[] {
	return [
		{
			displayName: '입력 방식',
			name: 'inputSource',
			type: 'options',
			noDataExpression: true,
			options: [
				{
					name: '바이너리 데이터',
					value: 'binary',
				},
				{
					name: 'URL',
					value: 'url',
				},
			],
			default: 'binary',
			displayOptions: {
				show: {
					resource: [resource],
					operation: operations,
				},
			},
			description: '파일 입력 방식을 선택합니다',
		},
		{
			displayName: '입력 바이너리 속성',
			name: 'inputBinaryPropertyName',
			type: 'string',
			default: 'data',
			required: true,
			displayOptions: {
				show: {
					resource: [resource],
					operation: operations,
					inputSource: ['binary'],
				},
			},
			description,
		},
		{
			displayName: '파일 URL',
			name: 'fileUrl',
			type: 'string',
			default: '',
			required: true,
			placeholder: 'https://example.com/document.hwpx',
			displayOptions: {
				show: {
					resource: [resource],
					operation: operations,
					inputSource: ['url'],
				},
			},
			description: '다운로드할 파일의 URL',
		},
	];
}

/**
 * Generate image input parameters (imageInputSource + imageBinaryPropertyName + imageUrl)
 * for operations that accept image files.
 */
export function makeImageInputParams(
	resource: string,
	operations: string[],
	description = '이미지 파일(PNG, JPG, GIF, BMP)이 포함된 바이너리 속성의 이름',
): INodeProperties[] {
	return [
		{
			displayName: '이미지 입력 방식',
			name: 'imageInputSource',
			type: 'options',
			noDataExpression: true,
			options: [
				{
					name: '바이너리 데이터',
					value: 'binary',
				},
				{
					name: 'URL',
					value: 'url',
				},
			],
			default: 'binary',
			displayOptions: {
				show: {
					resource: [resource],
					operation: operations,
				},
			},
			description: '이미지 입력 방식을 선택합니다',
		},
		{
			displayName: '이미지 바이너리 속성',
			name: 'imageBinaryPropertyName',
			type: 'string',
			default: 'image',
			required: true,
			displayOptions: {
				show: {
					resource: [resource],
					operation: operations,
					imageInputSource: ['binary'],
				},
			},
			description,
		},
		{
			displayName: '이미지 URL',
			name: 'imageUrl',
			type: 'string',
			default: '',
			required: true,
			placeholder: 'https://example.com/image.png',
			displayOptions: {
				show: {
					resource: [resource],
					operation: operations,
					imageInputSource: ['url'],
				},
			},
			description: '다운로드할 이미지 파일의 URL',
		},
	];
}
