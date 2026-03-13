import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';

import { documentOperations, documentFields } from './resources/document';
import { contentOperations, contentFields } from './resources/content';
import {
	createDocument,
	readDocument,
	validateDocument,
	toHtml,
	convertHwp,
	fillTemplate,
} from './shared/documentOps';
import {
	replaceText,
	replaceTextSequential,
	extractStructure,
	listTexts,
	pageSetup,
	insertImage,
	addTable,
	setHeaderFooter,
	mergeCells,
	insertEquation,
	setColumnLayout,
	insertColumnBreak,
	formatExamHeader,
	addTabStops,
	generateChart,
	addMemo,
	addShape,
	trackChanges,
	replaceByStyle,
	addNote,
	addBookmark,
	addWatermark,
	setPassword,
	toMarkdown,
} from './shared/contentOps';

export class Hwpx implements INodeType {
	description: INodeTypeDescription = {
		usableAsTool: true,
		displayName: 'HWPX',
		name: 'hwpx',
		icon: 'file:hwpx.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{ $parameter["operation"] + ": " + $parameter["resource"] }}',
		description: 'HWPX/HWP 문서 읽기, 생성, 편집, 변환, 유효성 검사',
		defaults: {
			name: 'HWPX',
		},
		inputs: ['main'],
		outputs: ['main'],
		properties: [
			{
				displayName: '리소스',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: '문서',
						value: 'document',
					},
					{
						name: '콘텐츠',
						value: 'content',
					},
				],
				default: 'document',
			},
			...documentOperations,
			...documentFields,
			...contentOperations,
			...contentFields,
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		const resource = this.getNodeParameter('resource', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;

		for (let i = 0; i < items.length; i++) {
			try {
				if (resource === 'document') {
					if (operation === 'create') {
						returnData.push(await createDocument.call(this, i));
					} else if (operation === 'read') {
						returnData.push(await readDocument.call(this, i, items[i]));
					} else if (operation === 'validate') {
						returnData.push(await validateDocument.call(this, i, items[i]));
					} else if (operation === 'toHtml') {
						returnData.push(await toHtml.call(this, i, items[i]));
					} else if (operation === 'convertHwp') {
						returnData.push(await convertHwp.call(this, i, items[i]));
					} else if (operation === 'fillTemplate') {
						returnData.push(await fillTemplate.call(this, i, items[i]));
					}
				} else if (resource === 'content') {
					if (operation === 'replaceText') {
						returnData.push(await replaceText.call(this, i, items[i]));
					} else if (operation === 'replaceTextSequential') {
						returnData.push(await replaceTextSequential.call(this, i, items[i]));
					} else if (operation === 'extractStructure') {
						returnData.push(await extractStructure.call(this, i, items[i]));
					} else if (operation === 'listTexts') {
						returnData.push(await listTexts.call(this, i, items[i]));
					} else if (operation === 'pageSetup') {
						returnData.push(await pageSetup.call(this, i, items[i]));
					} else if (operation === 'insertImage') {
						returnData.push(await insertImage.call(this, i, items[i]));
					} else if (operation === 'addTable') {
						returnData.push(await addTable.call(this, i, items[i]));
					} else if (operation === 'setHeaderFooter') {
						returnData.push(await setHeaderFooter.call(this, i, items[i]));
					} else if (operation === 'mergeCells') {
						returnData.push(await mergeCells.call(this, i, items[i]));
					} else if (operation === 'insertEquation') {
						returnData.push(await insertEquation.call(this, i, items[i]));
					} else if (operation === 'setColumnLayout') {
						returnData.push(await setColumnLayout.call(this, i, items[i]));
					} else if (operation === 'insertColumnBreak') {
						returnData.push(await insertColumnBreak.call(this, i, items[i]));
					} else if (operation === 'formatExamHeader') {
						returnData.push(await formatExamHeader.call(this, i, items[i]));
					} else if (operation === 'addTabStops') {
						returnData.push(await addTabStops.call(this, i, items[i]));
					} else if (operation === 'generateChart') {
						returnData.push(await generateChart.call(this, i));
					} else if (operation === 'addMemo') {
						returnData.push(await addMemo.call(this, i, items[i]));
					} else if (operation === 'addShape') {
						returnData.push(await addShape.call(this, i, items[i]));
					} else if (operation === 'trackChanges') {
						returnData.push(await trackChanges.call(this, i, items[i]));
					} else if (operation === 'replaceByStyle') {
						returnData.push(await replaceByStyle.call(this, i, items[i]));
					} else if (operation === 'addNote') {
						returnData.push(await addNote.call(this, i, items[i]));
					} else if (operation === 'addBookmark') {
						returnData.push(await addBookmark.call(this, i, items[i]));
					} else if (operation === 'addWatermark') {
						returnData.push(await addWatermark.call(this, i, items[i]));
					} else if (operation === 'setPassword') {
						returnData.push(await setPassword.call(this, i, items[i]));
					} else if (operation === 'toMarkdown') {
						returnData.push(await toMarkdown.call(this, i, items[i]));
					}
				}
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({
						json: { error: (error as Error).message },
						pairedItem: { item: i },
					});
					continue;
				}
				throw new NodeOperationError(this.getNode(), error as Error, { itemIndex: i });
			}
		}

		return [returnData];
	}
}
