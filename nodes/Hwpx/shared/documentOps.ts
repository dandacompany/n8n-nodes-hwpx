import type { IExecuteFunctions, INodeExecutionData, IDataObject } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import JSZip from 'jszip';
import { HwpxReader, HwpConverter  } from '@ssabrojs/hwpxjs';
import {
	parseMarkdown,
	parseStructuredJson,
	collectStyles,
	buildCharPrEntries,
	injectCharProperties,
	buildStyledSectionXml,
} from './formatParser';

/**
 * Base64-encoded blank HWPX template generated from python-hwpx reference.
 * Contains full OWPML structure: META-INF, mimetype, content.hpf,
 * header.xml (42KB fonts/styles/layout), settings.xml, version.xml, Preview.
 * mimetype is STORED (OCF spec), all others DEFLATE.
 */
const BLANK_HWPX_BASE64 =
	'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';

function buildSectionXml(text: string): string {
	const escapedText = text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');

	const paragraphs = escapedText
		.split(/\n/)
		.map(
			(line) =>
				`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
				`<hp:run charPrIDRef="0">` +
				`<hp:rPr/>` +
				`<hp:t>${line}</hp:t>` +
				`</hp:run>` +
				`</hp:p>`,
		)
		.join('');

	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"` +
		` xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">` +
		paragraphs +
		`</hs:sec>`
	);
}

/** Number of charPr entries already defined in the template header.xml */
const TEMPLATE_CHARPR_COUNT = 7;

/**
 * Create a new HWPX document from a validated python-hwpx template.
 * Supports plain text, Markdown, and structured JSON input formats.
 *
 * For styled formats (markdown/structuredJson), dynamically injects
 * charPr definitions into header.xml and generates styled section XML.
 */
export async function createDocument(
	this: IExecuteFunctions,
	itemIndex: number,
): Promise<INodeExecutionData> {
	const fileName = this.getNodeParameter('fileName', itemIndex) as string;
	const initialText = this.getNodeParameter('initialText', itemIndex, '') as string;
	const binaryPropertyName = this.getNodeParameter(
		'binaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		inputFormat?: string;
	};
	const inputFormat = options.inputFormat || 'plainText';

	const templateBuffer = Buffer.from(BLANK_HWPX_BASE64, 'base64');
	const zip = await JSZip.loadAsync(templateBuffer);

	if (initialText) {
		if (inputFormat === 'plainText') {
			zip.file('Contents/section0.xml', buildSectionXml(initialText));
		} else {
			// Parse content based on format
			const paragraphs =
				inputFormat === 'markdown'
					? parseMarkdown(initialText)
					: parseStructuredJson(initialText);

			// Collect unique styles, starting IDs after template's existing charPr entries
			const styles = collectStyles(paragraphs, TEMPLATE_CHARPR_COUNT);

			// Only modify header.xml if there are new styles beyond the default
			const newStyleCount = styles.size - 1; // subtract the default (id=0) mapping
			if (newStyleCount > 0) {
				const charPrXml = buildCharPrEntries(styles);
				const headerFile = zip.file('Contents/header.xml');
				if (headerFile) {
					const headerXml = await headerFile.async('string');
					const updatedHeader = injectCharProperties(
						headerXml,
						charPrXml,
						TEMPLATE_CHARPR_COUNT + newStyleCount,
					);
					zip.file('Contents/header.xml', updatedHeader);
				}
			}

			// Generate styled section XML
			zip.file('Contents/section0.xml', buildStyledSectionXml(paragraphs, styles));
		}
	}

	// OCF spec: mimetype must be STORED (uncompressed)
	const mimetypeContent = await zip.file('mimetype')?.async('string');
	if (mimetypeContent) {
		zip.file('mimetype', mimetypeContent, { compression: 'STORE' });
	}

	const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
	const binaryData = await this.helpers.prepareBinaryData(buffer, fileName, 'application/hwp+zip');

	return {
		json: { fileName, size: buffer.length, inputFormat },
		binary: { [binaryPropertyName]: binaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Fill a template HWPX document with batch and sequential replacements.
 *
 * Combines two replacement strategies in one operation:
 *   1. Batch: each find text is replaced everywhere with its replace value
 *   2. Sequential: each occurrence of a find text is replaced with successive values
 */
export async function fillTemplate(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const outputBinaryPropertyName = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const fileName = this.getNodeParameter('fileName', itemIndex, 'filled.hwpx') as string;

	// Batch replacements (optional)
	const replacementsParam = this.getNodeParameter('replacements', itemIndex, {}) as {
		pairs?: Array<{ find: string; replace: string }>;
	};

	// Sequential replacements (optional)
	const sequentialParam = this.getNodeParameter('sequentialReplacements', itemIndex, {}) as {
		groups?: Array<{ find: string; values: string }>;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const pairs = replacementsParam.pairs ?? [];
	const seqGroups = (sequentialParam.groups ?? []).map((g) => ({
		find: g.find,
		values: g.values.split('\n').filter((v: string) => v.length > 0),
	}));

	if (pairs.length === 0 && seqGroups.length === 0) {
		throw new NodeOperationError(
			this.getNode(),
			'At least one batch or sequential replacement is required',
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	let batchCount = 0;
	let seqCount = 0;

	// Track sequential value indices per group
	const seqIndices = seqGroups.map(() => 0);

	// Sort filenames for consistent section ordering
	const filenames = Object.keys(zip.files).sort();

	for (const filename of filenames) {
		const file = zip.files[filename];
		if (file.dir) continue;
		if (!filename.startsWith('Contents/') || !filename.endsWith('.xml')) continue;

		let content = await file.async('string');
		let modified = false;

		// Phase 1: Batch replacements (all occurrences)
		for (const pair of pairs) {
			if (pair.find && content.includes(pair.find)) {
				const occurrences = content.split(pair.find).length - 1;
				content = content.split(pair.find).join(pair.replace);
				batchCount += occurrences;
				modified = true;
			}
		}

		// Phase 2: Sequential replacements (one occurrence per value)
		for (let gi = 0; gi < seqGroups.length; gi++) {
			const group = seqGroups[gi];
			while (seqIndices[gi] < group.values.length && content.includes(group.find)) {
				content = content.replace(group.find, group.values[seqIndices[gi]]);
				seqIndices[gi]++;
				seqCount++;
				modified = true;
			}
		}

		if (modified) {
			zip.file(filename, content);
		}
	}

	// OCF spec: mimetype must be STORED
	const mimetypeContent = await zip.file('mimetype')?.async('string');
	if (mimetypeContent) {
		zip.file('mimetype', mimetypeContent, { compression: 'STORE' });
	}

	const outputBuffer = await zip.generateAsync({
		type: 'nodebuffer',
		compression: 'DEFLATE',
	});

	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: {
			fileName,
			size: outputBuffer.length,
			batchReplacements: batchCount,
			sequentialReplacements: seqCount,
			totalReplacements: batchCount + seqCount,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Read HWPX document using HwpxReader — extracts text, metadata, and images.
 */
export async function readDocument(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		includeMetadata?: boolean;
		includeFileList?: boolean;
		includeImages?: boolean;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const reader = new HwpxReader();
	await reader.loadFromArrayBuffer(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer);

	const text = await reader.extractText();

	const result: IDataObject = {
		text,
		textLength: text.length,
	};

	if (options.includeMetadata !== false) {
		const info = await reader.getDocumentInfo();
		result.metadata = info.metadata as unknown as IDataObject;
	}

	if (options.includeImages) {
		const images = await reader.listImages();
		result.images = images;
	}

	if (options.includeFileList) {
		const info = await reader.getDocumentInfo();
		result.files = info.summary.contentsFiles;
	}

	return {
		json: result,
		pairedItem: { item: itemIndex },
	};
}

/**
 * Validate the structure of an HWPX document.
 */
export async function validateDocument(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const errors: string[] = [];
	const warnings: string[] = [];

	let zip: JSZip;
	try {
		zip = await JSZip.loadAsync(buffer);
	} catch {
		return {
			json: { valid: false, errors: ['File is not a valid ZIP archive'], warnings: [] },
			pairedItem: { item: itemIndex },
		};
	}

	const fileList = Object.keys(zip.files);

	for (const path of ['Contents/content.hpf', 'Contents/header.xml', 'version.xml']) {
		if (!fileList.some((f) => f === path)) {
			errors.push(`Missing required file: ${path}`);
		}
	}

	if (!fileList.some((f) => f.startsWith('Contents/section') && f.endsWith('.xml'))) {
		errors.push('No section XML files found in Contents/');
	}

	if (!fileList.some((f) => f === 'mimetype')) {
		warnings.push('Missing mimetype file');
	}

	if (!fileList.some((f) => f === 'META-INF/container.xml')) {
		warnings.push('Missing META-INF/container.xml');
	}

	return {
		json: {
			valid: errors.length === 0,
			errors,
			warnings,
			fileCount: fileList.filter((f) => !zip.files[f].dir).length,
			sections: fileList.filter((f) => f.startsWith('Contents/section') && f.endsWith('.xml')),
		},
		pairedItem: { item: itemIndex },
	};
}

/**
 * Convert HWPX document to HTML using HwpxReader.
 */
export async function toHtml(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		renderImages?: boolean;
		renderTables?: boolean;
		renderStyles?: boolean;
		embedImages?: boolean;
		paragraphTag?: string;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const reader = new HwpxReader();
	await reader.loadFromArrayBuffer(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer);

	const html = await reader.extractHtml({
		renderImages: options.renderImages ?? true,
		renderTables: options.renderTables ?? true,
		renderStyles: options.renderStyles ?? true,
		embedImages: options.embedImages ?? false,
		paragraphTag: options.paragraphTag ?? 'p',
	});

	return {
		json: {
			html,
			htmlLength: html.length,
		},
		pairedItem: { item: itemIndex },
	};
}

/**
 * Convert HWP binary to HWPX using HWPReader (pure TypeScript).
 */
export async function convertHwp(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const outputBinaryPropertyName = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const fileName = this.getNodeParameter('fileName', itemIndex, 'converted.hwpx') as string;

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	// Write to temp file, convert, read back (HwpConverter is file-path based)
	const os = await import('os');
	const fs = await import('fs/promises');
	const path = await import('path');
	const tmpDir = os.tmpdir();
	const inputPath = path.join(tmpDir, `hwp-input-${Date.now()}.hwp`);
	const outputPath = path.join(tmpDir, `hwpx-output-${Date.now()}.hwpx`);

	try {
		await fs.writeFile(inputPath, buffer);
		const converter = new HwpConverter();
		const result = await converter.convertHwpToHwpx(inputPath, outputPath);

		if (!result.success) {
			throw new Error(result.error ?? 'HWP to HWPX conversion failed');
		}

		const outputBuffer = await fs.readFile(outputPath);
		const newBinaryData = await this.helpers.prepareBinaryData(
			outputBuffer,
			fileName,
			'application/hwp+zip',
		);

		return {
			json: {
				fileName,
				size: outputBuffer.length,
				convertedFrom: binaryData.fileName ?? 'unknown.hwp',
			},
			binary: { [outputBinaryPropertyName]: newBinaryData },
			pairedItem: { item: itemIndex },
		};
	} finally {
		const fs2 = await import('fs/promises');
		await fs2.unlink(inputPath).catch(() => {});
		await fs2.unlink(outputPath).catch(() => {});
	}
}
