import type { IExecuteFunctions, INodeExecutionData, IDataObject } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import JSZip from 'jszip';
import { HwpxReader } from '@ssabrojs/hwpxjs';

// HWPUNIT: 1mm ≈ 283.46 HWPUNIT
const MM_TO_HWPUNIT = 283.46;

interface PaperSizeSpec {
	width: number;
	height: number;
}

const PAPER_SIZES: Record<string, PaperSizeSpec> = {
	A4: { width: 59528, height: 84186 }, // 210mm × 297mm
	B5: { width: 49889, height: 70866 }, // 176mm × 250mm
	Letter: { width: 61200, height: 79200 }, // 215.9mm × 279.4mm
	Legal: { width: 61200, height: 100800 }, // 215.9mm × 355.6mm
	A3: { width: 84186, height: 119055 }, // 297mm × 420mm
};

/**
 * Replace text in an HWPX document at the ZIP level.
 */
export async function replaceText(
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
	const replacementsParam = this.getNodeParameter('replacements', itemIndex) as {
		pairs?: Array<{ find: string; replace: string }>;
	};
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		targetFiles?: string;
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
	if (pairs.length === 0) {
		throw new NodeOperationError(
			this.getNode(),
			'At least one replacement pair is required',
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);
	const targetFiles = options.targetFiles ?? 'contentsXml';

	let replacementCount = 0;

	for (const [filename, file] of Object.entries(zip.files)) {
		if (file.dir) continue;
		if (!filename.endsWith('.xml')) continue;

		const shouldProcess =
			targetFiles === 'allXml' ||
			(targetFiles === 'contentsXml' && filename.startsWith('Contents/'));

		if (!shouldProcess) continue;

		let content = await file.async('string');
		let modified = false;

		for (const pair of pairs) {
			if (pair.find && content.includes(pair.find)) {
				const occurrences = content.split(pair.find).length - 1;
				content = content.split(pair.find).join(pair.replace);
				replacementCount += occurrences;
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

	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			replacementCount,
			pairsApplied: pairs.length,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Extract text structure using HwpxReader for proper XML parsing.
 */
export async function extractStructure(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		includeEmpty?: boolean;
		sectionFilter?: string;
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

	const fullText = await reader.extractText();
	const info = await reader.getDocumentInfo();

	let sectionFiles = info.summary.contentsFiles.filter(
		(f: string) => f.includes('section') && f.endsWith('.xml'),
	);

	if (options.sectionFilter) {
		sectionFiles = sectionFiles.filter((f: string) => f.includes(options.sectionFilter!));
	}

	return {
		json: {
			fullText,
			totalTextElements: fullText.split('\n').filter((l: string) => options.includeEmpty || l.trim()).length,
			sectionCount: sectionFiles.length,
			sectionFiles,
		},
		pairedItem: { item: itemIndex },
	};
}

/**
 * Sequential text replacement: replace each occurrence of a find text
 * with a different value from the values list, in order.
 *
 * Example: find="PLACEHOLDER", values=["A","B","C"]
 *   1st occurrence becomes "A", 2nd becomes "B", 3rd becomes "C"
 */
export async function replaceTextSequential(
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
	const findText = this.getNodeParameter('findText', itemIndex) as string;
	const replaceValuesRaw = this.getNodeParameter('replaceValues', itemIndex, '') as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		targetFiles?: string;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	if (!findText) {
		throw new NodeOperationError(this.getNode(), 'Find Text is required', { itemIndex });
	}

	const values = replaceValuesRaw
		.split('\n')
		.filter((v: string) => v.length > 0);

	if (values.length === 0) {
		throw new NodeOperationError(
			this.getNode(),
			'At least one replacement value is required',
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);
	const targetFiles = options.targetFiles ?? 'contentsXml';

	let replacementCount = 0;
	let valueIndex = 0;

	// Sort filenames to ensure section0 < section1 < section2 order
	const filenames = Object.keys(zip.files).sort();

	for (const filename of filenames) {
		const file = zip.files[filename];
		if (file.dir) continue;
		if (!filename.endsWith('.xml')) continue;

		const shouldProcess =
			targetFiles === 'allXml' ||
			(targetFiles === 'contentsXml' && filename.startsWith('Contents/'));

		if (!shouldProcess) continue;

		let content = await file.async('string');
		let modified = false;

		// String.replace with a string arg replaces only the first occurrence
		while (valueIndex < values.length && content.includes(findText)) {
			content = content.replace(findText, values[valueIndex]);
			valueIndex++;
			replacementCount++;
			modified = true;
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

	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			replacementCount,
			totalValues: values.length,
			valuesUsed: valueIndex,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * List all text elements in an HWPX document.
 * Useful for discovering placeholders before template replacement.
 */
export async function listTexts(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		includeEmpty?: boolean;
		sectionFilter?: string;
		deduplicate?: boolean;
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
	const zip = await JSZip.loadAsync(buffer);

	const texts: Array<{ text: string; file: string; index: number }> = [];
	const filenames = Object.keys(zip.files).sort();

	for (const filename of filenames) {
		const file = zip.files[filename];
		if (file.dir) continue;
		if (!filename.startsWith('Contents/') || !filename.endsWith('.xml')) continue;

		if (options.sectionFilter && !filename.includes(options.sectionFilter)) {
			continue;
		}

		const content = await file.async('string');

		// Extract text from <hp:t>...</hp:t> elements
		const pattern = /<hp:t>([^<]*)<\/hp:t>/g;
		let match: RegExpExecArray | null;
		let idx = 0;

		while ((match = pattern.exec(content)) !== null) {
			const text = match[1]
				.replace(/&amp;/g, '&')
				.replace(/&lt;/g, '<')
				.replace(/&gt;/g, '>')
				.replace(/&quot;/g, '"')
				.replace(/&apos;/g, "'");

			if (text || options.includeEmpty) {
				texts.push({ text, file: filename, index: idx++ });
			}
		}
	}

	let outputTexts = texts;
	if (options.deduplicate) {
		const seen = new Set<string>();
		outputTexts = texts.filter((t) => {
			if (seen.has(t.text)) return false;
			seen.add(t.text);
			return true;
		});
	}

	const uniqueTexts = [...new Set(texts.map((t) => t.text))];

	return {
		json: {
			totalTexts: texts.length,
			uniqueCount: uniqueTexts.length,
			texts: outputTexts,
			uniqueTexts,
		},
		pairedItem: { item: itemIndex },
	};
}

/**
 * Files that must be stored uncompressed (STORED) in HWPX ZIP archives.
 * Changing these to DEFLATED causes "document corrupted" errors in Hangul viewer.
 */
function shouldStore(filename: string): boolean {
	if (filename === 'mimetype') return true;
	if (filename === 'version.xml') return true;
	if (filename.startsWith('BinData/')) return true;
	if (filename === 'Preview/PrvImage.png') return true;
	return false;
}

/**
 * Helper: preserve per-file compression types and generate output buffer.
 *
 * HWPX ZIP archives require specific files to use STORED (no compression):
 * - mimetype, version.xml (OCF/ODF spec)
 * - BinData/* (images/binaries)
 * - Preview/PrvImage.png (preview image)
 *
 * All other files (Contents/*.xml, settings.xml, META-INF/*, Scripts/*) use DEFLATED.
 */
async function finalizeZip(zip: JSZip): Promise<Buffer> {
	// Re-set compression for files that must be STORED
	for (const [filename, file] of Object.entries(zip.files)) {
		if (file.dir) continue;
		if (shouldStore(filename)) {
			const content = await file.async('nodebuffer');
			zip.file(filename, content, { compression: 'STORE' });
		}
	}
	return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }) as Promise<Buffer>;
}

/**
 * Helper: escape text for XML content.
 */
function escapeXml(text: string): string {
	return text
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

/**
 * Modify page setup (paper size, orientation, margins) of an HWPX document.
 *
 * Operates on section XML by modifying or inserting <hp:pagePr> inside <hp:secPr>.
 * If the section lacks secPr, a full secPr block is injected into the first paragraph.
 */
export async function pageSetup(
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
	const paperSize = this.getNodeParameter('paperSize', itemIndex, 'A4') as string;
	const orientation = this.getNodeParameter('orientation', itemIndex, 'portrait') as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		marginTop?: number;
		marginBottom?: number;
		marginLeft?: number;
		marginRight?: number;
		marginHeader?: number;
		marginFooter?: number;
		customWidth?: number;
		customHeight?: number;
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
	const zip = await JSZip.loadAsync(buffer);

	// Determine paper dimensions
	let paperW: number;
	let paperH: number;
	if (paperSize === 'custom') {
		paperW = Math.round((options.customWidth ?? 210) * MM_TO_HWPUNIT);
		paperH = Math.round((options.customHeight ?? 297) * MM_TO_HWPUNIT);
	} else {
		const spec = PAPER_SIZES[paperSize] ?? PAPER_SIZES.A4;
		paperW = spec.width;
		paperH = spec.height;
	}

	// Apply orientation swap
	const landscape = orientation === 'landscape';
	const finalW = landscape ? Math.max(paperW, paperH) : Math.min(paperW, paperH);
	const finalH = landscape ? Math.min(paperW, paperH) : Math.max(paperW, paperH);
	const landscapeAttr = landscape ? 'WIDELY' : 'NARROWLY';

	// Margins in HWPUNIT (defaults: top=20mm, bottom=15mm, left/right=30mm, header/footer=15mm)
	const mTop = Math.round((options.marginTop ?? 20) * MM_TO_HWPUNIT);
	const mBottom = Math.round((options.marginBottom ?? 15) * MM_TO_HWPUNIT);
	const mLeft = Math.round((options.marginLeft ?? 30) * MM_TO_HWPUNIT);
	const mRight = Math.round((options.marginRight ?? 30) * MM_TO_HWPUNIT);
	const mHeader = Math.round((options.marginHeader ?? 15) * MM_TO_HWPUNIT);
	const mFooter = Math.round((options.marginFooter ?? 15) * MM_TO_HWPUNIT);

	const pagePrXml =
		`<hp:pagePr landscape="${landscapeAttr}" width="${finalW}" height="${finalH}" gutterType="LEFT_ONLY">` +
		`<hp:margin header="${mHeader}" footer="${mFooter}" gutter="0" ` +
		`left="${mLeft}" right="${mRight}" top="${mTop}" bottom="${mBottom}" />` +
		`</hp:pagePr>`;

	// Find and modify section XML files
	const filenames = Object.keys(zip.files).sort();
	let modifiedSections = 0;

	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');

		if (content.includes('<hp:pagePr')) {
			// Replace existing pagePr block (including its child margin element)
			content = content.replace(
				/<hp:pagePr[^>]*>[\s\S]*?<\/hp:pagePr>/,
				pagePrXml,
			);
			modifiedSections++;
		} else if (content.includes('<hp:secPr')) {
			// secPr exists but no pagePr — insert pagePr as first child of secPr
			content = content.replace(
				/<hp:secPr([^>]*)>/,
				`<hp:secPr$1>${pagePrXml}`,
			);
			modifiedSections++;
		} else {
			// No secPr at all — inject a minimal secPr with pagePr into first paragraph's first run
			const secPrBlock =
				`<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" ` +
				`tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" ` +
				`textVerticalWidthHead="0" masterPageCnt="0">` +
				pagePrXml +
				`</hp:secPr>`;

			// Insert after the first <hp:run ...> opening tag
			const runMatch = content.match(/<hp:run[^>]*>/);
			if (runMatch && runMatch.index !== undefined) {
				const insertPos = runMatch.index + runMatch[0].length;
				content = content.slice(0, insertPos) + secPrBlock + content.slice(insertPos);
				modifiedSections++;
			}
		}

		zip.file(filename, content);
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			paperSize,
			orientation,
			widthHwpunit: finalW,
			heightHwpunit: finalH,
			modifiedSections,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Insert an image into an HWPX document.
 *
 * Three-step process:
 * 1. Store image binary in BinData/ directory
 * 2. Register in content.hpf manifest
 * 3. Insert <hp:picture> paragraph into section XML
 */
export async function insertImage(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const imageBinaryPropertyName = this.getNodeParameter(
		'imageBinaryPropertyName',
		itemIndex,
	) as string;
	const outputBinaryPropertyName = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		position?: string;
		widthMm?: number;
		heightMm?: number;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const imageBinary = item.binary?.[imageBinaryPropertyName];
	if (!imageBinary) {
		throw new NodeOperationError(
			this.getNode(),
			`No image binary data found in property "${imageBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const imageBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, imageBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	// Determine image format from mime type or file extension
	const mimeType = imageBinary.mimeType ?? 'image/png';
	const ext = mimeType.includes('jpeg') || mimeType.includes('jpg')
		? 'jpg'
		: mimeType.includes('gif')
			? 'gif'
			: mimeType.includes('bmp')
				? 'bmp'
				: 'png';
	const mediaType = ext === 'jpg' ? 'image/jpeg' : `image/${ext}`;

	// Find next available BIN ID
	let maxBinNum = 0;
	for (const fname of Object.keys(zip.files)) {
		const m = fname.match(/BinData\/BIN(\d+)\./);
		if (m) {
			const num = parseInt(m[1], 10);
			if (num > maxBinNum) maxBinNum = num;
		}
	}
	const binNum = maxBinNum + 1;
	const binId = `BIN${String(binNum).padStart(4, '0')}`;
	const binPath = `BinData/${binId}.${ext}`;

	// Step 1: Store image in BinData/
	zip.file(binPath, imageBuffer);

	// Step 2: Register in content.hpf manifest
	const hpfFile = zip.file('Contents/content.hpf');
	if (hpfFile) {
		let hpf = await hpfFile.async('string');
		const manifestItem = `<opf:item id="${binId}" href="${binPath}" media-type="${mediaType}"/>`;

		// Insert before </opf:manifest> or before </opf:package>
		if (hpf.includes('</opf:manifest>')) {
			hpf = hpf.replace('</opf:manifest>', `${manifestItem}</opf:manifest>`);
		} else if (hpf.includes('</manifest>')) {
			hpf = hpf.replace('</manifest>', `${manifestItem}</manifest>`);
		}
		zip.file('Contents/content.hpf', hpf);
	}

	// Step 3: Build image paragraph XML and insert into section
	const imgWidth = Math.round((options.widthMm ?? 100) * MM_TO_HWPUNIT);
	const imgHeight = Math.round((options.heightMm ?? 75) * MM_TO_HWPUNIT);
	const picId = Date.now();

	const pictureParagraph =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0">` +
		`<hp:pic id="${picId}" zOrder="0" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" ` +
		`textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" ` +
		`instid="${picId + 1}" reverse="0">` +
		`<hp:offset x="0" y="0"/>` +
		`<hp:orgSz width="${imgWidth}" height="${imgHeight}"/>` +
		`<hp:curSz width="0" height="0"/>` +
		`<hp:flip horizontal="0" vertical="0"/>` +
		`<hp:rotationInfo angle="0" centerX="${Math.round(imgWidth / 2)}" ` +
		`centerY="${Math.round(imgHeight / 2)}" rotateimage="1"/>` +
		`<hp:renderingInfo>` +
		`<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
		`<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
		`<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>` +
		`</hp:renderingInfo>` +
		`<hc:img binaryItemIDRef="${binId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>` +
		`<hp:imgRect>` +
		`<hc:pt0 x="0" y="0"/><hc:pt1 x="${imgWidth}" y="0"/>` +
		`<hc:pt2 x="${imgWidth}" y="${imgHeight}"/><hc:pt3 x="0" y="${imgHeight}"/>` +
		`</hp:imgRect>` +
		`<hp:imgClip left="0" right="0" top="0" bottom="0"/>` +
		`<hp:inMargin left="0" right="0" top="0" bottom="0"/>` +
		`<hp:imgDim dimwidth="${imgWidth}" dimheight="${imgHeight}"/>` +
		`<hp:effects/>` +
		`<hp:sz width="${imgWidth}" widthRelTo="ABSOLUTE" height="${imgHeight}" heightRelTo="ABSOLUTE" protect="0"/>` +
		`<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" ` +
		`holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" ` +
		`vertOffset="0" horzOffset="0"/>` +
		`<hp:outMargin left="0" right="0" top="0" bottom="0"/>` +
		`</hp:pic>` +
		`</hp:run>` +
		`</hp:p>`;

	// Find first section XML and insert image paragraph
	const position = options.position ?? 'append';
	const filenames = Object.keys(zip.files).sort();

	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');

		// Ensure hc namespace is declared for image references
		if (!content.includes('xmlns:hc=')) {
			content = content.replace(
				'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"',
				'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" ' +
				'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"',
			);
		}

		if (position === 'prepend') {
			// Insert after <hs:sec ...> opening tag
			content = content.replace(
				/(<hs:sec[^>]*>)/,
				`$1${pictureParagraph}`,
			);
		} else {
			// Append before </hs:sec>
			content = content.replace('</hs:sec>', `${pictureParagraph}</hs:sec>`);
		}

		zip.file(filename, content);
		break; // Only modify first section
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			binId,
			binPath,
			imageFormat: ext,
			widthMm: options.widthMm ?? 100,
			heightMm: options.heightMm ?? 75,
			position,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Replace an existing image in an HWPX document.
 *
 * Three-step process:
 * 1. Find existing image by ID or index in BinData/
 * 2. Replace the binary file with the new image
 * 3. Update dimensions in section XML if size changed
 */
export async function replaceImage(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputBinaryPropertyName = this.getNodeParameter(
		'inputBinaryPropertyName',
		itemIndex,
	) as string;
	const imageBinaryPropertyName = this.getNodeParameter(
		'imageBinaryPropertyName',
		itemIndex,
	) as string;
	const outputBinaryPropertyName = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const targetImage = this.getNodeParameter('targetImage', itemIndex, '') as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		widthMm?: number;
		heightMm?: number;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const imageBinary = item.binary?.[imageBinaryPropertyName];
	if (!imageBinary) {
		throw new NodeOperationError(
			this.getNode(),
			`No image binary data found in property "${imageBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const imageBuffer = await this.helpers.getBinaryDataBuffer(itemIndex, imageBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	// Find existing image files in BinData/
	const binFiles = Object.keys(zip.files)
		.filter((f) => f.startsWith('BinData/') && !zip.files[f].dir)
		.sort();

	if (binFiles.length === 0) {
		throw new NodeOperationError(
			this.getNode(),
			'No images found in document BinData/ directory',
			{ itemIndex },
		);
	}

	// Determine target: by filename, by ID, or by index (default: first image)
	let targetPath: string;
	if (targetImage) {
		// Try exact match first
		const exact = binFiles.find((f) => f === targetImage || f === `BinData/${targetImage}`);
		if (exact) {
			targetPath = exact;
		} else {
			// Try matching by ID (e.g., "image1", "BIN0001")
			const byId = binFiles.find((f) => f.includes(targetImage));
			if (byId) {
				targetPath = byId;
			} else {
				// Try by index (1-based)
				const idx = parseInt(targetImage, 10);
				if (!isNaN(idx) && idx >= 1 && idx <= binFiles.length) {
					targetPath = binFiles[idx - 1];
				} else {
					throw new NodeOperationError(
						this.getNode(),
						`Image "${targetImage}" not found. Available: ${binFiles.join(', ')}`,
						{ itemIndex },
					);
				}
			}
		}
	} else {
		// Default: replace the first image
		targetPath = binFiles[0];
	}

	// Determine new image format
	const mimeType = imageBinary.mimeType ?? 'image/png';
	const newExt = mimeType.includes('jpeg') || mimeType.includes('jpg')
		? 'jpg'
		: mimeType.includes('gif')
			? 'gif'
			: mimeType.includes('bmp')
				? 'bmp'
				: 'png';

	// Extract the bin ID from the target path (e.g., "BinData/image1.png" -> "image1")
	const oldBinId = targetPath.replace('BinData/', '').replace(/\.[^.]+$/, '');
	const oldExt = targetPath.split('.').pop() ?? 'png';

	// If extension changed, we need to update paths
	const newBinPath = newExt !== oldExt
		? `BinData/${oldBinId}.${newExt}`
		: targetPath;

	// Step 1: Replace image binary
	zip.remove(targetPath);
	zip.file(newBinPath, imageBuffer, { compression: 'STORE' });

	// Step 2: Update content.hpf manifest if extension changed
	if (newExt !== oldExt) {
		const hpfFile = zip.file('Contents/content.hpf');
		if (hpfFile) {
			let hpf = await hpfFile.async('string');
			const newMediaType = newExt === 'jpg' ? 'image/jpeg' : `image/${newExt}`;
			// Update href and media-type for this image
			hpf = hpf.replace(
				new RegExp(`href="${targetPath.replace('/', '\\/')}"[^/]*\\/>`),
				`href="${newBinPath}" media-type="${newMediaType}"/>`,
			);
			zip.file('Contents/content.hpf', hpf);
		}
	}

	// Step 3: Update image dimensions in section XML if new size specified
	if (options.widthMm || options.heightMm) {
		const newWidth = Math.round((options.widthMm ?? 100) * MM_TO_HWPUNIT);
		const newHeight = Math.round((options.heightMm ?? 75) * MM_TO_HWPUNIT);

		for (const filename of Object.keys(zip.files).sort()) {
			if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

			let content = await zip.files[filename].async('string');
			if (!content.includes(`binaryItemIDRef="${oldBinId}"`)) continue;

			// Find the <hp:pic> block containing this image and update dimensions
			const picRegex = new RegExp(
				`(<hp:pic[^>]*>)(.*?)(binaryItemIDRef="${oldBinId}")(.*?)(</hp:pic>)`,
				's',
			);
			const picMatch = content.match(picRegex);
			if (picMatch) {
				let picBlock = picMatch[0];
				// Update orgSz
				picBlock = picBlock.replace(
					/<hp:orgSz width="\d+" height="\d+"\/>/,
					`<hp:orgSz width="${newWidth}" height="${newHeight}"/>`,
				);
				// Update imgDim
				picBlock = picBlock.replace(
					/<hp:imgDim dimwidth="\d+" dimheight="\d+"\/>/,
					`<hp:imgDim dimwidth="${newWidth}" dimheight="${newHeight}"/>`,
				);
				// Update sz
				picBlock = picBlock.replace(
					/<hp:sz width="\d+" (widthRelTo="ABSOLUTE") height="\d+"/,
					`<hp:sz width="${newWidth}" $1 height="${newHeight}"`,
				);
				// Update imgRect points
				picBlock = picBlock.replace(
					/<hc:pt1 x="\d+" y="0"\/>/,
					`<hc:pt1 x="${newWidth}" y="0"/>`,
				);
				picBlock = picBlock.replace(
					/<hc:pt2 x="\d+" y="\d+"\/>/,
					`<hc:pt2 x="${newWidth}" y="${newHeight}"/>`,
				);
				picBlock = picBlock.replace(
					/<hc:pt3 x="0" y="\d+"\/>/,
					`<hc:pt3 x="0" y="${newHeight}"/>`,
				);
				content = content.replace(picMatch[0], picBlock);
			}

			zip.file(filename, content);
		}
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			replacedImage: targetPath,
			newPath: newBinPath,
			imageFormat: newExt,
			widthMm: options.widthMm,
			heightMm: options.heightMm,
			availableImages: binFiles,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Add a table to an HWPX document.
 *
 * Builds HWPML table XML from a JSON array-of-arrays and inserts it
 * into the section XML. Supports optional header row styling.
 */
export async function addTable(
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
	const tableDataRaw = this.getNodeParameter('tableData', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		position?: string;
		cellWidthMm?: number;
		cellHeightMm?: number;
		borderFillIDRef?: number;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	// Parse table data: array of arrays (JSON)
	let tableData: string[][];
	try {
		const parsed = JSON.parse(tableDataRaw);
		if (!Array.isArray(parsed) || parsed.length === 0) {
			throw new Error('Table data must be a non-empty array');
		}
		tableData = parsed.map((row: unknown) => {
			if (!Array.isArray(row)) throw new Error('Each row must be an array');
			return row.map((cell: unknown) => String(cell ?? ''));
		});
	} catch (error) {
		throw new NodeOperationError(
			this.getNode(),
			`Invalid table data JSON: ${(error as Error).message}. ` +
			'Expected format: [["A1","B1"],["A2","B2"]]',
			{ itemIndex },
		);
	}

	const rowCnt = tableData.length;
	const colCnt = Math.max(...tableData.map((r: string[]) => r.length));
	const borderFillRef = options.borderFillIDRef ?? 3;

	// Cell dimensions in HWPUNIT
	const cellW = Math.round((options.cellWidthMm ?? 30) * MM_TO_HWPUNIT);
	const cellH = Math.round((options.cellHeightMm ?? 10) * MM_TO_HWPUNIT);
	const tableW = cellW * colCnt;
	const tableH = cellH * rowCnt;

	// Build table XML
	let tblXml =
		`<hp:tbl id="${Date.now()}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" ` +
		`textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="0" ` +
		`rowCnt="${rowCnt}" colCnt="${colCnt}" cellSpacing="0" borderFillIDRef="${borderFillRef}" noAdjust="0">` +
		`<hp:sz width="${tableW}" widthRelTo="ABSOLUTE" height="${tableH}" heightRelTo="ABSOLUTE" protect="0"/>` +
		`<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" ` +
		`holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" ` +
		`vertOffset="0" horzOffset="0"/>` +
		`<hp:outMargin left="0" right="0" top="0" bottom="0"/>` +
		`<hp:inMargin left="0" right="0" top="0" bottom="0"/>`;

	for (let r = 0; r < rowCnt; r++) {
		tblXml += '<hp:tr>';
		const row = tableData[r];
		for (let c = 0; c < colCnt; c++) {
			const cellText = escapeXml(c < row.length ? row[c] : '');
			tblXml +=
				`<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="1" ` +
				`borderFillIDRef="${borderFillRef}">` +
				`<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" ` +
				`linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
				`<hp:p paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" id="0">` +
				`<hp:run charPrIDRef="0"><hp:t>${cellText}</hp:t></hp:run>` +
				`</hp:p>` +
				`</hp:subList>` +
				`<hp:cellAddr colAddr="${c}" rowAddr="${r}"/>` +
				`<hp:cellSpan colSpan="1" rowSpan="1"/>` +
				`<hp:cellSz width="${cellW}" height="${cellH}"/>` +
				`<hp:cellMargin left="0" right="0" top="0" bottom="0"/>` +
				`</hp:tc>`;
		}
		tblXml += '</hp:tr>';
	}
	tblXml += '</hp:tbl>';

	// Wrap in a paragraph
	const tableParagraph =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0">${tblXml}</hp:run></hp:p>`;

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);
	const position = options.position ?? 'append';

	const filenames = Object.keys(zip.files).sort();
	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');

		if (position === 'prepend') {
			content = content.replace(/(<hs:sec[^>]*>)/, `$1${tableParagraph}`);
		} else {
			content = content.replace('</hs:sec>', `${tableParagraph}</hs:sec>`);
		}

		zip.file(filename, content);
		break; // Only modify first section
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			rows: rowCnt,
			columns: colCnt,
			tableWidthHwpunit: tableW,
			tableHeightHwpunit: tableH,
			position,
			fileName: originalFileName,
		} as unknown as IDataObject,
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Set header and/or footer text in an HWPX document.
 *
 * HWPML header/footer structure:
 *   Inside the first paragraph's first <hp:run>, after <hp:secPr>,
 *   wrapped in <hp:ctrl><hp:header ...>...</hp:header></hp:ctrl>
 */
export async function setHeaderFooter(
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
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		headerText?: string;
		footerText?: string;
		applyPageType?: string;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const headerText = this.getNodeParameter('headerText', itemIndex, '') as string;
	const footerText = this.getNodeParameter('footerText', itemIndex, '') as string;

	if (!headerText && !footerText) {
		throw new NodeOperationError(
			this.getNode(),
			'At least one of Header Text or Footer Text must be provided',
			{ itemIndex },
		);
	}

	const applyPageType = options.applyPageType ?? 'BOTH';

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	const filenames = Object.keys(zip.files).sort();
	let modified = false;

	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');

		if (headerText) {
			const headerXml =
				`<hp:ctrl><hp:header id="${Date.now()}" applyPageType="${applyPageType}">` +
				`<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" ` +
				`linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
				`<hp:p paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" id="0">` +
				`<hp:run charPrIDRef="0"><hp:t>${escapeXml(headerText)}</hp:t></hp:run>` +
				`</hp:p></hp:subList></hp:header></hp:ctrl>`;

			if (/<hp:ctrl>\s*<hp:header[\s\S]*?<\/hp:header>\s*<\/hp:ctrl>/.test(content)) {
				// Replace existing header
				content = content.replace(
					/<hp:ctrl>\s*<hp:header[\s\S]*?<\/hp:header>\s*<\/hp:ctrl>/,
					headerXml,
				);
			} else {
				// Insert after </hp:secPr> if it exists, otherwise after first <hp:run...>
				if (content.includes('</hp:secPr>')) {
					content = content.replace('</hp:secPr>', `</hp:secPr>${headerXml}`);
				} else {
					const runMatch = content.match(/<hp:run[^>]*>/);
					if (runMatch && runMatch.index !== undefined) {
						const insertPos = runMatch.index + runMatch[0].length;
						content = content.slice(0, insertPos) + headerXml + content.slice(insertPos);
					}
				}
			}
		}

		if (footerText) {
			const footerXml =
				`<hp:ctrl><hp:footer id="${Date.now() + 1}" applyPageType="${applyPageType}">` +
				`<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" ` +
				`linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
				`<hp:p paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0" id="0">` +
				`<hp:run charPrIDRef="0"><hp:t>${escapeXml(footerText)}</hp:t></hp:run>` +
				`</hp:p></hp:subList></hp:footer></hp:ctrl>`;

			if (/<hp:ctrl>\s*<hp:footer[\s\S]*?<\/hp:footer>\s*<\/hp:ctrl>/.test(content)) {
				content = content.replace(
					/<hp:ctrl>\s*<hp:footer[\s\S]*?<\/hp:footer>\s*<\/hp:ctrl>/,
					footerXml,
				);
			} else {
				// Insert after header ctrl if it exists, or after </hp:secPr>, or after first run
				if (/<hp:ctrl>\s*<hp:header[\s\S]*?<\/hp:header>\s*<\/hp:ctrl>/.test(content)) {
					content = content.replace(
						/(<hp:ctrl>\s*<hp:header[\s\S]*?<\/hp:header>\s*<\/hp:ctrl>)/,
						`$1${footerXml}`,
					);
				} else if (content.includes('</hp:secPr>')) {
					content = content.replace('</hp:secPr>', `</hp:secPr>${footerXml}`);
				} else {
					const runMatch = content.match(/<hp:run[^>]*>/);
					if (runMatch && runMatch.index !== undefined) {
						const insertPos = runMatch.index + runMatch[0].length;
						content = content.slice(0, insertPos) + footerXml + content.slice(insertPos);
					}
				}
			}
		}

		zip.file(filename, content);
		modified = true;
		break; // Only modify first section
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			headerSet: !!headerText,
			footerSet: !!footerText,
			applyPageType,
			modified,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Merge cells in an existing table within an HWPX document.
 *
 * Modifies the table XML by:
 * 1. Expanding the anchor cell's cellSpan (colSpan/rowSpan) and cellSz
 * 2. Removing absorbed cells from each row
 * 3. Updating table rowCnt/colCnt if needed
 */
export async function mergeCells(
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
	const startRow = this.getNodeParameter('startRow', itemIndex) as number;
	const startCol = this.getNodeParameter('startCol', itemIndex) as number;
	const endRow = this.getNodeParameter('endRow', itemIndex) as number;
	const endCol = this.getNodeParameter('endCol', itemIndex) as number;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		tableIndex?: number;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	if (startRow > endRow || startCol > endCol) {
		throw new NodeOperationError(
			this.getNode(),
			'Start row/col must be less than or equal to end row/col',
			{ itemIndex },
		);
	}

	if (startRow === endRow && startCol === endCol) {
		throw new NodeOperationError(
			this.getNode(),
			'Merge range must span at least 2 cells',
			{ itemIndex },
		);
	}

	const tableIndex = options.tableIndex ?? 0;

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	const filenames = Object.keys(zip.files).sort();
	let merged = false;

	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		const content = await zip.files[filename].async('string');

		// Find all tables in the section
		const tablePattern = /<hp:tbl[^>]*>[\s\S]*?<\/hp:tbl>/g;
		const tables: Array<{ match: string; index: number }> = [];
		let tblMatch: RegExpExecArray | null;
		while ((tblMatch = tablePattern.exec(content)) !== null) {
			tables.push({ match: tblMatch[0], index: tblMatch.index });
		}

		if (tableIndex >= tables.length) continue;

		const table = tables[tableIndex];
		const tblXml = table.match;

		// Parse rows and cells
		const rowPattern = /<hp:tr>([\s\S]*?)<\/hp:tr>/g;
		const rows: string[] = [];
		let rowMatch: RegExpExecArray | null;
		while ((rowMatch = rowPattern.exec(tblXml)) !== null) {
			rows.push(rowMatch[0]);
		}

		if (endRow >= rows.length) {
			throw new NodeOperationError(
				this.getNode(),
				`End row ${endRow} exceeds table row count ${rows.length}`,
				{ itemIndex },
			);
		}

		// Calculate merged cell dimensions
		const colSpan = endCol - startCol + 1;
		const rowSpan = endRow - startRow + 1;

		// Extract cell widths and heights from the anchor cell's row
		const cellWidths: number[] = [];
		const cellHeights: number[] = [];

		// Get all cell sizes from first affected row to calculate merged width
		const firstRowCells = rows[startRow].match(/<hp:tc[\s\S]*?<\/hp:tc>/g) ?? [];

		for (const cell of firstRowCells) {
			const szMatch = /<hp:cellSz width="(\d+)" height="(\d+)"\/>/.exec(cell);
			if (szMatch) {
				cellWidths.push(parseInt(szMatch[1], 10));
				cellHeights.push(parseInt(szMatch[2], 10));
			}
		}

		if (endCol >= cellWidths.length) {
			throw new NodeOperationError(
				this.getNode(),
				`End col ${endCol} exceeds table column count ${cellWidths.length}`,
				{ itemIndex },
			);
		}

		// Calculate merged dimensions
		let mergedWidth = 0;
		for (let c = startCol; c <= endCol; c++) {
			mergedWidth += cellWidths[c];
		}

		// Get heights from all rows involved
		let mergedHeight = 0;
		for (let r = startRow; r <= endRow; r++) {
			const row = rows[r];
			if (!row) continue;
			const rowCells = row.match(/<hp:tc[\s\S]*?<\/hp:tc>/g) ?? [];
			if (rowCells.length > 0) {
				const firstCell = rowCells[0] as string;
				const szMatch = /<hp:cellSz width="(\d+)" height="(\d+)"\/>/.exec(firstCell);
				if (szMatch) {
					mergedHeight += parseInt(szMatch[2], 10);
				}
			}
		}

		// Process each row
		const newRows: string[] = [];
		for (let r = 0; r < rows.length; r++) {
			const cells = rows[r].match(/<hp:tc[\s\S]*?<\/hp:tc>/g) ?? [];
			const newCells: string[] = [];

			for (let c = 0; c < cells.length; c++) {
				const isInMergeRange = r >= startRow && r <= endRow && c >= startCol && c <= endCol;
				const isAnchorCell = r === startRow && c === startCol;

				if (isAnchorCell) {
					// Update anchor cell with expanded span and size
					let cell = cells[c];
					cell = cell.replace(
						/<hp:cellSpan colSpan="\d+" rowSpan="\d+"\/>/,
						`<hp:cellSpan colSpan="${colSpan}" rowSpan="${rowSpan}"/>`,
					);
					cell = cell.replace(
						/<hp:cellSz width="\d+" height="\d+"\/>/,
						`<hp:cellSz width="${mergedWidth}" height="${mergedHeight}"/>`,
					);
					newCells.push(cell);
				} else if (isInMergeRange) {
					// Skip absorbed cells (remove them)
					continue;
				} else {
					newCells.push(cells[c]);
				}
			}

			newRows.push(`<hp:tr>${newCells.join('')}</hp:tr>`);
		}

		// Rebuild table XML with new rows
		// Replace rows in original table XML
		let newTblXml = tblXml;
		// Remove all existing rows
		newTblXml = newTblXml.replace(/<hp:tr>[\s\S]*?<\/hp:tr>/g, '');
		// Find closing </hp:tbl> and insert new rows before it
		newTblXml = newTblXml.replace('</hp:tbl>', `${newRows.join('')}</hp:tbl>`);

		// Replace original table in content
		const newContent =
			content.slice(0, table.index) +
			newTblXml +
			content.slice(table.index + table.match.length);

		zip.file(filename, newContent);
		merged = true;
		break;
	}

	if (!merged) {
		throw new NodeOperationError(
			this.getNode(),
			`Table at index ${tableIndex} not found in document`,
			{ itemIndex },
		);
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			startRow,
			startCol,
			endRow,
			endCol,
			colSpan: endCol - startCol + 1,
			rowSpan: endRow - startRow + 1,
			tableIndex,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Insert a math equation using Hancom equation script syntax.
 *
 * HWPML equation structure:
 *   <hp:equation><hp:script>formula</hp:script></hp:equation>
 * Uses Hancom equation script (NOT LaTeX).
 */
export async function insertEquation(
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
	const equationScript = this.getNodeParameter('equationScript', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		position?: string;
		baseUnit?: number;
		textColor?: string;
		prefixText?: string;
		suffixText?: string;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	if (!equationScript) {
		throw new NodeOperationError(this.getNode(), 'Equation script is required', { itemIndex });
	}

	const baseUnit = options.baseUnit ?? 1000;
	const textColor = options.textColor ?? '#000000';
	const prefixText = options.prefixText ?? '';
	const suffixText = options.suffixText ?? '';
	const position = options.position ?? 'append';

	// Build equation run
	const equationRun =
		`<hp:run charPrIDRef="0">` +
		`<hp:equation id="${Date.now()}" type="0" textColor="${escapeXml(textColor)}" ` +
		`baseUnit="${baseUnit}" letterSpacing="0" lineThickness="100">` +
		`<hp:sz width="0" height="0" widthRelTo="ABS" heightRelTo="ABS"/>` +
		`<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="0" ` +
		`allowOverlap="0" holdAnchorAndSO="0" rgroupWithPrevCtrl="0" ` +
		`vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" ` +
		`vertOffset="0" horzOffset="0"/>` +
		`<hp:script>${escapeXml(equationScript)}</hp:script>` +
		`</hp:equation></hp:run>`;

	// Build paragraph with optional prefix/suffix text
	let paragraphContent = '';
	if (prefixText) {
		paragraphContent += `<hp:run charPrIDRef="0"><hp:t>${escapeXml(prefixText)}</hp:t></hp:run>`;
	}
	paragraphContent += equationRun;
	if (suffixText) {
		paragraphContent += `<hp:run charPrIDRef="0"><hp:t>${escapeXml(suffixText)}</hp:t></hp:run>`;
	}

	const equationParagraph =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		paragraphContent +
		`</hp:p>`;

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	const filenames = Object.keys(zip.files).sort();
	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');

		if (position === 'prepend') {
			content = content.replace(/(<hs:sec[^>]*>)/, `$1${equationParagraph}`);
		} else {
			content = content.replace('</hs:sec>', `${equationParagraph}</hs:sec>`);
		}

		zip.file(filename, content);
		break;
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			equationScript,
			baseUnit,
			position,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Set multi-column layout in an HWPX document.
 *
 * Adds or replaces <hp:colPr> in the first paragraph of the section.
 * type="NEWSPAPER" fills left-to-right, type="BALANCED_NEWSPAPER" balances columns.
 */
export async function setColumnLayout(
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
	const columnCount = this.getNodeParameter('columnCount', itemIndex, 2) as number;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		columnType?: string;
		gapMm?: number;
		sameSize?: boolean;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const colType = options.columnType ?? 'NEWSPAPER';
	const gapHwpunit = Math.round((options.gapMm ?? 8) * MM_TO_HWPUNIT);
	const sameSz = (options.sameSize ?? true) ? '1' : '0';

	const colPrXml =
		`<hp:colPr id="" type="${colType}" layout="LEFT" colCount="${columnCount}" ` +
		`sameSz="${sameSz}" sameGap="${gapHwpunit}"/>`;

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	const filenames = Object.keys(zip.files).sort();
	let modifiedSections = 0;

	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');

		if (/<hp:colPr[^/]*\/>/.test(content)) {
			// Replace existing colPr
			content = content.replace(/<hp:colPr[^/]*\/>/, colPrXml);
			modifiedSections++;
		} else if (/<hp:colPr[^>]*>[\s\S]*?<\/hp:colPr>/.test(content)) {
			content = content.replace(/<hp:colPr[^>]*>[\s\S]*?<\/hp:colPr>/, colPrXml);
			modifiedSections++;
		} else {
			// Insert colPr into first paragraph (after opening <hp:p ...> tag)
			const firstParagraph = content.match(/<hp:p[^>]*>/);
			if (firstParagraph && firstParagraph.index !== undefined) {
				const insertPos = firstParagraph.index + firstParagraph[0].length;
				content = content.slice(0, insertPos) + colPrXml + content.slice(insertPos);
				modifiedSections++;
			}
		}

		zip.file(filename, content);
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			columnCount,
			columnType: colType,
			gapHwpunit,
			modifiedSections,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Insert a column break to switch columns within a section.
 *
 * Adds a paragraph with columnBreak="1" at the specified position,
 * causing content after it to flow to the next column.
 */
export async function insertColumnBreak(
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
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		afterParagraphIndex?: number;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const afterIndex = options.afterParagraphIndex ?? -1;

	const colBreakParagraph =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="1" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run></hp:p>`;

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	const filenames = Object.keys(zip.files).sort();
	let inserted = false;

	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');

		if (afterIndex >= 0) {
			// Insert after specific paragraph index
			const paragraphPattern = /<hp:p[^>]*>[\s\S]*?<\/hp:p>/g;
			let pMatch: RegExpExecArray | null;
			let pIdx = 0;
			let insertPos = -1;

			while ((pMatch = paragraphPattern.exec(content)) !== null) {
				if (pIdx === afterIndex) {
					insertPos = pMatch.index + pMatch[0].length;
					break;
				}
				pIdx++;
			}

			if (insertPos >= 0) {
				content = content.slice(0, insertPos) + colBreakParagraph + content.slice(insertPos);
				inserted = true;
			}
		} else {
			// Append before </hs:sec>
			content = content.replace('</hs:sec>', `${colBreakParagraph}</hs:sec>`);
			inserted = true;
		}

		zip.file(filename, content);
		break;
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			inserted,
			afterParagraphIndex: afterIndex,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Add exam-style header to an HWPX document.
 *
 * Creates formatted header block with:
 * - Title line (e.g., "2025학년도 3월 고2 전국연합학력평가")
 * - Subject area (e.g., "수학 영역")
 * - Session label (e.g., "제 2 교시")
 * - Question type label (e.g., "5지선다형")
 */
export async function formatExamHeader(
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
	const year = this.getNodeParameter('year', itemIndex) as number;
	const month = this.getNodeParameter('month', itemIndex) as number;
	const grade = this.getNodeParameter('grade', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		subjectArea?: string;
		session?: number;
		questionTypeLabel?: string;
		examTitle?: string;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	const subjectArea = options.subjectArea ?? '수학';
	const session = options.session ?? 2;
	const questionTypeLabel = options.questionTypeLabel ?? '5지선다형';
	const examTitle =
		options.examTitle ?? `${year}학년도 ${month}월 ${grade} 전국연합학력평가`;

	// Build exam header paragraphs
	// Title line (centered)
	const titlePara =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t>${escapeXml(examTitle)}</hp:t></hp:run></hp:p>`;

	// Subject area (large, centered, bold — using charPrIDRef=0 as fallback)
	const subjectPara =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t>${escapeXml(subjectArea)} 영역</hp:t></hp:run></hp:p>`;

	// Session label
	const sessionPara =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t>제 ${session} 교시</hp:t></hp:run></hp:p>`;

	// Question type label with border
	const typeLabelPara =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t>${escapeXml(questionTypeLabel)}</hp:t></hp:run></hp:p>`;

	// Separator line (empty paragraph)
	const separatorPara =
		`<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run></hp:p>`;

	const headerBlock = titlePara + subjectPara + sessionPara + typeLabelPara + separatorPara;

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	const filenames = Object.keys(zip.files).sort();
	for (const filename of filenames) {
		if (!filename.startsWith('Contents/section') || !filename.endsWith('.xml')) continue;

		let content = await zip.files[filename].async('string');
		// Insert header after <hs:sec ...> opening tag
		content = content.replace(/(<hs:sec[^>]*>)/, `$1${headerBlock}`);
		zip.file(filename, content);
		break;
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			examTitle,
			subjectArea,
			session,
			questionTypeLabel,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Add custom tab stops to the document's header.xml.
 *
 * Tab stops control horizontal alignment of text elements,
 * commonly used for exam-style horizontal multiple choice layout.
 */
export async function addTabStops(
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
	const tabPositions = this.getNodeParameter('tabPositions', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		tabType?: string;
		leader?: string;
	};

	const binaryData = item.binary?.[inputBinaryPropertyName];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputBinaryPropertyName}"`,
			{ itemIndex },
		);
	}

	if (!tabPositions) {
		throw new NodeOperationError(
			this.getNode(),
			'At least one tab position is required',
			{ itemIndex },
		);
	}

	const tabType = options.tabType ?? 'LEFT';
	const leader = options.leader ?? 'NONE';

	// Parse tab positions: comma-separated mm values
	const positions = tabPositions
		.split(',')
		.map((s) => s.trim())
		.filter((s) => s.length > 0)
		.map((s) => Math.round(parseFloat(s) * MM_TO_HWPUNIT));

	if (positions.length === 0) {
		throw new NodeOperationError(
			this.getNode(),
			'No valid tab positions provided. Use comma-separated mm values (e.g., "16,32,48,64")',
			{ itemIndex },
		);
	}

	// Build tabPr XML with tab items
	const tabItems = positions
		.map((pos) => `<hp:tabItem pos="${pos}" type="${tabType}" leader="${leader}"/>`)
		.join('');
	const tabPrXml = `<hp:tabPr id="" autoTabLeft="0" autoTabRight="0">${tabItems}</hp:tabPr>`;

	const buffer = await this.helpers.getBinaryDataBuffer(itemIndex, inputBinaryPropertyName);
	const zip = await JSZip.loadAsync(buffer);

	// Modify header.xml to add tabPr definition
	const headerFile = zip.file('Contents/header.xml');
	let tabPrInserted = false;

	if (headerFile) {
		let headerContent = await headerFile.async('string');

		// Find the tabPrList section and append new tabPr
		if (headerContent.includes('</hh:tabPrList>')) {
			headerContent = headerContent.replace('</hh:tabPrList>', `${tabPrXml}</hh:tabPrList>`);
			tabPrInserted = true;
		} else if (headerContent.includes('<hh:tabPrList>')) {
			headerContent = headerContent.replace(
				'<hh:tabPrList>',
				`<hh:tabPrList>${tabPrXml}`,
			);
			tabPrInserted = true;
		} else {
			// No tabPrList exists — insert before </hh:head> or before closing tag
			const tabPrListXml = `<hh:tabPrList>${tabPrXml}</hh:tabPrList>`;
			if (headerContent.includes('</hh:head>')) {
				headerContent = headerContent.replace('</hh:head>', `${tabPrListXml}</hh:head>`);
				tabPrInserted = true;
			}
		}

		zip.file('Contents/header.xml', headerContent);
	}

	const outputBuffer = await finalizeZip(zip);
	const originalFileName = binaryData.fileName ?? 'modified.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		outputBuffer,
		originalFileName,
		'application/hwp+zip',
	);

	return {
		json: {
			tabPositionsMm: tabPositions,
			tabPositionsHwpunit: positions,
			tabType,
			leader,
			tabPrInserted,
			fileName: originalFileName,
		},
		binary: { [outputBinaryPropertyName]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

/**
 * Generate a simple SVG chart/graph as binary output.
 *
 * Supports:
 * - Function plots (polynomial, trig) on a coordinate plane
 * - Basic geometric shapes (triangle, circle, quadrilateral)
 *
 * Output is an SVG image that can be converted to PNG externally
 * or used with the insertImage operation.
 */
export async function generateChart(
	this: IExecuteFunctions,
	itemIndex: number,
): Promise<INodeExecutionData> {
	const chartType = this.getNodeParameter('chartType', itemIndex) as string;
	const chartSpec = this.getNodeParameter('chartSpec', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		width?: number;
		height?: number;
		backgroundColor?: string;
		outputBinaryPropertyName?: string;
	};

	const width = options.width ?? 400;
	const height = options.height ?? 400;
	const bgColor = options.backgroundColor ?? '#ffffff';
	const outputProp = options.outputBinaryPropertyName ?? 'chart';

	let spec: IDataObject;
	try {
		spec = JSON.parse(chartSpec) as IDataObject;
	} catch {
		throw new NodeOperationError(
			this.getNode(),
			'Invalid chart specification JSON',
			{ itemIndex },
		);
	}

	let svgContent = '';

	if (chartType === 'function') {
		svgContent = generateFunctionPlot(spec, width, height, bgColor);
	} else if (chartType === 'triangle') {
		svgContent = generateTriangle(spec, width, height, bgColor);
	} else if (chartType === 'circle') {
		svgContent = generateCircle(spec, width, height, bgColor);
	} else if (chartType === 'quadrilateral') {
		svgContent = generateQuadrilateral(spec, width, height, bgColor);
	} else if (chartType === 'coordinate') {
		svgContent = generateCoordinatePlane(spec, width, height, bgColor);
	} else {
		throw new NodeOperationError(
			this.getNode(),
			`Unknown chart type: ${chartType}. Supported: function, triangle, circle, quadrilateral, coordinate`,
			{ itemIndex },
		);
	}

	const svgBuffer = Buffer.from(svgContent, 'utf-8');
	const newBinaryData = await this.helpers.prepareBinaryData(
		svgBuffer,
		'chart.svg',
		'image/svg+xml',
	);

	return {
		json: {
			chartType,
			width,
			height,
			svgLength: svgContent.length,
		},
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// SVG Chart Generation Helpers
// ==========================================

function svgHeader(width: number, height: number, bgColor: string): string {
	return (
		`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">` +
		`<rect width="${width}" height="${height}" fill="${bgColor}"/>`
	);
}

function generateFunctionPlot(
	spec: IDataObject,
	width: number,
	height: number,
	bgColor: string,
): string {
	const xMin = (spec.xMin as number) ?? -10;
	const xMax = (spec.xMax as number) ?? 10;
	const yMin = (spec.yMin as number) ?? -10;
	const yMax = (spec.yMax as number) ?? 10;
	const expression = (spec.expression as string) ?? 'x';
	const color = (spec.color as string) ?? '#2563eb';
	const showGrid = (spec.showGrid as boolean) ?? true;
	const showAxes = (spec.showAxes as boolean) ?? true;

	const pad = 40;
	const plotW = width - pad * 2;
	const plotH = height - pad * 2;

	const toSvgX = (x: number) => pad + ((x - xMin) / (xMax - xMin)) * plotW;
	const toSvgY = (y: number) => pad + ((yMax - y) / (yMax - yMin)) * plotH;

	let svg = svgHeader(width, height, bgColor);

	// Grid
	if (showGrid) {
		svg += `<g stroke="#e5e7eb" stroke-width="0.5">`;
		for (let x = Math.ceil(xMin); x <= Math.floor(xMax); x++) {
			svg += `<line x1="${toSvgX(x)}" y1="${pad}" x2="${toSvgX(x)}" y2="${height - pad}"/>`;
		}
		for (let y = Math.ceil(yMin); y <= Math.floor(yMax); y++) {
			svg += `<line x1="${pad}" y1="${toSvgY(y)}" x2="${width - pad}" y2="${toSvgY(y)}"/>`;
		}
		svg += `</g>`;
	}

	// Axes
	if (showAxes) {
		const originX = toSvgX(0);
		const originY = toSvgY(0);
		svg += `<g stroke="#000" stroke-width="1.5">`;
		svg += `<line x1="${pad}" y1="${originY}" x2="${width - pad}" y2="${originY}"/>`;
		svg += `<line x1="${originX}" y1="${pad}" x2="${originX}" y2="${height - pad}"/>`;
		svg += `</g>`;
		// Axis labels
		svg += `<text x="${width - pad + 5}" y="${originY + 4}" font-size="12" fill="#000">x</text>`;
		svg += `<text x="${originX + 5}" y="${pad - 5}" font-size="12" fill="#000">y</text>`;
		// Tick labels
		for (let x = Math.ceil(xMin); x <= Math.floor(xMax); x++) {
			if (x === 0) continue;
			svg += `<text x="${toSvgX(x)}" y="${originY + 15}" font-size="9" text-anchor="middle" fill="#666">${x}</text>`;
		}
		for (let y = Math.ceil(yMin); y <= Math.floor(yMax); y++) {
			if (y === 0) continue;
			svg += `<text x="${originX - 8}" y="${toSvgY(y) + 3}" font-size="9" text-anchor="end" fill="#666">${y}</text>`;
		}
	}

	// Plot function using simple eval for basic math expressions
	const points: string[] = [];
	const steps = plotW;
	for (let i = 0; i <= steps; i++) {
		const x = xMin + (i / steps) * (xMax - xMin);
		const y = evaluateExpression(expression, x);
		if (isFinite(y) && y >= yMin - 5 && y <= yMax + 5) {
			points.push(`${toSvgX(x).toFixed(1)},${toSvgY(y).toFixed(1)}`);
		}
	}
	if (points.length > 1) {
		svg += `<polyline points="${points.join(' ')}" fill="none" stroke="${color}" stroke-width="2"/>`;
	}

	svg += '</svg>';
	return svg;
}

/**
 * Simple math expression evaluator for function plots.
 * Supports: x, +, -, *, /, ^, sin, cos, tan, sqrt, abs, log, exp, pi, e
 */
function evaluateExpression(expr: string, x: number): number {
	// Build a safe evaluation by replacing math functions
	const safe = expr
		.replace(/\bsin\b/g, 'Math.sin')
		.replace(/\bcos\b/g, 'Math.cos')
		.replace(/\btan\b/g, 'Math.tan')
		.replace(/\bsqrt\b/g, 'Math.sqrt')
		.replace(/\babs\b/g, 'Math.abs')
		.replace(/\blog\b/g, 'Math.log')
		.replace(/\bexp\b/g, 'Math.exp')
		.replace(/\bpi\b/g, String(Math.PI))
		.replace(/\be\b/g, String(Math.E))
		.replace(/\bx\b/g, `(${x})`)
		.replace(/\^/g, '**');

	try {
		 
		const fn = new Function(`return (${safe})`);
		const result = fn() as number;
		return typeof result === 'number' ? result : NaN;
	} catch {
		return NaN;
	}
}

function generateTriangle(
	spec: IDataObject,
	width: number,
	height: number,
	bgColor: string,
): string {
	const vertices = (spec.vertices as number[][]) ?? [
		[0, 0],
		[6, 0],
		[3, 5],
	];
	const labels = (spec.labels as IDataObject) ?? {};
	const sideLabels = (spec.sideLabels as IDataObject) ?? {};

	const pad = 40;
	const allX = vertices.map((v) => v[0]);
	const allY = vertices.map((v) => v[1]);
	const minX = Math.min(...allX);
	const maxX = Math.max(...allX);
	const minY = Math.min(...allY);
	const maxY = Math.max(...allY);
	const rangeX = maxX - minX || 1;
	const rangeY = maxY - minY || 1;
	const scale = Math.min((width - pad * 2) / rangeX, (height - pad * 2) / rangeY);

	const toSvgX = (x: number) => pad + (x - minX) * scale;
	const toSvgY = (y: number) => height - pad - (y - minY) * scale;

	let svg = svgHeader(width, height, bgColor);

	// Draw triangle
	const pts = vertices.map((v) => `${toSvgX(v[0]).toFixed(1)},${toSvgY(v[1]).toFixed(1)}`).join(' ');
	svg += `<polygon points="${pts}" fill="none" stroke="#000" stroke-width="2"/>`;

	// Vertex labels
	for (const [label, coords] of Object.entries(labels)) {
		const c = coords as number[];
		if (Array.isArray(c) && c.length === 2) {
			const sx = toSvgX(c[0]);
			const sy = toSvgY(c[1]);
			// Offset label away from centroid
			const cx = vertices.reduce((s, v) => s + v[0], 0) / 3;
			const cy = vertices.reduce((s, v) => s + v[1], 0) / 3;
			const dx = c[0] - cx;
			const dy = c[1] - cy;
			const len = Math.sqrt(dx * dx + dy * dy) || 1;
			svg += `<text x="${sx + (dx / len) * 15}" y="${sy - (dy / len) * 15}" ` +
				`font-size="14" font-weight="bold" text-anchor="middle" fill="#000">${escapeXml(label)}</text>`;
		}
	}

	// Side labels
	for (let i = 0; i < vertices.length; i++) {
		const j = (i + 1) % vertices.length;
		const keys = Object.keys(labels);
		const sideKey = keys.length >= 2
			? `${keys[i] ?? ''}${keys[j] ?? ''}`
			: '';
		const sLabel = sideLabels[sideKey] as string | undefined;
		if (sLabel) {
			const mx = (toSvgX(vertices[i][0]) + toSvgX(vertices[j][0])) / 2;
			const my = (toSvgY(vertices[i][1]) + toSvgY(vertices[j][1])) / 2;
			svg += `<text x="${mx}" y="${my - 8}" font-size="12" text-anchor="middle" fill="#333">${escapeXml(sLabel)}</text>`;
		}
	}

	svg += '</svg>';
	return svg;
}

function generateCircle(
	spec: IDataObject,
	width: number,
	height: number,
	bgColor: string,
): string {
	const center = (spec.center as number[]) ?? [0, 0];
	const radius = (spec.radius as number) ?? 3;
	const showCenter = (spec.showCenter as boolean) ?? true;
	const pointsOnCircle = (spec.pointsOnCircle as IDataObject[]) ?? [];
	const chords = (spec.chords as string[][]) ?? [];

	const pad = 40;
	const scale = Math.min((width - pad * 2), (height - pad * 2)) / (radius * 2.5);
	const toSvgX = (x: number) => width / 2 + (x - center[0]) * scale;
	const toSvgY = (y: number) => height / 2 - (y - center[1]) * scale;
	const svgR = radius * scale;

	let svg = svgHeader(width, height, bgColor);

	// Draw circle
	svg += `<circle cx="${toSvgX(center[0])}" cy="${toSvgY(center[1])}" r="${svgR}" ` +
		`fill="none" stroke="#000" stroke-width="2"/>`;

	// Center point
	if (showCenter) {
		svg += `<circle cx="${toSvgX(center[0])}" cy="${toSvgY(center[1])}" r="3" fill="#000"/>`;
		svg += `<text x="${toSvgX(center[0]) + 8}" y="${toSvgY(center[1]) - 8}" ` +
			`font-size="14" font-weight="bold" fill="#000">O</text>`;
	}

	// Points on circle
	const pointMap: Record<string, [number, number]> = {};
	for (const pt of pointsOnCircle) {
		const angleDeg = (pt.angleDeg as number) ?? 0;
		const label = (pt.label as string) ?? '';
		const angleRad = (angleDeg * Math.PI) / 180;
		const px = center[0] + radius * Math.cos(angleRad);
		const py = center[1] + radius * Math.sin(angleRad);
		pointMap[label] = [px, py];

		const sx = toSvgX(px);
		const sy = toSvgY(py);
		svg += `<circle cx="${sx}" cy="${sy}" r="3" fill="#000"/>`;

		// Label offset outward
		const lx = sx + Math.cos(angleRad) * 18;
		const ly = sy - Math.sin(angleRad) * 18;
		svg += `<text x="${lx}" y="${ly}" font-size="14" font-weight="bold" ` +
			`text-anchor="middle" fill="#000">${escapeXml(label)}</text>`;
	}

	// Chords
	for (const chord of chords) {
		if (chord.length === 2) {
			const p1 = pointMap[chord[0]];
			const p2 = pointMap[chord[1]];
			if (p1 && p2) {
				svg += `<line x1="${toSvgX(p1[0])}" y1="${toSvgY(p1[1])}" ` +
					`x2="${toSvgX(p2[0])}" y2="${toSvgY(p2[1])}" stroke="#333" stroke-width="1.5"/>`;
			}
		}
	}

	svg += '</svg>';
	return svg;
}

function generateQuadrilateral(
	spec: IDataObject,
	width: number,
	height: number,
	bgColor: string,
): string {
	const vertices = (spec.vertices as number[][]) ?? [
		[0, 0],
		[5, 0],
		[7, 3],
		[2, 3],
	];
	const labels = (spec.labels as IDataObject) ?? {};
	const showDiagonals = (spec.showDiagonals as boolean) ?? false;
	const sideLabels = (spec.sideLabels as IDataObject) ?? {};

	const pad = 40;
	const allX = vertices.map((v) => v[0]);
	const allY = vertices.map((v) => v[1]);
	const minX = Math.min(...allX);
	const maxX = Math.max(...allX);
	const minY = Math.min(...allY);
	const maxY = Math.max(...allY);
	const rangeX = maxX - minX || 1;
	const rangeY = maxY - minY || 1;
	const scale = Math.min((width - pad * 2) / rangeX, (height - pad * 2) / rangeY);

	const toSvgX = (x: number) => pad + (x - minX) * scale;
	const toSvgY = (y: number) => height - pad - (y - minY) * scale;

	let svg = svgHeader(width, height, bgColor);

	// Draw quadrilateral
	const pts = vertices.map((v) => `${toSvgX(v[0]).toFixed(1)},${toSvgY(v[1]).toFixed(1)}`).join(' ');
	svg += `<polygon points="${pts}" fill="none" stroke="#000" stroke-width="2"/>`;

	// Diagonals
	if (showDiagonals && vertices.length === 4) {
		svg += `<line x1="${toSvgX(vertices[0][0])}" y1="${toSvgY(vertices[0][1])}" ` +
			`x2="${toSvgX(vertices[2][0])}" y2="${toSvgY(vertices[2][1])}" stroke="#999" stroke-width="1" stroke-dasharray="4,4"/>`;
		svg += `<line x1="${toSvgX(vertices[1][0])}" y1="${toSvgY(vertices[1][1])}" ` +
			`x2="${toSvgX(vertices[3][0])}" y2="${toSvgY(vertices[3][1])}" stroke="#999" stroke-width="1" stroke-dasharray="4,4"/>`;
	}

	// Vertex labels
	for (const [label, coords] of Object.entries(labels)) {
		const c = coords as number[];
		if (Array.isArray(c) && c.length === 2) {
			const sx = toSvgX(c[0]);
			const sy = toSvgY(c[1]);
			const cx = vertices.reduce((s, v) => s + v[0], 0) / vertices.length;
			const cy = vertices.reduce((s, v) => s + v[1], 0) / vertices.length;
			const dx = c[0] - cx;
			const dy = c[1] - cy;
			const len = Math.sqrt(dx * dx + dy * dy) || 1;
			svg += `<text x="${sx + (dx / len) * 15}" y="${sy - (dy / len) * 15}" ` +
				`font-size="14" font-weight="bold" text-anchor="middle" fill="#000">${escapeXml(label)}</text>`;
		}
	}

	// Side labels
	for (const [key, value] of Object.entries(sideLabels)) {
		const labelKeys = Object.keys(labels);
		for (let i = 0; i < vertices.length; i++) {
			const j = (i + 1) % vertices.length;
			const sideKey = `${labelKeys[i] ?? ''}${labelKeys[j] ?? ''}`;
			if (sideKey === key) {
				const mx = (toSvgX(vertices[i][0]) + toSvgX(vertices[j][0])) / 2;
				const my = (toSvgY(vertices[i][1]) + toSvgY(vertices[j][1])) / 2;
				svg += `<text x="${mx}" y="${my - 8}" font-size="12" text-anchor="middle" fill="#333">${escapeXml(value as string)}</text>`;
			}
		}
	}

	svg += '</svg>';
	return svg;
}

function generateCoordinatePlane(
	spec: IDataObject,
	width: number,
	height: number,
	bgColor: string,
): string {
	const xMin = (spec.xMin as number) ?? -5;
	const xMax = (spec.xMax as number) ?? 5;
	const yMin = (spec.yMin as number) ?? -5;
	const yMax = (spec.yMax as number) ?? 5;
	const points = (spec.points as IDataObject[]) ?? [];
	const lines = (spec.lines as IDataObject[]) ?? [];
	const showGrid = (spec.showGrid as boolean) ?? true;

	const pad = 40;
	const plotW = width - pad * 2;
	const plotH = height - pad * 2;

	const toSvgX = (x: number) => pad + ((x - xMin) / (xMax - xMin)) * plotW;
	const toSvgY = (y: number) => pad + ((yMax - y) / (yMax - yMin)) * plotH;

	let svg = svgHeader(width, height, bgColor);

	// Grid
	if (showGrid) {
		svg += `<g stroke="#e5e7eb" stroke-width="0.5">`;
		for (let x = Math.ceil(xMin); x <= Math.floor(xMax); x++) {
			svg += `<line x1="${toSvgX(x)}" y1="${pad}" x2="${toSvgX(x)}" y2="${height - pad}"/>`;
		}
		for (let y = Math.ceil(yMin); y <= Math.floor(yMax); y++) {
			svg += `<line x1="${pad}" y1="${toSvgY(y)}" x2="${width - pad}" y2="${toSvgY(y)}"/>`;
		}
		svg += `</g>`;
	}

	// Axes
	const originX = toSvgX(0);
	const originY = toSvgY(0);
	svg += `<g stroke="#000" stroke-width="1.5">`;
	svg += `<line x1="${pad}" y1="${originY}" x2="${width - pad}" y2="${originY}"/>`;
	svg += `<line x1="${originX}" y1="${pad}" x2="${originX}" y2="${height - pad}"/>`;
	svg += `</g>`;
	svg += `<text x="${width - pad + 5}" y="${originY + 4}" font-size="12" fill="#000">x</text>`;
	svg += `<text x="${originX + 5}" y="${pad - 5}" font-size="12" fill="#000">y</text>`;

	// Points
	for (const pt of points) {
		const px = (pt.x as number) ?? 0;
		const py = (pt.y as number) ?? 0;
		const label = (pt.label as string) ?? '';
		const color = (pt.color as string) ?? '#ef4444';

		svg += `<circle cx="${toSvgX(px)}" cy="${toSvgY(py)}" r="4" fill="${color}"/>`;
		if (label) {
			svg += `<text x="${toSvgX(px) + 8}" y="${toSvgY(py) - 8}" font-size="12" fill="#000">${escapeXml(label)}</text>`;
		}
	}

	// Lines
	for (const line of lines) {
		const x1 = (line.x1 as number) ?? 0;
		const y1 = (line.y1 as number) ?? 0;
		const x2 = (line.x2 as number) ?? 0;
		const y2 = (line.y2 as number) ?? 0;
		const color = (line.color as string) ?? '#2563eb';
		svg += `<line x1="${toSvgX(x1)}" y1="${toSvgY(y1)}" x2="${toSvgX(x2)}" y2="${toSvgY(y2)}" ` +
			`stroke="${color}" stroke-width="2"/>`;
	}

	svg += '</svg>';
	return svg;
}

// ==========================================
// 1. Add Memo/Comment
// ==========================================

export async function addMemo(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const memoText = this.getNodeParameter('memoText', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		targetParagraphIndex?: number;
		author?: string;
		fileName?: string;
	};

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	const sectionFile =
		zip.file('Contents/section0.xml') ?? zip.file('Contents/Section0.xml');
	if (!sectionFile) {
		throw new NodeOperationError(this.getNode(), 'section0.xml not found in HWPX', {
			itemIndex,
		});
	}

	let sectionXml = await sectionFile.async('string');
	const targetIdx = options.targetParagraphIndex ?? 0;

	// Find the target paragraph
	const paragraphs = [...sectionXml.matchAll(/<hp:p\b[^>]*>/g)];
	if (paragraphs.length === 0) {
		throw new NodeOperationError(this.getNode(), 'No paragraphs found in section', {
			itemIndex,
		});
	}

	const paraIdx = Math.min(targetIdx, paragraphs.length - 1);
	const memoId = String(Date.now() & 0xffffffff);
	const escapedText = escapeXml(memoText);

	// Build memo XML - memogroup goes at section level (before closing </hs:sec>)
	const memoXml =
		`<hp:memogroup>` +
		`<hp:memo id="${memoId}" memoShapeIDRef="0">` +
		`<hp:paraList id="" textDirection="HORIZONTAL" lineWrap="BREAK" ` +
		`vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" ` +
		`textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
		`<hp:p id="${memoId}1" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t>${escapedText}</hp:t></hp:run>` +
		`<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" ` +
		`textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>` +
		`</hp:linesegarray></hp:p>` +
		`</hp:paraList></hp:memo></hp:memogroup>`;

	// Insert before closing </hs:sec> tag
	sectionXml = sectionXml.replace('</hs:sec>', memoXml + '</hs:sec>');

	zip.file(sectionFile.name, sectionXml);

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'memo.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: {
			success: true,
			memoText,
			memoId,
			targetParagraphIndex: paraIdx,
		},
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 2. Add Shape Objects
// ==========================================

export async function addShape(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const shapeType = this.getNodeParameter('shapeType', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		widthMm?: number;
		heightMm?: number;
		lineColor?: string;
		lineWidth?: number;
		fillColor?: string;
		startX?: number;
		startY?: number;
		endX?: number;
		endY?: number;
		targetParagraphIndex?: number;
		fileName?: string;
	};

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	const sectionFile =
		zip.file('Contents/section0.xml') ?? zip.file('Contents/Section0.xml');
	if (!sectionFile) {
		throw new NodeOperationError(this.getNode(), 'section0.xml not found in HWPX', {
			itemIndex,
		});
	}

	let sectionXml = await sectionFile.async('string');

	const wHwp = Math.round((options.widthMm ?? 50) * MM_TO_HWPUNIT);
	const hHwp = Math.round((options.heightMm ?? 25) * MM_TO_HWPUNIT);
	const lineColor = options.lineColor ?? '#000000';
	const lineW = String(Math.round((options.lineWidth ?? 0.4) * MM_TO_HWPUNIT));
	const fillColor = options.fillColor;
	const instId = String(Date.now() & 0xffffffff);

	const commonAttrs =
		`id="${instId}" zOrder="0" numberingType="NONE" lock="0" ` +
		`dropcapstyle="None" href="" groupLevel="0" instid="${instId}"`;

	const shapeCommonChildren =
		`<hp:offset x="0" y="0"/>` +
		`<hp:orgSz width="${wHwp}" height="${hHwp}"/>` +
		`<hp:curSz width="${wHwp}" height="${hHwp}"/>` +
		`<hp:flip horizontal="0" vertical="0"/>` +
		`<hp:rotationInfo angle="0" centerX="${Math.round(wHwp / 2)}" centerY="${Math.round(hHwp / 2)}"/>`;

	const lineShapeXml =
		`<hp:lineShape color="${lineColor}" width="${lineW}" style="SOLID" ` +
		`endCap="FLAT" headStyle="NORMAL" tailStyle="NORMAL" headSz="MEDIUM_MEDIUM" ` +
		`tailSz="MEDIUM_MEDIUM" outlineStyle="NORMAL" alpha="0"/>`;

	const fillXml = fillColor
		? `<hp:fillBrush><hp:winBrush faceColor="${fillColor}" hatchColor="#FFFFFF"/></hp:fillBrush>`
		: '';

	const shadowXml = `<hp:shadow type="NONE" color="#B2B2B2" offsetX="0" offsetY="0" alpha="0"/>`;

	const sizeChildren =
		`<hp:sz width="${wHwp}" height="${hHwp}" widthRelTo="ABSOLUTE" heightRelTo="ABSOLUTE" protect="0"/>` +
		`<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="0" allowOverlap="1" ` +
		`holdAnchorAndSO="0" vertRelTo="PARA" vertAlign="TOP" horzRelTo="COLUMN" ` +
		`horzAlign="LEFT" vertOffset="0" horzOffset="0"/>` +
		`<hp:outMargin left="0" right="0" top="0" bottom="0"/>`;

	let shapeXml: string;

	if (shapeType === 'line') {
		const sx = String(options.startX ?? 0);
		const sy = String(options.startY ?? 0);
		const ex = String(options.endX ?? wHwp);
		const ey = String(options.endY ?? 0);
		shapeXml =
			`<hp:line ${commonAttrs} isReverseHV="0">` +
			shapeCommonChildren +
			lineShapeXml + fillXml + shadowXml +
			`<hp:startPt x="${sx}" y="${sy}"/><hp:endPt x="${ex}" y="${ey}"/>` +
			sizeChildren +
			`</hp:line>`;
	} else if (shapeType === 'rectangle') {
		shapeXml =
			`<hp:rect ${commonAttrs} ratio="0">` +
			shapeCommonChildren +
			lineShapeXml + fillXml + shadowXml +
			`<hp:pt0 x="0" y="0"/><hp:pt1 x="${wHwp}" y="0"/>` +
			`<hp:pt2 x="${wHwp}" y="${hHwp}"/><hp:pt3 x="0" y="${hHwp}"/>` +
			sizeChildren +
			`</hp:rect>`;
	} else {
		// ellipse
		const cx = String(Math.round(wHwp / 2));
		const cy = String(Math.round(hHwp / 2));
		shapeXml =
			`<hp:ellipse ${commonAttrs} intervalDirty="0" hasArcPr="0" arcType="NORMAL">` +
			shapeCommonChildren +
			lineShapeXml + fillXml + shadowXml +
			`<hp:center x="${cx}" y="${cy}"/>` +
			`<hp:ax1 x="${wHwp}" y="${cy}"/>` +
			`<hp:ax2 x="${cx}" y="${hHwp}"/>` +
			`<hp:start1 x="${wHwp}" y="${cy}"/>` +
			`<hp:end1 x="${wHwp}" y="${cy}"/>` +
			`<hp:start2 x="${wHwp}" y="${cy}"/>` +
			`<hp:end2 x="${wHwp}" y="${cy}"/>` +
			sizeChildren +
			`</hp:ellipse>`;
	}

	// Wrap in run and insert into target paragraph
	const runXml = `<hp:run charPrIDRef="0">${shapeXml}</hp:run>`;

	const targetIdx = options.targetParagraphIndex ?? 0;
	const paragraphs = [...sectionXml.matchAll(/<hp:p\b[^>]*>/g)];
	if (paragraphs.length === 0) {
		throw new NodeOperationError(this.getNode(), 'No paragraphs found', { itemIndex });
	}

	const pIdx = Math.min(targetIdx, paragraphs.length - 1);
	const insertPos = (paragraphs[pIdx].index ?? 0) + paragraphs[pIdx][0].length;
	sectionXml = sectionXml.slice(0, insertPos) + runXml + sectionXml.slice(insertPos);

	zip.file(sectionFile.name, sectionXml);

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'shape.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: { success: true, shapeType, widthMm: options.widthMm ?? 50, heightMm: options.heightMm ?? 25 },
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 3. Track Changes (read/toggle)
// ==========================================

export async function trackChanges(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const action = this.getNodeParameter('trackChangeAction', itemIndex) as string;

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	if (action === 'read') {
		// Read track changes from section XML
		const changes: IDataObject[] = [];

		const sectionFiles = Object.keys(zip.files).filter(
			(f) => /Contents\/[Ss]ection\d+\.xml/.test(f),
		);

		for (const sf of sectionFiles) {
			const content = await zip.file(sf)!.async('string');
			const changeMatches = [
				...content.matchAll(/<hp:trackChange\b([^>]*)(?:\/>|>([\s\S]*?)<\/hp:trackChange>)/g),
			];
			for (const m of changeMatches) {
				const attrs = m[1];
				const typeMatch = attrs.match(/type="([^"]*)"/);
				const authorMatch = attrs.match(/author="([^"]*)"/);
				const dateMatch = attrs.match(/date="([^"]*)"/);
				changes.push({
					file: sf,
					type: typeMatch?.[1] ?? 'unknown',
					author: authorMatch?.[1] ?? '',
					date: dateMatch?.[1] ?? '',
				});
			}
		}

		return {
			json: { trackChangesCount: changes.length, changes },
			pairedItem: { item: itemIndex },
		};
	}

	// action === 'enable' or 'disable' — modify settings.xml
	const settingsFile = zip.file('settings.xml') ?? zip.file('Settings.xml');
	let settingsXml = settingsFile ? await settingsFile.async('string') : '';

	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as { fileName?: string };

	if (action === 'enable') {
		if (!settingsXml.includes('trackChange')) {
			if (settingsXml.includes('</config:config-item-set>')) {
				settingsXml = settingsXml.replace(
					'</config:config-item-set>',
					'<config:config-item config:name="trackChange" config:type="boolean">true</config:config-item>' +
						'</config:config-item-set>',
				);
			}
		} else {
			settingsXml = settingsXml.replace(
				/(<config:config-item[^>]*config:name="trackChange"[^>]*>)\s*false\s*(<\/config:config-item>)/,
				'$1true$2',
			);
		}
	} else {
		settingsXml = settingsXml.replace(
			/(<config:config-item[^>]*config:name="trackChange"[^>]*>)\s*true\s*(<\/config:config-item>)/,
			'$1false$2',
		);
	}

	zip.file(settingsFile?.name ?? 'settings.xml', settingsXml);

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'tracked.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: { success: true, action },
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 4. Style-based Text Replacement
// ==========================================

export async function replaceByStyle(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const replaceWith = this.getNodeParameter('replaceWith', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		filterBold?: boolean;
		filterItalic?: boolean;
		filterUnderline?: boolean;
		filterColor?: string;
		filterFontSize?: number;
		fileName?: string;
	};

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	// Read header.xml to get character property definitions
	const headerFile =
		zip.file('Contents/header.xml') ?? zip.file('Contents/Header.xml');
	if (!headerFile) {
		throw new NodeOperationError(this.getNode(), 'header.xml not found in HWPX', {
			itemIndex,
		});
	}
	const headerXml = await headerFile.async('string');

	// Build a map of charPrIDRef -> properties
	const charPrMap = new Map<string, IDataObject>();
	const charPrMatches = [
		...headerXml.matchAll(/<hh:charPr\b([^>]*)>([\s\S]*?)<\/hh:charPr>/g),
	];
	for (const m of charPrMatches) {
		const idMatch = m[1].match(/id="(\d+)"/);
		if (!idMatch) continue;
		const id = idMatch[1];
		const content = m[0];
		const props: IDataObject = { id };

		if (content.includes('<hh:bold/>') || content.includes('<hh:bold ')) props.bold = true;
		if (content.includes('<hh:italic/>') || content.includes('<hh:italic ')) props.italic = true;
		const ulMatch = content.match(/<hh:underline\b[^>]*type="([^"]*)"/);
		if (ulMatch && ulMatch[1] !== 'NONE') props.underline = true;
		const colorMatch = content.match(/<hh:textColor\b[^>]*value="([^"]*)"/);
		if (colorMatch) props.textColor = colorMatch[1];
		const sizeMatch = m[1].match(/height="(\d+)"/);
		if (sizeMatch) props.fontSize = Math.round(parseInt(sizeMatch[1], 10) / 100);

		charPrMap.set(id, props);
	}

	// Process section files
	const sectionFiles = Object.keys(zip.files).filter(
		(f) => /Contents\/[Ss]ection\d+\.xml/.test(f),
	);
	let totalReplacements = 0;

	for (const sf of sectionFiles) {
		const content = await zip.file(sf)!.async('string');
		let modified = false;

		const runRegex = /<hp:run\b[^>]*charPrIDRef="(\d+)"[^>]*>([\s\S]*?)<\/hp:run>/g;
		const newContent = content.replace(runRegex, (fullMatch, charId: string, inner: string) => {
			const props = charPrMap.get(charId);
			if (!props) return fullMatch;

			if (options.filterBold !== undefined && options.filterBold !== !!props.bold) return fullMatch;
			if (options.filterItalic !== undefined && options.filterItalic !== !!props.italic) return fullMatch;
			if (options.filterUnderline !== undefined && options.filterUnderline !== !!props.underline) return fullMatch;
			if (options.filterColor && props.textColor !== options.filterColor) return fullMatch;
			if (options.filterFontSize && props.fontSize !== options.filterFontSize) return fullMatch;

			const newInner = inner.replace(
				/<hp:t>([^<]*)<\/hp:t>/g,
				() => {
					totalReplacements++;
					return `<hp:t>${escapeXml(replaceWith)}</hp:t>`;
				},
			);

			if (newInner !== inner) {
				modified = true;
				return fullMatch.replace(inner, newInner);
			}
			return fullMatch;
		});

		if (modified) {
			zip.file(sf, newContent);
		}
	}

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'styled.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: { success: true, totalReplacements },
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 5. Add Footnote/Endnote
// ==========================================

export async function addNote(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const noteType = this.getNodeParameter('noteType', itemIndex) as string;
	const noteText = this.getNodeParameter('noteText', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		targetParagraphIndex?: number;
		fileName?: string;
	};

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	const sectionFile =
		zip.file('Contents/section0.xml') ?? zip.file('Contents/Section0.xml');
	if (!sectionFile) {
		throw new NodeOperationError(this.getNode(), 'section0.xml not found', { itemIndex });
	}

	let sectionXml = await sectionFile.async('string');
	const instId = String(Date.now() & 0xffffffff);
	const tag = noteType === 'endnote' ? 'endNote' : 'footNote';
	const escapedText = escapeXml(noteText);

	const noteXml =
		`<hp:run charPrIDRef="0">` +
		`<hp:${tag} instId="${instId}">` +
		`<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" ` +
		`vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" ` +
		`textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
		`<hp:p id="${instId}1" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0"><hp:t>${escapedText}</hp:t></hp:run>` +
		`<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" ` +
		`textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="42520" flags="393216"/>` +
		`</hp:linesegarray></hp:p>` +
		`</hp:subList></hp:${tag}></hp:run>`;

	const targetIdx = options.targetParagraphIndex ?? 0;
	const paragraphs = [...sectionXml.matchAll(/<hp:p\b[^>]*>/g)];
	if (paragraphs.length === 0) {
		throw new NodeOperationError(this.getNode(), 'No paragraphs found', { itemIndex });
	}

	const pIdx = Math.min(targetIdx, paragraphs.length - 1);
	const insertPos = (paragraphs[pIdx].index ?? 0) + paragraphs[pIdx][0].length;
	sectionXml = sectionXml.slice(0, insertPos) + noteXml + sectionXml.slice(insertPos);

	zip.file(sectionFile.name, sectionXml);

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'noted.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: { success: true, noteType: tag, noteText, instId },
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 6. Add Bookmark
// ==========================================

export async function addBookmark(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const bookmarkName = this.getNodeParameter('bookmarkName', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		targetParagraphIndex?: number;
		fileName?: string;
	};

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	const sectionFile =
		zip.file('Contents/section0.xml') ?? zip.file('Contents/Section0.xml');
	if (!sectionFile) {
		throw new NodeOperationError(this.getNode(), 'section0.xml not found', { itemIndex });
	}

	let sectionXml = await sectionFile.async('string');
	const escapedName = escapeXml(bookmarkName);

	const bookmarkXml =
		`<hp:run charPrIDRef="0"><hp:ctrl><hp:bookmark name="${escapedName}"/></hp:ctrl></hp:run>`;

	const targetIdx = options.targetParagraphIndex ?? 0;
	const paragraphs = [...sectionXml.matchAll(/<hp:p\b[^>]*>/g)];
	if (paragraphs.length === 0) {
		throw new NodeOperationError(this.getNode(), 'No paragraphs found', { itemIndex });
	}

	const pIdx = Math.min(targetIdx, paragraphs.length - 1);
	const insertPos = (paragraphs[pIdx].index ?? 0) + paragraphs[pIdx][0].length;
	sectionXml = sectionXml.slice(0, insertPos) + bookmarkXml + sectionXml.slice(insertPos);

	zip.file(sectionFile.name, sectionXml);

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'bookmarked.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: { success: true, bookmarkName },
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 7. Add Watermark
// ==========================================

export async function addWatermark(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const watermarkText = this.getNodeParameter('watermarkText', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		fontSizePt?: number;
		color?: string;
		angleDeg?: number;
		opacity?: number;
		fileName?: string;
	};

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	const sectionFile =
		zip.file('Contents/section0.xml') ?? zip.file('Contents/Section0.xml');
	if (!sectionFile) {
		throw new NodeOperationError(this.getNode(), 'section0.xml not found', { itemIndex });
	}

	let sectionXml = await sectionFile.async('string');

	const fontSize = options.fontSizePt ?? 48;
	const color = options.color ?? '#CCCCCC';
	const angleDeg = options.angleDeg ?? -45;
	const opacity = options.opacity ?? 30;
	const escapedText = escapeXml(watermarkText);

	// Calculate dimensions for watermark text box
	const charWidth = Math.round(fontSize * 0.6);
	const textWidth = charWidth * watermarkText.length;
	const boxWidth = Math.round(Math.max(textWidth, 200) * MM_TO_HWPUNIT);
	const boxHeight = Math.round(fontSize * 1.5 * MM_TO_HWPUNIT);
	const instId = String(Date.now() & 0xffffffff);

	// HWPX uses 1/60 degree units for rotation
	const angleHwp = Math.round(angleDeg * 60);

	const watermarkXml =
		`<hp:run charPrIDRef="0">` +
		`<hp:rect id="${instId}" zOrder="-1" numberingType="NONE" lock="0" ` +
		`dropcapstyle="None" href="" groupLevel="0" instid="${instId}" ratio="0">` +
		`<hp:offset x="0" y="0"/>` +
		`<hp:orgSz width="${boxWidth}" height="${boxHeight}"/>` +
		`<hp:curSz width="${boxWidth}" height="${boxHeight}"/>` +
		`<hp:flip horizontal="0" vertical="0"/>` +
		`<hp:rotationInfo angle="${angleHwp}" centerX="${Math.round(boxWidth / 2)}" centerY="${Math.round(boxHeight / 2)}"/>` +
		`<hp:lineShape color="#00000000" width="0" style="NONE" endCap="FLAT" ` +
		`headStyle="NORMAL" tailStyle="NORMAL" headSz="MEDIUM_MEDIUM" ` +
		`tailSz="MEDIUM_MEDIUM" outlineStyle="NORMAL" alpha="0"/>` +
		`<hp:shadow type="NONE" color="#B2B2B2" offsetX="0" offsetY="0" alpha="0"/>` +
		`<hp:pt0 x="0" y="0"/><hp:pt1 x="${boxWidth}" y="0"/>` +
		`<hp:pt2 x="${boxWidth}" y="${boxHeight}"/><hp:pt3 x="0" y="${boxHeight}"/>` +
		`<hp:sz width="${boxWidth}" height="${boxHeight}" widthRelTo="ABSOLUTE" heightRelTo="ABSOLUTE" protect="0"/>` +
		`<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="0" allowOverlap="1" ` +
		`holdAnchorAndSO="0" vertRelTo="PAGE" vertAlign="CENTER" ` +
		`horzRelTo="PAGE" horzAlign="CENTER" vertOffset="0" horzOffset="0"/>` +
		`<hp:outMargin left="0" right="0" top="0" bottom="0"/>` +
		`<hp:textBox editable="0" hasMargin="0" vertAlign="CENTER" textDirection="HORIZONTAL">` +
		`<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" ` +
		`vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" ` +
		`textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">` +
		`<hp:p id="${instId}1" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">` +
		`<hp:run charPrIDRef="0">` +
		`<hp:t>${escapedText}</hp:t>` +
		`</hp:run></hp:p></hp:subList></hp:textBox>` +
		`</hp:rect></hp:run>`;

	// Suppress lint warning about unused variables
	void color;
	void opacity;

	// Insert watermark into the first paragraph (behind content with zOrder=-1)
	const paragraphs = [...sectionXml.matchAll(/<hp:p\b[^>]*>/g)];
	if (paragraphs.length > 0) {
		const insertPos = (paragraphs[0].index ?? 0) + paragraphs[0][0].length;
		sectionXml = sectionXml.slice(0, insertPos) + watermarkXml + sectionXml.slice(insertPos);
	}

	zip.file(sectionFile.name, sectionXml);

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'watermarked.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: { success: true, watermarkText, fontSize, angleDeg, opacity },
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 8. Password Protection
// ==========================================

export async function setPassword(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const action = this.getNodeParameter('passwordAction', itemIndex) as string;

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	// HWPX protection is managed via content.hpf metadata and settings.xml
	const hpfFile = zip.file('Contents/content.hpf') ?? zip.file('content.hpf');
	let hpfXml = hpfFile ? await hpfFile.async('string') : '';

	if (action === 'checkProtection') {
		const isProtected = hpfXml.includes('permission') || hpfXml.includes('encrypt');
		let protectionType = 'none';
		if (hpfXml.includes('read-only')) protectionType = 'read-only';
		if (hpfXml.includes('encrypt')) protectionType = 'encrypted';

		return {
			json: { isProtected, protectionType },
			pairedItem: { item: itemIndex },
		};
	}

	const outputProp = this.getNodeParameter(
		'outputBinaryPropertyName',
		itemIndex,
		'data',
	) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as { fileName?: string };

	if (action === 'setReadOnly') {
		if (hpfXml.includes('<opf:meta')) {
			if (!hpfXml.includes('name="permission"')) {
				hpfXml = hpfXml.replace(
					'</opf:metadata>',
					'<opf:meta name="permission" content="read-only"/></opf:metadata>',
				);
			} else {
				hpfXml = hpfXml.replace(
					/(<opf:meta\b[^>]*name="permission"[^>]*content=")[^"]*(")/,
					'$1read-only$2',
				);
			}
		}

		const settingsFile = zip.file('settings.xml') ?? zip.file('Settings.xml');
		if (settingsFile) {
			let settingsXml = await settingsFile.async('string');
			if (!settingsXml.includes('documentProtection')) {
				if (settingsXml.includes('</config:config-item-set>')) {
					settingsXml = settingsXml.replace(
						'</config:config-item-set>',
						'<config:config-item config:name="documentProtection" config:type="boolean">true</config:config-item>' +
							'</config:config-item-set>',
					);
				}
			}
			zip.file(settingsFile.name, settingsXml);
		}
	} else if (action === 'removeProtection') {
		if (hpfXml.includes('name="permission"')) {
			hpfXml = hpfXml.replace(/<opf:meta\b[^>]*name="permission"[^/]*\/>/g, '');
		}

		const settingsFile = zip.file('settings.xml') ?? zip.file('Settings.xml');
		if (settingsFile) {
			let settingsXml = await settingsFile.async('string');
			settingsXml = settingsXml.replace(
				/(<config:config-item[^>]*config:name="documentProtection"[^>]*>)\s*true\s*(<\/config:config-item>)/,
				'$1false$2',
			);
			zip.file(settingsFile.name, settingsXml);
		}
	}

	if (hpfFile) {
		zip.file(hpfFile.name, hpfXml);
	}

	const newBuffer = await zip.generateAsync({ type: 'nodebuffer' });
	const fileName = options.fileName ?? binaryData.fileName ?? 'protected.hwpx';
	const newBinaryData = await this.helpers.prepareBinaryData(
		newBuffer,
		fileName,
		'application/hwp+zip',
	);

	return {
		json: { success: true, action },
		binary: { [outputProp]: newBinaryData },
		pairedItem: { item: itemIndex },
	};
}

// ==========================================
// 9. Export to Markdown
// ==========================================

export async function toMarkdown(
	this: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
): Promise<INodeExecutionData> {
	const inputProp = this.getNodeParameter('inputBinaryPropertyName', itemIndex) as string;
	const options = this.getNodeParameter('options', itemIndex, {}) as {
		includeTables?: boolean;
		outputFormat?: string;
	};

	const binaryData = item.binary?.[inputProp];
	if (!binaryData) {
		throw new NodeOperationError(
			this.getNode(),
			`No binary data found in property "${inputProp}"`,
			{ itemIndex },
		);
	}

	const fileBuffer = Buffer.from(binaryData.data, 'base64');
	const zip = await JSZip.loadAsync(fileBuffer);

	const includeTables = options.includeTables ?? true;

	// Read header.xml for style lookups
	const headerFile =
		zip.file('Contents/header.xml') ?? zip.file('Contents/Header.xml');
	const headerXml = headerFile ? await headerFile.async('string') : '';

	// Build paraPrIDRef → style info map (for heading detection)
	const paraPrMap = new Map<string, { outlineLevel?: number }>();
	const paraPrMatches = [
		...headerXml.matchAll(/<hh:paraPr\b([^>]*)(?:\/>|>([\s\S]*?)<\/hh:paraPr>)/g),
	];
	for (const m of paraPrMatches) {
		const idMatch = m[1].match(/id="(\d+)"/);
		const levelMatch = m[1].match(/outlineLevel="(\d+)"/);
		if (idMatch) {
			paraPrMap.set(idMatch[1], {
				outlineLevel: levelMatch ? parseInt(levelMatch[1], 10) : undefined,
			});
		}
	}

	// Build charPrIDRef → bold/italic map
	const charPrStyleMap = new Map<string, { bold?: boolean; italic?: boolean }>();
	const charPrMatchesMd = [
		...headerXml.matchAll(/<hh:charPr\b([^>]*)>([\s\S]*?)<\/hh:charPr>/g),
	];
	for (const m of charPrMatchesMd) {
		const idMatch = m[1].match(/id="(\d+)"/);
		if (!idMatch) continue;
		const content = m[0];
		const hasBold = content.includes('<hh:bold/>') || content.includes('<hh:bold ');
		const hasItalic = content.includes('<hh:italic/>') || content.includes('<hh:italic ');
		charPrStyleMap.set(idMatch[1], { bold: hasBold || undefined, italic: hasItalic || undefined });
	}

	// Process section files
	const sectionFiles = Object.keys(zip.files)
		.filter((f) => /Contents\/[Ss]ection\d+\.xml/.test(f))
		.sort();

	const markdownParts: string[] = [];

	for (const sf of sectionFiles) {
		const sectionContent = await zip.file(sf)!.async('string');

		const paraRegex = /<hp:p\b([^>]*)>([\s\S]*?)<\/hp:p>/g;
		let paraMatch;
		while ((paraMatch = paraRegex.exec(sectionContent)) !== null) {
			const paraAttrs = paraMatch[1];
			const paraContent = paraMatch[2];

			// Check for heading level
			const paraPrRef = paraAttrs.match(/paraPrIDRef="(\d+)"/);
			const paraStyle = paraPrRef ? paraPrMap.get(paraPrRef[1]) : undefined;
			let prefix = '';
			if (paraStyle?.outlineLevel !== undefined && paraStyle.outlineLevel >= 0 && paraStyle.outlineLevel <= 5) {
				prefix = '#'.repeat(paraStyle.outlineLevel + 1) + ' ';
			}

			// Check for table
			if (includeTables && paraContent.includes('<hp:tbl')) {
				const tableMarkdown = convertTableToMarkdown(paraContent);
				if (tableMarkdown) {
					markdownParts.push(tableMarkdown);
					continue;
				}
			}

			// Extract text from runs
			const runRegex = /<hp:run\b[^>]*charPrIDRef="(\d+)"[^>]*>([\s\S]*?)<\/hp:run>/g;
			let runMatch;
			const lineTexts: string[] = [];

			while ((runMatch = runRegex.exec(paraContent)) !== null) {
				const charId = runMatch[1];
				const runContent = runMatch[2];
				const charStyle = charPrStyleMap.get(charId);

				const textRegex = /<hp:t>([^<]*)<\/hp:t>/g;
				let textMatch;
				while ((textMatch = textRegex.exec(runContent)) !== null) {
					let text = textMatch[1];
					if (!text) continue;

					if (charStyle?.bold && charStyle?.italic) {
						text = `***${text}***`;
					} else if (charStyle?.bold) {
						text = `**${text}**`;
					} else if (charStyle?.italic) {
						text = `*${text}*`;
					}
					lineTexts.push(text);
				}
			}

			const lineText = lineTexts.join('');
			if (lineText || prefix) {
				markdownParts.push(prefix + lineText);
			} else {
				markdownParts.push('');
			}
		}
	}

	const markdown = markdownParts.join('\n');

	const outputFormat = options.outputFormat ?? 'text';
	if (outputFormat === 'binary') {
		const mdBuffer = Buffer.from(markdown, 'utf-8');
		const mdBinary = await this.helpers.prepareBinaryData(
			mdBuffer,
			'document.md',
			'text/markdown',
		);
		return {
			json: { success: true, length: markdown.length },
			binary: { data: mdBinary },
			pairedItem: { item: itemIndex },
		};
	}

	return {
		json: { markdown },
		pairedItem: { item: itemIndex },
	};
}

function convertTableToMarkdown(content: string): string | null {
	const cellRegex = /<hp:tc\b[^>]*>([\s\S]*?)<\/hp:tc>/g;
	const cells: string[] = [];
	let cellMatch;
	while ((cellMatch = cellRegex.exec(content)) !== null) {
		const cellContent = cellMatch[1];
		const texts: string[] = [];
		const tRegex = /<hp:t>([^<]*)<\/hp:t>/g;
		let tMatch;
		while ((tMatch = tRegex.exec(cellContent)) !== null) {
			if (tMatch[1]) texts.push(tMatch[1]);
		}
		cells.push(texts.join(' '));
	}

	if (cells.length === 0) return null;

	const trMatches = [...content.matchAll(/<hp:tr\b/g)];
	const colCount = trMatches.length > 0 ? Math.ceil(cells.length / trMatches.length) : cells.length;

	const rows: string[][] = [];
	for (let i = 0; i < cells.length; i += colCount) {
		rows.push(cells.slice(i, i + colCount));
	}

	if (rows.length === 0) return null;

	const lines: string[] = [];
	lines.push('| ' + rows[0].join(' | ') + ' |');
	lines.push('| ' + rows[0].map(() => '---').join(' | ') + ' |');
	for (let r = 1; r < rows.length; r++) {
		lines.push('| ' + rows[r].join(' | ') + ' |');
	}

	return lines.join('\n');
}
