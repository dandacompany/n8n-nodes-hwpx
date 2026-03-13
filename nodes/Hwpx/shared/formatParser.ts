/**
 * Text formatting parser for HWPX document creation.
 * Supports Markdown and Structured JSON input formats.
 *
 * HWPML uses a centralized styling model:
 *   - header.xml defines charPr entries (each with a unique ID)
 *   - section XML references them via charPrIDRef on <hp:run>
 */

export interface TextRun {
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strikeout?: boolean;
	fontSize?: number; // in points (default: 10)
	textColor?: string; // hex color e.g. "#FF0000"
}

export interface Paragraph {
	runs: TextRun[];
	heading?: number; // 1-3 for heading levels
}

const HEADING_SIZES: Record<number, number> = {
	1: 16,
	2: 14,
	3: 12,
};
const DEFAULT_FONT_SIZE = 10;

// ─── Markdown Parser ───────────────────────────────────────

/**
 * Parse Markdown text into structured paragraphs with styled runs.
 *
 * Supported syntax:
 *   - `# Heading 1`, `## Heading 2`, `### Heading 3`
 *   - `***bold+italic***`
 *   - `**bold**`
 *   - `*italic*`
 *   - `~~strikeout~~`
 */
export function parseMarkdown(text: string): Paragraph[] {
	const lines = text.split('\n');
	const paragraphs: Paragraph[] = [];

	for (const line of lines) {
		let heading: number | undefined;
		let content = line;

		const headingMatch = content.match(/^(#{1,3})\s+(.*)$/);
		if (headingMatch) {
			heading = headingMatch[1].length;
			content = headingMatch[2];
		}

		const runs = parseInlineFormatting(content);

		if (heading) {
			const size = HEADING_SIZES[heading] || DEFAULT_FONT_SIZE;
			for (const run of runs) {
				if (!run.fontSize) run.fontSize = size;
				run.bold = true;
			}
		}

		paragraphs.push({ runs, heading });
	}

	return paragraphs;
}

/**
 * Parse inline markdown formatting into TextRuns.
 * Order of precedence: ***bold+italic*** > **bold** > *italic* > ~~strikeout~~
 */
function parseInlineFormatting(text: string): TextRun[] {
	if (!text) return [{ text: '' }];

	const runs: TextRun[] = [];
	const pattern = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|~~(.+?)~~)/g;

	let lastIndex = 0;
	let match: RegExpExecArray | null;

	while ((match = pattern.exec(text)) !== null) {
		if (match.index > lastIndex) {
			runs.push({ text: text.slice(lastIndex, match.index) });
		}

		if (match[2] !== undefined) {
			runs.push({ text: match[2], bold: true, italic: true });
		} else if (match[3] !== undefined) {
			runs.push({ text: match[3], bold: true });
		} else if (match[4] !== undefined) {
			runs.push({ text: match[4], italic: true });
		} else if (match[5] !== undefined) {
			runs.push({ text: match[5], strikeout: true });
		}

		lastIndex = match.index + match[0].length;
	}

	if (lastIndex < text.length) {
		runs.push({ text: text.slice(lastIndex) });
	}

	if (runs.length === 0) {
		runs.push({ text: '' });
	}

	return runs;
}

// ─── Structured JSON Parser ────────────────────────────────

/**
 * Parse structured JSON into paragraphs.
 *
 * Expected format:
 * {
 *   "paragraphs": [
 *     {
 *       "runs": [
 *         { "text": "Bold text", "bold": true, "fontSize": 12 },
 *         { "text": " normal" }
 *       ]
 *     }
 *   ]
 * }
 */
export function parseStructuredJson(json: string | object): Paragraph[] {
	const data = typeof json === 'string' ? JSON.parse(json as string) : json;

	if (!data.paragraphs || !Array.isArray(data.paragraphs)) {
		throw new Error('Structured JSON must have a "paragraphs" array');
	}

	return (data.paragraphs as Array<Record<string, unknown>>).map(
		(p: Record<string, unknown>) => ({
			runs: (
				(p.runs as Array<Record<string, unknown>>) || []
			).map((r: Record<string, unknown>) => ({
				text: (r.text as string) || '',
				bold: (r.bold as boolean) || false,
				italic: (r.italic as boolean) || false,
				underline: (r.underline as boolean) || false,
				strikeout: (r.strikeout as boolean) || false,
				fontSize: (r.fontSize as number) || undefined,
				textColor: (r.textColor as string) || undefined,
			})),
			heading: (p.heading as number) || undefined,
		}),
	);
}

// ─── Style Collection ──────────────────────────────────────

/**
 * Generate a unique key for a style combination (for deduplication).
 */
function styleKey(run: TextRun): string {
	return [
		run.bold ? 'B' : '',
		run.italic ? 'I' : '',
		run.underline ? 'U' : '',
		run.strikeout ? 'S' : '',
		`sz${run.fontSize || DEFAULT_FONT_SIZE}`,
		`c${(run.textColor || '#000000').toUpperCase()}`,
	].join('|');
}

/**
 * Collect unique styles from paragraphs and assign charPr IDs.
 * IDs start from startId to avoid conflicts with template charPr entries.
 */
export function collectStyles(
	paragraphs: Paragraph[],
	startId: number,
): Map<string, { id: number; style: TextRun }> {
	const styles = new Map<string, { id: number; style: TextRun }>();
	let nextId = startId;

	// Default plain style maps to charPr id=0 in the template
	const defaultRun: TextRun = { text: '', fontSize: DEFAULT_FONT_SIZE, textColor: '#000000' };
	styles.set(styleKey(defaultRun), { id: 0, style: defaultRun });

	for (const para of paragraphs) {
		for (const run of para.runs) {
			const key = styleKey(run);
			if (!styles.has(key)) {
				styles.set(key, { id: nextId++, style: { ...run } });
			}
		}
	}

	return styles;
}

// ─── HWPML XML Generation ──────────────────────────────────

/**
 * Generate HWPML charPr XML entries for dynamically created styles.
 * Skips id=0 (already in template).
 */
export function buildCharPrEntries(
	styles: Map<string, { id: number; style: TextRun }>,
): string {
	let xml = '';

	for (const [, { id, style }] of styles) {
		if (id === 0) continue;

		const height = (style.fontSize || DEFAULT_FONT_SIZE) * 100;
		const textColor = style.textColor || '#000000';

		xml += `<hh:charPr id="${id}" height="${height}" textColor="${textColor}" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">`;
		xml += `<hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>`;
		xml += `<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>`;
		xml += `<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>`;
		xml += `<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>`;
		xml += `<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>`;

		if (style.bold) xml += `<hh:bold/>`;
		if (style.italic) xml += `<hh:italic/>`;

		xml += style.underline
			? `<hh:underline type="BOTTOM" shape="SOLID" color="${textColor}"/>`
			: `<hh:underline type="NONE" shape="SOLID" color="#000000"/>`;

		xml += style.strikeout
			? `<hh:strikeout shape="SOLID" color="${textColor}"/>`
			: `<hh:strikeout shape="NONE" color="#000000"/>`;

		xml += `<hh:outline type="NONE"/>`;
		xml += `<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>`;
		xml += `</hh:charPr>`;
	}

	return xml;
}

/**
 * Inject new charPr entries into header.xml content.
 * Updates the itemCnt attribute and appends entries before closing tag.
 */
export function injectCharProperties(
	headerXml: string,
	newCharPrXml: string,
	totalCount: number,
): string {
	const updated = headerXml.replace(
		/<hh:charProperties itemCnt="\d+">/,
		`<hh:charProperties itemCnt="${totalCount}">`,
	);

	return updated.replace(
		'</hh:charProperties>',
		newCharPrXml + '</hh:charProperties>',
	);
}

/**
 * Build styled section XML from paragraphs with charPrIDRef references.
 */
export function buildStyledSectionXml(
	paragraphs: Paragraph[],
	styles: Map<string, { id: number; style: TextRun }>,
): string {
	let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
	xml += `<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"`;
	xml += ` xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">`;

	for (const para of paragraphs) {
		xml += `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">`;

		for (const run of para.runs) {
			const key = styleKey(run);
			const entry = styles.get(key);
			const charPrId = entry ? entry.id : 0;

			const escaped = run.text
				.replace(/&/g, '&amp;')
				.replace(/</g, '&lt;')
				.replace(/>/g, '&gt;')
				.replace(/"/g, '&quot;');

			xml += `<hp:run charPrIDRef="${charPrId}">`;
			xml += `<hp:rPr/>`;
			xml += `<hp:t>${escaped}</hp:t>`;
			xml += `</hp:run>`;
		}

		xml += `</hp:p>`;
	}

	xml += `</hs:sec>`;
	return xml;
}
