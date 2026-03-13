import { build } from 'esbuild';
import { readdirSync, statSync } from 'fs';
import { join, relative } from 'path';

// Recursively find all .ts files in nodes/
function findTsFiles(dir) {
	const files = [];
	for (const entry of readdirSync(dir)) {
		const full = join(dir, entry);
		if (statSync(full).isDirectory()) {
			files.push(...findTsFiles(full));
		} else if (full.endsWith('.ts') && !full.endsWith('.d.ts')) {
			files.push(full);
		}
	}
	return files;
}

const entryPoints = findTsFiles('nodes');

await build({
	entryPoints,
	bundle: true,
	platform: 'node',
	target: 'node18',
	format: 'cjs',
	outdir: 'dist',
	// Bundle jszip, @ssabrojs/hwpxjs, fast-xml-parser into the output
	external: ['n8n-workflow'],
	sourcemap: true,
	// Preserve directory structure relative to project root
	outbase: '.',
});

console.log(`Bundled ${entryPoints.length} files with jszip + hwpxjs inlined`);
