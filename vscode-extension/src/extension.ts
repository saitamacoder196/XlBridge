import * as vscode from 'vscode';
import * as cp from 'child_process';
import * as util from 'util';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';

const exec = util.promisify(cp.exec);

const PARTICIPANT_ID = 'copatis.assistant';

// ─── Logger ───────────────────────────────────────────────────────────────────
// Visible at: View → Output → "Copatis"
// Format: [HH:MM:SS.mmm] LEVEL [Operation] message  key=value …

type LogLevel = 'INFO' | 'WARN' | 'ERROR';

class Logger {
    private readonly ch: vscode.OutputChannel;

    constructor() {
        this.ch = vscode.window.createOutputChannel('Copatis');
    }

    /** Print a horizontal separator before each top-level command. */
    section(label: string): void {
        const pad = Math.max(0, 50 - label.length - 4);
        this.ch.appendLine(`\n--- ${label} ${'-'.repeat(pad)}`);
    }

    info(op: string, msg: string, kv?: Record<string, string | number>): void {
        this.write('INFO ', op, msg, kv);
    }

    warn(op: string, msg: string, kv?: Record<string, string | number>): void {
        this.write('WARN ', op, msg, kv);
    }

    error(op: string, msg: string, kv?: Record<string, string | number>): void {
        this.write('ERROR', op, msg, kv);
    }

    /** Reveal the Output panel (preserves editor focus). */
    show(): void { this.ch.show(true); }

    dispose(): void { this.ch.dispose(); }

    // ── internals ──

    private ts(): string {
        const d = new Date();
        const hh = String(d.getHours()).padStart(2, '0');
        const mm = String(d.getMinutes()).padStart(2, '0');
        const ss = String(d.getSeconds()).padStart(2, '0');
        const ms = String(d.getMilliseconds()).padStart(3, '0');
        return `${hh}:${mm}:${ss}.${ms}`;
    }

    private write(
        level: LogLevel | string,
        op: string,
        msg: string,
        kv?: Record<string, string | number>,
    ): void {
        const opTag = op.padEnd(9);  // fixed-width column
        const kvStr = kv
            ? '  ' + Object.entries(kv).map(([k, v]) => `${k}=${JSON.stringify(String(v))}`).join('  ')
            : '';
        this.ch.appendLine(`[${this.ts()}] ${level} [${opTag}] ${msg}${kvStr}`);
    }
}

// Module-level singleton — created once in activate()
let log: Logger;

// ─── Dictionary ───────────────────────────────────────────────────────────────
// Persistent JSON map: Japanese source text → target language text.
// File format: { "日本語": "translation", ... } sorted by key.

class Dictionary {
    private readonly data = new Map<string, string>();
    private added = 0;   // entries added since construction

    constructor(private readonly filePath: string) {
        this.load();
    }

    lookup(key: string): string | undefined {
        return this.data.get(key);
    }

    set(key: string, value: string): void {
        if (!this.data.has(key)) this.added++;
        this.data.set(key, value);
    }

    get size(): number  { return this.data.size; }
    get newCount(): number { return this.added; }

    /** Write to disk. Returns number of new entries added. Idempotent. */
    save(): number {
        try {
            const obj: Record<string, string> = {};
            for (const k of [...this.data.keys()].sort()) { obj[k] = this.data.get(k)!; }
            fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
            fs.writeFileSync(this.filePath, JSON.stringify(obj, null, 2), 'utf-8');
            log.info('Dict', 'Saved', { path: this.filePath, total: this.data.size, new: this.added });
        } catch (err) {
            log.error('Dict', 'Save failed', { path: this.filePath, error: String(err) });
        }
        return this.added;
    }

    private load(): void {
        if (!fs.existsSync(this.filePath)) return;
        try {
            const obj = JSON.parse(fs.readFileSync(this.filePath, 'utf-8')) as Record<string, string>;
            for (const [k, v] of Object.entries(obj)) { this.data.set(k, v); }
        } catch (err) {
            log.warn('Dict', 'Load failed — starting empty', { path: this.filePath, error: String(err) });
        }
    }
}

// ─── Pattern Library ──────────────────────────────────────────────────────────
// Persistent JSON array of regex-based translation templates.
// File format: [{ "regex": "...", "en_template": "...", "vi_template": "..." }, ...]
// Capture groups in regex map to {0}, {1}, ... placeholders in templates.

interface PatternEntry {
    regex: string;        // ECMAScript regex with capture groups for variable parts
    en_template: string;  // English template, e.g. "Order No.: {0}"
    vi_template: string;  // Vietnamese template, e.g. "Số đơn hàng: {0}"
}

class PatternLibrary {
    private entries: PatternEntry[] = [];
    private added = 0;

    constructor(private readonly filePath: string) {
        this.load();
    }

    /** Try to match value against all patterns. Returns null if no match. */
    match(value: string): { entry: PatternEntry; groups: string[] } | null {
        for (const entry of this.entries) {
            try {
                const m = new RegExp(entry.regex).exec(value);
                if (m) { return { entry, groups: m.slice(1) }; }
            } catch { /* invalid regex — skip */ }
        }
        return null;
    }

    add(entry: PatternEntry): void {
        if (!this.entries.some(e => e.regex === entry.regex)) {
            this.entries.push(entry);
            this.added++;
        }
    }

    get size(): number  { return this.entries.length; }
    get newCount(): number { return this.added; }

    save(): number {
        try {
            fs.mkdirSync(path.dirname(this.filePath), { recursive: true });
            fs.writeFileSync(this.filePath, JSON.stringify(this.entries, null, 2), 'utf-8');
            log.info('Pattern', 'Saved', { path: this.filePath, total: this.entries.length, new: this.added });
        } catch (err) {
            log.error('Pattern', 'Save failed', { path: this.filePath, error: String(err) });
        }
        return this.added;
    }

    private load(): void {
        if (!fs.existsSync(this.filePath)) return;
        try {
            const arr = JSON.parse(fs.readFileSync(this.filePath, 'utf-8')) as PatternEntry[];
            if (Array.isArray(arr)) { this.entries = arr; }
        } catch (err) {
            log.warn('Pattern', 'Load failed — starting empty', { path: this.filePath, error: String(err) });
        }
    }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function workspacePath(): string | undefined {
    return vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
}

function xlbridgeCmd(): string {
    const cfg = vscode.workspace.getConfiguration('copatis');
    const python: string = cfg.get('pythonPath') || 'python';
    return `${python} -m xlbridge`;
}

async function runXlbridge(args: string[]): Promise<{ stdout: string; stderr: string }> {
    const cwd = workspacePath() || process.cwd();
    const cmd = `${xlbridgeCmd()} ${args.join(' ')}`;
    log.info('CLI', 'Exec', { cmd });
    return exec(cmd, { cwd });
}

async function findExcelFiles(): Promise<vscode.Uri[]> {
    return vscode.workspace.findFiles('**/*.xlsx', '{**/node_modules/**,**/.git/**}');
}

async function findTxtFiles(): Promise<vscode.Uri[]> {
    return vscode.workspace.findFiles('**/*.txt', '{**/node_modules/**,**/.git/**}');
}

/** Match a filename (supports Unicode) from a prompt string. */
function parseFilenameArg(prompt: string, ext: string): string | undefined {
    const re = new RegExp(`[^\\s"']+\\.${ext}`, 'i');
    return re.exec(prompt)?.[0];
}

function elapsed(startMs: number): string {
    return `${Date.now() - startMs}ms`;
}

function getCfg<T>(key: string, fallback: T): T {
    return vscode.workspace.getConfiguration('copatis').get<T>(key) ?? fallback;
}

function getBatchSize(): number {
    return Math.max(1, Math.min(200, getCfg<number>('batchSize', 25)));
}

function getMaxRetries(): number {
    return Math.max(1, Math.min(10, getCfg<number>('llmMaxRetries', 3)));
}

function getDictDir(): string {
    const configured = getCfg<string>('dictDir', '').trim();
    if (configured) return configured;
    const ws = workspacePath();
    return ws ? path.join(ws, 'copatis_dicts') : path.join(os.homedir(), 'copatis_dicts');
}

// ─── Language helpers ─────────────────────────────────────────────────────────

type TargetLang = 'en' | 'vi';

/**
 * Translated line format  : [Sheet]!Cell|Original|EN|VI
 * Original (no-translate) : [Sheet]!Cell|Original
 *
 * Returns { lang, usedFallback } where usedFallback=true means
 * the target-language column was missing → fell back to original.
 */
interface SelectResult {
    value: string;
    usedFallback: boolean;
}

function selectValue(rawAfterCell: string, lang: TargetLang): SelectResult {
    const parts = rawAfterCell.split('|');
    // parts[0]=Original, parts[1]=EN, parts[2]=VI
    if (lang === 'en') {
        const en = parts[1]?.trim();
        return en
            ? { value: en, usedFallback: false }
            : { value: parts[0], usedFallback: true };
    } else {
        const vi = parts[2]?.trim();
        return vi
            ? { value: vi, usedFallback: false }
            : { value: parts[0], usedFallback: true };
    }
}

/** Detect EN / VI from the user prompt. Returns undefined if not found. */
function detectLangFromPrompt(prompt: string): TargetLang | undefined {
    const p = prompt.toLowerCase();
    if (/\ben\b|english|tiếng[\s-]?anh/.test(p)) return 'en';
    if (/\bvi\b|vietnamese|tiếng[\s-]?việt/.test(p)) return 'vi';
    return undefined;
}

/** Detect whether the user is asking for an unsupported language. */
function detectOtherLang(prompt: string): string | undefined {
    const p = prompt.toLowerCase();
    const mapping: [RegExp, string][] = [
        [/\bjapanese\b|日本語|\bja\b|\bjp\b/, 'Tiếng Nhật'],
        [/\bchinese\b|中文|\bzh\b/, 'Tiếng Trung'],
        [/\bkorean\b|한국어|\bko\b/, 'Tiếng Hàn'],
        [/\bfrench\b|français|\bfr\b/, 'Tiếng Pháp'],
        [/\bgerman\b|deutsch|\bde\b/, 'Tiếng Đức'],
        [/\bspanish\b|español|\bes\b/, 'Tiếng Tây Ban Nha'],
        [/\bjapan\b/, 'Tiếng Nhật'],
    ];
    for (const [re, name] of mapping) {
        if (re.test(p)) return name;
    }
    return undefined;
}

/** Show a QuickPick and let user choose the target language. */
async function pickTargetLang(): Promise<TargetLang | undefined> {
    type Item = vscode.QuickPickItem & { lang: TargetLang };
    const items: Item[] = [
        { label: '$(globe) Tiếng Anh (EN)', description: 'Inject cột bản dịch Tiếng Anh', lang: 'en' },
        { label: '$(globe) Tiếng Việt (VI)', description: 'Inject cột bản dịch Tiếng Việt', lang: 'vi' },
    ];
    const picked = await vscode.window.showQuickPick(items, {
        placeHolder: 'Chọn ngôn ngữ đích để inject vào Excel (chỉ hỗ trợ EN và VI)',
        title: 'Copatis — Ngôn ngữ đích',
    });
    return picked?.lang;
}

/**
 * Transform a TXT file (original or translated) into a minimal
 * [Sheet]!Cell|Value file ready for `xlbridge inject`.
 *
 * Returns { content, stats }.
 */
interface TransformStats {
    total: number;
    translated: number;
    fallback: number;       // had no target-lang column → used original
    commentOrBlank: number;
}

function buildInjectContent(
    lines: string[],
    lang: TargetLang,
): { content: string; stats: TransformStats } {
    const stats: TransformStats = { total: 0, translated: 0, fallback: 0, commentOrBlank: 0 };
    const out: string[] = [];

    for (const raw of lines) {
        const line = raw.trimEnd();

        if (line.startsWith('#') || line === '') {
            stats.commentOrBlank++;
            out.push(line);
            continue;
        }

        const pipeIdx = line.indexOf('|');
        if (pipeIdx === -1) {
            out.push(line);     // malformed — pass through
            continue;
        }

        const cellRef      = line.slice(0, pipeIdx);
        const afterCell    = line.slice(pipeIdx + 1);
        const { value, usedFallback } = selectValue(afterCell, lang);

        stats.total++;
        if (usedFallback) { stats.fallback++; } else { stats.translated++; }

        out.push(`${cellRef}|${value}`);
    }

    return { content: out.join('\n'), stats };
}

// ─── /extract ─────────────────────────────────────────────────────────────────

async function handleExtract(
    request: vscode.ChatRequest,
    stream: vscode.ChatResponseStream,
    token: vscode.CancellationToken,
): Promise<void> {
    const t0 = Date.now();
    log.section('Extract');
    log.info('Extract', 'Start', { prompt: request.prompt || '(none)' });
    log.show();

    stream.markdown('### Copatis — Extract\n\n');

    // ── 1. Resolve input file ──
    const specifiedFile = parseFilenameArg(request.prompt, 'xlsx');
    let targetPath: string | undefined;

    if (specifiedFile) {
        const ws = workspacePath();
        const resolved = path.isAbsolute(specifiedFile)
            ? specifiedFile
            : ws ? path.join(ws, specifiedFile) : specifiedFile;
        if (fs.existsSync(resolved)) {
            targetPath = resolved;
            log.info('Extract', 'Resolved', { from: 'prompt', path: resolved });
        } else {
            log.warn('Extract', 'File not found', { specified: specifiedFile });
            stream.markdown(`❌ Không tìm thấy file \`${specifiedFile}\`\n\n`);
        }
    }

    if (!targetPath) {
        log.info('Extract', 'Scanning workspace for .xlsx');
        const xlsxFiles = await findExcelFiles();
        log.info('Extract', 'Scan result', { found: xlsxFiles.length });

        if (xlsxFiles.length === 0) {
            log.warn('Extract', 'No .xlsx files in workspace');
            stream.markdown(
                '❌ Không tìm thấy file `.xlsx` nào trong workspace.\n\n'
                + '```\n@copatis /extract ten-file.xlsx\n```',
            );
            return;
        }
        if (xlsxFiles.length === 1) {
            targetPath = xlsxFiles[0].fsPath;
            log.info('Extract', 'Auto-selected', { path: targetPath });
        } else {
            log.info('Extract', 'Multiple files found — awaiting user selection');
            stream.markdown('📁 Tìm thấy nhiều file Excel trong workspace:\n\n');
            xlsxFiles.forEach(f => stream.markdown(`- \`${vscode.workspace.asRelativePath(f)}\`\n`));
            stream.markdown('\n> ```\n> @copatis /extract ten-file.xlsx\n> ```');
            return;
        }
    }

    // ── 2. Build args ──
    const sheetMatch = request.prompt.match(/--sheet\s+(\S+)/g);
    const sheetArgs = sheetMatch
        ? sheetMatch.flatMap(m => ['--sheet', m.replace('--sheet ', '')])
        : [];

    const outputPath = targetPath.replace(/\.xlsx?$/i, '_export.txt');
    const relInput  = vscode.workspace.asRelativePath(targetPath);
    const relOutput = vscode.workspace.asRelativePath(outputPath);

    log.info('Extract', 'Plan', {
        input:  relInput,
        output: relOutput,
        sheets: sheetArgs.filter((_, i) => i % 2 !== 0).join(',') || 'all',
    });

    stream.markdown(`📊 Input : \`${relInput}\`\n`);
    stream.markdown(`📝 Output: \`${relOutput}\`\n\n`);
    stream.progress('Đang chạy xlbridge extract...');

    // ── 3. Run ──
    try {
        const { stdout, stderr } = await runXlbridge([
            'extract',
            '--input',  `"${targetPath}"`,
            '--output', `"${outputPath}"`,
            ...sheetArgs,
        ]);

        const cellMatch = (stdout + stderr).match(/Extracted (\d+) cells/);
        const cells = cellMatch ? parseInt(cellMatch[1]) : '?';

        log.info('Extract', 'Done', { cells, elapsed: elapsed(t0) });
        if (stderr.trim()) log.info('Extract', 'Stderr', { msg: stderr.trim() });

        stream.markdown('✅ **Extract thành công!**\n\n');
        if (stdout) stream.markdown(`\`\`\`\n${stdout.trim()}\n\`\`\`\n\n`);
        if (stderr) stream.markdown(`> ℹ️ ${stderr.trim()}\n\n`);

        if (fs.existsSync(outputPath)) {
            stream.button({
                command: 'vscode.open',
                arguments: [vscode.Uri.file(outputPath)],
                title: '📂 Mở file output',
            });
        }
    } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        log.error('Extract', 'Failed', { elapsed: elapsed(t0), error: msg.split('\n')[0] });
        stream.markdown(
            `❌ **Lỗi khi chạy xlbridge:**\n\`\`\`\n${msg}\n\`\`\`\n\n`
            + '💡 Đảm bảo XlBridge đã được cài đặt: `pip install -e .`',
        );
    }
}

// ─── /inject ──────────────────────────────────────────────────────────────────

async function handleInject(
    request: vscode.ChatRequest,
    stream: vscode.ChatResponseStream,
    token: vscode.CancellationToken,
): Promise<void> {
    const t0 = Date.now();
    log.section('Inject');
    log.info('Inject', 'Start', { prompt: request.prompt || '(none)' });
    log.show();

    stream.markdown('### Copatis — Inject\n\n');

    // ── 1. Resolve target language ──

    // Check for unsupported language first
    const otherLang = detectOtherLang(request.prompt);
    if (otherLang) {
        log.warn('Inject', 'Unsupported language requested', { lang: otherLang });
        stream.markdown(
            `❌ **Ngôn ngữ chưa được hỗ trợ: ${otherLang}**\n\n`
            + 'Copatis hiện chỉ hỗ trợ inject cho:\n'
            + '- **Tiếng Anh (EN)**: `@copatis /inject en`\n'
            + '- **Tiếng Việt (VI)**: `@copatis /inject vi`\n',
        );
        return;
    }

    let lang = detectLangFromPrompt(request.prompt);
    if (!lang) {
        log.info('Inject', 'Language not in prompt — showing QuickPick');
        lang = await pickTargetLang();
    }
    if (!lang) {
        log.warn('Inject', 'User cancelled language selection');
        stream.markdown('❌ Chưa chọn ngôn ngữ đích. Hãy thử:\n```\n@copatis /inject en\n@copatis /inject vi\n```');
        return;
    }

    const langLabel = lang === 'en' ? 'Tiếng Anh (EN)' : 'Tiếng Việt (VI)';
    log.info('Inject', 'Language selected', { lang, label: langLabel });
    stream.markdown(`🌐 Ngôn ngữ đích: **${langLabel}**\n\n`);

    // ── 2. Resolve files ──

    const specifiedXlsx = parseFilenameArg(request.prompt, 'xlsx');
    const specifiedTxt  = parseFilenameArg(request.prompt, 'txt');
    const ws = workspacePath() || '';

    log.info('Inject', 'Scanning workspace');
    const [xlsxFiles, txtFiles] = await Promise.all([findExcelFiles(), findTxtFiles()]);
    log.info('Inject', 'Scan result', { xlsx: xlsxFiles.length, txt: txtFiles.length });

    if (xlsxFiles.length === 0) {
        log.warn('Inject', 'No .xlsx files found');
        stream.markdown('❌ Không tìm thấy file `.xlsx` trong workspace.');
        return;
    }
    if (txtFiles.length === 0) {
        log.warn('Inject', 'No .txt files found');
        stream.markdown('❌ Không tìm thấy file `.txt` (bản dịch) trong workspace.');
        return;
    }

    const resolveFile = (name: string | undefined, list: vscode.Uri[]): string | undefined => {
        if (!name) return list.length === 1 ? list[0].fsPath : undefined;
        const abs = path.isAbsolute(name) ? name : path.join(ws, name);
        return fs.existsSync(abs) ? abs : list.find(f => f.fsPath.endsWith(name))?.fsPath;
    };

    const xlsxPath = resolveFile(specifiedXlsx, xlsxFiles);
    const txtPath  = resolveFile(specifiedTxt,  txtFiles);

    if (!xlsxPath || !txtPath) {
        log.warn('Inject', 'Cannot auto-resolve files — user selection required', {
            xlsx: specifiedXlsx ?? '(unspecified)',
            txt:  specifiedTxt  ?? '(unspecified)',
        });
        stream.markdown('📁 **File Excel tìm thấy:**\n');
        xlsxFiles.forEach(f => stream.markdown(`- \`${vscode.workspace.asRelativePath(f)}\`\n`));
        stream.markdown('\n📝 **File translation tìm thấy:**\n');
        txtFiles.slice(0, 8).forEach(f => stream.markdown(`- \`${vscode.workspace.asRelativePath(f)}\`\n`));
        stream.markdown('\n> ```\n> @copatis /inject en file.xlsx translated.txt\n> ```');
        return;
    }

    // ── 3. Transform TXT → target language ──

    log.info('Inject', 'Reading translation file', { path: vscode.workspace.asRelativePath(txtPath) });
    let rawContent: string;
    try {
        rawContent = fs.readFileSync(txtPath, 'utf-8');
    } catch (err) {
        log.error('Inject', 'Cannot read TXT file', { error: String(err) });
        stream.markdown(`❌ Không đọc được file TXT: ${err}`);
        return;
    }

    const lines = rawContent.replace(/\r\n/g, '\n').split('\n');
    const { content: injectContent, stats } = buildInjectContent(lines, lang);

    log.info('Inject', 'Transform done', {
        lang,
        total:      stats.total,
        translated: stats.translated,
        fallback:   stats.fallback,
    });

    if (stats.fallback > 0) {
        log.warn('Inject', 'Some cells missing target-lang column — used original', {
            fallback: stats.fallback,
            hint: 'Run @copatis /translate first to generate EN/VI columns',
        });
    }

    // ── 4. Write temp file ──

    const tmpFile = path.join(os.tmpdir(), `copatis_inject_${lang}_${Date.now()}.txt`);
    try {
        fs.writeFileSync(tmpFile, injectContent, 'utf-8');
        log.info('Inject', 'Temp file written', { path: tmpFile });
    } catch (err) {
        log.error('Inject', 'Cannot write temp file', { path: tmpFile, error: String(err) });
        stream.markdown(`❌ Không tạo được file tạm: ${err}`);
        return;
    }

    // ── 5. Build output path & log plan ──

    const langSuffix  = lang === 'en' ? '_en' : '_vi';
    const outputPath  = xlsxPath.replace(/\.xlsx?$/i, `${langSuffix}.xlsx`);

    log.info('Inject', 'Plan', {
        xlsx:       vscode.workspace.asRelativePath(xlsxPath),
        txt:        vscode.workspace.asRelativePath(txtPath),
        tmpFile:    path.basename(tmpFile),
        output:     vscode.workspace.asRelativePath(outputPath),
        translated: stats.translated,
        fallback:   stats.fallback,
    });

    stream.markdown(`📊 Excel  : \`${vscode.workspace.asRelativePath(xlsxPath)}\`\n`);
    stream.markdown(`📝 TXT    : \`${vscode.workspace.asRelativePath(txtPath)}\`\n`);
    stream.markdown(`📤 Output : \`${vscode.workspace.asRelativePath(outputPath)}\`\n`);
    stream.markdown(`📋 Cells  : ${stats.translated} đã dịch`
        + (stats.fallback > 0 ? `, ${stats.fallback} dùng bản gốc (chưa có ${langLabel})` : '')
        + '\n\n');

    if (stats.fallback > 0) {
        stream.markdown(
            `> ⚠️ **${stats.fallback} ô** chưa có bản dịch ${langLabel} → dùng văn bản gốc.\n`
            + '> Chạy `@copatis /translate` trước để tạo đủ bản dịch.\n\n',
        );
    }

    stream.progress(`Đang inject ${langLabel} vào Excel...`);

    // ── 6. Run xlbridge inject ──

    try {
        const { stdout, stderr } = await runXlbridge([
            'inject',
            '--input',       `"${xlsxPath}"`,
            '--translation', `"${tmpFile}"`,
            '--output',      `"${outputPath}"`,
        ]);

        const cellMatch = (stdout + stderr).match(/Injected (\d+)\/(\d+)/);
        const injected  = cellMatch ? `${cellMatch[1]}/${cellMatch[2]}` : `${stats.translated}/?`;

        log.info('Inject', 'Done', { lang, cells: injected, elapsed: elapsed(t0) });
        if (stderr.trim()) log.info('Inject', 'Stderr', { msg: stderr.trim() });

        stream.markdown('✅ **Inject thành công!**\n\n');
        if (stdout) stream.markdown(`\`\`\`\n${stdout.trim()}\n\`\`\`\n\n`);
        if (stderr) stream.markdown(`> ℹ️ ${stderr.trim()}\n\n`);

        if (fs.existsSync(outputPath)) {
            stream.button({
                command: 'revealInExplorer',
                arguments: [vscode.Uri.file(outputPath)],
                title: '📁 Hiện trong Explorer',
            });
        }
    } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        log.error('Inject', 'Failed', { elapsed: elapsed(t0), error: msg.split('\n')[0] });
        stream.markdown(`❌ **Lỗi:** \`\`\`\n${msg}\n\`\`\``);
    } finally {
        // Clean up temp file
        try { fs.unlinkSync(tmpFile); log.info('Inject', 'Temp file removed'); }
        catch { /* non-critical */ }
    }
}

// ─── /translate ───────────────────────────────────────────────────────────────

interface LLMTranslation { type: 'translation'; en: string; vi: string; }
interface LLMPattern    { type: 'pattern'; regex: string; en_template: string; vi_template: string; en: string; vi: string; }
type LLMResult = LLMTranslation | LLMPattern;

function parseDataLine(line: string): { prefix: string; value: string } | undefined {
    // Matches cell (A1), shape (shape:Name), and note (note:A1) address types.
    const match = line.match(/^(\[[^\]]+\]!(?:[A-Za-z]+\d+|shape:[^|]+|note:[A-Za-z]+\d+))\|(.+)$/);
    if (!match) return undefined;
    return { prefix: match[1], value: match[2] };
}

/** Escape literal newlines → \\n so each entry stays on a single line. */
function escNl(s: string): string { return s.replace(/\r?\n/g, '\\n'); }

/**
 * Returns true only if the value contains Japanese/CJK characters that need
 * translation. Values that are already English, numeric, date/time, symbolic,
 * or a mix of ASCII + punctuation are passed through unchanged.
 *
 * Covered ranges:
 *   U+3040–309F  Hiragana
 *   U+30A0–30FF  Katakana (full-width)
 *   U+31F0–31FF  Katakana Phonetic Extensions
 *   U+FF65–FF9F  Halfwidth Katakana
 *   U+4E00–9FFF  CJK Unified Ideographs (kanji)
 *   U+3400–4DBF  CJK Extension A
 *   U+F900–FAFF  CJK Compatibility Ideographs
 */
const JP_PATTERN = /[\u3040-\u309F\u30A0-\u30FF\u31F0-\u31FF\uFF65-\uFF9F\u4E00-\u9FFF\u3400-\u4DBF\uF900-\uFAFF]/;

/**
 * Strings consisting ONLY of these characters pass through unchanged.
 * These are common Japanese business/IPA document symbols that carry no
 * translatable meaning — they are status marks, punctuation, or decorative marks.
 *
 *   〇○●◎◯  — circle marks (OK / applicable / excellent)
 *   △▲▽▼    — triangle marks (partial / conditional)
 *   □■◆◇    — square / diamond marks
 *   ×✕✓✔   — cross / checkmark
 *   ★☆       — star ratings
 *   ー (U+30FC) — katakana long-vowel mark, used as a dash
 *   ・ (U+30FB) — katakana middle dot, used as a bullet/separator
 *   ゠ヽヾゝゞ — katakana/hiragana iteration marks
 *   ━―─…    — dash / horizontal rule / ellipsis
 */
const SYMBOL_ONLY = /^[\s〇○●◎◯△▲▽▼□■◆◇×✕✓✔★☆\u30FB\u30FC\u30A0\u30FD\u30FE\u30FF\u309D\u309E━―─…]+$/;

function needsTranslation(value: string): boolean {
    if (!value.trim()) return false;            // blank / whitespace-only
    if (SYMBOL_ONLY.test(value)) return false;  // document symbols — no translation needed
    return JP_PATTERN.test(value);
}

// ── Low-level LLM call ────────────────────────────────────────────────────────

function validateLLMResults(raw: unknown, expected: number): asserts raw is LLMResult[] {
    if (!Array.isArray(raw)) {
        throw new Error(`Response is not an array (got ${typeof raw})`);
    }
    if (raw.length !== expected) {
        throw new Error(`Length mismatch: expected ${expected}, got ${raw.length}`);
    }
    for (let i = 0; i < raw.length; i++) {
        const item = raw[i] as Record<string, unknown>;
        if (typeof item !== 'object' || item === null) {
            throw new Error(`Item [${i + 1}] is not an object`);
        }
        const t = item['type'];
        if (t === 'translation') {
            if (typeof item['en'] !== 'string') { throw new Error(`Item [${i + 1}] type=translation missing "en"`); }
            if (typeof item['vi'] !== 'string') { throw new Error(`Item [${i + 1}] type=translation missing "vi"`); }
        } else if (t === 'pattern') {
            if (typeof item['regex'] !== 'string')       { throw new Error(`Item [${i + 1}] type=pattern missing "regex"`); }
            if (typeof item['en_template'] !== 'string') { throw new Error(`Item [${i + 1}] type=pattern missing "en_template"`); }
            if (typeof item['vi_template'] !== 'string') { throw new Error(`Item [${i + 1}] type=pattern missing "vi_template"`); }
            if (typeof item['en'] !== 'string') { throw new Error(`Item [${i + 1}] type=pattern missing "en"`); }
            if (typeof item['vi'] !== 'string') { throw new Error(`Item [${i + 1}] type=pattern missing "vi"`); }
        } else {
            throw new Error(`Item [${i + 1}] has unknown type: ${JSON.stringify(t)}`);
        }
    }
}

/**
 * Substitute {0}, {1}, ... placeholders in a template string.
 * Each capture group is first looked up in `dict`; if not found the raw group value is used.
 */
function applyTemplate(template: string, groups: string[], dict: Dictionary): string {
    return template.replace(/\{(\d+)\}/g, (_, idx) => {
        const param = groups[parseInt(idx, 10)] ?? '';
        return dict.lookup(param) ?? param;
    });
}

async function callLLM(
    values: string[],
    model: vscode.LanguageModelChat,
    token: vscode.CancellationToken,
    retryHint?: string,
): Promise<LLMResult[]> {
    const numbered = values.map((v, i) => `${i + 1}. ${v}`).join('\n');
    const retryBlock = retryHint
        ? `\nPREVIOUS ATTEMPT FAILED — fix the issue before answering:\n"${retryHint}"\n`
        : '';

    const messages = [
        vscode.LanguageModelChatMessage.User(
            `You are a translation assistant for a Japanese software/business Excel file.
Translate each numbered text to English (EN) and Vietnamese (VI).
${retryBlock}
For each item, detect whether it is a TEMPLATE PATTERN — a fixed Japanese phrase containing variable parts (numbers, codes, IDs, dates, names, etc.). If it is a pattern:
  - output {"type":"pattern","regex":"ECMAScript-regex-with-capture-groups","en_template":"English with {0},{1}...","vi_template":"Việt with {0},{1}...","en":"translation of this exact input","vi":"translation of this exact input"}
  - regex must match the entire text; use () around each variable part.
  - Templates use {0},{1},... for capture groups in the same order.
If it is NOT a pattern (unique phrase), output {"type":"translation","en":"...","vi":"..."}

RULES:
1. Return ONLY a valid JSON array — no markdown, no explanation.
2. Exactly one object per input line, in the same order.
3. Represent newlines inside values as \\n (two characters). NEVER use actual newlines inside JSON strings.

Example — pattern:
  Input:  注文番号: 12345
  Output: {"type":"pattern","regex":"^注文番号: (.+)$","en_template":"Order No.: {0}","vi_template":"Số đơn hàng: {0}","en":"Order No.: 12345","vi":"Số đơn hàng: 12345"}

Example — plain translation:
  Input:  承認済み
  Output: {"type":"translation","en":"Approved","vi":"Đã phê duyệt"}

Output format:
[{...},...]

Input (${values.length} items):
${numbered}`,
        ),
    ];

    const t0 = Date.now();
    const response = await model.sendRequest(messages, {}, token);
    let raw = '';
    for await (const chunk of response.text) { raw += chunk; }
    log.info('Translate', 'LLM raw', { chars: raw.length, elapsed: elapsed(t0) });

    const jsonMatch = raw.match(/\[[\s\S]*\]/);
    if (!jsonMatch) {
        throw new Error(`No JSON array in response — preview: ${raw.slice(0, 200)}`);
    }
    let parsed: unknown;
    try {
        parsed = JSON.parse(jsonMatch[0]);
    } catch (e) {
        throw new Error(`JSON.parse failed: ${e} — preview: ${jsonMatch[0].slice(0, 200)}`);
    }
    validateLLMResults(parsed, values.length);
    return parsed;
}

// ── Retry wrapper ─────────────────────────────────────────────────────────────

async function translateBatch(
    values: string[],
    model: vscode.LanguageModelChat,
    token: vscode.CancellationToken,
): Promise<LLMResult[]> {
    const maxRetries = getMaxRetries();
    let lastError = '';

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const result = await callLLM(values, model, token, attempt > 1 ? lastError : undefined);
            if (attempt > 1) {
                log.info('Translate', 'LLM retry succeeded', { attempt });
            }
            return result;
        } catch (err) {
            lastError = err instanceof Error ? err.message : String(err);
            log.warn('Translate', 'LLM attempt failed', {
                attempt: `${attempt}/${maxRetries}`,
                reason: lastError.split('\n')[0],
            });
            if (attempt === maxRetries) {
                throw new Error(`LLM failed after ${maxRetries} attempts: ${lastError}`);
            }
        }
    }
    throw new Error('unreachable');
}

async function handleTranslate(
    request: vscode.ChatRequest,
    stream: vscode.ChatResponseStream,
    token: vscode.CancellationToken,
): Promise<void> {
    const t0 = Date.now();
    log.section('Translate');
    log.info('Translate', 'Start', { prompt: request.prompt || '(none)' });
    log.show();

    stream.markdown('### Copatis — Translate\n\n');

    // ── 1. Resolve file ──────────────────────────────────────────────────────
    let txtPath: string | undefined;
    const specifiedTxt = parseFilenameArg(request.prompt, 'txt');

    if (specifiedTxt) {
        const ws = workspacePath();
        const resolved = path.isAbsolute(specifiedTxt)
            ? specifiedTxt
            : ws ? path.join(ws, specifiedTxt) : specifiedTxt;
        if (fs.existsSync(resolved)) {
            txtPath = resolved;
            log.info('Translate', 'File resolved', { from: 'prompt', path: resolved });
        } else {
            log.warn('Translate', 'File not found', { specified: specifiedTxt });
            stream.markdown(`❌ Không tìm thấy file \`${specifiedTxt}\`\n\n`);
        }
    }

    if (!txtPath) {
        log.info('Translate', 'Opening file picker');
        const picked = await vscode.window.showOpenDialog({
            filters: { 'XlBridge Text Files': ['txt'] },
            canSelectMany: false,
            title: 'Chọn file TXT cần dịch',
        });
        if (!picked?.length) {
            log.warn('Translate', 'No file selected');
            stream.markdown('❌ Chưa chọn file. Hãy thử:\n```\n@copatis /translate ten-file.txt\n```');
            return;
        }
        txtPath = picked[0].fsPath;
        log.info('Translate', 'File resolved', { from: 'picker', path: txtPath });
    }

    // ── 2. Read & parse ──────────────────────────────────────────────────────
    let content: string;
    try {
        content = fs.readFileSync(txtPath, 'utf-8');
    } catch (err) {
        log.error('Translate', 'Cannot read file', { error: String(err) });
        stream.markdown(`❌ Không đọc được file: ${err}`);
        return;
    }

    const lines = content.replace(/\r\n/g, '\n').split('\n');
    const dataLines: Array<{ index: number; prefix: string; value: string }> = [];
    lines.forEach((line, i) => {
        const p = parseDataLine(line.trimEnd());
        if (p) dataLines.push({ index: i, ...p });
    });

    if (dataLines.length === 0) {
        log.warn('Translate', 'No data lines found');
        stream.markdown('⚠️ Không tìm thấy dòng dữ liệu `[Sheet]!Cell|Value` nào trong file.');
        return;
    }

    const relPath      = vscode.workspace.asRelativePath(txtPath);
    const batchSize    = getBatchSize();
    const commentLines = lines.filter(l => l.trimStart().startsWith('#')).length;
    const blankLines   = lines.filter(l => l.trim() === '').length;

    log.info('Translate', 'File parsed', {
        totalLines:   lines.length,
        dataLines:    dataLines.length,
        commentLines,
        blankLines,
        batchSize,
        file:         relPath,
    });

    // ── 3. Load dictionaries & pattern library ───────────────────────────────
    const dictDir    = getDictDir();
    const dictEN     = new Dictionary(path.join(dictDir, 'dict_ja_en.json'));
    const dictVI     = new Dictionary(path.join(dictDir, 'dict_ja_vi.json'));
    const patternLib = new PatternLibrary(path.join(dictDir, 'patterns_ja.json'));

    log.info('Translate', 'Resources loaded', {
        dir:         vscode.workspace.asRelativePath(dictDir) || dictDir,
        enSize:      dictEN.size,
        viSize:      dictVI.size,
        patterns:    patternLib.size,
    });

    // ── 4. Pick LLM model ────────────────────────────────────────────────────
    let models = await vscode.lm.selectChatModels({ vendor: 'copilot', family: 'gpt-4o' });
    if (models.length === 0) { models = await vscode.lm.selectChatModels({ vendor: 'copilot' }); }
    if (models.length === 0) {
        log.error('Translate', 'No LLM model available');
        stream.markdown('❌ Không tìm thấy Copilot LLM model. Hãy đảm bảo GitHub Copilot đang hoạt động.');
        return;
    }
    const model = models[0];
    log.info('Translate', 'Model', { name: model.name, family: model.family ?? 'unknown' });

    // ── 5. Summary header ────────────────────────────────────────────────────
    stream.markdown(`📄 File      : \`${relPath}\`\n`);
    stream.markdown(`🔢 Data lines: **${dataLines.length}**\n`);
    stream.markdown(`📚 Dict EN   : **${dictEN.size}** entries  |  Dict VI: **${dictVI.size}** entries\n`);
    stream.markdown(`🔖 Patterns  : **${patternLib.size}** templates\n`);
    stream.markdown(`📦 Batch size: **${batchSize}** (configurable via \`copatis.batchSize\`)\n`);
    stream.markdown(`🤖 Model     : \`${model.name}\`\n\n`);

    // ── 6. Main loop: passthrough → dict → pattern → LLM ────────────────────
    const outputLines = lines.map(l => l.trimEnd());
    let fromDict        = 0;
    let fromPassthrough = 0;   // non-Japanese: keep original as EN+VI
    let fromPattern     = 0;   // matched via PatternLibrary template
    let fromLLM         = 0;
    let failed          = 0;
    let llmBatchNum     = 0;

    // Progress tracking
    let processedCount  = 0;
    const checkpointEvery = Math.max(50, Math.floor(dataLines.length / 20)); // ~20 checkpoints

    log.info('Translate', 'Loop start', {
        total:          dataLines.length,
        batchSize,
        checkpointEvery,
        model:          model.name,
    });

    // value → all output-line positions waiting for this translation (deduplication map)
    // Each unique value is sent to LLM exactly once; duplicates reuse the same result.
    const pendingMap = new Map<string, Array<{ index: number; prefix: string }>>();

    /** Flush unique pending values to LLM; apply results to all waiting output lines. */
    const flushPending = async (): Promise<void> => {
        if (pendingMap.size === 0) return;

        const batch = [...pendingMap.entries()]; // [[value, [{index,prefix},...]], ...]
        pendingMap.clear();
        llmBatchNum++;
        const batchT0 = Date.now();
        const totalLines = batch.reduce((s, [, w]) => s + w.length, 0);

        const valPreview = batch.slice(0, 3)
            .map(([v]) => (v.length > 30 ? v.slice(0, 28) + '…' : v))
            .join(' │ ');
        log.info('Translate', 'LLM batch start', {
            batch:   llmBatchNum,
            unique:  batch.length,
            lines:   totalLines,
            preview: batch.length > 3 ? `${valPreview}  … +${batch.length - 3}` : valPreview,
        });
        stream.progress(`LLM batch ${llmBatchNum}: ${batch.length} unique (${totalLines} lines)...`);

        try {
            const results = await translateBatch(batch.map(([v]) => v), model, token);

            let batchPatterns = 0;
            batch.forEach(([value, waiters], j) => {
                const r = results[j];
                const en = escNl(r.en);
                const vi = escNl(r.vi);

                // Apply translation to every output line that shares this value
                for (const w of waiters) {
                    outputLines[w.index] = `${w.prefix}|${value}|${en}|${vi}`;
                }

                if (r.type === 'pattern') {
                    patternLib.add({ regex: r.regex, en_template: r.en_template, vi_template: r.vi_template });
                    batchPatterns++;
                }
                // Cache for future lines (cross-batch dedup via dict, cross-session persistence)
                dictEN.set(value, en);
                dictVI.set(value, vi);
            });

            fromLLM += totalLines;
            log.info('Translate', 'LLM batch done', {
                batch:       llmBatchNum,
                unique:      batch.length,
                lines:       totalLines,
                newPatterns: batchPatterns,
                elapsed:     elapsed(batchT0),
                dictENnew:   dictEN.newCount,
                dictVInew:   dictVI.newCount,
            });
            const dedup = totalLines - batch.length;
            stream.markdown(`✔ LLM batch ${llmBatchNum}: ${batch.length} unique`
                + (dedup > 0 ? ` (+${dedup} dedup)` : '')
                + (batchPatterns > 0 ? `, ${batchPatterns} patterns learned` : '')
                + ` — ${elapsed(batchT0)}\n`);

        } catch (err) {
            const failedLines = batch.reduce((s, [, w]) => s + w.length, 0);
            failed += failedLines;
            const msg = err instanceof Error ? err.message : String(err);
            log.warn('Translate', 'LLM batch failed', {
                batch:   llmBatchNum,
                failed:  failedLines,
                elapsed: elapsed(batchT0),
                reason:  msg.split('\n')[0],
            });
            stream.markdown(`⚠️ LLM batch ${llmBatchNum} thất bại (${batch.length} entries): ${msg.split('\n')[0]}\n`);
        }
    };

    for (const d of dataLines) {
        if (token.isCancellationRequested) {
            log.warn('Translate', 'Cancelled by user', {
                processed: `${processedCount}/${dataLines.length}`,
                fromDict, fromPattern, fromLLM, failed,
                elapsed: elapsed(t0),
            });
            await flushPending();
            stream.markdown('\n⚠️ Đã hủy bởi người dùng.\n');
            break;
        }

        if (!needsTranslation(d.value)) {
            // ① Non-Japanese / symbol-only — passthrough as-is
            outputLines[d.index] = `${d.prefix}|${d.value}|${d.value}|${d.value}`;
            fromPassthrough++;

        } else {
            const cachedEN = dictEN.lookup(d.value);
            const cachedVI = dictVI.lookup(d.value);

            if (cachedEN !== undefined && cachedVI !== undefined) {
                // ② Dict cache hit (cross-session or already translated earlier in this run)
                outputLines[d.index] = `${d.prefix}|${d.value}|${cachedEN}|${cachedVI}`;
                fromDict++;

            } else {
                const pm = patternLib.match(d.value);

                if (pm) {
                    // ③ Pattern library match — substitute template params
                    const en = escNl(applyTemplate(pm.entry.en_template, pm.groups, dictEN));
                    const vi = escNl(applyTemplate(pm.entry.vi_template, pm.groups, dictVI));
                    outputLines[d.index] = `${d.prefix}|${d.value}|${en}|${vi}`;
                    dictEN.set(d.value, en);
                    dictVI.set(d.value, vi);
                    fromPattern++;

                } else {
                    // ④ Unknown — queue for LLM (deduplicated: same value → shared result)
                    if (pendingMap.has(d.value)) {
                        pendingMap.get(d.value)!.push({ index: d.index, prefix: d.prefix });
                    } else {
                        pendingMap.set(d.value, [{ index: d.index, prefix: d.prefix }]);
                    }
                    if (pendingMap.size >= batchSize) {
                        await flushPending();
                    }
                }
            }
        }

        // ── Progress checkpoint ───────────────────────────────────────────────
        processedCount++;
        if (processedCount % checkpointEvery === 0 || processedCount === dataLines.length) {
            const pct = Math.round(processedCount / dataLines.length * 100);
            const queued = [...pendingMap.values()].reduce((s, w) => s + w.length, 0);
            log.info('Translate', 'Progress', {
                processed:   `${processedCount}/${dataLines.length}`,
                pct:         `${pct}%`,
                passthrough: fromPassthrough,
                dict:        fromDict,
                pattern:     fromPattern,
                llm:         fromLLM,
                queued,
                pending:     pendingMap.size,
                elapsed:     elapsed(t0),
            });
            stream.progress(
                `[${pct}%] ${processedCount}/${dataLines.length} dòng`
                + ` — ⚡${fromPassthrough} giữ nguyên`
                + ` | 📚${fromDict} từ điển`
                + (fromPattern > 0 ? ` | 🔖${fromPattern} pattern` : '')
                + ` | 🤖${fromLLM} LLM`
                + (queued > 0 ? ` | ⏳${queued} đang chờ` : ''),
            );
        }
    }

    // Flush any remaining entries
    await flushPending();

    log.info('Translate', 'Loop done', {
        total:       dataLines.length,
        passthrough: fromPassthrough,
        dict:        fromDict,
        pattern:     fromPattern,
        llm:         fromLLM,
        failed,
        llmBatches:  llmBatchNum,
        elapsed:     elapsed(t0),
    });

    // ── 7. Save dictionaries & pattern library ───────────────────────────────
    dictEN.save();
    dictVI.save();
    patternLib.save();

    // ── 8. Write output file ─────────────────────────────────────────────────
    const outputPath = txtPath.replace(/\.txt$/i, '_translated.txt');
    try {
        fs.writeFileSync(outputPath, outputLines.join('\n'), 'utf-8');

        log.info('Translate', 'Done', {
            fromPassthrough, fromDict, fromPattern, fromLLM, failed,
            llmBatches:      llmBatchNum,
            dictENadded:     dictEN.newCount,
            dictVIadded:     dictVI.newCount,
            patternsAdded:   patternLib.newCount,
            patternTotal:    patternLib.size,
            elapsed:         elapsed(t0),
            output:          vscode.workspace.asRelativePath(outputPath),
        });

        stream.markdown('\n---\n✅ **Hoàn thành!**\n\n');
        stream.markdown(`| Nguồn | Số dòng |\n|---|---|\n`);
        stream.markdown(`| Giữ nguyên (EN/số/ký hiệu) | **${fromPassthrough}** |\n`);
        stream.markdown(`| Từ điển (cache hit)         | **${fromDict}** |\n`);
        stream.markdown(`| Patterns (template match)   | **${fromPattern}** |\n`);
        stream.markdown(`| LLM (${llmBatchNum} batches)              | **${fromLLM}** |\n`);
        if (failed > 0) {
            stream.markdown(`| Lỗi                         | **${failed}** |\n`);
        }
        stream.markdown(`| Từ điển EN cập nhật         | +**${dictEN.newCount}** entries |\n`);
        stream.markdown(`| Từ điển VI cập nhật         | +**${dictVI.newCount}** entries |\n`);
        stream.markdown(`| Patterns mới học được       | +**${patternLib.newCount}** (tổng: ${patternLib.size}) |\n`);
        stream.markdown(`\n📤 Output: \`${vscode.workspace.asRelativePath(outputPath)}\`\n`);

        stream.button({
            command: 'vscode.open',
            arguments: [vscode.Uri.file(outputPath)],
            title: '📂 Mở file đã dịch',
        });
        stream.button({
            command: 'revealInExplorer',
            arguments: [vscode.Uri.file(dictDir)],
            title: '📚 Mở thư mục từ điển',
        });
    } catch (err) {
        log.error('Translate', 'Write failed', { error: String(err) });
        stream.markdown(`❌ Không ghi được file output: ${err}`);
    }
}

// ─── /help ────────────────────────────────────────────────────────────────────

async function handleHelp(stream: vscode.ChatResponseStream): Promise<void> {
    stream.markdown(`# Copatis Assistant

**Copatis** giúp extract nội dung file Excel ra TXT, dịch sang EN+VI bằng LLM, rồi inject bản dịch trở lại.

---

## Lệnh Chat

| Lệnh | Mô tả |
|------|-------|
| \`@copatis /extract\` | Extract tất cả sheet từ file .xlsx |
| \`@copatis /extract file.xlsx\` | Extract file cụ thể |
| \`@copatis /extract file.xlsx --sheet Sheet1\` | Chỉ extract sheet chỉ định |
| \`@copatis /inject en\` | Inject bản dịch Tiếng Anh (tự tìm file) |
| \`@copatis /inject vi\` | Inject bản dịch Tiếng Việt (tự tìm file) |
| \`@copatis /inject en file.xlsx translated.txt\` | Inject Tiếng Anh với file chỉ định |
| \`@copatis /inject vi file.xlsx translated.txt\` | Inject Tiếng Việt với file chỉ định |
| \`@copatis /translate\` | Dịch file TXT sang EN + VI (mở file picker) |
| \`@copatis /translate file.txt\` | Dịch file TXT chỉ định |
| \`@copatis /help\` | Hiển thị trợ giúp này |

---

## Format file TXT (sau extract)

\`\`\`
# XlBridge Export
# Source: filename.xlsx
# Date: 2024-01-01

[SheetName]!A1|Nội dung ô A1
[SheetName]!B2|Nội dung ô B2
\`\`\`

## Format sau khi /translate

\`\`\`
[SheetName]!A1|原文|English translation|Bản dịch tiếng Việt
[SheetName]!B2|原文 2|English translation 2|Bản dịch tiếng Việt 2
\`\`\`

---

## Log Output

Mở **View → Output → Copatis** để xem log chi tiết của mọi thao tác.

---

## CLI trực tiếp (terminal)

\`\`\`bash
xlbridge extract --input file.xlsx --output export.txt
xlbridge extract --input file.xlsx --output export.txt --sheet Sheet1
xlbridge inject  --input file.xlsx --translation translated.txt --output result.xlsx
\`\`\`
`);
}

// ─── General ──────────────────────────────────────────────────────────────────

async function handleGeneral(
    request: vscode.ChatRequest,
    stream: vscode.ChatResponseStream,
    token: vscode.CancellationToken,
): Promise<void> {
    const prompt = request.prompt.toLowerCase();
    log.info('General', 'Routing', { prompt: request.prompt });

    if (/extract|trích|xuất/.test(prompt))      { return handleExtract(request, stream, token); }
    if (/inject|nhập|translation/.test(prompt))  { return handleInject(request, stream, token); }
    if (/translat|dịch/.test(prompt))            { return handleTranslate(request, stream, token); }

    try {
        const models = await vscode.lm.selectChatModels({ vendor: 'copilot', family: 'gpt-4o' });
        if (models.length === 0) { return handleHelp(stream); }

        const messages = [
            vscode.LanguageModelChatMessage.User(
                `Bạn là trợ lý cho Copatis — công cụ Excel translation workflow:
1. /extract  — Extract cell content từ .xlsx → .txt
2. /translate — Dịch .txt sang EN + VI bằng LLM
3. /inject   — Inject bản dịch từ .txt → .xlsx

Người dùng hỏi: "${request.prompt}"

Trả lời ngắn gọn bằng tiếng Việt. Hướng dẫn dùng lệnh phù hợp nếu cần.`,
            ),
        ];

        const response = await models[0].sendRequest(messages, {}, token);
        for await (const chunk of response.text) { stream.markdown(chunk); }
    } catch {
        await handleHelp(stream);
    }
}

// ─── Activation ───────────────────────────────────────────────────────────────

export function activate(context: vscode.ExtensionContext): void {
    log = new Logger();
    log.info('System', 'Extension activated', { version: '0.1.0' });

    const participant = vscode.chat.createChatParticipant(
        PARTICIPANT_ID,
        async (request, _ctx, stream, token) => {
            try {
                log.info('System', 'Command received', {
                    command: request.command ?? '(default)',
                    prompt:  request.prompt || '(empty)',
                });
                switch (request.command) {
                    case 'extract':   await handleExtract(request, stream, token);   break;
                    case 'inject':    await handleInject(request, stream, token);    break;
                    case 'translate': await handleTranslate(request, stream, token); break;
                    case 'help':      await handleHelp(stream);                      break;
                    default:          await handleGeneral(request, stream, token);   break;
                }
            } catch (err: unknown) {
                const msg = err instanceof Error ? err.message : String(err);
                log.error('System', 'Unhandled exception', { error: msg });
                stream.markdown(`❌ Lỗi không mong đợi: ${msg}`);
            }
            return {};
        },
    );

    participant.iconPath = new vscode.ThemeIcon('table');

    participant.followupProvider = {
        provideFollowups(_result, _ctx, _token) {
            return [
                { label: '/extract   — Extract file Excel',    prompt: '/extract',   participant: PARTICIPANT_ID },
                { label: '/inject    — Inject bản dịch',       prompt: '/inject',    participant: PARTICIPANT_ID },
                { label: '/translate — Dịch file TXT EN + VI', prompt: '/translate', participant: PARTICIPANT_ID },
                { label: '/help      — Xem hướng dẫn',         prompt: '/help',      participant: PARTICIPANT_ID },
            ];
        },
    };

    context.subscriptions.push(
        participant,
        log as unknown as vscode.Disposable,

        vscode.commands.registerCommand('copatis.extract', async () => {
            const files = await vscode.window.showOpenDialog({
                filters: { 'Excel Files': ['xlsx'] },
                canSelectMany: false,
                title: 'Chọn file Excel cần extract',
            });
            if (!files?.length) return;

            const inputFile  = files[0].fsPath;
            const outputFile = inputFile.replace(/\.xlsx?$/i, '_export.txt');
            log.section('Extract (Command Palette)');
            log.info('Extract', 'Start', { input: inputFile });
            log.show();

            const terminal = vscode.window.createTerminal('Copatis');
            terminal.show();
            terminal.sendText(`${xlbridgeCmd()} extract --input "${inputFile}" --output "${outputFile}"`);
        }),

        vscode.commands.registerCommand('copatis.inject', async () => {
            const xlsxFiles = await vscode.window.showOpenDialog({
                filters: { 'Excel Files': ['xlsx'] },
                canSelectMany: false,
                title: 'Chọn file Excel gốc',
            });
            if (!xlsxFiles?.length) return;

            const txtFiles = await vscode.window.showOpenDialog({
                filters: { 'Text Files': ['txt'] },
                canSelectMany: false,
                title: 'Chọn file bản dịch (.txt)',
            });
            if (!txtFiles?.length) return;

            log.section('Inject (Command Palette)');
            log.info('Inject', 'Start', {
                xlsx: xlsxFiles[0].fsPath,
                txt:  txtFiles[0].fsPath,
            });
            log.show();

            const terminal = vscode.window.createTerminal('Copatis');
            terminal.show();
            terminal.sendText(
                `${xlbridgeCmd()} inject --input "${xlsxFiles[0].fsPath}" --translation "${txtFiles[0].fsPath}"`,
            );
        }),
    );
}

export function deactivate(): void {
    log?.info('System', 'Extension deactivated');
}
