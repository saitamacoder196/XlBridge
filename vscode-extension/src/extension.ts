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

const TRANSLATE_BATCH_SIZE = 25;

interface TranslationPair { en: string; vi: string; }

function parseDataLine(line: string): { prefix: string; value: string } | undefined {
    const match = line.match(/^(\[[^\]]+\]![A-Za-z]+\d+)\|(.+)$/);
    if (!match) return undefined;
    return { prefix: match[1], value: match[2] };
}

async function translateBatch(
    values: string[],
    model: vscode.LanguageModelChat,
    token: vscode.CancellationToken,
): Promise<TranslationPair[]> {
    const numbered = values.map((v, i) => `${i + 1}. ${v}`).join('\n');

    const messages = [
        vscode.LanguageModelChatMessage.User(
            `You are a translation assistant for a Japanese software/business Excel file.
Translate each numbered text to English (EN) and Vietnamese (VI).
Return ONLY a valid JSON array — no markdown fences, no explanation:
[{"en":"English text","vi":"Tiếng Việt"},...]

Input texts:
${numbered}`,
        ),
    ];

    const t0 = Date.now();
    const response = await model.sendRequest(messages, {}, token);
    let raw = '';
    for await (const chunk of response.text) { raw += chunk; }

    log.info('Translate', 'LLM response', { chars: raw.length, elapsed: elapsed(t0) });

    const jsonMatch = raw.match(/\[[\s\S]*\]/);
    if (!jsonMatch) {
        throw new Error(`LLM không trả về JSON hợp lệ:\n${raw.slice(0, 300)}`);
    }

    const parsed = JSON.parse(jsonMatch[0]) as TranslationPair[];
    if (!Array.isArray(parsed) || parsed.length !== values.length) {
        throw new Error(`Kỳ vọng ${values.length} bản dịch, nhận được ${parsed?.length ?? 0}`);
    }
    return parsed;
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

    // ── 1. Resolve file ──
    let txtPath: string | undefined;
    const specifiedTxt = parseFilenameArg(request.prompt, 'txt');

    if (specifiedTxt) {
        const ws = workspacePath();
        const resolved = path.isAbsolute(specifiedTxt)
            ? specifiedTxt
            : ws ? path.join(ws, specifiedTxt) : specifiedTxt;
        if (fs.existsSync(resolved)) {
            txtPath = resolved;
            log.info('Translate', 'Resolved', { from: 'prompt', path: resolved });
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
            log.warn('Translate', 'No file selected by user');
            stream.markdown('❌ Chưa chọn file. Hãy thử:\n```\n@copatis /translate ten-file.txt\n```');
            return;
        }
        txtPath = picked[0].fsPath;
        log.info('Translate', 'Resolved', { from: 'picker', path: txtPath });
    }

    // ── 2. Read & parse ──
    let content: string;
    try {
        content = fs.readFileSync(txtPath, 'utf-8');
    } catch (err) {
        log.error('Translate', 'Cannot read file', { path: txtPath, error: String(err) });
        stream.markdown(`❌ Không đọc được file: ${err}`);
        return;
    }

    const lines = content.replace(/\r\n/g, '\n').split('\n');
    const dataLines: Array<{ index: number; prefix: string; value: string }> = [];

    lines.forEach((line, i) => {
        const parsed = parseDataLine(line.trimEnd());
        if (parsed) dataLines.push({ index: i, ...parsed });
    });

    if (dataLines.length === 0) {
        log.warn('Translate', 'No data lines found in file', { path: txtPath });
        stream.markdown('⚠️ Không tìm thấy dòng dữ liệu `[Sheet]!Cell|Value` nào trong file.');
        return;
    }

    const totalBatches = Math.ceil(dataLines.length / TRANSLATE_BATCH_SIZE);
    const relPath = vscode.workspace.asRelativePath(txtPath);

    log.info('Translate', 'Parsed', {
        file:      relPath,
        totalLines: lines.length,
        dataLines:  dataLines.length,
        batches:    totalBatches,
        batchSize:  TRANSLATE_BATCH_SIZE,
    });

    stream.markdown(`📄 File  : \`${relPath}\`\n`);
    stream.markdown(`🔢 Dòng  : **${dataLines.length}** dòng cần dịch\n`);
    stream.markdown(`📦 Batch : ${TRANSLATE_BATCH_SIZE} dòng/lần × ${totalBatches} lần gọi LLM\n\n`);

    // ── 3. Pick model ──
    let models = await vscode.lm.selectChatModels({ vendor: 'copilot', family: 'gpt-4o' });
    if (models.length === 0) {
        models = await vscode.lm.selectChatModels({ vendor: 'copilot' });
    }
    if (models.length === 0) {
        log.error('Translate', 'No LLM model available');
        stream.markdown('❌ Không tìm thấy Copilot LLM model. Hãy đảm bảo GitHub Copilot đang hoạt động.');
        return;
    }
    const model = models[0];
    log.info('Translate', 'Model selected', { name: model.name, family: model.family ?? 'unknown' });
    stream.markdown(`🤖 Model : \`${model.name}\`\n\n`);

    // ── 4. Batch translate ──
    const outputLines = lines.map(l => l.trimEnd());
    let processed = 0;
    let failed = 0;

    for (let i = 0; i < dataLines.length; i += TRANSLATE_BATCH_SIZE) {
        if (token.isCancellationRequested) {
            log.warn('Translate', 'Cancelled by user', { processedSoFar: processed });
            stream.markdown('\n⚠️ Đã hủy bởi người dùng.\n');
            break;
        }

        const batch = dataLines.slice(i, i + TRANSLATE_BATCH_SIZE);
        const batchNum = Math.floor(i / TRANSLATE_BATCH_SIZE) + 1;
        const batchT0  = Date.now();

        stream.progress(`Batch ${batchNum}/${totalBatches} — đang dịch ${batch.length} dòng...`);

        try {
            const translations = await translateBatch(batch.map(d => d.value), model, token);

            batch.forEach((d, j) => {
                const t = translations[j];
                outputLines[d.index] = `${d.prefix}|${d.value}|${t.en}|${t.vi}`;
            });

            processed += batch.length;

            log.info('Translate', 'Batch OK', {
                batch:    `${batchNum}/${totalBatches}`,
                sent:     batch.length,
                ok:       batch.length,
                elapsed:  elapsed(batchT0),
                progress: `${processed}/${dataLines.length}`,
            });
            stream.markdown(`✔ Batch ${batchNum}/${totalBatches} — ${processed}/${dataLines.length} dòng\n`);

        } catch (err) {
            failed += batch.length;
            const msg = err instanceof Error ? err.message : String(err);
            log.warn('Translate', 'Batch FAILED', {
                batch:   `${batchNum}/${totalBatches}`,
                sent:    batch.length,
                failed:  batch.length,
                elapsed: elapsed(batchT0),
                reason:  msg.split('\n')[0],
            });
            stream.markdown(`⚠️ Batch ${batchNum} thất bại: ${msg.split('\n')[0]}\n`);
        }
    }

    // ── 5. Write output ──
    const outputPath = txtPath.replace(/\.txt$/i, '_translated.txt');
    try {
        fs.writeFileSync(outputPath, outputLines.join('\n'), 'utf-8');

        log.info('Translate', 'Written', {
            path:       vscode.workspace.asRelativePath(outputPath),
            translated: processed,
            failed,
            elapsed:    elapsed(t0),
        });

        stream.markdown(`\n---\n✅ **Hoàn thành!**\n\n`);
        stream.markdown(`- Dịch thành công : **${processed}** dòng\n`);
        if (failed > 0) stream.markdown(`- Lỗi             : **${failed}** dòng\n`);
        stream.markdown(`\n📤 Output: \`${vscode.workspace.asRelativePath(outputPath)}\`\n`);

        stream.button({
            command: 'vscode.open',
            arguments: [vscode.Uri.file(outputPath)],
            title: '📂 Mở file đã dịch',
        });
    } catch (err) {
        log.error('Translate', 'Write failed', { path: outputPath, error: String(err) });
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
