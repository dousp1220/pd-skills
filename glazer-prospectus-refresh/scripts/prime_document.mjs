#!/usr/bin/env node
import crypto from 'node:crypto';
import fs from 'node:fs';
import { open, rename, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const DEFAULT_PAODINGAI_API_BASE_URL = 'https://platform.paodingai.com/';
const DEFAULT_POLL_INTERVAL_MS = 20_000;
const DEFAULT_HEARTBEAT_INTERVAL_MS = 30_000;
const RETRY_DELAY_MS = 2_000;
const STATUS_REQUEST_TIMEOUT_MS = 30_000;
const SUBMIT_REQUEST_TIMEOUT_MS = 10 * 60 * 1000;
const DOWNLOAD_REQUEST_TIMEOUT_MS = 10 * 60 * 1000;
const STATUS_REQUEST_ATTEMPTS = 3;
const MAX_FILE_SIZE = 200 * 1024 * 1024;
const ZIP_SIGNATURES = new Set(['504b0304', '504b0506', '504b0708']);
const COMPOUND_DOC_SIGNATURE = Buffer.from('d0cf11e0a1b11ae1', 'hex');
const PDROUTER_SERVICE_CODE = 'glazer';
const SERVICE_PATH = '/prime/projects';
const PROSPECTUS_EXTENSIONS = new Set(['.docx']);
const AUDIT_EXTENSIONS = new Set(['.pdf', '.doc', '.docx', '.xls', '.xlsx']);
const DOWNLOADABLE_STATUSES = new Set(['succeeded', 'export_failed']);
const PERIOD_RANGE = /^(\d{4})年[-－](\d{4})年(?:(\d{1,2})月(?:末)?)?$/;
const PERIOD_SEGMENT = /^(\d{4})年(?:(\d{1,2})月(?:末)?|(\d{1,2})-(\d{1,2})月(?:末)?)?$/;
const PERIOD_SPLIT = /[、，,及与和]/;
const QUARTER_MONTH = [
  ['第一季度', '3月'],
  ['第二季度', '6月'],
  ['第三季度', '9月'],
  ['第四季度', '12月'],
  ['一季度', '3月'],
  ['二季度', '6月'],
  ['三季度', '9月'],
  ['四季度', '12月'],
  ['1季度', '3月'],
  ['2季度', '6月'],
  ['3季度', '9月'],
  ['4季度', '12月'],
  ['Q1', '3月'],
  ['Q2', '6月'],
  ['Q3', '9月'],
  ['Q4', '12月'],
  ['q1', '3月'],
  ['q2', '6月'],
  ['q3', '9月'],
  ['q4', '12月'],
];

function normalizePeriodText(text) {
  return String(text ?? '').trim().replace(/[\s\u3000]+/g, '').replace(/^(截至|截止到|截止)/, '');
}

function normalizePeriodSegment(segment) {
  let text = segment;
  for (const [token, month] of QUARTER_MONTH) text = text.replaceAll(token, month);
  text = text.replaceAll('上半年', '6月').replaceAll('下半年', '12月').replaceAll('半年度', '6月').replaceAll('半年', '6月');
  text = text.replaceAll('年报', '年').replaceAll('年度', '年').replace(/年+/g, '年');
  text = text.replace(/(\d{1,2})月[-－至到](\d{1,2})月/g, '$1-$2月');
  text = text.replace(/(\d{4})年[-－至到](\d{4})年/g, '$1年-$2年');
  text = text.replace(/^(\d{4})[-－至到](\d{4})年/, '$1年-$2年');
  text = text.replace(/(\d{4})年[-－至到](\d{4})$/, '$1年-$2年');
  text = text.replace(/(\d{4})年(\d{1,2})月\d{1,2}日/g, '$1年$2月');
  if (/^\d{4}$/.test(text)) text = `${text}年`;
  return text;
}

class PrimeError extends Error {
  constructor(errorCode, message, options = {}) {
    super(message, options.cause ? { cause: options.cause } : undefined);
    this.name = 'PrimeError';
    this.errorCode = errorCode;
    this.projectId = options.projectId;
    this.httpStatus = options.httpStatus;
  }
}

function usage() {
  return `Usage:
  prime_document.mjs start <prospectus.docx> <audit.pdf> [audit2.xlsx ...] --period <历史募集报告期> --period <审计报告期> --new-period <新募集报告期> [--file-type prospectus] [--file-type audit] [--output-dir <dir>] [--json]
  prime_document.mjs watch <project_id> [--after-event-id <id>] [--json]
  prime_document.mjs status <project_id> [--json]
  prime_document.mjs download <project_id> [--output-dir <dir>] [--json]`;
}

function requireArgValue(argv, index, label) {
  const value = argv[index + 1];
  if (!value || value.startsWith('--')) {
    throw new PrimeError('prime.invalid_arguments', `${label} requires a value`);
  }
  return value;
}

function parseJsonText(value, label) {
  try {
    return JSON.parse(value);
  } catch (error) {
    throw new PrimeError('prime.invalid_arguments', `${label} must be valid JSON`, { cause: error });
  }
}

function parseJsonFile(filePath, label) {
  let text;
  try {
    text = fs.readFileSync(path.resolve(filePath), 'utf8');
  } catch (error) {
    throw new PrimeError('prime.invalid_arguments', `${label} file does not exist: ${filePath}`, { cause: error });
  }
  return parseJsonText(text, label);
}

function periodsFromFlags(periodList, newPeriod) {
  const periods = { __new__: newPeriod };
  periodList.forEach((item, index) => {
    periods[String(index)] = item;
  });
  return periods;
}

function positiveInteger(value, label) {
  if (!/^\d+$/.test(String(value ?? ''))) {
    throw new PrimeError('prime.invalid_arguments', `${label} must be a positive integer`);
  }
  const parsed = Number.parseInt(value, 10);
  if (!Number.isSafeInteger(parsed) || parsed <= 0) {
    throw new PrimeError('prime.invalid_arguments', `${label} must be a positive integer`);
  }
  return parsed;
}

function nonNegativeInteger(value, label) {
  if (!/^\d+$/.test(String(value ?? ''))) {
    throw new PrimeError('prime.invalid_arguments', `${label} must be a non-negative integer`);
  }
  const parsed = Number.parseInt(value, 10);
  if (!Number.isSafeInteger(parsed) || parsed < 0) {
    throw new PrimeError('prime.invalid_arguments', `${label} must be a non-negative integer`);
  }
  return parsed;
}

function requiredEnvironment(name, environment) {
  const value = environment[name]?.trim();
  if (!value) throw new PrimeError('prime.configuration_error', `Missing environment variable ${name}`);
  return value;
}

function normalizeBaseUrl(value) {
  let url;
  try {
    url = new URL(value);
  } catch (error) {
    throw new PrimeError('prime.configuration_error', 'PAODINGAI_API_BASE_URL must be a full HTTP(S) URL', { cause: error });
  }
  if (!['http:', 'https:'].includes(url.protocol)) {
    throw new PrimeError('prime.configuration_error', 'PDRouter base URL must use HTTP or HTTPS');
  }
  if (url.search || url.hash) {
    throw new PrimeError('prime.configuration_error', 'PAODINGAI_API_BASE_URL must be a base URL without query or hash');
  }
  url.pathname = `${url.pathname.replace(/\/+$/, '')}/`;
  return url;
}

function serviceUrl(baseUrl, suffix = '') {
  const prefix = `${baseUrl.pathname.replace(/\/+$/, '')}/openapi/${encodeURIComponent(PDROUTER_SERVICE_CODE)}`;
  return new URL(`${prefix}${SERVICE_PATH}${suffix}`, baseUrl);
}

function normalizeFileTypes(value) {
  if (!Array.isArray(value)) {
    throw new PrimeError('prime.invalid_arguments', '--file-types must be a JSON array');
  }
  return value;
}

function parseArgs(argv) {
  if (argv.length === 0 || argv[0] === '--help' || argv[0] === '-h') {
    return { help: true };
  }

  const mode = argv[0];
  if (!['start', 'watch', 'status', 'download'].includes(mode)) {
    throw new PrimeError('prime.invalid_arguments', `Unsupported command: ${mode}`);
  }

  const args = {
    help: false,
    mode,
    files: [],
    projectId: null,
    periods: null,
    periodFlags: [],
    newPeriod: null,
    fileTypes: [],
    fileTypeFlags: [],
    afterEventId: 0,
    outputDir: null,
    json: false,
  };

  for (let index = 1; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '--period') {
      args.periodFlags.push(requireArgValue(argv, index, '--period'));
      index += 1;
    } else if (arg === '--new-period') {
      args.newPeriod = requireArgValue(argv, index, '--new-period');
      index += 1;
    } else if (arg === '--file-type') {
      args.fileTypeFlags.push(requireArgValue(argv, index, '--file-type'));
      index += 1;
    } else if (arg === '--periods-file') {
      args.periods = parseJsonFile(requireArgValue(argv, index, '--periods-file'), '--periods-file');
      index += 1;
    } else if (arg === '--file-types-file') {
      args.fileTypes = normalizeFileTypes(parseJsonFile(requireArgValue(argv, index, '--file-types-file'), '--file-types-file'));
      index += 1;
    } else if (arg === '--after-event-id') {
      args.afterEventId = nonNegativeInteger(argv[index + 1], '--after-event-id');
      index += 1;
    } else if (arg === '--output-dir') {
      const value = argv[index + 1];
      if (!value || value.startsWith('--')) throw new PrimeError('prime.invalid_arguments', '--output-dir requires a directory');
      args.outputDir = value;
      index += 1;
    } else if (arg === '--json') {
      args.json = true;
    } else if (arg.startsWith('--')) {
      throw new PrimeError('prime.invalid_arguments', `Unsupported option: ${arg}`);
    } else if (mode === 'start') {
      args.files.push(arg);
    } else if (!args.projectId) {
      args.projectId = positiveInteger(arg, 'project_id');
    } else {
      throw new PrimeError('prime.invalid_arguments', `Unexpected positional argument: ${arg}`);
    }
  }

  if (mode === 'start') {
    if (args.files.length < 2) throw new PrimeError('prime.invalid_arguments', 'start requires one prospectus and at least one audit/financial report');
    const hasPeriodFlags = args.periodFlags.length > 0 || args.newPeriod;
    if (hasPeriodFlags && args.periods) {
      throw new PrimeError('prime.invalid_arguments', 'do not mix --period/--new-period with --periods-file');
    }
    if (hasPeriodFlags) {
      if (args.periodFlags.length !== args.files.length) {
        throw new PrimeError('prime.invalid_arguments', '--period must be passed once per input file, in the same order');
      }
      if (!args.newPeriod) throw new PrimeError('prime.invalid_arguments', 'start requires --new-period');
      args.periods = periodsFromFlags(args.periodFlags, args.newPeriod);
    }
    if (!args.periods || typeof args.periods !== 'object' || Array.isArray(args.periods)) {
      throw new PrimeError('prime.invalid_arguments', 'start requires --period and --new-period');
    }
    if (args.fileTypeFlags.length > 0 && args.fileTypes.length > 0) {
      throw new PrimeError('prime.invalid_arguments', 'do not mix --file-type with --file-types-file');
    }
    if (args.fileTypeFlags.length > 0) {
      args.fileTypes = resolveFileTypeFlags(args.files, args.fileTypeFlags);
    }
  } else {
    if (!args.projectId) throw new PrimeError('prime.invalid_arguments', `${mode} requires project_id`);
    if (args.periods || args.periodFlags.length > 0 || args.newPeriod) {
      throw new PrimeError('prime.invalid_arguments', 'period options are only valid for start');
    }
    if (args.fileTypes.length > 0 || args.fileTypeFlags.length > 0) {
      throw new PrimeError('prime.invalid_arguments', 'file type options are only valid for start');
    }
    if (!['start', 'download'].includes(mode) && args.outputDir) {
      throw new PrimeError('prime.invalid_arguments', '--output-dir is only valid for start and download');
    }
  }
  return args;
}

function resolveFileTypeFlags(files, fileTypeFlags) {
  if (fileTypeFlags.length > files.length) {
    throw new PrimeError('prime.invalid_arguments', '--file-type count cannot exceed the number of input files');
  }
  const types = fileTypeFlags.map((value, index) => normalizeKind(value, `--file-type #${index}`));
  if (types.length === files.length) return types;
  if (types.length === 1 && types[0] === 'prospectus') {
    return assignProspectusAndDefaultAudit(files);
  }
  if (types.includes('prospectus')) {
    while (types.length < files.length) types.push('audit');
  }
  return types;
}

function assignProspectusAndDefaultAudit(files) {
  const docxIndexes = [];
  const namedIndexes = [];
  files.forEach((filePath, index) => {
    if (path.extname(filePath).toLowerCase() === '.docx') docxIndexes.push(index);
    if (guessKind(filePath) === 'prospectus') namedIndexes.push(index);
  });
  let prospectusIndex = null;
  if (namedIndexes.length === 1) prospectusIndex = namedIndexes[0];
  else if (docxIndexes.length === 1) prospectusIndex = docxIndexes[0];
  if (prospectusIndex === null) {
    throw new PrimeError(
      'prime.invalid_arguments',
      'Cannot tell which file is the historical prospectus. Put the prospectus .docx first, or pass --file-type in file order',
    );
  }
  return files.map((_, index) => (index === prospectusIndex ? 'prospectus' : 'audit'));
}

function loadConfig(environment = process.env) {
  return {
    baseUrl: normalizeBaseUrl(environment.PAODINGAI_API_BASE_URL?.trim() || DEFAULT_PAODINGAI_API_BASE_URL),
    apiKey: requiredEnvironment('PAODINGAI_API_KEY', environment),
  };
}

function resolveOutputDir(args, inputs = []) {
  if (args.outputDir) return path.resolve(args.outputDir);
  const prospectus = inputs.find((item) => item.kind === 'prospectus');
  if (prospectus) return path.dirname(prospectus.absolutePath);
  throw new PrimeError('prime.invalid_arguments', 'download requires --output-dir when the historical prospectus path is not available');
}

function hasDuplicateBasenames(inputs) {
  const seen = new Set();
  for (const input of inputs) {
    if (seen.has(input.basename)) return true;
    seen.add(input.basename);
  }
  return false;
}

function lookupFileMap(map, index, basename, allowBasename = true) {
  if (map == null) return undefined;
  if (Array.isArray(map)) return map[index];
  if (Object.prototype.hasOwnProperty.call(map, String(index))) return map[String(index)];
  if (Object.prototype.hasOwnProperty.call(map, index)) return map[index];
  if (!allowBasename) return undefined;
  return map[basename];
}

function guessKind(filePath) {
  const lower = path.basename(filePath).toLowerCase();
  if (lower.includes('募集') || lower.includes('prospectus')) return 'prospectus';
  if (['审计', '财务', '财报', '年报', 'audit', 'annual'].some((item) => lower.includes(item))) return 'audit';
  return null;
}

function normalizeKind(value, label) {
  if (value === 'prospectus' || value === 'audit') return value;
  throw new PrimeError('prime.invalid_arguments', `Unsupported file type for ${label}: ${value}`);
}

async function validateInputFile(filePath, declaredKind = null) {
  const absolutePath = path.resolve(filePath);
  let metadata;
  try {
    metadata = await stat(absolutePath);
  } catch (error) {
    throw new PrimeError('prime.file_required', `Input file does not exist: ${filePath}`, { cause: error });
  }
  if (!metadata.isFile()) throw new PrimeError('prime.invalid_document', `Input path must be a regular file: ${filePath}`);
  if (metadata.size === 0) throw new PrimeError('prime.invalid_document', `Input file is empty: ${filePath}`);
  if (metadata.size > MAX_FILE_SIZE) throw new PrimeError('prime.file_too_large', `Input file exceeds 200 MiB: ${filePath}`);

  const extension = path.extname(absolutePath).toLowerCase();
  const basename = path.basename(absolutePath);
  const kind = declaredKind ? normalizeKind(declaredKind, basename) : guessKind(absolutePath);
  if (!kind) throw new PrimeError('prime.ambiguous_document', `Cannot determine whether this is a prospectus or audit report: ${basename}`);
  if (kind === 'prospectus' && !PROSPECTUS_EXTENSIONS.has(extension)) {
    throw new PrimeError('prime.invalid_document', 'The historical prospectus must be a DOCX file');
  }
  if (kind === 'audit' && !AUDIT_EXTENSIONS.has(extension)) {
    throw new PrimeError('prime.invalid_document', 'Audit/financial reports must be PDF, Word, or Excel files');
  }

  const handle = await open(absolutePath, 'r');
  const header = Buffer.alloc(8);
  try {
    await handle.read(header, 0, header.length, 0);
  } finally {
    await handle.close();
  }
  const valid =
    (extension === '.pdf' && header.subarray(0, 5).equals(Buffer.from('%PDF-'))) ||
    (extension === '.doc' && header.equals(COMPOUND_DOC_SIGNATURE)) ||
    (extension === '.xls' && header.equals(COMPOUND_DOC_SIGNATURE)) ||
    (['.docx', '.xlsx'].includes(extension) && ZIP_SIGNATURES.has(header.subarray(0, 4).toString('hex')));
  if (!valid) throw new PrimeError('prime.invalid_document', `File content does not match extension: ${path.basename(filePath)}`);

  return { absolutePath, basename, size: metadata.size, kind };
}

function validatePeriodsForFiles(inputs, periods) {
  const allowBasename = !hasDuplicateBasenames(inputs);
  const refs = [];
  for (const [index, input] of inputs.entries()) {
    const period = lookupFileMap(periods, index, input.basename, allowBasename);
    if (!period) {
      throw new PrimeError('prime.missing_period', `Missing period for file ${index} (${input.basename})`);
    }
    if (typeof period !== 'string') {
      throw new PrimeError('prime.invalid_period', `Period for file ${index} (${input.basename}) must be a string, for example "2024年"`);
    }
    refs.push(parseReportPeriod(period, `${index}:${input.basename}`));
  }
  if (!periods.__new__ || typeof periods.__new__ !== 'string') {
    throw new PrimeError('prime.missing_period', 'Missing new prospectus period: __new__');
  }
  const prospectusInputs = inputs.filter((item) => item.kind === 'prospectus');
  if (prospectusInputs.length !== 1) throw new PrimeError('prime.invalid_documents', 'Exactly one historical prospectus DOCX is required');
  const auditInputs = inputs.filter((item) => item.kind === 'audit');
  if (auditInputs.length === 0) throw new PrimeError('prime.invalid_documents', 'At least one audit/financial report is required');

  const prospectusRef = refs[inputs.indexOf(prospectusInputs[0])];
  const newRef = parseReportPeriod(periods.__new__, '__new__');
  const auditRefs = inputs
    .map((input, index) => (input.kind === 'audit' ? refs[index] : null))
    .filter(Boolean);
  validateReportPeriodRules(prospectusRef, newRef, auditRefs);
}

function buildIndexKeyedPayload(inputs, periods) {
  const allowBasename = !hasDuplicateBasenames(inputs);
  const periodPayload = { __new__: periods.__new__ };
  const fileTypes = [];
  for (const [index, input] of inputs.entries()) {
    const period = lookupFileMap(periods, index, input.basename, allowBasename);
    periodPayload[String(index)] = period;
    fileTypes[index] = input.kind;
  }
  return { periods: periodPayload, fileTypes };
}

function parseReportPeriod(text, label) {
  const segments = normalizePeriodText(text).split(PERIOD_SPLIT).filter(Boolean);
  if (segments.length === 0) throw new PrimeError('prime.invalid_period', `Period is empty for ${label}`);

  const parsed = segments.flatMap((segment) => parsePeriodSegment(normalizePeriodSegment(segment), label));
  validateContinuousSegments(parsed, label);
  const years = parsed.map((item) => item.year);
  const reversePeriod = years[0] > years[years.length - 1] ? 1 : 0;
  const recent = reversePeriod ? parsed[0] : parsed[parsed.length - 1];
  return {
    recent_year: recent.month > 1 ? recent.year - 1 : recent.year,
    recent_month: recent.month || 1,
    report_period: parsed.length,
    reverse_period: reversePeriod,
  };
}

function parsePeriodSegment(segment, label) {
  const range = segment.match(PERIOD_RANGE);
  if (range) {
    const startYear = Number.parseInt(range[1], 10);
    const endYear = Number.parseInt(range[2], 10);
    const endMonth = Number.parseInt(range[3] || '1', 10);
    if (startYear === endYear) throw new PrimeError('prime.invalid_period', `Invalid year range for ${label}: ${segment}`);
    if (endMonth && (endMonth < 1 || endMonth > 12)) {
      throw new PrimeError('prime.invalid_period', `Invalid month in period for ${label}: ${segment}`);
    }
    const step = endYear > startYear ? 1 : -1;
    const periods = [];
    for (let year = startYear; ; year += step) {
      periods.push({ year, month: year === endYear ? endMonth : 1 });
      if (year === endYear) break;
    }
    return periods;
  }

  const match = segment.match(PERIOD_SEGMENT);
  if (!match) throw new PrimeError('prime.invalid_period', `Cannot parse period segment for ${label}: ${segment}`);
  const year = Number.parseInt(match[1], 10);
  const startMonth = match[3] ? Number.parseInt(match[3], 10) : null;
  const endMonth = match[2] ? Number.parseInt(match[2], 10) : (match[4] ? Number.parseInt(match[4], 10) : 1);
  if ((match[2] || match[4]) && (endMonth < 1 || endMonth > 12)) {
    throw new PrimeError('prime.invalid_period', `Invalid month in period for ${label}: ${segment}`);
  }
  if (startMonth !== null && (startMonth < 1 || startMonth > endMonth)) {
    throw new PrimeError('prime.invalid_period', `Invalid month range in period for ${label}: ${segment}`);
  }
  return [{ year, month: endMonth }];
}

function validateContinuousSegments(parsed, label) {
  if (parsed.length <= 1) return;
  const years = parsed.map((item) => item.year);
  const step = years[0] < years[years.length - 1] ? 1 : -1;
  for (let index = 1; index < years.length; index += 1) {
    if (years[index] - years[index - 1] !== step) {
      throw new PrimeError('prime.invalid_period', `Periods for ${label} must be continuous years in ascending or descending order`);
    }
  }
}

function periodSet(ref, isAudit = false) {
  let reportPeriod = ref.report_period;
  if (isAudit && reportPeriod === 1) reportPeriod = 2;
  if (ref.recent_month > 1) reportPeriod -= 1;

  const periods = new Set();
  for (let index = 0; index < reportPeriod; index += 1) {
    periods.add(ref.recent_year - index);
  }
  if (ref.recent_month > 1) periods.add((ref.recent_year + 1) * 100 + ref.recent_month);
  return periods;
}

function formatPeriod(period) {
  if (period > 9999) return `${Math.floor(period / 100)}年${period % 100}月`;
  return `${period}年`;
}

function validateReportPeriodRules(prospectusRef, newRef, auditRefs) {
  const errors = [];
  if (prospectusRef.report_period < 2) errors.push('Historical prospectus period must contain at least two periods');
  if (newRef.report_period < 2) errors.push('New prospectus period must contain at least two periods');
  if (auditRefs.length === 0) errors.push('At least one audit/financial report is required');
  if (errors.length > 0) throw new PrimeError('prime.invalid_period', errors.join('; '));

  const oldPeriods = periodSet(prospectusRef);
  const newPeriods = periodSet(newRef);
  const auditPeriods = new Set();
  for (const ref of auditRefs) {
    for (const period of periodSet(ref, true)) auditPeriods.add(period);
  }

  const missing = [...newPeriods].filter((period) => !oldPeriods.has(period) && !auditPeriods.has(period));
  if (missing.length > 0) {
    throw new PrimeError(
      'prime.period_not_covered',
      `New prospectus periods are not covered by audit/financial reports: ${missing.map(formatPeriod).sort().join(', ')}`,
    );
  }
}

class GlazerPrimeClient {
  constructor({ baseUrl, apiKey, fetchImpl = fetch }) {
    this.baseUrl = baseUrl;
    this.apiKey = apiKey;
    this.fetchImpl = fetchImpl;
  }

  headers(extra = {}) {
    return {
      Authorization: `Bearer ${this.apiKey}`,
      'X-Request-Id': crypto.randomUUID(),
      ...extra,
    };
  }

  async submit(inputs, periods, fileTypes = []) {
    const form = new FormData();
    for (const [index, input] of inputs.entries()) {
      const bytes = await fs.promises.readFile(input.absolutePath);
      form.append('files', new Blob([bytes]), `${index}__${input.basename}`);
    }
    form.append('periods', JSON.stringify(periods));
    if (Array.isArray(fileTypes) && fileTypes.length > 0) form.append('file_types', JSON.stringify(fileTypes));
    const response = await this.fetchImpl(serviceUrl(this.baseUrl).href, {
      method: 'POST',
      headers: this.headers(),
      body: form,
      signal: AbortSignal.timeout(SUBMIT_REQUEST_TIMEOUT_MS),
    });
    const payload = await parseJsonResponse(response, 'prime.submit_failed');
    if (!response.ok) throw responseError(response, payload, 'prime.submit_failed');
    const data = unwrapData(payload);
    const projectId = data.project_id;
    if (!projectId) throw new PrimeError('prime.invalid_response', 'Submit response did not include project_id');
    return { project_id: projectId, status: data.status ?? 'accepted' };
  }

  async status(projectId) {
    const response = await this.fetchImpl(serviceUrl(this.baseUrl, `/${projectId}/status`).href, {
      method: 'GET',
      headers: this.headers(),
      signal: AbortSignal.timeout(STATUS_REQUEST_TIMEOUT_MS),
    });
    const payload = await parseJsonResponse(response, 'prime.status_failed');
    if (!response.ok) throw responseError(response, payload, 'prime.status_failed', projectId);
    return unwrapData(payload);
  }

  async watch(projectId, afterEventId) {
    const url = serviceUrl(this.baseUrl, `/${projectId}/watch`);
    url.searchParams.set('after_event_id', String(afterEventId));
    const response = await this.fetchImpl(url.href, {
      method: 'GET',
      headers: this.headers(),
      signal: AbortSignal.timeout(STATUS_REQUEST_TIMEOUT_MS),
    });
    const payload = await parseJsonResponse(response, 'prime.watch_failed');
    if (!response.ok) throw responseError(response, payload, 'prime.watch_failed', projectId);
    return unwrapData(payload);
  }

  async download(projectId) {
    const response = await this.fetchImpl(serviceUrl(this.baseUrl, `/${projectId}/download`).href, {
      method: 'GET',
      headers: this.headers(),
      signal: AbortSignal.timeout(DOWNLOAD_REQUEST_TIMEOUT_MS),
    });
    if (!response.ok) {
      const payload = await readMaybeJson(response);
      throw responseError(response, payload, 'prime.download_failed', projectId);
    }
    return response;
  }
}

async function parseJsonResponse(response, fallbackCode) {
  const text = await response.text();
  if (!text) throw new PrimeError(fallbackCode, 'Service returned an empty response', { httpStatus: response.status });
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new PrimeError(fallbackCode, 'Service returned invalid JSON', { cause: error, httpStatus: response.status });
  }
}

async function readMaybeJson(response) {
  const text = await response.text();
  if (!text) return {};
  try { return JSON.parse(text); } catch { return { message: text }; }
}

function unwrapData(payload) {
  if (payload && typeof payload === 'object' && payload.data && typeof payload.data === 'object') return payload.data;
  return payload;
}

function responseError(response, payload, fallbackCode, projectId = null) {
  const body = unwrapData(payload);
  const message = body?.detail ?? body?.message ?? body?.msg ?? payload?.detail ?? payload?.message ?? `Service request failed (HTTP ${response.status})`;
  const code = body?.code ?? payload?.code ?? fallbackCode;
  return new PrimeError(code, message, { projectId, httpStatus: response.status });
}

const PROGRESS_TYPE_LABELS = {
  start: '开始',
  finish: '已完成',
  log: '进行中',
  fail: '失败',
};

function progressTypeLabel(type) {
  return PROGRESS_TYPE_LABELS[type] || type || '进行中';
}

function emitProgress(payload, jsonProgress, stderr = process.stderr) {
  if (jsonProgress) {
    stderr.write(`${JSON.stringify(payload)}\n`);
    return;
  }
  const scope = payload.stage || 'glazer-prospectus-refresh';
  const typeLabel = progressTypeLabel(payload.type);
  stderr.write(`[${scope}][${typeLabel}] ${payload.message ?? payload.status ?? 'processing'}\n`);
}

function isRetryableError(error) {
  if (!error) return false;
  if (error.httpStatus === 408 || error.httpStatus === 429 || error.httpStatus >= 500) return true;
  if (error.name === 'TimeoutError' || error.name === 'AbortError') return true;
  if (error.name === 'TypeError') return true;
  if (error.cause) return isRetryableError(error.cause);
  return false;
}

async function requestWatchWithRetry(client, projectId, afterEventId) {
  let lastError;
  for (let attempt = 1; attempt <= STATUS_REQUEST_ATTEMPTS; attempt += 1) {
    try {
      return await client.watch(projectId, afterEventId);
    } catch (error) {
      lastError = error;
      if (!isRetryableError(error) || attempt === STATUS_REQUEST_ATTEMPTS) break;
      await sleep(RETRY_DELAY_MS);
    }
  }
  throw lastError;
}

async function pollWatch(projectId, client, options = {}) {
  const onProgress = options.onProgress ?? (() => {});
  const pollIntervalMs = options.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS;
  const heartbeatIntervalMs = options.heartbeatIntervalMs ?? DEFAULT_HEARTBEAT_INTERVAL_MS;
  const maxWaitMs = options.maxWaitMs ?? null;
  const startedAt = Date.now();
  let cursor = options.afterEventId ?? 0;
  let lastHeartbeatAt = startedAt;
  const seenEventIds = new Set();

  const heartbeat = () => {
    const now = Date.now();
    if (now - lastHeartbeatAt < heartbeatIntervalMs) return;
    onProgress({ stage: 'glazer-prospectus-refresh', type: 'log', message: '刷报仍在进行中，正在等待下一次进度更新。' });
    lastHeartbeatAt = now;
  };

  while (true) {
    const elapsed = Date.now() - startedAt;
    if (maxWaitMs !== null && elapsed >= maxWaitMs) {
      return { project_id: projectId, status: 'processing', max_event_id: cursor, timed_out: true };
    }

    const data = await requestWatchWithRetry(client, projectId, cursor);
    const events = Array.isArray(data.events) ? data.events : [];
    for (const event of events) {
      if (event.id !== undefined && seenEventIds.has(event.id)) continue;
      if (event.id !== undefined) seenEventIds.add(event.id);
      onProgress(event);
      lastHeartbeatAt = Date.now();
    }
    cursor = data.max_event_id ?? cursor;
    if (DOWNLOADABLE_STATUSES.has(data.status) || data.status === 'failed') {
      return { project_id: projectId, status: data.status, max_event_id: cursor };
    }
    heartbeat();
    await sleep(pollIntervalMs);
  }
}

async function downloadResult(projectId, client, outputDir) {
  const response = await client.download(projectId);
  const arrayBuffer = await response.arrayBuffer();
  const bytes = Buffer.from(arrayBuffer);
  const sha256 = crypto.createHash('sha256').update(bytes).digest('hex');
  const declaredSha256 = response.headers.get('x-file-sha256');
  const declaredSize = response.headers.get('x-file-size');
  if (declaredSha256 && declaredSha256 !== sha256) {
    throw new PrimeError('prime.checksum_mismatch', 'Downloaded file SHA-256 did not match server declaration', { projectId });
  }
  if (declaredSize && Number.parseInt(declaredSize, 10) !== bytes.length) {
    throw new PrimeError('prime.size_mismatch', 'Downloaded file size did not match server declaration', { projectId });
  }

  fs.mkdirSync(outputDir, { recursive: true });
  const filename = filenameFromDisposition(response.headers.get('content-disposition')) || `${projectId}.docx`;
  const outputPath = uniqueOutputPath(path.join(outputDir, filename));
  const tmpPath = `${outputPath}.${process.pid}.${Date.now()}.tmp`;
  await writeFile(tmpPath, bytes);
  await rename(tmpPath, outputPath);
  return { output_file: path.resolve(outputPath), file_size: bytes.length, sha256 };
}

function decodeFileName(value) {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function filenameFromDisposition(disposition) {
  if (!disposition) return null;
  const utf8 = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8) return decodeFileName(utf8[1]);
  const quoted = disposition.match(/filename="?([^";]+)"?/i);
  return quoted ? decodeFileName(quoted[1]) : null;
}

function uniqueOutputPath(targetPath) {
  if (!fs.existsSync(targetPath)) return targetPath;
  const dir = path.dirname(targetPath);
  const ext = path.extname(targetPath);
  const base = path.basename(targetPath, ext);
  for (let index = 1; index < 10_000; index += 1) {
    const candidate = path.join(dir, `${base}(${index})${ext}`);
    if (!fs.existsSync(candidate)) return candidate;
  }
  throw new PrimeError('prime.output_conflict', 'Could not allocate a unique output file name');
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function run(argv = process.argv.slice(2), environment = process.env) {
  const args = parseArgs(argv);
  if (args.help) return { status: 'help', message: usage() };
  const config = loadConfig(environment);
  const client = new GlazerPrimeClient({ baseUrl: config.baseUrl, apiKey: config.apiKey });

  if (args.mode === 'start') {
    const inputs = await Promise.all(args.files.map((filePath, index) => (
      validateInputFile(filePath, args.fileTypes[index])
    )));
    validatePeriodsForFiles(inputs, args.periods);
    const payload = buildIndexKeyedPayload(inputs, args.periods);
    const accepted = await client.submit(inputs, payload.periods, payload.fileTypes);
    process.stdout.write(`${JSON.stringify(accepted)}\n`);
    emitProgress({ stage: 'submit', type: 'finish', message: `项目已受理。project_id=${accepted.project_id}` }, args.json);
    const watched = await pollWatch(accepted.project_id, client, {
      onProgress: (event) => emitProgress(event, args.json),
    });
    if (!DOWNLOADABLE_STATUSES.has(watched.status)) return watched;
    const downloaded = await downloadResult(accepted.project_id, client, resolveOutputDir(args, inputs));
    return { ...watched, ...downloaded, status: 'downloaded' };
  }

  if (args.mode === 'watch') {
    return await pollWatch(args.projectId, client, {
      afterEventId: args.afterEventId,
      onProgress: (event) => emitProgress(event, args.json),
    });
  }

  if (args.mode === 'status') {
    return await client.status(args.projectId);
  }

  if (args.mode === 'download') {
    const downloaded = await downloadResult(args.projectId, client, resolveOutputDir(args));
    return { status: 'downloaded', project_id: args.projectId, ...downloaded };
  }

  throw new PrimeError('prime.invalid_arguments', `Unsupported command: ${args.mode}`);
}

async function main() {
  try {
    const result = await run();
    if (result.status === 'help') process.stdout.write(`${result.message}\n`);
    else process.stdout.write(`${JSON.stringify(result)}\n`);
  } catch (error) {
    const primeError = error instanceof PrimeError
      ? error
      : new PrimeError('prime.client_failed', 'Glazer client failed', { cause: error });
    process.stdout.write(`${JSON.stringify({
      status: 'failed',
      project_id: primeError.projectId,
      error_code: primeError.errorCode,
      message: primeError.message,
    })}\n`);
    process.exitCode = 1;
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await main();
}

export {
  PrimeError,
  GlazerPrimeClient,
  loadConfig,
  parseArgs,
  run,
};
