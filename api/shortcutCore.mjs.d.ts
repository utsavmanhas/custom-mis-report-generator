export const XLSX_MIME: string;

export function isShortcutConfigured(env?: NodeJS.ProcessEnv): boolean;
export function verifyShortcut(env?: NodeJS.ProcessEnv): Promise<Record<string, unknown>>;
export function startShortcutJob(input: unknown, env?: NodeJS.ProcessEnv): Promise<Record<string, unknown>>;
export function getShortcutStatus(runId: string, env?: NodeJS.ProcessEnv): Promise<Record<string, unknown>>;
export function downloadShortcutResult(runId: string, env?: NodeJS.ProcessEnv): Promise<Buffer>;
