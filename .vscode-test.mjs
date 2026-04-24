import { execFileSync } from 'node:child_process';
import { existsSync, realpathSync } from 'node:fs';
import * as path from 'node:path';
import { defineConfig } from '@vscode/test-cli';

function resolveVSCodeExecutablePath() {
	if (process.env.VSCODE_EXECUTABLE_PATH) {
		return process.env.VSCODE_EXECUTABLE_PATH;
	}

	try {
		const locator = process.platform === 'win32' ? 'where' : 'which';
		const cliPath = execFileSync(locator, ['code'], {
			encoding: 'utf8',
		})
			.split(/\r?\n/)
			.find(Boolean)
			?.trim();

		if (!cliPath) {
			return undefined;
		}

		const resolvedCliPath = realpathSync(cliPath);

		if (process.platform === 'darwin') {
			const executablePath = path.join(
				path.dirname(
					path.dirname(path.dirname(path.dirname(resolvedCliPath))),
				),
				'MacOS',
				'Electron',
			);
			return existsSync(executablePath) ? executablePath : undefined;
		}

		if (process.platform === 'linux') {
			const executablePath = path.join(
				path.dirname(path.dirname(resolvedCliPath)),
				'code',
			);
			return existsSync(executablePath) ? executablePath : undefined;
		}
	} catch {
		return undefined;
	}

	return undefined;
}

const vscodeExecutablePath = resolveVSCodeExecutablePath();

export default defineConfig({
	files: 'out/test/**/*.test.cjs',
	...(vscodeExecutablePath
		? {
				useInstallation: {
					fromPath: vscodeExecutablePath,
				},
			}
		: {}),
});
