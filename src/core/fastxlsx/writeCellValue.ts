import * as vscode from 'vscode';
import { getCellAddress } from '../model/cells';

/**
 * Writes a new display value to a specific cell in a local .xlsx file.
 * Only local `file://` URIs are supported; read-only/git URIs must be rejected before calling.
 */
export async function writeCellValue(
	fileUri: vscode.Uri,
	sheetName: string,
	rowNumber: number,
	columnNumber: number,
	value: string,
): Promise<void> {
	if (fileUri.scheme !== 'file') {
		throw new Error('Cell editing is only supported for local files.');
	}

	const { Workbook } = await import('fastxlsx');
	const workbook = await Workbook.open(fileUri.fsPath);
	const sheet = workbook.getSheet(sheetName);
	const address = getCellAddress(rowNumber, columnNumber);
	sheet.cell(address).setValue(value);
	await workbook.save(fileUri.fsPath);
}
