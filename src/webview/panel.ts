import type {
	CellDiffStatus,
	DiffCellLocation,
	GridCellView,
	GridRowView,
	PageSlice,
	RenderModel,
	RowFilterMode,
	SheetTabView,
	WorkbookFileView,
} from '../core/model/types';

// ─── VS Code webview API ──────────────────────────────────────────────────────

interface VsCodeApi {
	postMessage(message: OutgoingMessage): void;
}

declare function acquireVsCodeApi(): VsCodeApi;

type OutgoingMessage =
	| { type: 'ready' }
	| { type: 'setSheet'; sheetKey: string }
	| { type: 'setFilter'; filter: RowFilterMode }
	| { type: 'prevPage' }
	| { type: 'nextPage' }
	| { type: 'prevDiff' }
	| { type: 'nextDiff' }
	| { type: 'selectCell'; rowNumber: number; columnNumber: number }
	| { type: 'saveEdits'; edits: Array<{ sheetKey: string; side: 'left' | 'right'; rowNumber: number; columnNumber: number; value: string }> }
	| { type: 'swap' }
	| { type: 'reload' };

type IncomingMessage =
	| { type: 'loading'; message: string }
	| { type: 'error'; message: string }
	| { type: 'render'; payload: RenderModel; silent?: boolean; clearPendingEdits?: boolean };

// ─── State ────────────────────────────────────────────────────────────────────

const vscode = acquireVsCodeApi();

let model: RenderModel | null = null;
let isSyncingScroll = false;
interface CellPosition {
	rowNumber: number;
	columnNumber: number;
	side: 'left' | 'right';
}

let selectedCell: CellPosition | null = null;
let pendingSelectionReason: string | null = null;
let layoutSyncFrame = 0;
let shouldSyncSelectionAfterRender = false;

interface EditState extends CellPosition {
	sheetKey: string;
	cellEl: HTMLElement;
	input: HTMLInputElement;
}

let editState: EditState | null = null;

interface PendingEdit {
	sheetKey: string;
	side: 'left' | 'right';
	rowNumber: number;
	columnNumber: number;
	value: string;
}

const pendingEdits = new Map<string, PendingEdit>();

function getPendingEditKey(sheetKey: string, side: 'left' | 'right', rowNumber: number, columnNumber: number): string {
	return `${sheetKey}:${side}:${rowNumber}:${columnNumber}`;
}

const DEFAULT_STRINGS = {
	loading: 'Loading XLSX diff...',
	all: 'All',
	diffs: 'Diffs',
	same: 'Same',
	prevDiff: 'Prev Diff',
	nextDiff: 'Next Diff',
	prevPage: 'Prev Page',
	nextPage: 'Next Page',
	swap: 'Swap',
	reload: 'Reload',
	left: 'Left',
	right: 'Right',
	mergedRangesChanged: 'Merged ranges changed',
	noRowsAvailable: 'No rows available for this filter.',
	size: 'Size',
	modified: 'Modified',
	sheet: 'Sheet',
	rows: 'Rows',
	noRows: 'No rows',
	page: 'Page',
	filter: 'Filter',
	diffCells: 'Diff cells',
	diffRows: 'Diff rows',
	sameRows: 'Same rows',
	visibleRows: 'Visible rows',
	readOnly: 'Read-only',
	save: 'Save',
};

type Strings = typeof DEFAULT_STRINGS;

const STRINGS: Strings = (globalThis as Record<string, unknown>).__XLSX_DIFF_STRINGS__ as Strings ?? DEFAULT_STRINGS;

// ─── Utilities ────────────────────────────────────────────────────────────────

function escapeHtml(value: string): string {
	return String(value)
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&#39;');
}

// ─── Loading / error shells ───────────────────────────────────────────────────

function renderLoading(message: string): void {
	document.body.innerHTML = `
		<div id="app" class="loading-shell">
			<div class="loading-shell__message">${escapeHtml(message)}</div>
		</div>
	`;
}

function renderError(message: string): void {
	document.body.innerHTML = `
		<div id="app" class="empty-shell">
			<div class="empty-shell__message">${escapeHtml(message)}</div>
		</div>
	`;
}

// ─── Diff tone helpers ────────────────────────────────────────────────────────

function getColumnDiffTones(rows: GridRowView[]): Map<number, CellDiffStatus> {
	const tones = new Map<number, CellDiffStatus>();

	for (const row of rows) {
		row.cells.forEach((cell, index) => {
			if (cell.status === 'equal') {
				return;
			}

			const currentTone = tones.get(index);
			if (currentTone === 'modified' || currentTone === 'removed') {
				return;
			}

			if (cell.status === 'modified' || cell.status === 'removed') {
				tones.set(index, cell.status);
				return;
			}

			if (!currentTone) {
				tones.set(index, cell.status);
			}
		});
	}

	return tones;
}

function getDiffToneClass(diffTone: CellDiffStatus | undefined): string {
	return diffTone ? `diff-marker--${diffTone}` : '';
}

function getEffectiveDiffMarkerClass(diffTone: string, hasPending: boolean): string | null {
	if (hasPending) {
		return 'diff-marker--pending';
	}

	if (diffTone) {
		return getDiffToneClass(diffTone as CellDiffStatus);
	}

	return null;
}

function updateLabelMarker(labelEl: Element, markerClass: string | null, extraClass?: string): void {
	let markerEl = labelEl.querySelector<HTMLElement>('.diff-marker');

	if (!markerClass) {
		markerEl?.remove();
		return;
	}

	if (!markerEl) {
		markerEl = document.createElement('span');
		markerEl.setAttribute('aria-hidden', 'true');
		labelEl.insertBefore(markerEl, labelEl.firstChild);
	}

	markerEl.className = `diff-marker ${markerClass}${extraClass ? ` ${extraClass}` : ''}`;
}

function shouldHighlightCell(cell: GridCellView, side: 'left' | 'right', isHighlighted: boolean): boolean {
	if (!isHighlighted) {
		return false;
	}

	if (cell.status === 'modified') {
		return true;
	}

	if (cell.status === 'added') {
		return side === 'right';
	}

	if (cell.status === 'removed') {
		return side === 'left';
	}

	return false;
}

function getSideCellClass(cell: GridCellView, side: 'left' | 'right', isHighlighted: boolean): string {
	const classes = ['grid__cell'];

	if (cell.status === 'modified') {
		classes.push('grid__cell--modified');
	} else if (cell.status === 'added') {
		classes.push(side === 'right' ? 'grid__cell--added' : 'grid__cell--ghost');
	} else if (cell.status === 'removed') {
		classes.push(side === 'left' ? 'grid__cell--removed' : 'grid__cell--ghost');
	} else {
		classes.push('grid__cell--equal');
	}

	if (isHighlighted) {
		classes.push('grid__cell--highlighted');
	}

	return classes.join(' ');
}

// ─── Cell content ─────────────────────────────────────────────────────────────

function getCellTooltip(address: string, value: string, formula: string | null): string {
	const lines = [address];

	if (value) {
		lines.push(value);
	}

	if (formula) {
		lines.push(`fx ${formula}`);
	}

	return lines.join('\n');
}

function renderCellValue(value: string, formula: string | null): string {
	if (!value && !formula) {
		return '';
	}

	return `${
		value ? `<span class="grid__cell-value">${escapeHtml(value)}</span>` : ''
	}${formula ? `<span class="cell__formula" title="${escapeHtml(formula)}">fx</span>` : ''}`;
}

function getCellView(rowNumber: number, columnNumber: number): GridCellView | null {
	const row = model?.page.rows.find((item) => item.rowNumber === rowNumber);
	return row?.cells[columnNumber - 1] ?? null;
}

// ─── Label helpers ────────────────────────────────────────────────────────────

function getFilterLabel(filter: RowFilterMode): string {
	switch (filter) {
		case 'diffs':
			return STRINGS.diffs;
		case 'same':
			return STRINGS.same;
		case 'all':
		default:
			return STRINGS.all;
	}
}

function getSheetTooltip(sheet: SheetTabView): string {
	const s = STRINGS;
	return `${sheet.label} · ${sheet.diffCellCount} ${s.diffCells} · ${sheet.diffRowCount} ${s.diffRows}`;
}

// ─── Selection helpers ────────────────────────────────────────────────────────

function getPreferredSelectionSide(rowNumber: number, columnNumber: number): 'left' | 'right' {
	const cell = getCellView(rowNumber, columnNumber);
	if (cell?.status === 'removed') {
		return 'left';
	}

	return 'right';
}

function getHighlightedDiffSelection(): CellPosition | null {
	if (!model?.page.highlightedDiffCell) {
		return null;
	}

	const { rowNumber, columnNumber } = model.page.highlightedDiffCell;
	return {
		rowNumber,
		columnNumber,
		side: getPreferredSelectionSide(rowNumber, columnNumber),
	};
}

function getSelectedCellElements(): Element[] {
	if (!selectedCell) {
		return [];
	}

	return Array.from(
		document.querySelectorAll(
			`.pane[data-side="${selectedCell.side}"] [data-role="grid-cell"][data-row-number="${selectedCell.rowNumber}"][data-column-number="${selectedCell.columnNumber}"]`,
		),
	);
}

function clearSelectedCells(): void {
	for (const cell of document.querySelectorAll('.grid__cell--selected')) {
		cell.classList.remove('grid__cell--selected');
	}

	for (const cell of document.querySelectorAll('[data-role="grid-cell"][aria-selected="true"]')) {
		cell.setAttribute('aria-selected', 'false');
	}
}

// ─── Scroll helpers ───────────────────────────────────────────────────────────

function clampScrollPosition(value: number, maxValue: number): number {
	return Math.max(0, Math.min(value, Math.max(maxValue, 0)));
}

function getPaneScrollState(): { top: number; left: number } | null {
	const pane = document.querySelector('.pane__table');
	if (!pane) {
		return null;
	}

	return {
		top: (pane as HTMLElement).scrollTop,
		left: (pane as HTMLElement).scrollLeft,
	};
}

interface ScrollUpdate {
	pane: HTMLElement;
	top: number;
	left: number;
}

function setPaneScrollPositions(updates: ScrollUpdate[]): void {
	if (updates.length === 0) {
		return;
	}

	isSyncingScroll = true;
	for (const { pane, top, left } of updates) {
		pane.scrollTop = clampScrollPosition(top, pane.scrollHeight - pane.clientHeight);
		pane.scrollLeft = clampScrollPosition(left, pane.scrollWidth - pane.clientWidth);
	}

	requestAnimationFrame(() => {
		isSyncingScroll = false;
	});
}

function restorePaneScrollState(scrollState: { top: number; left: number } | null): void {
	if (!scrollState) {
		return;
	}

	setPaneScrollPositions(
		Array.from(document.querySelectorAll<HTMLElement>('.pane__table')).map((pane) => ({
			pane,
			top: scrollState.top,
			left: scrollState.left,
		})),
	);
}

function getStickyPaneInsets(pane: HTMLElement): { top: number; left: number } {
	const headerRow = pane.querySelector('thead tr');
	const firstColumn = pane.querySelector('thead th:first-child');

	return {
		top: headerRow?.getBoundingClientRect().height ?? 0,
		left: firstColumn?.getBoundingClientRect().width ?? 0,
	};
}

function getDesiredPaneScrollPosition(pane: HTMLElement, element: HTMLElement): { top: number; left: number } {
	const paneRect = pane.getBoundingClientRect();
	const elementRect = element.getBoundingClientRect();
	const stickyInsets = getStickyPaneInsets(pane);
	let top = pane.scrollTop;
	let left = pane.scrollLeft;

	if (elementRect.top < paneRect.top + stickyInsets.top) {
		top -= paneRect.top + stickyInsets.top - elementRect.top;
	} else if (elementRect.bottom > paneRect.bottom) {
		top += elementRect.bottom - paneRect.bottom;
	}

	if (elementRect.left < paneRect.left + stickyInsets.left) {
		left -= paneRect.left + stickyInsets.left - elementRect.left;
	} else if (elementRect.right > paneRect.right) {
		left += elementRect.right - paneRect.right;
	}

	return { top, left };
}

function revealSelectedCells(elements: Element[]): void {
	setPaneScrollPositions(
		elements
			.map((element) => {
				const pane = element.closest<HTMLElement>('.pane__table');
				if (!pane) {
					return null;
				}

				return {
					pane,
					...getDesiredPaneScrollPosition(pane, element as HTMLElement),
				};
			})
			.filter((x): x is ScrollUpdate => x !== null),
	);
}

function applySelectedCell({ reveal = false }: { reveal?: boolean } = {}): void {
	clearSelectedCells();

	if (!selectedCell) {
		return;
	}

	const elements = getSelectedCellElements();
	if (elements.length === 0) {
		return;
	}

	for (const element of elements) {
		element.classList.add('grid__cell--selected');
		element.setAttribute('aria-selected', 'true');
	}

	if (reveal) {
		revealSelectedCells(elements);
	}
}

function syncSelectedCellAfterRender(): void {
	const reason = pendingSelectionReason;
	pendingSelectionReason = null;

	if (reason === 'highlighted-diff') {
		selectedCell = getHighlightedDiffSelection() ?? selectedCell;
	}

	if (!selectedCell || getSelectedCellElements().length === 0) {
		selectedCell = getHighlightedDiffSelection();
	}

	applySelectedCell({ reveal: reason === 'highlighted-diff' });
}

// ─── Row height sync ──────────────────────────────────────────────────────────

function getGridRows(side: 'left' | 'right'): HTMLElement[] {
	return Array.from(
		document.querySelectorAll<HTMLElement>(`.pane[data-side="${side}"] [data-role="grid-row"]`),
	);
}

function syncTableRowHeights(): void {
	const leftRows = getGridRows('left');
	const rightRows = getGridRows('right');
	const rowCount = Math.min(leftRows.length, rightRows.length);

	for (const row of [...leftRows, ...rightRows]) {
		row.style.height = '';
	}

	for (let index = 0; index < rowCount; index += 1) {
		const leftRow = leftRows[index]!;
		const rightRow = rightRows[index]!;
		const syncedHeight = Math.ceil(
			Math.max(
				leftRow.getBoundingClientRect().height,
				rightRow.getBoundingClientRect().height,
			),
		);

		if (syncedHeight <= 0) {
			continue;
		}

		leftRow.style.height = `${syncedHeight}px`;
		rightRow.style.height = `${syncedHeight}px`;
	}
}

function scheduleLayoutSync({ afterRender = false }: { afterRender?: boolean } = {}): void {
	shouldSyncSelectionAfterRender = shouldSyncSelectionAfterRender || afterRender;

	if (layoutSyncFrame) {
		cancelAnimationFrame(layoutSyncFrame);
	}

	layoutSyncFrame = requestAnimationFrame(() => {
		const syncSelection = shouldSyncSelectionAfterRender;

		layoutSyncFrame = 0;
		shouldSyncSelectionAfterRender = false;
		syncTableRowHeights();

		if (syncSelection) {
			syncSelectedCellAfterRender();
			return;
		}

		applySelectedCell();
	});
}

// ─── Edit mode ────────────────────────────────────────────────────────────────

function canEditCell(status: CellDiffStatus, side: 'left' | 'right'): boolean {
	if (!model) {
		return false;
	}

	const isReadonly = side === 'left' ? model.leftFile.isReadonly : model.rightFile.isReadonly;
	if (isReadonly) {
		return false;
	}

	const sheetName = side === 'left' ? model.activeSheet.leftName : model.activeSheet.rightName;
	if (!sheetName) {
		return false;
	}

	// Ghost cells cannot be edited
	if (status === 'added' && side === 'left') {
		return false;
	}

	if (status === 'removed' && side === 'right') {
		return false;
	}

	return true;
}

function getCellModelValue(rowNumber: number, columnNumber: number, side: 'left' | 'right'): string {
	const cell = getCellView(rowNumber, columnNumber);
	if (!cell) {
		return '';
	}

	return side === 'left' ? cell.leftValue : cell.rightValue;
}

function updateSaveButtonState(): void {
	const saveBtn = document.querySelector<HTMLButtonElement>('[data-action="save-edits"]');
	if (!saveBtn) {
		return;
	}

	const hasPending = pendingEdits.size > 0;
	saveBtn.disabled = !hasPending;
	saveBtn.classList.toggle('is-dirty', hasPending);
}

function getCellFormula(rowNumber: number, columnNumber: number, side: 'left' | 'right'): string | null {
	const cell = getCellView(rowNumber, columnNumber);
	if (!cell) {
		return null;
	}

	return side === 'left' ? cell.leftFormula : cell.rightFormula;
}

function syncCellDisplay(
	cellEl: HTMLElement,
	sheetKey: string,
	rowNumber: number,
	columnNumber: number,
	side: 'left' | 'right',
): void {
	const content = cellEl.querySelector<HTMLElement>('.grid__cell-content');
	if (!content) {
		return;
	}

	const pendingKey = getPendingEditKey(sheetKey, side, rowNumber, columnNumber);
	const pendingEdit = pendingEdits.get(pendingKey);
	const value = pendingEdit ? pendingEdit.value : getCellModelValue(rowNumber, columnNumber, side);
	const formula = pendingEdit ? null : getCellFormula(rowNumber, columnNumber, side);

	content.innerHTML = renderCellValue(value, formula);
	cellEl.classList.toggle('grid__cell--pending', Boolean(pendingEdit));
}

function finishEdit(
	{ mode, clearSelection = false }: { mode: 'commit' | 'cancel'; clearSelection?: boolean },
): void {
	const session = editState;
	if (!session) {
		return;
	}

	editState = null;
	const { sheetKey, rowNumber, columnNumber, side, cellEl, input } = session;
	cellEl.classList.remove('grid__cell--editing');

	if (mode === 'commit') {
		commitEdit(sheetKey, rowNumber, columnNumber, side, input.value);
		if (clearSelection) {
			selectedCell = null;
			clearSelectedCells();
		}
		return;
	}

	syncCellDisplay(cellEl, sheetKey, rowNumber, columnNumber, side);
}

function cancelEdit(): void {
	finishEdit({ mode: 'cancel' });
}

function clearSelectedCellValue(): void {
	if (!model || !selectedCell) {
		return;
	}

	const { rowNumber, columnNumber, side } = selectedCell;
	const cell = getCellView(rowNumber, columnNumber);
	if (!cell || !canEditCell(cell.status, side)) {
		return;
	}

	commitEdit(model.activeSheet.key, rowNumber, columnNumber, side, '');
}

function isClearSelectedCellKey(event: KeyboardEvent): boolean {
	if (event.altKey || event.ctrlKey || event.metaKey) {
		return false;
	}

	return (
		event.key === 'Backspace' ||
		event.key === 'Delete' ||
		event.code === 'Backspace' ||
		event.code === 'Delete'
	);
}

function applyPendingEditStyles(): void {
	const activeSheetKey = model?.activeSheet.key;
	if (!activeSheetKey) {
		return;
	}

	const pendingSheetKeys = new Set<string>();
	const pendingRowsBySide = {
		left: new Set<number>(),
		right: new Set<number>(),
	};
	const pendingColumnsBySide = {
		left: new Set<number>(),
		right: new Set<number>(),
	};

	for (const pendingEdit of pendingEdits.values()) {
		pendingSheetKeys.add(pendingEdit.sheetKey);
		if (pendingEdit.sheetKey !== activeSheetKey) {
			continue;
		}

		pendingRowsBySide[pendingEdit.side].add(pendingEdit.rowNumber);
		pendingColumnsBySide[pendingEdit.side].add(pendingEdit.columnNumber);
	}

	// 1. Cells – update content and pending border
	for (const cellEl of document.querySelectorAll<HTMLElement>('[data-role="grid-cell"]')) {
		const pane = cellEl.closest<HTMLElement>('[data-side]');
		const side = pane?.getAttribute('data-side') as 'left' | 'right' | null;
		if (!side || cellEl.classList.contains('grid__cell--editing')) {
			continue;
		}

		const rowNumber = Number(cellEl.getAttribute('data-row-number'));
		const columnNumber = Number(cellEl.getAttribute('data-column-number'));
		syncCellDisplay(cellEl, activeSheetKey, rowNumber, columnNumber, side);
	}

	// 2. Row number headers – square marker with priority: pending > diff
	for (const rowEl of document.querySelectorAll<HTMLElement>('[data-role="grid-row"]')) {
		const rowNumber = Number(rowEl.getAttribute('data-row-number'));
		const pane = rowEl.closest<HTMLElement>('[data-side]');
		const side = pane?.getAttribute('data-side') as 'left' | 'right' | null;
		if (!side) {
			continue;
		}

		const hasPending = pendingRowsBySide[side].has(rowNumber);
		const rowHeader = rowEl.querySelector<HTMLElement>('th[data-diff-tone]');
		if (!rowHeader) {
			continue;
		}

		const diffTone = rowHeader.getAttribute('data-diff-tone') ?? '';
		rowHeader.classList.toggle('grid__row-number--pending', hasPending);
		const labelEl = rowHeader.querySelector('.grid__row-label');
		if (labelEl) {
			updateLabelMarker(labelEl, getEffectiveDiffMarkerClass(diffTone, hasPending));
		}
	}

	// 3. Column headers – square marker with priority: pending > diff
	for (const pane of document.querySelectorAll<HTMLElement>('.pane__table[data-side]')) {
		const side = pane.getAttribute('data-side') as 'left' | 'right';

		for (const colHeader of pane.querySelectorAll<HTMLElement>('thead th[data-column-number]')) {
			const columnNumber = Number(colHeader.getAttribute('data-column-number'));
			const diffTone = colHeader.getAttribute('data-diff-tone') ?? '';
			const hasPending = pendingColumnsBySide[side].has(columnNumber);
			const markerClass = getEffectiveDiffMarkerClass(diffTone, hasPending);
			// Ensure --diff class is present when pending adds a new marker to a non-diff column
			colHeader.classList.toggle('grid__column--diff', Boolean(diffTone) || hasPending);
			colHeader.classList.toggle('grid__column--pending', hasPending);
			const labelEl = colHeader.querySelector('.grid__column-label');
			if (labelEl) {
				updateLabelMarker(labelEl, markerClass);
			}
		}
	}

	// 4. Sheet tabs – square marker with priority: pending > diff
	for (const tabEl of document.querySelectorAll<HTMLElement>('[data-action="set-sheet"]')) {
		const sheetKey = tabEl.getAttribute('data-sheet-key');
		const hasPending = sheetKey ? pendingSheetKeys.has(sheetKey) : false;
		const diffTone = tabEl.getAttribute('data-diff-tone') ?? '';
		const markerClass = getEffectiveDiffMarkerClass(diffTone, hasPending);
		let markerEl = tabEl.querySelector<HTMLElement>('.tab__marker');

		if (!markerClass) {
			markerEl?.remove();
		} else {
			if (!markerEl) {
				markerEl = document.createElement('span');
				markerEl.setAttribute('aria-hidden', 'true');
				tabEl.insertBefore(markerEl, tabEl.firstChild);
			}

			markerEl.className = `diff-marker ${markerClass} tab__marker`;
		}
	}
}

function commitEdit(
	sheetKey: string,
	rowNumber: number,
	columnNumber: number,
	side: 'left' | 'right',
	value: string,
): void {
	const key = getPendingEditKey(sheetKey, side, rowNumber, columnNumber);
	const modelValue = getCellModelValue(rowNumber, columnNumber, side);

	if (value === modelValue) {
		// Value reverted to original — remove any existing pending for this cell
		pendingEdits.delete(key);
	} else {
		pendingEdits.set(key, { sheetKey, side, rowNumber, columnNumber, value });
	}

	applyPendingEditStyles();
	updateSaveButtonState();
}

function triggerSave(): void {
	if (pendingEdits.size === 0) {
		return;
	}

	const edits = Array.from(pendingEdits.values());
	pendingEdits.clear();
	updateSaveButtonState();
	vscode.postMessage({ type: 'saveEdits', edits });
}

function enterEditMode(
	cellEl: HTMLElement,
	rowNumber: number,
	columnNumber: number,
	side: 'left' | 'right',
	currentValue: string,
): void {
	finishEdit({ mode: 'commit' });

	const capturedSheetKey = model?.activeSheet.key;
	const content = cellEl.querySelector<HTMLElement>('.grid__cell-content');
	if (!content || !capturedSheetKey) {
		return;
	}

	content.innerHTML = '';
	const input = document.createElement('input');
	input.type = 'text';
	input.className = 'grid__cell-input';
	input.value = currentValue;
	content.appendChild(input);
	cellEl.classList.add('grid__cell--editing');
	editState = { rowNumber, columnNumber, side, sheetKey: capturedSheetKey, cellEl, input };

	input.focus();
	input.select();

	input.addEventListener('keydown', (e: KeyboardEvent) => {
		if (e.key === 'Enter' || e.key === 'Tab') {
			e.preventDefault();
			finishEdit({ mode: 'commit', clearSelection: true });
		} else if (e.key === 'Escape') {
			e.preventDefault();
			finishEdit({ mode: 'cancel' });
		}
	});

	input.addEventListener('blur', () => {
		if (editState?.input === input) {
			finishEdit({ mode: 'commit', clearSelection: true });
		}
	});
}

// ─── Table rendering ──────────────────────────────────────────────────────────

function renderTable(side: 'left' | 'right'): string {
	if (!model || model.page.rows.length === 0) {
		return `<div class="empty-table">${escapeHtml(STRINGS.noRowsAvailable)}</div>`;
	}

	const diffColumnTones = getColumnDiffTones(model.page.rows);
	const headerColumns = model.page.columns
		.map((column, index) => {
			const diffTone = diffColumnTones.get(index);

			return `<th class="grid__column ${diffTone ? `grid__column--diff grid__column--${diffTone}` : ''}" data-column-number="${index + 1}" data-diff-tone="${diffTone ?? ''}">
				<span class="grid__column-label">
					${diffTone ? `<span class="diff-marker ${getDiffToneClass(diffTone)}" aria-hidden="true"></span>` : ''}
					<span>${escapeHtml(column)}</span>
				</span>
			</th>`;
		})
		.join('');

	const bodyRows = model.page.rows
		.map((row) => {
			const rowClasses = [
				row.hasDiff ? 'row--diff' : '',
				row.hasDiff ? `row--diff-${row.diffTone}` : '',
				row.isHighlighted ? 'row--highlight' : '',
			]
				.filter(Boolean)
				.join(' ');

			const cells = row.cells
				.map((cell, columnIndex) => {
					const value = side === 'left' ? cell.leftValue : cell.rightValue;
					const formula = side === 'left' ? cell.leftFormula : cell.rightFormula;
					const highlightCell = shouldHighlightCell(cell, side, row.isHighlighted);

					const cellClass = getSideCellClass(cell, side, highlightCell);
					const cellTooltip = getCellTooltip(cell.address, value, formula);
					const editable = canEditCell(cell.status, side) ? 'true' : 'false';

					return `<td title="${escapeHtml(cellTooltip)}" class="${cellClass}" data-role="grid-cell" data-row-number="${row.rowNumber}" data-column-number="${columnIndex + 1}" data-cell-status="${cell.status}" data-editable="${editable}" aria-selected="false">
						<div class="grid__cell-content">${renderCellValue(value, formula)}</div>
					</td>`;
				})
				.join('');

			return `
				<tr class="${rowClasses}" data-role="grid-row" data-row-number="${row.rowNumber}" data-row-has-diff="${row.hasDiff ? 'true' : 'false'}">
					<th class="grid__row-number ${row.hasDiff ? `grid__row-number--diff grid__row-number--${row.diffTone}` : ''}" data-diff-tone="${row.hasDiff ? row.diffTone : ''}">
						<span class="grid__row-label">
							${row.hasDiff ? `<span class="diff-marker ${getDiffToneClass(row.diffTone)}" aria-hidden="true"></span>` : ''}
							<span>${row.rowNumber}</span>
						</span>
					</th>
					${cells}
				</tr>
			`;
		})
		.join('');

	return `
		<table class="grid">
			<thead>
				<tr>
					<th class="grid__row-number">#</th>
					${headerColumns}
				</tr>
			</thead>
			<tbody>${bodyRows}</tbody>
		</table>
	`;
}

// ─── File card rendering ──────────────────────────────────────────────────────

function renderFileCard(file: WorkbookFileView): string {
	const s = STRINGS;
	const readOnlyIcon = file.isReadonly
		? `<span class="codicon codicon-lock file-card__lock" title="${escapeHtml(s.readOnly)}" aria-label="${escapeHtml(s.readOnly)}"></span>`
		: '';

	return `
		<div class="file-card">
			<div class="file-card__name" title="${escapeHtml(file.filePath)}">
				<span class="file-card__name-text">${escapeHtml(file.fileName)}</span>
				${readOnlyIcon}
			</div>
			<div class="file-card__meta">
				<div class="file-card__path" title="${escapeHtml(file.filePath)}">${escapeHtml(file.filePath)}</div>
				<div class="file-card__facts">
					<span>${escapeHtml(s.size)}: ${escapeHtml(file.fileSizeLabel)}</span>
					${file.detailLabel && file.detailValue ? `<span>${escapeHtml(file.detailLabel)}: ${escapeHtml(file.detailValue)}</span>` : ''}
					<span>${escapeHtml(s.modified)}: ${escapeHtml(file.modifiedTimeLabel)}</span>
				</div>
			</div>
		</div>
	`;
}

// ─── Pane rendering ───────────────────────────────────────────────────────────

function renderPane(title: string, side: 'left' | 'right'): string {
	return `
		<section class="pane" data-side="${side}">
			<div class="pane__header">
				<div class="pane__title">${escapeHtml(title)}</div>
				${model!.activeSheet.mergedRangesChanged ? `<span class="badge badge--warn">${escapeHtml(STRINGS.mergedRangesChanged)}</span>` : ''}
			</div>
			<div class="pane__table" data-side="${side}">${renderTable(side)}</div>
		</section>
	`;
}

// ─── Toolbar rendering ────────────────────────────────────────────────────────

interface ToolbarButtonOptions {
	action: string;
	icon: string;
	label: string;
	active?: boolean;
	disabled?: boolean;
	filter?: string;
}

function renderToolbarButton({ action, icon, label, active = false, disabled = false, filter }: ToolbarButtonOptions): string {
	return `<button class="toolbar__button ${active ? 'is-active' : ''}" data-action="${action}" ${filter ? `data-filter="${filter}"` : ''} ${disabled ? 'disabled' : ''}>
		<span class="codicon ${icon} toolbar__button-icon" aria-hidden="true"></span>
		<span>${escapeHtml(label)}</span>
	</button>`;
}

function renderToolbar(): string {
	const m = model!;
	const s = STRINGS;

	return `
		<header class="toolbar">
			<div class="toolbar__group">
				${renderToolbarButton({ action: 'set-filter', filter: 'all', icon: 'codicon-list-flat', label: s.all, active: m.filter === 'all' })}
				${renderToolbarButton({ action: 'set-filter', filter: 'diffs', icon: 'codicon-diff-multiple', label: s.diffs, active: m.filter === 'diffs' })}
				${renderToolbarButton({ action: 'set-filter', filter: 'same', icon: 'codicon-check-all', label: s.same, active: m.filter === 'same' })}
			</div>
			<div class="toolbar__group">
				${renderToolbarButton({ action: 'prev-diff', icon: 'codicon-arrow-up', label: s.prevDiff, disabled: !m.canPrevDiff })}
				${renderToolbarButton({ action: 'next-diff', icon: 'codicon-arrow-down', label: s.nextDiff, disabled: !m.canNextDiff })}
				${renderToolbarButton({ action: 'prev-page', icon: 'codicon-arrow-left', label: s.prevPage, disabled: !m.canPrevPage })}
				${renderToolbarButton({ action: 'next-page', icon: 'codicon-arrow-right', label: s.nextPage, disabled: !m.canNextPage })}
				${renderToolbarButton({ action: 'swap', icon: 'codicon-arrow-swap', label: s.swap })}
				${renderToolbarButton({ action: 'reload', icon: 'codicon-refresh', label: s.reload })}			${renderToolbarButton({ action: 'save-edits', icon: 'codicon-save', label: s.save, disabled: true })}			</div>
		</header>
	`;
}

// ─── Status bar rendering ─────────────────────────────────────────────────────

function renderTabs(): string {
	return model!.sheets
		.map(
			(sheet) => `
				<button
					class="tab tab--${sheet.diffTone} ${sheet.isActive ? 'is-active' : ''} ${sheet.hasDiff ? 'has-diff' : ''}"
					data-action="set-sheet"
					data-sheet-key="${escapeHtml(sheet.key)}"
					data-diff-tone="${sheet.hasDiff ? sheet.diffTone : ''}"
					title="${escapeHtml(getSheetTooltip(sheet))}"
				>
					${sheet.hasDiff ? `<span class="diff-marker ${getDiffToneClass(sheet.diffTone)} tab__marker" aria-hidden="true"></span>` : ''}
					<span class="tab__label">${escapeHtml(sheet.label)}</span>
				</button>
			`,
		)
		.join('');
}

function renderStatus(): string {
	const m = model!;
	const s = STRINGS;
	const rowRangeLabel = m.page.visibleRowCount === 0 ? s.noRows : m.page.rangeLabel;

	return `
		<footer class="footer">
			<div class="tabs">${renderTabs()}</div>
			<div class="status">
				<span><strong>${escapeHtml(s.sheet)}:</strong> ${escapeHtml(m.activeSheet.label)}</span>
				<span><strong>${escapeHtml(s.rows)}:</strong> ${escapeHtml(rowRangeLabel)}</span>
				<span><strong>${escapeHtml(s.page)}:</strong> ${m.page.currentPage} / ${m.page.totalPages}</span>
				<span><strong>${escapeHtml(s.filter)}:</strong> ${escapeHtml(getFilterLabel(m.filter))}</span>
				<span><strong>${escapeHtml(s.diffRows)}:</strong> ${m.page.diffRowCount}</span>
				<span><strong>${escapeHtml(s.sameRows)}:</strong> ${m.page.sameRowCount}</span>
				<span><strong>${escapeHtml(s.visibleRows)}:</strong> ${m.page.visibleRowCount}</span>
			</div>
		</footer>
	`;
}

// ─── App render ───────────────────────────────────────────────────────────────

function renderApp({ silent = false }: { silent?: boolean } = {}): void {
	if (!model) {
		renderLoading(STRINGS.loading);
		return;
	}

	finishEdit({ mode: 'commit' });

	const previousPaneScrollState = getPaneScrollState();
	const s = STRINGS;

	document.body.innerHTML = `
		<div id="app" class="app">
			${renderToolbar()}
			<section class="files">
				${renderFileCard(model.leftFile)}
				${renderFileCard(model.rightFile)}
			</section>
			<section class="panes">
				${renderPane(s.left, 'left')}
				${renderPane(s.right, 'right')}
			</section>
			${renderStatus()}
		</div>
	`;

	attachPaneScrollSync();
	restorePaneScrollState(previousPaneScrollState);
	applyPendingEditStyles();
	updateSaveButtonState();
	scheduleLayoutSync({ afterRender: true });
}

// ─── Scroll sync ──────────────────────────────────────────────────────────────

function syncPaneScroll(sourcePane: HTMLElement): void {
	if (isSyncingScroll) {
		return;
	}

	const panes = Array.from(document.querySelectorAll<HTMLElement>('.pane__table'));
	if (panes.length < 2) {
		return;
	}

	isSyncingScroll = true;
	for (const pane of panes) {
		if (pane === sourcePane) {
			continue;
		}

		pane.scrollTop = clampScrollPosition(sourcePane.scrollTop, pane.scrollHeight - pane.clientHeight);
		pane.scrollLeft = clampScrollPosition(sourcePane.scrollLeft, pane.scrollWidth - pane.clientWidth);
	}

	requestAnimationFrame(() => {
		isSyncingScroll = false;
	});
}

function attachPaneScrollSync(): void {
	const panes = Array.from(document.querySelectorAll<HTMLElement>('.pane__table'));
	for (const pane of panes) {
		pane.addEventListener(
			'scroll',
			() => {
				syncPaneScroll(pane);
			},
			{ passive: true },
		);
	}
}

// ─── Global event listeners ───────────────────────────────────────────────────

window.addEventListener('resize', () => {
	if (!model) {
		return;
	}

	scheduleLayoutSync();
});

window.addEventListener('message', (event: MessageEvent<IncomingMessage>) => {
	const message = event.data;

	if (message.type === 'loading') {
		renderLoading(message.message);
		return;
	}

	if (message.type === 'error') {
		renderError(message.message);
		return;
	}

	if (message.type === 'render') {
		model = message.payload;
		if (message.clearPendingEdits) {
			pendingEdits.clear();
		}

		renderApp({ silent: message.silent });
	}
});

document.addEventListener('click', (event: MouseEvent) => {
	const eventTarget = event.target instanceof Element ? event.target : null;
	if (!eventTarget) {
		return;
	}

	const cellTarget = eventTarget.closest<HTMLElement>('[data-role="grid-cell"]');
	if (cellTarget) {
		const pane = cellTarget.closest<HTMLElement>('[data-side]');
		const side = pane?.getAttribute('data-side') as 'left' | 'right' | null;
		if (!side) {
			return;
		}

		const rowNumber = Number(cellTarget.getAttribute('data-row-number'));
		const columnNumber = Number(cellTarget.getAttribute('data-column-number'));
		const cellStatus = cellTarget.getAttribute('data-cell-status') as CellDiffStatus;

		selectedCell = { rowNumber, columnNumber, side };
		pendingSelectionReason = null;
		applySelectedCell();

		if (
			cellStatus !== 'equal' &&
			(
				model?.page.highlightedDiffCell?.rowNumber !== rowNumber ||
				model?.page.highlightedDiffCell?.columnNumber !== columnNumber
			)
		) {
			vscode.postMessage({ type: 'selectCell', rowNumber, columnNumber });
		}

		return;
	}

	const target = eventTarget.closest<HTMLElement>('[data-action]');
	if (!target) {
		return;
	}

	const action = target.getAttribute('data-action');
	switch (action) {
		case 'set-filter':
			vscode.postMessage({ type: 'setFilter', filter: target.getAttribute('data-filter') as RowFilterMode });
			return;
		case 'set-sheet':
			vscode.postMessage({ type: 'setSheet', sheetKey: target.getAttribute('data-sheet-key')! });
			return;
		case 'prev-page':
			vscode.postMessage({ type: 'prevPage' });
			return;
		case 'next-page':
			vscode.postMessage({ type: 'nextPage' });
			return;
		case 'prev-diff':
			pendingSelectionReason = 'highlighted-diff';
			vscode.postMessage({ type: 'prevDiff' });
			return;
		case 'next-diff':
			pendingSelectionReason = 'highlighted-diff';
			vscode.postMessage({ type: 'nextDiff' });
			return;
		case 'swap':
			vscode.postMessage({ type: 'swap' });
			return;
		case 'reload':
			vscode.postMessage({ type: 'reload' });
			return;
		case 'save-edits':
			triggerSave();
			return;
	}
});

document.addEventListener('keydown', (event: KeyboardEvent) => {
	if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
		event.preventDefault();
		if (editState) {
			finishEdit({ mode: 'commit', clearSelection: true });
		}

		triggerSave();
		return;
	}

	if (isClearSelectedCellKey(event) && !editState && selectedCell) {
		event.preventDefault();
		clearSelectedCellValue();
	}
});

document.addEventListener('dblclick', (event: MouseEvent) => {
	const eventTarget = event.target instanceof Element ? event.target : null;
	if (!eventTarget) {
		return;
	}

	const cellTarget = eventTarget.closest<HTMLElement>('[data-role="grid-cell"]');
	if (!cellTarget) {
		return;
	}

	const cellStatus = cellTarget.getAttribute('data-cell-status') as CellDiffStatus;
	const editable = cellTarget.getAttribute('data-editable') === 'true';
	if (!editable) {
		return;
	}

	const pane = cellTarget.closest<HTMLElement>('[data-side]');
	const side = pane?.getAttribute('data-side') as 'left' | 'right' | null;
	if (!side) {
		return;
	}

	const rowNumber = Number(cellTarget.getAttribute('data-row-number'));
	const columnNumber = Number(cellTarget.getAttribute('data-column-number'));

	// Use pending value if the cell has been staged but not yet saved
	const pendingKey = getPendingEditKey(model!.activeSheet.key, side, rowNumber, columnNumber);
	const pendingEdit = pendingEdits.get(pendingKey);
	const currentValue = pendingEdit ? pendingEdit.value : getCellModelValue(rowNumber, columnNumber, side);

	selectedCell = { rowNumber, columnNumber, side };
	applySelectedCell();
	enterEditMode(cellTarget, rowNumber, columnNumber, side, currentValue);
});

// ─── Init ─────────────────────────────────────────────────────────────────────

renderLoading(STRINGS.loading);
vscode.postMessage({ type: 'ready' });
