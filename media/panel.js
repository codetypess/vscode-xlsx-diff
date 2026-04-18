const vscode = acquireVsCodeApi();

let model = null;
let isSyncingScroll = false;
let selectedCell = null;
let pendingSelectionReason = null;
const STRINGS = globalThis.__XLSX_DIFF_STRINGS__ ?? {
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
};

function escapeHtml(value) {
	return String(value)
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&#39;');
}

function renderLoading(message) {
	document.body.innerHTML = `
		<div id="app" class="loading-shell">
			<div class="loading-shell__message">${escapeHtml(message)}</div>
		</div>
	`;
}

function renderError(message) {
	document.body.innerHTML = `
		<div id="app" class="empty-shell">
			<div class="empty-shell__message">${escapeHtml(message)}</div>
		</div>
	`;
}

function getColumnDiffTones(rows) {
	const tones = new Map();

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

function getDiffToneClass(diffTone) {
	return diffTone ? `diff-marker--${diffTone}` : '';
}

function shouldHighlightCell(cell, side, isHighlighted) {
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

function getSideCellClass(cell, side, isHighlighted) {
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

function getCellTooltip(address, value, formula) {
	const lines = [address];

	if (value) {
		lines.push(value);
	}

	if (formula) {
		lines.push(`fx ${formula}`);
	}

	return lines.join('\n');
}

function renderCellValue(value, formula) {
	if (!value && !formula) {
		return '';
	}

	return `${
		value ? `<span class="grid__cell-value">${escapeHtml(value)}</span>` : ''
	}${formula ? `<span class="cell__formula" title="${escapeHtml(formula)}">fx</span>` : ''}`;
}

function getFilterLabel(filter) {
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

function getSheetTooltip(sheet) {
	return `${sheet.label} · ${sheet.diffCellCount} ${STRINGS.diffCells} · ${sheet.diffRowCount} ${STRINGS.diffRows}`;
}

function getHighlightedDiffSelection() {
	if (!model || model.page.highlightedDiffRow === null) {
		return null;
	}

	const highlightedRow = model.page.rows.find(
		(row) => row.rowNumber === model.page.highlightedDiffRow,
	);
	if (!highlightedRow || highlightedRow.cells.length === 0) {
		return null;
	}

	const diffColumnIndex = highlightedRow.cells.findIndex(
		(cell) => cell.status !== 'equal',
	);
	return {
		rowNumber: highlightedRow.rowNumber,
		columnNumber: diffColumnIndex >= 0 ? diffColumnIndex + 1 : 1,
	};
}

function getSelectedCellElements() {
	if (!selectedCell) {
		return [];
	}

	return Array.from(
		document.querySelectorAll(
			`[data-role="grid-cell"][data-row-number="${selectedCell.rowNumber}"][data-column-number="${selectedCell.columnNumber}"]`,
		),
	);
}

function clearSelectedCells() {
	for (const cell of document.querySelectorAll('.grid__cell--selected')) {
		cell.classList.remove('grid__cell--selected');
	}

	for (const cell of document.querySelectorAll(
		'[data-role="grid-cell"][aria-selected="true"]',
	)) {
		cell.setAttribute('aria-selected', 'false');
	}
}

function revealSelectedCells(elements) {
	for (const element of elements) {
		element.scrollIntoView({
			block: 'nearest',
			inline: 'center',
		});
	}
}

function applySelectedCell({ reveal = false } = {}) {
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

function syncSelectedCellAfterRender() {
	const shouldRevealSelection = pendingSelectionReason === 'highlighted-diff';
	if (pendingSelectionReason === 'highlighted-diff') {
		selectedCell = getHighlightedDiffSelection() ?? selectedCell;
	}

	if (!selectedCell || getSelectedCellElements().length === 0) {
		selectedCell = getHighlightedDiffSelection();
	}

	applySelectedCell({ reveal: shouldRevealSelection });
	pendingSelectionReason = null;
}

function renderTable(side) {
	if (!model || model.page.rows.length === 0) {
		return `<div class="empty-table">${escapeHtml(STRINGS.noRowsAvailable)}</div>`;
	}

	const diffColumnTones = getColumnDiffTones(model.page.rows);
	const headerColumns = model.page.columns
		.map((column, index) => {
			const diffTone = diffColumnTones.get(index);

			return `<th class="grid__column ${diffTone ? `grid__column--diff grid__column--${diffTone}` : ''}">
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
					const highlightCell = shouldHighlightCell(
						cell,
						side,
						row.isHighlighted,
					);

					const cellClass = getSideCellClass(cell, side, highlightCell);
					const cellTooltip = getCellTooltip(cell.address, value, formula);

					return `<td title="${escapeHtml(cellTooltip)}" class="${cellClass}" data-role="grid-cell" data-row-number="${row.rowNumber}" data-column-number="${columnIndex + 1}" aria-selected="false">
						<div class="grid__cell-content">${renderCellValue(value, formula)}</div>
					</td>`;
				})
				.join('');

			return `
				<tr class="${rowClasses}">
					<th class="grid__row-number ${row.hasDiff ? `grid__row-number--diff grid__row-number--${row.diffTone}` : ''}">
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

function renderFileCard(file) {
	const readOnlyIcon = file.isReadonly
		? `<span class="codicon codicon-lock file-card__lock" title="${escapeHtml(STRINGS.readOnly)}" aria-label="${escapeHtml(STRINGS.readOnly)}"></span>`
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
					<span>${escapeHtml(STRINGS.size)}: ${escapeHtml(file.fileSizeLabel)}</span>
					${file.detailLabel && file.detailValue ? `<span>${escapeHtml(file.detailLabel)}: ${escapeHtml(file.detailValue)}</span>` : ''}
					<span>${escapeHtml(STRINGS.modified)}: ${escapeHtml(file.modifiedTimeLabel)}</span>
				</div>
			</div>
		</div>
	`;
}

function renderPane(title, side) {
	return `
		<section class="pane">
			<div class="pane__header">
				<div class="pane__title">${escapeHtml(title)}</div>
				${model.activeSheet.mergedRangesChanged ? `<span class="badge badge--warn">${escapeHtml(STRINGS.mergedRangesChanged)}</span>` : ''}
			</div>
			<div class="pane__table">${renderTable(side)}</div>
		</section>
	`;
}

function renderToolbarButton({
	action,
	icon,
	label,
	active = false,
	disabled = false,
	filter,
}) {
	return `<button class="toolbar__button ${active ? 'is-active' : ''}" data-action="${action}" ${filter ? `data-filter="${filter}"` : ''} ${disabled ? 'disabled' : ''}>
		<span class="codicon ${icon} toolbar__button-icon" aria-hidden="true"></span>
		<span>${escapeHtml(label)}</span>
	</button>`;
}

function renderToolbar() {
	return `
		<header class="toolbar">
			<div class="toolbar__group">
				${renderToolbarButton({ action: 'set-filter', filter: 'all', icon: 'codicon-list-flat', label: STRINGS.all, active: model.filter === 'all' })}
				${renderToolbarButton({ action: 'set-filter', filter: 'diffs', icon: 'codicon-diff-multiple', label: STRINGS.diffs, active: model.filter === 'diffs' })}
				${renderToolbarButton({ action: 'set-filter', filter: 'same', icon: 'codicon-check-all', label: STRINGS.same, active: model.filter === 'same' })}
			</div>
			<div class="toolbar__group">
				${renderToolbarButton({ action: 'prev-diff', icon: 'codicon-arrow-up', label: STRINGS.prevDiff, disabled: !model.canPrevDiff })}
				${renderToolbarButton({ action: 'next-diff', icon: 'codicon-arrow-down', label: STRINGS.nextDiff, disabled: !model.canNextDiff })}
				${renderToolbarButton({ action: 'prev-page', icon: 'codicon-arrow-left', label: STRINGS.prevPage, disabled: !model.canPrevPage })}
				${renderToolbarButton({ action: 'next-page', icon: 'codicon-arrow-right', label: STRINGS.nextPage, disabled: !model.canNextPage })}
				${renderToolbarButton({ action: 'swap', icon: 'codicon-arrow-swap', label: STRINGS.swap })}
				${renderToolbarButton({ action: 'reload', icon: 'codicon-refresh', label: STRINGS.reload })}
			</div>
		</header>
	`;
}

function renderFiles() {
	return `
		<section class="files">
			${renderFileCard(model.leftFile)}
			${renderFileCard(model.rightFile)}
		</section>
	`;
}

function renderTabs() {
	return model.sheets
		.map(
			(sheet) => `
				<button
					class="tab tab--${sheet.diffTone} ${sheet.isActive ? 'is-active' : ''} ${sheet.hasDiff ? 'has-diff' : ''}"
					data-action="set-sheet"
					data-sheet-key="${escapeHtml(sheet.key)}"
					title="${escapeHtml(getSheetTooltip(sheet))}"
				>
					${sheet.hasDiff ? `<span class="diff-marker ${getDiffToneClass(sheet.diffTone)} tab__marker" aria-hidden="true"></span>` : ''}
					<span class="tab__label">${escapeHtml(sheet.label)}</span>
				</button>
			`,
		)
		.join('');
}

function renderStatus() {
	const rowRangeLabel =
		model.page.visibleRowCount === 0 ? STRINGS.noRows : model.page.rangeLabel;

	return `
		<footer class="footer">
			<div class="tabs">${renderTabs()}</div>
			<div class="status">
				<span><strong>${escapeHtml(STRINGS.sheet)}:</strong> ${escapeHtml(model.activeSheet.label)}</span>
				<span><strong>${escapeHtml(STRINGS.rows)}:</strong> ${escapeHtml(rowRangeLabel)}</span>
				<span><strong>${escapeHtml(STRINGS.page)}:</strong> ${model.page.currentPage} / ${model.page.totalPages}</span>
				<span><strong>${escapeHtml(STRINGS.filter)}:</strong> ${escapeHtml(getFilterLabel(model.filter))}</span>
				<span><strong>${escapeHtml(STRINGS.diffRows)}:</strong> ${model.page.diffRowCount}</span>
				<span><strong>${escapeHtml(STRINGS.sameRows)}:</strong> ${model.page.sameRowCount}</span>
				<span><strong>${escapeHtml(STRINGS.visibleRows)}:</strong> ${model.page.visibleRowCount}</span>
			</div>
		</footer>
	`;
}

function renderApp() {
	if (!model) {
		renderLoading(STRINGS.loading);
		return;
	}

	document.body.innerHTML = `
		<div id="app" class="app">
			${renderToolbar()}
			${renderFiles()}
			<section class="panes">
				${renderPane(STRINGS.left, 'left')}
				${renderPane(STRINGS.right, 'right')}
			</section>
			${renderStatus()}
		</div>
	`;

	attachPaneScrollSync();
	syncSelectedCellAfterRender();
}

function syncPaneScroll(sourcePane) {
	if (isSyncingScroll) {
		return;
	}

	const panes = Array.from(document.querySelectorAll('.pane__table'));
	if (panes.length < 2) {
		return;
	}

	isSyncingScroll = true;
	for (const pane of panes) {
		if (pane === sourcePane) {
			continue;
		}

		pane.scrollTop = sourcePane.scrollTop;
		pane.scrollLeft = sourcePane.scrollLeft;
	}

	requestAnimationFrame(() => {
		isSyncingScroll = false;
	});
}

function attachPaneScrollSync() {
	const panes = Array.from(document.querySelectorAll('.pane__table'));
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

window.addEventListener('message', (event) => {
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
		renderApp();
	}
});

document.addEventListener('click', (event) => {
	const cellTarget = event.target.closest('[data-role="grid-cell"]');
	if (cellTarget) {
		selectedCell = {
			rowNumber: Number(cellTarget.getAttribute('data-row-number')),
			columnNumber: Number(cellTarget.getAttribute('data-column-number')),
		};
		pendingSelectionReason = null;
		applySelectedCell();
		return;
	}

	const target = event.target.closest('[data-action]');
	if (!target) {
		return;
	}

	const action = target.getAttribute('data-action');
	switch (action) {
		case 'set-filter':
			vscode.postMessage({
				type: 'setFilter',
				filter: target.getAttribute('data-filter'),
			});
			return;
		case 'set-sheet':
			vscode.postMessage({
				type: 'setSheet',
				sheetKey: target.getAttribute('data-sheet-key'),
			});
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
	}
});

renderLoading(STRINGS.loading);
vscode.postMessage({ type: 'ready' });
