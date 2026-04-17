const vscode = acquireVsCodeApi();

let model = null;
let isSyncingScroll = false;
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
	page: 'Page',
	filter: 'Filter',
	diffRows: 'Diff rows',
	sameRows: 'Same rows',
	visibleRows: 'Visible rows',
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

function renderCellValue(value, formula) {
	if (!value && !formula) {
		return '';
	}

	return `${escapeHtml(value)}${
		formula
			? `<span class="cell__formula" title="${escapeHtml(formula)}">fx</span>`
			: ''
	}`;
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
				.map((cell) => {
					const value = side === 'left' ? cell.leftValue : cell.rightValue;
					const formula = side === 'left' ? cell.leftFormula : cell.rightFormula;
					const highlightCell = shouldHighlightCell(
						cell,
						side,
						row.isHighlighted,
					);

					const cellClass = getSideCellClass(cell, side, highlightCell);

					return `<td title="${escapeHtml(cell.address)}" class="${cellClass}">
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
			<div class="file-card">
				<div class="file-card__name" title="${escapeHtml(model.leftFile.filePath)}">${escapeHtml(model.leftFile.fileName)}</div>
				<div class="file-card__meta">
					<div class="file-card__path" title="${escapeHtml(model.leftFile.filePath)}">${escapeHtml(model.leftFile.filePath)}</div>
					<div class="file-card__facts">
						<span>${escapeHtml(STRINGS.size)}: ${escapeHtml(model.leftFile.fileSizeLabel)}</span>
						${model.leftFile.detailLabel && model.leftFile.detailValue ? `<span>${escapeHtml(model.leftFile.detailLabel)}: ${escapeHtml(model.leftFile.detailValue)}</span>` : ''}
						<span>${escapeHtml(STRINGS.modified)}: ${escapeHtml(model.leftFile.modifiedTimeLabel)}</span>
					</div>
				</div>
			</div>
			<div class="file-card">
				<div class="file-card__name" title="${escapeHtml(model.rightFile.filePath)}">${escapeHtml(model.rightFile.fileName)}</div>
				<div class="file-card__meta">
					<div class="file-card__path" title="${escapeHtml(model.rightFile.filePath)}">${escapeHtml(model.rightFile.filePath)}</div>
					<div class="file-card__facts">
						<span>${escapeHtml(STRINGS.size)}: ${escapeHtml(model.rightFile.fileSizeLabel)}</span>
						${model.rightFile.detailLabel && model.rightFile.detailValue ? `<span>${escapeHtml(model.rightFile.detailLabel)}: ${escapeHtml(model.rightFile.detailValue)}</span>` : ''}
						<span>${escapeHtml(STRINGS.modified)}: ${escapeHtml(model.rightFile.modifiedTimeLabel)}</span>
					</div>
				</div>
			</div>
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
					title="${escapeHtml(`${sheet.label} · ${sheet.diffCellCount} changed cells · ${sheet.diffRowCount} changed rows`)}"
				>
					${sheet.hasDiff ? `<span class="diff-marker ${getDiffToneClass(sheet.diffTone)} tab__marker" aria-hidden="true"></span>` : ''}
					<span class="tab__label">${escapeHtml(sheet.label)}</span>
				</button>
			`,
		)
		.join('');
}

function renderStatus() {
	return `
		<footer class="footer">
			<div class="tabs">${renderTabs()}</div>
			<div class="status">
				<span><strong>${escapeHtml(STRINGS.sheet)}:</strong> ${escapeHtml(model.activeSheet.label)}</span>
				<span><strong>${escapeHtml(STRINGS.rows)}:</strong> ${model.page.rangeLabel}</span>
				<span><strong>${escapeHtml(STRINGS.page)}:</strong> ${model.page.currentPage} / ${model.page.totalPages}</span>
				<span><strong>${escapeHtml(STRINGS.filter)}:</strong> ${escapeHtml(model.filter)}</span>
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
			vscode.postMessage({ type: 'prevDiff' });
			return;
		case 'next-diff':
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
