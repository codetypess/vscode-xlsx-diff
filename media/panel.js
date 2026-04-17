const vscode = acquireVsCodeApi();

let model = null;
let isSyncingScroll = false;

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
	if (cell.status === 'modified') {
		return `cell cell--modified ${isHighlighted ? 'cell--highlighted' : ''}`.trim();
	}

	if (cell.status === 'added') {
		return side === 'right'
			? `cell cell--added ${isHighlighted ? 'cell--highlighted' : ''}`.trim()
			: 'cell cell--ghost';
	}

	if (cell.status === 'removed') {
		return side === 'left'
			? `cell cell--removed ${isHighlighted ? 'cell--highlighted' : ''}`.trim()
			: 'cell cell--ghost';
	}

	return 'cell';
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
		return '<div class="empty-table">No rows available for this filter.</div>';
	}

	const diffColumnTones = getColumnDiffTones(model.page.rows);
	const headerColumns = model.page.columns
		.map((column, index) => {
			const diffTone = diffColumnTones.get(index);

			return `<th class="grid__column ${diffTone ? `grid__column--diff grid__column--${diffTone}` : ''}">
				<span class="grid__column-label">
					${diffTone ? '<span class="grid__diff-marker" aria-hidden="true"></span>' : ''}
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

					return `<td title="${escapeHtml(cell.address)}"><div class="${getSideCellClass(
						cell,
						side,
						highlightCell,
					)}">${renderCellValue(value, formula)}</div></td>`;
				})
				.join('');

			return `
				<tr class="${rowClasses}">
					<th class="grid__row-number ${row.hasDiff ? `grid__row-number--diff grid__row-number--${row.diffTone}` : ''}">
						<span class="grid__row-label">
							${row.hasDiff ? '<span class="grid__diff-marker" aria-hidden="true"></span>' : ''}
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
				${model.activeSheet.mergedRangesChanged ? '<span class="badge badge--warn">Merged ranges changed</span>' : ''}
			</div>
			<div class="pane__table">${renderTable(side)}</div>
		</section>
	`;
}

function renderToolbar() {
	return `
		<header class="toolbar">
			<div class="toolbar__group">
				<button class="toolbar__button ${model.filter === 'all' ? 'is-active' : ''}" data-action="set-filter" data-filter="all"><span class="toolbar__button-icon" aria-hidden="true">#</span><span>All</span></button>
				<button class="toolbar__button ${model.filter === 'diffs' ? 'is-active' : ''}" data-action="set-filter" data-filter="diffs"><span class="toolbar__button-icon" aria-hidden="true">≠</span><span>Diffs</span></button>
				<button class="toolbar__button ${model.filter === 'same' ? 'is-active' : ''}" data-action="set-filter" data-filter="same"><span class="toolbar__button-icon" aria-hidden="true">=</span><span>Same</span></button>
			</div>
			<div class="toolbar__group">
				<button class="toolbar__button" data-action="prev-diff" ${model.canPrevDiff ? '' : 'disabled'}><span class="toolbar__button-icon" aria-hidden="true">↑</span><span>Prev Diff</span></button>
				<button class="toolbar__button" data-action="next-diff" ${model.canNextDiff ? '' : 'disabled'}><span class="toolbar__button-icon" aria-hidden="true">↓</span><span>Next Diff</span></button>
				<button class="toolbar__button" data-action="prev-page" ${model.canPrevPage ? '' : 'disabled'}><span class="toolbar__button-icon" aria-hidden="true">←</span><span>Prev Page</span></button>
				<button class="toolbar__button" data-action="next-page" ${model.canNextPage ? '' : 'disabled'}><span class="toolbar__button-icon" aria-hidden="true">→</span><span>Next Page</span></button>
				<button class="toolbar__button" data-action="swap"><span class="toolbar__button-icon" aria-hidden="true">⇄</span><span>Swap</span></button>
				<button class="toolbar__button" data-action="reload"><span class="toolbar__button-icon" aria-hidden="true">↻</span><span>Reload</span></button>
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
						<span>Size: ${escapeHtml(model.leftFile.fileSizeLabel)}</span>
						${model.leftFile.detailLabel && model.leftFile.detailValue ? `<span>${escapeHtml(model.leftFile.detailLabel)}: ${escapeHtml(model.leftFile.detailValue)}</span>` : ''}
						<span>Modified: ${escapeHtml(model.leftFile.modifiedTimeLabel)}</span>
					</div>
				</div>
			</div>
			<div class="file-card">
				<div class="file-card__name" title="${escapeHtml(model.rightFile.filePath)}">${escapeHtml(model.rightFile.fileName)}</div>
				<div class="file-card__meta">
					<div class="file-card__path" title="${escapeHtml(model.rightFile.filePath)}">${escapeHtml(model.rightFile.filePath)}</div>
					<div class="file-card__facts">
						<span>Size: ${escapeHtml(model.rightFile.fileSizeLabel)}</span>
						${model.rightFile.detailLabel && model.rightFile.detailValue ? `<span>${escapeHtml(model.rightFile.detailLabel)}: ${escapeHtml(model.rightFile.detailValue)}</span>` : ''}
						<span>Modified: ${escapeHtml(model.rightFile.modifiedTimeLabel)}</span>
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
				<span><strong>Sheet:</strong> ${escapeHtml(model.activeSheet.label)}</span>
				<span><strong>Rows:</strong> ${model.page.rangeLabel}</span>
				<span><strong>Page:</strong> ${model.page.currentPage} / ${model.page.totalPages}</span>
				<span><strong>Filter:</strong> ${escapeHtml(model.filter)}</span>
				<span><strong>Diff rows:</strong> ${model.page.diffRowCount}</span>
				<span><strong>Same rows:</strong> ${model.page.sameRowCount}</span>
				<span><strong>Visible rows:</strong> ${model.page.visibleRowCount}</span>
			</div>
		</footer>
	`;
}

function renderApp() {
	if (!model) {
		renderLoading('Loading XLSX diff...');
		return;
	}

	document.body.innerHTML = `
		<div id="app" class="app">
			${renderToolbar()}
			${renderFiles()}
			<section class="panes">
				${renderPane('Left', 'left')}
				${renderPane('Right', 'right')}
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

renderLoading('Loading XLSX diff...');
vscode.postMessage({ type: 'ready' });
