const vscode = acquireVsCodeApi();

let model = null;

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

function getSideCellClass(cell, side) {
	if (cell.status === 'modified') {
		return 'cell cell--modified';
	}

	if (cell.status === 'added') {
		return side === 'right' ? 'cell cell--added' : 'cell cell--ghost';
	}

	if (cell.status === 'removed') {
		return side === 'left' ? 'cell cell--removed' : 'cell cell--ghost';
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

	const headerColumns = model.page.columns
		.map((column) => `<th class="grid__column">${escapeHtml(column)}</th>`)
		.join('');

	const bodyRows = model.page.rows
		.map((row) => {
			const rowClasses = [
				row.hasDiff ? 'row--diff' : '',
				row.isHighlighted ? 'row--highlight' : '',
			]
				.filter(Boolean)
				.join(' ');

			const cells = row.cells
				.map((cell) => {
					const value = side === 'left' ? cell.leftValue : cell.rightValue;
					const formula = side === 'left' ? cell.leftFormula : cell.rightFormula;

					return `<td title="${escapeHtml(cell.address)}"><div class="${getSideCellClass(
						cell,
						side,
					)}">${renderCellValue(value, formula)}</div></td>`;
				})
				.join('');

			return `
				<tr class="${rowClasses}">
					<th class="grid__row-number">${row.rowNumber}</th>
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

function renderPane(title, sideName, side) {
	return `
		<section class="pane">
			<div class="pane__header">
				<div>
					<div class="pane__title">${escapeHtml(title)}</div>
					<div class="pane__subtitle">${escapeHtml(sideName ?? 'No matching sheet')}</div>
				</div>
				${model.activeSheet.mergedRangesChanged ? '<span class="badge badge--warn">Merged ranges changed</span>' : ''}
			</div>
			<div class="pane__table">${renderTable(side)}</div>
		</section>
	`;
}

function renderSummary() {
	return `
		<section class="summary">
			<span class="badge"><strong>${model.summary.totalSheets}</strong> sheets</span>
			<span class="badge"><strong>${model.summary.diffSheets}</strong> changed sheets</span>
			<span class="badge"><strong>${model.summary.diffRows}</strong> changed rows</span>
			<span class="badge"><strong>${model.summary.diffCells}</strong> changed cells</span>
		</section>
	`;
}

function renderToolbar() {
	return `
		<header class="toolbar">
			<div class="toolbar__group">
				<button class="toolbar__button ${model.filter === 'all' ? 'is-active' : ''}" data-action="set-filter" data-filter="all">All</button>
				<button class="toolbar__button ${model.filter === 'diffs' ? 'is-active' : ''}" data-action="set-filter" data-filter="diffs">Diffs</button>
				<button class="toolbar__button ${model.filter === 'same' ? 'is-active' : ''}" data-action="set-filter" data-filter="same">Same</button>
			</div>
			<div class="toolbar__group">
				<button class="toolbar__button" data-action="prev-diff" ${model.canPrevDiff ? '' : 'disabled'}>Prev Diff</button>
				<button class="toolbar__button" data-action="next-diff" ${model.canNextDiff ? '' : 'disabled'}>Next Diff</button>
				<button class="toolbar__button" data-action="prev-page" ${model.canPrevPage ? '' : 'disabled'}>Prev Page</button>
				<button class="toolbar__button" data-action="next-page" ${model.canNextPage ? '' : 'disabled'}>Next Page</button>
				<button class="toolbar__button" data-action="swap">Swap</button>
				<button class="toolbar__button" data-action="reload">Reload</button>
			</div>
		</header>
	`;
}

function renderFiles() {
	return `
		<section class="files">
			<div class="file-card">
				<div class="file-card__label">Left workbook</div>
				<div class="file-card__name" title="${escapeHtml(model.leftFile.filePath)}">${escapeHtml(model.leftFile.fileName)}</div>
				<div class="file-card__meta">
					<div class="file-card__path" title="${escapeHtml(model.leftFile.filePath)}">${escapeHtml(model.leftFile.filePath)}</div>
					<div class="file-card__facts">
						<span>Size: ${escapeHtml(model.leftFile.fileSizeLabel)}</span>
						<span>Modified: ${escapeHtml(model.leftFile.modifiedTimeLabel)}</span>
					</div>
				</div>
			</div>
			<div class="file-card">
				<div class="file-card__label">Right workbook</div>
				<div class="file-card__name" title="${escapeHtml(model.rightFile.filePath)}">${escapeHtml(model.rightFile.fileName)}</div>
				<div class="file-card__meta">
					<div class="file-card__path" title="${escapeHtml(model.rightFile.filePath)}">${escapeHtml(model.rightFile.filePath)}</div>
					<div class="file-card__facts">
						<span>Size: ${escapeHtml(model.rightFile.fileSizeLabel)}</span>
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
					class="tab ${sheet.isActive ? 'is-active' : ''} ${sheet.hasDiff ? 'has-diff' : ''}"
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
			${renderSummary()}
			${renderFiles()}
			<section class="panes">
				${renderPane('Left', model.activeSheet.leftName, 'left')}
				${renderPane('Right', model.activeSheet.rightName, 'right')}
			</section>
			${renderStatus()}
		</div>
	`;
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
