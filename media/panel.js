"use strict";
(() => {
  // src/webview/panel.ts
  var vscode = acquireVsCodeApi();
  var model = null;
  var isSyncingScroll = false;
  var selectedCell = null;
  var pendingSelectionReason = null;
  var layoutSyncFrame = 0;
  var shouldSyncSelectionAfterRender = false;
  var editState = null;
  var DEFAULT_STRINGS = {
    loading: "Loading XLSX diff...",
    all: "All",
    diffs: "Diffs",
    same: "Same",
    prevDiff: "Prev Diff",
    nextDiff: "Next Diff",
    prevPage: "Prev Page",
    nextPage: "Next Page",
    swap: "Swap",
    reload: "Reload",
    left: "Left",
    right: "Right",
    mergedRangesChanged: "Merged ranges changed",
    noRowsAvailable: "No rows available for this filter.",
    size: "Size",
    modified: "Modified",
    sheet: "Sheet",
    rows: "Rows",
    noRows: "No rows",
    page: "Page",
    filter: "Filter",
    diffCells: "Diff cells",
    diffRows: "Diff rows",
    sameRows: "Same rows",
    visibleRows: "Visible rows",
    readOnly: "Read-only"
  };
  var STRINGS = globalThis.__XLSX_DIFF_STRINGS__ ?? DEFAULT_STRINGS;
  function escapeHtml(value) {
    return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;");
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
    const tones = /* @__PURE__ */ new Map();
    for (const row of rows) {
      row.cells.forEach((cell, index) => {
        if (cell.status === "equal") {
          return;
        }
        const currentTone = tones.get(index);
        if (currentTone === "modified" || currentTone === "removed") {
          return;
        }
        if (cell.status === "modified" || cell.status === "removed") {
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
    return diffTone ? `diff-marker--${diffTone}` : "";
  }
  function shouldHighlightCell(cell, side, isHighlighted) {
    if (!isHighlighted) {
      return false;
    }
    if (cell.status === "modified") {
      return true;
    }
    if (cell.status === "added") {
      return side === "right";
    }
    if (cell.status === "removed") {
      return side === "left";
    }
    return false;
  }
  function getSideCellClass(cell, side, isHighlighted) {
    const classes = ["grid__cell"];
    if (cell.status === "modified") {
      classes.push("grid__cell--modified");
    } else if (cell.status === "added") {
      classes.push(side === "right" ? "grid__cell--added" : "grid__cell--ghost");
    } else if (cell.status === "removed") {
      classes.push(side === "left" ? "grid__cell--removed" : "grid__cell--ghost");
    } else {
      classes.push("grid__cell--equal");
    }
    if (isHighlighted) {
      classes.push("grid__cell--highlighted");
    }
    return classes.join(" ");
  }
  function getCellTooltip(address, value, formula) {
    const lines = [address];
    if (value) {
      lines.push(value);
    }
    if (formula) {
      lines.push(`fx ${formula}`);
    }
    return lines.join("\n");
  }
  function renderCellValue(value, formula) {
    if (!value && !formula) {
      return "";
    }
    return `${value ? `<span class="grid__cell-value">${escapeHtml(value)}</span>` : ""}${formula ? `<span class="cell__formula" title="${escapeHtml(formula)}">fx</span>` : ""}`;
  }
  function getFilterLabel(filter) {
    switch (filter) {
      case "diffs":
        return STRINGS.diffs;
      case "same":
        return STRINGS.same;
      case "all":
      default:
        return STRINGS.all;
    }
  }
  function getSheetTooltip(sheet) {
    const s = STRINGS;
    return `${sheet.label} \xB7 ${sheet.diffCellCount} ${s.diffCells} \xB7 ${sheet.diffRowCount} ${s.diffRows}`;
  }
  function getHighlightedDiffSelection() {
    if (!model?.page.highlightedDiffCell) {
      return null;
    }
    return {
      rowNumber: model.page.highlightedDiffCell.rowNumber,
      columnNumber: model.page.highlightedDiffCell.columnNumber
    };
  }
  function getSelectedCellElements() {
    if (!selectedCell) {
      return [];
    }
    return Array.from(
      document.querySelectorAll(
        `[data-role="grid-cell"][data-row-number="${selectedCell.rowNumber}"][data-column-number="${selectedCell.columnNumber}"]`
      )
    );
  }
  function clearSelectedCells() {
    for (const cell of document.querySelectorAll(".grid__cell--selected")) {
      cell.classList.remove("grid__cell--selected");
    }
    for (const cell of document.querySelectorAll('[data-role="grid-cell"][aria-selected="true"]')) {
      cell.setAttribute("aria-selected", "false");
    }
  }
  function clampScrollPosition(value, maxValue) {
    return Math.max(0, Math.min(value, Math.max(maxValue, 0)));
  }
  function getPaneScrollState() {
    const pane = document.querySelector(".pane__table");
    if (!pane) {
      return null;
    }
    return {
      top: pane.scrollTop,
      left: pane.scrollLeft
    };
  }
  function setPaneScrollPositions(updates) {
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
  function restorePaneScrollState(scrollState) {
    if (!scrollState) {
      return;
    }
    setPaneScrollPositions(
      Array.from(document.querySelectorAll(".pane__table")).map((pane) => ({
        pane,
        top: scrollState.top,
        left: scrollState.left
      }))
    );
  }
  function getStickyPaneInsets(pane) {
    const headerRow = pane.querySelector("thead tr");
    const firstColumn = pane.querySelector("thead th:first-child");
    return {
      top: headerRow?.getBoundingClientRect().height ?? 0,
      left: firstColumn?.getBoundingClientRect().width ?? 0
    };
  }
  function getDesiredPaneScrollPosition(pane, element) {
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
  function revealSelectedCells(elements) {
    setPaneScrollPositions(
      elements.map((element) => {
        const pane = element.closest(".pane__table");
        if (!pane) {
          return null;
        }
        return {
          pane,
          ...getDesiredPaneScrollPosition(pane, element)
        };
      }).filter((x) => x !== null)
    );
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
      element.classList.add("grid__cell--selected");
      element.setAttribute("aria-selected", "true");
    }
    if (reveal) {
      revealSelectedCells(elements);
    }
  }
  function syncSelectedCellAfterRender() {
    const reason = pendingSelectionReason;
    pendingSelectionReason = null;
    if (reason === "highlighted-diff") {
      selectedCell = getHighlightedDiffSelection() ?? selectedCell;
    }
    if (!selectedCell || getSelectedCellElements().length === 0) {
      selectedCell = getHighlightedDiffSelection();
    }
    applySelectedCell({ reveal: reason === "highlighted-diff" });
  }
  function getGridRows(side) {
    return Array.from(
      document.querySelectorAll(`.pane[data-side="${side}"] [data-role="grid-row"]`)
    );
  }
  function syncTableRowHeights() {
    const leftRows = getGridRows("left");
    const rightRows = getGridRows("right");
    const rowCount = Math.min(leftRows.length, rightRows.length);
    for (const row of [...leftRows, ...rightRows]) {
      row.style.height = "";
    }
    for (let index = 0; index < rowCount; index += 1) {
      const leftRow = leftRows[index];
      const rightRow = rightRows[index];
      const syncedHeight = Math.ceil(
        Math.max(
          leftRow.getBoundingClientRect().height,
          rightRow.getBoundingClientRect().height
        )
      );
      if (syncedHeight <= 0) {
        continue;
      }
      leftRow.style.height = `${syncedHeight}px`;
      rightRow.style.height = `${syncedHeight}px`;
    }
  }
  function scheduleLayoutSync({ afterRender = false } = {}) {
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
  function canEditCell(status, side) {
    if (!model) {
      return false;
    }
    const isReadonly = side === "left" ? model.leftFile.isReadonly : model.rightFile.isReadonly;
    if (isReadonly) {
      return false;
    }
    const sheetName = side === "left" ? model.activeSheet.leftName : model.activeSheet.rightName;
    if (!sheetName) {
      return false;
    }
    if (status === "added" && side === "left") {
      return false;
    }
    if (status === "removed" && side === "right") {
      return false;
    }
    return true;
  }
  function getCellModelValue(rowNumber, columnNumber, side) {
    const row = model?.page.rows.find((r) => r.rowNumber === rowNumber);
    const cell = row?.cells[columnNumber - 1];
    if (!cell) {
      return "";
    }
    return side === "left" ? cell.leftValue : cell.rightValue;
  }
  function cancelEdit() {
    editState = null;
  }
  function commitEdit(value) {
    if (!editState) {
      return;
    }
    const { rowNumber, columnNumber, side } = editState;
    editState = null;
    vscode.postMessage({ type: "editCell", side, rowNumber, columnNumber, value });
  }
  function enterEditMode(cellEl, rowNumber, columnNumber, side, currentValue) {
    cancelEdit();
    editState = { rowNumber, columnNumber, side };
    const content = cellEl.querySelector(".grid__cell-content");
    if (!content) {
      return;
    }
    content.innerHTML = "";
    const input = document.createElement("input");
    input.type = "text";
    input.className = "grid__cell-input";
    input.value = currentValue;
    content.appendChild(input);
    cellEl.classList.add("grid__cell--editing");
    input.focus();
    input.select();
    let committed = false;
    const commit = (inputValue) => {
      if (committed) {
        return;
      }
      committed = true;
      cellEl.classList.remove("grid__cell--editing");
      commitEdit(inputValue);
    };
    const cancel = () => {
      if (committed) {
        return;
      }
      committed = true;
      editState = null;
      cellEl.classList.remove("grid__cell--editing");
      renderApp();
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === "Tab") {
        e.preventDefault();
        commit(input.value);
      } else if (e.key === "Escape") {
        e.preventDefault();
        cancel();
      }
    });
    input.addEventListener("blur", () => {
      if (!committed) {
        commit(input.value);
      }
    });
  }
  function renderTable(side) {
    if (!model || model.page.rows.length === 0) {
      return `<div class="empty-table">${escapeHtml(STRINGS.noRowsAvailable)}</div>`;
    }
    const diffColumnTones = getColumnDiffTones(model.page.rows);
    const headerColumns = model.page.columns.map((column, index) => {
      const diffTone = diffColumnTones.get(index);
      return `<th class="grid__column ${diffTone ? `grid__column--diff grid__column--${diffTone}` : ""}">
				<span class="grid__column-label">
					${diffTone ? `<span class="diff-marker ${getDiffToneClass(diffTone)}" aria-hidden="true"></span>` : ""}
					<span>${escapeHtml(column)}</span>
				</span>
			</th>`;
    }).join("");
    const bodyRows = model.page.rows.map((row) => {
      const rowClasses = [
        row.hasDiff ? "row--diff" : "",
        row.hasDiff ? `row--diff-${row.diffTone}` : "",
        row.isHighlighted ? "row--highlight" : ""
      ].filter(Boolean).join(" ");
      const cells = row.cells.map((cell, columnIndex) => {
        const value = side === "left" ? cell.leftValue : cell.rightValue;
        const formula = side === "left" ? cell.leftFormula : cell.rightFormula;
        const highlightCell = shouldHighlightCell(cell, side, row.isHighlighted);
        const cellClass = getSideCellClass(cell, side, highlightCell);
        const cellTooltip = getCellTooltip(cell.address, value, formula);
        const editable = canEditCell(cell.status, side) ? "true" : "false";
        return `<td title="${escapeHtml(cellTooltip)}" class="${cellClass}" data-role="grid-cell" data-row-number="${row.rowNumber}" data-column-number="${columnIndex + 1}" data-cell-status="${cell.status}" data-editable="${editable}" aria-selected="false">
						<div class="grid__cell-content">${renderCellValue(value, formula)}</div>
					</td>`;
      }).join("");
      return `
				<tr class="${rowClasses}" data-role="grid-row" data-row-number="${row.rowNumber}" data-row-has-diff="${row.hasDiff ? "true" : "false"}">
					<th class="grid__row-number ${row.hasDiff ? `grid__row-number--diff grid__row-number--${row.diffTone}` : ""}">
						<span class="grid__row-label">
							${row.hasDiff ? `<span class="diff-marker ${getDiffToneClass(row.diffTone)}" aria-hidden="true"></span>` : ""}
							<span>${row.rowNumber}</span>
						</span>
					</th>
					${cells}
				</tr>
			`;
    }).join("");
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
    const s = STRINGS;
    const readOnlyIcon = file.isReadonly ? `<span class="codicon codicon-lock file-card__lock" title="${escapeHtml(s.readOnly)}" aria-label="${escapeHtml(s.readOnly)}"></span>` : "";
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
					${file.detailLabel && file.detailValue ? `<span>${escapeHtml(file.detailLabel)}: ${escapeHtml(file.detailValue)}</span>` : ""}
					<span>${escapeHtml(s.modified)}: ${escapeHtml(file.modifiedTimeLabel)}</span>
				</div>
			</div>
		</div>
	`;
  }
  function renderPane(title, side) {
    return `
		<section class="pane" data-side="${side}">
			<div class="pane__header">
				<div class="pane__title">${escapeHtml(title)}</div>
				${model.activeSheet.mergedRangesChanged ? `<span class="badge badge--warn">${escapeHtml(STRINGS.mergedRangesChanged)}</span>` : ""}
			</div>
			<div class="pane__table" data-side="${side}">${renderTable(side)}</div>
		</section>
	`;
  }
  function renderToolbarButton({ action, icon, label, active = false, disabled = false, filter }) {
    return `<button class="toolbar__button ${active ? "is-active" : ""}" data-action="${action}" ${filter ? `data-filter="${filter}"` : ""} ${disabled ? "disabled" : ""}>
		<span class="codicon ${icon} toolbar__button-icon" aria-hidden="true"></span>
		<span>${escapeHtml(label)}</span>
	</button>`;
  }
  function renderToolbar() {
    const m = model;
    const s = STRINGS;
    return `
		<header class="toolbar">
			<div class="toolbar__group">
				${renderToolbarButton({ action: "set-filter", filter: "all", icon: "codicon-list-flat", label: s.all, active: m.filter === "all" })}
				${renderToolbarButton({ action: "set-filter", filter: "diffs", icon: "codicon-diff-multiple", label: s.diffs, active: m.filter === "diffs" })}
				${renderToolbarButton({ action: "set-filter", filter: "same", icon: "codicon-check-all", label: s.same, active: m.filter === "same" })}
			</div>
			<div class="toolbar__group">
				${renderToolbarButton({ action: "prev-diff", icon: "codicon-arrow-up", label: s.prevDiff, disabled: !m.canPrevDiff })}
				${renderToolbarButton({ action: "next-diff", icon: "codicon-arrow-down", label: s.nextDiff, disabled: !m.canNextDiff })}
				${renderToolbarButton({ action: "prev-page", icon: "codicon-arrow-left", label: s.prevPage, disabled: !m.canPrevPage })}
				${renderToolbarButton({ action: "next-page", icon: "codicon-arrow-right", label: s.nextPage, disabled: !m.canNextPage })}
				${renderToolbarButton({ action: "swap", icon: "codicon-arrow-swap", label: s.swap })}
				${renderToolbarButton({ action: "reload", icon: "codicon-refresh", label: s.reload })}
			</div>
		</header>
	`;
  }
  function renderTabs() {
    return model.sheets.map(
      (sheet) => `
				<button
					class="tab tab--${sheet.diffTone} ${sheet.isActive ? "is-active" : ""} ${sheet.hasDiff ? "has-diff" : ""}"
					data-action="set-sheet"
					data-sheet-key="${escapeHtml(sheet.key)}"
					title="${escapeHtml(getSheetTooltip(sheet))}"
				>
					${sheet.hasDiff ? `<span class="diff-marker ${getDiffToneClass(sheet.diffTone)} tab__marker" aria-hidden="true"></span>` : ""}
					<span class="tab__label">${escapeHtml(sheet.label)}</span>
				</button>
			`
    ).join("");
  }
  function renderStatus() {
    const m = model;
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
  function renderApp() {
    if (!model) {
      renderLoading(STRINGS.loading);
      return;
    }
    cancelEdit();
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
				${renderPane(s.left, "left")}
				${renderPane(s.right, "right")}
			</section>
			${renderStatus()}
		</div>
	`;
    attachPaneScrollSync();
    restorePaneScrollState(previousPaneScrollState);
    scheduleLayoutSync({ afterRender: true });
  }
  function syncPaneScroll(sourcePane) {
    if (isSyncingScroll) {
      return;
    }
    const panes = Array.from(document.querySelectorAll(".pane__table"));
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
  function attachPaneScrollSync() {
    const panes = Array.from(document.querySelectorAll(".pane__table"));
    for (const pane of panes) {
      pane.addEventListener(
        "scroll",
        () => {
          syncPaneScroll(pane);
        },
        { passive: true }
      );
    }
  }
  window.addEventListener("resize", () => {
    if (!model) {
      return;
    }
    scheduleLayoutSync();
  });
  window.addEventListener("message", (event) => {
    const message = event.data;
    if (message.type === "loading") {
      renderLoading(message.message);
      return;
    }
    if (message.type === "error") {
      renderError(message.message);
      return;
    }
    if (message.type === "render") {
      model = message.payload;
      renderApp();
    }
  });
  document.addEventListener("click", (event) => {
    const eventTarget = event.target instanceof Element ? event.target : null;
    if (!eventTarget) {
      return;
    }
    const cellTarget = eventTarget.closest('[data-role="grid-cell"]');
    if (cellTarget) {
      const rowNumber = Number(cellTarget.getAttribute("data-row-number"));
      const columnNumber = Number(cellTarget.getAttribute("data-column-number"));
      const cellStatus = cellTarget.getAttribute("data-cell-status");
      selectedCell = { rowNumber, columnNumber };
      pendingSelectionReason = null;
      applySelectedCell();
      if (cellStatus !== "equal" && (model?.page.highlightedDiffCell?.rowNumber !== rowNumber || model?.page.highlightedDiffCell?.columnNumber !== columnNumber)) {
        vscode.postMessage({ type: "selectCell", rowNumber, columnNumber });
      }
      return;
    }
    const target = eventTarget.closest("[data-action]");
    if (!target) {
      return;
    }
    const action = target.getAttribute("data-action");
    switch (action) {
      case "set-filter":
        vscode.postMessage({ type: "setFilter", filter: target.getAttribute("data-filter") });
        return;
      case "set-sheet":
        vscode.postMessage({ type: "setSheet", sheetKey: target.getAttribute("data-sheet-key") });
        return;
      case "prev-page":
        vscode.postMessage({ type: "prevPage" });
        return;
      case "next-page":
        vscode.postMessage({ type: "nextPage" });
        return;
      case "prev-diff":
        pendingSelectionReason = "highlighted-diff";
        vscode.postMessage({ type: "prevDiff" });
        return;
      case "next-diff":
        pendingSelectionReason = "highlighted-diff";
        vscode.postMessage({ type: "nextDiff" });
        return;
      case "swap":
        vscode.postMessage({ type: "swap" });
        return;
      case "reload":
        vscode.postMessage({ type: "reload" });
        return;
    }
  });
  document.addEventListener("dblclick", (event) => {
    const eventTarget = event.target instanceof Element ? event.target : null;
    if (!eventTarget) {
      return;
    }
    const cellTarget = eventTarget.closest('[data-role="grid-cell"]');
    if (!cellTarget) {
      return;
    }
    const cellStatus = cellTarget.getAttribute("data-cell-status");
    const editable = cellTarget.getAttribute("data-editable") === "true";
    if (!editable) {
      return;
    }
    const pane = cellTarget.closest("[data-side]");
    const side = pane?.getAttribute("data-side");
    if (!side) {
      return;
    }
    const rowNumber = Number(cellTarget.getAttribute("data-row-number"));
    const columnNumber = Number(cellTarget.getAttribute("data-column-number"));
    const currentValue = getCellModelValue(rowNumber, columnNumber, side);
    selectedCell = { rowNumber, columnNumber };
    enterEditMode(cellTarget, rowNumber, columnNumber, side, currentValue);
  });
  renderLoading(STRINGS.loading);
  vscode.postMessage({ type: "ready" });
})();
//# sourceMappingURL=panel.js.map
