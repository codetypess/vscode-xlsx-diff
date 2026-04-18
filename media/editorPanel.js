"use strict";
(() => {
  // src/constants.ts
  var DEFAULT_PAGE_SIZE = 200;

  // src/webview/editorPanel.ts
  var vscode = acquireVsCodeApi();
  var model = null;
  var selectedCell = null;
  var editState = null;
  var isSaving = false;
  var lastPendingNotification = null;
  var pendingSelectionAfterRender = null;
  var pendingEdits = /* @__PURE__ */ new Map();
  function getPendingEditKey(sheetKey, rowNumber, columnNumber) {
    return `${sheetKey}:${rowNumber}:${columnNumber}`;
  }
  var DEFAULT_STRINGS = {
    loading: "Loading XLSX editor...",
    reload: "Reload",
    prevPage: "Prev Page",
    nextPage: "Next Page",
    size: "Size",
    modified: "Modified",
    sheet: "Sheet",
    rows: "Rows",
    noRows: "No rows",
    page: "Page",
    visibleRows: "Visible rows",
    readOnly: "Read-only",
    save: "Save",
    totalSheets: "Sheets",
    totalRows: "Rows",
    nonEmptyCells: "Non-empty cells",
    mergedRanges: "Merged ranges",
    noRowsAvailable: "No rows available on this page.",
    readOnlyBadge: "Read-only"
  };
  var STRINGS = globalThis.__XLSX_EDITOR_STRINGS__ ?? DEFAULT_STRINGS;
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
  function updateLabelMarker(labelEl, hasPending, extraClass) {
    let markerEl = labelEl.querySelector(".diff-marker");
    if (!hasPending) {
      markerEl?.remove();
      return;
    }
    if (!markerEl) {
      markerEl = document.createElement("span");
      markerEl.setAttribute("aria-hidden", "true");
      labelEl.insertBefore(markerEl, labelEl.firstChild);
    }
    markerEl.className = `diff-marker diff-marker--pending${extraClass ? ` ${extraClass}` : ""}`;
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
  function getCellView(rowNumber, columnNumber) {
    const row = model?.page.rows.find((item) => item.rowNumber === rowNumber);
    return row?.cells[columnNumber - 1] ?? null;
  }
  function clearSelectedCells() {
    for (const cell of document.querySelectorAll(".grid__cell--selected")) {
      cell.classList.remove("grid__cell--selected");
    }
    for (const cell of document.querySelectorAll('[data-role="grid-cell"][aria-selected="true"]')) {
      cell.setAttribute("aria-selected", "false");
    }
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
  function restorePaneScrollState(scrollState) {
    if (!scrollState) {
      return;
    }
    const pane = document.querySelector(".pane__table");
    if (!pane) {
      return;
    }
    pane.scrollTop = clampScrollPosition(scrollState.top, pane.scrollHeight - pane.clientHeight);
    pane.scrollLeft = clampScrollPosition(scrollState.left, pane.scrollWidth - pane.clientWidth);
  }
  function getStickyPaneInsets(pane) {
    const headerRow = pane.querySelector("thead tr");
    const firstColumn = pane.querySelector("thead th:first-child");
    return {
      top: headerRow?.getBoundingClientRect().height ?? 0,
      left: firstColumn?.getBoundingClientRect().width ?? 0
    };
  }
  function revealSelectedCells(elements) {
    const pane = document.querySelector(".pane__table");
    if (!pane || elements.length === 0) {
      return;
    }
    const target = elements[0];
    const paneRect = pane.getBoundingClientRect();
    const elementRect = target.getBoundingClientRect();
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
    pane.scrollTop = clampScrollPosition(top, pane.scrollHeight - pane.clientHeight);
    pane.scrollLeft = clampScrollPosition(left, pane.scrollWidth - pane.clientWidth);
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
  function syncSelectedCellToHost() {
    if (!model || !selectedCell) {
      return;
    }
    vscode.postMessage({
      type: "selectCell",
      rowNumber: selectedCell.rowNumber,
      columnNumber: selectedCell.columnNumber
    });
  }
  function setSelectedCellLocal(nextCell, { reveal = false, syncHost = true } = {}) {
    selectedCell = nextCell;
    applySelectedCell({ reveal });
    if (syncHost) {
      syncSelectedCellToHost();
    }
  }
  function syncSelectedCellAfterRender({ reveal = false } = {}) {
    if (pendingSelectionAfterRender && getCellView(
      pendingSelectionAfterRender.rowNumber,
      pendingSelectionAfterRender.columnNumber
    )) {
      selectedCell = {
        rowNumber: pendingSelectionAfterRender.rowNumber,
        columnNumber: pendingSelectionAfterRender.columnNumber
      };
      pendingSelectionAfterRender = null;
      applySelectedCell({ reveal: true });
      syncSelectedCellToHost();
      return;
    }
    pendingSelectionAfterRender = null;
    if (selectedCell && getCellView(selectedCell.rowNumber, selectedCell.columnNumber)) {
      applySelectedCell({ reveal });
      syncSelectedCellToHost();
      return;
    }
    selectedCell = model?.selection ? {
      rowNumber: model.selection.rowNumber,
      columnNumber: model.selection.columnNumber
    } : model?.page.rows[0] ? { rowNumber: model.page.rows[0].rowNumber, columnNumber: 1 } : null;
    applySelectedCell({ reveal });
    if (selectedCell) {
      syncSelectedCellToHost();
    }
  }
  function canEditCell() {
    return Boolean(model?.canEdit);
  }
  function getCellModelValue(rowNumber, columnNumber) {
    return getCellView(rowNumber, columnNumber)?.value ?? "";
  }
  function getCellFormula(rowNumber, columnNumber) {
    return getCellView(rowNumber, columnNumber)?.formula ?? null;
  }
  function notifyPendingEditState() {
    const hasPendingEdits = pendingEdits.size > 0;
    if (lastPendingNotification === hasPendingEdits) {
      return;
    }
    lastPendingNotification = hasPendingEdits;
    vscode.postMessage({ type: "pendingEditStateChanged", hasPendingEdits });
  }
  function updateSaveButtonState() {
    notifyPendingEditState();
    const saveBtn = document.querySelector('[data-action="save-edits"]');
    if (!saveBtn) {
      return;
    }
    const hasPendingEdits = pendingEdits.size > 0;
    saveBtn.disabled = !model?.canEdit || !hasPendingEdits || isSaving;
    saveBtn.classList.toggle("is-dirty", hasPendingEdits);
    if (isSaving) {
      saveBtn.setAttribute("aria-busy", "true");
    } else {
      saveBtn.removeAttribute("aria-busy");
    }
  }
  function syncCellDisplay(cellEl, sheetKey, rowNumber, columnNumber) {
    const content = cellEl.querySelector(".grid__cell-content");
    if (!content) {
      return;
    }
    const pendingKey = getPendingEditKey(sheetKey, rowNumber, columnNumber);
    const pendingEdit = pendingEdits.get(pendingKey);
    const value = pendingEdit ? pendingEdit.value : getCellModelValue(rowNumber, columnNumber);
    const formula = pendingEdit ? null : getCellFormula(rowNumber, columnNumber);
    content.innerHTML = renderCellValue(value, formula);
    cellEl.classList.toggle("grid__cell--pending", Boolean(pendingEdit));
  }
  function finishEdit({ mode, clearSelection = false }) {
    const session = editState;
    if (!session) {
      return;
    }
    editState = null;
    const { sheetKey, rowNumber, columnNumber, cellEl, input } = session;
    cellEl.classList.remove("grid__cell--editing");
    if (mode === "commit") {
      commitEdit(sheetKey, rowNumber, columnNumber, input.value);
      if (clearSelection) {
        selectedCell = null;
        clearSelectedCells();
      }
      return;
    }
    syncCellDisplay(cellEl, sheetKey, rowNumber, columnNumber);
  }
  function clearSelectedCellValue() {
    if (!model || !selectedCell || !canEditCell()) {
      return;
    }
    commitEdit(
      model.activeSheet.key,
      selectedCell.rowNumber,
      selectedCell.columnNumber,
      ""
    );
  }
  function isClearSelectedCellKey(event) {
    if (event.altKey || event.ctrlKey || event.metaKey) {
      return false;
    }
    return event.key === "Backspace" || event.key === "Delete" || event.code === "Backspace" || event.code === "Delete";
  }
  function applyPendingEditStyles() {
    const activeSheetKey = model?.activeSheet.key;
    if (!activeSheetKey) {
      return;
    }
    const pendingSheetKeys = /* @__PURE__ */ new Set();
    const pendingRows = /* @__PURE__ */ new Set();
    const pendingColumns = /* @__PURE__ */ new Set();
    for (const pendingEdit of pendingEdits.values()) {
      pendingSheetKeys.add(pendingEdit.sheetKey);
      if (pendingEdit.sheetKey !== activeSheetKey) {
        continue;
      }
      pendingRows.add(pendingEdit.rowNumber);
      pendingColumns.add(pendingEdit.columnNumber);
    }
    for (const cellEl of document.querySelectorAll('[data-role="grid-cell"]')) {
      if (cellEl.classList.contains("grid__cell--editing")) {
        continue;
      }
      const rowNumber = Number(cellEl.getAttribute("data-row-number"));
      const columnNumber = Number(cellEl.getAttribute("data-column-number"));
      syncCellDisplay(cellEl, activeSheetKey, rowNumber, columnNumber);
    }
    for (const rowHeader of document.querySelectorAll('th[data-role="grid-row-header"]')) {
      const rowNumber = Number(rowHeader.getAttribute("data-row-number"));
      const hasPending = pendingRows.has(rowNumber);
      rowHeader.classList.toggle("grid__row-number--pending", hasPending);
      const labelEl = rowHeader.querySelector(".grid__row-label");
      if (labelEl) {
        updateLabelMarker(labelEl, hasPending);
      }
    }
    for (const colHeader of document.querySelectorAll("thead th[data-column-number]")) {
      const columnNumber = Number(colHeader.getAttribute("data-column-number"));
      const hasPending = pendingColumns.has(columnNumber);
      colHeader.classList.toggle("grid__column--diff", hasPending);
      colHeader.classList.toggle("grid__column--pending", hasPending);
      const labelEl = colHeader.querySelector(".grid__column-label");
      if (labelEl) {
        updateLabelMarker(labelEl, hasPending);
      }
    }
    for (const tabEl of document.querySelectorAll('[data-action="set-sheet"]')) {
      const sheetKey = tabEl.getAttribute("data-sheet-key");
      const hasPending = sheetKey ? pendingSheetKeys.has(sheetKey) : false;
      updateLabelMarker(tabEl, hasPending, "tab__marker");
    }
  }
  function commitEdit(sheetKey, rowNumber, columnNumber, value) {
    const key = getPendingEditKey(sheetKey, rowNumber, columnNumber);
    const modelValue = getCellModelValue(rowNumber, columnNumber);
    if (value === modelValue) {
      pendingEdits.delete(key);
    } else {
      pendingEdits.set(key, { sheetKey, rowNumber, columnNumber, value });
    }
    applyPendingEditStyles();
    updateSaveButtonState();
  }
  function triggerSave() {
    if (!model || pendingEdits.size === 0 || isSaving) {
      return;
    }
    finishEdit({ mode: "commit", clearSelection: true });
    isSaving = true;
    updateSaveButtonState();
    vscode.postMessage({
      type: "saveEdits",
      edits: Array.from(pendingEdits.values())
    });
  }
  function normalizePastedRows(text) {
    const lines = text.replaceAll("\r", "").split("\n");
    if (lines.length > 0 && lines[lines.length - 1] === "") {
      lines.pop();
    }
    return lines.map((line) => line.split("	"));
  }
  function applyPastedGrid(grid) {
    if (!model || !selectedCell || grid.length === 0) {
      return;
    }
    const maxRow = model.activeSheet.rowCount;
    const maxColumn = model.activeSheet.columnCount;
    for (let rowOffset = 0; rowOffset < grid.length; rowOffset += 1) {
      const targetRow = selectedCell.rowNumber + rowOffset;
      if (targetRow > maxRow) {
        break;
      }
      const values = grid[rowOffset] ?? [];
      for (let columnOffset = 0; columnOffset < values.length; columnOffset += 1) {
        const targetColumn = selectedCell.columnNumber + columnOffset;
        if (targetColumn > maxColumn) {
          break;
        }
        commitEdit(model.activeSheet.key, targetRow, targetColumn, values[columnOffset] ?? "");
      }
    }
    applySelectedCell({ reveal: true });
  }
  function enterEditMode(cellEl, rowNumber, columnNumber, currentValue) {
    finishEdit({ mode: "commit" });
    const capturedSheetKey = model?.activeSheet.key;
    const content = cellEl.querySelector(".grid__cell-content");
    if (!content || !capturedSheetKey) {
      return;
    }
    content.innerHTML = "";
    const input = document.createElement("input");
    input.type = "text";
    input.className = "grid__cell-input";
    input.value = currentValue;
    content.appendChild(input);
    cellEl.classList.add("grid__cell--editing");
    editState = { sheetKey: capturedSheetKey, rowNumber, columnNumber, cellEl, input };
    input.focus();
    input.select();
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === "Tab") {
        event.preventDefault();
        finishEdit({ mode: "commit", clearSelection: true });
      } else if (event.key === "Escape") {
        event.preventDefault();
        finishEdit({ mode: "cancel" });
      }
    });
    input.addEventListener("blur", () => {
      if (editState?.input === input) {
        finishEdit({ mode: "commit", clearSelection: true });
      }
    });
  }
  function getSelectionBounds() {
    if (!model || model.page.rows.length === 0) {
      return null;
    }
    return {
      minRow: model.page.rows[0].rowNumber,
      maxRow: model.page.rows[model.page.rows.length - 1].rowNumber,
      minColumn: 1,
      maxColumn: model.page.columns.length
    };
  }
  function ensureSelection() {
    if (selectedCell) {
      return selectedCell;
    }
    if (model?.selection) {
      selectedCell = {
        rowNumber: model.selection.rowNumber,
        columnNumber: model.selection.columnNumber
      };
      return selectedCell;
    }
    if (model?.page.rows[0]) {
      selectedCell = {
        rowNumber: model.page.rows[0].rowNumber,
        columnNumber: 1
      };
      return selectedCell;
    }
    return null;
  }
  function moveSelection(rowDelta, columnDelta) {
    const selection = ensureSelection();
    const bounds = getSelectionBounds();
    if (!selection || !bounds) {
      return;
    }
    const nextRow = Math.max(bounds.minRow, Math.min(bounds.maxRow, selection.rowNumber + rowDelta));
    const nextColumn = Math.max(
      bounds.minColumn,
      Math.min(bounds.maxColumn, selection.columnNumber + columnDelta)
    );
    setSelectedCellLocal({ rowNumber: nextRow, columnNumber: nextColumn }, { reveal: true });
  }
  function moveSelectionByPage(direction) {
    const selection = ensureSelection();
    if (!model || !selection) {
      return;
    }
    const canMove = direction < 0 ? model.canPrevPage : model.canNextPage;
    if (!canMove) {
      return;
    }
    pendingSelectionAfterRender = {
      rowNumber: Math.max(
        1,
        Math.min(model.activeSheet.rowCount, selection.rowNumber + direction * DEFAULT_PAGE_SIZE)
      ),
      columnNumber: selection.columnNumber,
      reveal: true
    };
    vscode.postMessage({ type: direction < 0 ? "prevPage" : "nextPage" });
  }
  function renderTable() {
    if (!model || model.page.rows.length === 0) {
      return `<div class="empty-table">${escapeHtml(STRINGS.noRowsAvailable)}</div>`;
    }
    const headerColumns = model.page.columns.map(
      (column, index) => `<th class="grid__column" data-column-number="${index + 1}">
				<span class="grid__column-label"><span>${escapeHtml(column)}</span></span>
			</th>`
    ).join("");
    const bodyRows = model.page.rows.map((row) => {
      const cells = row.cells.map((cell, columnIndex) => {
        const cellClass = ["grid__cell", cell.isSelected ? "grid__cell--selected" : ""].filter(Boolean).join(" ");
        const editable = canEditCell() ? "true" : "false";
        return `<td title="${escapeHtml(getCellTooltip(cell.address, cell.value, cell.formula))}" class="${cellClass}" data-role="grid-cell" data-row-number="${row.rowNumber}" data-column-number="${columnIndex + 1}" data-editable="${editable}" aria-selected="${cell.isSelected ? "true" : "false"}">
						<div class="grid__cell-content">${renderCellValue(cell.value, cell.formula)}</div>
					</td>`;
      }).join("");
      return `
				<tr data-role="grid-row" data-row-number="${row.rowNumber}">
					<th class="grid__row-number" data-role="grid-row-header" data-row-number="${row.rowNumber}">
						<span class="grid__row-label"><span>${row.rowNumber}</span></span>
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
    const readOnlyIcon = file.isReadonly ? `<span class="codicon codicon-lock file-card__lock" title="${escapeHtml(STRINGS.readOnly)}" aria-label="${escapeHtml(STRINGS.readOnly)}"></span>` : "";
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
					${file.detailLabel && file.detailValue ? `<span>${escapeHtml(file.detailLabel)}: ${escapeHtml(file.detailValue)}</span>` : ""}
					<span>${escapeHtml(STRINGS.modified)}: ${escapeHtml(file.modifiedTimeLabel)}</span>
				</div>
			</div>
		</div>
	`;
  }
  function renderToolbarButton({
    action,
    icon,
    label,
    disabled = false
  }) {
    return `<button class="toolbar__button" data-action="${action}" ${disabled ? "disabled" : ""}>
		<span class="codicon ${icon} toolbar__button-icon" aria-hidden="true"></span>
		<span>${escapeHtml(label)}</span>
	</button>`;
  }
  function renderToolbar() {
    const currentModel = model;
    return `
		<header class="toolbar toolbar--editor">
			<div class="toolbar__group">
				${renderToolbarButton({ action: "reload", icon: "codicon-refresh", label: STRINGS.reload })}
				${renderToolbarButton({ action: "save-edits", icon: "codicon-save", label: STRINGS.save, disabled: !currentModel.canSave })}
			</div>
		</header>
	`;
  }
  function renderPane() {
    return `
		<section class="pane pane--single">
			<div class="pane__header">
				<div class="pane__title">${escapeHtml(model.activeSheet.label)}</div>
				${model.activeSheet.hasMergedRanges ? `<span class="badge badge--warn">${escapeHtml(STRINGS.mergedRanges)}: ${model.activeSheet.mergedRangeCount}</span>` : ""}
			</div>
			<div class="pane__table">${renderTable()}</div>
		</section>
	`;
  }
  function renderTabs() {
    return model.sheets.map(
      (sheet) => `
				<button
					class="tab ${sheet.isActive ? "is-active" : ""}"
					data-action="set-sheet"
					data-sheet-key="${escapeHtml(sheet.key)}"
					title="${escapeHtml(sheet.label)}"
				>
					<span class="tab__label">${escapeHtml(sheet.label)}</span>
				</button>
			`
    ).join("");
  }
  function renderStatus() {
    const currentModel = model;
    const rowRangeLabel = currentModel.page.visibleRowCount === 0 ? STRINGS.noRows : currentModel.page.rangeLabel;
    return `
		<footer class="footer">
			<div class="tabs">${renderTabs()}</div>
			<div class="status">
				<span><strong>${escapeHtml(STRINGS.sheet)}:</strong> ${escapeHtml(currentModel.activeSheet.label)}</span>
				<span><strong>${escapeHtml(STRINGS.rows)}:</strong> ${escapeHtml(rowRangeLabel)}</span>
				<span><strong>${escapeHtml(STRINGS.page)}:</strong> ${currentModel.page.currentPage} / ${currentModel.page.totalPages}</span>
				<span><strong>${escapeHtml(STRINGS.visibleRows)}:</strong> ${currentModel.page.visibleRowCount}</span>
			</div>
		</footer>
	`;
  }
  function renderApp({ revealSelection = false } = {}) {
    if (!model) {
      renderLoading(STRINGS.loading);
      return;
    }
    finishEdit({ mode: "commit" });
    const previousScrollState = getPaneScrollState();
    document.body.innerHTML = `
		<div id="app" class="app app--editor">
			${renderToolbar()}
			<section class="files files--single">
				${renderFileCard(model.file)}
			</section>
			<section class="panes panes--single">
				${renderPane()}
			</section>
			${renderStatus()}
		</div>
	`;
    restorePaneScrollState(previousScrollState);
    applyPendingEditStyles();
    updateSaveButtonState();
    syncSelectedCellAfterRender({ reveal: revealSelection });
  }
  window.addEventListener("message", (event) => {
    const message = event.data;
    if (message.type === "loading") {
      renderLoading(message.message);
      return;
    }
    if (message.type === "error") {
      isSaving = false;
      updateSaveButtonState();
      renderError(message.message);
      return;
    }
    if (message.type === "render") {
      model = message.payload;
      isSaving = false;
      if (message.clearPendingEdits) {
        pendingEdits.clear();
      }
      renderApp({ revealSelection: !message.silent });
    }
  });
  document.addEventListener("click", (event) => {
    const eventTarget = event.target instanceof Element ? event.target : null;
    if (!eventTarget) {
      return;
    }
    const cellTarget = eventTarget.closest('[data-role="grid-cell"]');
    if (cellTarget) {
      if (editState) {
        finishEdit({ mode: "commit" });
      }
      selectedCell = {
        rowNumber: Number(cellTarget.getAttribute("data-row-number")),
        columnNumber: Number(cellTarget.getAttribute("data-column-number"))
      };
      setSelectedCellLocal(selectedCell);
      return;
    }
    const target = eventTarget.closest("[data-action]");
    if (!target) {
      return;
    }
    finishEdit({ mode: "commit" });
    const action = target.getAttribute("data-action");
    switch (action) {
      case "set-sheet":
        vscode.postMessage({ type: "setSheet", sheetKey: target.getAttribute("data-sheet-key") });
        return;
      case "reload":
        vscode.postMessage({ type: "reload" });
        return;
      case "save-edits":
        triggerSave();
        return;
    }
  });
  document.addEventListener("keydown", (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      triggerSave();
      return;
    }
    if (editState) {
      return;
    }
    if (isClearSelectedCellKey(event) && selectedCell) {
      event.preventDefault();
      clearSelectedCellValue();
      return;
    }
    if (event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
      return;
    }
    switch (event.key) {
      case "ArrowUp":
        event.preventDefault();
        moveSelection(-1, 0);
        return;
      case "ArrowDown":
        event.preventDefault();
        moveSelection(1, 0);
        return;
      case "ArrowLeft":
        event.preventDefault();
        moveSelection(0, -1);
        return;
      case "ArrowRight":
        event.preventDefault();
        moveSelection(0, 1);
        return;
      case "Tab":
        event.preventDefault();
        moveSelection(0, 1);
        return;
      case "Enter":
        event.preventDefault();
        moveSelection(1, 0);
        return;
      case "PageUp":
        event.preventDefault();
        moveSelectionByPage(-1);
        return;
      case "PageDown":
        event.preventDefault();
        moveSelectionByPage(1);
        return;
    }
  });
  document.addEventListener("paste", (event) => {
    if (editState || !model || !selectedCell || !canEditCell()) {
      return;
    }
    const text = event.clipboardData?.getData("text/plain");
    if (!text) {
      return;
    }
    event.preventDefault();
    applyPastedGrid(normalizePastedRows(text));
  });
  document.addEventListener("dblclick", (event) => {
    const eventTarget = event.target instanceof Element ? event.target : null;
    if (!eventTarget || !model || !canEditCell()) {
      return;
    }
    const cellTarget = eventTarget.closest('[data-role="grid-cell"]');
    if (!cellTarget || cellTarget.getAttribute("data-editable") !== "true") {
      return;
    }
    const rowNumber = Number(cellTarget.getAttribute("data-row-number"));
    const columnNumber = Number(cellTarget.getAttribute("data-column-number"));
    const pendingKey = getPendingEditKey(model.activeSheet.key, rowNumber, columnNumber);
    const pendingEdit = pendingEdits.get(pendingKey);
    const currentValue = pendingEdit ? pendingEdit.value : getCellModelValue(rowNumber, columnNumber);
    setSelectedCellLocal({ rowNumber, columnNumber }, { syncHost: true });
    enterEditMode(cellTarget, rowNumber, columnNumber, currentValue);
  });
  renderLoading(STRINGS.loading);
  vscode.postMessage({ type: "ready" });
})();
//# sourceMappingURL=editorPanel.js.map
