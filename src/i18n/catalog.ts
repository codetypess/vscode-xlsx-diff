import type { EditorPanelStrings } from "../webview/editor-panel-types";

export type I18nLanguage = "en" | "zh-cn";

export interface DiffPanelStrings {
    loading: string;
    reload: string;
    swap: string;
    all: string;
    diffs: string;
    same: string;
    diffRows: string;
    sameRows: string;
    prevDiff: string;
    nextDiff: string;
    sheets: string;
    diffCells: string;
    rows: string;
    filter: string;
    visibleRows: string;
    currentDiff: string;
    selected: string;
    save: string;
    none: string;
    modified: string;
    size: string;
    readOnly: string;
    noSheet: string;
    noRows: string;
}

export interface RuntimeMessages {
    commands: {
        workbookFilterLabel: string;
        openEditorSelectLocalWorkbook: string;
        compareTwoFilesSelectLeftWorkbook: string;
        compareTwoFilesSelectRightWorkbook: string;
        compareActiveWithOpenWorkbookFirst: string;
        compareActiveWithSelectTargetWorkbook: string;
    };
    workbook: {
        schemeResource: string;
        savingWorkbook: string;
        savingChanges: string;
        untitledSheet: string;
        newSheetBaseName: string;
    };
    scm: {
        sourceLabel: string;
        commitLabel: string;
        committerLabel: string;
        authorLabel: string;
        indexLabel: string;
        stageLabel: string;
        indexBaseLabel: string;
        gitRefLabel: string;
        svnRefLabel: string;
        emptyWorkbookLabel: string;
    };
    editorPanel: EditorPanelStrings;
    diffPanel: DiffPanelStrings;
}

export type I18nParams = Record<string, string | number>;

export function formatI18nMessage(
    template: string,
    params: I18nParams | undefined = undefined
): string {
    if (!params) {
        return template;
    }

    return template.replace(/\{(\w+)\}/g, (match, key: string) =>
        Object.prototype.hasOwnProperty.call(params, key) ? String(params[key]) : match
    );
}

export const RUNTIME_MESSAGES = {
    en: {
        commands: {
            workbookFilterLabel: "Excel Workbooks",
            openEditorSelectLocalWorkbook: "Select or open a local .xlsx file first.",
            compareTwoFilesSelectLeftWorkbook: "Select the left XLSX workbook",
            compareTwoFilesSelectRightWorkbook: "Select the right XLSX workbook",
            compareActiveWithOpenWorkbookFirst:
                'Open an .xlsx file first, or run "Compare Two XLSX Files" from the Command Palette.',
            compareActiveWithSelectTargetWorkbook:
                "Select the XLSX workbook to compare against",
        },
        workbook: {
            schemeResource: "{scheme} resource",
            savingWorkbook: "Saving {workbookName}...",
            savingChanges: "Saving workbook changes...",
            untitledSheet: "Untitled Sheet",
            newSheetBaseName: "Sheet",
        },
        scm: {
            sourceLabel: "Source",
            commitLabel: "Commit",
            committerLabel: "Committer",
            authorLabel: "Author",
            indexLabel: "Index",
            stageLabel: "Stage {stage}",
            indexBaseLabel: "Index · base {commit}",
            gitRefLabel: "Git ref: {ref}",
            svnRefLabel: "SVN ref: {ref}",
            emptyWorkbookLabel: "Empty workbook",
        },
        editorPanel: {
            search: "Search",
            searchFind: "Find",
            searchReplace: "Replace",
            searchReplaceComingSoon: "Coming soon",
            searchScopeSheet: "Current sheet",
            searchScopeSelection: "Selected range",
            searchScopeSelectionDisabled: "Select multiple cells to enable",
            searchClose: "Close",
            loading: "Loading XLSX editor...",
            reload: "Reload",
            undo: "Undo",
            redo: "Redo",
            searchPlaceholder: "Search values or formulas",
            findPrev: "Prev Match",
            findNext: "Next Match",
            gotoPlaceholder: "A1 or Sheet1!B2",
            goto: "Go",
            cancelInput: "Cancel input",
            confirmInput: "Apply input",
            save: "Save",
            lockView: "Lock View",
            unlockView: "Unlock View",
            addSheet: "Add Sheet",
            deleteSheet: "Delete Sheet",
            renameSheet: "Rename Sheet",
            insertRowAbove: "Insert Row Above",
            deleteRow: "Delete Row",
            insertColumnLeft: "Insert Column Left",
            deleteColumn: "Delete Column",
            renameSheetPrompt: "Enter a new sheet name",
            renameSheetTitle: "Rename Sheet",
            sheetNameEmpty: "Sheet name cannot be empty.",
            sheetNameDuplicate: "A sheet with this name already exists.",
            sheetNameTooLong: "Sheet names must be 31 characters or fewer.",
            sheetNameInvalidChars: "Sheet names cannot contain \\ / ? * [ ] or :.",
            selectedCell: "Selected cell",
            multipleCellsSelected: "Multiple cells selected",
            noCellSelected: "None",
            noRowsAvailable: "No rows available in this view.",
            localChangesBlockedReload:
                "The workbook changed on disk. Save or discard your pending edits before reloading.",
            displayLanguageRefreshBlocked:
                "Pending edits are open. Display language changes will apply after you save or reload the editor.",
            noSearchMatches: "No matching cells were found.",
            invalidCellReference:
                "Unable to locate that cell. Use A1 or Sheet1!B2 and stay within the workbook range.",
            invalidSearchPattern: "The search pattern is invalid.",
            searchRegex: "Use Regular Expression",
            searchMatchCase: "Match Case",
            searchWholeWord: "Match Whole Word",
        },
        diffPanel: {
            loading: "Loading XLSX diff...",
            reload: "Reload",
            swap: "Swap",
            all: "All",
            diffs: "Diffs",
            same: "Same",
            diffRows: "Diff Rows",
            sameRows: "Same Rows",
            prevDiff: "Prev Diff",
            nextDiff: "Next Diff",
            sheets: "Sheets",
            diffCells: "Diff Cells",
            rows: "Rows",
            filter: "Filter",
            visibleRows: "Visible Rows",
            currentDiff: "Current Diff",
            selected: "Selected",
            save: "Save",
            none: "-",
            modified: "Modified",
            size: "Size",
            readOnly: "Read-only",
            noSheet: "No sheet is available.",
            noRows: "No rows are available in this sheet.",
        },
    },
    "zh-cn": {
        commands: {
            workbookFilterLabel: "Excel 工作簿",
            openEditorSelectLocalWorkbook: "请先选择或打开一个本地 .xlsx 文件。",
            compareTwoFilesSelectLeftWorkbook: "选择左侧 XLSX 工作簿",
            compareTwoFilesSelectRightWorkbook: "选择右侧 XLSX 工作簿",
            compareActiveWithOpenWorkbookFirst:
                "请先打开一个 .xlsx 文件，或从命令面板运行“比较两个 XLSX 文件”。",
            compareActiveWithSelectTargetWorkbook: "选择要与当前文件比较的 XLSX 工作簿",
        },
        workbook: {
            schemeResource: "{scheme} 资源",
            savingWorkbook: "正在保存 {workbookName}...",
            savingChanges: "正在保存工作簿修改...",
            untitledSheet: "未命名工作表",
            newSheetBaseName: "工作表",
        },
        scm: {
            sourceLabel: "来源",
            commitLabel: "提交",
            committerLabel: "提交者",
            authorLabel: "提交者",
            indexLabel: "暂存区",
            stageLabel: "阶段 {stage}",
            indexBaseLabel: "暂存区 · 基线 {commit}",
            gitRefLabel: "Git 引用: {ref}",
            svnRefLabel: "SVN 引用: {ref}",
            emptyWorkbookLabel: "空工作簿",
        },
        editorPanel: {
            search: "搜索",
            searchFind: "查找",
            searchReplace: "替换",
            searchReplaceComingSoon: "即将推出",
            searchScopeSheet: "当前工作表",
            searchScopeSelection: "选定区域",
            searchScopeSelectionDisabled: "请先选中多个单元格后再使用",
            searchClose: "关闭",
            loading: "正在加载 XLSX 编辑器...",
            reload: "刷新",
            undo: "撤销",
            redo: "重做",
            searchPlaceholder: "搜索值或公式",
            findPrev: "上一个",
            findNext: "下一个",
            gotoPlaceholder: "A1 或 Sheet1!B2",
            goto: "定位",
            cancelInput: "取消输入",
            confirmInput: "确认输入",
            save: "保存",
            lockView: "锁定视图",
            unlockView: "取消锁定视图",
            addSheet: "添加工作表",
            deleteSheet: "删除工作表",
            renameSheet: "重命名工作表",
            insertRowAbove: "在上方插入行",
            deleteRow: "删除行",
            insertColumnLeft: "在左侧插入列",
            deleteColumn: "删除列",
            renameSheetPrompt: "输入新的工作表名称",
            renameSheetTitle: "重命名工作表",
            sheetNameEmpty: "工作表名称不能为空。",
            sheetNameDuplicate: "工作表名称已存在。",
            sheetNameTooLong: "工作表名称不能超过 31 个字符。",
            sheetNameInvalidChars: "工作表名称不能包含 \\ / ? * [ ] : 等字符。",
            selectedCell: "当前单元格",
            multipleCellsSelected: "已选择多个单元格",
            noCellSelected: "未选择",
            noRowsAvailable: "当前视图没有可显示的行。",
            localChangesBlockedReload:
                "工作簿文件已在磁盘上变化。请先保存或放弃当前未保存修改，再刷新。",
            displayLanguageRefreshBlocked:
                "当前有未保存修改，语言变更将在保存或手动刷新后生效。",
            noSearchMatches: "没有找到匹配的单元格。",
            invalidCellReference:
                "无法定位该单元格，请使用 A1 或 Sheet1!B2 格式，并确保目标在当前工作簿范围内。",
            invalidSearchPattern: "搜索表达式无效。",
            searchRegex: "使用正则表达式",
            searchMatchCase: "区分大小写",
            searchWholeWord: "匹配整个单词",
        },
        diffPanel: {
            loading: "正在加载 XLSX 对比...",
            reload: "刷新",
            swap: "交换",
            all: "全部",
            diffs: "差异",
            same: "相同",
            diffRows: "差异行",
            sameRows: "相同行",
            prevDiff: "上一处差异",
            nextDiff: "下一处差异",
            sheets: "工作表",
            diffCells: "差异单元格",
            rows: "行",
            filter: "筛选",
            visibleRows: "可见行",
            currentDiff: "当前差异",
            selected: "选中",
            save: "保存",
            none: "-",
            modified: "修改时间",
            size: "大小",
            readOnly: "只读",
            noSheet: "没有可显示的工作表。",
            noRows: "当前工作表没有可显示的行。",
        },
    },
} satisfies Record<I18nLanguage, RuntimeMessages>;

export function getRuntimeMessagesForLanguage(language: I18nLanguage): RuntimeMessages {
    return RUNTIME_MESSAGES[language];
}
