# XLSX Regression Fixtures

这里收录真实 `.xlsx` 误报/回归样本，优先服务于 `loadWorkbookSnapshot`、`buildWorkbookDiff` 和资源侧兼容测试。

## Naming Convention

- 每个案例单独建目录，目录名统一使用 `现象-根因`
- 目录名只用小写英文、数字和中划线
- 优先写“用户看到的现象”，再写“我们确认的根因”
- 如果根因还没完全确认，先使用 `现象-suspected-cause`

推荐模式：

- `newline-only-cell-diff`
- `trimmed-whitespace-cell-diff`
- `same-display-different-type`
- `formula-cache-mismatch`

## Case Contract

- 每个案例至少包含 `base.xlsx`、`head.xlsx` 和 `README.md`
- 所有 `.xlsx` 统一通过 `fastxlsx` 生成，避免手改压缩包内容
- 说明文档优先按 [CASE_TEMPLATE.md](./CASE_TEMPLATE.md) 填写
- 如果案例覆盖特定资源链路，要在文档里写清 `local file`、`git`、`svn` 的覆盖范围

## Add A Case

1. 新建一个符合命名规范的目录。
2. 通过 `npm run generate:test-fixtures` 或补充生成脚本产出 `base.xlsx` / `head.xlsx`。
3. 按 [CASE_TEMPLATE.md](./CASE_TEMPLATE.md) 填写复现说明。
4. 为该案例补最小回归测试，至少覆盖本地文件链路。
5. 如果案例与 SCM 资源相关，再补 `git` / `svn` 链路用例。

## Current Cases

- `newline-only-cell-diff`
  说明：`define!F5` 仅换行风格不同，UI 看起来一致，不应继续报 diff
- `empty-string-vs-blank-cell`
  说明：`define!F5` 一侧为空白，另一侧为显式空字符串，两边看起来都为空，不应继续报 diff
- `style-only-background-color`
  说明：`define!F5` 值相同，只改了背景色样式，不应继续报 diff
- `freeze-pane-only-view-change`
  说明：`define!F5` 值相同，但冻结窗格发生变化，应作为结构性差异显示
- `sheet-visibility-only-structure-change`
  说明：`define!F5` 值相同，但工作表显示状态发生变化，应作为结构性差异显示
- `sheet-order-only-structure-change`
  说明：`define!F5` 值相同，但工作表顺序发生变化，应作为结构性差异显示
- `defined-name-only-structure-change`
  说明：`define!F5` 值相同，但 workbook defined name 发生变化，应作为结构性差异显示

生成命令：

```bash
npm run generate:test-fixtures
```
