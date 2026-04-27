# newline-only-cell-diff

首个真实工作簿回归案例，对应线上遇到的“看起来一样，但 diff panel 在 `define!F5` 误报不同”问题。

## Naming

- 目录名：`newline-only-cell-diff`
- 含义：用户看到的是“单元格被判定不同”，根因是“只差换行风格”

## Incident

- 来源：真实线上误报案例
- 首次发现时间：2026-04
- 首次归档时间：2026-04
- 影响资源链路：`local file`、`git HEAD`、`svn BASE`
- 关键定位：`define!F5`

## Symptom

- 用户看到的现象：diff panel 在 `define!F5` 标记单元格不同
- 当时误判结果：分页/F5 被判定为差异，但肉眼看两边内容一致

## Root Cause

- 底层真实差异：同一段文本分别使用 `LF` 与 `CRLF` 换行
- 为什么 UI 看起来相同：渲染后的多行文本内容一致，只是底层换行符不同

## Fixture

- `base.xlsx`：保留 `LF` 文本版本
- `head.xlsx`：保留 `CRLF` 文本版本
- 关键 sheet / cell：`define!F5`

## Expected Behavior

- 应显示的差异：无
- 应忽略的差异：仅换行风格不同的文本差异

## Coverage

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`

## Regenerate

- 运行 `npm run generate:test-fixtures`
- 脚本入口：`scripts/generate-test-fixtures.mts`
