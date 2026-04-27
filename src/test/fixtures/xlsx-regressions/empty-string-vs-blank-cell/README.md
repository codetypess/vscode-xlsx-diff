# empty-string-vs-blank-cell

这个案例用来防住“一个单元格是空白，另一个单元格是显式空字符串，但 diff panel 仍然报不同”的误报。

## Naming

- 目录名：`empty-string-vs-blank-cell`
- 含义：用户看到的是“空白单元格对空白单元格”，根因是“底层一个没有 cell，一个是空字符串 cell”

## Incident

- 来源：基于现有 diff 口径补充的高风险误报场景
- 首次归档时间：2026-04
- 影响资源链路：`local file`、`git HEAD`、`svn BASE`
- 关键定位：`define!F5`

## Symptom

- 用户看到的现象：两边单元格都显示为空，但 diff panel 可能仍标记该位置不同
- 当时误判结果：空白与显式空字符串被当成内容差异

## Root Cause

- 底层真实差异：一侧没有实际 cell，另一侧存在 `inlineStr` 类型且值为 `""` 的 cell
- 为什么 UI 看起来相同：两边渲染结果都是空白文本

## Fixture

- `base.xlsx`：`define!F5` 保持真正空白
- `head.xlsx`：`define!F5` 写入显式空字符串 `""`
- 关键 sheet / cell：`define!F5`

## Expected Behavior

- 应显示的差异：无
- 应忽略的差异：仅“空白”与“显式空字符串”之间的差异

## Coverage

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`

## Regenerate

- 运行 `npm run generate:test-fixtures`
- 脚本入口：`scripts/generate-test-fixtures.mts`
