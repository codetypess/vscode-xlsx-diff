# sheet-visibility-only-structure-change

这个案例用来防住“单元格值没变，只改了 sheet 显示状态，但 diff panel 没把它当结构变化”的漏报。

## Naming

- 目录名：`sheet-visibility-only-structure-change`
- 含义：用户看到的是“内容相同”，根因是“只改了工作表显示状态”

## Incident

- 来源：结构性 diff 能力补齐
- 首次归档时间：2026-04
- 影响资源链路：`local file`、`git HEAD`、`svn BASE`
- 关键定位：`define!F5`

## Symptom

- 用户看到的现象：两边内容一致，但一侧 `define` sheet 被隐藏了
- 当时漏判结果：如果只看 cell diff，会把这类变化完全忽略掉

## Root Cause

- 底层真实差异：一侧 `define` 保持 `visible`，另一侧变成 `hidden`
- 为什么需要显示：这属于工作簿结构变化，会直接影响用户是否能看到目标 sheet

## Fixture

- `base.xlsx`：保留 `define` 和 `helper` 两张 sheet，`define!F5` 写入 `same`，`define` 保持可见
- `head.xlsx`：同样保留 `define` 和 `helper`，`define!F5` 仍为 `same`，再把 `define` 设为 `hidden`
- 关键 sheet / cell：`define!F5`

## Expected Behavior

- 应显示的差异：sheet visibility 结构变化
- 应忽略的差异：无

## Coverage

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`

## Regenerate

- 运行 `npm run generate:test-fixtures`
- 脚本入口：`scripts/generate-test-fixtures.mts`
