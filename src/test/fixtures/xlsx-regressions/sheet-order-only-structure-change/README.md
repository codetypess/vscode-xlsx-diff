# sheet-order-only-structure-change

这个案例用来防住“单元格值没变，只改了 sheet 顺序，但 diff panel 没把它当结构变化”的漏报。

## Naming

- 目录名：`sheet-order-only-structure-change`
- 含义：用户看到的是“内容相同”，根因是“只改了工作表顺序”

## Incident

- 来源：结构性 diff 能力补齐
- 首次归档时间：2026-04
- 影响资源链路：`local file`、`git HEAD`、`svn BASE`
- 关键定位：`define!F5`

## Symptom

- 用户看到的现象：两边 `define!F5` 内容一致，但 sheet tab 顺序不同
- 当时漏判结果：如果只看 cell diff，会把这类变化完全忽略掉

## Root Cause

- 底层真实差异：`base.xlsx` 顺序为 `define -> helper`，`head.xlsx` 顺序为 `helper -> define`
- 为什么需要显示：sheet 顺序会影响用户浏览和上下文定位，属于工作簿结构变化

## Fixture

- `base.xlsx`：包含 `define` 和 `helper` 两张 sheet，`define!F5` 写入 `same`
- `head.xlsx`：保留相同内容，但把 `define` 移到索引 `1`
- 关键 sheet / cell：`define!F5`

## Expected Behavior

- 应显示的差异：sheet order 结构变化
- 应忽略的差异：无

## Coverage

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`

## Regenerate

- 运行 `npm run generate:test-fixtures`
- 脚本入口：`scripts/generate-test-fixtures.mts`
