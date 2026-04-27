# defined-name-only-structure-change

这个案例用来防住“单元格值没变，只改了 workbook defined name，但 diff panel 没把它当结构变化”的漏报。

## Naming

- 目录名：`defined-name-only-structure-change`
- 含义：用户看到的是“内容相同”，根因是“只改了工作簿命名区域”

## Incident

- 来源：结构性 diff 能力补齐
- 首次归档时间：2026-04
- 影响资源链路：`local file`、`git HEAD`、`svn BASE`
- 关键定位：`define!F5`

## Symptom

- 用户看到的现象：两边 `define!F5` 内容一致，但 workbook 级命名区域不同
- 当时漏判结果：如果只看 sheet / cell diff，会把这类变化完全忽略掉

## Root Cause

- 底层真实差异：`head.xlsx` 新增了全局 defined name `DefineCell=define!$F$5`
- 为什么需要显示：defined names 会参与公式引用、数据绑定和业务表关联，属于工作簿级结构变化

## Fixture

- `base.xlsx`：`define!F5` 写入 `same`，不设置 defined name
- `head.xlsx`：`define!F5` 同样写入 `same`，再新增 `DefineCell`
- 关键 sheet / cell：`define!F5`

## Expected Behavior

- 应显示的差异：defined names 工作簿结构变化
- 应忽略的差异：无

## Coverage

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`

## Regenerate

- 运行 `npm run generate:test-fixtures`
- 脚本入口：`scripts/generate-test-fixtures.mts`
