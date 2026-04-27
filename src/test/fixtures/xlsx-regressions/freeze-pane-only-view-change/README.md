# freeze-pane-only-view-change

这个案例用来防住“单元格值没变，只改了冻结窗格，但 diff panel 没把它当结构变化”的漏报。

## Naming

- 目录名：`freeze-pane-only-view-change`
- 含义：用户看到的是“内容相同”，根因是“只改了视图层冻结窗格”

## Incident

- 来源：结构性 diff 能力补齐
- 首次归档时间：2026-04
- 影响资源链路：`local file`、`git HEAD`、`svn BASE`
- 关键定位：`define!F5`

## Symptom

- 用户看到的现象：两边内容一致，但工作表视图锁定状态不同
- 当时漏判结果：如果只看 cell diff，会把这类变化完全忽略掉

## Root Cause

- 底层真实差异：一侧没有冻结窗格，另一侧设置了 `freezePane(1, 1)`
- 为什么需要显示：这属于工作表结构/视图变化，不是普通单元格值变化

## Fixture

- `base.xlsx`：`define!F5` 写入 `same`，不设置冻结窗格
- `head.xlsx`：`define!F5` 同样写入 `same`，再设置 `freezePane(1, 1)`
- 关键 sheet / cell：`define!F5`

## Expected Behavior

- 应显示的差异：`freezePane` 结构变化
- 应忽略的差异：无

## Coverage

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`

## Regenerate

- 运行 `npm run generate:test-fixtures`
- 脚本入口：`scripts/generate-test-fixtures.mts`
