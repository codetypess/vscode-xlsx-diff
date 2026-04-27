# style-only-background-color

这个案例用来防住“单元格值完全一样，只改了背景色，但 diff panel 仍然报不同”的误报。

## Naming

- 目录名：`style-only-background-color`
- 含义：用户看到的是“值相同”，根因是“只有背景色样式变化”

## Incident

- 来源：与当前产品范围对齐的回归补强
- 首次归档时间：2026-04
- 影响资源链路：`local file`、`git HEAD`、`svn BASE`
- 关键定位：`define!F5`

## Symptom

- 用户看到的现象：两边单元格内容一致，但如果把样式也算进 diff，就可能误报
- 当时误判结果：值未变化，只有背景色不同

## Root Cause

- 底层真实差异：一侧是默认样式，另一侧仅修改了 `backgroundColor`
- 为什么 UI 看起来相同：当前 diff 关注值和公式，不关注样式层差异

## Fixture

- `base.xlsx`：`define!F5` 写入 `same`
- `head.xlsx`：`define!F5` 同样写入 `same`，再把背景色改为红色
- 关键 sheet / cell：`define!F5`

## Expected Behavior

- 应显示的差异：无
- 应忽略的差异：仅样式层变化导致的差异

## Coverage

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`

## Regenerate

- 运行 `npm run generate:test-fixtures`
- 脚本入口：`scripts/generate-test-fixtures.mts`
