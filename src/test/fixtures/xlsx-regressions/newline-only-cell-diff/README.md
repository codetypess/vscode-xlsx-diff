# newline-only-cell-diff

首个真实工作簿回归案例，对应线上遇到的“看起来一样，但 diff panel 在 `define!F5` 误报不同”问题。

案例说明：

- `base.xlsx` 的 `define!F5` 使用 `LF` 换行
- `head.xlsx` 的 `define!F5` 使用 `CRLF` 换行
- 两边展示文本一致，预期不应再标记为 diff

当前覆盖：

- `local file -> local file`
- `git HEAD -> local file`
- `svn BASE -> local file`
