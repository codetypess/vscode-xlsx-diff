# XLSX Regression Fixtures

这里收录真实 `.xlsx` 误报/回归样本，优先服务于 `loadWorkbookSnapshot`、`buildWorkbookDiff` 和资源侧兼容测试。

约定：

- 每个案例单独建目录，目录名使用 `问题类别-简述`
- 每个案例至少包含 `base.xlsx`、`head.xlsx` 和一份说明文档
- 所有 `.xlsx` 统一通过 `fastxlsx` 生成，避免手改压缩包内容

生成命令：

```bash
npm run generate:test-fixtures
```
