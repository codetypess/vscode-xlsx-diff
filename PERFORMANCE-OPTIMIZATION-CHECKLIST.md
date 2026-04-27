# Project Roadmap And Performance Checklist

面向 `vscode-xlsx-diff` 的项目路线与大文件性能优化清单。

说明：

- 当前文件名仍然保留 `PERFORMANCE-OPTIMIZATION-CHECKLIST.md`
- 但内容已经扩展为项目级 backlog，包含性能、正确性、结构性差异、merge 流程、环境支持与工程化事项
- 样式级 diff 暂不纳入本清单

目标：

- 降低打开大 `.xlsx` 时的首屏等待时间
- 降低 diff / editor 面板的卡顿感
- 控制内存占用
- 提升 diff 正确性与用户信任感
- 为后续 merge、远程只读支持和工程化演进打基础
- 为后续 `fastxlsx` 能力升级预留清晰边界

## Priority Backlog

### P0

- [ ] `Diff` 正确性加固
- [ ] 结构性差异展示
- [ ] `diff-panel` / `editor-panel` 共享基础层
- [ ] 真实工作簿回归集

说明：

- 这一层不以“性能”单点为目标，但会直接影响后续优化改造的稳定性和可维护性
- 如果基础正确性和共享层不稳，性能优化很容易把问题放大

### P1

- [ ] 大文件性能优化第一阶段：不改 `fastxlsx`
- [ ] 建立性能观测、缓存、惰性加载、分阶段渲染
- [ ] 补性能基准和性能回归保护线
- [ ] diff panel 搜索与导航增强

说明：

- 这一层优先解决“体感卡顿”和“重复计算”问题
- 目标是在不动底层读取 API 的情况下先把首屏和切换体验拉起来

### P2

- [ ] 大文件性能优化第二阶段：需要 `fastxlsx` 支持
- [ ] 推进更轻的元数据读取、非空单元格遍历、范围读取、sheet 级懒打开
- [ ] 明确扩展层与 `fastxlsx` 的接口边界
- [ ] merge / 接受变更流程
- [ ] merge 历史、pending 状态与保存模型统一

说明：

- 这一层解决“真正的大文件读取成本”问题
- 当瓶颈已经明确落在 workbook 解析层时，再进入这一阶段
- 同时推进从“查看 diff”到“利用 diff 处理改动”的能力跃迁

### P3

- [ ] 联调、验收与发布准备
- [ ] 完整性能回归集
- [ ] Git / SVN / 本地文件三类资源的稳定性验证
- [ ] 输出后续 issue/backlog 与 `fastxlsx` API 需求文档
- [ ] 只读远程环境支持
- [ ] 发布与维护工程化
- [ ] 图表 / 透视表 / 宏 diff 预研

说明：

- 这一层偏收口和落地，确保优化不只是局部实验，而是可长期维护的方案

## Workstream A: Diff 正确性与结构性差异

### 1. `Diff` 正确性加固

- [ ] 系统梳理“看起来一样但底层不同”的误报场景
- [ ] 补换行、空白、富文本纯文本相同、公式缓存值相关回归测试
- [ ] 评估日期/数字格式化值相同但底层类型不同的展示策略
- [ ] 明确哪些差异属于“应显示”，哪些属于“应忽略”

### 2. 结构性差异展示

- [ ] 展示 `defined names` 差异
- [x] 展示 `freeze pane` 差异（已接入结构性 diff 判定与面板提示）
- [x] 展示 sheet 隐藏/显示状态变化（已接入结构性 diff 判定与面板提示）
- [ ] 展示 sheet 顺序变化
- [ ] 展示工作表增删改名的更明确提示
- [ ] 先做 sheet 级提示，再评估是否做更细粒度展示

### 3. 真实工作簿回归集

- [x] 引入真实 `.xlsx` fixture 驱动测试（已落地首个 `newline-only-cell-diff` 回归样本）
- [x] 建立误报案例归档目录（已收录 `define!F5` 换行差异案例）
- [x] 建立 Git / SVN / 本地文件三类资源的最小回归样本（已覆盖 `git HEAD`、`svn BASE` 与本地文件链路）
- [x] 为线上踩过的坑建立命名规范和复现说明（已补案例模板与命名约定）

进度：

- 已新增真实 `.xlsx` fixture 生成脚本、案例目录与 5 条回归样本
- 已补 Git / SVN / 本地文件三类资源样本
- 已补案例命名规范、复现说明模板与首个案例文档
- 下一步优先继续扩充“看起来一样但底层不同”的误报案例

## Workstream B: Webview 结构与可维护性

### 1. 共享基础层

- [ ] 抽离 diff/editor 共用的 tabs、按钮、宽度观测、基础工具函数
- [ ] 抽离共用的 sheet tab overflow 布局逻辑
- [ ] 减少两套面板重复维护的交互代码
- [ ] 为后续 merge 与搜索增强预留统一组件入口

### 2. 重复派生数据收敛

- [ ] 统一 pending summary、selection summary、tab width cache 等派生结构
- [ ] 评估是否建立更明确的 view-model 边界
- [ ] 避免主文件继续膨胀，维持 feature 拆分方向

## Workstream C: 大文件性能

说明：

- 这一部分保留现有 `Phase 1 / Phase 2 / Phase 3` 结构
- 其中 `Phase 1` 以扩展层优化为主
- `Phase 2` 以 `fastxlsx` 能力升级为主
- `Phase 3` 以联调与验收为主

## Phase 1: 不改 `fastxlsx` 先做

### 1. 观测与基线

- [ ] 增加 workbook 打开、snapshot 构建、diff 构建、webview render 的耗时打点
- [ ] 准备 3 组基准样本：中型表、大型表、超大表
- [ ] 记录首屏打开耗时、切 sheet 耗时、内存占用、diff 构建耗时
- [ ] 为性能回归建立最小基线文档

### 2. Snapshot / Diff 缓存

- [ ] 为 `loadWorkbookSnapshot` 增加基于文件路径 + mtime 的缓存
- [ ] 对 Git / SVN / SCM 资源增加基于资源 URI 的只读 snapshot 缓存
- [ ] 同一次 diff 打开流程里避免重复解析同一 workbook
- [ ] 在 auto-reload 时只失效受影响的缓存项

### 3. 惰性加载与惰性计算

- [ ] 打开 diff panel 时只优先构建当前激活 sheet 的 view model
- [ ] 非激活 sheet 延后计算 diff rows / diff cells
- [ ] 打开 editor panel 时优先构建激活 sheet，其他 sheet 延后准备
- [ ] 将“工作簿整体加载”和“sheet 级别 render model 构建”拆开

### 4. 分阶段渲染

- [ ] 先显示基本框架和 sheet tabs，再补充当前 sheet 内容
- [ ] 大型 diff 先显示 row/column 计数和 loading 状态，再填充详细差异
- [ ] 对重计算流程增加取消旧任务能力，避免切 sheet 时无意义计算继续运行
- [ ] 对频繁触发的 reload / resize / selection 更新增加节流或合并策略

### 5. 减少重复计算

- [ ] 复查 `buildWorkbookDiff` 是否存在可复用的 row/column signature 计算结果
- [ ] 复查 diff panel / editor panel 是否重复派生相同数据结构
- [ ] 对 sheet tabs 宽度、pending summary、selection summary 等派生数据做更稳定的复用
- [ ] 减少 render 期间的全量 Map / Array 重建

### 6. 大表交互体验

- [ ] 切换分页或 sheet 时优先保留上一次内容，避免白屏闪烁
- [ ] 为大表显示“正在分析”或“正在加载差异”提示
- [ ] 评估 diff panel 是否需要按页构建 diff 视图而不是一次铺满
- [ ] 评估 editor panel 的搜索、跳转、选区同步是否可以异步化

## Phase 2: 需要 `fastxlsx` 支持

### 1. 更轻的元数据读取

- [ ] 支持只读取 workbook 的 sheet 列表、sheet 名称、rowCount、columnCount、merge、freezePane
- [ ] 支持先读 sheet metadata，后按需读 cell 内容

### 2. 非空单元格遍历能力

- [ ] 提供类似 `sheet.getUsedCells()` 的能力，直接遍历非空单元格
- [ ] 避免扩展层按 `1..rowCount * 1..columnCount` 全矩形扫描
- [ ] 保证遍历结果能拿到 `displayValue`、`formula`、`styleId`

### 3. 范围读取能力

- [ ] 支持按 sheet + range 读取指定区域
- [ ] 支持优先读取当前可视区域附近的 cells
- [ ] 支持 editor panel / diff panel 分页范围读取

### 4. Sheet 级懒打开

- [ ] 支持按 sheet 单独打开或单独 materialize
- [ ] 避免一次性构造整本 workbook 的所有 sheet 数据
- [ ] 评估 SVN / Git 历史资源场景下的兼容方式

### 5. 流式/增量读取

- [ ] 评估 `Workbook.open` 是否可以暴露流式读取或增量解析接口
- [ ] 评估对超大 sheet 的分块读取能力
- [ ] 明确扩展层需要的最小 API，避免过早设计过大接口

## Phase 3: 联调与验收

### 1. 回归测试

- [ ] 为大文件场景增加 fixture 驱动测试
- [ ] 为缓存失效、reload、切 sheet、分页、搜索增加回归测试
- [ ] 为 Git / SVN / 本地文件三类资源各补一个性能路径回归样本

### 2. 性能验收

- [ ] 对比优化前后首屏耗时
- [ ] 对比优化前后内存峰值
- [ ] 对比优化前后切 sheet 和搜索响应时间
- [ ] 记录哪些优化不需要 `fastxlsx`，哪些必须依赖 `fastxlsx`

## Workstream D: Merge 与差异处理流程

### 1. Merge 交互

- [ ] 从 diff panel 接受左侧单元格
- [ ] 从 diff panel 接受右侧单元格
- [ ] 支持接受整行 / 整列 / 整张 sheet
- [ ] 明确 merge 后的 pending 状态展示

### 2. 历史与保存模型

- [ ] 为 merge 操作接入 undo / redo
- [ ] 统一 merge pending 与 editor pending 模型
- [ ] 明确保存前校验、冲突提示与回滚策略
- [ ] 评估是否需要单独的 merge session 状态

## Workstream E: 环境支持与体验补强

### 1. 只读与受限环境支持

- [ ] 支持 virtual workspace / 受限环境下的只读查看模式
- [ ] 评估 `untrustedWorkspaces` 的只读支持策略
- [ ] 区分“可读不可写”和“完全不可读”两类提示
- [ ] 保持 Git / SVN 历史资源的一致只读体验

### 2. 搜索与导航增强

- [ ] diff panel 支持按差异结果搜索
- [ ] 支持只搜 diff rows / 当前 sheet / 全 workbook
- [ ] 支持按公式、按值、按结构变化跳转
- [ ] 打开 diff 时自动定位首个变化 sheet / cell

## Workstream F: 发布与维护工程化

### 1. 工程化与发布

- [ ] 输出 issue/backlog 拆分文档
- [ ] 建立 release checklist
- [ ] 建立性能回归检查清单
- [ ] 建立真实样本回归维护规范

### 2. 中长期预研

- [ ] 图表 diff 预研
- [ ] 透视表 diff 预研
- [ ] 宏相关差异展示预研
- [ ] 明确哪些能力属于核心路线，哪些属于扩展路线

## 建议执行顺序

1. `Workstream A / Diff 正确性加固`
2. `Workstream A / 结构性差异展示`
3. `Workstream B / 共享基础层`
4. `Phase 1 / 观测与基线`
5. `Phase 1 / Snapshot / Diff 缓存`
6. `Phase 1 / 惰性加载与惰性计算`
7. `Phase 1 / 分阶段渲染`
8. 评估瓶颈是否仍主要在 workbook 读取层
9. 若仍然卡在读取层，再进入 `Phase 2`
10. 性能主线稳定后进入 `Workstream D / Merge`
11. 再推进 `Workstream E / 环境支持与体验补强`

## 预期输出文件

- `PERFORMANCE-OPTIMIZATION-CHECKLIST.md`
- 后续可拆出的 issue/backlog 文档
- 可能新增的 benchmark/fixture 目录
- 可能新增的 `fastxlsx` API 需求文档
- 可能新增的 merge 流程设计文档
- 可能新增的真实工作簿回归样本目录
