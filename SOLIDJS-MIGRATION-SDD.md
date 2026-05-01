# SolidJS Rewrite SDD

Status: Draft
Date: 2026-05-01
Scope: Webview runtime only
Decision: Rebuild the editor and diff webviews as new SolidJS applications optimized for current product behavior, without preserving the current React component structure, DOM shape, CSS organization, or internal host-webview protocol unless they are still the best choice.

## 1. Context

The current webview runtime is implemented directly on React and React DOM.

- `src/webview/editor-panel/editor-panel-app.tsx` is a very large editor runtime that mixes shell composition, message intake, virtualization, grid rendering, selection logic, and DOM effects.
- `src/webview/editor-panel/editor-panel-chrome.tsx` holds toolbar, search panel, tabs, and other editor chrome.
- `src/webview/diff-panel/diff-panel-app.tsx` implements the diff UI.
- `tsconfig.webview.json` is configured for React JSX.
- `scripts/esbuild.mts` bundles the extension host and the webviews from the same custom script.
- `package.json` still includes React runtime and type dependencies.

Recent profiling showed that the current editor path is structurally sensitive to React update mechanics such as `flushSync`, `React.memo`, prop identity stability, and effect timing. The immediate bugs can continue to be fixed on React, but the architecture remains tuned around avoiding broad rerenders rather than expressing narrow ownership directly.

This document therefore does not describe a conservative React-to-Solid port. It describes a rewrite of the webview layer using SolidJS as the primary architecture.

## 2. Product Boundary

The rewrite only needs to preserve the current product surface.

The following user-visible capabilities must still exist after the rewrite:

1. Single-file XLSX editor.
2. XLSX diff panel opened from commands, Explorer, and Source Control.
3. Sheet tabs.
4. Search, replace, and goto in the editor.
5. Cell editing, pending edits, save, undo, and redo.
6. Alignment controls.
7. Selection, range extension, fill drag, and keyboard navigation.
8. Row and column operations already supported by the editor.
9. Diff navigation, row filters, inline diff editing, and save in the diff panel.
10. Read-only behavior for non-writable or history-backed resources.

Everything below that surface is open for redesign.

## 3. Goals

The rewrite must achieve the following goals.

1. Replace the entire React webview runtime with SolidJS.
2. Rebuild the UI around fine-grained reactivity rather than tree-wide invalidation and memo comparators.
3. Split editor and diff runtimes into explicit state, rendering, and DOM-effect layers.
4. Make localized operations stay localized, especially cell edits, alignment changes, selection moves, and viewport updates.
5. Simplify the mental model of state ownership between the extension host and the webview.
6. Improve maintainability by replacing the current large TSX monoliths with smaller bounded modules.
7. Keep the extension host focused on workbook I/O, SCM integration, persistence, and authoritative workbook session state.

## 4. Non-Goals

The rewrite explicitly does not optimize for the following.

1. Preserving the current React component tree or file layout.
2. Preserving CSS class names, DOM structure, or snapshot stability.
3. Preserving the current internal host-webview message schema.
4. Keeping the current custom esbuild-only webview build if a different webview build path is better for Solid.
5. Minimizing the diff size.
6. Shipping a mixed React and Solid runtime on the main branch.
7. Preserving existing helper boundaries if they no longer make sense in the Solid architecture.

## 5. Primary Architectural Decision

The extension host remains TypeScript and continues to own workbook loading, saving, SCM integration, and authoritative workbook state.

The webviews are rebuilt as two dedicated Solid applications.

Key decisions:

1. Use SolidJS as the only webview runtime.
2. Use Vite with `vite-plugin-solid` for webview builds because it is the primary Solid build path and gives a better development model than stretching the current custom esbuild script further.
3. Keep the extension host build on the current TypeScript and esbuild path unless a separate future change is needed.
4. Replace the internal host-webview message contract with a cleaner typed protocol designed for the new apps.
5. Treat the rewrite as a greenfield webview implementation constrained only by feature scope, not by the old React architecture.

## 6. Target Architecture

### 6.1 High-Level Shape

The target system has three layers.

1. Extension host layer.
2. Webview session layer.
3. Webview UI layer.

The host layer owns workbook persistence and authoritative session mutations.

The webview session layer owns normalized state for the current view model plus ephemeral UI state such as selection, focus, editing drafts, search panel state, viewport state, and filter menu state.

The webview UI layer renders from Solid signals, stores, and derived memos and is responsible for DOM measurement, scrolling, and focus coordination.

### 6.2 Extension Host Responsibilities

The host remains authoritative for:

1. Opening and saving workbooks.
2. Loading workbook snapshots and structure.
3. Applying edits that mutate workbook state.
4. Managing undo or redo state that must survive across view refreshes.
5. SCM resource resolution and compare flows.
6. Error handling and resource lifecycle.

The host does not remain constrained by the current render payload shape. It may expose a new protocol better suited to Solid.

### 6.3 Webview Responsibilities

The webview owns:

1. Viewport state.
2. Selection model.
3. Editing drafts and focus transitions.
4. Search panel and filter menu presentation state.
5. Sheet tab selection UI.
6. Virtualized grid rendering.
7. DOM measurement and scroll synchronization.

The webview should operate on normalized state rather than repeatedly materializing large view objects during render.

### 6.4 State Ownership Model

The Solid runtime should use explicit ownership rules.

1. Use `createStore` for structured app state.
2. Use `createSignal` for local scalar state and DOM-related state.
3. Use `createMemo` for derived slices with real fan-out.
4. Use `batch` when multiple host messages or UI actions update related state together.
5. Keep cell-local derivations as close as possible to cell ownership.
6. Avoid a single top-level store mutation pattern that invalidates the whole grid.

### 6.5 Message Protocol

The host-webview protocol may change.

The target protocol should have these characteristics.

1. Typed commands and events.
2. Clear separation between authoritative workbook data and ephemeral UI state.
3. Initial bootstrap message for full session hydration.
4. Patch-style updates for workbook deltas when practical.
5. Explicit lifecycle events for loading, saving, error, and disposal.

Recommended protocol groups:

- `session:init`
- `session:patch`
- `session:status`
- `editor:command`
- `diff:command`
- `dialog:request`
- `telemetry:debug`

The old protocol can be retired as part of the rewrite.

### 6.6 Rendering Model

The new UI should be split by ownership, not by the current file boundaries.

Recommended editor structure:

- `editor-app/boot.tsx`
- `editor-app/store.ts`
- `editor-app/protocol.ts`
- `editor-app/session.ts`
- `editor-app/components/AppShell.tsx`
- `editor-app/components/Toolbar.tsx`
- `editor-app/components/SearchPanel.tsx`
- `editor-app/components/SheetTabs.tsx`
- `editor-app/components/GridViewport.tsx`
- `editor-app/components/FrozenGridLayer.tsx`
- `editor-app/components/CellLayer.tsx`
- `editor-app/components/Cell.tsx`
- `editor-app/components/SelectionOverlay.tsx`
- `editor-app/components/FilterMenu.tsx`
- `editor-app/services/selection.ts`
- `editor-app/services/viewport.ts`
- `editor-app/services/editing.ts`
- `editor-app/services/dom-effects.ts`

Recommended diff structure:

- `diff-app/boot.tsx`
- `diff-app/store.ts`
- `diff-app/protocol.ts`
- `diff-app/components/AppShell.tsx`
- `diff-app/components/Toolbar.tsx`
- `diff-app/components/Grid.tsx`
- `diff-app/components/SheetTabs.tsx`
- `diff-app/components/SelectionPreview.tsx`
- `diff-app/services/selection.ts`
- `diff-app/services/scroll.ts`

The actual file names may vary, but the rewrite must land with explicit stores, protocol adapters, components, and DOM services rather than recreating a single giant application file.

### 6.7 Virtualization Strategy

The virtualization math can be reused where it is correct, but the rendering strategy should be redesigned for Solid.

Requirements:

1. Visible rows and columns are derived from viewport signals.
2. Frozen panes are rendered as separate owned layers.
3. Cells are keyed by cell identity and owned by narrow reactive inputs.
4. Column metrics and row metrics are stable domain values rather than transient wrapper objects.
5. The grid must not depend on component-level custom comparators to remain fast.

The rewrite should prefer a model where cell components subscribe only to the reactive values they need, rather than receiving broad parent props whose identity must be preserved manually.

### 6.8 DOM Effects and Measurement

DOM effects should be isolated from rendering logic.

Rules:

1. Use `onMount` for setup.
2. Use `createEffect` or `createRenderEffect` only for true DOM synchronization work.
3. Keep scroll, resize, focus, and text-measurement code in services or effect modules.
4. Avoid hidden effect chains that indirectly update broad UI state.

### 6.9 Styling

The rewrite does not need to preserve the current CSS structure.

Recommended styling rules:

1. Keep the current product look and VS Code integration quality.
2. Reorganize styles into tokens, layout, chrome, grid, overlays, and dialogs.
3. Prefer codicons and inline SVG over React-only icon packages.
4. Keep class names stable only where existing tests or host HTML wiring need them.

## 7. Build and Tooling

### 7.1 Build Split

The repository should use two build paths.

1. Extension host build: keep the current TypeScript plus esbuild path.
2. Webview build: move to Vite plus `vite-plugin-solid`.

This is the preferred split because the extension host and webview have different optimization needs and different framework requirements.

### 7.2 Required Dependency Changes

Add:

- `solid-js`
- `vite`
- `vite-plugin-solid`
- `vitest`
- `jsdom` or `happy-dom` for webview tests
- optional Solid testing helpers if needed

Remove after cutover:

- `react`
- `react-dom`
- `@types/react`
- `@types/react-dom`
- `react-icons`

### 7.3 TypeScript Changes

The Solid webview build should use a dedicated TypeScript config with:

- `jsx: preserve`
- `jsxImportSource: solid-js`

The extension host config remains separate.

### 7.4 Output Contract

The extension can keep consuming `media/editor-panel.js` and `media/diff-panel.js`, but this is only an output wiring choice, not an architectural constraint. The rewrite may generate those files from Vite build output.

## 8. Testing Strategy

### 8.1 Test Stack

The rewrite should separate testing by layer.

1. Keep `vscode-test` and current extension integration tests for extension-host behavior.
2. Introduce Vitest-based unit tests for Solid webview stores, services, and components.
3. Add focused DOM tests for selection, toolbar drafts, viewport updates, and diff interactions.

### 8.2 Required Coverage Areas

At minimum, the rewrite must have automated coverage for:

1. Editor sheet switch.
2. Editor search, replace, and goto.
3. Editor pending edits, save, undo, and redo.
4. Editor alignment and localized cell updates.
5. Editor selection drag and keyboard navigation.
6. Diff panel row filtering.
7. Diff panel edit staging and save.
8. Diff panel synchronized horizontal scrolling.

### 8.3 Performance Validation

Manual profiling remains required for:

1. Editor single-cell alignment.
2. Editor single-cell value edit.
3. Editor selection movement across viewport edges.
4. Editor sheet switch.
5. Diff panel filter changes.
6. Diff panel save after staged edits.

The Solid rewrite is successful only if these flows do not regress in perceived responsiveness.

## 9. Delivery Plan

This work still lands on a feature branch and merges only after the Solid rewrite is complete, but the implementation plan is a rewrite plan rather than a constrained migration plan.

### Phase 0: Rewrite Baseline

Tasks:

1. Approve this SDD.
2. Freeze unrelated webview refactors.
3. Capture baseline product and performance behavior for current flows.
4. Decide final target folder structure for the new Solid apps.

Exit criteria:

1. Rewrite branch exists.
2. Feature boundary is explicit.
3. Baseline traces and manual test matrix are captured.

### Phase 1: New Webview Platform

Tasks:

1. Add Vite plus Solid build pipeline.
2. Add new Solid entry points for editor and diff.
3. Keep extension host build unchanged.
4. Wire generated assets back into the current extension HTML.
5. Add Vitest and basic webview test harness.

Exit criteria:

1. Empty Solid apps build and load inside VS Code.
2. Development sourcemaps work.
3. Webview test harness runs in CI or local automation.

### Phase 2: Protocol and Session Model

Tasks:

1. Define the new typed host-webview protocol.
2. Define normalized session state for editor and diff.
3. Separate authoritative workbook data from ephemeral UI state.
4. Implement protocol adapters in the host and webviews.

Exit criteria:

1. Protocol is typed and no longer React-shaped.
2. Webview state can hydrate from the host without using legacy render payload assumptions.

### Phase 3: Diff App Rewrite

Tasks:

1. Build the diff app from scratch on Solid.
2. Re-implement sheet tabs, row filters, diff navigation, selection preview, inline edit, and save.
3. Re-implement synchronized horizontal scrolling with explicit scroll services.

Exit criteria:

1. Diff flows match current product behavior.
2. The old React diff runtime is no longer used.

### Phase 4: Editor Shell Rewrite

Tasks:

1. Build editor shell, toolbar, search panel, tabs, context menu, and filter menu in Solid.
2. Implement editor stores for selection, drafts, save state, search, and filter state.
3. Re-implement host command dispatch against the new protocol.

Exit criteria:

1. Editor shell behavior works without the old React runtime.
2. Toolbar draft synchronization and search panel flows are stable.

### Phase 5: Editor Grid Rewrite

Tasks:

1. Build the virtual grid, frozen panes, headers, cell layer, overlays, and editing surfaces in Solid.
2. Re-implement selection, fill drag, keyboard movement, and active cell editing.
3. Wire localized cell subscriptions so localized mutations remain localized.
4. Rebuild DOM effect coordination for scroll, focus, measurement, and overlay positioning.

Exit criteria:

1. The editor grid fully works on Solid.
2. No React editor runtime remains.
3. Localized updates remain localized in performance traces.

### Phase 6: Product Completion

Tasks:

1. Close remaining feature gaps.
2. Revalidate read-only behavior, structural edits, and save flows.
3. Finish test coverage and manual regression checks.
4. Remove dead React files and compatibility shims.

Exit criteria:

1. Current product capability is complete.
2. The React webview runtime is fully removed.

### Phase 7: Cleanup and Merge

Tasks:

1. Remove React dependencies and old webview build wiring.
2. Update documentation and contributor setup.
3. Verify build, test, and profiling gates.

Exit criteria:

1. Main branch is ready to run only the Solid webviews.
2. Legacy React webview files are deleted or archived outside the active runtime.

## 10. Acceptance Criteria

The rewrite is complete only if all conditions below are met.

### 10.1 Product Completion

1. All current editor and diff user-visible features are present.
2. Current commands, settings, and extension entry points continue to work.
3. Read-only and writable workbook behaviors remain correct.

### 10.2 Technical Completion

1. No webview runtime dependency on React remains.
2. No webview type dependency on React remains.
3. The webview build is fully Solid-based.
4. The extension host is wired to the new webview assets and protocol.

### 10.3 Performance Completion

1. Localized editor changes do not invalidate broad visible grid slices unless the viewport or layout truly changed.
2. Single-cell edit and alignment flows avoid long main-thread stalls in normal use.
3. Scroll, sheet switch, and diff navigation are at least as responsive as the current implementation.

### 10.4 Codebase Quality

1. The new Solid webviews are modular and do not collapse back into a single giant runtime file.
2. Stores, protocol adapters, services, and components have explicit ownership boundaries.
3. DOM effect code is isolated from domain logic.

## 11. Risks

### 11.1 Technical Risks

1. The editor grid rewrite is the dominant risk because it combines virtualization, selection, editing, overlays, and measurement.
2. Replacing the current protocol may surface hidden assumptions in host code.
3. Fine-grained reactivity can still be misused if the state graph is too coarse.
4. Moving webview builds to Vite introduces a new toolchain split that must be kept disciplined.

### 11.2 Product Risks

1. Temporary feature gaps during rewrite phases.
2. Keyboard and pointer interaction regressions.
3. Read-only and history-backed resource behavior regressions.

### 11.3 Schedule Risks

1. The editor grid phase can take longer than expected.
2. Test coverage may need to expand significantly once the new architecture exposes better boundaries.

## 12. Risk Mitigations

1. Rewrite diff first to validate tooling and protocol on a smaller surface.
2. Capture baseline traces before cutting over high-risk editor flows.
3. Keep host workbook logic stable where possible while replacing the UI runtime.
4. Use explicit store and protocol contracts early so implementation does not drift.
5. Add focused DOM and state tests as each subsystem lands.

## 13. Rollback Strategy

Rollback remains branch-based.

1. All rewrite work stays on a feature branch until completion.
2. The main branch keeps shipping the current React runtime until the rewrite is ready.
3. If product scope or performance goals are not met, the rewrite branch does not merge.
4. There is no production mixed-runtime fallback plan.

## 14. Work Breakdown and Rough Estimate

This remains a medium-to-large rewrite.

| Phase   | Scope                      | Estimate   |
| ------- | -------------------------- | ---------- |
| Phase 0 | baseline and approval      | 1-2 days   |
| Phase 1 | Solid plus Vite platform   | 2-4 days   |
| Phase 2 | protocol and session model | 3-5 days   |
| Phase 3 | diff app rewrite           | 4-7 days   |
| Phase 4 | editor shell rewrite       | 4-8 days   |
| Phase 5 | editor grid rewrite        | 10-20 days |
| Phase 6 | product completion         | 4-7 days   |
| Phase 7 | cleanup and merge          | 2-4 days   |

Expected total: roughly 5-9 weeks for one engineer already familiar with the codebase.

## 15. Merge Gate

The rewrite branch is ready to merge only when all conditions below are true.

1. Current product scope is fully implemented.
2. The Solid runtime is the only active webview runtime.
3. Performance validation is acceptable on key editor and diff flows.
4. Build, typecheck, lint, and tests pass.
5. Remaining architectural debt is not simply a re-creation of the current React monolith under Solid syntax.

## 16. Immediate Next Actions

1. Approve this rewrite-oriented SDD.
2. Create the rewrite branch.
3. Land the Vite plus Solid webview platform.
4. Define the new protocol and session model before implementing editor grid behavior.
