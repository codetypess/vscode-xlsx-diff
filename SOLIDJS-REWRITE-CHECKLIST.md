# SolidJS Rewrite Implementation Checklist

Companion document for `SOLIDJS-MIGRATION-SDD.md`.

This checklist is execution-oriented. It assumes the project is taking the rewrite path described in the SDD: preserve the current product surface, but rebuild the webview runtime in SolidJS without preserving the current React architecture.

## 1. Working Rules

- [ ] Do all rewrite work on a dedicated feature branch.
- [ ] Freeze unrelated webview refactors while the rewrite is in progress.
- [ ] Keep extension-host workbook logic stable unless a protocol or state boundary change requires host work.
- [ ] Do not mechanically translate React code into Solid syntax.
- [ ] Keep the rewrite modular: store, protocol, services, and components must land as separate layers.
- [ ] Treat editor grid performance as a first-class acceptance condition, not as a post-merge optimization task.

## 2. Product Scope Lock

Before implementation starts, confirm the required product surface.

- [ ] Single-file XLSX editor remains supported.
- [ ] XLSX diff panel remains supported.
- [ ] Explorer, editor, and SCM entry points still open the correct views.
- [ ] Sheet tabs remain supported.
- [ ] Editor search, replace, and goto remain supported.
- [ ] Cell editing, pending edits, save, undo, and redo remain supported.
- [ ] Alignment controls remain supported.
- [ ] Selection, range extension, fill drag, and keyboard navigation remain supported.
- [ ] Existing row and column operations remain supported.
- [ ] Diff navigation, row filters, inline edit, and save remain supported.
- [ ] Read-only behavior for non-writable and history-backed resources remains supported.

## 3. Baseline Capture

Complete this before changing build or runtime architecture.

- [ ] Capture the current editor single-cell alignment trace.
- [ ] Capture the current editor single-cell edit trace.
- [ ] Capture the current editor sheet-switch trace.
- [ ] Capture the current editor selection-move trace.
- [ ] Capture the current diff filter-change trace.
- [ ] Capture the current diff save-after-edits trace.
- [ ] Save a short manual regression matrix for editor and diff behaviors.
- [ ] Record current commands, settings, and resource modes that must still work after rewrite.

Definition of done:

- [ ] Baseline traces and behavior matrix are stored somewhere discoverable in the branch.

## 4. Phase 1: Webview Platform

Goal: stand up a Solid-only webview platform without touching product behavior yet.

### Build and dependencies

- [x] Add `solid-js`.
- [x] Add `vite`.
- [x] Add `vite-plugin-solid`.
- [x] Add `vitest`.
- [x] Add `jsdom` or `happy-dom`.
- [x] Remove nothing yet unless it blocks the new platform work.

### Project structure

- [x] Create a dedicated Solid webview source root for the editor app.
- [x] Create a dedicated Solid webview source root for the diff app.
- [x] Create a dedicated Vite config for webview assets.
- [x] Create a dedicated TypeScript config for Solid webview code.
- [x] Keep extension-host build config separate from webview build config.

### Build output wiring

- [x] Make the Solid editor app produce the asset consumed by the editor webview HTML.
- [x] Make the Solid diff app produce the asset consumed by the diff webview HTML.
- [x] Preserve extension startup and packaging scripts after the build split.
- [x] Add watch mode for the Solid webview build.

### Smoke tests

- [ ] Open the extension in Extension Development Host with empty Solid roots.
- [ ] Verify the editor webview loads without runtime errors.
- [ ] Verify the diff webview loads without runtime errors.
- [ ] Verify sourcemaps work in webview devtools.

Definition of done:

- [ ] Solid webviews build and load.
- [x] Extension-host build still works.
- [ ] Packaging still produces a runnable extension.

## 5. Phase 2: Protocol and Session Model

Goal: replace the old React-shaped render payload path with a clean session protocol and normalized webview state.

### Protocol design

- [x] Define a typed host-to-webview bootstrap message.
- [x] Define a typed patch/update message family.
- [x] Define typed command messages from webview to host.
- [x] Define typed status and error messages.
- [x] Define protocol ownership rules: authoritative workbook data vs ephemeral UI state.

### Host changes

- [x] Build a host adapter that emits the new session bootstrap message.
- [x] Build a host adapter that emits incremental session patches where practical.
- [ ] Keep extension commands and editor providers wired to the new protocol.

### Webview session model

- [x] Create editor session store shape.
- [x] Create diff session store shape.
- [x] Split workbook-backed data from UI-only state.
- [x] Define how selection, editing drafts, viewport state, and panel state are stored.
- [ ] Ensure session hydration does not depend on recreating large denormalized render objects.

### Validation

- [x] Add unit tests for protocol encoding and decoding.
- [x] Add unit tests for session hydration.
- [x] Add unit tests for patch application.

Definition of done:

- [x] New protocol is typed end to end.
- [x] Webview state hydrates through the new session model.
- [ ] New work no longer depends on React-era render payload assumptions.

## 6. Phase 3: Diff App Rewrite

Goal: deliver the diff panel first as the lower-risk Solid app.

### UI shell

- [x] Build Solid app shell for diff panel.
- [x] Build diff toolbar.
- [x] Build sheet tabs.
- [x] Build row filter controls.
- [x] Build selection preview.

### Grid and interaction

- [x] Build the diff grid.
- [x] Implement diff navigation.
- [x] Implement inline diff editing.
- [x] Implement pending edits.
- [x] Implement save flow.
- [x] Implement synchronized horizontal scrolling.

### Validation

- [x] Test open-from-command flow.
- [x] Test row filter switching.
- [x] Test selection preview updates.
- [x] Test inline edit commit and cancel.
- [x] Test save after staged edits.
- [x] Test synchronized horizontal scrolling.

Definition of done:

- [ ] Diff panel matches current product behavior.
- [ ] Old React diff runtime is no longer used.

## 7. Phase 4: Editor Shell Rewrite

Goal: deliver the non-grid editor UX on top of the new protocol and session model.

### App shell

- [x] Build editor app shell.
- [x] Build toolbar.
- [x] Build search panel.
- [x] Build replace controls.
- [x] Build goto controls.
- [x] Build sheet tabs.
- [x] Build context menu.
- [x] Build filter menu shell.

### State and commands

- [x] Implement search state store.
- [x] Implement toolbar draft state.
- [x] Implement save and dirty state indicators.
- [x] Implement command dispatch back to the host.
- [x] Implement read-only gating at the UI layer.

### Validation

- [x] Test toolbar draft synchronization.
- [x] Test search open, close, drag, and replace.
- [x] Test goto submission.
- [x] Test sheet switching from tabs.
- [x] Test read-only command disabling.

Definition of done:

- [ ] Editor shell features work without the React runtime.
- [ ] Search and toolbar flows are stable on Solid.

## 8. Phase 5: Editor Grid Rewrite

Goal: rebuild the highest-risk subsystem around Solid fine-grained ownership.

### Grid foundation

- [x] Define viewport store for rows and columns.
- [x] Define stable row and column metrics representation.
- [x] Build scroll container and viewport services.
- [x] Build frozen pane layers.
- [x] Build row header layer.
- [x] Build column header layer.

### Cell rendering

- [x] Build cell layer.
- [x] Build cell component keyed by stable cell identity.
- [ ] Ensure cell subscriptions only read the state each cell actually needs.
- [ ] Avoid broad parent props that recreate visible cell state on every view update.
- [x] Re-implement overflow and alignment rendering.

### Selection and editing

- [x] Rebuild selection model.
- [x] Rebuild selection overlay rendering.
- [x] Rebuild active-cell editing surface.
- [x] Rebuild range extension.
- [x] Rebuild fill drag.
- [x] Rebuild keyboard navigation.

### DOM coordination

- [x] Rebuild scroll synchronization.
- [ ] Rebuild focus management.
- [ ] Rebuild overlay positioning.
- [ ] Rebuild DOM measurement effects.
- [ ] Keep DOM effects isolated from session and domain logic.

### Validation

- [x] Test single-cell value edit.
- [x] Test single-cell alignment change.
- [x] Test keyboard navigation.
- [x] Test selection drag.
- [x] Test fill drag.
- [x] Test viewport scrolling.
- [x] Test frozen pane behavior.

Definition of done:

- [ ] Editor grid works fully on Solid.
- [ ] Old React editor runtime is no longer used.
- [ ] Localized changes remain localized in profiling.

## 9. Phase 6: Product Completion

Goal: close feature gaps and reach full product parity.

- [ ] Revalidate save flows across editor and diff.
- [ ] Revalidate pending edits across sheet switches.
- [ ] Revalidate read-only behavior.
- [ ] Revalidate row and column operations.
- [ ] Revalidate SCM-backed compare flows.
- [ ] Revalidate error and loading states.
- [ ] Revalidate bilingual language behavior.

Definition of done:

- [ ] Current product scope is complete on Solid.

## 10. Phase 7: Cleanup and Merge Readiness

Goal: remove legacy runtime pieces and make the branch mergeable.

### Cleanup

- [ ] Remove `react`.
- [ ] Remove `react-dom`.
- [ ] Remove `@types/react`.
- [ ] Remove `@types/react-dom`.
- [ ] Remove `react-icons` if no longer needed.
- [ ] Remove old React webview entry points.
- [ ] Remove dead React-only helpers.
- [ ] Remove obsolete build scripts and configs.

### Documentation and tooling

- [ ] Update contributor setup instructions.
- [ ] Update build instructions.
- [ ] Document new webview build commands.
- [ ] Document new test commands.

### Final validation

- [ ] Run full typecheck.
- [ ] Run full lint.
- [ ] Run extension tests.
- [ ] Run webview unit tests.
- [ ] Run manual regression matrix.
- [ ] Re-run baseline performance scenarios.

Definition of done:

- [ ] No active React runtime remains.
- [ ] Build and tests pass.
- [ ] Performance is acceptable on key editor and diff flows.

## 11. Critical Performance Gates

Do not merge if any of these are still failing.

- [ ] Single-cell alignment no longer causes broad visible-grid invalidation.
- [ ] Single-cell edit no longer causes long main-thread stalls in normal use.
- [ ] Selection movement remains responsive.
- [ ] Viewport scrolling remains responsive.
- [ ] Sheet switching remains responsive.
- [ ] Diff filter changes remain responsive.

## 12. Final Merge Checklist

- [ ] Product scope complete.
- [ ] Solid webviews are the only active webview runtime.
- [ ] React dependencies removed.
- [ ] Vite-based webview build integrated into packaging.
- [ ] Extension-host integration still works.
- [ ] Test suite and manual regression pass.
- [ ] Performance gates pass.
- [ ] Old React webview files are deleted or clearly retired from the active runtime.
