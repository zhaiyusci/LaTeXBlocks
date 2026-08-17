# Changelog

All notable product-facing changes are recorded here. Version history begins with the first self-contained
Word-and-PowerPoint package line.

## [Unreleased]

## [0.2.127] — 2026-08-17

### Word

- Fixed fixed-size LaTeX Blocks being stretched after a mouse resize instead of
  being recompiled for the new frame. The resize monitor now observes Word mouse
  release as a fallback for drawing-layer gestures that omit the native capture-
  end accessibility event.
- Added a real foreground-Word interaction check for the resize mouse-release
  fallback, in addition to the existing exact-frame reflow and persistence test.

## [0.2.126] — 2026-08-17

### Word and PowerPoint

- Replaced the generic Office glyphs for Inline Math, Display Math, Numbered
  Math, LaTeX Block, and Equation Reference with a coherent mathematical Ribbon
  icon family. Equation Reference uses the selected Word-style document,
  bookmark, and action-arrow artwork without changing the add-in's behavior.
- PowerPoint's Insert Block command now shares the same LaTeX Block icon as Word.

## [0.2.125] — 2026-08-17

### Brand and product surfaces

- Adopted the new full-canvas LaTeX Blocks icon across the repository README, the
  Word and PowerPoint About dialogs, and the Windows installer/uninstaller.
- Added an About command to both Office Ribbon tabs with the installed version,
  host-specific product description, project home, and support link.
- Applied the same product icon to every add-in-owned editor and equation-reference
  window instead of leaving the default WinForms application glyph in the title bar.

### Word and PowerPoint

- Preserve authored LF line boundaries when transferring multi-line TeX between
  the Alternative Text persistence envelope and the Windows editor control.

## [0.2.121] — 2026-08-12

### Word

- Removed the second fixed wait after a Font Color palette commit and reduced the monitor boundary to the next
  scheduler turn, so SVG formula paint no longer trails Word's native text colour by roughly 175 ms.
- Existing visible solid Graphics Fill now changes through one RGB assignment without rebuilding the fill through
  `Solid()`, avoiding an intermediate/default paint frame.

## [0.2.120] — 2026-08-12

### Word

- Font Color synchronization now leaves Word's native character colour untouched and updates only an SVG formula's
  Graphics Fill. Theme colours therefore retain their theme slot and tint instead of being downgraded to direct RGB.
- Avoided redundant Graphics Fill writes, `ScreenUpdating` toggles, and unrelated ribbon invalidation in the
  colour-only path, preventing the delayed redraw visible with light theme colours.

## [0.2.119] — 2026-08-12

### Word

- Fixed mixed-range Font Color commands updating only some formulas. The committed selection colour is now the single
  target for every captured Inline, Display, and Numbered Math object; each drawing character and its SVG Graphics
  Fill are reconciled even when Word applies the native colour inconsistently across the selected pictures.

## [0.2.118] — 2026-08-12

### Word

- Fixed Numbered Math conversion to Display Math or another math kind being rejected because the conversion validator
  treated the source object's own numbered-equation scaffold as an external conflict.
- Added a complete Inline/Display/Numbered conversion matrix that verifies all six directions, including magic-header
  kind/role/source, inline boundaries, display tabs, numbered SEQ fields, and equation bookmarks.

## [0.2.117] — 2026-08-12

### Word

- Display Math and Numbered Math now unconditionally take ownership of their paragraph's custom tab-stop layout.
  Existing custom stops no longer block insertion; the add-in clears them and installs its center/right math stops.

## [0.2.116] — 2026-08-12

### Word and PowerPoint

- Replaced Title/JSON and PowerPoint tag persistence with one shared TeX magic-header envelope in `AlternativeText`.
  The source begins after `% !end-latexblocks`, retains meaningful blank lines, and is separated from metadata without
  escaping or flattening its LaTeX text. Committed objects keep `Title` empty. Only the new format is recognized.
- Fixed Word's Flat OPC SVG-media refresh path to locate the new `descr` contract while preserving native Graphics
  Fill, and derived non-persisted baseline/frame state from the live Office object.

## [0.2.115] — 2026-08-12

### Word and PowerPoint

- Fixed styled Blocks becoming solid or invisible after insertion. Fixed
  `LaTeXBlock` objects now define foreground and background through TeX only and do
  not use Office colour/fill APIs; Inline, Display, and Numbered Math retain their
  existing host-colour behavior. PowerPoint keeps searchable LaTeX directly in
  `AlternativeText`; Office may normalize CR/CRLF to LF but preserves line count,
  blank lines, Unicode, and terminal newlines. Identity, layout, SVG dimensions, and
  explicitly separate text/background style fields now share one versioned JSON
  object in `Title`. Word and PowerPoint use the same metadata schema; old
  semicolon Titles and PowerPoint metadata Tags are not read.
  The fixed-Block TeX wrapper passes the stored source without trimming boundary
  newlines, preserving paragraph-break semantics.
  The editor's Fill label is now Background.
- The unified installer now registers the Word and PowerPoint VSTO manifests
  directly instead of installing them as two independent ClickOnce products.
  Programs and Features therefore exposes one LaTeX Blocks entry, owned by the
  stable installer AppId, while upgrades remove registrations created by the old
  packaging path.
- Fixed the unified installer's local VSTO layout. The installer now packages
  each project's signed flat Release layout—the same layout used by development
  registration—rather than the ClickOnce `app.publish` tree, allowing Word and
  PowerPoint to load via `|vstolocal` without separate ClickOnce products.

## [0.2.107] — 2026-08-10

### Word

- Migrating an older Display Math object to the center-tab scaffold during
  Update now belongs to the same Word Undo record and rollback transaction as
  the SVG replacement.

## [0.2.106] — 2026-08-10

### Word

- Display Math now owns a Word center TabStop and a leading tab on its visual
  line, matching the centering model used by Numbered Math without becoming a
  separate paragraph or floating object.
- Display Math update, conversion, deletion, and Copy as LaTeX paths now migrate,
  preserve, or remove the centering scaffold atomically as appropriate.
- Display and numbered math can share one paragraph's owned center/right tab
  layout while ordinary author tabs remain protected from silent replacement.

## [0.2.105] — 2026-08-10

### Word

- Display Math updates now retain display-line caret semantics and reset the
  trailing Shift+Enter boundary to the body-text baseline.
- Converting to or from Numbered Math now refreshes all Word-owned equation
  numbers in the same undo transaction.
- In-flight format renders now include the persisted math kind in their stale
  result guard, preventing an Inline render from replacing a converted Display
  formula or vice versa.
- Numbered Math now uses its Word-owned numbering source validation through the
  production render path, rejecting TeX-side tags and unsupported environments.

## [0.2.104] — 2026-08-10

### Word

- Fixed baseline leakage after Display Math and Numbered Math. Their TeX-derived
  picture offset remains unchanged, while the trailing Shift+Enter boundary and
  insertion point are reset to the body-text baseline before subsequent typing.
- Numbered Math now locates its literal closing parenthesis and trailing manual
  line break from the actual Word range instead of assuming fixed field offsets.

## [0.2.103] — 2026-08-10

### Word

- Display and numbered math now reapply their derived TeX baseline after Word has
  created line/number scaffolding and moved the insertion point. This prevents
  Word from merging the drawing run back to `Font.Position = 0` when no U+2060
  boundary is present.

## [0.2.102] — 2026-08-10

### Word

- Inline Math, Display Math, Numbered Math, and LaTeX Block now have explicit,
  persisted object kinds and four distinct Ribbon insertion commands.
- The three math objects store only their delimiter-free math body. Their inline or
  display wrapper is added only for rendering, while LaTeX Block continues to accept
  unrestricted text-mode LaTeX source.
- Only Inline Math uses U+2060 boundaries. Display and numbered math remain
  InlineShapes in Word's text flow but own display-line scaffolding instead.
- The shared math editor can atomically convert inline, display, and numbered math,
  including Word SEQ fields, bookmarks, tab stops, and line scaffolding.

## [0.2.101] — 2026-08-10

### Word

- The Shift+Enter spacing toggle now maps Word's underlying compatibility flag
  to the native UI behavior correctly, and its Ribbon label is consistently English.

## [0.2.100] — 2026-08-10

### Word

- The Ribbon now exposes the current document's native “Don't expand character
  spaces on a line ending with Shift+Enter” compatibility option as a toggle.

## [0.2.99] — 2026-08-10

### Word

- Numbered equations now participate in the same Auto InlineShape format batch as
  ordinary inline and display formulas. Their role only selects display-style
  rendering and validates the available width; the surrounding SEQ field, bookmark,
  tabs, and paragraph scaffold remain host-owned and survive the shared SVG update.

## [0.2.98] — 2026-08-10

### Word

- A numbered display equation interleaved with ordinary formulas now completes a
  Ctrl+A Font Size update after save and reopen. If the ordinary direct-media batch
  reconstructs the surrounding OpenXML envelope, the independent equation update
  reacquires its current InlineShape by the persisted formula GUID instead of using
  the deleted COM wrapper.

## [0.2.97] — 2026-08-10

### Word

- Ctrl+A Font Size changes now retain the value committed in Word's native size
  control and reconcile the selection after Word has propagated the command. Inline
  and unnumbered display formulas therefore enter the same Auto InlineShape batch,
  even when the immediate selection size is temporarily mixed or stale.

## [0.2.96] — 2026-08-10

### Word

- Existing inline display formulas written by older Paste From LaTeX versions as
  unstyled Fixed Content are now interpreted as natural-size Auto formulas. Their
  next font-size update persists the corrected mode; explicitly styled Blocks remain
  Fixed even when their TeX source contains a display environment.

## [0.2.95] — 2026-08-10

### Word

- Native Font Size commits are now observed when an individual inline or display
  formula is selected as an InlineShape. The exact-selection Font Color guard remains
  separate, so formula colour continues to use Office Graphics Fill without rendering.

## [0.2.94] — 2026-08-10

### Word

- Removed the obsolete render-batch colour fallback now that every external
  colour-only change is handled directly through Office Graphics Fill. Added
  separate profiling for StemTeX font-size rendering and the Word batch commit.

## [0.2.93] — 2026-08-10

### Word

- Unnumbered display math pasted from LaTeX is now a natural-size Auto formula,
  not a user-sized Fixed Block. It therefore follows native Word font-size changes
  like inline math while retaining display-style TeX rendering.

## [0.2.92] — 2026-08-10

### Word and PowerPoint

- Pure external foreground changes are now always host paint operations. Word Font
  Color and Office Graphics Fill update ordinary formulas, numbered equations, and
  Fixed Blocks directly without creating a StemTeX render task or replacing the
  Office object. TeX-authored explicit child colours remain independent.

## [0.2.91] — 2026-08-09

### Word

- Mixed selections containing ordinary inline formulas together with numbered
  equations or Fixed Blocks now route each object through its correct update
  contract. Ordinary Auto formulas retain the fast batch path, while special
  objects no longer trigger the “Only ordinary Auto inline formulas” error.
- Font-size-only direct SVG replacement now clears stale Word subscript and
  superscript transforms before deriving the new formula baseline, while retaining
  independent run formatting and Graphics Fill per formula across paragraphs.

## [0.2.90] — 2026-08-09

### Word

- Font-size-only refreshes of automatic inline formulas now replace SVG media
  directly in one Flat OPC transaction, avoiding the preliminary `AddPicture` and
  its redundant PNG generation for every formula. Word still regenerates its PNG
  compatibility fallback during the final `InsertXML`; unsupported package layouts
  fall back to the normal picture-import path. The direct transaction may span
  multiple paragraphs in one Word story while preserving their text and paragraph
  marks; the old same-paragraph restriction applies only to the `AddPicture` path.
  Because Word does not serialize an Office Graphic's current Graphics Fill into
  `Range.WordOpenXML`, the direct path explicitly captures each old drawing's fill
  colour and replays it after import instead of incorrectly inferring it from the
  separate run `Font.Color` property.
- Font Color palette tracking now identifies the dropdown once on its click. Swatch
  hover events only record the native event tuple and make no UI Automation or MSAA
  provider calls; accessibility classification is deferred until an actual click.
  Escape closes and cancels the active palette transaction through a provider-free
  keyboard signal, preventing stale sessions after a popup is dismissed.
- Mixed text/formula colour reconciliation now suppresses intermediate Word repaints
  while applying native Graphics Fill to the formulas, then repaints once after the
  atomic undo transaction.

## [0.2.84] — 2026-08-09

### Word

- Unified single- and multi-formula colour reconciliation on the native Office
  Graphic Fill path, so selection-change fallbacks no longer rerender a lone SVG.
- Removed the superseded OpenXML/SVG colour-migration pipeline, obsolete host-colour
  markers, collapsed-caret invocation suppression, and unregistered UI Automation
  callbacks. The active path retains its live host-state validation and atomic undo.
- Unified styled Block foregrounds with formulas: Word and PowerPoint now leave the
  default SVG ink unset and apply the Block text colour through native Office Graphics
  Fill. SVG background, border, clipping, and explicit TeX colours remain independent.

## [0.2.83] — 2026-08-08

### Word

- Changed whole-formula colour to Word's native Office Graphic operation
  (`InlineShape.Fill.ForeColor.RGB`). Colour-only updates no longer render TeX,
  replace SVG drawings, move selections, or touch formula baselines and run
  highlighting.
- Exact formula selections are now owned exclusively by Word's **Graphics Fill**
  command. Removed the collapsed-caret `FontColorPicker` proxy.
- Restricted the native formatting monitor's UI Automation hit testing to Office
  Ribbon/gallery `NetUI` windows. Ordinary document clicks no longer enqueue
  accessibility provider calls, preventing colour-menu backlogs and delayed Word
  shutdown; the monitor is also disabled before host teardown.

## [0.2.82] — 2026-08-08

### Word

- Corrected the colour architecture for automatic formulas: default SVG paint now remains genuinely unset, so Word's
  native `Font.Color` is the complete operation. The common path performs no StemTeX render, media rewrite, drawing
  replacement, selection restore, or `ScreenUpdating` toggle; baseline and highlight therefore remain untouched.
- Existing SVGs with the concrete host-colour wrapper introduced by 0.2.80/0.2.81 are detached from that wrapper on
  their first colour change. Subsequent changes use the zero-write native path.

## [0.2.81] — 2026-08-08

### Word

- Fixed colour-only SVG refreshes accepting a transient zero `w:position` after Word applied Font Color to an exact
  formula selection. The same atomic OpenXML transaction now restores the formula character's derived baseline from
  its existing TeX depth without rerendering or resetting highlight and other run properties.

## [0.2.80] — 2026-08-08

### Word

- Removed the remaining per-formula contract reads, range lookups, deletion calls, and final identity checks from
  same-paragraph format batches. New SVGs are imported beside the old drawings, then the old drawing runs are removed
  together in the single OpenXML transaction. A 10-formula Word write profile fell from 1189 ms to 592 ms while
  preserving text and every formula contract.
- Moved host text colour out of TeX and into a marked, inherited SVG paint layer for new automatic formulas and styled
  Blocks. Whole-formula colour changes now preserve explicit source colours, geometry, and baseline while skipping
  StemTeX and AddPicture entirely. The embedded PNG compatibility fallback is updated selectively in the same Flat OPC
  transaction; a 10-formula local profile completed in 259 ms.
- Colour-only batches may span multiple paragraphs. They are prepared per paragraph, written back in reverse document
  order, and remain one Word undo operation without changing paragraph marks or surrounding text.

## [0.2.78] — 2026-08-07

### Word

- Batched same-paragraph font-size and font-colour refreshes now import their SVGs first and normalize all formula
  drawings in one OpenXML transaction. The original drawing-run properties are transplanted verbatim, while formula
  size and baseline are recalculated. A 10-formula local profile reduced Word replacement time from 3566 ms to
  1189 ms without losing highlight, language, emphasis, spacing, scaling, colour, or surrounding text.

## [0.2.77] — 2026-08-07

### Word

- Removed a redundant host-font and baseline write that was immediately discarded when Word normalized a newly
  imported SVG. A 10-formula local profile reduced Word replacement time from 3975 ms to 3566 ms (about 10%).

### Packaging

- Release builds now preserve older installers in `dist/release` and refuse to overwrite an existing installer or
  checksum for the requested version.

## [0.2.76] — 2026-08-07

### Word

- Added an immediate Font Size transaction for the Fluent Ribbon control and its detached popup list. Mouse selection
  and legacy `CommandBarComboBox.Change` signals now converge on the same deduplicated batch refresh without requiring
  the user to leave the selection.

## [0.2.75] — 2026-08-07

### Word

- Added an explicit mouse-release signal for the main half of Word's Font Color split button, whose accessibility
  provider does not reliably publish `EVENT_OBJECT_INVOKED`. Provider and mouse signals for the same gesture are
  deduplicated before entering the existing colour refresh transaction.

## [0.2.74] — 2026-08-07

### Word

- Replaced the ineffective extra-dispatch-turn Font Color synchronization with one bounded post-command delay before
  reading Word's committed range colour. Formula refresh no longer depends on leaving the selection.

## [0.2.73] — 2026-08-07

### Word

- Deferred native Font Color commits by one Word UI turn so the selected formulas read the colour after Word has
  applied it and refresh immediately without requiring the user to clear the selection.

## [0.2.72] — 2026-08-07

### Word

- Generalized the mixed-selection redraw batch from Font Color to renderer-backed Word formatting. Multi-formula
  Font Size changes now render as one batch, commit under one suspended-screen-update window, and restore the range
  selection once; combined size-and-colour fallback changes use the same path.

## [0.2.71] — 2026-08-07

### Word

- Routed pure text-colour refresh lists through the same batch writer regardless of whether they originated from the
  Ribbon colour monitor or the selection-change fallback. The fallback no longer bypasses batching and repainting
  formulas one by one.

## [0.2.70] — 2026-08-07

### Word

- Batch formula redraws produced by one mixed-selection Font Color command. All renders finish before Word replaces
  any drawing, the whole batch is committed with screen updating suspended, and the original selection is restored
  once. Selection-change fallback recognizes the batch as pending and no longer queues duplicate redraws.

## [0.2.69] — 2026-08-07

### Word

- Removed the cross-callback custom undo transaction introduced in 0.2.68. Word cannot safely keep that global
  transaction open while asynchronous formula renders complete; doing so could repeatedly retrigger selection-format
  reconciliation after Select All. Restored the stable asynchronous colour-refresh behavior.

## [0.2.68] — 2026-08-07

### Word

- Grouped a native Font Color change and all resulting LaTeX formula redraws into one custom Word undo record, so
  changing a mixed text-and-formula selection is reverted as one semantic operation.

## [0.2.67] — 2026-08-07

### Word

- Removed all desktop-root UI Automation event subscriptions from the in-process add-in. Font Color detection now
  uses the existing WinEvent/MSAA/mouse-confirmation stream and short-lived UIA classification queries only, avoiding
  both slow UIA unregistration and the post-Quit callback drain that could leave WINWORD hot for minutes.

## [0.2.66] — 2026-08-07

### Word

- Moved Font Color UI Automation unregistration to the final document's `DocumentBeforeClose`, while Word's Ribbon
  provider is still live. This prevents the windowless WINWORD/VSTO teardown path from spending minutes in a hot UIA
  callback drain after real colour/formula interactions.

## [0.2.65] — 2026-08-07

### Word

- Replaced selection range leases with scalar document/story/start/end descriptors. Restoring a selection now creates
  a short-lived Range only at the moment it is needed; document close performs neither `FinalReleaseComObject` nor a
  global GC. This removes the COM reentrancy wait reproduced inside `Document.Close()` on an otherwise blank document.

## [0.2.64] — 2026-08-07

### Word

- Removed the two full-heap collections and global finalizer wait from final-document close. LaTeX Blocks now releases
  only the persistent Range/Shape RCWs held by its selection and rerender state, preserving prompt process teardown
  without pausing for unrelated WebBrowser, UIA, or Office finalizers.

## [0.2.63] — 2026-08-07

### Word

- Corrected formula replacement after confirming that `InlineShapes.AddPicture` inserts beside an existing drawing
  even when given that drawing's one-character Range. Update now deletes the old character inside the custom undo
  transaction and inserts the new SVG at the saved position, so neither a transient nor surviving duplicate remains.
- Released persistent selection and rerender `Range`/`Shape` RCWs at `DocumentBeforeClose`. These references previously
  kept a windowless WINWORD process and its RenderHost/StemTeX process tree alive before Word could raise `Quit`.

## [0.2.62] — 2026-08-07

### Word

- Changed formula Update from insert-then-delete to Word's single-operation range replacement. The existing formula's
  non-collapsed drawing range is passed directly to `InlineShapes.AddPicture`, eliminating the temporary duplicate
  formula and its extra paragraph reflow while retaining custom-record rollback on failure.

## [0.2.61] — 2026-08-06

### Word

- Removed every Office COM call from the confirmed host-exit path. Word now owns destruction of its Quit, window,
  and CommandBar connection points instead of the add-in issuing synchronous event-unsubscribe RPCs while Office is
  already tearing those servers down.

## [0.2.60] — 2026-08-06

### Word

- Prevented Word/VSTO shutdown from waiting for desktop-wide UI Automation event removal. On an actual host exit the
  colour monitor makes callbacks inert, stops timers, releases native hooks, and lets process teardown reclaim UIA
  registrations; ordinary disposal still unregisters them synchronously.

## [0.2.59] — 2026-08-06

### Word

- Corrected inline baseline placement to use only the newly rendered TeX depth: `Font.Position = -round(depth)`.
  Word already interprets this value relative to the current line baseline, so neighboring text, manual breaks, and
  previously compensated formulas can no longer influence insertion, Update, Font Color, or Font Size refreshes.
## [0.2.58] — 2026-08-06

### Word

- Fixed inline-formula baseline drift at the start of a visual line separated by a manual break (`Shift+Enter`).
  Baseline resolution now treats Word's `\v` character as a hard boundary instead of inheriting the preceding
  line-break run's compensated `Font.Position` and subtracting the TeX depth again during an update.

## [0.2.57] — 2026-08-06

### Word

- Abstracted native Font Color as value-free `Began`/`Committed`/`Canceled` transactions instead of exposing raw
  UIA/MSAA events to the Word document layer. A gallery hover identifies a scoped swatch but does not commit; an actual
  click on that same swatch does. Opening or canceling the palette and canceling More Colors no longer apply a stale
  last-used colour; the main button, a committed swatch, and More Colors **OK** each produce one update. Popup close
  ordering cannot discard the active swatch, while a paired down/up on the live popup prevents an Escape followed by
  a stale-coordinate click from committing. Generation-bound mouse confirmation is reconciled only after Word
  processes its native command, and a single FIFO UI-thread queue preserves Begin-before-terminal ordering across
  callback threads.
- Made an ordinary text range containing one or more Auto formulas behave like text under Word's native Font Color:
  Word colours the text and formula drawing characters, then LaTeX Blocks immediately rerenders each changed formula.
  A shared range lease preserves the mixed selection across asynchronous replacements without stealing it back after
  the user moves elsewhere.
- Kept the formula as an exact `InlineShape` selection. On a confirmed colour commit, LaTeX Blocks briefly uses a
  collapsed caret to read Word's current picker value, restores the exact picture selection, applies the value to the
  drawing run, and queues the SVG refresh. The U+2060 spacing boundaries are never turned into a persistent text
  selection, so copy, arrow navigation, and picture handles retain normal Word behavior.
- Preserved Word's native colour semantics across SVG replacement: Automatic, direct BGR, and theme slot plus
  tint/shade remain distinct, while StemTeX receives the resolved display RGB. Independent highlight, underline,
  proofing/language properties still survive and the formula baseline is still recomputed.

## [0.2.56] — 2026-08-06

### Word

- Attempted to observe native Font Color through the resize mouse-capture completion path. Exact InlineShape selection
  proved to be a Word command no-op, and opening/canceling the gallery uses the same generic signal; 0.2.57 replaces
  this insufficient route with command-specific accessibility events and a transactional caret probe.

## [0.2.55] — 2026-08-06

### Word

- Made native formula-format refreshes property-aware rather than restoring the old drawing run wholesale. Font Size
  and Font Color remain renderer inputs, independent Word formatting such as highlight, underline, proofing, language,
  and other direct character attributes survives SVG replacement, while baseline position and script flags are always
  recomputed from the new formula contract. The completion path also rejects an in-flight render when either live size
  or colour has changed, including native colour races on legacy unstyled Fixed Blocks.
- Kept an exactly selected formula highlighted after a native Font Color or font-size refresh replaces its SVG.
  Selection is transferred only when the same old InlineShape is still selected at commit time; moving the caret
  while TeX renders is respected and never causes the add-in to steal the selection back.
- Preserved an existing paragraph mark when inserting or updating a formula at the end of a non-final paragraph.
  SVG normalization now distinguishes Word's temporary `InsertXML` separator from the document-owned paragraph
  terminator, so the following paragraph is no longer merged into the formula's paragraph.

## [0.2.54] — 2026-08-06

### Word and PowerPoint

- Restored a stable TeX line box on the first and final lines of ordinary Block text. Lowercase-only runs such as
  `aa` therefore retain the selected font's full typographic ascent/depth instead of collapsing to x-height at a
  Top-aligned frame edge. Standalone display math remains free of paragraph or strut injection, and all fixed-frame
  vertical placement remains inside TeX rather than the SVG shell.
- Applied the same standard zero-width TeX strut to every natural-width, single-baseline formula box, including
  inline math and Word-native displaystyle equations. These SVGs now use the strut as their minimum height/depth with
  `PreviewBorder=0pt`; their horizontal TeX width and visible baseline remain unchanged, while taller math still
  expands naturally.

## [0.2.53] — 2026-08-06

### Word and PowerPoint

- Removed the obsolete fixed-Block line-box strut while retaining Top/Middle/Bottom in both editors and existing v1
  style payloads. The authored outer frame minus padding now defines one exact, horizontally left-aligned TeX content
  box; TeX alone performs the selected vertical placement inside its fixed-height `vbox`. The SVG layer places that
  box at the padding origin and only adds fill, an inside border, and clipping, with no second alignment calculation.

## [0.2.51] — 2026-08-06

### Word

- **Paste from LaTeX** now parses adjacent dollar-delimited formulas such as `$a$$b$` as two inline formulas,
  without changing the existing `$$...$$` display-math syntax.

## [0.2.50] — 2026-08-06

### Word

- **Paste from LaTeX** now follows TeX newline semantics: a single physical newline becomes interword whitespace,
  while two or more consecutive newlines collapse to one Word paragraph break. Explicit `\\` remains a manual
  line break.

## [0.2.49] — 2026-08-06

### Word

- Updating an inline formula now recomputes its baseline from adjacent prose and the newly rendered TeX depth,
  repairing formulas whose drawing run has lost `w:position` instead of preserving the damaged position.

## [0.2.48] — 2026-08-06

### Word

- Mixed **Paste from LaTeX** now restores the original Word `Font.Position` before every formula insertion, so the
  TeX-depth compensation of an earlier inline formula cannot become the host baseline of a later formula.

## [0.2.47] — 2026-08-06

### Word and PowerPoint

- Simplified the Fixed Block box model: the authored outer frame minus exactly twice the padding is the TeX content
  box. Border thickness no longer changes the content measure; background and an inside border are painted only
  after StemTeX returns the exact content SVG.
- Wrapper setup is now space-neutral without ending or vertically shifting the surrounding TeX paragraph, so top
  padding is preserved as literally as left padding.

## [0.2.46] — 2026-08-06

### Word and PowerPoint

- Removed the real 1.67 em leading glue produced when StemTeX entered the request file in horizontal mode. Styled
  Block content now begins at the SVG content-box origin, with only the explicitly configured padding remaining.

## [0.2.45] — 2026-08-06

### Word and PowerPoint

- Fixed styled Block paragraphs now reset every inherited TeX left-margin mechanism (`parindent`, `leftskip`,
  `parshape`, `hangindent`, and `everypar`) so the content origin is exactly the configured padding edge.

## [0.2.44] — 2026-08-06

### Word and PowerPoint

- Fixed-width Blocks now keep the TeX viewport horizontally left-anchored. Padding and border define the left inset;
  widening a frame adds room on the right, while narrowing it clips the right edge instead of centering the content.
- Styled Fixed Blocks now derive an exact TeX content box from the authored outer frame minus padding and border.
  TeX owns paragraph indentation (`0pt`) and Top/Middle/Bottom placement; SVG only paints the background/border shell
  and clips at the requested outer bounds.

## [0.2.43] — 2026-08-06

### Word

- Scoped imported text styles now restore the original Word insertion format, so `\textit`, `\textsf`, and related
  commands cannot affect following plain text. Font-family commands now apply separate Western and CJK faces from
  the selected profile semantics, including Arial + SimHei for `\textsf` under CJK profiles.

## [0.2.42] — 2026-08-06

### Word

- Fixed **Paste from LaTeX** failing with Word error `0x800A16D4` when `\textsf`, `\textrm`, or `\texttt` tried to
  apply a Western font family through Word's incompatible East Asian font property.

## [0.2.41] — 2026-08-06

### Word

- **Paste from LaTeX** now converts `\textit`, `\emph`, `\textbf`, `\textsf`, `\textrm`, and `\texttt` into
  corresponding Word text formatting, including nested combinations, instead of inserting the commands literally.

## [0.2.40] — 2026-08-06

### Word

- Added **Paste from LaTeX** for mixed clipboard text. The mode-aware parser leaves escaped text such as `\%` as
  ordinary Word characters, ignores comments, rejects unmatched math delimiters, and creates editable Blocks only
  for genuine inline/display math delimiters and standard mathematical environments.

## [0.2.39] — 2026-08-06

### Word

- Added **Copy as LaTeX** for selected Word text. Ordinary text is escaped for LaTeX, recognized inline Blocks are
  replaced losslessly by their authoritative Alternative Text source, and Word-only inline boundaries and numbered
  equation tab/field scaffolds are omitted.

## [0.2.38] — 2026-08-05

### Block vertical alignment

- Styled Word and PowerPoint Blocks now align their Top/Middle/Bottom frame positions
  against a real TeX text line box rather than the visible ink bounds of the first
  glyph. A lowercase `x`, capital `A`, and descender-bearing `g` therefore reserve
  the same ascender/depth space at a given font size and selected leading.
- The line-box strut is rebuilt after the final leading is selected, so its height
  follows the editor's line-spacing value instead of a stale profile default.
- Standalone display math (`\[...\]`, equation/align-style environments) preserves
  its own TeX vertical list exactly: no wrapper paragraph, strut, or outer TeX
  colour command is injected. Its default foreground colour is inherited from the
  SVG root, while explicit colours in the author source still take precedence.

## [0.2.37] — 2026-08-05

### Word

- Fixed Content Blocks now have the same persistent style editor as PowerPoint: TeX font size,
  line spacing, uniform padding, Top/Middle/Bottom placement, text color, background fill, and
  border color/width.
- The shared style model keeps typography in TeX and the outer shell in SVG. Styled Word Blocks
  preserve their source in Alternative Text, persist their style in compact Title metadata, and
  repaint padding/fill/border at the correct edges after an inline or floating frame resize.
- Word and PowerPoint now give an editor-confirmed default style its literal meaning: 1.20× leading
  is authored in TeX and the SVG owns the viewport. Existing default blocks stay on their compatible
  bare-SVG route until edited, so opening old documents does not reformat them.
- Inline formulas and Word-native numbered equations remain deliberately unstyled so their
  running-text baseline and same-paragraph tab/field semantics are unchanged.

## [0.2.36] — 2026-08-05

### Word

- Hardened fixed-Block frame reflow around real Word gestures: the add-in now compares the
  frame before and after each gesture, so moving or rotating a Block never queues a render.
  Rapid consecutive resizes accumulate from the latest intended TeX measure rather than from
  stale document metadata.
- Reflow work is now keyed to the actual Word drawing object, rather than copied metadata, so
  independently resized copies cannot overwrite one another. A native text-colour or width
  refresh for a fixed Block is committed through the same framed SVG path.
- Physical Block frames no longer silently cap at 2000 pt. The TeX layout-width policy remains
  bounded at 30–2000 pt, while a valid user-owned SVG frame is preserved exactly.
- Fixed Content Blocks use this same resize-on-release contract in line with text and under every
  ordinary floating wrapping mode. Flow participation, moving, rotating, and changing wrapping
  never themselves cause a re-render.

### Rendering lifecycle

- Moved the native StemTeX renderer behind the x64 **LaTeXBlocks Render Host** process. Word and
  PowerPoint now use a versioned local named-pipe protocol for profile warm-up, latest-only
  previews, durable renders, and cancellation; Office never owns an in-process native renderer.
- The Office add-in owns its Render Host through a Windows Job Object. VSTO unload and application
  shutdown release the job immediately, so Word does not wait for a native renderer create or
  render call to return.

## [0.2.33] — 2026-08-05

### Word

- Fixed-width ordinary LaTeX Blocks now use one resize/reflow contract whether Word keeps them in line with text or
  exposes them as a floating object under any wrapping mode. Changing either outer-frame axis rerenders the TeX box
  and rebuilds an exact SVG viewport rather than persisting a Word image-scale transform.
- Native resize commits now begin at Word's mouse-capture end (mouse-up), via a process-scoped WinEvent whose callback
  performs no COM work and schedules one UI-thread continuation. The existing selection-transition path remains a
  non-polling fallback; **Reflow Frame** accepts either fixed Block representation.
- Editing a resized inline fixed Block now preserves its author-owned outer frame, matching floating Block editing.

## [0.2.32] — 2026-08-05

### Word

- Floating fixed-width LaTeX Blocks now persist an exact SVG frame separately from their TeX layout width. Native
  frame changes are rerendered and reframed without stretching TeX artwork; a width change derives a fresh measure,
  while a height-only change preserves the measure and clips or adds viewport space as needed.
- Added **Reflow Frame** for the selected floating Block. Word has no shape-resize event, so the same operation is
  also queued asynchronously when the selection leaves a resized Block; moving or rotating it does not rerender.

## [0.2.30] — 2026-08-05

### Word

- Fixed-width LaTeX Blocks remain editable after Word Layout Options converts them to floating SVG objects. Editing
  preserves their floating position, relative frame, wrapping, margins, supported object formatting, and metadata.

## [0.2.29] — 2026-08-05

### Word

- Removed the unsuccessful post-command Font Size refresh experiment. Existing selection-change refresh remains the
  supported native-format synchronization path.

## [0.2.27] — 2026-08-04

### Word

- Numbered equations now have an **Equation Reference** picker. It inserts a native, hyperlink-enabled Word `REF`
  field to the individual equation number, so references persist in DOCX files and follow **Update Numbers**.

## [0.2.26] — 2026-08-04

### Word

- Fixed a color-rendering regression where the wrapper could append a TeX word space to an auto-width inline formula.
  Word Font Color now changes paint only: the exact TeX box width is unchanged, including when the source ends in a line break.

## [0.2.25] — 2026-08-04

### Word

- Inline formulas, fixed-width blocks, and numbered display equations now inherit the native Word **Font Color** at
  insertion. Recoloring an existing block uses that same Word formatting as the source of truth and asynchronously
  rerenders its SVG without changing the authoritative LaTeX stored in Alternative Text.

## [0.2.24] — 2026-08-04

### PowerPoint

- Fixed fixed-height bare blocks with **Vertical: Top**: the TeX SVG viewport now
  begins at the host frame's top edge instead of remaining vertically centered.
  Middle and Bottom retain their respective viewport placements, including when
  the block otherwise uses the default style.

## [0.2.23] — 2026-08-04

### PowerPoint

- Native PowerPoint frame dimensions are now authoritative. When a reflowed TeX block still exceeds a user-shrunk
  frame, its SVG keeps the exact requested viewport and clips overflow instead of expanding back to a natural size.

## [0.2.22] — 2026-08-04

### PowerPoint

- Bundled a corrected StemTeX worker template for full-width request content: the worker now suppresses its own outer
  paragraph indentation before starting the request minipage.
- Moved PowerPoint block padding, background, border, and vertical placement out of TeX boxes and into the final SVG
  shell. Typography (leading and text color) remains in TeX. SVG borders are four in-viewport filled strips, so the
  trailing edge cannot be clipped by a centered stroke or an incorrect TeX box viewport.
- Added a regression check that verifies every generated frame rectangle fits inside the emitted SVG `viewBox`.

## [0.2.16] — 2026-08-04

### Reliability

- PowerPoint now defers the embedded preview browser until the Ribbon callback has returned and retries transient OLE
  `RPC_E_SERVERCALL_RETRYLATER` / call-rejected responses. A temporarily busy PowerPoint instance no longer reports a
  successful TeX render as a preview failure.

## [0.2.15] — 2026-08-04

### PowerPoint

- Unified every native PowerPoint resize handle under one host-frame contract. Every actual size change now queues a
  real asynchronous StemTeX layout pass: width changes derive a new stored typesetting measure, while height-only
  changes rerender the current measure. Translation and rotation do not reflow. The TeX SVG remains 1:1 and is never
  stretched or cropped.
- Removed the `VisualScale` concept entirely. Actual formula size is controlled only by **TeX size (pt)** and always
  rerenders.
- Added per-block styling in the PowerPoint editor: ordinary-paragraph line spacing, uniform padding,
  Top/Middle/Bottom placement, text color, background fill, and border color/width. The original author source
  remains in Alternative Text.

## [0.2.14] — 2026-08-03

### Reliability

- Preview cancellation is now isolated from queued insert/update renders. Closing an editor promptly cancels only
  obsolete preview work, and the renderer can recover from a failed or canceled profile initialization.
- Word and PowerPoint lifecycle, profile-switch, document mutation, undo, and Office-event paths now clean up
  transactionally. A failed operation preserves the existing object instead of partially replacing it.
- PowerPoint block recovery preserves geometry and visual-scale metadata during exceptional render/update paths.

### Verification and packaging

- Release smoke coverage exercises active-preview cancellation, renderer recovery, immediate shutdown, U+2060
  persistence, Word equation numbering, and PowerPoint replacement geometry.
- The release procedure now validates the installed PowerPoint VSTO package. VSTO cannot safely swap a solution
  identity between an installed codebase and a development directory through registry edits alone.

## [0.2.13] — 2026-08-03

### Word

- Auto-width inline formulas now use an exact TeX SVG box with a U+2060 WORD JOINER immediately on each side.
  Existing boundaries are reused on edit, adjacent formulas share their middle boundary, and a conversion to a fixed
  block removes unshared boundaries.
- The old horizontal signed-`wp:effectExtent` / adjacent-space measurement path is removed. All drawing effect
  extents are normalized to zero; TeX depth remains the only baseline-mapping input.
- Word smoke coverage now verifies U+2060 insertion, repeated updates, save/reopen, caret placement, adjacent
  formulas, and Auto-to-Fixed conversion against desktop Word.

### Packaging

- The `0.2.13` installer publishes both VSTO add-ins and bundles StemTeX `0.12.4` with the supported profiles.
- Documentation now separates product scope, object contracts, StemTeX integration, developer workflows, testing,
  release operations, and design decisions.
