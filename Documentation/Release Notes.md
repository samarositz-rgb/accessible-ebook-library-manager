# Changelog

## Tester Repairs - 2026-09-28

- New imports no longer assign file formats or sources as tags. Imported EPUB subjects and online categories are restricted to the predefined dropdown categories; unsupported subjects are left out. Existing library tags are not migrated or deleted. File format remains stored separately.

- Book List Reading now uses a Book list speech dropdown instead of separate preset buttons. Both reading and sound choices use Apply, Save, and Cancel; existing custom speech selections are preserved.

- Reproduced the Golden Gate search problem against an isolated copy of the library: unrelated full-text word matches were sorted before the matching title. Multiword full-text queries now match phrases, and title matches rank before author, other metadata, and content matches. Ctrl+F selects the previous query for replacement. Keyboard tests cover every view, repeated searches, empty results, and Escape.

- Settings stays open after commands. Missing-metadata sound and supplementary speech have Apply, Save, and Cancel controls; choosing a sound mode no longer closes Settings.
- Missing-metadata alerts use a short tone independent of the Windows notification sound scheme, and category changes now check for missing metadata too.
- Search checks the whole library instead of silently retaining previous filters. Empty results clear the accessible current book and keep focus in Search. Escape restores the unfiltered library.
- Avoided redundant accessible-name updates when book focus returns and corrected focus transitions around empty categories. The metadata wizard opens its review editor before returning to the main book list.
- Required-title feedback stays in the editor; the removal confirmation has a safe Cancel default without repeated message labels. Removed bulk editor boilerplate and combined folder-import options into one dialog.
- Single-book text detection runs in a background worker. Successful imports clear old search filters and display the new book immediately.
- EPUB metadata detection preserves paragraph boundaries, recognizes cataloging author statements, and leaves unknown sources blank instead of guessing Personal. Google Books rate-limit errors explain what happened and do not immediately retry the same service.
- A read-only check confirmed The Golden Gate Bridge remains in the live database and its stored EPUB exists. No live library or backup records were changed during diagnosis.

## Release Candidate - 2026-09-27

- Tag categories are alphabetical in both row and column layouts, including Other and Untagged.

- Routine setting changes, successful covers/backups/exports/transfers/scans, and no-result messages use brief accessibility announcements with details in Recent Messages instead of OK dialogs. Automatic scan results stay quiet; failures and partial-result warnings remain available.
- Online metadata lookup now opens one review dialog with Fill Blank Fields (default) or Replace Existing. Switching the choice preserves manually edited fields, and the book's source is not replaced with the metadata provider.
- Background watched-folder refreshes preserve the current book and do not take keyboard focus away from other controls.

- Find Cover Image Online now starts immediately without an extra online-lookup confirmation.

- Added Organize > Category Layout > Rows or Columns. Rows place categories vertically and books horizontally: Up/Down changes categories and Left/Right changes books. Columns remain the default. The preference persists across restarts, applies to every grouped view, and preserves the current book and selections; single-list navigation is unchanged.

- Removed ISBN and series number from bulk editing and protected them from bulk database updates. Both remain available in the individual metadata editor; shared fields are unchanged.

- Removed the extra OK dialog after successful bulk metadata editing; completion remains in the status bar and recent messages, with focus returned directly to the book list.

- Explicit metadata and selection feedback now sends accessibility announcements carrying the requested text, independent of supplementary status speech. Marked books include Selected when re-entered; held shortcuts no longer repeatedly toggle selection or open editing, and double-press metadata editing is tied to the same book.

- Fixed Bulk Edit Metadata Apply so it reaches confirmation, saves selected books, and returns focus to the current book. Added Ctrl+S and Alt+S and an explicit keyboard tab order.

- Tags now use a keyboard-accessible dropdown with checkable predefined categories. Multiple categories can be selected; Space toggles a category without closing the dropdown, and Escape returns to the editor.
- Existing imported/custom tags are preserved under Existing Tags. Tag view is bounded to ten categories plus Untagged; unrecognized existing tags appear under Other without changing their stored values.
- All enabled metadata alert modes now include missing tags. Added Missing Tags Only; Off still silences the alert.
- Reorganized the project into Application, Documentation, Source, Tools, and Archive; rebuilt the application at a stable, clearly named path.

## Qt Preview - 2026-09-26

- Replaced the repeated discard-question message box with native Keep Editing and Discard Changes buttons; Keep Editing is the default.
- Removed settings scroll-area wrappers and their keyboard stops. Moved sound choices into Book List Reading and File Explorer integration into General; consolidated backup folder selection.
- Recent Messages now retains the latest ten session messages as individual, arrow-navigable list entries, newest first.
- Keep invalid metadata in the editor and focus Title for correction; confirm discarding unsaved individual or bulk metadata edits.
- Added safe-point cancellation for folder scans/imports, reindexing, multipart reading, copy/export/conversion batches, online lookups, and watched-folder scans. Completed work is retained; cancelling reindexing no longer clears unprocessed books' indexes.
- Deliver background results only after threads finish, and keep shutdown responsive while active steps finish safely.
- Consolidated settings into categorized tabs. Added a session history of recent status messages and optional supplementary screen-reader alerts, off by default; explicit Alt+number reading remains unchanged.
- Fixed EPUB imports that treated a print-edition ISBN in dc:source as the book's supplying source. ISBN metadata now comes from the EPUB's own valid identifier, preferring the package identifier, while provider names remain intact.
- Consolidated display choices under Organize > Book List Display, with Ctrl+3 for letter columns and Ctrl+4 for exact-tag columns. Alt+number metadata shortcuts are unchanged.
- Alphabetical sorting now ignores leading A and The consistently with grouping and letter navigation.
- Added editable common-category tag dropdowns, retaining custom and multiple tags.
- Added a background metadata wizard: local detection with a larger text sample, Open Library, then Google Books, followed by review and save. Existing values and book source are preserved; unmatched records are not applied. Work-level publication details from Open Library are not mistaken for edition details.
- Removed the confirmation before the explicit Look Up Metadata Online command.

- Added conversion format choices to folder-based Send To commands for NLS eReader, Voice Dream, and Dolphin EasyReader.
- Added keyboard-accessible braille conversion reports for detected graphics, tables, mathematical markup, language mismatches, and converter warnings.
- Researched chaptered M4B narration and recorded the proposed approach in M4B_RESEARCH.md; audio export is not yet implemented.

- Added cancellation checks during EPUB preparation and streamed embedded resources to reduce peak memory use.
- Clean up unlocked temporary conversion files between books; Windows-locked source copies are removed when the converter shuts down.
- Strengthened eBraille checks for reading order, missing chapters, and translated body content, and BRF checks for page height.
- Keep multiple BRF volumes in numeric order and validate saved braille settings before opening the export dialog.
- Expanded regression tests for large books, cancellation, output rollback, and temporary-file cleanup.

## Qt Preview - 2026-09-25

- Added separate BRF and eBraille (.ebrl) choices to File > Export Book Copy (Ctrl+E).
- Bundled the DAISY Pipeline braille engine and Java runtime; users do not need to install conversion tools or send their books to an online service.
- Added remembered English UEB contracted/uncontracted settings and BRF page dimensions, with native keyboard-accessible controls.
- Braille export handles selected EPUB books in the background, supports cancellation, preserves originals, and never overwrites an existing exported copy.
- Unsupported source formats report an explicit error rather than producing an incomplete braille book.

## 0.3 - 2026-05-16

- Improved metadata detection for law textbook PDFs and DOCX files, including stronger title-page detection, noisy PDF title rejection, ISBN-13 preference, edition detection, and safer author and publisher guesses.
- Added Home and End keyboard navigation in the books list.
- Expanded library backup and restore so cloud backups include both the database and imported book files.
- Added regression tests for importer metadata detection and folder backup syncing.

## 0.2 - 2026-04-28

- Added EPUB accessibility checking that inspects package metadata, navigation, page list, headings, language, and image alt text coverage, then stores the results with the book metadata.
- Added optional Windows File Explorer right-click integration for importing supported files directly into the manager.
- Added a No edition marker in the metadata editor so books without an edition statement can be identified clearly.
- Polished menu labels and common error messages for tester-facing builds.
- Moved filter commands into an Organize, Filter submenu and removed the older Read Current Book command.
- Expanded the Organize, Sort menu with title A to Z, title Z to A, author A to Z, author Z to A, published year newest or oldest first, and date added newest or oldest first.
- Removed the NVDA announcement setting. NVDA book list announcements now turn on automatically when NVDA is detected and running.
- Added library database backup and restore through a user-chosen synced cloud folder, including Google Drive, OneDrive, iCloud Drive, another folder, daily, weekly, monthly, and on-demand schedules.
- Fixed Alt+number single-press reading so it exposes the requested field through a normal focused read-only edit control for screen readers such as JAWS, even when that field is not included in the current Book List Speech setting.
- Added Alt+number shortcuts in the books list:
  - Alt+1 reads the title.
  - Alt+2 reads the author.
  - Alt+3 reads the edition.
  - Alt+4 reads the year.
  - Alt+5 reads the ISBN.
  - Alt+6 reads the publisher.
  - Alt+7 reads the source.
  - Alt+8 reads the tags.
  - Alt+9 reads the format.
  - Alt+0 reads the date added.
- Added quick-repeat editing for editable metadata fields. Pressing the same Alt+number twice quickly opens the metadata editor focused on that field.
- Updated the metadata editor so it can open directly on title, author, edition, year, ISBN, publisher, source, tags, or notes.
- Added Help text explaining the new books-list shortcuts.
- Added `.gitignore` rules so build output, Python cache files, PyInstaller spec files, and Codex scratch index files are not uploaded as source changes.
- Updated the GitHub repository so it includes the source files and build script, while preserving the existing Windows executable.

## 0.1 - 2026-04-27

- Initial working version of Accessible Ebook Library Manager.
- Included the Windows executable in the GitHub repository.
