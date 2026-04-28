# Changelog

## 0.2 - 2026-04-28

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
