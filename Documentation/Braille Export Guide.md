# Braille Export

Select one or more books, press Ctrl+E (File > Export Book Copy), and choose a
destination folder. Choose BRF or eBraille, set the braille code, and activate
Export. Tab and Shift+Tab move through the settings and buttons.

The folder-based Send To commands for NLS eReader, Voice Dream, and Dolphin
EasyReader now offer the same format choices and put the converted copies in the
configured destination. Original-format copying is still available. HumanWare
MTP and Kindle email transfers are unchanged. Choose a format your destination
supports; the app does not infer device compatibility.

Braille conversion results flag detected images/graphics, tables, marked-up
mathematics, non-English language labels, and warnings returned by the converter.
The report is a read-only text field reachable with the keyboard; Tab moves to
Close and Shift+Tab returns to the report. These checks do not identify every
possible transcription problem, especially notation represented only as plain
text or images, and do not certify braille accuracy.

- BRF: paginated ASCII braille for devices that accept .brf files, including
  NLS braille eReaders. Choose English UEB contracted (grade 2) or uncontracted
  (grade 1). The default is 40 cells per line and 25 lines per page; these are
  adjustable and remembered.
- eBraille: structured .ebrl output with Unicode braille and navigation for
  compatible reading systems, including newer multiline displays. Do not assume
  an older reader supports eBraille merely because it supports BRF.
- This first implementation converts EPUB books. Other source formats must be
  converted to EPUB first using the existing conversion command. Image-only
  content is not transcribed. Existing BRF/eBraille files exported to their own
  format are copied without retranslation.
- Conversion is local. There is no conversion service account or end-user
  installation. The converter starts only when needed and shuts down afterwards.
  Starting it can take several seconds; the book manager remains responsive.
- Originals and library metadata are unchanged. Existing destination files are
  preserved using numbered filenames. Multiple BRF volumes are all exported.
- Cancel stops the remaining batch, retains already completed books, and removes
  partial copies of the current book. Failures identify the affected books.

Automated transcription needs human review, especially for mathematics, complex
tables, mixed languages, and image-heavy books. The English UEB setting is not a
promise of correct transcription for every language or specialist notation.

## Development and Packaging

Run build_qt_preview_windows.bat. It verifies the bundled engine, fetching pinned
dependencies at build time only if they are not already staged. Keep the whole
generated Ebook Manager folder together, including _internal/vendor/braille.
The embedded converter increases the application's download and installed size.

Unit tests: python -m unittest tests.test_braille_export tests.test_qt_library_manager

Set AELM_TEST_BRAILLE=1 to additionally run real conversions with both output
formats and both grades. These tests use a synthetic book, not a user's library.
The tests check translated content, final text, navigation, original-file
preservation, collisions, cancellation, and accessible settings.
Set AELM_TEST_BRAILLE_LARGE=1 as well to include a synthetic book exceeding
500,000 characters and check both formats for complete final text.

## Third-Party Components

DAISY Pipeline is obtained from the official pipeline-ui 1.12.0 Windows release:
https://github.com/daisy/pipeline-ui/releases/tag/1.12.0

The engine modules are used unmodified. Upstream source and notices:
https://github.com/daisy/pipeline
https://github.com/daisy/pipeline-assembly
See _internal/vendor/braille/NOTICE.txt and the META-INF license/notice resources
within its system/common and system/webservice JARs.

The included JRE is Eclipse Temurin 17.0.20.1+1, replacing the older JRE in the
engine distribution. Its license and third-party notices are retained under
_internal/vendor/braille/jre/legal. Release and corresponding source information:
https://github.com/adoptium/temurin17-binaries/releases/tag/jdk-17.0.20.1%2B1
https://github.com/adoptium/jdk17u/tree/jdk-17.0.20.1%2B1

7-Zip is used only during build preparation and is not included in the app.
Pinned download URLs and SHA-256 checksums are in prepare_braille_runtime.py.
The shipped bundle.json records the version and hashes of the staged runtime.

Before public distribution, include corresponding source materials and license
notices for the exact dependency versions in the release distribution; do not
strip notices or restrict replacement of the embedded runtime components.
