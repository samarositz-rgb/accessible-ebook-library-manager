# M4B Export Feasibility

Research date: September 26, 2026. This is a proposal, not an implemented export.

## Recommendation

Start with local, chaptered narration from EPUB, using an installed Windows voice
and a bundled audio packager. Keep it separate from the screen reader: generate
audio to files, never route interface announcements or a user's screen-reader
speech into the audiobook. Offer a short voice preview, voice and speed settings,
chapter-level progress, cancellation, and recovery after an interrupted job.

Existing recorded audio is a separate path. Packaging recordings requires no
speech synthesis, but this app currently imports text/document formats, not an
audio library. Adding MP3/M4A/WAV imports would need its own metadata, ordering,
duplicate detection, and duration handling. Do not silently expand imports as a
side effect of adding text-to-audio export.

## Verified Building Blocks

Windows SpeechSynthesizer can produce an audio stream from text or SSML and
enumerate compatible installed voices. Only Microsoft-signed voices supported by
this API are available; do not promise that every JAWS, NVDA, or Narrator voice
will appear. Voice quality and availability need testing on the user's computer.
[Microsoft SpeechSynthesizer documentation](https://learn.microsoft.com/en-us/uwp/api/windows.media.speechsynthesis.speechsynthesizer)

FFmpeg provides native AAC encoding and MP4-family packaging, with chapter and
book metadata support. Its iPod muxer includes the M4B extension. We can use a
bundled, pinned build rather than require an end-user installation.
[AAC encoder](https://ffmpeg.org/ffmpeg-codecs.html#aac),
[formats and chapter metadata](https://ffmpeg.org/ffmpeg-formats.html),
[MP4-family implementation](https://github.com/FFmpeg/FFmpeg/blob/master/libavformat/movenc.c)

Bundling requires choosing an appropriate FFmpeg build and retaining its license,
build configuration, and corresponding source materials. Optional components can
change its licensing; this must be checked for the exact binary we distribute.
[FFmpeg distribution guidance](https://ffmpeg.org/legal.html)

## Proposed Implementation Work

1. Read every EPUB spine document in order, using its navigation for chapter
   boundaries. Do not reuse the search index extraction: that is length-limited.
2. Define handling for footnotes, repeated running headers, tables, page numbers,
   and image descriptions. Keep these choices explicit; never quietly omit
   substantive text. Image-only pages require OCR, which is outside this first
   audio implementation.
3. Generate narration in bounded chunks, cache completed chapters, and record
   voice/speed/source hashes so only compatible partial work can be resumed.
4. Measure actual generated durations to construct chapter timestamps, encode
   AAC, and package an .m4b with title, author, chapter names, and optional cover.
5. Verify duration, nonempty audio, chapter count/order, final content, and player
   compatibility before publishing the result. Use collision-safe filenames and
   preserve the source book.
6. Run the work in the background without playing narration while the user is
   navigating the app. Reuse the accessible conversion report for content risks.

This is a substantial feature rather than a new filename extension. Packaging is
the smaller portion; reliable narration and reading-order handling need the most
testing. Cloud voices could be a later optional backend, with explicit consent,
clear costs, and no automatic uploading of books. A local voice backend avoids
that upload requirement, but voice selection and practical generation speed
remain to be tested. No runtime estimate is asserted before a prototype.

## Decisions Before Implementation

- Is the first goal narration from existing EPUB books, packaging existing audio,
  or both? Narration best matches the library's current text-book focus.
- Which available voices are satisfactory? Test a short sample before generating
  a full textbook.
- How should footnotes and page numbers be spoken?
- Which reader apps/devices should be the compatibility test targets?
