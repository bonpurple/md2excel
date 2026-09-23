# AGENTS.md

## Project

- Java 8 project.
- Keep compatibility with Java 8.
- Do not use Java 9+ language features or APIs.
- This project is built with Eclipse JDT, not Maven or Gradle.
- Dependencies are managed as local jars under `lib/`.

## Source layout

- Production code: `src/md2excel/`
- Tests: `test/md2excel/`
- Test suite: `test/md2excel/AllTests.java`
- Compile output and temporary verification files may be placed under `bin/`.

## Testing

- Tests use JUnit 4.
- During implementation, prefer the smallest relevant test scope needed for feedback.
- Before completing an independently reviewable Java change, compile all production and test sources with Java 8 and run the complete test suite once.
- Run the complete test suite with:
  `org.junit.runner.JUnitCore md2excel.AllTests`
- Do not repeatedly run the full suite after intermediate edits unless needed to diagnose a failure.
- Do not delete, disable, or weaken existing tests to make a change pass.
- When adding a new test class, register it in `AllTests`.
- Java 8 compiler:
  `C:\Program Files (x86)\Java\jdk1.8.0_181\bin\javac.exe`
- Use this compiler for verification. Do not search for other JDKs unless this path is unavailable.

## Refactoring policy

- Preserve existing externally observable behavior unless the task explicitly requests a behavior change.
- Preserve generated Excel values, positions, and formatting during refactoring.
- Preserve public class names, public method signatures, constructors, public fields, and enum values unless explicitly instructed otherwise.
- Prefer small, independently reviewable changes.
- Do not combine unrelated refactorings.
- For behavior that is insufficiently protected, add characterization tests before changing the relevant implementation.
- Do not infer desired Markdown behavior from external Markdown specifications when writing characterization tests. Verify the current implementation behavior first.
- Treat unusual existing behavior as current behavior unless the task explicitly requests a bug fix.

## High-risk areas

Changes involving the following classes require particular care because behavior depends on shared state and processing order:

- `RenderState`
- `RenderContext`
- `MarkdownRenderer`
- `MdBlockBoundary`
- `ParagraphUtil`
- `MarkdownLineParser`
- `MarkdownInline`
- Do not re-analyze these classes broadly unless the requested change requires it.

Do not perform broad redesigns of these classes as part of an unrelated task.

## Formatting

- Follow the existing Eclipse Java formatting style.
- Do not use the Eclipse Oxygen CLI formatter directly on repository files because it may corrupt UTF-8 Japanese text in this environment.
- Prefer minimal formatting changes consistent with surrounding code.
- If formatting is necessary, use the Eclipse IDE manually or verify formatter output on a copy before applying it.
- Java source files must remain UTF-8 without BOM and use LF line endings.
- Do not format unrelated files.
- Run `git diff --check` after changes.

## Git

- Do not commit unless explicitly requested.
- Do not stage unrelated files.
- Keep test additions and implementation refactoring in separate commits when practical.
- Existing commit messages are written in Japanese; follow the existing repository style when proposing commit messages.