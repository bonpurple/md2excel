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
- After changing Java code, compile all production and test sources with Java 8.
- Run the complete test suite with:
  `org.junit.runner.JUnitCore md2excel.AllTests`
- Do not delete, disable, or weaken existing tests to make a change pass.
- When adding a new test class, register it in `AllTests`.
- Run the full test suite after each independently reviewable change.

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

Do not perform broad redesigns of these classes as part of an unrelated task.

## Formatting

- Use the existing Eclipse Java Formatter configuration.
- Format only Java files changed by the current task.
- Java source files must remain UTF-8 without BOM.
- Java source files must use LF line endings, as required by `.gitattributes`.
- The Eclipse CLI formatter may produce CRLF on Windows; normalize changed Java files back to LF after formatting.
- Do not format unrelated files.
- Run `git diff --check` after formatting.

## Git

- Do not commit unless explicitly requested.
- Do not stage unrelated files.
- Keep test additions and implementation refactoring in separate commits when practical.
- Existing commit messages are written in Japanese; follow the existing repository style when proposing commit messages.