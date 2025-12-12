<!-- Import workspace-level CLAUDE.md configuration -->
<!-- See /home/lab/workspace/.claude/CLAUDE.md for complete rules -->

# Project-Specific Configuration

This file extends workspace-level configuration with project-specific rules.

## Project Context

JupyterLab 4 extension for viewing Microsoft Word documents (DOCX, DOC) and Rich Text Format (RTF) files directly in JupyterLab. Converts documents to PDF on-the-fly using pure Python libraries.

**Technology Stack**:
- Frontend: TypeScript, JupyterLab 4 extension API, Lumino widgets
- Backend: Python server extension, python-docx, reportlab
- Build: jlpm (Jupyter's yarn), webpack via @jupyterlab/builder
- Testing: Jest (frontend), Pytest (backend), Playwright (integration)
- CI/CD: GitHub Actions, npm, PyPI

**Package Names**:
- npm: `jupyterlab_doc_reader_extension`
- PyPI: `jupyterlab-doc-reader-extension`
- GitHub: `stellarshenson/jupyterlab_doc_reader_extension`

## Code and Content Generation Rules

- Always consider JOURNAL instructions
- Always consider markdown instructions and guidelines when creating documentation
- Always obey mermaid diagramming rules when creating diagrams
- Never add claude as coauthor into commit messages

## Context Persistence

**MANDATORY FIRST STEP**: At the start of EVERY session, you MUST:
1. Read `.claude/JOURNAL.md` (if it exists) before responding to any user query
2. Acknowledge what previous work was done based on the journal
3. Ask the user how to proceed based on that context

**MANDATORY AFTER EVERY TASK**: After completing substantive work, you MUST:
1. Update `.claude/JOURNAL.md` with the entry
2. Confirm to the user that the journal was updated

**Journal Entry Rules**:
- ONLY log substantive work on documents, diagrams, or documentation content
- DO NOT log: git commits, git pushes, file cleanup, maintenance tasks, or conversational queries
- Index entries incrementally: '1', '2', etc.
- Use single bullet points, not sections
- Merge related consecutive entries when natural

**Format** (include version number):
```
<number>. **Task - <short 3-5 word depiction>** (v1.2.3): task description / query description / summary<br>
    **Result**: summary of the work done
```

**When NOT creating journal entry**: State explicitly "Not logging to journal: <reason>"

## Folders

### DO NOT LOOK INTO

- `**/@archive`: folder that has outdated and unused content
- `**/.ipynb_checkpoints`: folder that has jupyterlab checkpoint files
- `**/node_modules`: npm dependencies

## Content Guidelines

### Markdown Standards

- No emojis - maintain professional, technical documentation style
- Balance concise narrative with structured bullet points
- Bullet points capture key takeaways and essential information
- Narrative focuses on value proposition, concrete benefits, and implementation details
- Include brief introductions but avoid fluff
- Explicitly state caveats and limitations where relevant
- Do not use em-dashes, use hyphens with spaces (` - `) instead
- Do not use full stop after a bullet point
- For mermaid diagrams use standard colours and not overloaded complex content
    - Use standard colours, no custom styles
    - Use diagrams to illustrate complex processes, workflows, or architectures
    - Do not overload diagrams with details, provide text narrative above or below
    - Do not use images and emojis
    - Only type of styling allowed: stroke and stroke-width for graph elements
    - **DO NOT use** `%%{init: {'theme':'neutral'}}%%` as it obscures colours in dark mode

## Documentation Standards

- Focus on concrete business value and technical implementation
- Include specific technology stacks and methodologies
- Maintain consistency across service descriptions
- Provide clear implementation timelines and phases
- Document success criteria and measurable outcomes

## Git Commit Standards

- Use conventional commit format: `feat: <description>` or `fix: <description>` or `chore: <description>`
- Keep descriptions concise and descriptive
- Use lowercase for commit messages
- IMPORTANT: Never attribute content creation to Claude - all content is authored by Konrad Jelen, Claude only assists with organization
- Do not include "Generated with Claude Code" or "Co-Authored-By: Claude" in commit messages
- Examples:
  - `feat: add unicode font detection for polish characters`
  - `fix: resolve vega-dataflow dependency issue`
  - `chore: update readme badges`

## GitHub Project Instructions

**Badge Template** (shields.io style):
```markdown
[![GitHub Actions](https://github.com/stellarshenson/jupyterlab_doc_reader_extension/actions/workflows/build.yml/badge.svg)](https://github.com/stellarshenson/jupyterlab_doc_reader_extension/actions/workflows/build.yml)
[![npm version](https://img.shields.io/npm/v/jupyterlab_doc_reader_extension.svg)](https://www.npmjs.com/package/jupyterlab_doc_reader_extension)
[![PyPI version](https://img.shields.io/pypi/v/jupyterlab-doc-reader-extension.svg)](https://pypi.org/project/jupyterlab-doc-reader-extension/)
[![Total PyPI downloads](https://static.pepy.tech/badge/jupyterlab-doc-reader-extension)](https://pepy.tech/project/jupyterlab-doc-reader-extension)
[![JupyterLab 4](https://img.shields.io/badge/JupyterLab-4-orange.svg)](https://jupyterlab.readthedocs.io/en/stable/)
```

**Link Checker Configuration**:
When using `jupyterlab/maintainer-tools/.github/actions/check-links@v1`, configure `ignore_links` parameter to skip badge URLs:
```yaml
- uses: jupyterlab/maintainer-tools/.github/actions/check-links@v1
  with:
    ignore_links: "https://www.npmjs.com/package/.* https://pepy.tech/.* https://static.pepy.tech/.*"
```
