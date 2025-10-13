<!-- #region -->

# Code and Content Generation Rules
- always consider JOURNAL instructions
- always consider markdown instructions and guidelines when creating documentation
- always obey mermaid diagramming rules when creating diagrams
- never add claude as coauthor into commit messages


## Context Persistance

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

**Format**:
```
<number>. **Task - <short 3-5 word depiction>**: task description / query description / summary<br>
    **Result**: summary of the work done
```

**When NOT creating journal entry**: State explicitly "Not logging to journal: <reason>"

## Folders

### DO NOT LOOK INTO: 
- `**/@archive`: folder that has outdated and unused content
- `**/.ipynb_checkpoints`: folder that has jupyterlab checkpoint files

## Content Guidelines

### Markdown Standards
- No emojis - maintain professional, technical documentation style
- Balance concise narrative with structured bullet points
- Bullet points capture key takeaways and essential information
- Narrative focuses on value proposition, concrete benefits, and implementation details
- Include brief introductions but avoid fluff
- Explicitly state caveats and limitations where relevant
- Do not use —, use - for bullet points and hyphens in text
- Do not use full stop after a bullet point
- For mermaid diagrams use standard colours and not overloaded complex content
    - use standard colours, no custom styles
    - Use diagrams to illustrate complex processes, workflows, or architectures
    - do not overload diagrams with details, provide text narrative above or below
    - do not use images and emojis
    - only type of styling allowed: stroke and stroke-width for graph elements


## Documentation Standards
- Focus on concrete business value and technical implementation
- Include specific technology stacks and methodologies
- Maintain consistency across service descriptions
- Provide clear implementation timelines and phases
- Document success criteria and measurable outcomes

## Git Commit Standards
- Use conventional commit format: `feat(CP-0000): <description>`
- Keep descriptions concise and descriptive
- Use lowercase for commit messages
- IMPORTANT: Never attribute content creation to Claude - all content is authored by Konrad Jelen, Claude only assists with organization
- Do not include "Generated with Claude Code" or "Co-Authored-By: Claude" in commit messages
- Examples:
  - `feat(CP-0000): add context management section`
  - `feat(CP-0000): generate high-res diagrams with mermaid-cli`

## Tooling Installation

### Mermaid Diagram Generation

For generating PNG diagrams from Mermaid source files:

1. Install Mermaid CLI globally:
```bash
npm install -g @mermaid-js/mermaid-cli
```

2. Install minimal required system libraries (Ubuntu/Debian):
```bash
sudo apt-get update
sudo apt-get install -y libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 libdrm2 libgbm1 libxshmfence1
```

3. Generate diagrams with transparent backgrounds and neutral theme:

Single diagram:
```bash
echo '{"args": ["--no-sandbox", "--disable-setuid-sandbox"]}' > puppeteer-config.json
mmdc -i diagram.mmd -o diagram.png -b transparent -p puppeteer-config.json -w 2400
rm puppeteer-config.json
```

Batch conversion (all diagrams in one command):
```bash
echo '{"args": ["--no-sandbox", "--disable-setuid-sandbox"]}' > puppeteer-config.json && \
for diagram in component-integration agent-workflow framework-orchestration context-management; do \
  mmdc -i agentic-solution-components-${diagram}.mmd -o agentic-solution-components-${diagram}.png -b transparent -p puppeteer-config.json -w 2400; \
done && \
rm puppeteer-config.json
```

**IMPORTANT**:
- Always use `-p puppeteer-config.json` flag, NOT `--no-sandbox` directly (mmdc doesn't recognize it)
- Create puppeteer-config.json before running mmdc, then clean it up afterward
- All diagrams should use `%%{init: {'theme':'neutral'}}%%` in the Mermaid source
- Use `-b transparent` for transparent backgrounds
- Use `-w 2400` for high-resolution output

### Diagram Numbering Convention

All diagrams within a document should be numbered with two-digit prefixes (01, 02, 03, etc.) in the order they appear in the document. This ensures diagrams are easily referenced and maintained in proper sequence.

**Naming pattern:**
```
<document-name>-<NN>-<diagram-name>.mmd
<document-name>-<NN>-<diagram-name>.png
```

**Example:**
```
agentic-solution-components-01-success-factors.mmd
agentic-solution-components-01-success-factors.png
agentic-solution-components-02-knowledge-graph.mmd
agentic-solution-components-02-knowledge-graph.png
```

**When adding new diagrams:**
- Determine the diagram's position in the document flow
- Assign the next sequential two-digit number
- If inserting a diagram between existing ones, renumber subsequent diagrams

<!-- #endregion -->

