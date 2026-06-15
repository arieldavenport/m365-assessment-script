# Spec Kit Runbook — Building the M365 License Roadmap Web App

This is a start-to-finish guide for building the app described in
[`DESIGN.md`](./DESIGN.md) using **GitHub Spec Kit** inside **Claude Code**.
Written for someone new to the Claude Code CLI.

## What you're doing

Claude Code is Claude running in your terminal, inside a folder on your machine,
so it can read and write your files directly. Spec Kit adds a structured
`/speckit.*` workflow on top. Order of operations:

1. Install Claude Code.
2. Clone this repo locally.
3. Install Spec Kit into the repo.
4. Drive the build with the `/speckit.*` slash commands, seeding them from
   `DESIGN.md`.

> **Prefer a GUI?** The Claude Code **Desktop app** (Mac/Windows) runs the same
> slash commands without a terminal. The steps below are the terminal path.

---

## Step 1 — Open a terminal

- **Mac:** ⌘+Space → type "Terminal" → Enter.
- **Windows:** Start → type "PowerShell" → open Windows PowerShell.

## Step 2 — Install Claude Code

```bash
# Mac / Linux
curl -fsSL https://claude.ai/install.sh | bash
```
```powershell
# Windows PowerShell
irm https://claude.ai/install.ps1 | iex
```

Confirm it installed (reopen the terminal first if you get `command not found`):

```bash
claude --version
```

## Step 3 — Get the repo onto your computer

Requires **Git** (https://git-scm.com/downloads).

```bash
git clone https://github.com/arieldavenport/m365-assessment-script.git
cd m365-assessment-script
git checkout claude/vibrant-davinci-rwwmmw
```

## Step 4 — Start Claude Code and log in

```bash
claude
```

First run opens your browser to log in. Requires a **Claude Pro, Max, Team, or
Enterprise** plan (or an API/Console account); the free plan does not include
Claude Code. Once logged in you get a `>` prompt. Useful basics:

- `/help` — list commands.
- `/exit` — quit (rerun `claude` to return).
- Claude asks permission before editing files or running commands; approve each,
  or choose "always allow" for routine actions.

## Step 5 — Install Spec Kit

Spec Kit needs **Python 3.11+** and **`uv`**.

```bash
# install uv (Mac/Linux)
curl -LsSf https://astral.sh/uv/install.sh | sh
# install uv (Windows PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# install the Spec Kit CLI
uv tool install specify-cli --from git+https://github.com/github/spec-kit.git
```

From inside the `m365-assessment-script` folder, initialize Spec Kit and choose
**Claude Code** when prompted for the agent:

```bash
specify init --here --force
```

This adds a `.specify/` folder and registers the `/speckit.*` commands.
(`--here` = use this existing folder; `--force` = allow merging into a non-empty
repo.) Verify with `specify self check`.

## Step 6 — Drive the build

Start Claude Code again (`claude`) in the folder, then run these in order. The
key move: reference the existing design with `@DESIGN.md` so Claude doesn't
re-derive it.

| # | Command | Seed prompt / what to provide |
|---|---------|-------------------------------|
| 1 | `/speckit.constitution` | "Establish these principles: deterministic Python owns all math, pricing, timeline, and Excel formulas; the AI model only classifies accounts and infers frontline vs knowledge workers and writes rationale — it never computes prices or quantities. Quotes must be auditable and reproducible. No secrets in code (use managed identity). Every cost/quote cell uses live Excel formulas. Tests gate the CSV normalization, the deterministic engine, and the LLM output schema." |
| 2 | `/speckit.specify` | "Build a web app that ingests Microsoft 365 user and subscription data — either uploaded CSVs (from the m365-assessment.sh script OR the M365 admin-center export) or fetched live from Microsoft Graph — reviews it against the license-migration roadmap rules in SKILL.md, and outputs a 6-tab Excel workbook: User Inventory, Current Subscriptions, Migration Timeline, Cost Summary, License Overlaps, Quote Summary. Support both data sources and an interactive clarifying-questions step." |
| 3 | `/speckit.clarify` | Let it surface ambiguities. Key ones from the design: the purchase-channel data gap (Commercial Direct vs CSP isn't in the script CSV or Graph), the pricing source of truth (externalized, dated catalog), and the wizard questions (E3/E5 keep-or-migrate, shared-mailbox conversion, frontline segmentation, stale-day threshold). |
| 4 | `/speckit.plan` | "Use the architecture and technology stack in @DESIGN.md — FastAPI, Azure AI Foundry (gpt-4o, managed identity), openpyxl, App Service with an async job model, Bicep IaC, the `app/` directory tree, and the dual-source CSV normalization tables." |
| 5 | `/speckit.tasks` | Generate the ordered task list. Steer it to the phased build in @DESIGN.md: Phase 0 scaffolding → Phase 1 MVP (normalize → engine → Foundry classify → Excel → async job) → Phase 2 Azure deploy → Phase 3 live Graph fetch. |
| 6 | `/speckit.analyze` | Cross-artifact consistency check before coding (e.g. catch any task that lets the LLM compute prices — that violates the constitution). |
| 7 | `/speckit.implement` | Execute. Do it **one phase at a time** — verify Phase 1 produces a correct workbook from sample CSVs before deploying anything. |

Optional Spec Kit commands: `/speckit.checklist` (custom quality checklist —
useful for the Excel formatting spec) and `/speckit.taskstoissues` (push tasks
into GitHub issues on this repo).

## Tips

- You can just *tell* Claude Code in plain English to run these setup steps
  (e.g. "install uv and set up Spec Kit"); it will run them and ask permission.
  The manual commands above are so you understand what's happening.
- Commit after each phase: tell Claude "commit and push to
  claude/vibrant-davinci-rwwmmw".
- `DESIGN.md` maps closely onto Spec Kit's artifacts: the constitution =
  its principles, the body = `/speckit.plan`, the risks = `/speckit.clarify`,
  and the phased build = `/speckit.tasks`.
