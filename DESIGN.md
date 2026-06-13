# M365 License Roadmap Web App — Design

## Context

Two assets exist today and are used as separate manual steps:

1. **`m365-assessment.sh`** — a Bash script (built for Azure Cloud Shell) that
   reuses the Azure CLI access token, calls Microsoft Graph
   (`/organization`, `/subscribedSkus`, `/directory/subscriptions`, `/users`),
   maps license GUIDs to friendly names, and exports three CSVs
   (`M365_Users`, `M365_Products`, `M365_StaleUsers`).
2. **The `m365-license-roadmap` skill** — an LLM skill that ingests M365
   admin-center CSV exports and produces a 6-tab Excel license-migration
   roadmap/quote (300-seat-cap strategy, account categorization, renewal/channel
   timeline, current-vs-future cost, mandatory Centre Technologies services).

This document describes an **Azure-hosted web app** that fuses both: ingest the
CSV files, review them with an **Azure AI Foundry** setup, and output the Excel
workbook — replacing the two-step manual workflow with one app.

### Decisions locked in
- **Data source:** both upload-CSV *and* live Graph fetch, delivered in phases.
- **AI role:** **hybrid** — deterministic code owns the math; the Foundry model
  owns fuzzy judgment only.
- **Foundry:** the plan includes standing up the Foundry project + model
  deployment.

## Core idea

```
Upload CSVs ─┐
             ├─► Normalize ─► Deterministic engine ─► Excel (6 tabs)
Live Graph ──┘        │            ▲
fetch                 └─► Azure AI Foundry (classification / judgment)
```

The split that keeps quotes auditable and reproducible:

- **Deterministic Python owns anything with a defined right answer** — Step 0
  licensed-user count + 300-seat-cap path selection, renewal/channel timeline
  math, pricing catalog, overlap quantification, and all Excel formulas.
- **The Foundry model owns only fuzzy judgment** — per-account classification
  (User / Service Account / Shared-Functional / Unused), frontline-vs-knowledge
  inference from titles/departments, and narrative rationale.

The LLM returns **structured JSON**; the deterministic layer cross-checks every
SKU pick against the allowed-SKU enum and the tenant-size path, overriding or
flagging out-of-policy picks. The LLM never chooses prices or quantities.

## High-level architecture

```
                          ┌─────────────────────────────────────────┐
                          │            Browser (SPA wizard)           │
                          │ upload / sign in / answer clarifying Qs /  │
                          │ poll status / download xlsx                │
                          └───────────────┬───────────────────────────┘
                                          │ HTTPS (REST + poll)
                                          ▼
   ┌──────────────────────────────────────────────────────────────────────┐
   │                  FastAPI app  (App Service or Container Apps)          │
   │   /api/jobs (create)  /api/jobs/{id} (status)  /api/jobs/{id}/answers  │
   │   /api/jobs/{id}/result (xlsx)   /api/graph/fetch (optional live)      │
   │                                                                        │
   │   ingest/normalize ─► engine (deterministic) ─► excel_builder(openpyxl)│
   │                              │                                         │
   │                              ▼                                         │
   │                       foundry_client ── azure-ai-projects/agents       │
   └──────────────────────────────────────────────────────────────────────┘
                │ Managed Identity (no secrets in code)
        ┌───────┴────────┐  ┌──────────────┐  ┌──────────────────────────┐
        │ Azure Blob      │  │ Azure Key    │  │ Azure AI Foundry project │
        │ (uploads,       │  │ Vault        │  │  + model deployment      │
        │  results, state)│  │ (opt secret) │  │  (gpt-4o / o-series)     │
        └─────────────────┘  └──────────────┘  └──────────────────────────┘
        ┌───────────────────────────────────┐
        │ Microsoft Graph (optional live)    │  delegated OAuth (Phase 3)
        └───────────────────────────────────┘
   Async layer: Storage Queue + worker (in-process asyncio for MVP;
   Container Apps Job later). Job state in Blob/Table.
   Auth (Phase 3): Entra ID via App Service EasyAuth, reused for delegated Graph.
```

**Why async:** the AI review iterates over potentially thousands of user rows;
a >300-user tenant can take minutes and would exceed App Service's 230s request
timeout. Jobs are created with `202 + jobId`; the frontend polls for status.

## Repo structure to add

The existing script stays at the repo root (Cloud Shell entry point + reference
implementation). New code lives under `app/`:

```
app/
├── pyproject.toml                 # deps (below)
├── Dockerfile                     # Container Apps path
├── azure/
│   ├── main.bicep                 # all resources
│   └── parameters.json
├── backend/
│   ├── main.py                    # FastAPI app + router wiring
│   ├── config.py                  # pydantic-settings; env / KV refs
│   ├── api/{jobs.py, graph.py}
│   ├── ingest/
│   │   ├── canonical.py           # canonical pydantic models (User, Product)
│   │   ├── normalize.py           # dual-source → canonical mapping (lynchpin)
│   │   ├── detect_source.py       # sniff script-CSV vs admin-center-CSV
│   │   └── sku_friendly.py        # port of script SKU friendly-name map
│   ├── engine/
│   │   ├── deterministic.py       # 300-cap math, pricing, phase grouping
│   │   ├── renewal_math.py        # port of months_until / days_since
│   │   ├── decision_matrix.py     # target-SKU rules, overlap detection
│   │   └── pricing.py             # MS list prices + Centre services catalog
│   ├── ai/
│   │   ├── foundry_client.py      # azure-ai-projects wrapper
│   │   ├── prompts.py             # ported SKILL.md instructions (system prompt)
│   │   ├── schemas.py             # JSON Schemas for structured output
│   │   └── classify.py            # batched, concurrent LLM orchestration
│   ├── excel/
│   │   ├── builder.py             # 6-tab workbook orchestrator
│   │   ├── styles.py              # Arial, colors, fills
│   │   ├── tabs/                  # one module per tab (1..6)
│   │   └── formulas.py            # formula strings + recalc handling
│   ├── graph/
│   │   ├── client.py              # httpx Graph client + pagination
│   │   ├── auth.py                # delegated (MSAL) vs app-perm token paths
│   │   └── extract.py             # /organization,/subscribedSkus,/users,...
│   ├── jobs/{store.py, worker.py, models.py}
│   └── storage/blob.py
├── frontend/                      # minimal SPA wizard (React+Vite or HTMX)
└── tests/                         # fixtures (both CSV sources) + golden xlsx
```

**Python packages:** `fastapi`, `uvicorn[standard]`, `pydantic>=2`,
`pydantic-settings`, `pandas`, `openpyxl`, `httpx`, `azure-ai-projects`,
`azure-ai-agents` (or `openai` SDK pointed at the Foundry endpoint),
`azure-identity`, `azure-storage-blob`, `azure-data-tables`, `msal` (Phase 3),
`python-multipart`. Dev: `pytest`, `pytest-asyncio`, `respx`, `ruff`.

## CSV normalization layer

A **canonical schema** decoupled from both sources. `detect_source.py` sniffs
the header row (`UserPrincipalName`+`DaysSinceLastActivity` → script source;
`User principal name`+`Block credential` → admin-center source).

**Users mapping**

| Canonical | Script CSV | Admin-center CSV |
|---|---|---|
| `display_name` | `DisplayName` | `Display name` |
| `upn` | `UserPrincipalName` | `User principal name` |
| `title` | `JobTitle` | `Title` |
| `department` | `Department` | `Department` |
| `office` | `OfficeLocation` | `Office` |
| `licenses` | `AssignedLicenses` (split `;`) | `Licenses` (split `+`/`;`) |
| `account_enabled` | `AccountEnabled` | derived from `Block credential` (inverted) |
| `days_since_activity` | `DaysSinceLastActivity` | absent → `None` |

**Products mapping**

| Canonical | Script CSV | Admin-center CSV |
|---|---|---|
| `product_name` | `ProductName` | `Product name` |
| `purchased_qty` | `SkuTotalLicenses` | `Purchased quantity` |
| `assigned_qty` | `SkuConsumedLicenses` | `Assigned licenses` |
| `available_qty` | `SkuAvailableLicenses` | `Available licenses` |
| `renewal_date` | `NextRenewalDate` | `Renewal or expiration date` |
| `months_until_renewal` | `MonthsUntilNextRenewal` | compute via `renewal_math` |
| `subscription_status` | `SubscriptionStatus` | `Subscription status` |
| `purchase_channel` | **absent → `Unknown`** (flag) | `Purchase channel` |

Notes: normalize license names via a small alias table; Step 0 licensed-user
count uses canonical `license_count > 0` (works for both sources); missing
`purchase_channel` on script source defaults to the conservative term-locked
assumption and/or a clarifying question.

## Deterministic engine vs Foundry LLM

**Deterministic (`engine/`)**
- Step 0 licensed-user count → ≤300 (Business Premium) vs >300 (E3/E5/F3/F1/EXO P1) path.
- Renewal/channel timeline: Commercial Direct = term-locked (schedule at renewal);
  CSP/Reseller = change anytime (Phase 0); mid-term upgrade rules. Port the
  script's `months_until` math.
- Pricing catalog (data, not LLM output): MS list prices per target SKU +
  mandatory Centre Technologies services (Cloud Backup, Value Add, DMARC, service
  account). Current vs post-migration cost.
- Overlap quantification where it is a rule (Teams Essentials redundancy, Audio
  Conferencing add-ons, Copilot on service accounts, unassigned seats).

**Foundry LLM (`ai/`)** — only:
- Per-account classification: User / Service Account / Shared-Functional / Unused.
- Frontline vs knowledge-worker inference from title/department.
- Narrative rationale + overlap narrative text.

**Structured output:** call the model deployment with a JSON Schema
(`ai/schemas.py`) returning per-user `{account_type, worker_class,
recommended_target_sku, confidence, rationale}`. Temperature 0, strict schema;
validate every response with pydantic; on failure retry once then fall back to a
deterministic keyword heuristic so a job never hard-fails. Batch users
~50–100/call, run chunks concurrently with a semaphore.

**Porting the skill:** SKILL.md prose becomes the system prompt for the
*classification + decision-matrix* rules only. The math (Step 0 count, Step 3
timeline, Step 4 workbook) is removed from the prompt and implemented in code;
the prompt explicitly tells the model to classify and explain only, never to
compute prices or totals.

**Clarifying questions (Step 1):** modeled as a two-phase job — after
ingest/normalize the engine computes which questions are needed (E3/E5
keep-or-migrate, shared-mailbox conversion, output format, frontline
segmentation, optional supplemental reports for >300, missing purchase channel).
Job enters `AWAITING_ANSWERS`; the SPA renders them; `POST /answers` resumes.
Questions are deterministic (derived from data), not LLM-generated.

## Excel generation (replacing the external xlsx skill + recalc.py)

Built directly with **openpyxl** — one module per tab + shared `styles.py`
(Arial, header fills, color-coding constants, `freeze_panes="A2"`,
`auto_filter`, column autosize).

1. **User Inventory** — one row/user: canonical fields + LLM classification +
   rationale; conditional fill by recommended action.
2. **Current Subscriptions** — qty/assigned/available/renewal/channel/status.
3. **Migration Timeline** — Phase 0 + renewal-grouped phases from the engine.
4. **Cost Summary** — current vs post-migration with **live `=qty*price` and
   `=SUM(...)` formulas** (not hardcoded values).
5. **License Overlaps** — quantified findings + LLM narrative.
6. **Quote Summary** — MS license line items + mandatory Centre Technologies
   services + Centre service account; grand total via `=SUM`.

**Formula recalculation** (what `recalc.py` did): default to
`wb.calculation.fullCalcOnLoad = True` so Excel/LibreOffice/Sheets recompute on
open. If a numeric value is needed server-side (e.g. PDF render), round-trip via
**LibreOffice headless** in the container. Avoid `formulas`/`pycel`.

## Microsoft Graph integration — auth options

The Cloud Shell `az account get-access-token` approach cannot work in a hosted
app. Port the Graph REST logic (`get_all_pages` pagination, v1.0→beta
subscription fallback, `signInActivity` graceful degradation) into `graph/` with
`httpx`, and pick an auth model:

- **Option A — Upload-only (MVP).** No Graph permissions, no consent friction.
- **Option B — Delegated OAuth (target).** MSAL auth-code; user consents to
  `User.Read.All`, `Organization.Read.All`, `Directory.Read.All`,
  `AuditLog.Read.All`. Mirrors the script's delegated model; reuses the EasyAuth
  identity.
- **Option C — App registration (app-only).** Unattended/scheduled runs; biggest
  security surface; only if automation is required.

**Channel gap:** `Purchase channel` (Commercial Direct vs CSP) is not reliably
exposed by Graph — so even with live fetch, channel still likely needs the
admin-center export or a clarifying question.

## Azure infrastructure & deployment

- **Compute:** App Service (Linux/Python) for MVP — built-in EasyAuth doubles as
  the Phase 3 delegated-Graph identity; managed identity; deployment slots. Move
  the worker to a **Container Apps Job + Storage Queue** as load grows. The
  `Dockerfile` supports both.
- **AI Foundry:** Foundry project (hub + project) + model deployment (gpt-4o for
  cost/speed, o-series if classification quality needs it). `foundry_client.py`
  authenticates via `DefaultAzureCredential` (managed identity granted
  `Azure AI Developer` / `Cognitive Services OpenAI User`) — **no API keys**.
- **Secrets/identity:** system-assigned managed identity + RBAC. Key Vault only
  for a Graph app secret if Option C is used.
- **Storage:** one account with `uploads/`, `results/` blob containers, a Table
  for job state, and a `jobs` queue. Lifecycle rule auto-deletes inputs/results
  after N days (PII hygiene).
- **IaC:** `azure/main.bicep` provisions all of the above + role assignments.

## Phased build sequence

- **Phase 0 — Scaffolding.** `app/` tree, `pyproject.toml`, FastAPI skeleton,
  canonical models, test fixtures (sample CSVs for both sources).
- **Phase 1 — MVP (local-runnable): upload → AI review → Excel.** Normalization,
  deterministic engine, Foundry classification + fallback, 6-tab Excel; then wrap
  in the async job model + clarifying-questions pause/resume; minimal SPA.
  *Exit criterion: upload sample CSVs from either source → correct 6-tab xlsx.*
- **Phase 2 — Deploy to Azure.** Bicep, managed identity, Foundry deployment,
  Blob/Queue/Table state, App Service + EasyAuth.
- **Phase 3 — Live Graph fetch.** Port pagination/extract, wire delegated OAuth
  (Option B), generate canonical dataset straight from Graph; upload stays as
  fallback.
- **Phase 4 — Polish/scale.** Container Apps Job worker, LibreOffice recalc path
  if needed, supplemental reports for >300, PDF export, multi-tenant (Option C).

## Verification & testing

- **Normalization:** both CSV sources → identical canonical output; `account_enabled`
  from `Block credential`; license split for both delimiters; Step 0 count.
- **Deterministic:** golden tests on `months_until`/`days_since` vs bash output;
  300-cap path at 299/300/301; pricing totals; phase grouping with mixed channels.
- **LLM contract:** mock Foundry; every response validates against the schema;
  out-of-policy SKU overridden; fallback engages on malformed output. A small
  hand-labeled eval set tracks classification accuracy (metric, not hard gate).
- **Excel:** open generated workbook; assert 6 sheets, headers, frozen panes,
  auto_filter, Arial, fills, that cost/quote cells are formula strings (`=…`),
  and `fullCalcOnLoad` set. Optional LibreOffice convert in CI to confirm no
  `#REF!`/`#NAME?`.
- **Graph (Phase 3):** `respx`-mock pagination, v1.0→beta fallback, and the
  `signInActivity` 403 degradation.
- **E2E:** largest fixture through the async worker; assert state transitions
  `QUEUED→AWAITING_ANSWERS→RUNNING→DONE` and a downloadable xlsx.

## Key risks & open questions

1. **Purchase-channel data gap (highest correctness impact).** Neither the
   script CSV nor Graph reliably exposes Commercial Direct vs CSP, which drives
   the timeline. *Mitigation:* require the admin-center subscriptions export, or
   ask per subscription, defaulting to term-locked.
2. **Pricing source of truth.** MS + Centre rates change. Externalize the catalog
   with an effective-date stamp shown on the Quote tab.
3. **LLM classification accuracy.** Frontline-vs-knowledge inference is fuzzy and
   changes the quote. *Mitigation:* temp 0 + confidence scores + rationale +
   deterministic fallback + flag low-confidence rows for human review.
4. **Formula recalc fidelity.** `fullCalcOnLoad` works on open but a programmatic
   reader sees blanks until recalc; adopt the LibreOffice round-trip if values are
   needed server-side.
5. **PII & data residency.** Short Blob TTL, no row-content logging,
   managed-identity-only access, send the model minimal fields
   (title/department/license/activity, not the full directory record).
6. **Graph consent friction (Phase 3).** Port the script's graceful degradation
   so missing `AuditLog.Read.All` / `Directory.Read.All` never hard-fails.
7. **App Service 230s timeout.** Mitigated by the async job model — fix the
   `202 + poll` contract from Phase 1.
8. **"Unused" classification.** Feed the script's stale definition (enabled,
   non-guest, `DaysSinceLastActivity ≥ threshold`, or never-signed-in-and-old) to
   the LLM as the deterministic input; keep the threshold (default 90)
   user-configurable in the wizard.
