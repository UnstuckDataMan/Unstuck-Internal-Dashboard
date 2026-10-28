# Choosing the Copy Studio writing provider

Choose using blind reviews of the same client briefs, approved references and writing instructions. Factual accuracy and editing effort come before price. Neither provider is declared the better writer by this implementation, and passing automated tests does not establish subjective writing quality.

The evaluator prepares a private review pack from supplied material. **It makes no API calls by default**, creates no example client data and never supplies missing model outputs. Paid generation requires both `--generate` and `--allow-paid-calls`, plus explicit model selections and company server credentials. All opted-in writing and independent-review calls use the same tool cost ledger as dashboard jobs.

## Curate the held-out set

Start with 5–10 varied pilot clients and work toward 50–100 real briefs. Include email and LinkedIn, differing sectors and geographies, missing proof, audience exclusions, sensitive client names, conflicting onboarding statements, long transcripts and stale performance examples. Keep a smaller synthetic fixture set for regression tests; mark it separately and do not count it as a real quality benchmark.

Use immutable approved brief and example records exported from Copy Studio. Keep each client's verified facts, complete constraints and evidence citations. A brief with unresolved contradictions or missing audience/offer facts must be corrected before model comparison. LinkedIn needs an approved LinkedIn Copy Bank example after leakage filtering. No email-winner evidence or Meet Alfred metrics are introduced.

Give every case and reference stable lineage tags:

- `client_family`: the same identity for a client and its aliases.
- `campaign_family`: the same identity for a campaign and all related copy revisions, steps and variants. Changing a variant ID or saving a new copy version must not create a new family.

Default `--holdout-level campaign_family` excludes references associated with **any** held-out campaign family in the input pack. `client_family` performs a stricter client-held-out comparison. Client-held-out comparisons may need explicitly approved, sanitized agency examples, since ordinary examples from other clients remain prohibited by the live writer. Client facts still come from the selected client's approved brief in either mode.

References with missing lineage are excluded. Exact long copy passages matching a held-out human baseline are also excluded even when their lineage claims otherwise. The manifest records each exclusion. These checks rely on accurate curator-supplied lineage; they do not discover every paraphrased copy relationship automatically. Review the filtered corpus before treating results as held out.

## Input format

Each UTF-8 JSONL line contains one case with these fields:

| Field | Meaning |
|---|---|
| `case_id` | Stable unique case identifier |
| `client_id` | Authorized client identifier matching the approved brief |
| `origin` | `real` or `synthetic_fixture` |
| `lineage` | Object containing `client_family` and `campaign_family` |
| `brief` | Full immutable approved `brief` record, including `id`, `client_id`, `kind`, `status`, `content_hash`, and `content` |
| `request` | Live writer request: `channel`, `mode: "draft"`, audience, angle count, objective and optional channel limits |
| `examples` | Approved immutable example records, each with its own top-level `lineage` |
| `methodology` | Approved methodology records; empty means methodology review remains pending |
| `baseline` | Optional current human draft, using the writer's `{"angles": [...]}` structure; never passed to a model |
| `outputs` | Optional previously generated candidates: `[{"candidate_id": "unique label", "provider": "...", "model": "...", "draft": {...}}]` |

The draft structure is defined by `DRAFT_SCHEMA` in `app/utils/copy_studio/writer.py`. Each angle has a name, rationale, subject list, sequence steps, evidence IDs and claims. Step types and limits are channel-specific. Export existing records and drafts rather than inventing case material to fill this structure.

An approved record is not proof that its `origin: real` tag is authentic. Source authenticity, permission to use claims, transcript completeness and lineage are curator/reviewer responsibilities; the manifest states this limitation explicitly.

## Prepare a pack without model calls

```bash
.venv/bin/python scripts/copy_studio_evaluate.py \
  --input /secure/copy-evaluation/cases.jsonl \
  --out-dir /secure/copy-evaluation/review-001
```

The script refuses to overwrite an existing pack. Keep outputs outside the repository. It writes private files (mode 0600 in a mode 0700 directory):

- `blind-review.csv`: randomly ordered cases and candidates, with copy, audience, approved facts, evidence links, constraints and completeness. Provider/model labels and draft provenance are omitted. Scores start blank.
- `answer-key.private.json`: mappings back to provider/model or current human baseline, source versions, prompt hashes, automated review and evaluation cost-ledger IDs. Do not give this file to blind reviewers.
- `manifest.json`: real/fixture counts, held-out filtering, missing outputs and generation status. Keep it from reviewers if generation status reveals provider identities.
- `ratings-summary.json`: initially `pending`, with no fabricated scores, pass or winning model.

If no outputs or baseline are supplied, the pack truthfully contains no candidate copy. If actual outreach text identifies a model/provider, preparation stops rather than silently rewriting that candidate and biasing the comparison. A fixed `--seed` can reproduce ordering for a controlled run; keep it with the private answer key.

## Optional paid comparison

Set company API credentials securely on the machine running the evaluation; never place them in case files or command arguments. Configure Copy Studio storage and an agency ID, as for the dashboard, so API usage is recorded. An explicit local SQLite ledger is available for development; no unmetered fallback occurs when storage is missing.

```bash
.venv/bin/python scripts/copy_studio_evaluate.py \
  --input /secure/copy-evaluation/cases.jsonl \
  --out-dir /secure/copy-evaluation/review-002 \
  --generate --allow-paid-calls \
  --model 'openai=<company-approved-model-id>' \
  --model 'anthropic=<company-approved-model-id>' \
  --actor '<operator-identity>'
```

Use currently available company-approved model IDs; this harness does not select an untested model as the quality default. OpenAI prefers `COPY_STUDIO_OPENAI_API_KEY`, with the existing `OPENAI_API_KEY` as a fallback. Anthropic prefers `COPY_STUDIO_ANTHROPIC_API_KEY`, with the existing `ANTHROPIC_API_KEY` as a fallback. The tool-specific credential takes precedence when both are set, matching the dashboard provider configuration. Neither a personal ChatGPT/Claude subscription nor a custom GPT imports knowledge into these API calls.

Each provider runs the same live writing and independent-review pipeline with the same frozen brief, filtered examples, methodology, audience, angle count and limits. The harness verifies that the writing-prompt hashes match before admitting the comparison. Independent critiques naturally contain each provider's own draft. It records failures explicitly and generates no substitute copy.

Each case/provider run has an `evaluation_job_id` in the private answer key and tool cost ledger. Writing, critique, retries and charged/incomplete responses are metered by the provider boundary. Missing token usage or pricing stays unknown; supplied historical outputs get no invented cost. Compare cost **after** quality gates and human review. Runs do not send campaigns or approve any generated copy.

## Blind review and decision

Review against the exact client facts and constraints, without the answer key. Score each dimension 1–5:

1. Factual accuracy.
2. Audience relevance.
3. Unstuck/client voice.
4. Specificity and commercial usefulness.
5. Channel fit.
6. Strategic variety across the requested angles.

Record actual editing minutes, preference rank within each case, reviewer identity and notes. Mark the four release gates `yes` or `no`: claims supported, no client leakage, exclusions respected, valid variables/format. Reviewers must check evidence meaning and permissions, not merely whether a citation ID is present. Missing context or incomplete evidence should be resolved, not scored as if verified.

Use a second independent reviewer on a representative subset or the full set. Prepare a separate pack or retain each reviewer's original CSV; the scorer accepts one rating per candidate in a run and does not pretend correlated duplicate ratings are separate independent briefs. Investigate disagreements and examine client/channel slices before averaging.

After ratings are completed:

```bash
.venv/bin/python scripts/copy_studio_evaluate.py \
  --out-dir /secure/copy-evaluation/review-002 \
  --ratings /secure/copy-evaluation/completed-review.csv
```

The scorer writes a new timestamped summary. Blank and incomplete ratings remain pending; a failed factual or scope gate remains visible even if stylistic scores are high. It does not select an automatic winner, declare statistical significance, approve copy or deploy a model.

Choose the provider that passes factual/scope gates and earns better paired blind preference with less editing effort across the varied real sample. If results are close, use cost, latency and operational reliability as tie-breakers, or collect more paired cases. Confirm improved reviewer preference/editing effort over the existing human workflow before broad client access. Re-run evaluation after material model, prompt, retrieval or corpus changes.

## Current validation boundary

Automated fixture tests cover the paid-call guard, group leakage filters, client/channel isolation, identical provider writing context, blinded output, private answer keys, failure handling, pending scores and cost-ledger integration. They do not constitute live model trials, a collected 50–100-brief corpus, real campaign-performance evidence or human quality approval.
