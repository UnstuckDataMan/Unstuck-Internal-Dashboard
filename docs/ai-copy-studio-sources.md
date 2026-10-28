# Copy Studio source contracts and validation

These connectors are read-only. They never create campaigns, send outreach, modify a Google Sheet, or connect to Meet Alfred. Source material is untrusted data, and importing it does not approve it for reuse.

## Configuration and access

The server needs an explicit `COPY_STUDIO_AGENCY_ID` (or `PULSE_AGENCY_ID`), `SUPABASE_URL`, and `SUPABASE_SERVICE_ROLE_KEY` to read existing mapped sources. Studio never falls back to the anonymous database credential or creates an agency during a read. A client must have an explicit matching `clients.agency_id`; legacy rows with a null agency are not silently included.

Smartlead additionally needs `SMARTLEAD_API_KEY`. Manual campaigns need the existing `GOOGLE_SHEETS_SA_JSON` and Viewer sharing for their linked spreadsheets. Copy Studio creates a separate Google `spreadsheets.readonly` session and does not change the other dashboard tools' Google credentials or scopes.

The Read AI connector reads all relevant client calls in `pulse_meetings`, populated by the existing webhook ingestion, rather than only rows flagged as onboarding. It does not assume that a new direct Read AI API connection exists. Missing calls must arrive through that existing ingestion or be supplied as an explicitly reviewed source. The `MeetAlfredProvider` boundary reports `disabled`, exposes no endpoint, and refuses sync even when a similarly named environment variable exists.

Provider configuration is reported as `configured_not_verified`, not healthy/live. A successful source job and its observation timestamp supply the separate verification evidence.

## Ingestion API

`sources.list_campaigns(client_id, source, offset=0, limit=100)` returns `campaigns`, `next_offset`, and `complete`. Only `smartlead` and `manual` sources are supported. Smartlead records must match both agency and client. Legacy manual campaigns have no agency column; their agency is inherited only after verifying the parent client's ownership. Stored job mappings are rechecked before each campaign read.

`sync_copy_bank(client_id)`, `sync_onboarding(client_id)`, and `sync_campaign(client_id, source, campaign, cursor=None, max_pages=5)` return:

```text
items: [{kind, client_id, source_key, content}]
complete: boolean
next_cursor: object or null
coverage: {complete, ...source-specific counts}
warnings: [safe actionable messages]
```

The worker saves `items` through the immutable store, then checkpoints the continuation. It checks its lease before writing source results. Campaign enumeration pages are checkpointed so insertion or deletion of a campaign cannot shift a saved index onto another campaign's cursor. Missing sheet access and unavailable sources produce partial reports while unaffected campaigns/sources continue. A completed job attempt may have `result.complete=false`; it is not a claim that every source synced. Only complete source reports carry `last_successful_sync_at`.

Smartlead statistics use bounded pages, retry/backoff for network failures, HTTP 429 and 5xx, repeated-page detection and a safety ceiling. Cursors contain hashes, stable identifiers and counters, not sent bodies or prospect details. A later incremental refresh replays historical rows because replies/statuses can change on old sends; immutable content identities prevent duplicate copy versions. No append-only source offset is assumed to capture late outcomes.

`reconcile_smartlead(client_id)` compares the current provider campaign IDs against mapped campaign IDs. It reports account/mapping counts, the number of provider campaigns without agency mappings, and this client's mapped IDs missing from the provider. It does not disclose other clients' names or silently equate imported and live campaign counts. Unmapped campaigns need a reviewed client mapping before ingestion.

## Copy versions and result attribution

Smartlead reads these existing API contracts:

- `GET /campaigns`: current campaign ID reconciliation.
- `GET /campaigns/{id}/sequences`: current `sequence_variants` with stable ID, label, subject, `email_body`, and update time.
- `GET /campaigns/{id}/statistics?offset=...&limit=...`: actual sent subject/message with `sequence_number`, `seq_variant_id`, and `email_campaign_seq_id`.
- `GET /campaigns/{id}/sequence-analytics`: provider-native counts over the requested UTC window.

Current editable templates are pending `example` records labelled `current_template_only`. Sequence-default wording without a provider variant ID is also preserved, without inventing an ID or attributing historical sends to it. Actual sent wording is stored separately as private `sent_copy` records keyed by provider + campaign + sequence ID + step + variant ID + content hash. Aggregate `performance` records reference copy source keys rather than duplicating bodies per event. A reviewer prepares and approves a separate sanitized reusable reference before sent wording can be retrieved by the writer.

Analytics are attributed to one sent version only when stable IDs agree, all source pages have completed, one observed content version exists for that variant, and its observed send count matches the provider denominator. Multiple personalized bodies, edited wording, missing IDs or incomplete counts remain explicitly unresolved. Today's template is never assumed to have been sent throughout history. The dashboard's campaign-import timestamp is not used as the start of source history.

Manual sources read the first campaign worksheet, preserving the merger's source contract. The optional `records=` argument on `google_sheets.read_ab_stats` lets Studio reuse the same aggregation as Outbound Pulse with its separate read-only source session. `Subject Line` and `Email Body` recover exact sent text without requiring Copy Bank territory/industry links. Missing bodies leave usable campaign statistics with unresolved attribution. A legacy sheet lacking Send Status is labelled `allocated_legacy`; its rate is not promoted as a verified send-based reference.

Manual outcome snapshots remain **all-time status snapshots**. Sent dates do not date later Lead/Interested outcomes. `S1/B1` is scoped to its campaign; unrelated campaigns are never pooled. Manual sent copies and performance rows use the same deterministic label identifier. If one label has different sent bodies, or outcome statuses appear on unverified/unsent rows, native reporting counts are retained but attribution remains unresolved. Subject/body pair comparisons cannot isolate subject-line effects. A `Meeting Request` is retained as a request and is never converted into a booking or held meeting.

## Performance interpretation

`performance.comparison` calls Pulse's existing `_decide` function on copied rows. It preserves send denominators, sample counts, losing arms, positive-response leaders and reply-rate fallbacks. Smartlead native positives remain separate from manual Lead + Interested. A heuristic leader is **not** a statistically proven winner; every report includes `statistical_significance: not_tested` and `proven_winner: false`.

Reference eligibility additionally requires a positive leader, verified sends, exact non-empty sent-copy attribution, complete fresh source coverage, and a non-test/non-draft campaign. `refresh_evidence(report)` applies freshness at read time; observations over seven days old, missing timestamps and future timestamps cannot retain current-reference eligibility. Historical observations and their original winner decisions remain immutable.

Relevance is evaluated within channel, client, audience, offer and step. Audience, deliverability, offer and timing remain possible confounders. Losing arms are retained for evaluation rather than constructing a winner-only corpus.

## Copy Bank and onboarding

Copy Studio reads the same canonical `clients` rows as Client Profiles. `client_profile(client_id)` enforces the explicit agency/client mapping, and the pure `profile_metadata(row)` helper provides the same shape for local fixtures. A website link can be derived from a normalized Client Profiles domain, but is labelled inferred and unverified; sending inboxes never establish an official website. Unsafe URLs, local/IP destinations and personal-email domains are not offered as client websites. `profile_fingerprint` hashes client ID, domains and profile update time so the application can detect a profile change requiring context refresh.

Copy Bank reads only exact client-prefixed keys, with no fallback to another client or bare Biz Dev keys. Email/flyout variations become separate alternative examples, not a concatenated sequence. LinkedIn and chaser arrays retain their sequence order. Known LinkedIn step type aliases are normalized to `connection`, `initial`, `follow_up`, and `closing`; an unknown type is flagged for review. Positions in mutable Copy Bank arrays are not historical attribution IDs.

LinkedIn examples originate only from Copy Bank and have no inferred performance label. Client onboarding facts and constraints can ground LinkedIn writing, but email metrics never prove LinkedIn wording performs.

The backwards-compatible `sync_onboarding` job now imports every eligible explicitly mapped client call, including later strategy/update calls, with every stored transcript block and its original index/timestamps. It does not use the older summary/truncated prompt tail. It also considers unassigned calls only when participant domains match this client uniquely across a fully paginated agency Client Profiles directory. An incomplete directory disables unassigned recovery. Title text alone never establishes an assignment. Ignored/ambiguous calls, other clients' assignments and other agencies' records are excluded; no upstream assignment is modified.

Each transcript retains `source_assignment` provenance. Coverage reports mapped/recovered counts, missing transcripts, ambiguous candidates, missing profile domains, stored listing completeness, and the current `profile_fingerprint`. An authoritative `imported_source_keys` manifest identifies currently eligible non-empty transcripts, allowing later brief extraction to exclude stale records when mappings change. Partial directory/listing reads never claim an authoritative manifest.

The existing webhook schema cannot independently certify that the upstream transcript covers the full call, so imports explicitly retain `stored_payload_complete=true`, `source_complete=false`, and `upstream_call_completeness=unverified`. Reviewers must verify source completeness rather than silently treating all stored blocks as the full call. `account_history_verified` remains false: stored webhook records are not a claim to have fetched every call in the Read AI account.

## Validation status for this build

The source implementation was tested with offline fixtures covering rate-limit retries, stable-ID mismatches, changed wording, duplicate statistics, resumed pages, non-advancing pagination, missing sheet sharing, unequal sends, the supplied anonymized manual A/B count example, long transcript tails, separate email variations, LinkedIn-only reference rules, stale evidence and disabled Meet Alfred. Worker tests cover partial-source isolation, checkpoints and loss of a job lease.

The research brief reports successful prior Smartlead read-only probes. Initial implementation used offline fixtures because credentials were unavailable. After environment credentials became available on 28 October 2026, bounded read-only production probes verified the canonical Client Profiles schema/lookup and stored Read AI meeting schema. A 25-profile sample had two configured canonical domains; those produce inferred links, not independently verified websites. Three sampled stored calls had transcripts, and none was flagged as onboarding, supporting the broader call-selection requirement. No source records were written or persisted by these probes.

Additional bounded read-only checks verified all three Smartlead source endpoint envelopes against a mapped completed campaign: two sequence steps, actual sent subject/body and matching sequence/variant IDs, plus the 1–28 October 2026 analytics window. A baseline analytic row omitted its stable variant ID and remains unresolved; fixture coverage includes this observed case. The linked Kelton manual sheet exposes all required/copy columns and reproduced the research aggregates: S1/B1 sent 138 with seven positives, and S2/B2 sent 137 with nine positives. Only counts and field names were recorded, not copy, recipients or credentials. No source data was changed.

These bounded checks do not establish full historical source coverage, compatibility with every provider payload, model output quality, or a successful staging/production rollout. Check the release verification record for additional integration tests; passing fixtures alone is not a live integration or subjective quality result.
