# Word inspection evidence

`assessments.jsonl` and `inspection-log.jsonl` each contain 144 unique fixture paths, all recorded as `proven` at this revision. These are the recorded Microsoft Word investigations from the PR; the follow-up code review did not perform new Word opens. Earlier PR descriptions saying 83 proven / 53 pending were stale.

Each assessment has a `path`, `status`, and `defects` array. Defect entries describe the isolated edit and Word's before/after verdict. Clean-open assessments deliberately have an empty `defects` array and retain the original observation under top-level `defect`, `change`, and `word` fields. An empty array means no demonstrated defect, not missing evidence; consumers must also read the clean-open observation. The inspection log records the corresponding investigations and conclusion.

The follow-up fixes recognize the implicitly bound XML `xml` prefix, share read-only DOMs within each Word-openability pass, and use native signature search for earlier ZIP end records. Repeated validation builds a fresh parse cache. A final EOCD at EOF does not allow skipping the earlier-record check: the proven concatenated-archive fixtures have exactly that layout.

Two broader questions still need targeted Word A/B evidence before changing the verdict rules:

- Which otherwise unreferenced parts without content types Word tolerates. The existing corpus also proves that some unreferenced undeclared parts trigger recovery, so weakening the rule to reachable parts alone would miss a demonstrated defect.
- How each new rule should handle unselected markup-compatibility branches across supported Word versions. Existing textbox-specific fallback handling is retained; a general branch-selection policy has not been established by these fixes.

The `xml` prefix regression test is a namespace-binding check, and the cache regression test checks repeated validation after an edit. Neither is a new claim about a fixture's Microsoft Word verdict.
