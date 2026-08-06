# Skill Benchmark: docx-research-report

**Model**: <model-name>
**Date**: 2026-08-06T02:56:52Z
**Evals**: 1 (1 run each per configuration)

## Summary

| Metric | With Skill | Without Skill | Delta |
|--------|------------|---------------|-------|
| Pass Rate | 100% ± 0% | 42% ± 0% | +0.58 |
| Time | 329.6s ± 0.0s | 156.6s ± 0.0s | +173.0s |
| Tokens | 77702 ± 0 | 49346 ± 0 | +28356 |

## Notes

- Non-discriminating assertions: font_is_malgun_gothic and content_is_researched_not_fabricated passed for BOTH configurations — Claude already defaults to 맑은 고딕 for Korean docx and already does careful sourced research without the skill. The skill's real value-add is structural/formatting: TOC, dedicated hyperlinked references section, consistent title/body point sizes, and one consistent main color — all 7 of these failed in the baseline.
- Single run per configuration (n=1) — no stddev signal; a true delta estimate would need 2-3 runs per config given research tasks have some run-to-run variance in topic coverage and source selection.
- with_skill took ~2.1x longer (330s vs 157s) and used ~1.6x more tokens (77.7k vs 49.3k) than the baseline, mostly from more WebSearch/WebFetch calls (13+9) and building the intermediate content.json. This is an expected time/quality tradeoff worth flagging to the user, not a bug.
- The with-skill agent hit a JSON syntax error in its first content.json draft and had to self-correct by running `python -m json.tool` before invoking build_report.py. The skill doesn't currently instruct this validation step — worth adding to reduce wasted iterations in future runs.