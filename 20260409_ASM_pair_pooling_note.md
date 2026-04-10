# ASM: Pair Pooling Diagnostic — Internal Note

**Date:** 2026-04-09  
**Context:** ASM unified ANOVA script (`260114_ANOVA_ASM_unified.r`)  
**Author:** Anezka / Claude working note

---

## The situation

In the ASM experiment, each of the 3 solutions (water_1/2/3 for SNC, Lactose/Stannum/Silicea for Verum) is produced in duplicate — two independently prepared bottles from the same source (28X). These are labelled as **pair 1** and **pair 2**. Each pair contributes 7 crystallization plates per experiment, giving 14 plates per solution per experiment.

The current ANOVA model is:

```
parameter ~ decoded_solution * experiment_number
```

where `decoded_solution` has **3 levels** (not 6). The two pairs within each solution are pooled: pair 1 and pair 2 plates are treated as 14 interchangeable replicates within the same group. The `pair` variable exists in the data but is not entered into the model.

## The question

Is this pooling justified? If the production step introduces systematic differences between the two bottles, then the 14 plates are not truly independent — 7 come from one bottle, 7 from another. This would mean:

- The effective sample size per cell is closer to 2 (pair means) than 14 (individual plates)
- Standard errors from the current model may be too small
- Significant solution effects could be partially driven by pair-level differences rather than true solution differences

## The problem with "pair" as a factor

The pair label (1 or 2) is arbitrary — it just means "first production run" and "second production run" of that solution. Crucially, **pair 1 of water_1 has nothing in common with pair 1 of water_2**. They are different bottles, produced at different times, from different source solutions. The number "1" carries no shared identity across solutions.

This rules out treating `pair` as a simple crossed factor in the ANOVA (e.g., `decoded_solution * pair * experiment_number`), because the `pair` main effect — "is pair 1 different from pair 2 on average" — would be comparing things that have no reason to be comparable. The `decoded_solution × pair` interaction would test something more sensible (do pairs differ within each solution?), but the main effect would be uninterpretable.

## Options considered

### Option 1: Three-way ANOVA with pair crossed (rejected)

```
parameter ~ base_solution * pair * experiment_number
```

**Problem:** The `pair` main effect and `pair × experiment_number` interaction are meaningless (pair 1 across different solutions is not a coherent group). Only the `base_solution:pair` interaction would be informative. The other terms waste df on uninterpretable comparisons.

**Verdict:** Structurally wrong for this design.

### Option 2: Nested fixed effect (chosen for internal diagnostic)

```
parameter ~ decoded_solution / pair + experiment_number + decoded_solution:experiment_number
```

The `/` operator nests `pair` within `decoded_solution`. This estimates a specific pair 1 vs. pair 2 shift *within each solution separately* — so you get three tests: does pair 1 ≠ pair 2 for water_1? for water_2? for water_3?

**Pros:**

- Stays within the `aov()` / Type III framework already used in the script — no new packages or model types
- Gives per-solution pair differences you can inspect individually — if one solution drives the effect, you see it directly
- Easy to explain; familiar ANOVA output
- Good enough as a screening check: if all three nested terms are non-significant, pooling is justified

**Cons:**

- Burns 3 df on specific shift estimates that are not the primary interest (you care about *whether* pairs matter, not *which way* each pair deviates)
- No single summary statistic for "do pairs matter overall" — you get three p-values and must synthesize
- Does not adjust the standard errors of the `decoded_solution` F-test. Even after absorbing the pair mean shifts, the residual treats remaining plates as independent within each pair. If pair-level correlation is real, the solution test is still slightly anticonservative
- Treats these specific pairs as fixed — doesn't generalize to "what if we produced new pairs?"

### Option 3: Random effect / mixed model (better for formal reporting)

```
parameter ~ decoded_solution * experiment_number + (1 | decoded_solution:pair)
```

Using `lmer()` from `lme4`/`lmerTest`. This adds a random intercept for each specific preparation (e.g., water_1 pair 1, water_1 pair 2, etc.), estimating how much variance the production step contributes.

**Pros:**

- Gives one clean summary: the pair-level **variance component** and its derived **ICC** (intraclass correlation = pair variance / total within-cell variance)
- ICC near zero → pairs are interchangeable → pooling is fine
- ICC substantial → production step matters → SEs should be adjusted
- Automatically adjusts denominator df for the fixed effects — if pair variance is large, the solution F-test becomes appropriately more conservative
- Treats pairs as a sample from a population of possible production runs — matches the conceptual question ("would new pairs also vary?")

**Cons:**

- Introduces mixed model machinery (`lmer`, Satterthwaite df) which is a different framework than the rest of the script
- Harder to explain to reviewers unfamiliar with variance components
- Lumps all solutions into one variance estimate — could mask if only one solution has a pair problem
- For a quick internal check, it's more infrastructure than needed

### Option 4: Diagnostic plots (complementary to any model)

- **Pair agreement scatter:** For each solution × experiment cell, plot mean(pair 1) vs. mean(pair 2). Points near the identity line → pairs agree → pooling is fine.
- **Pair difference plot:** Plot (pair 1 mean − pair 2 mean) across experiments for each solution. Centered on zero with no trend → no systematic pair effect.
- **Already partially visible** in the existing 6-line plots (pairs as separate lines per solution), but formal diagnostics are clearer.

## Decision

**For now: Option 2 (nested fixed effect) as an internal screening check, supplemented by Option 4 (pair diagnostic plots).**

If the nested pair terms are non-significant across parameters and the diagnostic plots show no structure, we document that pooling is justified and move on. If pair effects emerge, we escalate to Option 3 (mixed model) for formal adjustment.

This keeps the analysis within the current `aov()` framework and avoids introducing mixed models unless they're needed.

## Implementation note

The `base_solution` variable needed for this analysis already exists in the plotting section of the unified script (created around line 1010). For the nested ANOVA, we can use the existing `decoded_solution` and `pair` columns directly — no new variables needed. The nested term `decoded_solution/pair` in the model formula handles the nesting automatically.
