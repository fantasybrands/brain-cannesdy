# Brain Cannesdy — Phase 3: Rewrite All Entry Descriptions

Paste this into a fresh Claude session inside the Brain Cannesdy Website project.

---

## Context

Brain Cannesdy is a database of Cannes Lions 2025 award-winning work. The data lives in two places:

- **`Brain_Cannesdy_Brand_Experience_2025.xlsx`** — source of truth for entry data (Tier, Campaign, Brand, Agency, Country, Subcategory, Image URL, YouTube ID, Entry URL).
- **`index.html`** — the live site. Contains a `const DATA = { ... }` JavaScript object with 872 entries across 32 category sheets. Each entry has a `desc:` field.

The `desc:` field is currently unreliable. About 558 of them are hand-authored but inaccurate, and the rest are auto-generated templates like:

> "Grand Prix winner in B01 - Use of Music. One Second Ads for Budweiser by Africa Creative DDB (Brazil)."

These templates just restate the metadata. **Phase 3 replaces every desc with a 1–2 sentence accurate, specific description of what the campaign actually is and why it won.**

## What's already done (do not redo)

- All 146 high-resolution local images are in `digital-presentation-images/` (WebP q80, 2048px max). Use the originals at `digital-presentation-images/originals/` for visual reference when writing descs — they preserve full source quality and include small text/copy.
- The xlsx Image URL column is populated with relative paths to the local WebP files for those 146 entries. The remaining ~726 entries use filespin URLs that are now paywalled.
- 133 entries in `index.html` already point at the local WebP files. The descs themselves are still wrong on all of them.

## Goal

Rewrite the `desc:` field for every entry in `index.html`. End state: each entry's description tells a creative reader, in one or two sentences, the actual idea behind the campaign — concept, insight, or twist — not just the metadata.

## Recommended approach

### 1. Add a Description column to the xlsx

Make the xlsx the source of truth for descriptions going forward, so this work is resumable across sessions and reviewable in a spreadsheet view.

Add a column called `Description` after `Entry URL` in every sheet. Leave it blank initially. Don't touch any other column.

```python
import openpyxl
wb = openpyxl.load_workbook('Brain_Cannesdy_Brand_Experience_2025.xlsx')
for sheet in wb.sheetnames:
    if sheet == 'Summary': continue
    ws = wb[sheet]
    headers = [c.value for c in ws[1]]
    if 'Description' not in headers:
        col = len(headers) + 1
        ws.cell(row=1, column=col, value='Description')
wb.save('Brain_Cannesdy_Brand_Experience_2025.xlsx')
```

Backup before writing: copy the xlsx to `Brain_Cannesdy_Brand_Experience_2025.bak_before_phase3.xlsx`.

### 2. Process category by category, write descs into xlsx

For each category sheet, walk the rows. For each row:

- **Gather context** from the row's Brand, Campaign, Agency, Country, Subcategory, Entry URL.
- **Pull visual context** if a local image exists at `digital-presentation-images/originals/<category-slug>-<brand>-<campaign>.<ext>` — open it with the Read tool to see the creative directly. Small text/copy in the image often reveals the campaign idea.
- **Try the lovethework Entry URL** via WebFetch. Many are paywalled now, but some still return useful metadata. `lovethework.com` and `lovetheworkmore.com` are on the network allowlist.
- **Use general knowledge / web search** when needed. Most major Cannes 2025 winners have press coverage on Adweek, AdAge, LBB, Campaign, etc.
- **Write a 1–2 sentence desc** following the quality bar below.
- **Save into the Description column** of the xlsx.

Save the xlsx after each category so progress isn't lost if the session ends.

### 3. Sync xlsx descriptions into index.html

After all (or some) categories are filled, run a small patch script analogous to `scripts/patch_imgurls.py` that:

1. Reads each xlsx row's Brand, Campaign, and Description.
2. Finds the matching entry block in `index.html` by brand+campaign (handle Unicode `\uXXXX` escapes the same way `patch_imgurls.py` does — see that script for the matching variants).
3. Replaces only the `desc:` field; leaves everything else (id, tier, agency, tags, keywords, youtubeId, videoUrl, caseUrl, imgUrl) untouched.

Backup `index.html` before writing.

The Description column in the xlsx becomes the durable record. Future Cannes updates just refill it.

## Quality bar for descriptions

**Good descs are specific, concrete, and revealing.** Examples:

❌ Template trash to delete:
- "Grand Prix winner in B01 - Use of Music. One Second Ads for Budweiser by Africa Creative DDB (Brazil)."
- "Bronze winner in Brand Storytelling. TANTRUM GIRL for TOBLERONE by LEPUB (Italy)."

✅ Accurate, idea-driven (illustrative; verify the actual concept):
- "Budweiser ran 1-second TV spots during the most expensive ad windows of the year — proving the brand is so iconic, it only needs a flash to register."
- "The launch piece of Burger King UK's Bundles of Joy: new parents, fresh from hospital, craving a Whopper — captured with disarming honesty on outdoor posters."

Rules:

- **Lead with the idea, not the metadata.** Don't start with "Gold winner in..." — that's already shown via badges.
- **Name the insight, twist, or mechanic** when it's clear. ("turned waiting time into payment time"; "let listeners name the next Oreo flavor"; etc.)
- **One sentence is fine; two if needed for setup + payoff.** Resist the urge to write three.
- **Avoid generic awards-jargon** ("groundbreaking", "innovative", "powerful"). Specific verbs > superlatives.
- **Don't restate the campaign name** unless it's load-bearing. The user already sees the title.
- **Stay grounded.** If you can't determine the actual concept from available sources, say what's known — don't invent a creative idea.

## Things to watch for

- **Cross-category duplicates** (campaigns appearing in multiple sheets — e.g., Apple "Find Your Friends" in Film and Film Craft). The site dedupes campaign+brand into a single card, so they share one desc on the live site. Write the desc once based on the campaign as a whole; don't tailor per category.
- **Subcategory-level duplicates within a sheet** (e.g., Burger King "Bundles of Joy OOH1/OOH2/OOH3/OOH4" each has its own row). These ARE separate cards on the site. Each board may have its own concept worth describing. Write distinct descs.
- **Existing descs that happen to be accurate** — there are some hand-authored gems in the current data. When you check an entry and the existing desc is genuinely good, keep it. Don't rewrite for the sake of rewriting.

## Files & paths

- xlsx: `Brain_Cannesdy_Brand_Experience_2025.xlsx`
- site: `index.html`
- local images: `digital-presentation-images/originals/<category-slug>-<brand>-<campaign>.<ext>`
- existing patch script reference: `scripts/patch_imgurls.py` (model your `desc:` patcher after this)
- existing data sync skill: `cannesdy-site-sync` — reference for understanding the DATA structure, but **do not run the full skill workflow** because it regenerates the entire DATA block from scratch and would lose the descs you just wrote. Only do surgical desc patches.

## Tracking progress

- The Description column in xlsx is the progress tracker — empty cells = not done.
- Process one category at a time. Save xlsx after each.
- Suggest checkpointing memory after each big category: "Phase 3: completed Audio & Radio (21/21), Film (52/52)..." so the next session can pick up cleanly.

## Validation

After syncing descs into `index.html`:

```python
html = open('index.html').read()
# Sanity checks
assert html.count('<script') == html.count('</script>'), "script tags unbalanced"
assert 'const DATA = {' in html, "DATA block missing"
# Count entries with template-shaped descs (these would mean unfinished work)
import re
template_count = len(re.findall(r'desc:"[A-Z][a-z]+ Prix winner in|desc:"Gold winner in|desc:"Silver winner in|desc:"Bronze winner in', html))
print(f"Entries still on the auto-template: {template_count}")
```

When `template_count` is 0, Phase 3 is complete.

## Don't break

- Don't change `id`, `tier`, `tags`, `keywords`, `youtubeId`, `videoUrl`, `caseUrl`, `imgUrl`, `brand`, `campaign`, `agency`, `country`, or `grad` on any entry.
- Don't rename or rearrange categories.
- Don't run the full cannesdy-site-sync skill — it regenerates DATA from xlsx and would lose any unsaved desc state.
- Backup before writing. Always.
