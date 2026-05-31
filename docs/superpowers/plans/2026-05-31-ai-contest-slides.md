# AI Contest Slides Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a polished browser-based HTML slide deck from `SLIDE_DESIGN.md`.

**Architecture:** Use a plain static GitHub Pages structure with `index.html` as the reveal.js deck, `src/styles.css` for all visual styling, and `src/interactions.js` for slide-specific interactions. Add a small Node smoke test that verifies the deck has the expected 20-slide structure and key assets before manual visual review.

**Tech Stack:** Static HTML, reveal.js CDN, lucide icons CDN, custom CSS, vanilla JavaScript, Node built-in test runner.

---

## File Structure

- Create: `index.html` - reveal.js presentation with 20 sections based on `SLIDE_DESIGN.md`.
- Create: `src/styles.css` - responsive 16:9 slide styling, component styles, hero visuals, and print-friendly rules.
- Create: `src/interactions.js` - reveal initialization, OOXML reveal interaction, safe-runner demo interaction, and keyboard affordances.
- Create: `tests/slide-smoke-test.mjs` - Node assertions for deck structure and core implementation hooks.
- Create: `README.md` - submission copy and local viewing instructions.

## Chunk 1: Deck Skeleton and Smoke Test

### Task 1: Smoke Test

**Files:**
- Create: `tests/slide-smoke-test.mjs`

- [ ] **Step 1: Write the failing test**

```js
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const html = await readFile(new URL("../index.html", import.meta.url), "utf8");

assert.equal((html.match(/<section class="slide/g) || []).length, 20);
assert.match(html, /reveal\.initialize/i);
assert.match(html, /OOXML/);
assert.match(html, /safe-runner/i);
assert.match(html, /src\/styles\.css/);
assert.match(html, /src\/interactions\.js/);
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/slide-smoke-test.mjs`

Expected: FAIL because `index.html` does not exist yet.

### Task 2: Static Deck Implementation

**Files:**
- Create: `index.html`
- Create: `src/styles.css`
- Create: `src/interactions.js`
- Create: `README.md`

- [ ] **Step 1: Implement the deck**

Build all 20 slides from `SLIDE_DESIGN.md`, with extra polish on the OOXML reveal, automation pipeline, safe-runner product mockup, safety model, and demo simulation slides.

- [ ] **Step 2: Run automated verification**

Run: `node tests/slide-smoke-test.mjs`

Expected: PASS.

- [ ] **Step 3: Run a local static server**

Run: `python3 -m http.server 8000`

Expected: The deck opens at `http://localhost:8000/`.

- [ ] **Step 4: Visual verification**

Use browser screenshots at desktop and mobile widths to confirm the deck renders, slides are navigable, and interactive elements work.
