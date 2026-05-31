import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const html = await readFile(new URL("../index.html", import.meta.url), "utf8");
const js = await readFile(new URL("../src/interactions.js", import.meta.url), "utf8");

const slideCount = (html.match(/<section class="slide/g) || []).length;

assert.equal(slideCount, 20, "deck should contain exactly 20 slides");
assert.match(js, /Reveal\.initialize/i, "deck should initialize reveal.js");
assert.match(html, /문서를 자동화 자산으로 바꾸는 AI 활용법/, "deck should include the title");
assert.match(html, /ooxml-reveal/i, "deck should include the OOXML reveal hero");
assert.match(html, /automation-pipeline/i, "deck should include the automation pipeline hero");
assert.match(html, /safe-runner/i, "deck should include the safe runner hero");
assert.match(html, /src\/styles\.css/, "deck should link the custom stylesheet");
assert.match(html, /src\/interactions\.js/, "deck should load the interaction script");
