# Indicators of AI-Generated Writing

*Research compiled May 4, 2026. Sources: QuillBot, GPTZero, academic AI detection literature.*

---

## Overview

AI detectors (GPTZero, QuillBot AI Detector, etc.) use two primary statistical signals to flag AI-generated text:

- **Perplexity** — how predictable the word choices are. AI favors safe, statistically likely words, producing low perplexity. Human writing surprises more.
- **Burstiness** — how varied sentence structure and length are. AI writes at a uniform rhythm; humans naturally mix short punchy sentences with longer complex ones.

When perplexity is low *and* burstiness is low, AI authorship is likely. Human reviewers picking up on the same patterns do so intuitively — the text feels "smooth but flat."

---

## 1. Vocabulary Indicators

### Overused AI Words

These words appear disproportionately in AI output. A single occurrence is not a flag; a cluster of them is.

| Category | Words |
|---|---|
| **Abstract nouns** | tapestry, landscape, ecosystem, nexus, interplay, synergy |
| **Elevated verbs** | delve, embark, leverage, harness, elevate, underscore, navigate |
| **Filler adjectives** | crucial, intricate, robust, vibrant, dynamic, groundbreaking, transformative, innovative, comprehensive, nuanced |
| **Vague amplifiers** | world-class, cutting-edge, state-of-the-art, unprecedented, remarkable |
| **Corporate verbs** | resonate, enhance, foster, facilitate, champion, bolster |
| **Hedge verbs** | offerings, ensure, enable, empower, unlock |

### Overused AI Phrases

- "It is important to note that…"
- "In today's fast-paced world…"
- "Delving into…" / "By delving deeper…"
- "A key aspect of…"
- "This highlights the importance of…"
- "In conclusion, it is clear that…"
- "Moreover," / "Furthermore," / "Additionally," (used to start nearly every paragraph)
- "It is worth noting that…"
- "At the intersection of…"
- "A testament to…"
- "This underscores the need for…"
- "In the realm of…"
- "A holistic approach to…"

### Problematic Synonym Substitution

AI sometimes substitutes unusual synonyms in an attempt to vary vocabulary, producing unnatural phrasing:

| Natural phrasing | AI substitution |
|---|---|
| mix / combination | tapestry |
| start / begin | embark on |
| use | leverage / harness |
| show | underscore / illuminate |
| important | crucial / paramount |
| wide-ranging | multifaceted |

---

## 2. Grammar & Sentence Structure Indicators

### Uniform Sentence Length (Low Burstiness)

Human writers naturally vary sentence length — a long complex sentence is followed by a short one for punch. AI writing tends to use similarly-lengthed, similarly-structured sentences throughout, creating a smooth but monotonous rhythm.

**AI pattern:** Every sentence is a medium-length declarative statement.

**Human pattern:** Short sentence. Then a longer, more elaborate one that builds on what came before — with subordinate clauses, qualifications, or emphasis placed deliberately. Then another short one.

### Formulaic Transitions

AI uses transition words at a much higher frequency than humans, especially at the start of paragraphs:

> "Moreover, the technology has…"  
> "Furthermore, studies have shown…"  
> "Additionally, this approach enables…"  
> "In addition, the team has…"

### Parallel Structure Overuse

AI defaults to heavily parallel grammatical structures within a sentence or paragraph — every item in a list uses the same verb form, every paragraph follows the same subject-verb-object pattern. Human writing breaks this pattern naturally.

### Passive Voice Overuse

AI defaults to passive constructions to sound formal and neutral:
- "It has been demonstrated that…"
- "This was achieved through…"
- "The results were validated by…"

---

## 3. Tone & Style Indicators

### Cheerfully Formal / Impersonal Register

AI defaults to a tone that is simultaneously formal and eager-to-please — neutral, positive, and polished — even when a more direct, urgent, or personal voice would be more appropriate. This is sometimes described as sounding like a well-written brochure rather than a genuine human voice.

### Lack of Personal Voice, Anecdote, or Specificity

AI generates *generic* explanations. It rarely:
- References a specific personal memory or moment
- Includes an observation that only someone with direct experience could make
- Shows hesitation, doubt, or nuanced ambivalence
- Makes an unexpected comparison or uses an original metaphor

### Repetition of Concepts Across Different Words

AI may use different words but repeat the same concept in adjacent sentences. For example (from QuillBot's analysis of ChatGPT describing cities):
> "A vibrant metropolis…diverse neighborhoods…vibrant culture…"

The word "vibrant" appeared in descriptions of New York, London, and Tokyo. "World-class" and "diverse neighborhoods" similarly recurred. Human writers select and commit to specifics; AI circles back to the same conceptual ground.

### Excessive Superlatives and Enthusiasm

AI writing tends toward unearned enthusiasm — words like "exceptional," "remarkable," "extraordinary," "unprecedented," and "transformative" appear at a higher rate than a neutral human analyst would use. This is particularly noticeable in introductory and closing sentences.

### No Thematic Arc

Human writing typically has themes woven throughout, with a conclusion that closes those themes meaningfully. AI writing often has a **summary conclusion** (restating what was said) rather than a **resonant conclusion** (bringing the reader somewhere new).

---

## 4. Structural & Formatting Indicators

### Default Bullet-Point / Header Structure

AI defaults to a characteristic presentation format:

- **Bolded key term** followed by a colon and explanation
- Nested bullet points where plain prose would be more natural
- Headers for every paragraph or two (over-segmentation)
- Symmetrical paragraph lengths

Human professional writing, by contrast, often uses flowing prose for nuanced arguments and reserves lists for genuinely enumerable items.

### The "Three-Part" Pattern

AI responses commonly follow:
1. An introductory paragraph that restates the prompt or question
2. A body with numbered or bulleted sections
3. A conclusion paragraph that summarizes the body

This structure is fine for some contexts but becomes a tell when used uniformly across all content types.

### Excessive Bold Emphasis

AI over-bolds text, treating nearly every key phrase as worthy of bolding. Human writers use bold sparingly, for truly critical terms.

### Unnecessary Tables

AI frequently presents information in tables even when a simple sentence or short paragraph would communicate it more naturally and persuasively.

---

## 5. Content & Accuracy Indicators

### AI Hallucination (Fabricated Facts or Sources)

AI can confidently assert incorrect information — fabricated quotes attributed to real experts, non-existent studies, incorrect dates, or URLs that do not resolve. This is one of the most reliable hard indicators of AI authorship.

**Example (from QuillBot research):** When asked to find a quote from ADHD expert Dr. Russell Barkley, ChatGPT produced a plausible-sounding quote and cited a webpage. The quote did not appear on that page, and the page contained no mention of Dr. Barkley.

### Internal Inconsistency

AI may contradict itself between paragraphs — stating one figure or characterization early, then a different one later — because it generates text sequentially without maintaining a full awareness of what it has already written.

### Failure to Answer the Actual Question (Intent Mismatch)

AI does not always pick up on nuanced reader intent. It may address a general topic when the reader needed a specific application, or provide encyclopedia-style background when the context called for an opinion or recommendation.

### Missing Tacit Knowledge

Genuine expert writing contains observations, caveats, or framings that only someone with real domain experience would include. AI writing tends to stay on the surface of a topic — comprehensive-sounding but lacking the unexpected insight that signals authentic expertise.

---

## 6. Presentation-Level Indicators (Documents and Letters)

### Over-Structured Professional Documents

In letters, reports, and professional communications, AI-written content tends to:
- Use headers within a letter body (unusual in formal correspondence)
- Include bulleted breakdowns where a sentence would suffice
- Open with a highly formulaic salutation and closing
- State the document's own structure ("I will discuss three key areas…")

### Symmetrical Section Lengths

AI writes sections of nearly equal length, regardless of the relative importance or complexity of the content. Human writers allocate space based on what actually needs elaboration.

### Consistent Paragraph Opening Pattern

AI frequently starts every paragraph the same way — "This technology…", "Furthermore, this approach…", "The result is…" — creating a detectable repetitive cadence.

### Lack of Authorial Idiosyncrasies

Experienced human writers have habits: a preferred sentence-opening rhythm, characteristic punctuation choices (em-dashes, semicolons), preferred analogies, or a distinctive way of qualifying claims. AI writing is stylistically smooth but characteristically *neutral* — it has no idiosyncrasies.

---

## 7. Markdown-First Artifacts (AI → Markdown → Word Pipeline)

A common AI-assisted workflow is: prompt an LLM → receive markdown output → convert to DOCX via pandoc or a similar tool. Pandoc is capable and handles many elements gracefully, but certain conventions that are natural in markdown look unconventional in a Word document intended for professional correspondence. These artifacts indicate the document was *composed in markdown*, not typed natively in a word processor.

### What Pandoc Converts Cleanly (Not a Reliable Tell)

These elements render normally after conversion and are not reliable indicators on their own:

- `**bold**` and `*italic*` → standard bold and italic
- `- bullet` lists and `1.` numbered lists → proper Word list styles
- `[text](url)` hyperlinks → Word hyperlinks with display text
- `#`/`##`/`###` headers → Word Heading 1/2/3 styles (though see below)
- Standard paragraph breaks → normal paragraph spacing

### What Still Reads as Markdown-Origin After Conversion

**Blockquotes used for visual indentation**
Markdown `>` blockquotes become pandoc's "Block Text" paragraph style in Word — an indented, often bordered block. In a formal letter or report, this is an unusual visual element. Letter writers in Word indent with tab stops or adjust paragraph margins; they do not create indented prose blocks. Using blockquotes to structure the body sections of a letter (rather than to quote a source) is a clear markdown convention that survives conversion as an obvious structural oddity.

**Section headers inside a letter body**
If markdown uses `##` headers to divide sections of a letter, pandoc converts these to Word Heading 2 style. A heading mid-letter — bold, larger, with paragraph spacing above and below — looks like a document template, not a letter. Formal correspondence does not use heading hierarchy; it uses paragraph breaks and transitional prose.

**Trailing backslash line breaks**
Markdown uses `\` at the end of a line to force a line break within a paragraph. Pandoc converts these to Word soft returns (Shift+Enter). The visual result may look fine, but the source markdown file retains the `\` characters, which are immediately recognizable as markdown syntax if the `.md` file is ever inspected.

**Horizontal rules as section dividers**
`---` in markdown becomes a Word horizontal rule. In a letter this is highly unusual — formal letters use white space and paragraph structure to separate sections, not graphical dividers. Horizontal rules in Word are typically used in newsletters, reports, or résumés, not in correspondence.

**Inconsistent spacing behavior**
Pandoc applies its default `reference.docx` paragraph spacing, which may differ from what the document's context calls for. The result can be noticeably uniform paragraph spacing that doesn't match the conventions of the document type (e.g., a letter with single-spaced lines but pandoc's default `Space After` applied uniformly).

**Over-structured body sections**
A Word-native letter writer structures arguments in prose. A markdown-first document tends to have explicit sections, sub-bullets, and headers that reflect the markdown composer's visual organization habits. After conversion, the structure is preserved — but it looks like a formatted report, not a letter.

**Density of inline bold**
AI writing in markdown frequently bolds key phrases throughout the body (`**transformative impact**`, `**98–100%**`, `**drop-in retrofits**`). After conversion, these all render as bold. The result is a letter body with bolded phrases every sentence or two — a pattern that looks like a marketing brochure rather than professional correspondence, and that no one composing directly in Word would produce at that density.

**Special Unicode characters typed directly**
AI outputs typographically correct Unicode characters — em dash (—), en dash (–), curly quotes (" "), ellipsis (…), etc. — as literal characters in the markdown source. A human typing natively in Word typically gets these through autocorrect (e.g., `--` → —) or Insert > Symbol, and uses them more sparingly. In the rendered document the characters themselves are not wrong, but their *pattern of use* is a tell:

- **Em dash (—)** is the biggest offender — and has become so closely associated with AI output that many readers now assume a document was AI-generated the moment they see one, regardless of how the character was inserted. This is not purely about the Unicode character itself; it is about the fact that AI models use em dashes constantly as a stylistic default, and that pattern has trained a widespread cultural reflex. A human letter-writer more often uses a comma, parentheses, or a new sentence. Even a single em dash in a support letter or professional correspondence can trigger suspicion. Multiple em dashes are nearly disqualifying.
- **En dash (–)** in numeric ranges (98–100%) is typographically correct, but a human typing in Word would often just use a hyphen (98-100%) and not bother with the proper character unless they are meticulous about typography.
- **Ellipsis character (…)** vs. three separate periods (...). AI outputs the Unicode ellipsis; humans usually type three periods.

**When special characters are *not* a tell:**
The key is whether the character fits the document type and field conventions:
- Degree symbol (°) in any document discussing temperature, angles, or coordinates — entirely expected.
- Greek letters (α, β, θ, μ, etc.) in a technical paper, scientific report, or equation — appropriate and normal.
- A single em dash used deliberately in formal prose — not a flag on its own.

The tell is an **unlikely character for the document type, used at an unlikely frequency**. The em dash deserves special emphasis: it has become so culturally synonymous with AI writing that its presence in correspondence (letters, emails, memos) now triggers an AI assumption in many readers even when it appears only once. Replace em dashes in correspondence with commas, parentheses, colons, or restructured sentences.

### Summary Table

| Element | Pandoc conversion | Residual tell? |
|---|---|---|
| Bold / italic | Clean | No |
| Bullet lists | Clean | Only if overused |
| Hyperlinks | Clean | No |
| `>` Blockquotes | Renders as Block Text style | **Yes** — unusual in letters |
| `##` Headers in letter body | Renders as Heading style | **Yes** — looks like a report |
| `---` Horizontal rules | Renders as Word rule | **Yes** — unusual in correspondence |
| Trailing `\` line breaks | Renders correctly | Visible in source `.md` file |
| Uniform pandoc paragraph spacing | Renders consistently | **Yes** — may not match doc type conventions |
| Heavy inline bolding | Renders as bold | **Yes** — density is the tell |
| Em dash (—) in correspondence | Renders as em dash | **Yes** — frequency and context are the tell |
| Greek letters / degree symbol in technical writing | Renders correctly | No — appropriate for document type |

---

## 8. Quick Reference Checklist

Use this checklist when reviewing a document for AI indicators:

**Vocabulary**
- [ ] Multiple overused AI words in close proximity (tapestry, leverage, delve, robust, etc.)
- [ ] Unnatural synonym substitutions
- [ ] Excessive superlatives or enthusiasm words

**Grammar & Structure**
- [ ] Uniform sentence length throughout (no short/long variation)
- [ ] Transitions ("Moreover," "Furthermore,") starting most paragraphs
- [ ] Overly parallel sentence structure

**Tone & Voice**
- [ ] Cheerfully formal, impersonal register inconsistent with context
- [ ] No personal observations, anecdotes, or unexpected specifics
- [ ] Repeated concepts with different words
- [ ] No genuine thematic arc — summary conclusion instead of resonant one

**Formatting**
- [ ] Bullet points where flowing prose is more appropriate
- [ ] Excessive bolding of phrases
- [ ] Headers every few sentences
- [ ] Symmetrical section/paragraph lengths

**Content**
- [ ] Any facts or citations that cannot be independently verified
- [ ] Internal contradictions between sections
- [ ] Generic explanations that lack expert-level specificity

**Markdown-First Artifacts**
- [ ] Blockquotes used as indentation rather than to quote a source
- [ ] Section headers (`##` / Heading styles) inside a letter or memo body
- [ ] Horizontal rules as section dividers in correspondence
- [ ] Inline bold at high density throughout body paragraphs
- [ ] Em dashes (—) appearing anywhere in a letter, email, or memo (even one triggers suspicion; multiple are nearly disqualifying)
- [ ] Other Unicode characters (–, …, curly quotes) at higher frequency than the document type warrants

---

## 9. Important Caveats

- **No single indicator is proof.** These are signals, not proof. A human can use "leverage" once; an AI can avoid it if prompted. Look for *clusters* of indicators.
- **AI detectors are imperfect.** Tools like GPTZero and QuillBot AI Detector are useful starting points but produce both false positives (flagging human writing as AI) and false negatives (missing well-edited AI content).
- **Edited AI output is harder to detect.** AI text that has been substantially revised by a human may not trigger these indicators. The more editing a human applies, the less detectable the AI origin becomes.
- **Context matters.** Some legitimate human writing — highly technical, procedural, or formulaic by necessity — may resemble AI output. Flagging should prompt inquiry, not accusation.

---

*Sources: QuillBot Blog (2024–2026); GPTZero FAQ (2026); Originality.ai research summaries; academic literature on perplexity/burstiness in AI detection.*
