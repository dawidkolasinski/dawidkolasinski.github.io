# Privacy and Terms localization

This directory prepares local translations of the four canonical Sunchucks and
Moonchucks Privacy and Terms pages. The HTML files remain the source of the
policy clauses. Existing Polish pages are retained directly, including local
changes. No English-language precedence clause is added.

The current package is a local working-tree snapshot. Its `sourceHash` does not
claim that these files are published or contained in the recorded Git commit.
`provenance.gitHead` records context only; `sourceFiles` records the actual
SHA-256 of all four English and four Polish HTML files. Publishing requires a
separate authorized operation.

## Contract

`source-index.json` has schema version 1:

- `documents`: four ordered source documents, with `id`, `app`, `kind`,
  `sourceLanguage`, `path`, `title`, `updatedLabel`, `footerHTML`, and `blocks`.
  Each block has a stable positional `id`, an outer `tag`, and its complete
  `html` contents. Top-level lists retain their full `<li>` structure.
- `uiLabels`: navigation, language, table-of-contents and accessibility labels.
- `privacyHostingVariants`: complete paragraphs for GitHub Pages, Cloudflare
  Pages, and Firebase Hosting for each Privacy document. A renderer selects
  exactly the paragraph for its actual host, preserving the other policy
  clauses. These alternatives do not announce a deployment.
- `translationUnits`: deterministic transport bindings for full headings,
  paragraphs and list items. Inline hyperlinks are never separate fragments.
- `sourceHash`, `sourceFiles`, and `provenance`: snapshot identity and evidence.

Calculate `sourceHash` as SHA-256 of UTF-8 JSON for **all index fields except**
`sourceHash`, `provenance`, and `translationUnits`, using Python
`json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(',', ':'))`.
The parser also recomputes the translation-unit bindings and every source-file
hash. A changed manifest, stale HTML, or different DOCX source pin fails closed.

Each `locales/<locale>.json` has `schemaVersion`, `sourceHash`, `locale`,
`documents`, `uiLabels`, `privacyHostingVariants`, and `translation` provenance.
`documents` is keyed by source document ID; `blocks` is keyed by source block ID
and contains translated complete inner HTML. Titles and dates are plain text;
footer and hosting fields contain complete inner HTML. All source hyperlinks
remain real links with their original destinations. Website rendering may route
internal policy links to the active local language.

The final catalog contains 52 explicit locale files. The following same-language
regional variants share their reviewed base content: `en-AU/en-GB/en-IN → en`,
`es-419/es-US → es`, `fr-CA → fr`, and `zh-HK → zh-Hant`.
Portuguese for Portugal has its own reviewed regional copy rather than being
an unedited Brazilian-Portuguese alias.
The finalizer records this relationship in `locale-sources.json` and each derived
catalog. Unavailable translations never fall back silently to English.

## Local workflow

Use a Python runtime with the dependencies in `legal/requirements.txt`.
Create an ignored virtual environment with `python3 -m venv .venv` and install
them using `.venv/bin/python -m pip install -r legal/requirements.txt`. This does
not depend on a particular Codex installation or machine-local bundled runtime.

1. Finish and review the source HTML changes before rebuilding the source index.
2. Run `python legal/scripts/legal_localization.py build`.
3. Export with `python legal/scripts/legal_localization.py export --output /tmp/legal/source.docx`.
4. Translate the document as a whole. Keep `[[LGL0001]]` identifiers and the
   `[[LGLSOURCE…]]` source pin. The DOCX also stores the pin in its subject.
   The verified Hindi Google export transliterates only the transport prefix
   `[[LGL0001]]` to `[[एलजीएल0001]]`. The importer normalizes that exact known
   prefix, preserves the numbered binding, and still requires the matching source
   hash and every unit. Unknown prefixes fail rather than being guessed.
5. Import into a review folder with
   `python legal/scripts/legal_localization.py import translated.docx --locale de --output /tmp/legal/raw/de.json`.
6. Review the complete translated clauses and save corrected catalogs to a
   separate `reviewed` folder. Preserve negations, permission conditions,
   subscription cancellation and restore rules,
   provider identities, NASA/NOAA limitations, numbers, and links according to
   the frozen source. A product-scope change requires a new source snapshot.
7. Finalize only the reviewed complete set:
   `python legal/scripts/legal_localization.py finalize --reviewed /tmp/legal/reviewed`.
8. Run `python legal/scripts/legal_localization.py check` and
   `python legal/scripts/test_legal_localization.py`.

`polish` regenerates only the Polish JSON from the existing canonical Polish
HTML. `check --allow-incomplete --directory /tmp/legal/raw` checks a partial
review batch without claiming complete coverage. `--allow-source` is reserved
for the exporter/importer round-trip test.

The DOCX exporter preserves real Word hyperlinks, list-item text, emphasis, and
full-sentence context. Each import checks source identity, all unit IDs,
hyperlink destinations, lists, numeric clauses, and field coverage. These checks
do not establish legal equivalence or replace language review. The source guard rejects retired rewarded-access clauses in English and Polish,
including the former 30-minute Premium offer. These clauses were removed from
the prepared release copy at the user’s direction. The package
records machine translation and local review separately from publication and
independent professional review.
