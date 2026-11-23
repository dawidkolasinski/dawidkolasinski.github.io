#!/usr/bin/env python3
"""Failure-oriented checks for the frozen legal translation transport."""
import contextlib
import copy
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn
from lxml import html
import legal_localization as legal


class LegalLocalizationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = legal.load_source(legal.DEFAULT_INDEX)
        cls.english = json.loads((legal.ROOT / 'legal/locales/en.json').read_text())

    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix='legal-transport-')
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name)
        self.docx = self.folder / 'source.docx'
        with contextlib.redirect_stdout(io.StringIO()):
            legal.export_docx(self.source, self.docx)

    def import_source(self, path=None):
        with contextlib.redirect_stdout(io.StringIO()):
            return legal.import_docx(self.source, path or self.docx, 'en', self.folder / 'result.json', allow_source=True)

    def edit_docx(self, change):
        doc = Document(self.docx)
        change(doc)
        path = self.folder / 'changed.docx'
        doc.save(path)
        return path

    def test_roundtrip_preserves_every_full_clause_and_hyperlink(self):
        result = self.import_source()
        for doc in self.source['documents']:
            actual = result['documents'][doc['id']]
            self.assertEqual(actual['title'], doc['title'])
            self.assertEqual(actual['updatedLabel'], doc['updatedLabel'])
            for block in doc['blocks']:
                value = actual['blocks'][block['id']]
                if block['tag'] in ('ul', 'ol'):
                    original = html.fragment_fromstring('<div>' + block['html'] + '</div>')
                    translated = html.fragment_fromstring('<div>' + value + '</div>')
                    self.assertEqual([item.text_content().strip() for item in translated],
                                     [item.text_content().strip() for item in original])
                else:
                    self.assertEqual(legal.plain(value), legal.plain(block['html']))
                self.assertEqual(legal.hrefs(value), legal.hrefs(block['html']))
        self.assertEqual(set(result['uiLabels']), set(self.source['uiLabels']))
        self.assertEqual(set(result['privacyHostingVariants']), set(self.source['privacyHostingVariants']))

    def test_metadata_pin_suffices_for_the_verified_first_google_pilot(self):
        def change(doc):
            for paragraph in doc.paragraphs:
                if paragraph.text.startswith('[[LGLSOURCE'):
                    paragraph._p.getparent().remove(paragraph._p)
                    break
        self.import_source(self.edit_docx(change))

    def test_visible_pin_survives_a_missing_document_subject(self):
        self.import_source(self.edit_docx(lambda doc: setattr(doc.core_properties, 'subject', '')))

    def test_verified_hindi_transport_prefix_preserves_full_unit_bindings(self):
        def change(doc):
            for paragraph in doc.paragraphs:
                if legal.MARKER.fullmatch(paragraph.text):
                    paragraph.text = paragraph.text.replace('LGL', 'एलजीएल')
        result = self.import_source(self.edit_docx(change))
        self.assertEqual(result['documents'], self.import_source()['documents'])

    def test_unknown_transport_prefix_is_not_guessed(self):
        def change(doc):
            for paragraph in doc.paragraphs:
                if paragraph.text == '[[LGL0023]]':
                    paragraph.text = '[[OTHER0023]]'
                    return
        with self.assertRaisesRegex(ValueError, 'Unit coverage mismatch'):
            self.import_source(self.edit_docx(change))

    def test_conflicting_source_pins_are_rejected(self):
        def change(doc):
            doc.core_properties.subject = 'sourceHash ' + '0' * 64
        with self.assertRaisesRegex(ValueError, 'DOCX sourceHash'):
            self.import_source(self.edit_docx(change))

    def test_missing_source_pin_is_rejected(self):
        def change(doc):
            doc.core_properties.subject = ''
            for paragraph in doc.paragraphs:
                if paragraph.text.startswith('[[LGLSOURCE'):
                    paragraph._p.getparent().remove(paragraph._p)
                    break
        with self.assertRaisesRegex(ValueError, 'DOCX sourceHash'):
            self.import_source(self.edit_docx(change))

    def test_missing_full_unit_is_rejected(self):
        def change(doc):
            paragraphs = doc.paragraphs
            marker = next(i for i, paragraph in enumerate(paragraphs) if paragraph.text == '[[LGL0023]]')
            for paragraph in paragraphs[marker:marker + 2]:
                paragraph._p.getparent().remove(paragraph._p)
        with self.assertRaisesRegex(ValueError, 'Unit coverage mismatch'):
            self.import_source(self.edit_docx(change))

    def test_duplicate_unit_is_rejected(self):
        def change(doc):
            doc.add_paragraph('[[LGL0023]]')
            doc.add_paragraph('Duplicate clause.')
        with self.assertRaisesRegex(ValueError, 'Duplicate unit'):
            self.import_source(self.edit_docx(change))

    def test_changed_link_target_is_rejected(self):
        def change(doc):
            link = next(iter(doc.element.body.iter(qn('w:hyperlink'))))
            relation = doc.part.rels[link.get(qn('r:id'))]
            relation._target = 'https://example.com/changed'
        with self.assertRaisesRegex(ValueError, 'Hyperlinks changed'):
            self.import_source(self.edit_docx(change))

    def test_changed_provider_identifier_is_rejected(self):
        def change(doc):
            for paragraph in doc.paragraphs:
                if 'NOAA-20' in paragraph.text:
                    for run in paragraph.runs:
                        if 'NOAA-20' in run.text:
                            run.text = run.text.replace('NOAA-20', 'NOAA-21')
                            return
        with self.assertRaisesRegex(ValueError, 'Numeric clause mismatch'):
            self.import_source(self.edit_docx(change))

    def test_retired_rewarded_clause_is_rejected_in_translated_html(self):
        result = copy.deepcopy(self.english)
        result['documents']['moonchucks_terms']['blocks']['block-005'] += ' Optional rewarded access grants 30 minutes of Premium.'
        with self.assertRaisesRegex(ValueError, 'Retired rewarded Premium clause'):
            legal.validate_locale(self.source, result, 'en')

    def test_retired_rewarded_clause_blocks_rebuilding_english_or_polish_source(self):
        for relative in self.source['sourceFiles']:
            path = self.folder / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes((legal.ROOT / relative).read_bytes())
        for relative, clause in (
            ('privacy/index.html', 'An optional rewarded ad grants 30 minutes of Premium.'),
            ('pl/terms/index.html', 'Reklama z nagrodą daje 30 minut dostępu Premium.'),
        ):
            with self.subTest(source=relative):
                path = self.folder / relative
                original = path.read_bytes()
                try:
                    path.write_bytes(original + clause.encode())
                    with patch.object(legal, 'ROOT', self.folder):
                        with self.assertRaisesRegex(ValueError, 'Retired rewarded Premium clause'):
                            legal.build(self.folder / 'new-index.json')
                finally:
                    path.write_bytes(original)

    def test_manifest_content_change_requires_a_new_hash(self):
        source = copy.deepcopy(self.source)
        source['documents'][0]['blocks'][0]['html'] += ' Changed.'
        path = self.folder / 'source-index.json'
        legal.write_json(path, source)
        with self.assertRaisesRegex(ValueError, 'manifest hash mismatch'):
            legal.load_source(path)

    def test_translation_unit_bindings_cannot_bypass_the_source_hash(self):
        source = copy.deepcopy(self.source)
        source['translationUnits'][0]['target']['document'] = 'moonchucks_terms'
        path = self.folder / 'source-index.json'
        legal.write_json(path, source)
        with self.assertRaisesRegex(ValueError, 'binding differs'):
            legal.load_source(path)

    def test_rehashed_manifest_cannot_replace_the_canonical_html_clauses(self):
        source = copy.deepcopy(self.source)
        source['documents'][0]['blocks'][0]['html'] += ' Changed clause.'
        source['translationUnits'] = legal.units_for(source)
        source['sourceHash'] = legal.source_hash(source)
        path = self.folder / 'source-index.json'
        legal.write_json(path, source)
        with self.assertRaisesRegex(ValueError, 'blocks differ from canonical HTML'):
            legal.load_source(path)

    def test_dirty_source_html_is_checked_against_its_file_hash(self):
        for relative in self.source['sourceFiles']:
            path = self.folder / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes((legal.ROOT / relative).read_bytes())
        (self.folder / 'moonchucks/terms/index.html').write_text('Changed source')
        with patch.object(legal, 'ROOT', self.folder):
            with self.assertRaisesRegex(ValueError, 'Source HTML hash mismatch'):
                legal.load_source(legal.DEFAULT_INDEX)

    def test_canonical_polish_keeps_existing_source_blocks(self):
        result = legal.canonical_pl(self.source)
        legal.validate_locale(self.source, result, 'pl')
        for doc_id, app, kind, _ in legal.SOURCES:
            relative = ('pl/' if app == 'sunchucks' else 'moonchucks/pl/') + kind + '/index.html'
            original = legal.extract(doc_id, app, kind, relative)
            self.assertEqual(result['documents'][doc_id]['blocks'], {b['id']: b['html'] for b in original['blocks']})

    def test_missing_list_item_or_hosting_alternative_is_rejected(self):
        result = copy.deepcopy(self.english)
        block = next(b for b in self.source['documents'][0]['blocks'] if b['tag'] == 'ul')
        result['documents']['sunchucks_privacy']['blocks'][block['id']] = '<li>Lost clauses.</li>'
        with self.assertRaises(ValueError):
            legal.validate_locale(self.source, result, 'en')
        result = copy.deepcopy(self.english)
        result['privacyHostingVariants']['moonchucks_privacy'].pop('firebase')
        with self.assertRaisesRegex(ValueError, 'Hosting provider coverage'):
            legal.validate_locale(self.source, result, 'en')

    def test_foreign_english_fallback_is_rejected(self):
        result = copy.deepcopy(self.english)
        result['locale'] = 'ja'
        with self.assertRaisesRegex(ValueError, 'Untranslated full paragraph'):
            legal.validate_locale(self.source, result, 'ja')

    def test_authored_assembly_has_no_fictitious_document_input(self):
        units = {unit['id']: unit['html'] for unit in self.source['translationUnits']}
        result = legal.assemble_locale(self.source, 'en', units, allow_source=True)
        self.assertEqual(result['translation']['method'], 'assistant-authored')
        self.assertIs(result['translation']['independentNativeReview'], False)
        self.assertNotIn('inputFile', result['translation'])
        self.assertNotIn('inputSHA256', result['translation'])

    def test_authored_assembly_requires_every_full_unit(self):
        units = {unit['id']: unit['html'] for unit in self.source['translationUnits']}
        units.pop('LGL0054')
        with self.assertRaisesRegex(ValueError, 'Unit coverage mismatch'):
            legal.assemble_locale(self.source, 'en', units, allow_source=True)

    def test_full_coverage_never_writes_partial_catalogs_on_missing_review(self):
        output = self.folder / 'locales'
        with self.assertRaisesRegex(ValueError, 'Missing reviewed locale'):
            legal.finalize(self.source, output, self.folder / 'reviewed')
        self.assertFalse(output.exists())


if __name__ == '__main__':
    unittest.main()
