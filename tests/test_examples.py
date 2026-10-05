import json
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from excel_converter.core import convert, headers, load_profile, load_source, suggest_header, suggest_mapping
from excel_converter.example_pack import create_example_pack


class ExamplePackTests(unittest.TestCase):
    def test_all_examples_match_catalog_and_settings(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            paths = create_example_pack(directory)
            self.assertEqual(len(paths), 24)
            catalog = json.loads((directory / 'examples.json').read_text(encoding='utf-8'))
            self.assertEqual(len(catalog), 24)
            for case in catalog:
                with self.subTest(file=case['file']):
                    source = load_source(directory / case['file'])
                    profile = load_profile()
                    self.assertEqual(suggest_header(source.sheets[case['sheet']], profile), case['header_row'])
                    mapping = suggest_mapping(headers(source.sheets[case['sheet']], case['header_row']), profile)
                    result = convert(source, case['sheet'], case['header_row'], profile, mapping)
                    self.assertEqual(len(result.records), case['valid_rows'])
                    self.assertEqual(bool(result.errors), case['status'] != 'ready')
                    if case['settings']:
                        for field in profile['fields']:
                            if field['kind'] in ('date', 'amount'):
                                field.update(case['settings'])
                        result = convert(source, case['sheet'], case['header_row'], profile, mapping)
                        self.assertFalse(result.errors)
                        self.assertEqual(len(result.records), case['records_after_settings'])
                    name = Path(case['file']).name
                    if name.startswith('02_'):
                        self.assertTrue(all(row['date'] == date(2026, 12, 31) for row in result.records))
                    elif name.startswith('03_'):
                        self.assertTrue(all(row['amount'] == Decimal('1234.56') for row in result.records))
                    elif name.startswith('12_'):
                        self.assertEqual(sum(row['amount'] for row in result.records), Decimal('100000'))
            first = paths[0].read_bytes()
            with self.assertRaises(ValueError):
                create_example_pack(directory)
            self.assertEqual(paths[0].read_bytes(), first)


if __name__ == '__main__':
    unittest.main()
