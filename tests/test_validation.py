import tempfile
import unittest
from copy import deepcopy
from datetime import date
from pathlib import Path
import shutil
import uuid

from openpyxl import Workbook

from app.runner import (
    EXCEL_ROW_NUM_KEY,
    expected_input_headers,
    load_config,
    make_input_audit_logger,
    prepare_input_rows,
    read_excel_rows,
    validate_and_prepare,
)


TODAY = date(2026, 5, 7)


class ValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cfg = load_config("config/config.yaml")

    def base_row(self):
        cols = self.cfg["columns"]
        return {
            EXCEL_ROW_NUM_KEY: 2,
            cols["first_name"]: "Ja'Kyren",
            cols["last_name"]: "Smith",
            cols["ssid"]: "1234567890",
            cols["dob"]: "05/06/2010",
            cols["gender"]: "M",
            cols["grade"]: "9",
            cols["disability"]: "Specific Learning Disability (SLD)",
            cols["ethcd"]: "Y",
            cols["race1"]: "White",
            cols["aeries_school"]: "CS - BARBARA PHELPS CS",
        }

    def validate(self, row, logs=None):
        return validate_and_prepare(
            row,
            self.cfg,
            validation_log_fn=(logs.append if logs is not None else None),
            today=TODAY,
        )

    def validate_address_patch(self, row, logs=None):
        return validate_and_prepare(
            row,
            self.cfg,
            validation_log_fn=(logs.append if logs is not None else None),
            today=TODAY,
            require_address_fields=True,
        )

    def assert_invalid(self, row, expected_text):
        with self.assertRaises(ValueError) as cm:
            self.validate(row)
        self.assertIn(expected_text, str(cm.exception))

    def test_valid_row_passes(self):
        student = self.validate(self.base_row())
        self.assertEqual(student.first_name, "Ja'Kyren")
        self.assertEqual(student.ssid, "1234567890")
        self.assertEqual(student.gender_ui, "Male")
        self.assertEqual(student.ethnicity_ui, "Yes")
        self.assertEqual(student.grade_ui, "9")
        self.assertEqual(student.race_ui, "White")

    def test_name_sanitization_removes_unsupported_characters(self):
        logs = []
        row = self.base_row()
        row[self.cfg["columns"]["first_name"]] = "Jayne (Jay)"
        row[self.cfg["columns"]["last_name"]] = "O\u2019Connor-Smith #2"

        student = self.validate(row, logs=logs)

        self.assertEqual(student.first_name, "Jayne Jay")
        self.assertEqual(student.last_name, "O'Connor-Smith")
        self.assertTrue(any("First Name sanitized" in msg for msg in logs))
        self.assertTrue(any("Last Name sanitized" in msg for msg in logs))

    def test_first_name_required_and_alpha(self):
        row = self.base_row()
        row[self.cfg["columns"]["first_name"]] = ""
        self.assert_invalid(row, "First Name is required")

        row = self.base_row()
        row[self.cfg["columns"]["first_name"]] = "----"
        self.assert_invalid(row, "First Name must contain")

    def test_last_name_required_and_alpha(self):
        row = self.base_row()
        row[self.cfg["columns"]["last_name"]] = None
        self.assert_invalid(row, "Last Name is required")

        row = self.base_row()
        row[self.cfg["columns"]["last_name"]] = "123"
        self.assert_invalid(row, "Last Name must contain")

    def test_ssid_required_and_exactly_10_digits(self):
        row = self.base_row()
        row[self.cfg["columns"]["ssid"]] = ""
        self.assert_invalid(row, "Missing SSID")

        row = self.base_row()
        row[self.cfg["columns"]["ssid"]] = "12345"
        self.assert_invalid(row, "exactly 10 digits")

        row = self.base_row()
        row[self.cfg["columns"]["ssid"]] = "123456789A"
        self.assert_invalid(row, "exactly 10 digits")

    def test_dob_required_real_date_and_min_age(self):
        row = self.base_row()
        row[self.cfg["columns"]["dob"]] = ""
        self.assert_invalid(row, "Birthdate is required")

        row = self.base_row()
        row[self.cfg["columns"]["dob"]] = "02/30/2010"
        self.assert_invalid(row, "Invalid DOB")

        row = self.base_row()
        row[self.cfg["columns"]["dob"]] = "05/08/2013"
        self.assert_invalid(row, "must be at least 13")

    def test_gender_case_insensitive_and_invalid(self):
        logs = []
        row = self.base_row()
        row[self.cfg["columns"]["gender"]] = "f"
        student = self.validate(row, logs=logs)
        self.assertEqual(student.gender_ui, "Female")
        self.assertTrue(any("Gender 'f' normalized to 'F'" in msg for msg in logs))

        row = self.base_row()
        row[self.cfg["columns"]["gender"]] = "X"
        self.assert_invalid(row, "Invalid gender")

    def test_grade_required_and_7_to_12(self):
        row = self.base_row()
        row[self.cfg["columns"]["grade"]] = ""
        self.assert_invalid(row, "Grade is required")

        row = self.base_row()
        row[self.cfg["columns"]["grade"]] = "6"
        self.assert_invalid(row, "Invalid grade")

        row = self.base_row()
        row[self.cfg["columns"]["grade"]] = 12
        student = self.validate(row)
        self.assertEqual(student.grade_ui, "12")

    def test_disability_required_and_mapped(self):
        row = self.base_row()
        row[self.cfg["columns"]["disability"]] = ""
        self.assert_invalid(row, "Description_CSE_DI is required")

        row = self.base_row()
        row[self.cfg["columns"]["disability"]] = "No Disability"
        self.assert_invalid(row, "Invalid disability")

    def test_ethcd_case_insensitive_y_or_n_only(self):
        logs = []
        row = self.base_row()
        row[self.cfg["columns"]["ethcd"]] = "n"
        student = self.validate(row, logs=logs)
        self.assertEqual(student.ethnicity_ui, "No")
        self.assertTrue(any("EthCd 'n' normalized to 'N'" in msg for msg in logs))

        row = self.base_row()
        row[self.cfg["columns"]["ethcd"]] = "D"
        self.assert_invalid(row, "Invalid ethnicity code")

    def test_race_required_and_unmapped_falls_back_to_other_with_log(self):
        row = self.base_row()
        row[self.cfg["columns"]["race1"]] = ""
        self.assert_invalid(row, "Description_STU_RC1 is required")

        logs = []
        row = self.base_row()
        row[self.cfg["columns"]["race1"]] = "Not In Mapping"
        student = self.validate(row, logs=logs)
        self.assertEqual(student.race_ui, "Other")
        self.assertTrue(any("falling back to 'Other'" in msg for msg in logs))

    def test_school_required_case_insensitive_mapped_and_active(self):
        row = self.base_row()
        row[self.cfg["columns"]["aeries_school"]] = ""
        self.assert_invalid(row, "School name is required")

        logs = []
        row = self.base_row()
        row[self.cfg["columns"]["aeries_school"]] = "cs - barbara phelps cs"
        student = self.validate(row, logs=logs)
        self.assertEqual(student.wai_school, "Barbara Phelps CHS (San Bernardino County Office of Education)")
        self.assertTrue(any("case-insensitive lookup" in msg for msg in logs))

        row = self.base_row()
        row[self.cfg["columns"]["aeries_school"]] = "CS - CHAFFEY NORTH CS"
        self.assert_invalid(row, "inactive/closed")

    def test_prepare_input_rows_splits_valid_and_invalid(self):
        valid = self.base_row()
        invalid = deepcopy(valid)
        invalid[EXCEL_ROW_NUM_KEY] = 3
        invalid[self.cfg["columns"]["ssid"]] = ""

        prepared, bad = prepare_input_rows([valid, invalid], self.cfg, today=TODAY)
        self.assertEqual(len(prepared), 1)
        self.assertEqual(len(bad), 1)
        self.assertEqual(bad[0].action, "SKIPPED_MISSING_SSID")

    def test_prepare_input_rows_rejects_duplicate_ssids(self):
        first = self.base_row()
        second = deepcopy(first)
        second[EXCEL_ROW_NUM_KEY] = 3
        second[self.cfg["columns"]["first_name"]] = "Duplicate"

        prepared, bad = prepare_input_rows([first, second], self.cfg, today=TODAY)
        self.assertEqual(len(prepared), 1)
        self.assertEqual(len(bad), 1)
        self.assertEqual(bad[0].action, "NEEDS_INPUT_DATA")
        self.assertIn("Duplicate State Student ID", bad[0].details)

    def test_address_patch_requires_expanded_fields(self):
        row = self.base_row()
        row[self.cfg["columns"]["street_address"]] = "1537 W 7th Street #101"
        row[self.cfg["columns"]["city"]] = "Upland"
        row[self.cfg["columns"]["state"]] = "ca"
        row[self.cfg["columns"]["zip"]] = 91786
        row[self.cfg["columns"]["phone_number"]] = "(626) 840-3963"
        row[self.cfg["columns"]["parent_name"]] = "Elly Rejuso"

        student = self.validate_address_patch(row)
        self.assertEqual(student.state, "CA")
        self.assertEqual(student.zip_code, "91786")
        self.assertEqual(student.parent_name, "Elly Rejuso")

        row = self.base_row()
        row[self.cfg["columns"]["street_address"]] = ""
        row[self.cfg["columns"]["city"]] = "Upland"
        row[self.cfg["columns"]["state"]] = "CA"
        row[self.cfg["columns"]["zip"]] = 91786
        row[self.cfg["columns"]["phone_number"]] = "(626) 840-3963"
        row[self.cfg["columns"]["parent_name"]] = "Elly Rejuso"
        self.assert_invalid_address_patch(row, "Street address is required")

    def assert_invalid_address_patch(self, row, expected_text):
        with self.assertRaises(ValueError) as cm:
            self.validate_address_patch(row)
        self.assertIn(expected_text, str(cm.exception))


class HeaderValidationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cfg = load_config("config/config.yaml")

    def write_workbook(self, headers, data_rows=None):
        tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp.close()
        path = Path(tmp.name)

        wb = Workbook()
        ws = wb.active
        ws.append(headers)
        for row in data_rows or []:
            ws.append(row)
        wb.save(path)
        return path

    def test_valid_headers_in_correct_locations_pass(self):
        headers = expected_input_headers(self.cfg)
        path = self.write_workbook(headers, [["A", "B", "1234567890", "05/06/2010", "M", "9", "Specific Learning Disability (SLD)", "Y", "White", "CS - BARBARA PHELPS CS"]])
        try:
            rows = read_excel_rows(str(path), cfg=self.cfg)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0][EXCEL_ROW_NUM_KEY], 2)
        finally:
            path.unlink(missing_ok=True)

    def test_name_sanitization_happens_while_reading_excel(self):
        headers = expected_input_headers(self.cfg)
        path = self.write_workbook(headers, [["Jayne (Jay)", "O\u2019Connor-Smith #2", "1234567890", "05/06/2010", "M", "9", "Specific Learning Disability (SLD)", "Y", "White", "CS - BARBARA PHELPS CS"]])
        logs = []
        try:
            rows = read_excel_rows(str(path), cfg=self.cfg, log_fn=logs.append)
            self.assertEqual(rows[0][self.cfg["columns"]["first_name"]], "Jayne Jay")
            self.assertEqual(rows[0][self.cfg["columns"]["last_name"]], "O'Connor-Smith")
            self.assertTrue(any("column A (First Name)" in msg and "removed unsupported name characters" in msg for msg in logs))
            self.assertTrue(any("column B (Last Name)" in msg and "removed unsupported name characters" in msg for msg in logs))
        finally:
            path.unlink(missing_ok=True)

    def test_wrong_header_location_fails(self):
        headers = expected_input_headers(self.cfg)
        headers[0], headers[1] = headers[1], headers[0]
        path = self.write_workbook(headers)
        try:
            with self.assertRaises(ValueError) as cm:
                read_excel_rows(str(path), cfg=self.cfg)
            self.assertIn("INPUT_HEADER_ERROR", str(cm.exception))
            self.assertIn("column A", str(cm.exception))
        finally:
            path.unlink(missing_ok=True)

    def test_unexpected_extra_header_fails(self):
        headers = expected_input_headers(self.cfg) + ["Extra"]
        path = self.write_workbook(headers)
        try:
            with self.assertRaises(ValueError) as cm:
                read_excel_rows(str(path), cfg=self.cfg)
            self.assertIn("expected 'Street address'", str(cm.exception))
        finally:
            path.unlink(missing_ok=True)

    def test_expanded_address_headers_pass_when_required(self):
        headers = expected_input_headers(self.cfg) + [
            self.cfg["columns"]["street_address"],
            self.cfg["columns"]["city"],
            self.cfg["columns"]["state"],
            self.cfg["columns"]["zip"],
            self.cfg["columns"]["phone_number"],
            self.cfg["columns"]["parent_name"],
        ]
        row = [
            "Kenneth Leandro",
            "Rejuso",
            "3448323742",
            "06/04/2007",
            "M",
            "12",
            "Autism (AUT)",
            "N",
            "Filipino",
            "SE - WE - UPLAND HS",
            "1537 W 7th Street #101",
            "Upland",
            "CA",
            "91786",
            "(626) 840-3963",
            "Elly Rejuso",
        ]
        path = self.write_workbook(headers, [row])
        try:
            rows = read_excel_rows(str(path), cfg=self.cfg, require_address_columns=True)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0][self.cfg["columns"]["street_address"]], "1537 W 7th Street #101")
        finally:
            path.unlink(missing_ok=True)

    def test_address_headers_required_for_address_patch_mode(self):
        headers = expected_input_headers(self.cfg)
        path = self.write_workbook(headers)
        try:
            with self.assertRaises(ValueError) as cm:
                read_excel_rows(str(path), cfg=self.cfg, require_address_columns=True)
            self.assertIn("missing required address headers", str(cm.exception))
        finally:
            path.unlink(missing_ok=True)

    def test_audit_logger_writes_temp_file_and_is_cleaned_up(self):
        headers = expected_input_headers(self.cfg)
        headers[0] = "First\u00a0Name"
        row = [
            "Ja\u2019Kyren",
            "Smith",
            "1234567890",
            "05/06/2010",
            "f",
            "9",
            "Specific Learning Disability (SLD)",
            "n",
            "Not In Mapping",
            "cs - barbara phelps cs",
        ]
        workbook_path = self.write_workbook(headers, [row])
        dev_logs = []

        tmpdir = Path("output") / f"test_audit_{uuid.uuid4().hex}"
        try:
            tmpdir.mkdir(parents=True, exist_ok=True)
            audit_path = tmpdir / "logs" / "input_sanitization.log"
            audit_log = make_input_audit_logger(str(audit_path), "test-run", log_fn=dev_logs.append)

            rows = read_excel_rows(str(workbook_path), cfg=self.cfg, log_fn=audit_log)
            prepared, invalid = prepare_input_rows(rows, self.cfg, validation_log_fn=audit_log, today=TODAY)

            self.assertEqual(len(prepared), 1)
            self.assertEqual(invalid, [])
            self.assertTrue(audit_path.exists())

            text = audit_path.read_text(encoding="utf-8")
            self.assertIn("[test-run]", text)
            self.assertIn("Excel sanitized row 1, column A", text)
            self.assertIn("normalized smart apostrophes", text)
            self.assertIn("Gender 'f' normalized to 'F'", text)
            self.assertIn("EthCd 'n' normalized to 'N'", text)
            self.assertIn("falling back to 'Other'", text)
            self.assertIn("case-insensitive lookup", text)
            self.assertTrue(any("falling back to 'Other'" in msg for msg in dev_logs))
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

        self.assertFalse(audit_path.exists())
        workbook_path.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
