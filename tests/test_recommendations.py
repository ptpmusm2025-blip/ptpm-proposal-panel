"""Regression tests for the transparent panel recommendation rules."""

import unittest

import pandas as pd

import app


def appointment(date, student, supervisor, panel):
    return {
        "Proposal Year": 2026,
        "Date": pd.Timestamp(date) if date else pd.NaT,
        "Student": student,
        "Supervisor": supervisor,
        "Panel Member": panel,
        "Panel Category": app.CATEGORY_INTERNAL,
        "Panel Code": panel,
    }


class RecommendationTests(unittest.TestCase):
    def setUp(self):
        rows = []
        # WAJ has many students; CKT is deliberately used often overall and
        # repeatedly for WAJ, while ABN has zero historical appointments.
        for index in range(6):
            rows.append(appointment(f"2026-0{index + 1}-10", f"Student {index}", "WAJ", "CKT"))
        rows.extend([
            appointment("2026-04-01", "Student 7", "WAJ", "INU"),
            appointment("2026-05-01", "Student 8", "NJ", "CKT"),
            appointment("2026-05-15", "Student 9", "MS", "CKT"),
            appointment("2026-06-01", "Student 10", "NJ", "MTA"),
        ])
        self.data = pd.DataFrame(rows)

    def test_supervisor_and_manual_exclusions(self):
        pool = app.generate_eligible_lecturer_pool(self.data, "WAJ", ["INU"], ["NJ"])
        self.assertNotIn("WAJ", pool)
        self.assertNotIn("INU", pool)
        self.assertNotIn("NJ", pool)

    def test_zero_and_high_frequency_lecturers(self):
        pool = app.generate_eligible_lecturer_pool(self.data, "WAJ")
        stats, dates_available = app.calculate_lecturer_statistics(
            self.data, pool, "WAJ", pd.Timestamp("2026-06-30")
        )
        scored, _ = app.calculate_recommendation_scores(stats, dates_available)
        abn = scored.loc[scored["Lecturer Code"] == "ABN"].iloc[0]
        ckt = scored.loc[scored["Lecturer Code"] == "CKT"].iloc[0]
        self.assertEqual(abn["Total Proposal Appointments"], 0)
        self.assertGreater(abn["Recommendation Score"], ckt["Recommendation Score"])
        self.assertEqual(ckt["Appointments for Selected Supervisor"], 6)

    def test_missing_dates_redistributes_weight(self):
        missing_dates = self.data.copy()
        missing_dates["Date"] = pd.NaT
        pool = app.generate_eligible_lecturer_pool(missing_dates, "WAJ")
        stats, dates_available = app.calculate_lecturer_statistics(
            missing_dates, pool, "WAJ", pd.Timestamp("2026-06-30")
        )
        scored, components = app.calculate_recommendation_scores(stats, dates_available)
        self.assertFalse(dates_available)
        self.assertTrue(scored["Recommendation Score"].between(0, 100).all())
        recent = components.loc[components["Score Component"] == "Recent Workload"].iloc[0]
        self.assertEqual(recent["Applied Weight"], "0.0%")

    def test_many_and_few_supervisor_records(self):
        pool = app.generate_eligible_lecturer_pool(self.data, "WAJ")
        many = app.validate_recommendation_inputs(self.data, "WAJ", pool, 3)
        few = app.validate_recommendation_inputs(self.data, "MS", pool, 3)
        self.assertFalse(any("fewer than three" in warning for warning in many))
        self.assertTrue(any("fewer than three" in warning for warning in few))

    def test_selection_review_balance(self):
        stats, _ = app.calculate_lecturer_statistics(
            self.data, ["CKT", "ABN"], "WAJ", pd.Timestamp("2026-06-30")
        )
        review = app.review_proposed_panel_selection(stats, ["CKT", "ABN"], "WAJ", [])
        self.assertEqual(review["indicator"], "Highly concentrated distribution")


if __name__ == "__main__":
    unittest.main()
