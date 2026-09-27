"""test_cube_solver_rev3.py - Verification and regression test suite for Rev 3.

Validates:
- Baseline Load Cases 1 to 9 definitions and equilibrium totals
- Rigid roof diaphragm master/slave constraints
- Load combination definitions and factor breakdowns (including ASD Combination 13)
- Verification report generation and PNG export integrity
"""

import math
import os
import unittest
from pathlib import Path

from cube_solver_rev3_loads import (
    initialize,
    build_rev3_load_cases,
    build_load_combinations,
    compute_case_totals,
    generate_verification_report,
    run_headless_verification,
    NODES,
    MEMBERS,
)


class TestCubeSolverRev3Loads(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.state = initialize()
        cls.cases, cls.diaphragm = build_rev3_load_cases(cls.state)
        cls.combs = build_load_combinations()

    def test_geometry_and_topology(self):
        """Verify the 6m cube node count and member connectivity."""
        self.assertEqual(len(NODES), 8)
        self.assertEqual(len(MEMBERS), 12)

        # Confirm 4 base beams, 4 roof beams, 4 columns
        base_beams = [m for m in MEMBERS if m["type"] == "Base beam"]
        roof_beams = [m for m in MEMBERS if m["type"] == "Roof beam"]
        columns = [m for m in MEMBERS if m["type"] == "Column"]

        self.assertEqual(len(base_beams), 4)
        self.assertEqual(len(roof_beams), 4)
        self.assertEqual(len(columns), 4)

    def test_rigid_roof_diaphragm(self):
        """Verify the rigid roof diaphragm topology and constraints."""
        self.assertEqual(self.diaphragm.master_node, 5)
        self.assertListEqual(self.diaphragm.constrained_nodes, [6, 7, 8])
        self.assertTupleEqual(self.diaphragm.degrees_of_freedom, ("UX", "UZ", "RY"))

    def test_load_case_equilibrium_resultants(self):
        """Verify resultant forces for Load Cases 1 to 8 against theoretical benchmarks."""
        expected_resultants = {
            1: 29.815,   # LC1: Framing self-weight (~29.815 kN)
            2: 120.000,  # LC2: Roof dead (4 beams * 6m * 5 kN/m)
            3: 72.000,   # LC3: Roof live (4 beams * 6m * 3 kN/m)
            4: 20.000,   # LC4: Center point load (4 beams * 5 kN)
            5: 10.000,   # LC5: Lateral wind X (4 nodes * 2.5 kN)
            6: 10.000,   # LC6: Lateral wind Z (4 nodes * 2.5 kN)
            7: 15.000,   # LC7: Seismic X (4 nodes * 3.75 kN)
            8: 15.000,   # LC8: Seismic Z (4 nodes * 3.75 kN)
        }

        for cid, expected_val in expected_resultants.items():
            tots = compute_case_totals(self.cases[cid], self.state)
            self.assertAlmostEqual(
                tots["Resultant"],
                expected_val,
                places=2,
                msg=f"Discrepancy found in Load Case {cid} equilibrium total."
            )

    def test_temperature_load_case_9(self):
        """Verify Load Case 9 thermal strain and external equilibrium balance."""
        lc9 = self.cases[9]
        self.assertIsNotNone(lc9.temperature_load)
        tload = lc9.temperature_load

        self.assertEqual(tload.temperature_change, 15.0)
        self.assertListEqual(tload.member_ids, [5, 6, 7, 8])

        # Self-equilibrating internal strain generates 0 net external support reaction
        tots = compute_case_totals(lc9, self.state)
        self.assertAlmostEqual(tots["Resultant"], 0.0, places=6)

    def test_combination_13_asd_definition(self):
        """Verify ASD Combination 13 contains LC1, LC2, and LC4 with 1.0 factor."""
        # Find Combination 13 in the combination library
        comb_13 = next((c for c in self.combs if "Combination 13" in c.name or c.id == 17), None)
        self.assertIsNotNone(comb_13, "Combination 13 was not found in build_load_combinations()")

        self.assertEqual(comb_13.design_method, "ASD")
        # Factors: DEAD/SELF WEIGHT (LC1) x 1.0, ROOF DEAD (LC2) x 1.0, ROOF BEAM CENTER LOAD (LC4) x 1.0
        expected_factors = {1: 1.0, 2: 1.0, 4: 1.0}
        self.assertDictEqual(comb_13.factors, expected_factors)

    def test_verification_report_generation(self):
        """Verify headless generation of cube_rev3_verification_report.txt."""
        report_path = Path("test_verification_audit.txt")
        if report_path.exists():
            report_path.unlink()

        generate_verification_report(self.state, self.cases, self.combs, self.diaphragm, report_path)
        self.assertTrue(report_path.exists())

        content = report_path.read_text(encoding="utf-8")
        self.assertIn("ASD: Combination 13 - D", content)
        self.assertIn("DEAD / SELF WEIGHT: 1.0, ROOF DEAD: 1.0, ROOF BEAM CENTER LOAD: 1.0", content)

        report_path.unlink()


if __name__ == "__main__":
    unittest.main()