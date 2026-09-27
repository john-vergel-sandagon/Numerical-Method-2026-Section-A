"""generate_report.py - Standalone report generation script.

Executes the model building and verification audit from cube_solver_rev3_loads
and writes the formatted 'cube_rev3_verification_report.txt'.
"""

from pathlib import Path
from cube_solver_rev3_loads import (
    initialize,
    build_rev3_load_cases,
    build_load_combinations,
    generate_verification_report,
)


def main():
    # 1. Initialize solver databases (materials, shapes, units)
    state = initialize()

    # 2. Build load cases (LC1 to LC9) and the rigid diaphragm constraint
    cases, diaphragm = build_rev3_load_cases(state)

    # 3. Build LRFD and ASD combinations (including Combination 13)
    combs = build_load_combinations()

    # 4. Generate the verification text report
    output_file = Path("cube_rev3_verification_report.txt")
    generate_verification_report(state, cases, combs, diaphragm, output_file)


if __name__ == "__main__":
    main()