"""
===============================================================================
FAPI PET Patient Database Processing
===============================================================================

Author: Robin Gutsche
Institute: Department of Nuclear Medicine, FAU Erlangen-Nürnberg
Email: robin.gutsche@fau.de

Description:
    Reads the FAPI PET patient database, standardizes selected indications,
    identifies follow-up examinations and creates formatted Excel tables.

Version: 1.0
Date: 2026-08-05

Copyright (c) 2026 Robin Gutsche

This script was developed for scientific research.
Please contact the author before redistribution or modification.
===============================================================================
"""

from processing import run_pipeline


def main() -> None:
    """Startet die vollständige Datenauswertung."""
    run_pipeline()


if __name__ == "__main__":
    main()
