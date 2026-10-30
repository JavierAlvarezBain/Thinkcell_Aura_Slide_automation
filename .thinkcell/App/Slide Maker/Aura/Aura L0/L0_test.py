from __future__ import annotations

"""Compatibility launcher for the accepted modular Aura L0 workflow.

Existing commands that call ``L0_test.py`` continue to work.  The implementation
now lives in dedicated L0 modules and ``run_aura_l0.py``.
"""

from run_aura_l0 import main


if __name__ == "__main__":
    main()
