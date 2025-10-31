"""
Application entry point for the GridLamEdit GUI.

Running ``python -m gridlamedit.app.main`` will launch a minimal PySide6
window. Additional modules (models, services, ui) can build on this later.
"""

from __future__ import annotations

import sys
from typing import Optional

from PySide6.QtWidgets import QApplication, QWidget


def _get_or_create_app(argv: Optional[list[str]] = None) -> QApplication:
    """Return the existing QApplication or create a new one for this process."""
    existing_app = QApplication.instance()
    if existing_app is not None:
        return existing_app
    # Ensure argv defaults to sys.argv but remains modifiable for tests.
    return QApplication(argv if argv is not None else sys.argv)


def main(argv: Optional[list[str]] = None) -> int:
    """Launch the GridLamEdit window and return the Qt event loop exit code."""
    app = _get_or_create_app(argv)

    window = QWidget()
    window.setWindowTitle("GridLamEdit")
    window.resize(800, 600)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
