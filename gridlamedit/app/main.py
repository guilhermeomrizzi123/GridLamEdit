"""
Application entry point for the GridLamEdit GUI.

Running ``python -m gridlamedit.app.main`` will launch a minimal PySide6
window. Additional modules (models, services, ui) can build on this later.
"""

from __future__ import annotations

import sys
from typing import Optional

from PySide6.QtWidgets import QApplication

if __package__ in (None, ""):
    # Running as a script: ensure the project root is on sys.path for absolute imports.
    from pathlib import Path

    project_root = Path(__file__).resolve().parents[2]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    from gridlamedit.app.main_window import MainWindow
else:
    from .main_window import MainWindow

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

    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    screen = app.primaryScreen()
    if screen is not None:
        geometry = screen.availableGeometry()
        frame = window.frameGeometry()
        frame.moveCenter(geometry.center())
        window.move(frame.topLeft())

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
