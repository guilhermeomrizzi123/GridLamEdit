"""GridLamEdit package root."""

from .core.paths import package_path, resource_path, is_frozen

APP_NAME = "GridLamEdit"
__version__ = "0.1.0"
__version_date__ = "2026-01-19"
__contact__ = "Guilherme Rizzi Contato +5512988508042"

__all__ = [
	"package_path",
	"resource_path",
	"is_frozen",
	"APP_NAME",
	"__version__",
	"__version_date__",
	"__contact__",
]
