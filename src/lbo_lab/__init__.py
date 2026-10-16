"""lbo-lab: a leveraged buyout model with a Python engine and a live-formula Excel workbook."""

from .assumptions import Assumptions, Case
from .engine import ModelResult, run_model

__version__ = "1.0.0"
__all__ = ["Assumptions", "Case", "ModelResult", "run_model", "__version__"]
