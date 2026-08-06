"""Generateur Python de reporting risque middle-office (VaR, Expected Shortfall, tracking error, beta, drawdown) produisant un classeur Excel complet."""

from riskreporting import bloomberg_stub, data, metrics, validate, workbook

__all__ = ["data", "metrics", "workbook", "validate", "bloomberg_stub"]

__version__ = "1.0.0"
