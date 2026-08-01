"""
riskreporting
=============

Generateur Python de reporting risque middle-office (VaR, Expected Shortfall,
tracking error, beta, drawdown...) produisant un classeur Excel complet et
directement ouvrable, avec formules natives, mise en forme conditionnelle et
graphiques.

Ce package sert aussi de reference de validation croisee pour les UDF VBA du
dossier ``vba/`` du meme depot (voir ``examples/02_python_vs_vba.py``).
"""

from riskreporting import bloomberg_stub, data, metrics, validate, workbook

__all__ = ["data", "metrics", "workbook", "validate", "bloomberg_stub"]

__version__ = "1.0.0"
