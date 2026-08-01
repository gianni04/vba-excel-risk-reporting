"""
test_bloomberg_stub.py
=======================

Verifie que le simulateur local Bloomberg (riskreporting.bloomberg_stub)
renvoie des types et des formes de donnees coherents avec l'usage attendu par
le classeur (BDP/BDH/BDS), sans jamais appeler de service reseau.
"""

from __future__ import annotations

import pandas as pd
import pytest

from riskreporting import bloomberg_stub as bbg


def test_bdp_renvoie_un_dataframe_indexe_par_ticker():
    df = bbg.bdp(["AAPL US Equity", "FR0000131104 Equity"], ["PX_LAST", "CRNCY"])
    assert isinstance(df, pd.DataFrame)
    assert list(df.index) == ["AAPL US Equity", "FR0000131104 Equity"]
    assert list(df.columns) == ["PX_LAST", "CRNCY"]


def test_bdp_accepte_ticker_et_champ_scalaires():
    df = bbg.bdp("AAPL US Equity", "PX_LAST")
    assert df.shape == (1, 1)
    assert isinstance(df.iloc[0, 0], float)


def test_bdp_champ_inconnu_leve_erreur_dediee():
    with pytest.raises(bbg.BloombergIndisponibleError):
        bbg.bdp("AAPL US Equity", "CHAMP_INEXISTANT")


def test_bdp_est_deterministe():
    df1 = bbg.bdp("AAPL US Equity", "PX_LAST", graine=123)
    df2 = bbg.bdp("AAPL US Equity", "PX_LAST", graine=123)
    pd.testing.assert_frame_equal(df1, df2)


def test_bdh_serie_temporelle_forme_attendue():
    df = bbg.bdh("AAPL US Equity", "PX_LAST", "2026-01-01", "2026-01-31")
    assert isinstance(df, pd.DataFrame)
    assert df.index.name == "date"
    assert (df["PX_LAST"] > 0).all()
    assert len(df) > 0


def test_bdh_plusieurs_tickers_et_champs_multiindex():
    df = bbg.bdh(["AAPL US Equity", "MSFT US Equity"], ["PX_LAST", "VOLATILITY_90D"], "2026-01-01", "2026-01-31")
    assert isinstance(df.columns, pd.MultiIndex)
    assert set(df.columns.get_level_values("ticker")) == {"AAPL US Equity", "MSFT US Equity"}


def test_bdh_champ_inconnu_leve_erreur():
    with pytest.raises(bbg.BloombergIndisponibleError):
        bbg.bdh("AAPL US Equity", "CHAMP_INEXISTANT", "2026-01-01", "2026-01-31")


def test_bds_composition_indice_pourcentages_sensés():
    df = bbg.bds("SXXP Index", "INDX_MWEIGHT", n_lignes=10)
    assert isinstance(df, pd.DataFrame)
    assert len(df) == 10
    assert "Percentage Weight" in df.columns
    assert (df["Percentage Weight"] >= 0).all()


def test_bds_champ_non_supporte_leve_erreur():
    with pytest.raises(bbg.BloombergIndisponibleError):
        bbg.bds("SXXP Index", "AUTRE_CHAMP_BULK")


def test_documentation_champs_couvre_tous_les_champs_documentes():
    df = bbg.documentation_champs()
    assert set(df["champ"]) == set(bbg.CHAMPS_DOCUMENTES.keys())
