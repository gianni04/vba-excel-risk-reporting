# VBA / Excel Risk Reporting

Excel risk report for an equity fund: VBA functions (UDFs) for risk metrics
and option pricing, a limit check sheet, and a Python script that builds the
workbook with native Excel formulas and checks the results.

![Python](https://img.shields.io/badge/Python-3.11-blue) ![VBA](https://img.shields.io/badge/VBA-Excel-217346) ![License](https://img.shields.io/badge/license-MIT-green) ![Tests](https://img.shields.io/badge/tests-46%20passed-brightgreen)

All data (25 positions, 260 days of NAV, Bloomberg fields) is synthetic,
generated with a fixed seed.

## Contents

```
vba/
  modRiskMetrics.bas    VaR (historical, parametric, Cornish-Fisher), ES, volatility,
                        EWMA vol, tracking error, beta, Sharpe, information ratio, max drawdown
  modOptions.bas        Black-Scholes price, delta, gamma, vega, theta, rho, implied vol
  modLimites.bas        limit checks (OK / warning / breach)
  modBloomberg.bas      builds BDP/BDH/BDS formulas, falls back to local data without a terminal
  modReporting.bas      summary sheet, charts, PDF export
  modDataUtils.bas      CSV import, sorting, business days
  clsPortefeuille.cls   portfolio class (exposures, concentration)
  modTests.bas          VBA unit tests (run inside Excel)
src/riskreporting/      Python version of the same metrics, workbook builder, checks
examples/               build the workbook, print reference values
tests/                  46 pytest tests
docs/exemple_reporting_risque.xlsx   generated workbook
```

The function names in the VBA modules are in French (for example
`VaRHistorique`, `VolImplicite`).

## Run

```bash
pip install -r requirements.txt
python examples/01_build_report.py    # builds the workbook
python examples/02_python_vs_vba.py   # reference values for the VBA tests
python -m pytest tests -q
```

To use the VBA functions: open the workbook, press Alt+F11, import the `.bas`
and `.cls` files from `vba/`, and save as `.xlsm`. Then for example
`=VaRHistorique(B2:B251;0,99)` in a cell. The VBA tests run with
`RunAllTests` in `modTests`.

The pytest suite tests the Python version and the generated workbook. The VBA
code itself runs only inside Excel.

## Workbook

| Sheet | Content |
|---|---|
| Summary | 14 risk indicators, all as native Excel formulas |
| Positions | 25 positions: sector, currency, rating, exposure, duration, beta |
| History | 260 days of fund and benchmark NAV, returns, drawdown |
| Limits | VaR, TE, concentration and sector limits, with traffic-light status |
| Charts | NAV vs benchmark, risk contribution by sector |

The workbook contains 1,162 native formulas (`SUMPRODUCT`, `XLOOKUP`,
`PERCENTILE.INC`, `STDEV.S`, `SUMIFS`, `SLOPE`, named ranges).

## Results (synthetic fund, Python version)

| Indicator | Value |
|---|---|
| Historical VaR 99%, 1 day | 1.93% |
| Parametric VaR 99%, 1 day | 1.93% |
| Cornish-Fisher VaR 99% | 1.85% |
| Expected Shortfall 97.5% | 1.93% |
| Annualised volatility | 13.1% |
| EWMA volatility (lambda 0.94), daily | 0.89% |
| Tracking error | 6.3% |
| Beta | 0.78 |
| Max drawdown | -9.2% |
| Largest issuer weight | 13.3% |

`examples/02_python_vs_vba.py` prints each value next to the matching Excel
formula, as a reference for the VBA functions. `modTests.bas` checks the VBA
Black-Scholes functions against the same reference (S = K = 100, T = 1,
r = 5%, vol = 20%): call 10.4506, put 5.5735, delta 0.6368.

![NAV and drawdown](docs/img/vl_drawdown.png)

## Limitations

- Synthetic data only.
- The sector risk contribution is a simple approximation (sum of absolute
  exposures, correlation of 1 between positions).
- Bloomberg functions need a terminal; without one, local data is used.
