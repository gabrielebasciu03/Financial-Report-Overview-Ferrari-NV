# Financial-Report-Overview-Ferrari-NV
Implied Discount Factors &amp; Dividend Term Structure from Ferrari (RACE) Option Prices.  Extracting the risk-free discount curve and the implied dividend structure of Ferrari N.V. (RACE) from listed option prices, using only no-arbitrage relationships.
Overview

# Implied Discount Factors & Dividend Term Structure from Ferrari (RACE) Option Prices

Extracting the risk-free discount curve and the implied dividend structure of **Ferrari N.V. (RACE)** from listed option prices, using only no-arbitrage relationships.

## Overview

Market data as of **24 March 2026** (spot S₀ = 329.93). All option prices are **mid-prices** (average of bid and ask). Listed options are American but treated as European, an approximation discussed below.

## Methodology

### 1. Discount factor via box spread / put-call parity

Using two strikes (K₁ = 300, K₂ = 360, expiry 15 May 2026):

```
D(0,T) = [C(K₁) − P(K₁) − C(K₂) + P(K₂)] / (K₂ − K₁)
```

The risk-free rate follows as `r = −ln(D) / T`.

### 2. Forward price from ATM parity

With K_ATM = 330:

```
F(0,T) = K + [C − P] / D(0,T)
```

### 3. Implied dividends

- Discrete dividend: `div = S₀ · e^(rT) − F(0,T)`
- Proportional yield: `q = r − ln(F/S₀) / T`

### 4. Term structure

The procedure is repeated for 1, 2, 3, 6 and 9-month maturities.

## Results

| Maturity | D₀(T) | r | F₀(T) | Div | q |
|----------|-------|--------|--------|-------|--------|
| 1M | 0.9981 | 2.25% | 326.95 | 3.60 | 13.17% |
| 2M | 0.9992 | 0.50% | 327.05 | 3.16 | 5.77% |
| 3M | 0.9883 | 4.69% | 328.25 | 5.58 | 6.82% |
| 6M | 0.9883 | 2.35% | 330.90 | 2.92 | 1.78% |
| 9M | 0.9450 | 7.54% | 334.60 | 14.53 | 6.00% |

## Key Findings

- Short-dated discount factors are consistent with a low-rate environment.
- Implied rates rise with maturity (up to ~7.5% at 9M), likely reflecting bid-ask noise amplified by thinner liquidity rather than a true term premium.
- Forward prices increase monotonically (~327 → ~335).
- Implied dividends are non-monotonic; the large 9M value is consistent with Ferrari's dividend cycle.
- Short-dated q is numerically unstable (division by small T), so absolute dividends are the more reliable measure.

## Limitations

- American-vs-European approximation (early-exercise premium biases short maturities).
- Wide bid-ask spreads at long tenors.
- Single-day snapshot, no smoothing or curve fitting.

## Files

- `basciu gabriele report ferrari.xlsm`: Excel calculations
- Report (PDF): methodology, results and discussion

## Tech

Excel (VBA-enabled workbook) and financial derivatives theory.
