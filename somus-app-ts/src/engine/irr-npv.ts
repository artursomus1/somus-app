/**
 * irr-npv.ts - Core financial math functions
 * Port of nasa_engine_hd.py helper functions
 * Somus Capital - Mesa de Produtos
 *
 * All rates are internally as decimals (0.01 = 1%).
 */

/**
 * Net Present Value of a series of monthly cashflows.
 * @param rateMonthly - monthly discount rate (decimal)
 * @param cashflows - array of cashflows, index 0 = period 0
 */
export function npv(rateMonthly: number, cashflows: number[]): number {
  if (rateMonthly === 0) {
    let sum = 0;
    for (let i = 0; i < cashflows.length; i++) {
      sum += cashflows[i];
    }
    return sum;
  }
  let total = 0.0;
  for (let t = 0; t < cashflows.length; t++) {
    try {
      const d = Math.pow(1 + rateMonthly, t);
      if (!isFinite(d) || d === 0) break;
      total += cashflows[t] / d;
    } catch {
      break;
    }
  }
  return total;
}

/**
 * Internal Rate of Return (monthly) via Newton-Raphson with bisection fallback.
 * @param cashflows - array of cashflows
 * @param guess - initial guess (default 0.01)
 * @param tol - tolerance (default 1e-9)
 * @param maxIter - max iterations (default 500)
 */
export function irr(
  cashflows: number[],
  guess: number = 0.01,
  tol: number = 1e-9,
  maxIter: number = 500
): number {
  // Try Newton-Raphson first
  let rate = guess;
  for (let iter = 0; iter < maxIter; iter++) {
    let npvVal = 0.0;
    let dnpv = 0.0;
    let ok = true;
    for (let t = 0; t < cashflows.length; t++) {
      const d = Math.pow(1 + rate, t);
      if (Math.abs(d) < 1e-30 || !isFinite(d)) {
        ok = false;
        break;
      }
      npvVal += cashflows[t] / d;
      if (t > 0) {
        dnpv -= (t * cashflows[t]) / Math.pow(1 + rate, t + 1);
      }
    }
    if (!ok) break;
    if (Math.abs(npvVal) < tol) return rate;
    if (Math.abs(dnpv) < 1e-15) break;
    const step = npvVal / dnpv;
    rate -= step;
    if (rate <= -1) {
      rate = 0.0001;
    }
  }

  // Fallback: bisection
  return irrBisection(cashflows, -0.5, 2.0, tol, maxIter > 500 ? maxIter : 1000);
}

/**
 * IRR via bisection (more robust for difficult cash flows).
 */
function irrBisection(
  cashflows: number[],
  lo: number = -0.5,
  hi: number = 2.0,
  tol: number = 1e-9,
  maxIter: number = 1000
): number {
  // Use conservative bounds for long flows to avoid overflow
  if (hi > 1.0 && cashflows.length > 100) {
    hi = 0.5;
  }
  if (lo < -0.9 && cashflows.length > 100) {
    lo = -0.3;
  }

  let npvLo = npv(lo, cashflows);
  let npvHi = npv(hi, cashflows);

  // Search for valid bounds
  if (npvLo * npvHi > 0) {
    const testLos = [-0.3, -0.1, -0.01, 0.0];
    const testHis = [0.05, 0.1, 0.2, 0.5];
    let found = false;
    for (const testLo of testLos) {
      for (const testHi of testHis) {
        const nLo = npv(testLo, cashflows);
        const nHi = npv(testHi, cashflows);
        if (nLo * nHi < 0) {
          lo = testLo;
          hi = testHi;
          npvLo = nLo;
          npvHi = nHi;
          found = true;
          break;
        }
      }
      if (found) break;
    }
    if (!found) return 0.0; // No solution
  }

  for (let iter = 0; iter < maxIter; iter++) {
    const mid = (lo + hi) / 2;
    const npvMid = npv(mid, cashflows);
    if (Math.abs(npvMid) < tol || (hi - lo) / 2 < tol) {
      return mid;
    }
    if (npvMid * npvLo < 0) {
      hi = mid;
      npvHi = npvMid;
    } else {
      lo = mid;
      npvLo = npvMid;
    }
  }
  return (lo + hi) / 2;
}

/**
 * PMT: Payment calculation (Price table / annuity).
 * @param rate - periodic interest rate (decimal)
 * @param nper - number of periods
 * @param pv - present value (principal)
 * @returns payment amount (positive)
 */
export function pmt(rate: number, nper: number, pv: number): number {
  if (rate === 0) {
    return nper > 0 ? pv / nper : 0.0;
  }
  const factor = Math.pow(1 + rate, nper);
  return (pv * rate * factor) / (factor - 1);
}

/**
 * Convert monthly rate to annual equivalent.
 * @param rm - monthly rate (decimal)
 */
export function annualFromMonthly(rm: number): number {
  return Math.pow(1 + rm, 12) - 1;
}

/**
 * Convert annual rate to monthly equivalent.
 * @param ra - annual rate (decimal)
 */
export function monthlyFromAnnual(ra: number): number {
  if (ra <= -1) return 0.0;
  return Math.pow(1 + ra, 1 / 12) - 1;
}

/**
 * GoalSeek: Generic solver using bisection + Newton numerical derivative.
 *
 * Finds x such that targetFunc(x) ~ targetValue.
 *
 * @param targetFunc - function f(x) -> number
 * @param targetValue - desired output value
 * @param lo - lower bound for bisection
 * @param hi - upper bound for bisection
 * @param tolerance - precision (default 1e-9)
 * @param maxIter - max iterations (default 500)
 * @returns x found
 */
export function goalSeek(
  targetFunc: (x: number) => number,
  targetValue: number,
  lo: number,
  hi: number,
  tolerance: number = 1e-9,
  maxIter: number = 500
): number {
  const func = (x: number) => targetFunc(x) - targetValue;

  let fLo = func(lo);
  let fHi = func(hi);

  if (fLo * fHi > 0) {
    // Try expanding bounds
    let expanded = false;
    let eLo = lo;
    let eHi = hi;
    for (let i = 0; i < 20; i++) {
      eLo *= 0.5;
      eHi *= 2.0;
      fLo = func(eLo);
      fHi = func(eHi);
      if (fLo * fHi <= 0) {
        lo = eLo;
        hi = eHi;
        expanded = true;
        break;
      }
    }
    if (!expanded) {
      // Fallback to Newton numerical
      return goalSeekNewton(func, (lo + hi) / 2, tolerance, maxIter);
    }
  }

  // Bisection
  for (let iter = 0; iter < maxIter; iter++) {
    const mid = (lo + hi) / 2;
    const fMid = func(mid);
    if (Math.abs(fMid) < tolerance || (hi - lo) / 2 < tolerance) {
      return mid;
    }
    if (fMid * fLo < 0) {
      hi = mid;
      fHi = fMid;
    } else {
      lo = mid;
      fLo = fMid;
    }
  }
  return (lo + hi) / 2;
}

/**
 * Newton numerical with finite-difference derivative (internal fallback).
 */
function goalSeekNewton(
  func: (x: number) => number,
  guess: number,
  tolerance: number,
  maxIter: number
): number {
  let x = guess;
  let h = Math.max(Math.abs(x) * 1e-6, 1e-10);
  for (let iter = 0; iter < maxIter; iter++) {
    const fx = func(x);
    if (Math.abs(fx) < tolerance) return x;
    const fxh = func(x + h);
    const deriv = (fxh - fx) / h;
    if (Math.abs(deriv) < 1e-15) {
      h *= 10;
      continue;
    }
    const step = fx / deriv;
    x -= step;
    h = Math.max(Math.abs(x) * 1e-6, 1e-10);
  }
  return x;
}
