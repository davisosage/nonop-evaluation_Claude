"""Lateral-length scaling of a monthly type curve (Davis's rule, Nov 2026).

A type curve built for one lateral length is moved to another like this:
  * within +/- BAND ft (default 2,500) of the length it was built for: linear (volumes x L_to / L_from)
  * further away: the EUR still scales linearly, but the shape does not
      - months 1..EARLY_MONTHS (default 8, ramp included) get EARLY_PCT (default 90%) of the linear scaling
        when going to a longer lateral, and linear / EARLY_PCT when going to a shorter one (mirror image)
      - from month EARLY_MONTHS + 1 the decline is re-solved so the 360-month EUR lands exactly on the linear EUR:
        the tail keeps its b and starts from the scaled month-9 rate (no jump), and the hyperbolic decline is
        made shallower (longer lateral) or steeper (shorter lateral). The 7% terminal decline is kept.
        In practice the decline segment is stretched in time by a factor s:
            q_new(t) = r * e * q(9 + (t - 9) * s)                       while 9 + (t - 9) * s <= Tsw
            q_new(t) = r * e * q(Tsw) * g ** ((9 + (t - 9) * s - Tsw) / s)  after (terminal, g = monthly ratio)
        (for a hyperbolic, a time stretch is exactly "same b, Di x s"), with s solved so sum(q_new) = r * sum(q).

Used by build_tc.py (offsets -> the deal's lateral -> the 10,000 ft basis), engine.py (10,000 ft basis -> each
well's lateral) and lookback.py. The workbook's 'Lateral Curves' tab does the same arithmetic with the s, Tsw and g
values written by write_model.py.
"""
import numpy as np

BASIS = 10000.0          # type curves are kept per 10,000 ft
BAND = 2500.0            # +/- ft within which scaling is linear
EARLY_MONTHS = 8         # months held to EARLY_PCT of linear
EARLY_PCT = 0.90
N = 360
F_TERM = 0.93 ** (1 / 12)      # 7% effective terminal decline, monthly ratio


def params(band=None, months=None, pct=None):
    return dict(band=BAND if band is None else float(band), months=EARLY_MONTHS if months is None else int(months),
                pct=EARLY_PCT if pct is None else float(pct))


def terminal(c):
    """(Tsw, g): first month (1-based) from which the curve declines at a constant monthly ratio g (the terminal
    exponential), or (360, last ratio) when it never gets there inside 360 months."""
    c = np.asarray(c, float)
    n = len(c)
    if n < 3 or c[-1] <= 0 or c[-2] <= 0:
        return n, F_TERM
    rho = c[1:] / np.where(c[:-1] > 0, c[:-1], np.nan)      # rho[k] = c[k+2] / c[k+1] in 1-based months
    g = rho[-1]
    if not np.isfinite(g) or abs(g - F_TERM) > 1e-6:
        return n, float(g) if np.isfinite(g) and 0 < g < 1 else F_TERM
    ok = np.abs(rho - F_TERM) <= 1e-6
    k = len(ok)
    while k > 0 and ok[k - 1]:
        k -= 1
    return k + 1, F_TERM                                     # 1-based month where the terminal run starts


def _interp(c, pos):
    i = np.floor(pos).astype(int)
    fr = pos - i
    i0 = np.clip(i, 1, len(c)) - 1
    i1 = np.clip(i + 1, 1, len(c)) - 1
    return c[i0] * (1 - fr) + c[i1] * fr


def _tail(c, s, m1, tsw, g, n):
    t = np.arange(m1, n + 1, dtype=float)
    pos = m1 + (t - m1) * s
    v = np.where(pos <= tsw, _interp(c, np.minimum(pos, tsw)), c[tsw - 1] * g ** ((pos - tsw) / s))
    return v


def scale_curve(c, l_from, l_to, band=None, months=None, pct=None):
    """Move a monthly curve from lateral l_from to l_to. Returns (curve, info) where info has
    mode ('linear' / 'shaped'), r, e (early factor relative to r), s (decline stretch), tsw, g."""
    P = params(band, months, pct)
    c = np.asarray(c, float)
    n = len(c)
    if not l_from or not l_to or l_from <= 0 or l_to <= 0:
        return np.zeros(n), dict(mode="linear", r=0.0, e=1.0, s=1.0, tsw=n, g=F_TERM)
    r = l_to / l_from
    if abs(l_to - l_from) <= P["band"] + 1e-9 or c.sum() <= 0 or n <= P["months"] + 1:
        return c * r, dict(mode="linear", r=r, e=1.0, s=1.0, tsw=n, g=F_TERM)
    M, m1 = P["months"], P["months"] + 1
    e = P["pct"] if r > 1 else 1 / P["pct"]
    tsw, g = terminal(c)
    tsw = max(tsw, m1)
    need = c.sum() / e - c[:M].sum()                         # tail volume (in source units) that hits the linear EUR
    lo, hi = np.log(1e-3), np.log(1e3)
    for _ in range(100):                                     # tail volume falls as s rises
        mid = (lo + hi) / 2
        if _tail(c, np.exp(mid), m1, tsw, g, n).sum() > need:
            lo = mid
        else:
            hi = mid
    s = float(np.exp((lo + hi) / 2))
    out = np.concatenate([c[:M], _tail(c, s, m1, tsw, g, n)]) * r * e
    return out, dict(mode="shaped", r=r, e=e, s=s, tsw=int(tsw), g=float(g))


def scale_data(d, info, months=None):
    """Apply a scaling found on a curve (scale_curve info) to a well's reported monthly volumes. The decline part is
    stretched in time the same way (no terminal handling: reported data stops well before it)."""
    d = np.asarray(d, float)
    if info["mode"] == "linear":
        return d * info["r"]
    M = EARLY_MONTHS if months is None else int(months)
    m1, n = M + 1, len(d)
    if n <= M:
        return d * info["r"] * info["e"]
    s = info["s"]
    t_max = m1 + int(np.floor((n - m1) / s + 1e-9))
    t = np.arange(m1, t_max + 1, dtype=float)
    tail = _interp(d, m1 + (t - m1) * s)
    return np.concatenate([d[:M], tail]) * info["r"] * info["e"]
