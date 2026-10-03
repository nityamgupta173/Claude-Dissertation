"""Shared estimation tools: TWFE via linearmodels.PanelOLS + wild cluster bootstrap-t (Webb weights, null imposed)."""
import numpy as np
import pandas as pd
from scipy import stats
from linearmodels.panel import PanelOLS

B_REPS = 19999
SEED = 20231116
MACRO = ["Repo_Rate_Pct", "Bond_Yield_5Y_Pct", "GDP_Growth_YoY_Pct"]

# ------------------------------------------------------------------ inference engine
def _basis(M, tol=1e-9):
    u, s, _ = np.linalg.svd(M, full_matrices=False)
    return u[:, s > tol * s[0]]


def wcb_t(y, X_test, W, groups, B=B_REPS, seed=SEED):
    """Wild cluster bootstrap-t (Cameron-Gelbach-Miller 2008), null imposed (WCR),
    Webb 6-point weights. Tests every column of X_test one at a time
    (other test columns stay in the model). Returns list of dict."""
    rng = np.random.default_rng(seed)
    g_idx, G = pd.factorize(groups)
    G = len(G)
    Gm = np.zeros((len(y), G)); Gm[np.arange(len(y)), g_idx] = 1
    webb = np.array([-np.sqrt(1.5), -1, -np.sqrt(.5), np.sqrt(.5), 1, np.sqrt(1.5)])
    Wt = webb[rng.integers(0, 6, size=(G, B))]            # G x B
    Wobs = Wt[g_idx, :]                                     # n x B
    res = []
    n = len(y)
    for j in range(X_test.shape[1]):
        others = np.delete(X_test, j, axis=1)
        Wj = np.column_stack([W, others]) if others.shape[1] else W
        Q = _basis(Wj)
        d = X_test[:, j]
        dt = d - Q @ (Q.T @ d)
        dd = dt @ dt
        K = Q.shape[1] + 1
        adj = G / (G - 1) * (n - 1) / (n - K)

        def tstat(ytil):                                   # ytil: n x m
            b = dt @ ytil / dd
            e = ytil - np.outer(dt, b)
            S = Gm.T @ (dt[:, None] * e)
            se = np.sqrt(adj * (S ** 2).sum(0)) / dd
            return b, se
        yt = y - Q @ (Q.T @ y)
        b, se = tstat(yt[:, None])
        b, se = b[0], se[0]
        t = b / se
        # restricted model: y has beta_j = 0 -> residual u_r = M_Wj y
        ur = yt
        ystar = Wobs * ur[:, None]
        ystar_t = ystar - Q @ (Q.T @ ystar)
        bs, ses = tstat(ystar_t)
        ts = bs / ses
        p = (np.abs(ts) >= abs(t)).mean()
        crit = np.quantile(np.abs(ts), 0.95)
        p_t = 2 * stats.t.sf(abs(t), G - 1)
        res.append(dict(beta=b, se_cr1=se, t=t, p_wcb=p, p_cr1_t=p_t, G=G,
                        ci_lo=b - crit * se, ci_hi=b + crit * se))
    return res


def wcb_F(y, X_test, W, groups, B=B_REPS, seed=SEED):
    """Joint test that all X_test coefficients = 0: classical SSR-based F with
    null-imposed wild cluster bootstrap p-value (cluster-robust Wald is rank
    deficient when #restrictions >= #clusters)."""
    rng = np.random.default_rng(seed)
    g_idx, uniq = pd.factorize(groups)
    G = len(uniq); n = len(y)
    webb = np.array([-np.sqrt(1.5), -1, -np.sqrt(.5), np.sqrt(.5), 1, np.sqrt(1.5)])
    Wobs = webb[rng.integers(0, 6, size=(G, B))][g_idx, :]
    Qr = _basis(W)
    Qu = _basis(np.column_stack([W, X_test]))
    q = Qu.shape[1] - Qr.shape[1]
    df2 = n - Qu.shape[1]

    def F(Y):
        sr = (Y ** 2).sum(0) - ((Qr.T @ Y) ** 2).sum(0)
        su = (Y ** 2).sum(0) - ((Qu.T @ Y) ** 2).sum(0)
        return ((sr - su) / q) / (su / df2)
    y1 = y[:, None]
    F0 = F(y1)[0]
    ur = y - Qr @ (Qr.T @ y)
    Fs = F(Wobs * ur[:, None])
    return dict(F=F0, q=q, df2=df2, p_classical=stats.f.sf(F0, q, df2), p_wcb=(Fs >= F0).mean())


# ------------------------------------------------------------------ model runner
def design(data, dv, treat, ctrl, extra_x=(), min_obs=8):
    cols = [dv] + list(treat) + list(ctrl) + list(extra_x)
    d = data.dropna(subset=cols).copy()
    # keep only firms with >= 8 usable obs
    d = d.groupby("Company").filter(lambda x: len(x) >= min_obs)
    return d


def run(label, data, dv, treat, ctrl=MACRO, boot=True, min_obs=8):
    d = design(data, dv, treat, ctrl, min_obs=min_obs)
    p = d.set_index(["Company", "Date_End"])
    X = p[list(treat) + list(ctrl)]
    mod = PanelOLS(p[dv], X, entity_effects=True, time_effects=True, drop_absorbed=True)
    fit = mod.fit(cov_type="clustered", cluster_entity=True)
    kept = list(fit.params.index)
    absorbed = [c for c in list(treat) + list(ctrl) if c not in kept]
    # LSDV objects for the bootstrap / adj R2
    ent = pd.get_dummies(d.Company).values.astype(float)
    tim = pd.get_dummies(d.Date_End).values.astype(float)[:, 1:]
    Wfe = np.column_stack([ent, tim] + [d[c].values for c in kept if c not in treat])
    Xt = d[list(treat)].values
    y = d[dv].values
    out = wcb_t(y, Xt, Wfe, d.Company.values) if boot else None
    # check against PanelOLS point estimates
    for j, c in enumerate(treat):
        assert abs(fit.params[c] - out[j]["beta"]) < 1e-6, (label, c, fit.params[c], out[j]["beta"])
    Qall = _basis(np.column_stack([Wfe, Xt]))
    n, K = len(y), Qall.shape[1]
    resid = y - Qall @ (Qall.T @ y)
    r2_lsdv = 1 - (resid ** 2).sum() / ((y - y.mean()) ** 2).sum()
    adj = 1 - (1 - r2_lsdv) * (n - 1) / (n - K)
    return dict(label=label, dv=dv, treat=list(treat), ctrl=list(ctrl), absorbed=absorbed, res=out,
                fit=fit, n=n, firms=d.Company.nunique(), adj_r2=adj, within_r2=float(fit.rsquared),
                sample=d, firm_ctrl=[c for c in kept if c not in treat],
                coef_other={c: (fit.params[c], fit.std_errors[c]) for c in kept if c not in treat})


def stars(p):
    return "***" if p < .01 else "**" if p < .05 else "*" if p < .10 else ""


