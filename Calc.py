# -*- coding: utf-8 -*-
#================================================================================#
#  ECOLOG - Sistema Gerenciador de Banco de Dados para Levantamentos Ecológicos  #
#        ECOLOG - Database Management System for Ecological Surveys              #
#      Copyright (c) 1990-2014 Mauro J. Cavalcanti. All rights reserved.         #
#                                                                                #
#   Este programa é software livre; você pode redistribuí-lo e/ou                #
#   modificá-lo sob os termos da Licença Pública Geral GNU, conforme             #
#   publicada pela Free Software Foundation; tanto a versão 2 da                 #
#   Licença como (a seu critério) qualquer versão mais nova.                     #
#                                                                                # 
#   Este programa é distribuído na expectativa de ser útil, mas SEM              #
#   QUALQUER GARANTIA; sem mesmo a garantia implícita de                         #
#   COMERCIALIZAÇÃO ou de ADEQUAÇÃO A QUALQUER PROPÓSITO EM                      #
#   PARTICULAR. Consulte a Licença Pública Geral GNU para obter mais             #
#   detalhes.                                                                    #
#                                                                                #
#   This program is free software: you can redistribute it and/or                #
#   modify it under the terms of the GNU General Public License as published     #
#   by the Free Software Foundation, either version 2 of the License, or         #
#   version 3 of the License, or (at your option) any later version.             #
#                                                                                #
#   This program is distributed in the hope that it will be useful,              #
#   but WITHOUT ANY WARRANTY; without even the implied warranty of               #
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See                     #
#   the GNU General Public License for more details.                             #
#                                                                                #
#  Dependências / Dependencies:                                                  #
#    Python 2.6+  (www.python.org)                                               #
#    NumPy 1.4+ (www.numpy.org)                                                  #
#    SciPy 0.12+ (www.scipy.org)                                                 #
#================================================================================#

from __future__ import division
import math
import numpy as np
from numpy import sum, ones, log, exp
from scipy.misc import comb

def biodiv(s, n, ni):
    c = 0
    h = 0 
    pie = 0
    u = 0
    nmax = ni[0]
    for i in range(s):
        c = c + ((ni[i] * (ni[i] - 1)) / (n * (n - 1)))
        h = h + -(ni[i] / n) * math.log(ni[i] / n)  #  Shannon-Weaver
        pie = pie + (ni[i] / n) * (ni[i] / n)
        u = u + (ni[i] * ni[i])
        if ni[i] > nmax:
            nmax = ni[i]
    d1 = (s - 1) / math.log(n)                      #  Margalef
    d2 = s / math.sqrt(n)                           #  Menhinick
    c = 1 - c                                       #  Simpson
    j = h / math.log(s)                             #  Equitability
    pie = (n / (n - 1)) * (1 - pie)                 #  Hurlbert
    d = 1 - (nmax / n)                              #  Berger-Parker
    m = (n - math.sqrt(u)) / (n - math.sqrt(n))     #  MacIntosh
    return d1, d2, c, h, pie, d, m, j

def ca(X, scaling):
    # From scikit-bio
    # http://scikit-bio.org/
    
    r, c = X.shape
    grand_total = X.sum()
    Q = X / grand_total
    column_marginals = Q.sum(axis=0)
    row_marginals = Q.sum(axis=1)
    expected = np.outer(row_marginals, column_marginals)
    Q_bar = (Q - expected) / np.sqrt(expected)
    U_hat, W, Ut = np.linalg.svd(Q_bar, full_matrices=False)
    rank = svd_rank(Q_bar.shape, W)
    assert rank <= min(r, c) - 1
    U_hat = U_hat[:, :rank]
    W = W[:rank]
    U = Ut[:rank].T
    V = column_marginals[:, None]**-0.5 * U
    V_hat = row_marginals[:, None]**-0.5 * U_hat
    F = V_hat * W
    F_hat = V * W
    eigvals = W**2
    percvar = 100 * eigvals / np.sum(eigvals)
    cumvar = np.cumsum(eigvals)
    cumvar /= cumvar[-1]
    cumvar *= 100
    species_scores = [V, F_hat][scaling - 1]
    site_scores = [F, V_hat][scaling - 1]
    return eigvals, site_scores, species_scores, percvar, cumvar

def cca(X, Y, scaling):
    # From scikit-bio
    # http://scikit-bio.org/
    
    if X.shape[0] != Y.shape[0]:
        raise ValueError("Contingency and environmental tables must have"
                         " the same number of rows (sites). X has {0}"
                         " rows but Y has {1}.".format(X.shape[0], Y.shape[0]))
    row_max = Y.max(axis=1)
    if np.any(row_max <= 0):
        raise ValueError("Contingency table cannot contain row of only 0s")
    grand_total = Y.sum()
    Q = Y / grand_total
    column_marginals = Q.sum(axis=0)
    row_marginals = Q.sum(axis=1)
    expected = np.outer(row_marginals, column_marginals)
    Q_bar = (Q - expected) / np.sqrt(expected)
    X = scale(X, weights=row_marginals, ddof=0)
    X_weighted = row_marginals[:, None]**0.5 * X
    B, _, rank_lstsq, _ = np.linalg.lstsq(X_weighted, Q_bar)
    Y_hat = X_weighted.dot(B)
    Y_res = Q_bar - Y_hat
    u, s, vt = np.linalg.svd(Y_hat, full_matrices=False)
    rank = svd_rank(Y_hat.shape, s)
    s = s[:rank]
    u = u[:, :rank]
    vt = vt[:rank]
    U = vt.T
    U_hat = Q_bar.dot(U) * s**-1
    u_res, s_res, vt_res = np.linalg.svd(Y_res, full_matrices=False)
    rank = svd_rank(Y_res.shape, s_res)
    s_res = s_res[:rank]
    u_res = u_res[:, :rank]
    vt_res = vt_res[:rank]
    U_res = vt_res.T
    U_hat_res = Y_res.dot(U_res) * s_res**-1
    eigvals = np.r_[s, s_res]**2
    proportion_explained = 100 * eigvals / eigvals.sum()
    cumulative_proportion = np.cumsum(eigvals)
    cumulative_proportion /= cumulative_proportion[-1]
    cumulative_proportion *= 100
    V = (column_marginals**-0.5)[:, None] * U
    V_hat = (row_marginals**-0.5)[:, None] * U_hat
    F = V_hat * s
    F_hat = V * s
    Z_scaling1 = ((row_marginals**-0.5)[:, None] * Y_hat.dot(U))
    Z_scaling2 = Z_scaling1 * s**-1
    V_res = (column_marginals**-0.5)[:, None] * U_res
    V_hat_res = (row_marginals**-0.5)[:, None] * U_hat_res
    F_res = V_hat_res * s_res
    F_hat_res = V_res * s_res
    if scaling == 1:
        species_scores = np.hstack((V, V_res))
        site_scores = np.hstack((F, F_res))
        site_constraints = np.hstack((Z_scaling1, F_res))
    elif scaling == 2:
        species_scores = np.hstack((F_hat, F_hat_res))
        site_scores = np.hstack((V_hat, V_hat_res))
        site_constraints = np.hstack((Z_scaling2, V_hat_res))
    biplot_scores = corr(X_weighted, u)
    return eigvals, site_scores, species_scores, biplot_scores, proportion_explained, cumulative_proportion, site_constraints

def corr(x, y=None):
    # From scikit-bio
    # http://scikit-bio.org/
    
    if y is not None:
        if y.shape[0] != x.shape[0]:
            raise ValueError("Both matrices must have the same number of rows")
        x, y = scale(x), scale(y)
    else:
        x = scale(x)
        y = x
    return x.T.dot(y) / x.shape[0]

def hill_number(counts, q):
    # From Jonathan Friedman
    # http://yonatanfriedman.com/docs/survey/index.html
    
    counts = counts[counts>0]
    p = 1.* counts / sum(counts)
    lq = sum(p**q)
    if q == 1:
        plp = p * log(p)
        plp[p == 0] = 0
        lq = -sum(plp) 
        N = exp(lq)     
    else:      
        N  = lq**(1./ (1 - q))
    return N, lq

def mean_and_std(a, axis=None, weights=None, with_mean=True, with_std=True, ddof=0):
    # From scikit-bio
    # http://scikit-bio.org/
    
    if weights is None:
        avg = a.mean(axis=axis) if with_mean else None
        std = a.std(axis=axis, ddof=ddof) if with_std else None
    else:
        avg = np.average(a, axis=axis, weights=weights)
        if with_std:
            if axis is None:
                variance = np.average((a - avg)**2, weights=weights)
            else:
                a_rolled = np.rollaxis(a, axis)
                variance = np.average((a_rolled - avg)**2, axis=0,
                                    weights=weights)
            if ddof != 0:
                if axis is None:
                    variance *= a.size / (a.size - ddof)
                else:
                    variance *= a.shape[axis] / (a.shape[axis] - ddof)
            std = np.sqrt(variance)
        else:
            std = None
        avg = avg if with_mean else None
    return avg, std

def morisita_horn(datamtx):
    # From scikit-bio
    # http://scikit-bio.org/
    
    numrows, numcols = np.shape(datamtx)

    if numrows == 0 or numcols == 0:
        return np.zeros((0,0),'d')
    dists = np.zeros((numrows,numrows),'d')
    
    rowsums = datamtx.sum(axis=1, dtype="float")
    row_ds = (datamtx**2).sum(axis=1, dtype="float")
    
    for i in range(numrows):
        if row_ds[i] !=0.:
            row_ds[i] = row_ds[i] / rowsums[i]**2
    for i in range(numrows):
        row1 = datamtx[i]
        N1 = rowsums[i]
        d1 = row_ds[i]
        for j in range(i):
            row2 = datamtx[j]
            N2 = rowsums[j]
            d2 = row_ds[j]
            if N2 == 0.0 and N1 == 0.0:
                dist = 0.0
            elif N2 == 0.0 or N1 == 0.0:
                dist = 1.0
            else:
                similarity = 2 * sum(row1 * row2)
                similarity = similarity / ( (d1 + d2) * N1 * N2 )
                dist = 1 - similarity
            dists[i][j] = dists[j][i] = dist
    return dists

def ochiai(datamtx):
    # From scikit-bio
    # http://scikit-bio.org/
    
    datamtx = datamtx.astype(bool)
    datamtx = datamtx.astype(float)
    numrows, numcols = np.shape(datamtx)

    if numrows == 0 or numcols == 0:
        return np.zeros((0,0),'d')
    dists = np.zeros((numrows,numrows),'d')
    rowsums = datamtx.sum(axis=1)
    
    for i in range(numrows):
        first = datamtx[i]
        a = rowsums[i]
        for j in range(i):
            second = datamtx[j]
            b = rowsums[j]
            c = float(np.logical_and(first, second).sum())
            if a == 0.0 and b == 0.0:
                dist = 0.0
            elif a == 0.0 or b == 0.0:
                dist = 1.0
            else:
                dist = 1.0 - (c / math.sqrt(a * b))
            dists[i][j] = dists[j][i] = dist
    return dists

def pca(X, index=1, center=False, stand=False):
    if center:
        # Center data
        X = scale(X, with_std=False)
    if stand:
        # Standardize data
        X = scale(X, with_std=True)
    if index == 1:
        # Compute the covariance matrix
        cor_mat = np.cov(X.T)
    elif index == 2:
        # Compute the correlation matrix
        cor_mat = np.corrcoef(X.T)
    # Eigendecomposition of the covariance or correlation matrix
    eig_val, eig_vec = np.linalg.eig(cor_mat)
    # Drop imaginary component, if we got one
    eig_val = eig_val.real
    eig_vec = eig_vec.real
    # Compute percentage variances
    sumvariance = 100 * eig_val / eig_val.sum()
    cumvariance = np.cumsum(eig_val)
    cumvariance /= cumvariance[-1]
    cumvariance *= 100
    # Compute principal components
    scores = eig_vec.T.dot(X.T).T
    return eig_val, eig_vec, scores, sumvariance, cumvariance

def pcoa(distance_matrix):
    # From scikit-bio
    # http://scikit-bio.org/
    
    E_matrix = distance_matrix * distance_matrix / -2.0
    row_means = E_matrix.mean(axis=1, keepdims=True)
    col_means = E_matrix.mean(axis=0, keepdims=True)
    matrix_mean = E_matrix.mean()
    F_matrix = E_matrix - row_means - col_means + matrix_mean
    eigvals, eigvecs = np.linalg.eigh(F_matrix)
    negative_close_to_zero = np.isclose(eigvals, 0)
    eigvals[negative_close_to_zero] = 0
    idxs_descending = eigvals.argsort()[::-1]
    eigvals = eigvals[idxs_descending]
    eigvecs = eigvecs[:, idxs_descending]
    num_positive = (eigvals >= 0).sum()
    eigvecs[:, num_positive:] = np.zeros(eigvecs[:, num_positive:].shape)
    eigvals[num_positive:] = np.zeros(eigvals[num_positive:].shape)
    coordinates = eigvecs * np.sqrt(eigvals)
    coordinates *= -1
    proportion_explained = 100 * eigvals / eigvals.sum()
    cumulative_proportion = np.cumsum(eigvals)
    cumulative_proportion /= cumulative_proportion[-1]
    cumulative_proportion *= 100
    return eigvals, eigvecs, coordinates, proportion_explained, cumulative_proportion

def rarefact(freq):
    freq = freq[:]
    i = np.nonzero(freq==0)
    if not i:
        freq[i] = None
    S = len(freq)
    N = sum(freq)
    ES = np.zeros(N)
    for n in range(0,N):
        es = 0
        for s in range(0,S):
            es += (1 - comb(N - freq[s],n) / comb(N,n))
        ES[n] = es
    return ES, S, N

def rda(Y, X, scale_Y, scaling):
    # From scikit-bio
    # http://scikit-bio.org/
    
    n, p = Y.shape
    n_, m = X.shape
    if n != n_:
        raise ValueError("Both data matrices must have the same number of rows.")
    if n < m:
        raise ValueError("Explanatory variables cannot have less rows than columns.")
    Y = scale(Y, with_std=scale_Y)
    X = scale(X, with_std=False)
    B, _, rank_X, _ = np.linalg.lstsq(X, Y)
    Y_hat = X.dot(B)
    u, s, vt = np.linalg.svd(Y_hat, full_matrices=False)
    rank = svd_rank(Y_hat.shape, s)
    U = vt[:rank].T
    F = Y.dot(U)
    Z = Y_hat.dot(U)
    Y_res = Y - Y_hat
    u_res, s_res, vt_res = np.linalg.svd(Y_res, full_matrices=False)
    rank_res = svd_rank(Y_res.shape, s_res)
    U_res = vt_res[:rank_res].T
    F_res = Y_res.dot(U_res)
    eigvals = np.r_[s[:rank], s_res[:rank_res]]
    proportion_explained = eigvals * 100 / eigvals.sum()
    cumulative_proportion = np.cumsum(eigvals)
    cumulative_proportion /= cumulative_proportion[-1]
    cumulative_proportion *= 100
    const = np.sum(eigvals**2)**0.25
    if scaling == 1:
        scaling_factor = const
    elif scaling == 2:
        scaling_factor = eigvals / const
    species_scores = np.hstack((U, U_res)) * scaling_factor
    site_scores = np.hstack((F, F_res)) / scaling_factor
    site_constraints = np.hstack((Z, F_res)) / scaling_factor
    biplot_scores = corr(X, u)
    return eigvals, site_scores, species_scores, biplot_scores, proportion_explained, cumulative_proportion, site_constraints

def sample_diversity(counts_mat, indices="shannon", **kwargs):
    # From Jonathan Friedman
    # http://yonatanfriedman.com/docs/survey/index.html
    
    if isinstance(indices, str): indices = [indices]
    n_smpls = np.shape(counts_mat)[0]
    n_inds = len(indices)
    D = np.zeros((n_smpls, n_inds))
    for i in range(n_smpls):
        counts = counts_mat[i,:]
        for j in range(n_inds):
            ind = indices[j].lower()
            if ind in ["shannon", "hill"]:
                d, lq = hill_number(counts, 1)
                d = np.log(d)
                if ind == "hill": d = exp(d)
            elif ind in ["simpson", "simpson_inv"]:
                d, lq = hill_number(counts, 2)
                if ind == "simpson": d = 1 / d
            else: raise ValueError("%s is not a supported divesity index." %ind)
            D[i,j] = d
    return D, indices

def scale(a, weights=None, with_mean=True, with_std=True, ddof=0, copy=True):
    # From scikit-bio
    # http://scikit-bio.org/
    
    if copy:
        a = a.copy()
    avg, std = mean_and_std(a, axis=0, weights=weights, with_mean=with_mean,
                            with_std=with_std, ddof=ddof)
    if with_mean:
        a -= avg
    if with_std:
        std[std == 0] = 1.0
        a /= std
    return a

def svd_rank(M_shape, S, tol=None):
    # From scikit-bio
    # http://scikit-bio.org/
    
    if tol is None:
        tol = S.max() * max(M_shape) * np.finfo(S.dtype).eps
    return np.sum(S > tol)