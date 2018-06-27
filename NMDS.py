#!usr/bin/env python
# -*- coding: utf-8 -*-
#================================================================================#
#  ECOLOG - Sistema Gerenciador de Banco de Dados para Levantamentos Ecológicos  #
#        ECOLOG - Database Management System for Ecological Surveys              #
#      Copyright (c) 1990-2015 Mauro J. Cavalcanti. All rights reserved.         #
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
from numpy import array, multiply, sum, zeros, size, shape, diag, dot, mean,\
    sqrt, transpose, trace, argsort, newaxis, finfo, all
from numpy.random import seed, normal as random_gauss
from numpy.linalg import norm, svd
from operator import itemgetter
import scipy.optimize as optimize
from Calc import pcoa

class nmds(object):
    # From Justin Kuczynski, The Cogent Project
    # http://pycogent.org/index.html

    def __init__(self, dissimilarity_mtx, initial_pts="pcoa", 
        dimension=2, rand_seed=None, max_iterations=50, 
        min_rel_improvement = 1e-3,
        min_abs_stress = 1e-5):
        
        self.min_rel_improvement = min_rel_improvement
        self.min_abs_stress = min_abs_stress

        if dimension >= len(dissimilarity_mtx) - 1:
            raise RuntimeError("NMDS requires N-1 dimensions or fewer, "+\
             "where N is the number of samples, or rows in the dissim matrix"+\
             " got %s rows for a %s dimension NMDS" % \
             (len(dissimilarity_mtx), dimension))

        if rand_seed != None:
            seed(rand_seed)
        
        num_points = len(dissimilarity_mtx)
        point_range = range(num_points)
        self.dimension = dimension
        
        self._calc_dissim_order(dissimilarity_mtx, point_range)
        
        if initial_pts == "random":
            self.points = self._get_initial_pts(dimension, point_range)
        elif initial_pts == "pcoa":
            pcoa_eigs, pcoa_evecs, pcoa_pts, percvars, cumvars = pcoa(dissimilarity_mtx)
            order = argsort(pcoa_eigs)[::-1]
            pcoa_pts = pcoa_pts[order].T
            self.points = pcoa_pts[:,:dimension]
        else:
            self.points = initial_pts
        self.points = self._center(self.points)
        self._rescale()
        self._calc_distances() 
        self._update_dhats()
        self._calc_stress()
        self.stresses = [self.stress]

        for i in range(max_iterations):
            if (self.stresses[-1] < self.min_abs_stress):
                break
            self._move_points()
            self._calc_distances()
            self._update_dhats()
            self._calc_stress()
            self.stresses.append(self.stress)
            if (self.stresses[-2]-self.stresses[-1]) / self.stresses[-2] <\
                self.min_rel_improvement:
                break
        self.points = self._center(self.points)
        u,s,vh = svd(self.points, full_matrices=False)
        S = diag(s)
        self.points = dot(u,S)
        self._rescale()
    
    @property
    def dhats(self):
        return [self._dhats[i,j] for (i,j) in self.order]
        
    @property
    def dists(self):
        return [self._dists[i,j] for (i,j) in self.order]

    def getPoints(self):
        return self.points
    
    def getStress(self):
        return self.stresses[-1]
        
    def getDimension(self):
        return self.dimension
        
    def _center(self, mtx):
        result = array(mtx, 'd')
        result -= mean(result, 0) 
        return result
    
    def _calc_dissim_order(self, dissim_mtx, point_range):
        dissim_list = []
        for i in point_range:
            for j in point_range:
                if j > i:
                    dissim_list.append([i, j, dissim_mtx[i,j]])
        dissim_list.sort(key = itemgetter(2))
        for elem in dissim_list:
            elem.pop()
        self.order = dissim_list

    def _get_initial_pts(self, dimension, pt_range):
        points = [[random_gauss(0., 1) for axis in range(dimension)] \
            for pt_idx in pt_range]
        return array(points, 'd')

    def _calc_distances(self):
        diffv = self.points[newaxis, :, :] - self.points[:, newaxis, :]
        squared_dists = (diffv**2).sum(axis=-1)
        self._dists = sqrt(squared_dists)
        self._squared_dist_sums = squared_dists.sum(axis=-1)
             
    def _update_dhats(self):
        new_dhats = self._dists.copy()
        ordered_dhats = [new_dhats[i,j] for (i,j) in self.order]
        ordered_dhats = self._do_monotone_regression(ordered_dhats)
        for ((i,j),d) in zip(self.order, ordered_dhats):
            new_dhats[i,j] = new_dhats[j, i] = d
        self._dhats = new_dhats
        
    def _do_monotone_regression(self, dhats):
        blocklist = []
        for top_dhat in dhats:
            top_total = top_dhat
            top_size = 1
            while blocklist and top_dhat <= blocklist[-1][0]:
                (dhat, total, size) = blocklist.pop()
                top_total += total
                top_size += size
                top_dhat = top_total / top_size
            blocklist.append((top_dhat, top_total, top_size))
        result_dhats = []
        for (val, total, size) in blocklist:
            result_dhats.extend([val]*size)
        return result_dhats
        
    def _calc_stress(self):
        diffs = (self._dists - self._dhats)
        diffs **= 2
        self._squared_diff_sums = diffs.sum(axis=-1)
        self._total_squared_diff = self._squared_diff_sums.sum() / 2
        self._total_squared_dist = self._squared_dist_sums.sum() / 2
        self.stress = sqrt(self._total_squared_diff/self._total_squared_dist)
        
    def _nudged_stress(self, v, d, epsilon):
        delta_epsilon = zeros([self.dimension], float)
        delta_epsilon[d] = epsilon
        moved_point = self.points[v] + delta_epsilon
        squared_dists = ((moved_point - self.points)**2).sum(axis=-1)
        squared_dists[v] = 0.0
        delta_squared_dist = squared_dists.sum() - self._squared_dist_sums[v]
        diffs = sqrt(squared_dists) - self._dhats[v]
        diffs **= 2
        delta_squared_diff = diffs.sum() - self._squared_diff_sums[v]
        return sqrt(
            (self._total_squared_diff + delta_squared_diff) /
            (self._total_squared_dist + delta_squared_dist))

    def _rescale(self):
        factor = array([norm(vec) for vec in self.points]).mean()
        self.points = self.points/factor

    def _move_points(self):
        numrows, numcols = shape(self.points)
        pts = self.points.ravel().copy()
        maxiter = 100
        while True:
            if maxiter <= 1:
                raise RuntimeError("could not run scipy optimizer")
            try:
                optpts = optimize.fmin_bfgs(
                 self._recalc_stress_from_pts, pts,
                 fprime=self._calc_stress_gradients,
                 disp=False, maxiter=maxiter, gtol=1e-3)
                break
            except FloatingPointError:
                maxiter = int(maxiter/2)
        self.points = optpts.reshape((numrows, numcols))

    def _recalc_stress_from_pts(self, pts):
        pts = pts.reshape(self.points.shape)
        changed = not all(pts == self.points)
        self.points = pts
        if changed:
            self._calc_distances()
        self._calc_stress()
        return self.stress
    
    def _calc_stress_gradients(self, pts):
        epsilon = sqrt(finfo(float).eps)
        f0 = self._recalc_stress_from_pts(pts)
        grad = zeros(pts.shape, float)
        for k in range(len(pts)):
            (point, dim) = divmod(k, self.dimension)
            f1 = self._nudged_stress(point, dim, epsilon)
            grad[k] = (f1 - f0)/epsilon
        return grad