import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import yfinance as yf
import scipy
from datetime import datetime

class SharpeOptimisation:
    
    def __init__(self, portfolio: list, start=None, end=None):
        
        self.portfolio = portfolio
        self.df = pd.DataFrame()
        
        self.start = start
        self.end = end
        
        if self.start == None or self.end == None:
            
            date = datetime.now()
            
            self.start = f'{date.year-1}-{date.month}-{date.day}'
            self.end = f'{date.year}-{date.month}-{date.day}'
        
        
        for ticker in self.portfolio:
            download = yf.download(ticker, self.start, self.end)['Close']
            
            self.df = pd.concat([self.df, download], axis=1)
        
        self.df.columns = self.portfolio
        
        
        self.log_ret = np.log(self.df / self.df.shift(1))

    def __get_ret_vol_sr(self, weights):
        weights = np.array(weights)
        ret = np.sum(self.log_ret.mean() * weights * 252)
        vol = np.sqrt(np.dot(weights.T, np.dot(self.log_ret.cov()*252, weights)))
        sr = ret / vol

        return np.array([ret, vol, sr])
        
    
    def __minimize_volatility(weights):
        return get_ret_vol_sr(weights)[1]
        
        
    def optimize_sharpe_ratio(self):
        
        def __neg_sharpe(weights):
            return self.__get_ret_vol_sr(weights)[2] * (-1)

        def __check_sum(weights):

            return np.sum(weights) - 1
        
        cons = ({'type' : 'eq', 'fun' : __check_sum})
        bounds = tuple([(0,1) for _ in range(len(self.portfolio))])
        init_guess = list(np.ones(len(self.portfolio)) / len(self.portfolio))

        opt_results = scipy.optimize.minimize(__neg_sharpe, init_guess, method='SLSQP', 
                       bounds=bounds, constraints=cons)
        
        weights = opt_results.x
        results = self.__get_ret_vol_sr(weights)
        
        return {'Best Sharpe ratio' : np.round(results[2], 4),
               'Estimated returns' : np.round(results[0], 4),
               'Estimated volatility': np.round(results[1], 4),
               'Weights': np.round(weights, 4)}


