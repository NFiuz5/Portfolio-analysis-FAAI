import numpy as np
import pandas as pd
import yfinance as yf

#Custom function in order to save multiple df on the same excel sheet, might not work in the future due to updates

def multiple_dfs(df_list, sheets, file_name, spaces, new_sheet=False):

    if not new_sheet:
        with pd.ExcelWriter(file_name,engine='openpyxl') as writer:
            col = 0
            for dataframe in df_list:
                dataframe.to_excel(writer,sheet_name=sheets,startrow=0 , startcol=col)   
                col = col + len(dataframe.columns) + spaces + 1
        
    
    else:
         with pd.ExcelWriter(file_name,engine='openpyxl', mode='a') as writer:

            col = 0
            for dataframe in df_list:
                dataframe.to_excel(writer,sheet_name=sheets,startrow=0 , startcol=col)   
                col = col + len(dataframe.columns) + spaces + 1

#Class allowing all the computing and modelling

class PortfolioAnalysis:
    
    def __init__(self, weighted_portfolio: dict, save_as=None, sheet_name='Analysis'):
        #Main argument to give : a weighted portfolio as shown in the example below
        
        self.file_name = save_as
        self.sheet_name = sheet_name
        self.df = pd.DataFrame()
        self.cov_matrix = pd.DataFrame()
        self.analysis_df = pd.DataFrame()
        self.tickers = list(weighted_portfolio.keys())
        self.weights = list(weighted_portfolio.values())
        self.ret_col = list()
        self.individual_ret = list()
        self.yearly_return = 0
        self.yearly_var = 0
        self.yearly_vol = 0
        
    
    #Download the 'Close' data of each ticker given with yfinance and creates one df to store it all
    def download_data(self):

        for t in self.tickers:
            data = yf.download(t, period='1y')
            self.df[t] = data['Close']
    
    #Compute the daily return of each stock and creates new cols in the df
    def compute_ret(self):
        
        for t in self.tickers:
            col_name = f'Ret_{t}'
            self.df[col_name] = self.df[t].pct_change(1)
            self.ret_col.append(col_name)
    
    #Compute the average yearly return of each stock
    def compute_individual_ret(self):
        self.individual_ret = [np.mean(self.df[ticker])*252 for ticker in self.ret_col]
    
    #Compute the yearly return of the portfolio
    def compute_yearly_return(self):
        
        avg = 0
        for i in range(len(self.weights)):
            avg += self.df[self.ret_col[i]].mean() * self.weights[i]
        
        self.yearly_return = avg * 252
            
    #Compute the covariance matrix, stored in a new df
    def compute_cov_matrix(self):
        
        self.cov_matrix = self.df[self.ret_col].cov()
    
    #Compute the yearly variance of the portfolio thanks to the covariance matrix
    def compute_yearly_var(self):
        
        cols = self.cov_matrix.columns
        results = list()
        
        for i in range(len(cols)):
    
            sum_prod = 0
            for y in range(len(cols)):
                sum_prod += self.cov_matrix[cols[i]].iloc[y] * self.weights[y]
            
            results.append(self.weights[i]*sum_prod)
        
        self.yearly_var = np.sum(results) * 252
    
    #Converts the yearly variance into a yearly std
    def compute_yearly_vol(self):
        
        self.yearly_vol = np.sqrt(self.yearly_var)
        
    #Process everything at the same time, once the Portfolio Analysis instance is created you only need to run this function
    def complete_analysis(self):
        
        self.download_data()
        self.compute_ret()
        self.compute_individual_ret()
        self.compute_yearly_return()
        
        self.compute_cov_matrix()
        self.compute_yearly_var()
        self.compute_yearly_vol()
        
        if self.file_name == None:
            print(f'Returns : {self.yearly_return}\nVolatily : {self.yearly_vol}\nReward/Risk :{self.yearly_return/self.yearly_vol} ')
        else:
            self.save_data()
        
    #Saving data in a single excel sheet 
    def save_data(self):
        
        analysis_dict = {'Yearly Return' : self.yearly_return,
                        'Yearly Variance' : self.yearly_var,
                        'Yearly Volatility' : self.yearly_vol,
                        'Risk Reward' : self.yearly_return/self.yearly_vol
                        }
        
        analysis_df = pd.DataFrame(analysis_dict, index=['Results']).transpose()
        df_yearly_ret = pd.DataFrame(list(zip(self.individual_ret, self.weights)), index=self.ret_col, columns=['Average yearly return', 'Weights'])

        if self.sheet_name == 'Optimization':

            new_weighted_portfolio = ""
            for i in range(len(self.tickers)):
                new_weighted_portfolio += f'{self.tickers[i]} : {self.weights[i]} / '
            analysis_df.loc['Weights'] = new_weighted_portfolio
            multiple_dfs([analysis_df], self.sheet_name, self.file_name, 1, True)

        else:
            multiple_dfs([self.df, self.cov_matrix, df_yearly_ret, analysis_df], self.sheet_name, self.file_name, 1)
        
        print('Download completed')

if __name__ == '__main__':
    w_p = {'LTIM.NS': 0.35,
       'HDFCBANK.NS' : 0.2,
       'INFY.NS': 0.15,
      'RELIANCE.NS': 0.3}

    portfolio = PortfolioAnalysis(w_p)
    print(portfolio.complete_analysis())