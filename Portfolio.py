from Optimization import SharpeOptimisation
from PortfolioAnalysis import PortfolioAnalysis

class Portfolio:

    def __init__(self, w_p: dict, file_name, optimization=True):
        self.w_p = w_p
        self.file_name = file_name
        self.optimization = optimization
        self.tickers = self.w_p.keys()
    
    def analysis(self):

        PortfolioAnalysis(self.w_p,save_as=self.file_name).complete_analysis()

        if self.optimization:
            
            opti = SharpeOptimisation(self.tickers).optimize_sharpe_ratio()
            new_weights = opti.get('Weights')
            new_w_p = dict(zip(self.tickers, new_weights))
            PortfolioAnalysis(new_w_p, save_as=self.file_name, sheet_name='Optimization').complete_analysis()

if __name__ == '__main__':
    w_p = {'KAYNES.NS': 0.2,
       'HDFCBANK.NS' : 0.1,
       'LTIM.NS': 0.2,
      'ULTRACEMCO.NS': 0.2,
      'ASIANPAINT.NS' : 0.3}
    
    Portfolio(w_p, file_name='Assignment Analysis.xlsx').analysis()
