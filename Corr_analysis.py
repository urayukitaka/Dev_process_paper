import numpy as np
import matplotlib.pyplot as plt

# Corr analysis
from scipy.stats import pearsonr, spearmanr, kendalltau
# regression
import statsmodels.api as sm

def make_random_value():

    np.random.seed(332)

    # base plot
    num = 10
    x = np.linspace(0,50,num)
    std = (x+1)**2
    y = abs(np.random.normal(loc=0, scale=std))/5000
    #np.exp(x/10)/(np.random.random(num)/10)

    # add outlier
    x = list(x) + [0, 0, 0, 40, 80, 0]
    y = list(y) + [0.32, 0.3, 0.4, 0.3, 0.3, 0.5]

    # add 

    plt.figure(figsize=(10,6))
    plt.scatter(x,y)

    return x, y

def corr_analysis(x, y):

    # pearsonr
    cor_pearson, p_pearson = pearsonr(x, y)

    print("Pearson Corr coef / p : {} / {}".format(cor_pearson, p_pearson))

    # spearman
    cor_spearman, p_spearman = spearmanr(x, y)

    print("Spearman Corr coef / p : {} / {}".format(cor_spearman, p_spearman))

    # kendall
    cor_kendall, p_kendall = kendalltau(x, y)

    print("Kendall Corr coef / p : {} / {}".format(cor_kendall, p_kendall))


def quantile_regression(x, y, q):

    # Quantile regression
    # create instance and fit
    model_quat = sm.QuantReg(y, sm.add_constant(x)).fit(q=q)
    print("Psuedo R square : {}".format(model_quat.prsquared))

    # Linear regression
    model_simple = sm.OLS(y, sm.add_constant(x)).fit()
    print("R square : {}".format(model_simple.rsquared))

    # plot
    plt.figure(figsize=(10,6))
    plt.scatter(x, y, alpha=0.5)
    plt.plot(x, model_quat.predict(sm.add_constant(x)), color='red', label="quantile_regression q = {}".format(q))
    plt.plot(x, model_simple.predict(sm.add_constant(x)), color='blue', label="simple_regression")
    plt.legend()
    plt.show()



if __name__ == "__main__":

    x, y = make_random_value()

    print("")
    print("## Corr analysis ##")
    corr_analysis(x, y)
    print("")

    print("## Quantile regression")
    quantile_regression(x, y, q=0.2)

