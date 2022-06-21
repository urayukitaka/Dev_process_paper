import os

def objective(trial): # trial is optuna function, for suggest params

    # parameter
    x = trial.suggest_uniform("x", 0, 5) # 変数xを上下限0~5の範囲で連続値
    y = trial.suggest_uniform("y", 0, 3) # 変数yを上下限0~3の範囲で連続値

    # return target values
    v0 = 4 * x ** 2 + 4 * y ** 2
    v1 = (x - 5) ** 2 + (y - 5) ** 2

    return v0, v1