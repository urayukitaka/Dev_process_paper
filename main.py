import os
import optuna
import matplotlib.pyplot as plt
from PIL import Image

import warnings
warnings.simplefilter('ignore')

import function as f
import utils as u

def exe_opt():

    # 最適化の条件設定, 変数を与える関数
    study = optuna.multi_objective.create_study(
        directions=["minimize", "minimize"], # "minimize" "maximize"
        sampler=optuna.multi_objective.samplers.NSGAIIMultiObjectiveSampler(seed = 1)
    )
    # 最適化の実行
    study.optimize(f.objective, n_trials=200)

    return study

def analyze(study):
    # 最適化過程で得た履歴データの取得。get_trials()メソッドを使用
    trials = {str(trial.values): trial for trial in study.get_trials()} # target values gotten y1 and y2
    trials = list(trials.values())

    # グラフにプロットするため、目的変数をリストに格納する
    y1_all_list = []
    y2_all_list = []
    for i, trial in enumerate(trials, start=1):
        y1_all_list.append(trial.values[0])
        y2_all_list.append(trial.values[1])

    # パレート解の取得。get_pareto_front_trials()メソッドを使用
    trials = {str(trial.values): trial for trial in study.get_pareto_front_trials()}
    trials = list(trials.values())
    trials.sort(key=lambda t: t.values)

    # グラフプロット用にリストで取得。またパレート解の目的変数と説明変数をcsvに保存する
    y1_list = []
    y2_list = []
    with open('pareto_data_real.csv', 'w') as f:
        for i, trial in enumerate(trials, start=1):
            if i == 1:
                columns_name_str = 'trial_no,y1,y2'
            data_list = []
            data_list.append(trial.number)
            y1_value = trial.values[0]
            y2_value = trial.values[1]
            y1_list.append(y1_value)
            y2_list.append(y2_value)
            data_list.append(y1_value)
            data_list.append(y2_value)
            for key, value in trial.params.items():
                data_list.append(value)
                if i == 1:
                    columns_name_str += ',' + key
            if i == 1:
                f.write(columns_name_str + '\n')
            data_list = list(map(str, data_list))
            data_list_str = ','.join(data_list)
            f.write(data_list_str + '\n')

    # パレート解を図示
    # setting of graph
    u.graph_setting()
    # plot
    plot_fig_list = list()
    for i in range(len(y1_all_list)):
        plt.figure(figsize=(10,6))
        plt.scatter(y1_all_list[:i], y2_all_list[:i], c='blue', label='all trials', s=20)
        plt.title("multiobjective optimization")
        plt.xlabel("Y1")
        plt.ylabel("Y2")
        plt.grid()
        plt.legend()
        plt.tight_layout()
        plot_fig_list.append("plots/pareto_graph_real_{}.png".format(i))
        plt.savefig("plots/pareto_graph_real_{}.png".format(i))
    # make gif image
    # create gif file
    images = list(map(lambda file : Image.open(file), plot_fig_list))
    images[0].save("optimization_step.gif", save_all=True, append_images=images[1:], duration=400, loop=0)

    # last result
    plt.figure(figsize=(10,6))
    plt.scatter(y1_all_list, y2_all_list, c='blue', label='all trials', s=20)
    plt.scatter(y1_list, y2_list, c='red', label='pareto front', s=20)
    plt.title("multiobjective optimization")
    plt.xlabel("Y1")
    plt.ylabel("Y2")
    plt.grid()
    plt.legend()
    plt.tight_layout()
    plt.savefig("pareto_graph_real.png")
    plt.close()

if __name__ == "__main__":

    # exe optimization
    print("-"*20)
    print("Exe optimization")
    study = exe_opt()
    print("-"*20)

    # analyze
    print("Exe analysis")
    analyze(study)
    print("-"*20)