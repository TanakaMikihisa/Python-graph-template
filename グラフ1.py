#!/usr/bin/python3

import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os
import subprocess
from matplotlib import font_manager

#=========================================
# カスタマイズ可能な設定項目
#=========================================

# データファイルの設定
DATA_FILE = 'data1.xlsm'
X_COL = 'Time'          # X軸データの列名
Y1_COL = 'Temp1'        # Y軸データ1の列名
Y2_COL = 'Temp2'        # Y軸データ2の列名

# グラフのラベル設定
X_LABEL = "経過時間 (分)"
Y_LABEL = "温度 (℃)"

# 各グラフの表示名設定
DATA_LABELS = {
    'data1': {
        'data': "温度センサー1の測定値",      # 実測値
        'theory': "温度1の理論曲線",         # 理論曲線
        'fit': "温度1の近似曲線"            # 近似曲線
    },
    'data2': {
        'data': "温度センサー2の測定値",      # 実測値
        'theory': "温度2の理論曲線",         # 理論曲線
        'fit': "温度2の近似曲線"            # 近似曲線
    }
}

# 理論式の定数
PARAM_A = 50.0          # 最大到達温度 (℃)
PARAM_B = 25.0          # 室温 (℃)
PARAM_C = 15.0          # 温度1の時定数 (分)
PARAM_D = 25.0          # 温度2の時定数 (分)

def calc_theory(x_data):
    """温度の理論式を計算する関数
    
    T(t) = Tmax + (T0 - Tmax)exp(-t/τ)
    
    Parameters:
        x_data: 経過時間 (分)
    Returns:
        (y1, y2): データ1とデータ2の理論値
    """
    # データ1の理論値（温度1）
    y1 = PARAM_A + (PARAM_B - PARAM_A) * np.exp(-x_data/PARAM_C)
    
    # データ2の理論値（温度2）
    y2 = PARAM_A + (PARAM_B - PARAM_A) * np.exp(-x_data/PARAM_D)
    
    return y1, y2

# グラフの表示設定
PLOT_SIZE = (15, 5)      # グラフのサイズ (幅, 高さ)
FONT_SIZE = 17           # フォントサイズ
DPI = 300               # 保存時の解像度

# グラフスタイルの設定
PLOT_STYLES = {
    'data1': {
        'data': {                 # 実測値のスタイル
            'color': 'red',
            'marker': 'o',
            'label': DATA_LABELS['data1']['data']
        },
        'theory': {              # 理論曲線のスタイル
            'color': 'red',
            'linestyle': '-',
            'label': DATA_LABELS['data1']['theory']
        },
        'fit': {                 # 近似曲線のスタイル
            'color': 'blue',
            'linestyle': '--',
            'label': DATA_LABELS['data1']['fit']
        }
    },
    'data2': {
        'data': {                 # 実測値のスタイル
            'color': 'green',
            'marker': '^',
            'label': DATA_LABELS['data2']['data']
        },
        'theory': {              # 理論曲線のスタイル
            'color': 'green',
            'linestyle': '-',
            'label': DATA_LABELS['data2']['theory']
        },
        'fit': {                 # 近似曲線のスタイル
            'color': 'blue',
            'linestyle': ':',
            'label': DATA_LABELS['data2']['fit']
        }
    }
}

# 軸とグリッドの設定
AXIS_LOG = {
    'x': False,   # x軸を線形スケールに
    'y': False    # y軸を線形スケールに
}
SHOW_GRID = True  # グリッド線の表示

# 追加機能の設定
PLOT_THEORY = True      # 理論曲線を表示するか
PLOT_FIT = True        # 近似曲線を表示するか


#=========================================
# 以下、メインの処理
#=========================================

# 日本語フォントの設定
try:
    font_manager.fontManager.addfont('/System/Library/Fonts/ヒラギノ角ゴシック W4.ttc')
    plt.rcParams['font.family'] = 'Hiragino Sans'
except:
    print("日本語フォントの設定に失敗しました。デフォルトフォントを使用します。")
    plt.rcParams['font.family'] = 'sans-serif'

plt.rcParams['axes.labelsize'] = FONT_SIZE
plt.rcParams['legend.fontsize'] = FONT_SIZE

# ファイルパスの設定
if '__file__' in globals():
    # 通常の実行時（ターミナルやIDEの実行ボタンから）
    current_dir = os.path.dirname(os.path.abspath(__file__))
else:
    # Jupyter等でセルから実行時
    current_dir = os.getcwd()
current_file = os.path.splitext(os.path.basename(__file__ if '__file__' in globals() else 'グラフ1'))[0]

# データ読み込み
df = pd.read_excel(os.path.join(current_dir, DATA_FILE), engine='openpyxl')
x_data = df[X_COL].copy()
y1_data = df[Y1_COL].copy()
y2_data = df[Y2_COL].copy()

# 無効なデータの除去
mask = (~np.isnan(x_data)) & (~np.isnan(y1_data)) & (~np.isnan(y2_data))
x_data = x_data[mask]
y1_data = y1_data[mask]
y2_data = y2_data[mask]

# グラフ作成
fig, ax = plt.subplots(figsize=PLOT_SIZE)

# 実測値のプロット
ax.scatter(x_data, y1_data, **PLOT_STYLES['data1']['data'])
ax.scatter(x_data, y2_data, **PLOT_STYLES['data2']['data'])

# 理論曲線の追加
if PLOT_THEORY:
    # 理論曲線用の時間点を生成（滑らかな曲線のため）
    t = np.linspace(0, max(x_data), 200)
    y1_theory, y2_theory = calc_theory(t)
    
    ax.plot(t, y1_theory, **PLOT_STYLES['data1']['theory'])
    ax.plot(t, y2_theory, **PLOT_STYLES['data2']['theory'])

# 近似直線の追加
if PLOT_FIT:
    try:
        # データ1の近似曲線（2次関数でフィッティング）
        coef1 = np.polyfit(x_data, y1_data, deg=2)
        fit1 = np.poly1d(coef1)
        t_fit = np.linspace(0, max(x_data), 100)
        ax.plot(t_fit, fit1(t_fit), **PLOT_STYLES['data1']['fit'])

        # データ2の近似曲線（2次関数でフィッティング）
        coef2 = np.polyfit(x_data, y2_data, deg=2)
        fit2 = np.poly1d(coef2)
        ax.plot(t_fit, fit2(t_fit), **PLOT_STYLES['data2']['fit'])
    except Exception as e:
        print(f"近似曲線の計算でエラーが発生しました: {e}")
        print("近似曲線の表示をスキップします")

# 軸の設定
ax.set_xscale('log' if AXIS_LOG['x'] else 'linear')
ax.set_yscale('log' if AXIS_LOG['y'] else 'linear')
ax.set_xlabel(X_LABEL)
ax.set_ylabel(Y_LABEL)

# y軸の範囲は元のまま
#ax.set_xlim(0, 10000000)
#plt.xticks(np.arange(0, 1.6 , 0.1 ))  # 0から50まで10刻み

# グラフの体裁設定
ax.grid(SHOW_GRID)
ax.tick_params(axis='both', which='both', top=True, right=True, direction='in')
ax.legend()
plt.tight_layout()

# 保存と表示
output_path = os.path.join(current_dir, f'{current_file}.png')
plt.savefig(output_path, dpi=DPI, bbox_inches='tight')
try:
    # macOSの場合
    subprocess.run(['open', output_path])
except:
    try:
        # Windowsの場合
        subprocess.run(['start', output_path], shell=True)
    except:
        try:
            # Linuxの場合
            subprocess.run(['xdg-open', output_path])
        except:
            print(f"グラフを保存しました: {output_path}")
            print("ファイルを手動で開いてください。")