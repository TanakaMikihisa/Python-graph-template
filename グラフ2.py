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
DATA_FILE = 'data2.xlsm'
X_COL = 'Hz'            # X軸データの列名
Y1_COL = 'G_s'          # Y軸データ1の列名
Y2_COL = 'G_b'          # Y軸データ2の列名

# グラフのラベル設定
X_LABEL = "入力波の周波数(Hz)"
Y_LABEL = "電圧利得"

# 各グラフの表示名設定
DATA_LABELS = {
    'data1': {
        'data': "積分器の電圧利得",           # 実測値
        'theory': "積分器の理論曲線",         # 理論曲線
        'fit': "積分器の近似曲線"            # 近似曲線
    },
    'data2': {
        'data': "微分器の電圧利得",           # 実測値
        'theory': "微分器の理論曲線",         # 理論曲線
        'fit': "微分器の近似曲線"            # 近似曲線
    }
}

# 理論式の定数
PARAM_A = 2 * np.pi      # 角周波数変換係数 (2π)
PARAM_B = 10 ** -6       # RC時定数 (秒)
                        # 積分器の遮断周波数 fc = 1/(2πRC)
                        # 現在値: fc ≈ 159.2 kHz
PARAM_C = 10 ** -3       # 微分器のゲイン係数
                        # 高周波での利得の上限を決定
                        # 現在値: 1/1000

def calc_theory(x_data):
    """周波数特性の理論式を計算する関数
    
    積分器: G(ω) = 1/sqrt(1 + (ωRC)^2)
    微分器: G(ω) = KωRC/sqrt(1 + (ωRC)^2)
    
    Parameters:
        x_data: 周波数データ (Hz)
    Returns:
        (y1, y2): データ1とデータ2の理論値
    """
    omega = PARAM_A * x_data  # 角周波数の計算 (ω = 2πf)
    denominator = np.sqrt(1 + PARAM_B * (omega ** 2))
    
    # データ1の理論値（積分器）
    y1 = 1 / denominator
    
    # データ2の理論値（微分器）
    y2 = PARAM_C * omega / denominator
    
    return y1, y2

# グラフの表示設定
PLOT_SIZE = (15, 5)      # グラフのサイズ (幅, 高さ)
FONT_SIZE = 17           # フォントサイズ
DPI = 300               # 保存時の解像度

# グラフスタイルの設定
PLOT_STYLES = {
    'data1': {
        'data': {                 # 実測値のスタイル
            'color': 'blue',
            'marker': 'o',
            'label': DATA_LABELS['data1']['data']
        },
        'theory': {              # 理論曲線のスタイル
            'color': 'blue',
            'linestyle': '-',
            'label': DATA_LABELS['data1']['theory']
        },
        'fit': {                 # 近似曲線のスタイル
            'color': 'red',
            'linestyle': '--',
            'label': DATA_LABELS['data1']['fit']
        }
    },
    'data2': {
        'data': {                 # 実測値のスタイル
            'color': 'orange',
            'marker': '^',
            'label': DATA_LABELS['data2']['data']
        },
        'theory': {              # 理論曲線のスタイル
            'color': 'orange',
            'linestyle': '-',
            'label': DATA_LABELS['data2']['theory']
        },
        'fit': {                 # 近似曲線のスタイル
            'color': 'red',
            'linestyle': ':',
            'label': DATA_LABELS['data2']['fit']
        }
    }
}

# 軸とグリッドの設定
AXIS_LOG = {
    'x': True,    # x軸を対数スケールに
    'y': True     # y軸を対数スケールに
}
SHOW_GRID = False  # グリッド線の表示

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
current_file = os.path.splitext(os.path.basename(__file__ if '__file__' in globals() else 'グラフ2'))[0]

# データ読み込み
df = pd.read_excel(os.path.join(current_dir, DATA_FILE), engine='openpyxl')
x_data = df[X_COL].copy()
y1_data = df[Y1_COL].copy()
y2_data = df[Y2_COL].copy()

# 無効なデータの除去
mask = (x_data > 0) & (y1_data > 0) & (y2_data > 0) & \
       (~np.isnan(x_data)) & (~np.isnan(y1_data)) & (~np.isnan(y2_data))
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
    # データ1と2の理論曲線を計算
    y1_theory, y2_theory = calc_theory(x_data)
    
    ax.plot(x_data, y1_theory, **PLOT_STYLES['data1']['theory'])
    ax.plot(x_data, y2_theory, **PLOT_STYLES['data2']['theory'])

# 近似直線の追加
if PLOT_FIT:
    try:
        # データ1の近似直線 (対数スケールで線形フィッティング)
        slope1, intercept1 = np.polyfit(np.log10(x_data), np.log10(y1_data), deg=1)
        fit1 = 10 ** (slope1 * np.log10(x_data) + intercept1)
        fit_label1 = f'{DATA_LABELS["data1"]["fit"]}: y ∝ x^{slope1:.2f}'
        ax.plot(x_data, fit1, **{**PLOT_STYLES['data1']['fit'], 'label': fit_label1})

        # データ2の近似直線 (対数スケールで線形フィッティング)
        slope2, intercept2 = np.polyfit(np.log10(x_data), np.log10(y2_data), deg=1)
        fit2 = 10 ** (slope2 * np.log10(x_data) + intercept2)
        fit_label2 = f'{DATA_LABELS["data2"]["fit"]}: y ∝ x^{slope2:.2f}'
        ax.plot(x_data, fit2, **{**PLOT_STYLES['data2']['fit'], 'label': fit_label2})
    except Exception as e:
        print(f"近似直線の計算でエラーが発生しました: {e}")
        print("近似直線の表示をスキップします")

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