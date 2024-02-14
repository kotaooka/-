import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
from tkinter import filedialog
import tkinter.messagebox as messagebox
import tkinter as tk
import os
import datetime


def run_program1():
    # ファイルが選択されているか確認
    file_path = txtBox3.get()
    if not file_path:
        messagebox.showerror('エラー - 工程能力算出ツール', 'ファイルを選択してください。')
        return
    
    # Excelファイルのデータを取得する
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror('エラー - 工程能力算出ツール', 'ファイルの読み込み中にエラーが発生しました。\n{}'.format(e))
        return
    
    # データフレームが空の場合のエラーハンドリング
    if df.empty:
        messagebox.showerror('エラー - 工程能力算出ツール', 'ファイルにデータがありません。')
        return

    # データフレームの最初の列を対象として計算を行う
    data_column = df.columns[0]

    # 空のセルがあるかどうかチェック
    if df[data_column].isnull().values.any():
        messagebox.showerror('エラー - 工程能力算出ツール', 'ファイルに空のセルが含まれています。')
        return

    # 基本統計量を計算
    resultstd = float(df[data_column].std())
    resultvar = float(df[data_column].var())
    resultmean = float(df[data_column].mean())
    resultmax = float(df[data_column].max())
    resultmin = float(df[data_column].min())
    resultkurt = float(df[data_column].kurt())
    resultskew = float(df[data_column].skew())

    # 規格値が入力されているか確認
    usl_value = txtBox1.get()
    lsl_value = txtBox2.get()
    if not usl_value or not lsl_value:
        messagebox.showerror('エラー - 工程能力算出ツール', '規格値を入力してください。')
        return

    # 数値以外が入力された場合のエラーハンドリング
    try:
        usl = float(usl_value)
        lsl = float(lsl_value)
    except ValueError:
        messagebox.showerror('エラー - 工程能力算出ツール', '規格値は数値で入力してください。')
        return

    # CPKを計算
    cpu = (usl - resultmean) / (3 * resultstd)
    cpl = (resultmean - lsl) / (3 * resultstd)
    cp = (usl - lsl) / (6 * resultstd)

    # ヒストグラムを作成
    plt.hist(df[data_column])
    plt.show()

    # Q-Q plotを作成
    stats.probplot(df[data_column], dist="norm", plot=plt)
    plt.show()

    # Shapiro-Wilk検定
    result1 = stats.shapiro(df[data_column])

    # Kolmogorov-Smirnov検定
    result2 = stats.ks_1samp(df[data_column], stats.norm.cdf)

    # 日付を取得してフォーマットを指定
    dt_now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')

    # 計算結果をxlsxファイルに出力する
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    filename = os.path.join(desktop_path, dt_now + '_results.xlsx')

    with pd.ExcelWriter(filename) as writer:
        df_result = pd.DataFrame({
            '計算項目': ['標準偏差', '分散', '尖度', '歪度', '最大値', '最小値', '平均', 'Cp', '上限規格側 Cpk', '下限規格側 Cpk', 'Shapiro-Wilk検定', 'Kolmogorov-Smirnov検定', 'ファイルパス'],
            '計算結果': [resultstd, resultvar, resultkurt, resultskew, resultmax, resultmin, resultmean, cp, cpu, cpl, result1, result2, file_path]
        })
        df_result.to_excel(writer, index=False)

    # 完了メッセージを表示
    messagebox.showinfo('計算完了 - 工程能力算出ツール', '計算が完了しました。計算結果はデスクトップに保存されます。')


# GUI
baseGround = tk.Tk()
baseGround.title('工程能力算出ツール')
baseGround.geometry('512x288')

label = tk.Label(text='Excelファイルの数値からQQプロット、ヒストグラムを出力し、工程能力指数を算出します。')
label.place(x=50, y=10)

label3 = tk.Label(text='Excelファイルのパス')
label3.place(x=50, y=50)
txtBox3 = tk.Entry(width=50)
txtBox3.place(x=50, y=70)

label = tk.Label(text='注意:Excelファイルのデータは最初の列から取り込まれます。')
label.place(x=50, y=90)

label1 = tk.Label(text='規格上限値')
label1.place(x=50, y=130)
txtBox1 = tk.Entry(width=20)
txtBox1.place(x=50, y=150)

label2 = tk.Label(text='規格下限値')
label2.place(x=50, y=170)
txtBox2 = tk.Entry(width=20)
txtBox2.place(x=50, y=190)

buttonB = tk.Button(baseGround, text='Excelファイルを選択', command=lambda: txtBox3.insert(0, filedialog.askopenfilename(filetypes=[('Excelファイル', '*.xlsx *.xls')], initialdir='C:\\')))
buttonB.place(x=370, y=66)

buttonA = tk.Button(baseGround, text='計算実行', command=run_program1)
buttonA.place(x=150, y=240)

buttonC = tk.Button(baseGround, text='閉じる', command=baseGround.destroy)
buttonC.place(x=300, y=240)

baseGround.mainloop()