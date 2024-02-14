## CpkTools - 工程能力算出ツール

このツールは、Excelファイルから数値データを取得し、QQプロットやヒストグラムを出力し、工程能力指数を算出します。

![image](https://github.com/kotaooka/CpkTools/assets/115392256/632ac7b4-22cb-4a21-8379-2de238a0eaa4)

### 使用方法
1. **Excelファイルの選択**:
   - 「Excelファイルのパス」欄にあるテキストボックスをクリックし、Excelファイルを選択します。
   - 「Excelファイルを選択」ボタンをクリックすることでも、ファイルを選択できます。

2. **規格値の入力**:
   - 「規格上限値」と「規格下限値」欄に規格値を入力します。

3. **計算実行**:
   - 「計算実行」ボタンをクリックすると、工程能力指数が計算されます。
   - 同時に、ヒストグラムやQQプロットが表示されます。

### sample.xlsxから算出されたヒストグラム
![histogram](https://github.com/kotaooka/-/assets/115392256/c781d2fd-7b60-4675-9297-a5b20be49ab2)

### sample.xlsxから算出されたQQプロット
![QQplot](https://github.com/kotaooka/-/assets/115392256/554f4180-bbb5-4eb2-b738-e11b91be5ace)

4. **計算結果の保存**:
   - 計算結果は、デスクトップに「計算日時_results.xlsx」という名前で保存されます。

### sample.xlsxから算出された計算結果
![image](https://github.com/kotaooka/-/assets/115392256/e2cc8439-ffac-4ca6-9939-d3bc96589295)


### 注意事項
- Excelファイルのデータは、最初の列から取り込まれます。
- 空のセルや数値以外のデータが含まれている場合、エラーメッセージが表示されます。

### 要件
- Python 3.x
- pandas
- scipy
- matplotlib
- tkinter
- openpyxl


