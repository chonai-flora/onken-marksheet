# onken-marksheet

## 概要
音楽研究部の活動記録をマークシート形式で記録し、活動報告書を自動で作成します。

### 0. マークシートを印刷
サンプルとして `excel/マークシート.xlsx` を用意しています。

### 1. セットアップ
```bash
# 仮想環境の作成
python3 -m venv venv

# 仮想環境をアクティべート
source venv/bin/lib/activate

# 関連モジュールをインストール
pip3 install -r requirements.txt
```

### 2. `scanned/` に画像データを配置
`excel/マークシート.xlsx` と `scanned/sample/` 内の画像を参考にしてください。

### 3. `excel/` に元となるExcelファイルを配置
`excel/元ファイル.xlsx` を `excel/<month>月元ファイル.xlsx` に変更し、名簿を入力してください。  
元ファイルとマークシートの名簿は同じ順にしてください。

### 4. 集計
```bash
python3 src/main.py
```
`excel/<month>月活動報告書.xlsx` が作成されます。

## 余談
以前は専用フォームを利用して出席を取っていましたが、手間が多かったため紙媒体にしました。フォームのソースは https://github.com/chonai-flora/onken-form にあります。