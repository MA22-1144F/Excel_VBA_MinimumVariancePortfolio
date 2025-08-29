# Excel_VBA_MinimumVariancePortfolio
Minimum Variance Portfolio

## 概要
Excel上で株価データから最小分散フロンティアを作図します。  
ExcelのSolverアドインを使用します。

<img src="images/image_01.png" alt="フォームイメージ" width="300">
<img src="images/image_02.png" alt="グラフメージ" width="600">

## 動作環境
Microsoft Excel上で動作します。  

## インストール方法
1. Contentsフォルダ内の frmMVP.frm、frmMVP.frx を任意の同一フォルダに保存
2. 任意のExcelBookをマクロ有効ブック(.xlsm)として保存
3. [ファイル]>[オプション]>[アドイン]>[設定]でソルバー アドインにチェックを入れて[OK]またはEnter
4. [開発タブ]またはAlt+F11でVisual Basicを開く
5. [ツール]>[参照設定]またはAlt+T+RからSlverにチェックを入れて[OK]またはEnter
6. [▶Sub/ユーザー フォームの実行]またはF5でユーザーフォームを実行

## 機能
任意の株価データ(行方向に期間、列方向に銘柄)から以下のものを計算・表示する。
* 銘柄毎のログリターン及びその平均、標準偏差 ＊  
* 銘柄間の分散共分散行列 ＊  
* ポートフォリオの期待利益率を実現する最小の標準偏差及びその時の各銘柄への投資割合  
* 最小分散フロンティのグラフ  
* 縦軸：期待利益率、横軸：投資割合で各銘柄への投資割合の推移を表したグラフ
＊：参照形式で表示されます。



## 連絡先
[Instagram](https://www.instagram.com/nattotoasto?igsh=NWNtdHhnY3A4NDQ0 "nattotoasto")

## ライセンス
MIT License
