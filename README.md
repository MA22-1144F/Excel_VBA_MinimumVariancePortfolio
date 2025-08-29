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
4. VBAProject一覧から VBAProject (PERSONAL.XLSB) を選択し右クリック
5. ファイルのインポートで保存した frmZenkakuHankaku.frm を開く
6. 加えてファイルのインポートで保存した 全角半角変換.bas を開く
7. 上書き保存ボタンまたはCtrl+Sで保存
以下オプション設定  
8. 任意のExcel WorkSheetに戻り、ファイル→オプション→リボンのユーザー設定の順で遷移
9. 任意のユーザー設定グループ(無ければ新規作成)に PERSONAL.XLSB!ZenkakuHankakuConverter のマクロを追加
10. お好みで名前やアイコンを変更
11. Excel WorkSheetに戻り、設定したマクロのアイコンを選択して起動

## 機能
選択範囲について以下の機能を実行

* 変換方向の選択  
  全角→半角 または 半角→全角  

* 変換対象の選択  
  * 英数字  
  * 記号    
  * カタカナ  
  * スペース  

* 数式が入力されたセルに対する処理の選択  

## 連絡先
[Instagram](https://www.instagram.com/nattotoasto?igsh=NWNtdHhnY3A4NDQ0 "nattotoasto")

## ライセンス
MIT License
