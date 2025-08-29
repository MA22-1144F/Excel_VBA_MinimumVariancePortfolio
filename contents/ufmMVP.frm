VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmMVP 
   Caption         =   "完全版_最小分散フロンティア"
   ClientHeight    =   4710
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   6525
   OleObjectBlob   =   "ufmMVP.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ufmMVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'このコードはユーザーフォーム用に記述しています。 _
    利用にはユーザーフォームが必要です。
'VBAを利用する場合は、マクロ有効ブックとして保存する必要があるので注意。
'ソルバーの機能を利用しているので、Excelのアドインでソルバーアドインにチェックを入れた上で、ツール>参照設定からSolverにチェックを入れる必要があります。
'実行すると処理に数分間かかる場合があるので注意。
'株価Sheet名：分析したい株価情報が記録されているSheet名を入力してください。
'新規Sheet名：分析結果を表示するSheet名を入力してください(Sheetが自動で作成されます)。
'株価データ行・列：分析したい株価情報の範囲を指定してください(範囲が広いほど処理に時間がかかります)。 _
    株価の部分のみ指定し、会社名や証券コード、期間などは含めないでください。 _
    株価情報は行方向に期間、列方向に銘柄を羅列してください。 _
    行の指定は半角数字、列の指定は大文字または小文字の半角英字で行ってください。
'銘柄名∨コード列：分析したい銘柄の名称またはコードの範囲を指定してください。 _
    銘柄の名称またはコードは株価データと同一の行に記録してください。 _
    列の指定は大文字または小文字の半角英字で行ってください。
'最低投資割合：各銘柄に投資する最低限の割合を0以上の半角数字で入力してください。 _
    0を入力した場合は投資しない銘柄が生じる可能性があります｡
'期待利益率の段階：分析する期待利益率の細かさの程度を指定してください。 _
    10程度でも十分な最小分散フロンティアの作図が可能です(数字が大きいほど処理に時間がかかります)。
'オプションボタン：どこまでの処理を実行するかを3段階で選択できます。 _
    {ログリターンまで}、{ポートフォリオ標準偏差まで}は比較的短時間で処理が終了します。 _
    {最終分散フロンティア作図まで}は処理に時間を要します。
'クリアボタン：入力欄をすべて空欄に戻します。
'実行ボタン：処理を実行します(入力欄を埋めるまでは押せません)。


'列(英字入力)を(数字入力)へ変換
Function ColumnName2Idx(ByVal colName As String) As Integer
    ColumnName2Idx = Columns(colName).Column
End Function

'列(数字入力)を(英字入力)へ変換
Function ColumnIdx2Name(ByVal colNum As Integer) As String
    ColumnIdx2Name = Split(Columns(colNum).Address, "$")(2)
End Function

'文字列の全ての文字が英字の場合はTrue、そうでない場合はFalseを返す
Function IsAlpha(str As String) As Boolean
    IsAlpha = Not str Like "*[!a-zA-Zａ-ｚＡ-Ｚ]*"
End Function

'テキストボックスの初期設定
'初期値を灰色で入力
Private Sub UserForm_Initialize()
    cmdEX.Enabled = False
    txtSheet1.Value = "株価データサンプル"
    txtSheet1.ForeColor = &HC0C0C0
    txtSheet2.Text = "最小分散フロンティア"
    txtSheet2.ForeColor = &HC0C0C0
    txtStock1.Text = 4
    txtStock1.ForeColor = &HC0C0C0
    txtStock2.Text = 23
    txtStock2.ForeColor = &HC0C0C0
    txtStock3.Text = "c"
    txtStock3.ForeColor = &HC0C0C0
    txtStock4.Text = "bk"
    txtStock4.ForeColor = &HC0C0C0
    txtName.Text = "a"
    txtName.ForeColor = &HC0C0C0
    txtMinweight.Text = 1
    txtMinweight.ForeColor = &HC0C0C0
    txtStep.Text = 50
    txtStep.ForeColor = &HC0C0C0
End Sub

'すべてのテキストボックスに入力が無い間は実行ボタンを無効化
Private Sub CheckTextBoxes()
    Dim ctrl As Control     'テキストボックス
    Dim allFilled As Boolean        'すべてのテキストボックスに入力されているか
    allFilled = True
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            If ctrl.Text = "" Then
                allFilled = False
                Exit For
            End If
        End If
    Next ctrl
    cmdEX.Enabled = allFilled
End Sub

'テキストボックスの初期設定
'マウスで入力を始めるとテキストボックスの初期値を消して文字色を黒色にする
Private Sub txtSheet1_Change()
    CheckTextBoxes
End Sub
Private Sub txtSheet1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtSheet1.Text = ""
    txtSheet1.ForeColor = &H80000008
End Sub
Private Sub txtSheet2_Change()
    CheckTextBoxes
End Sub
Private Sub txtSheet2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtSheet2.Text = ""
    txtSheet2.ForeColor = &H80000008
End Sub
Private Sub txtStock1_Change()
    CheckTextBoxes
End Sub
Private Sub txtStock1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtStock1.Text = ""
    txtStock1.ForeColor = &H80000008
End Sub
Private Sub txtStock2_Change()
    CheckTextBoxes
End Sub
Private Sub txtStock2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtStock2.Text = ""
    txtStock2.ForeColor = &H80000008
End Sub
Private Sub txtStock3_Change()
    CheckTextBoxes
End Sub
Private Sub txtStock3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtStock3.Text = ""
    txtStock3.ForeColor = &H80000008
End Sub
Private Sub txtStock4_Change()
    CheckTextBoxes
End Sub
Private Sub txtStock4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtStock4.Text = ""
    txtStock4.ForeColor = &H80000008
End Sub
Private Sub txtName_Change()
    CheckTextBoxes
End Sub
Private Sub txtName_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtName.Text = ""
    txtName.ForeColor = &H80000008
End Sub
Private Sub txtMinweight_Change()
    CheckTextBoxes
End Sub
Private Sub txtMinweight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtMinweight.Text = ""
    txtMinweight.ForeColor = &H80000008
End Sub
Private Sub txtStep_Change()
    CheckTextBoxes
End Sub
Private Sub txtStep_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    txtStep.Text = ""
    txtStep.ForeColor = &H80000008
End Sub

'___________________________________________________________________________________________________
'___________________________________________________________________________________________________
'実行ボタンを推した際の動作
Private Sub cmdEX_Click()
    Application.ScreenUpdating = False  '画面更新を停止
    Application.EnableEvents = False    'イベントを抑制
    Application.Calculation = xlCalculationManual   '計算を手動
    
    '各テキストボックスの入力に関するエラーを表示
    If Not IsNumeric(txtStock1.Text) Then
        MsgBox "株価データ 行" & txtStock1.Text & "は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    If Not IsNumeric(txtStock2.Text) Then
        MsgBox "株価データ 行" & txtStock2.Text & "は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    If Not IsAlpha(txtStock3.Text) Then
        MsgBox "株価データ 列" & txtStock3.Text & "は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    If Not IsAlpha(txtStock4.Text) Then
        MsgBox "株価データ 列" & txtStock4.Text & "は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    If Not IsAlpha(txtName.Text) Then
        MsgBox "銘柄名∨コード 列" & txtName.Text & "は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    If Not IsNumeric(txtMinweight.Text) Then
        MsgBox "最低投資割合" & txtMinweight.Text & "%は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    If Not IsNumeric(txtStep.Text) Then
        MsgBox "期待利益の段階" & txtStep.Text & "段階は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    
    '最低投資割合が全銘柄均等投資の場合を超過する場合にはエラーを表示
    Dim Stock1 As Integer, Stock2 As Integer, Stock3 As Integer, Stock4 As Integer
    Stock1 = txtStock1.Text     '株価データの開始行
    Stock2 = txtStock2.Text     '株価データの最終行
    Dim sr As Integer       '株価データ終了行-開始行
    sr = Stock2 - Stock1
    Dim MinWeight As Double     '最低投資割合の数値
    MinWeight = txtMinweight.Text / 100
    If 1 / (sr + 1) < MinWeight Then
        MsgBox "最低投資割合" & txtMinweight.Text & "%は不正です。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    
    '株価Sheet名がExcelBookに存在しない場合にエラーを表示
    Dim flg2 As Boolean     '株価Sheet名が存在するか
    Dim chkWs As Worksheet      'ExcelBook内のSheet
    flg2 = False
    For Each chkWs In Worksheets
        If chkWs.Name = txtSheet1.Text Then
            flg2 = True
            Exit For
        End If
    Next chkWs
    If flg2 = False Then
        MsgBox "Sheet名「" & txtSheet1.Text & "」は存在しません。", vbCritical + vbOKOnly, "エラー"
        Exit Sub
    End If
    
    '新規Sheet名が既にExcelBookに存在する場合にエラーを表示
    Dim flg As Boolean      '新規Sheet名が既存か
    Dim addWs As Worksheet      '新規Sheet
    flg = True
    For Each chkWs In Worksheets
        If chkWs.Name = txtSheet2.Text Then
            flg = False
            MsgBox "Sheet名「" & txtSheet2.Text & "」は既存です。", vbCritical + vbOKOnly, "エラー"
            Exit Sub
        End If
    Next chkWs
    
    '新規Sheet名のSheetを最後尾に追加
    If flg Then
        Set addWs = Worksheets.Add(After:=Sheets(Worksheets.Count))
        addWs.Name = txtSheet2.Text
    End If
    
    '●株価情報からログリターンの計算
    Stock3 = ColumnName2Idx(txtStock3.Text)     '株価データの開始列(数字変換)
    Stock4 = ColumnName2Idx(txtStock4.Text)     '株価データの最終列(数字変換)
    Dim stkWs As Worksheet: Set stkWs = Worksheets(txtSheet1.Text)     '株価Sheet
    Dim sc As Integer     '株価データの期間列-1
    Dim lnr As Integer      '新規ワークシートの基準位置(行)
    Dim lnc As Integer      '新規ワークシートの基準位置(列)
    Dim Stock3a As String       '株価データの開始列+1(英字変換)
    lnr = 2
    lnc = 2
    sc = Stock4 - Stock3 - 1
    Stock3a = ColumnIdx2Name(Stock3 + 1)
    'ログリターンのラベルを表示
    addWs.Cells(lnr - 1, lnc) = "ログリターン"
    'ログリターンを表示
    addWs.Range(Cells(lnr, lnc), Cells(lnr + sr, lnc + sc)).Formula _
        = "=IFERROR(LN('" & stkWs.Name & "'!" & Stock3a & txtStock1.Text & "/'" & stkWs.Name & "'!" & txtStock3.Text & txtStock1.Text & "),0)"
'    'ログリターンの計算にエラーが生じた場合に0を表示
'    Dim lrTarget As Range
'    On Error Resume Next
'    Set lrTarget = addWs.Range(Cells(lnr, lnc), Cells(lnr + sr, lnc + sc)).SpecialCells(xlCellTypeFormulas, xlErrors)
'    On Error GoTo 0
'    If Not lrTarget Is Nothing Then
'        lrTarget.Value = 0
'    End If
    
    '銘柄名∨コードをログリターンの左横に表示
    addWs.Range(Cells(lnr, lnc - 1), Cells(lnr + sr, lnc - 1)).Formula _
        = "='" & stkWs.Name & "'!" & txtName.Text & txtStock1.Text
    
    '●ログリターンの平均と標準偏差
    Dim lnca As String      '新規ワークシートのログリターン開始列(英字変換)
    Dim lncb As String      '新規ワークシートのログリターン最終列(英字変換)
    lnca = ColumnIdx2Name(lnc)
    lncb = ColumnIdx2Name(lnc + sc)
    'ログリターン平均のラベルを表示
    addWs.Cells(lnr - 1, lnc + sc + 2) = "ログリターン平均"
    'ログリターン平均の関数を記述
    addWs.Range(Cells(lnr, lnc + sc + 2), Cells(lnr + sr, lnc + sc + 2)).Formula _
        = "=AVERAGE(" & lnca & lnr & ":" & lncb & lnr & ")"
    'ログリターン標準偏差のラベルを表示
    addWs.Cells(lnr - 1, lnc + sc + 3) = "ログリターン標準偏差"
    'ログリターン標準偏差の関数を記述
    addWs.Range(Cells(lnr, lnc + sc + 3), Cells(lnr + sr, lnc + sc + 3)).Formula _
        = "=STDEV.P(" & lnca & lnr & ":" & lncb & lnr & ")"
    
    '▲オプションボタンで「ログリターンまで」を選択した場合はここまでで処理を終了
    If opb1 Then
        Application.EnableEvents = True    'イベントを開始
        Application.ScreenUpdating = True  '画面更新を開始
        Application.Calculation = xlCalculationAutomatic   '計算を自動
        Unload ufm完全版_最小分散フロンティア       'ユーザーフォームを閉じる
        MsgBox "処理が完了しました。", vbInformation      '終了メッセージ
        Exit Sub
    End If
    
    '●ウエイトの初期設定
    Dim lncc As String      '新規ワークシートのウエイト最終列(英字変換)
    lncc = ColumnIdx2Name(lnc + sr)
    '一時的にすべての銘柄のウエイトを最低投資割合に設定
    addWs.Range(Cells(lnr + sr + 3, lnc), Cells(lnr + sr + 3, lnc + sr)) = MinWeight
    'ウエイト合計のラベルを表示
    addWs.Cells(lnr + sr + 2, lnc + sr + 2) = "ウエイト合計"
    'ウエイト合計の関数を記述
    addWs.Cells(lnr + sr + 3, lnc + sr + 2) _
        = "=SUM(" & lnca & lnr + sr + 3 & ":" & lncc & lnr + sr + 3 & ")"
    '銘柄名∨コードをウエイトの上に横並びに表示
    Dim n As Integer        '銘柄名∨コードの番目
    Dim lncf As String      '最初に表示した銘柄名∨コードの表示列(英字変換)
    lncf = ColumnIdx2Name(lnc - 1)
    For n = 0 To sr
        addWs.Cells(lnr + sr + 2, lnc + n) _
            = "=" & lncf & lnr + n
    Next n
    
    '●ポートフォリオの期待利益率
    Dim lncd As String      'ログリターン平均の関数を記述されている列(英字変換)
    lncd = ColumnIdx2Name(lnc + sc + 2)
    'ポートフォリオ期待利益率のラベルを表示
    addWs.Cells(lnr + sr + 5, lnc) = "ポートフォリオ期待利益率"
    'ポートフォリオ期待利益率を求める関数の記述
    addWs.Cells(lnr + sr + 6, lnc) _
        = "=MMULT(" & lnca & lnr + sr + 3 & ":" & lncc & lnr + sr + 3 & "," & lncd & lnr & ":" & lncd & lnr + sr & ")"
    
    '●分散共分散行列
    '分散共分散行列のラベルを表示
    addWs.Cells(lnr + sr + 8, lnc - 1) = "分散共分散行列"
    Dim var As Integer      'ログリターンの行の番目
    For var = 0 To sr
    '分散共分散行列の関数を記述
    addWs.Range(Cells(lnr + sr + 9, lnc + var), Cells(lnr + 2 * sr + 9, lnc + var)).Formula _
        = "=COVARIANCE.P(" & "$" & lnca & "$" & lnr + var & ":" & "$" & lncb & "$" & lnr + var & "," & "$" & lnca & lnr & ":" & "$" & lncb & lnr & ")"
    Next var
    Dim mm As Integer       '分散共分散行列の行の番目
    For mm = 0 To sr
        addWs.Cells(lnr + 2 * sr + 11, lnc + mm).Formula _
            = "=INDEX(MMULT(" & lnca & lnr + sr + 3 & ":" & lncc & lnr + sr + 3 & "," & lnca & lnr + sr + 9 & ":" & lncc & lnr + 2 * sr + 9 & ")," & mm + 1 & ")"
    Next mm
    '銘柄名∨コードを分散共分散行列の上に横並びに表示
    addWs.Range(Cells(lnr + sr + 8, lnc), Cells(lnr + sr + 8, lnc + sr)).Formula _
        = "=" & lnca & lnr + sr + 2
    '銘柄名∨コードを分散共分散行列の左横に縦並びに表示
    addWs.Range(Cells(lnr + sr + 9, lnc - 1), Cells(lnr + 2 * sr + 9, lnc - 1)).Formula _
        = "=" & lncf & lnr
    
    '●ウエイトの複製
    Dim weights As Integer      'ウエイトの行の番目
    Dim lnce As String      '各銘柄のウエイトの列(英字変換)
    'ウエイトの初期設定で設定したウエイトを縦並びに置き換える
    For weights = 0 To sr
        lnce = ColumnIdx2Name(lnc + weights)
        addWs.Cells(lnr + 2 * sr + 13 + weights, lnc).Formula _
            = "=" & lnce & lnr + sr + 3
    Next weights
    '銘柄名∨コードをウエイトの左横に縦並びに表示
    addWs.Range(Cells(lnr + 2 * sr + 13, lnc - 1), Cells(lnr + 3 * sr + 13, lnc - 1)).Formula _
        = "=" & lncf & lnr
    
    '●ポートフォリオの標準偏差
    'ポートフォリオ標準偏差のラベルを表示
    addWs.Cells(lnr + 3 * sr + 15, lnc) = "ポートフォリオ標準偏差"
    'ポートフォリオ標準偏差を求める関数の記述
    addWs.Cells(lnr + 3 * sr + 16, lnc).Formula _
        = "=SQRT(MMULT(" & lnca & lnr + 2 * sr + 11 & ":" & lncc & lnr + 2 * sr + 11 & "," & lnca & lnr + 2 * sr + 13 & ":" & lnca & lnr + 3 * sr + 13 & "))"
    
    '▲オプションボタンで「ポートフォリオ標準偏差まで」を選択した場合はここまでで処理を終了
    If opb2 Then
        Application.EnableEvents = True    'イベントを開始
        Application.ScreenUpdating = True  '画面更新を開始
        Application.Calculation = xlCalculationAutomatic  '計算を自動
        Unload ufm完全版_最小分散フロンティア       'ユーザーフォームを閉じる
        MsgBox "処理が完了しました。", vbInformation      '終了メッセージ
        Exit Sub
    End If
    
    '●実現期待利益率の範囲と段階の設定
    'ポートフォリオ標準偏差のラベルを表示
    addWs.Cells(lnr + 3 * sr + 18, lnc) = "ポートフォリオ標準偏差"
    'ポートフォリオ期待利益率のラベルを表示
    addWs.Cells(lnr + 3 * sr + 18, lnc + 1) = "ポートフォリオ期待利益率"
    Dim MaxWeight As Double     '設定した最低投資割合の上で実現可能な最大の投資割合
    Dim MaxReturn As Double     '実現可能な最大のポートフォリオ期待利益率
    Dim MinReturn As Double     '実現可能な最低のポートフォリオ期待利益率
    Dim MaxRe As Double     'すべての銘柄の中で最も高いログリターン平均
    Dim MinRe As Double     'すべての銘柄の中で最も低いログリターン平均
    Dim DifReturn As Double     '実現可能なポートフォリオ期待利益率の範囲を期待利益率の段階で分けたときの1段階の値
    Dim r As Double     'MinReturnにDifReturnを加えた実際に表示する値
    Dim counter As Integer     'MinReturnにDifReturnを加えた回数
    MaxWeight = 1 - (MinWeight * (sr + 1))
    With Application.WorksheetFunction
        MaxRe = .Max(addWs.Range(Cells(lnr, lnc + sc + 2), Cells(lnr + sr, lnc + sc + 2)))
        MaxReturn _
            = MaxRe * MaxWeight _
            + (.Sum(addWs.Range(Cells(lnr, lnc + sc + 2), Cells(lnr + sr, lnc + sc + 2))) - MaxRe) * MinWeight
        MinRe = .Min(addWs.Range(Cells(lnr, lnc + sc + 2), Cells(lnr + sr, lnc + sc + 2)))
        MinReturn _
            = MinRe * MaxWeight _
            + (.Sum(addWs.Range(Cells(lnr, lnc + sc + 2), Cells(lnr + sr, lnc + sc + 2))) - MinRe) * MinWeight
    End With
    DifReturn = (MaxReturn - MinReturn) / (txtStep.Text)
    r = MinReturn
    counter = 0
    Do While r <= MaxReturn
        addWs.Cells(lnr + 3 * sr + 19 + counter, lnc + 1) = r
        counter = counter + 1
        r = r + DifReturn
    Loop
    '銘柄名∨コードを縦並びに表示
    addWs.Range(Cells(lnr + 3 * sr + 18, lnc + 3), Cells(lnr + 3 * sr + 18, lnc + sr + 3)).Formula _
        = "=" & lnca & lnr + sr + 2

    '●ソルバーの実行
    'ソルバーをリセット
    SolverReset
    'ソルバーの精度を指定
    SolverOptions Precision:=0.000001       '初期値：0.000001
    Set SetObjectiveCells = addWs.Cells(lnr + 3 * sr + 16, lnc)     '目的セルの範囲をポートフォリオ標準偏差に指定
    Set ChangingVariableCells = addWs.Range(Cells(lnr + sr + 3, lnc), Cells(lnr + sr + 3, lnc + sr))    '変数セルの範囲をウエイトに指定
    Set PortfolioReturn = addWs.Cells(lnr + sr + 6, lnc)        'ポートフォリオの期待利益率の範囲を指定
    'ソルバーの目的セル、目標値、変数セル、解決方法を設定
    SolverOk SetCell:=SetObjectiveCells, _
        MaxMinVal:=2, _
        ByChange:=ChangingVariableCells, _
        EngineDesc:="GRG Nonlinear"
    Dim q As Integer        'ウエイトの行の番目
    'ウエイトの制約条件として>=MinWeightを設定
    For q = lnc To lnc + sr
        SolverAdd CellRef:=addWs.Cells(lnr + sr + 3, q), _
            Relation:=3, _
            FormulaText:=CDbl(MinWeight)
    Next q
    'ウエイト合計の制約条件として=1を設定
    SolverAdd CellRef:=addWs.Cells(lnr + sr + 3, lnc + sr + 2), _
        Relation:=2, _
        FormulaText:=1
    'ポートフォリオ期待利益率の制約条件を設定した実現期待利益率に従って変更しながらソルバーを稼働して結果を表示
    Dim i As Integer        '実現期待利益率の番目
    For i = lnr + 3 * sr + 19 To lnr + 3 * sr + 18 + counter
        'ポートフォリオ期待利益率について設定されている制約条件を削除
        SolverDelete CellRef:=PortfolioReturn, _
            Relation:=2
        Set RealizedReturn = addWs.Cells(i, lnc + 1)        '設定した実現期待利益率
        Set RiskOutcome = addWs.Cells(i, lnc)       'ソルバーを実行した結果表示されるポートフォリオ標準偏差
        Set WeightsOutcome = addWs.Cells(i, lnc + 3)        'ソルバーを実行した結果表示されるウエイト
        'ポートフォリオ期待利益率の制約条件として=実現期待利益率を設定
        SolverAdd CellRef:=PortfolioReturn, _
            Relation:=2, _
            FormulaText:=RealizedReturn
        Dim SolverResult As Integer     'ソルバーを実行した際の戻り値
        'ソルバーを実行
        SolverResult = SolverSolve(UserFinish:=True)
        'ソルバーで解が求められない場合にエラーを表示
        If SolverResult > 1 Then
            MsgBox "リターン" & addWs.Cells(i, lnc + 1) & "の実行可能解が見つかりませんでした。"
        Else
            'エラーが無ければ
            'ソルバーを実行した結果表示されたポートフォリオ標準偏差を所定の位置にコピー
            Set SetObjectiveCells = addWs.Cells(lnr + 3 * sr + 16, lnc)
            SetObjectiveCells.Copy
            RiskOutcome.PasteSpecial xlPasteValues
            'ソルバーを実行した結果表示されたウエイトを所定の位置にコピー
            Set ChangingVariableCells = addWs.Range(Cells(lnr + sr + 3, lnc), Cells(lnr + sr + 3, lnc + sr))
            ChangingVariableCells.Copy WeightsOutcome
        End If
    Next i
    
    '●最小分散フロンティアの作図
    Dim chart1 As ChartObject       '最小分散フロンティアのチャート名
    Set chart1 = addWs.ChartObjects.Add(10, 10, 300, 200)       '最小分散フロンティアのチャートの位置とサイズを設定
    '最小分散フロンティアのチャートを移動
    With chart1
        .Left = Cells(lnr + 3 * sr + 20 + counter, lnc).Left
        .Top = Cells(lnr + 3 * sr + 20 + counter, lnc).Top
    End With
    '最小分散フロンティアのチャートの体裁を整える
    With chart1.Chart
        .ChartType = xlXYScatterSmoothNoMarkers     'チャートの種類を平滑線付き散布図(データマーカーなし)に設定
        .SetSourceData Range(Cells(lnr + 3 * sr + 19, lnc), Cells(lnr + 3 * sr + 18 + counter, lnc + 1))     'チャートのデータ範囲を指定
        .HasTitle = True     'チャートのタイトルを追加
        .ChartTitle.Text = "最小分散フロンティア"     'チャートのタイトルを設定
        'チャートのタイトルのフォントの設定
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14      'フォントサイズ
            .Bold = msoFalse        '太字を無効
            .Italic = msoFalse      'イタリックを無効
            .Name = "Meiryo UI"     '字体
        End With
        'チャートの凡例を無効に設定
        .HasLegend = False
        'チャートのY軸の設定
        With .Axes(xlValue)
            .HasTitle = True      '軸ラベルを有効
            .AxisTitle.Text = "期待利益率"      '軸ラベル名
            '軸ラベル名のフォントの設定
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      'フォントサイズ
                .Bold = msoFalse        '太字を無効
                .Italic = msoFalse      'イタリックを無効
                .Name = "Meiryo UI"     '字体
            End With
            .MajorGridlines.Delete      '目盛り線を無効
            .TickLabels.NumberFormat = "0.00"      '軸ラベルの小数点以下表示桁数
        End With
        'チャートのX軸の設定
        With .Axes(xlCategory)
            .HasTitle = True      '軸ラベルを有効
            .AxisTitle.Text = "標準偏差"      '軸ラベル名
            '軸ラベル名のフォントの設定
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      'フォントサイズ
                .Bold = msoFalse        '太字を無効
                .Italic = msoFalse      'イタリックを無効
                .Name = "Meiryo UI"     '字体
            End With
            .MajorGridlines.Delete      '目盛り線を無効
            .TickLabels.NumberFormat = "0.00"      '軸ラベルの小数点以下表示桁数
        End With
        'チャートの平滑線の設定
        With .SeriesCollection(1).Format.Line
        .weight = 1.5        '線幅
        .ForeColor.RGB = RGB(30, 80, 150)        '線の色
        End With
        '表示されているポートフォリオ標準偏差の中で最も小さい値の点をマーカーとして描画する
        Dim xValues As Variant      '最小のポートフォリオ標準偏差の値
        Dim minX As Double      '最小値を保持
        Dim minIndex As Long      'チャートのデータ範囲の中での最小のポートフォリオ標準偏差の値の位置
        Dim v As Long      'ループ処理用のカウンタ
        xValues = .SeriesCollection(1).xValues
        minX = xValues(1)
        minIndex = 1
        For v = LBound(xValues) To UBound(xValues)
            If xValues(v) < minX Then
                minX = xValues(v)
                minIndex = v
            End If
        Next v
        'チャートのマーカーの設定
        With .SeriesCollection(1).Points(minIndex)
            .MarkerStyle = xlMarkerStyleCircle      'マーカーのスタイル
            .MarkerSize = 3      'マーカーのサイズ
            .Format.Fill.ForeColor.RGB = RGB(30, 80, 150)      'マーカーの色
            .HasDataLabel = True      'データラベルを有効
            'データラベルの設定
            With .DataLabel
                .Text = "最小分散ポートフォリオ"      'データラベルの内容
                .Font.Name = "Meiryo UI"      '字体
                .Font.Size = 6      'フォントサイズ
                .Font.Bold = False        '太字を無効
                .Format.Line.Visible = msoFalse        '枠線を無効
                .Format.Line.ForeColor.RGB = RGB(0, 0, 0)        '枠線の色
            End With
        End With
    End With

    '●各期待利益率における各銘柄の投資割合の作図
    Dim chart2 As ChartObject       '各期待利益率における各銘柄の投資割合のチャート名
    Set chart2 = addWs.ChartObjects.Add(10, 10, 600, 400)       '各期待利益率における各銘柄の投資割合のチャートの位置とサイズを設定
    Dim labelRange As Range     'X軸(縦軸)の目盛りに用いる実現期待利益率の範囲を指定
    Set labelRange = addWs.Range(Cells(lnr + 3 * sr + 19, lnc + 1), Cells(lnr + 3 * sr + 18 + counter, lnc + 1))
    '各期待利益率における各銘柄の投資割合のチャートを移動
    With chart2
        .Left = Cells(lnr + 3 * sr + 20 + counter, lnc + 7).Left
        .Top = Cells(lnr + 3 * sr + 20 + counter, lnc + 7).Top
    End With
    '各期待利益率における各銘柄の投資割合のチャートの体裁を整える
    With chart2.Chart
        .ChartType = xlBarStacked100     'チャートの種類を100% 積み上げ横棒に設定
        .SetSourceData Range(Cells(lnr + 3 * sr + 18, lnc + 3), Cells(lnr + 3 * sr + 18 + counter, lnc + sr + 3))     'チャートのデータ範囲を指定
        .HasTitle = True     'チャートのタイトルを追加
        .ChartTitle.Text = "各期待利益率における各銘柄の投資割合"     'チャートのタイトルを設定
        'チャートのタイトルのフォントの設定
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14      'フォントサイズ
            .Bold = msoFalse        '太字を無効
            .Italic = msoFalse      'イタリックを無効
            .Name = "Meiryo UI"      '字体
        End With
        'チャートの凡例を有効に設定
        .HasLegend = True
        'チャートの凡例のフォントの設定
        With .Legend.Format.TextFrame2.TextRange.Font
            .Size = 8      'フォントサイズ
            .Bold = msoFalse        '太字を無効
            .Italic = msoFalse      'イタリックを無効
            .Name = "Meiryo UI"      '字体
        End With
        'チャートの要素の間隔を0に設定
        .ChartGroups(1).GapWidth = 0
        'チャートのY軸(横軸)の設定
        With .Axes(xlValue)
            .HasTitle = True      '軸ラベルを有効
            .AxisTitle.Text = "投資割合"      '軸ラベル名
            '軸ラベル名のフォントの設定
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      'フォントサイズ
                .Bold = msoFalse        '太字を無効
                .Italic = msoFalse      'イタリックを無効
                .Name = "Meiryo UI"      '字体
            End With
            .MajorGridlines.Delete      '目盛り線を無効
        End With
        'チャートのX軸(縦軸)の設定
        With .Axes(xlCategory)
            .HasTitle = True      '軸ラベルを有効
            .AxisTitle.Text = "期待利益率"      '軸ラベル名
            '軸ラベル名のフォントの設定
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      'フォントサイズ
                .Bold = msoFalse        '太字を無効
                .Italic = msoFalse      'イタリックを無効
                .Name = "Meiryo UI"      '字体
            End With
            .MajorGridlines.Delete      '目盛り線を無効
            .CategoryNames = labelRange.Value       'X軸(縦軸)の目盛りのデータ範囲を設定
            .TickLabels.NumberFormat = "0.0000"      '軸ラベルの小数点以下表示桁数
        End With
    End With
    
    addWs.Cells(1, 1).Select     'セルA1を選択
    Application.CutCopyMode = False     'コピーの無効
    Application.EnableEvents = True    'イベントを開始
    Application.ScreenUpdating = True  '画面更新を開始
    Application.Calculation = xlCalculationAutomatic  '計算を自動
    Unload ufm完全版_最小分散フロンティア      'ユーザーフォームを閉じる
    MsgBox "処理が完了しました。", vbInformation      '終了メッセージ
End Sub
'___________________________________________________________________________________________________
'___________________________________________________________________________________________________

'クリアボタン
Private Sub cmdC_Click()
    Dim ctrls As Control         'テキストボックス
    'すべてのテキストボックスを空にする
    For Each ctrls In Controls
        If TypeName(ctrls) = "TextBox" Then _
            ctrls.Value = ""
    Next ctrls
End Sub


