VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmMVP 
   Caption         =   "���S��_�ŏ����U�t�����e�B�A"
   ClientHeight    =   4710
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   6525
   OleObjectBlob   =   "ufmMVP.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ufmMVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���̃R�[�h�̓��[�U�[�t�H�[���p�ɋL�q���Ă��܂��B _
    ���p�ɂ̓��[�U�[�t�H�[�����K�v�ł��B
'VBA�𗘗p����ꍇ�́A�}�N���L���u�b�N�Ƃ��ĕۑ�����K�v������̂Œ��ӁB
'�\���o�[�̋@�\�𗘗p���Ă���̂ŁAExcel�̃A�h�C���Ń\���o�[�A�h�C���Ƀ`�F�b�N����ꂽ��ŁA�c�[��>�Q�Ɛݒ肩��Solver�Ƀ`�F�b�N������K�v������܂��B
'���s����Ə����ɐ����Ԃ�����ꍇ������̂Œ��ӁB
'�R�[�h�̐����̓l�b�g������ChatGPT����LLM�����p����Ɨǂ��Ǝv���܂��B
'����Sheet���F���͂�����������񂪋L�^����Ă���Sheet������͂��Ă��������B
'�V�KSheet���F���͌��ʂ�\������Sheet������͂��Ă�������(Sheet�������ō쐬����܂�)�B
'�����f�[�^�s�E��F���͂������������͈̔͂��w�肵�Ă�������(�͈͂��L���قǏ����Ɏ��Ԃ�������܂�)�B _
    �����̕����̂ݎw�肵�A��Ж���،��R�[�h�A���ԂȂǂ͊܂߂Ȃ��ł��������B _
    �������͍s�����Ɋ��ԁA������ɖ����𗅗񂵂Ă��������B _
    �s�̎w��͔��p�����A��̎w��͑啶���܂��͏������̔��p�p���ōs���Ă��������B
'�������ɃR�[�h��F���͂����������̖��̂܂��̓R�[�h�͈̔͂��w�肵�Ă��������B _
    �����̖��̂܂��̓R�[�h�͊����f�[�^�Ɠ���̍s�ɋL�^���Ă��������B _
    ��̎w��͑啶���܂��͏������̔��p�p���ōs���Ă��������B
'�Œᓊ�������F�e�����ɓ�������Œ���̊�����0�ȏ�̔��p�����œ��͂��Ă��������B _
    0����͂����ꍇ�͓������Ȃ�������������\��������܂��
'���җ��v���̒i�K�F���͂�����җ��v���ׂ̍����̒��x���w�肵�Ă��������B _
    10���x�ł��\���ȍŏ����U�t�����e�B�A�̍�}���\�ł�(�������傫���قǏ����Ɏ��Ԃ�������܂�)�B
'�I�v�V�����{�^���F�ǂ��܂ł̏��������s���邩��3�i�K�őI���ł��܂��B _
    {���O���^�[���܂�}�A{�|�[�g�t�H���I�W���΍��܂�}�͔�r�I�Z���Ԃŏ������I�����܂��B _
    {�ŏI���U�t�����e�B�A��}�܂�}�͏����Ɏ��Ԃ�v���܂��B
'�N���A�{�^���F���͗������ׂċ󗓂ɖ߂��܂��B
'���s�{�^���F���������s���܂�(���͗��𖄂߂�܂ł͉����܂���)�B


'��(�p������)��(��������)�֕ϊ�
Function ColumnName2Idx(ByVal colName As String) As Integer
    ColumnName2Idx = Columns(colName).Column
End Function

'��(��������)��(�p������)�֕ϊ�
Function ColumnIdx2Name(ByVal colNum As Integer) As String
    ColumnIdx2Name = Split(Columns(colNum).Address, "$")(2)
End Function

'������̑S�Ă̕������p���̏ꍇ��True�A�����łȂ��ꍇ��False��Ԃ�
Function IsAlpha(str As String) As Boolean
    IsAlpha = Not str Like "*[!a-zA-Z��-���`-�y]*"
End Function

'�e�L�X�g�{�b�N�X�̏����ݒ�
'�����l���D�F�œ���
Private Sub UserForm_Initialize()
    cmdEX.Enabled = False
    txtSheet1.Value = "�����f�[�^�T���v��"
    txtSheet1.ForeColor = &HC0C0C0
    txtSheet2.Text = "�ŏ����U�t�����e�B�A"
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

'���ׂẴe�L�X�g�{�b�N�X�ɓ��͂������Ԃ͎��s�{�^���𖳌���
Private Sub CheckTextBoxes()
    Dim ctrl As Control     '�e�L�X�g�{�b�N�X
    Dim allFilled As Boolean        '���ׂẴe�L�X�g�{�b�N�X�ɓ��͂���Ă��邩
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

'�e�L�X�g�{�b�N�X�̏����ݒ�
'�}�E�X�œ��͂��n�߂�ƃe�L�X�g�{�b�N�X�̏����l�������ĕ����F�����F�ɂ���
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
'���s�{�^���𐄂����ۂ̓���
Private Sub cmdEX_Click()
    Application.ScreenUpdating = False  '��ʍX�V���~
    Application.EnableEvents = False    '�C�x���g��}��
    Application.Calculation = xlCalculationManual   '�v�Z���蓮
    
    '�e�e�L�X�g�{�b�N�X�̓��͂Ɋւ���G���[��\��
    If Not IsNumeric(txtStock1.Text) Then
        MsgBox "�����f�[�^ �s" & txtStock1.Text & "�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    If Not IsNumeric(txtStock2.Text) Then
        MsgBox "�����f�[�^ �s" & txtStock2.Text & "�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    If Not IsAlpha(txtStock3.Text) Then
        MsgBox "�����f�[�^ ��" & txtStock3.Text & "�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    If Not IsAlpha(txtStock4.Text) Then
        MsgBox "�����f�[�^ ��" & txtStock4.Text & "�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    If Not IsAlpha(txtName.Text) Then
        MsgBox "�������ɃR�[�h ��" & txtName.Text & "�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    If Not IsNumeric(txtMinweight.Text) Then
        MsgBox "�Œᓊ������" & txtMinweight.Text & "%�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    If Not IsNumeric(txtStep.Text) Then
        MsgBox "���җ��v�̒i�K" & txtStep.Text & "�i�K�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    
    '�Œᓊ���������S�����ϓ������̏ꍇ�𒴉߂���ꍇ�ɂ̓G���[��\��
    Dim Stock1 As Integer, Stock2 As Integer, Stock3 As Integer, Stock4 As Integer
    Stock1 = txtStock1.Text     '�����f�[�^�̊J�n�s
    Stock2 = txtStock2.Text     '�����f�[�^�̍ŏI�s
    Dim sr As Integer       '�����f�[�^�I���s-�J�n�s
    sr = Stock2 - Stock1
    Dim MinWeight As Double     '�Œᓊ�������̐��l
    MinWeight = txtMinweight.Text / 100
    If 1 / (sr + 1) < MinWeight Then
        MsgBox "�Œᓊ������" & txtMinweight.Text & "%�͕s���ł��B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    
    '����Sheet����ExcelBook�ɑ��݂��Ȃ��ꍇ�ɃG���[��\��
    Dim flg2 As Boolean     '����Sheet�������݂��邩
    Dim chkWs As Worksheet      'ExcelBook����Sheet
    flg2 = False
    For Each chkWs In Worksheets
        If chkWs.Name = txtSheet1.Text Then
            flg2 = True
            Exit For
        End If
    Next chkWs
    If flg2 = False Then
        MsgBox "Sheet���u" & txtSheet1.Text & "�v�͑��݂��܂���B", vbCritical + vbOKOnly, "�G���["
        Exit Sub
    End If
    
    '�V�KSheet��������ExcelBook�ɑ��݂���ꍇ�ɃG���[��\��
    Dim flg As Boolean      '�V�KSheet����������
    Dim addWs As Worksheet      '�V�KSheet
    flg = True
    For Each chkWs In Worksheets
        If chkWs.Name = txtSheet2.Text Then
            flg = False
            MsgBox "Sheet���u" & txtSheet2.Text & "�v�͊����ł��B", vbCritical + vbOKOnly, "�G���["
            Exit Sub
        End If
    Next chkWs
    
    '�V�KSheet����Sheet���Ō���ɒǉ�
    If flg Then
        Set addWs = Worksheets.Add(After:=Sheets(Worksheets.Count))
        addWs.Name = txtSheet2.Text
    End If
    
    '��������񂩂烍�O���^�[���̌v�Z
    Stock3 = ColumnName2Idx(txtStock3.Text)     '�����f�[�^�̊J�n��(�����ϊ�)
    Stock4 = ColumnName2Idx(txtStock4.Text)     '�����f�[�^�̍ŏI��(�����ϊ�)
    Dim stkWs As Worksheet: Set stkWs = Worksheets(txtSheet1.Text)     '����Sheet
    Dim sc As Integer     '�����f�[�^�̊��ԗ�-1
    Dim lnr As Integer      '�V�K���[�N�V�[�g�̊�ʒu(�s)
    Dim lnc As Integer      '�V�K���[�N�V�[�g�̊�ʒu(��)
    Dim Stock3a As String       '�����f�[�^�̊J�n��+1(�p���ϊ�)
    lnr = 2
    lnc = 2
    sc = Stock4 - Stock3 - 1
    Stock3a = ColumnIdx2Name(Stock3 + 1)
    '���O���^�[���̃��x����\��
    addWs.Cells(lnr - 1, lnc) = "���O���^�[��"
    '���O���^�[����\��
    addWs.Range(Cells(lnr, lnc), Cells(lnr + sr, lnc + sc)).Formula _
        = "=IFERROR(LN('" & stkWs.Name & "'!" & Stock3a & txtStock1.Text & "/'" & stkWs.Name & "'!" & txtStock3.Text & txtStock1.Text & "),0)"
'    '���O���^�[���̌v�Z�ɃG���[���������ꍇ��0��\��
'    Dim lrTarget As Range
'    On Error Resume Next
'    Set lrTarget = addWs.Range(Cells(lnr, lnc), Cells(lnr + sr, lnc + sc)).SpecialCells(xlCellTypeFormulas, xlErrors)
'    On Error GoTo 0
'    If Not lrTarget Is Nothing Then
'        lrTarget.Value = 0
'    End If
    
    '�������ɃR�[�h�����O���^�[���̍����ɕ\��
    addWs.Range(Cells(lnr, lnc - 1), Cells(lnr + sr, lnc - 1)).Formula _
        = "='" & stkWs.Name & "'!" & txtName.Text & txtStock1.Text
    
    '�����O���^�[���̕��ςƕW���΍�
    Dim lnca As String      '�V�K���[�N�V�[�g�̃��O���^�[���J�n��(�p���ϊ�)
    Dim lncb As String      '�V�K���[�N�V�[�g�̃��O���^�[���ŏI��(�p���ϊ�)
    lnca = ColumnIdx2Name(lnc)
    lncb = ColumnIdx2Name(lnc + sc)
    '���O���^�[�����ς̃��x����\��
    addWs.Cells(lnr - 1, lnc + sc + 2) = "���O���^�[������"
    '���O���^�[�����ς̊֐����L�q
    addWs.Range(Cells(lnr, lnc + sc + 2), Cells(lnr + sr, lnc + sc + 2)).Formula _
        = "=AVERAGE(" & lnca & lnr & ":" & lncb & lnr & ")"
    '���O���^�[���W���΍��̃��x����\��
    addWs.Cells(lnr - 1, lnc + sc + 3) = "���O���^�[���W���΍�"
    '���O���^�[���W���΍��̊֐����L�q
    addWs.Range(Cells(lnr, lnc + sc + 3), Cells(lnr + sr, lnc + sc + 3)).Formula _
        = "=STDEV.P(" & lnca & lnr & ":" & lncb & lnr & ")"
    
    '���I�v�V�����{�^���Łu���O���^�[���܂Łv��I�������ꍇ�͂����܂łŏ������I��
    If opb1 Then
        Application.EnableEvents = True    '�C�x���g���J�n
        Application.ScreenUpdating = True  '��ʍX�V���J�n
        Application.Calculation = xlCalculationAutomatic   '�v�Z������
        Unload ufm���S��_�ŏ����U�t�����e�B�A       '���[�U�[�t�H�[�������
        MsgBox "�������������܂����B", vbInformation      '�I�����b�Z�[�W
        Exit Sub
    End If
    
    '���E�G�C�g�̏����ݒ�
    Dim lncc As String      '�V�K���[�N�V�[�g�̃E�G�C�g�ŏI��(�p���ϊ�)
    lncc = ColumnIdx2Name(lnc + sr)
    '�ꎞ�I�ɂ��ׂĂ̖����̃E�G�C�g���Œᓊ�������ɐݒ�
    addWs.Range(Cells(lnr + sr + 3, lnc), Cells(lnr + sr + 3, lnc + sr)) = MinWeight
    '�E�G�C�g���v�̃��x����\��
    addWs.Cells(lnr + sr + 2, lnc + sr + 2) = "�E�G�C�g���v"
    '�E�G�C�g���v�̊֐����L�q
    addWs.Cells(lnr + sr + 3, lnc + sr + 2) _
        = "=SUM(" & lnca & lnr + sr + 3 & ":" & lncc & lnr + sr + 3 & ")"
    '�������ɃR�[�h���E�G�C�g�̏�ɉ����тɕ\��
    Dim n As Integer        '�������ɃR�[�h�̔Ԗ�
    Dim lncf As String      '�ŏ��ɕ\�������������ɃR�[�h�̕\����(�p���ϊ�)
    lncf = ColumnIdx2Name(lnc - 1)
    For n = 0 To sr
        addWs.Cells(lnr + sr + 2, lnc + n) _
            = "=" & lncf & lnr + n
    Next n
    
    '���|�[�g�t�H���I�̊��җ��v��
    Dim lncd As String      '���O���^�[�����ς̊֐����L�q����Ă����(�p���ϊ�)
    lncd = ColumnIdx2Name(lnc + sc + 2)
    '�|�[�g�t�H���I���җ��v���̃��x����\��
    addWs.Cells(lnr + sr + 5, lnc) = "�|�[�g�t�H���I���җ��v��"
    '�|�[�g�t�H���I���җ��v�������߂�֐��̋L�q
    addWs.Cells(lnr + sr + 6, lnc) _
        = "=MMULT(" & lnca & lnr + sr + 3 & ":" & lncc & lnr + sr + 3 & "," & lncd & lnr & ":" & lncd & lnr + sr & ")"
    
    '�����U�����U�s��
    '���U�����U�s��̃��x����\��
    addWs.Cells(lnr + sr + 8, lnc - 1) = "���U�����U�s��"
    Dim var As Integer      '���O���^�[���̍s�̔Ԗ�
    For var = 0 To sr
    '���U�����U�s��̊֐����L�q
    addWs.Range(Cells(lnr + sr + 9, lnc + var), Cells(lnr + 2 * sr + 9, lnc + var)).Formula _
        = "=COVARIANCE.P(" & "$" & lnca & "$" & lnr + var & ":" & "$" & lncb & "$" & lnr + var & "," & "$" & lnca & lnr & ":" & "$" & lncb & lnr & ")"
    Next var
    Dim mm As Integer       '���U�����U�s��̍s�̔Ԗ�
    For mm = 0 To sr
        addWs.Cells(lnr + 2 * sr + 11, lnc + mm).Formula _
            = "=INDEX(MMULT(" & lnca & lnr + sr + 3 & ":" & lncc & lnr + sr + 3 & "," & lnca & lnr + sr + 9 & ":" & lncc & lnr + 2 * sr + 9 & ")," & mm + 1 & ")"
    Next mm
    '�������ɃR�[�h�𕪎U�����U�s��̏�ɉ����тɕ\��
    addWs.Range(Cells(lnr + sr + 8, lnc), Cells(lnr + sr + 8, lnc + sr)).Formula _
        = "=" & lnca & lnr + sr + 2
    '�������ɃR�[�h�𕪎U�����U�s��̍����ɏc���тɕ\��
    addWs.Range(Cells(lnr + sr + 9, lnc - 1), Cells(lnr + 2 * sr + 9, lnc - 1)).Formula _
        = "=" & lncf & lnr
    
    '���E�G�C�g�̕���
    Dim weights As Integer      '�E�G�C�g�̍s�̔Ԗ�
    Dim lnce As String      '�e�����̃E�G�C�g�̗�(�p���ϊ�)
    '�E�G�C�g�̏����ݒ�Őݒ肵���E�G�C�g���c���тɒu��������
    For weights = 0 To sr
        lnce = ColumnIdx2Name(lnc + weights)
        addWs.Cells(lnr + 2 * sr + 13 + weights, lnc).Formula _
            = "=" & lnce & lnr + sr + 3
    Next weights
    '�������ɃR�[�h���E�G�C�g�̍����ɏc���тɕ\��
    addWs.Range(Cells(lnr + 2 * sr + 13, lnc - 1), Cells(lnr + 3 * sr + 13, lnc - 1)).Formula _
        = "=" & lncf & lnr
    
    '���|�[�g�t�H���I�̕W���΍�
    '�|�[�g�t�H���I�W���΍��̃��x����\��
    addWs.Cells(lnr + 3 * sr + 15, lnc) = "�|�[�g�t�H���I�W���΍�"
    '�|�[�g�t�H���I�W���΍������߂�֐��̋L�q
    addWs.Cells(lnr + 3 * sr + 16, lnc).Formula _
        = "=SQRT(MMULT(" & lnca & lnr + 2 * sr + 11 & ":" & lncc & lnr + 2 * sr + 11 & "," & lnca & lnr + 2 * sr + 13 & ":" & lnca & lnr + 3 * sr + 13 & "))"
    
    '���I�v�V�����{�^���Łu�|�[�g�t�H���I�W���΍��܂Łv��I�������ꍇ�͂����܂łŏ������I��
    If opb2 Then
        Application.EnableEvents = True    '�C�x���g���J�n
        Application.ScreenUpdating = True  '��ʍX�V���J�n
        Application.Calculation = xlCalculationAutomatic  '�v�Z������
        Unload ufm���S��_�ŏ����U�t�����e�B�A       '���[�U�[�t�H�[�������
        MsgBox "�������������܂����B", vbInformation      '�I�����b�Z�[�W
        Exit Sub
    End If
    
    '���������җ��v���͈̔͂ƒi�K�̐ݒ�
    '�|�[�g�t�H���I�W���΍��̃��x����\��
    addWs.Cells(lnr + 3 * sr + 18, lnc) = "�|�[�g�t�H���I�W���΍�"
    '�|�[�g�t�H���I���җ��v���̃��x����\��
    addWs.Cells(lnr + 3 * sr + 18, lnc + 1) = "�|�[�g�t�H���I���җ��v��"
    Dim MaxWeight As Double     '�ݒ肵���Œᓊ�������̏�Ŏ����\�ȍő�̓�������
    Dim MaxReturn As Double     '�����\�ȍő�̃|�[�g�t�H���I���җ��v��
    Dim MinReturn As Double     '�����\�ȍŒ�̃|�[�g�t�H���I���җ��v��
    Dim MaxRe As Double     '���ׂĂ̖����̒��ōł��������O���^�[������
    Dim MinRe As Double     '���ׂĂ̖����̒��ōł��Ⴂ���O���^�[������
    Dim DifReturn As Double     '�����\�ȃ|�[�g�t�H���I���җ��v���͈̔͂����җ��v���̒i�K�ŕ������Ƃ���1�i�K�̒l
    Dim r As Double     'MinReturn��DifReturn�����������ۂɕ\������l
    Dim counter As Integer     'MinReturn��DifReturn����������
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
    '�������ɃR�[�h���c���тɕ\��
    addWs.Range(Cells(lnr + 3 * sr + 18, lnc + 3), Cells(lnr + 3 * sr + 18, lnc + sr + 3)).Formula _
        = "=" & lnca & lnr + sr + 2

    '���\���o�[�̎��s
    '�\���o�[�����Z�b�g
    SolverReset
    '�\���o�[�̐��x���w��
    SolverOptions Precision:=0.000001       '�����l�F0.000001
    Set SetObjectiveCells = addWs.Cells(lnr + 3 * sr + 16, lnc)     '�ړI�Z���͈̔͂��|�[�g�t�H���I�W���΍��Ɏw��
    Set ChangingVariableCells = addWs.Range(Cells(lnr + sr + 3, lnc), Cells(lnr + sr + 3, lnc + sr))    '�ϐ��Z���͈̔͂��E�G�C�g�Ɏw��
    Set PortfolioReturn = addWs.Cells(lnr + sr + 6, lnc)        '�|�[�g�t�H���I�̊��җ��v���͈̔͂��w��
    '�\���o�[�̖ړI�Z���A�ڕW�l�A�ϐ��Z���A�������@��ݒ�
    SolverOk SetCell:=SetObjectiveCells, _
        MaxMinVal:=2, _
        ByChange:=ChangingVariableCells, _
        EngineDesc:="GRG Nonlinear"
    Dim q As Integer        '�E�G�C�g�̍s�̔Ԗ�
    '�E�G�C�g�̐�������Ƃ���>=MinWeight��ݒ�
    For q = lnc To lnc + sr
        SolverAdd CellRef:=addWs.Cells(lnr + sr + 3, q), _
            Relation:=3, _
            FormulaText:=CDbl(MinWeight)
    Next q
    '�E�G�C�g���v�̐�������Ƃ���=1��ݒ�
    SolverAdd CellRef:=addWs.Cells(lnr + sr + 3, lnc + sr + 2), _
        Relation:=2, _
        FormulaText:=1
    '�|�[�g�t�H���I���җ��v���̐��������ݒ肵���������җ��v���ɏ]���ĕύX���Ȃ���\���o�[���ғ����Č��ʂ�\��
    Dim i As Integer        '�������җ��v���̔Ԗ�
    For i = lnr + 3 * sr + 19 To lnr + 3 * sr + 18 + counter
        '�|�[�g�t�H���I���җ��v���ɂ��Đݒ肳��Ă��鐧��������폜
        SolverDelete CellRef:=PortfolioReturn, _
            Relation:=2
        Set RealizedReturn = addWs.Cells(i, lnc + 1)        '�ݒ肵���������җ��v��
        Set RiskOutcome = addWs.Cells(i, lnc)       '�\���o�[�����s�������ʕ\�������|�[�g�t�H���I�W���΍�
        Set WeightsOutcome = addWs.Cells(i, lnc + 3)        '�\���o�[�����s�������ʕ\�������E�G�C�g
        '�|�[�g�t�H���I���җ��v���̐�������Ƃ���=�������җ��v����ݒ�
        SolverAdd CellRef:=PortfolioReturn, _
            Relation:=2, _
            FormulaText:=RealizedReturn
        Dim SolverResult As Integer     '�\���o�[�����s�����ۂ̖߂�l
        '�\���o�[�����s
        SolverResult = SolverSolve(UserFinish:=True)
        '�\���o�[�ŉ������߂��Ȃ��ꍇ�ɃG���[��\��
        If SolverResult > 1 Then
            MsgBox "���^�[��" & addWs.Cells(i, lnc + 1) & "�̎��s�\����������܂���ł����B"
        Else
            '�G���[���������
            '�\���o�[�����s�������ʕ\�����ꂽ�|�[�g�t�H���I�W���΍�������̈ʒu�ɃR�s�[
            Set SetObjectiveCells = addWs.Cells(lnr + 3 * sr + 16, lnc)
            SetObjectiveCells.Copy
            RiskOutcome.PasteSpecial xlPasteValues
            '�\���o�[�����s�������ʕ\�����ꂽ�E�G�C�g������̈ʒu�ɃR�s�[
            Set ChangingVariableCells = addWs.Range(Cells(lnr + sr + 3, lnc), Cells(lnr + sr + 3, lnc + sr))
            ChangingVariableCells.Copy WeightsOutcome
        End If
    Next i
    
    '���ŏ����U�t�����e�B�A�̍�}
    Dim chart1 As ChartObject       '�ŏ����U�t�����e�B�A�̃`���[�g��
    Set chart1 = addWs.ChartObjects.Add(10, 10, 300, 200)       '�ŏ����U�t�����e�B�A�̃`���[�g�̈ʒu�ƃT�C�Y��ݒ�
    '�ŏ����U�t�����e�B�A�̃`���[�g���ړ�
    With chart1
        .Left = Cells(lnr + 3 * sr + 20 + counter, lnc).Left
        .Top = Cells(lnr + 3 * sr + 20 + counter, lnc).Top
    End With
    '�ŏ����U�t�����e�B�A�̃`���[�g�̑̍ق𐮂���
    With chart1.Chart
        .ChartType = xlXYScatterSmoothNoMarkers     '�`���[�g�̎�ނ𕽊����t���U�z�}(�f�[�^�}�[�J�[�Ȃ�)�ɐݒ�
        .SetSourceData Range(Cells(lnr + 3 * sr + 19, lnc), Cells(lnr + 3 * sr + 18 + counter, lnc + 1))     '�`���[�g�̃f�[�^�͈͂��w��
        .HasTitle = True     '�`���[�g�̃^�C�g����ǉ�
        .ChartTitle.Text = "�ŏ����U�t�����e�B�A"     '�`���[�g�̃^�C�g����ݒ�
        '�`���[�g�̃^�C�g���̃t�H���g�̐ݒ�
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14      '�t�H���g�T�C�Y
            .Bold = msoFalse        '�����𖳌�
            .Italic = msoFalse      '�C�^���b�N�𖳌�
            .Name = "Meiryo UI"     '����
        End With
        '�`���[�g�̖}��𖳌��ɐݒ�
        .HasLegend = False
        '�`���[�g��Y���̐ݒ�
        With .Axes(xlValue)
            .HasTitle = True      '�����x����L��
            .AxisTitle.Text = "���җ��v��"      '�����x����
            '�����x�����̃t�H���g�̐ݒ�
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      '�t�H���g�T�C�Y
                .Bold = msoFalse        '�����𖳌�
                .Italic = msoFalse      '�C�^���b�N�𖳌�
                .Name = "Meiryo UI"     '����
            End With
            .MajorGridlines.Delete      '�ڐ�����𖳌�
            .TickLabels.NumberFormat = "0.00"      '�����x���̏����_�ȉ��\������
        End With
        '�`���[�g��X���̐ݒ�
        With .Axes(xlCategory)
            .HasTitle = True      '�����x����L��
            .AxisTitle.Text = "�W���΍�"      '�����x����
            '�����x�����̃t�H���g�̐ݒ�
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      '�t�H���g�T�C�Y
                .Bold = msoFalse        '�����𖳌�
                .Italic = msoFalse      '�C�^���b�N�𖳌�
                .Name = "Meiryo UI"     '����
            End With
            .MajorGridlines.Delete      '�ڐ�����𖳌�
            .TickLabels.NumberFormat = "0.00"      '�����x���̏����_�ȉ��\������
        End With
        '�`���[�g�̕������̐ݒ�
        With .SeriesCollection(1).Format.Line
        .weight = 1.5        '����
        .ForeColor.RGB = RGB(30, 80, 150)        '���̐F
        End With
        '�\������Ă���|�[�g�t�H���I�W���΍��̒��ōł��������l�̓_���}�[�J�[�Ƃ��ĕ`�悷��
        Dim xValues As Variant      '�ŏ��̃|�[�g�t�H���I�W���΍��̒l
        Dim minX As Double      '�ŏ��l��ێ�
        Dim minIndex As Long      '�`���[�g�̃f�[�^�͈͂̒��ł̍ŏ��̃|�[�g�t�H���I�W���΍��̒l�̈ʒu
        Dim v As Long      '���[�v�����p�̃J�E���^
        xValues = .SeriesCollection(1).xValues
        minX = xValues(1)
        minIndex = 1
        For v = LBound(xValues) To UBound(xValues)
            If xValues(v) < minX Then
                minX = xValues(v)
                minIndex = v
            End If
        Next v
        '�`���[�g�̃}�[�J�[�̐ݒ�
        With .SeriesCollection(1).Points(minIndex)
            .MarkerStyle = xlMarkerStyleCircle      '�}�[�J�[�̃X�^�C��
            .MarkerSize = 3      '�}�[�J�[�̃T�C�Y
            .Format.Fill.ForeColor.RGB = RGB(30, 80, 150)      '�}�[�J�[�̐F
            .HasDataLabel = True      '�f�[�^���x����L��
            '�f�[�^���x���̐ݒ�
            With .DataLabel
                .Text = "�ŏ����U�|�[�g�t�H���I"      '�f�[�^���x���̓��e
                .Font.Name = "Meiryo UI"      '����
                .Font.Size = 6      '�t�H���g�T�C�Y
                .Font.Bold = False        '�����𖳌�
                .Format.Line.Visible = msoFalse        '�g���𖳌�
                .Format.Line.ForeColor.RGB = RGB(0, 0, 0)        '�g���̐F
            End With
        End With
    End With

    '���e���җ��v���ɂ�����e�����̓��������̍�}
    Dim chart2 As ChartObject       '�e���җ��v���ɂ�����e�����̓��������̃`���[�g��
    Set chart2 = addWs.ChartObjects.Add(10, 10, 600, 400)       '�e���җ��v���ɂ�����e�����̓��������̃`���[�g�̈ʒu�ƃT�C�Y��ݒ�
    Dim labelRange As Range     'X��(�c��)�̖ڐ���ɗp����������җ��v���͈̔͂��w��
    Set labelRange = addWs.Range(Cells(lnr + 3 * sr + 19, lnc + 1), Cells(lnr + 3 * sr + 18 + counter, lnc + 1))
    '�e���җ��v���ɂ�����e�����̓��������̃`���[�g���ړ�
    With chart2
        .Left = Cells(lnr + 3 * sr + 20 + counter, lnc + 7).Left
        .Top = Cells(lnr + 3 * sr + 20 + counter, lnc + 7).Top
    End With
    '�e���җ��v���ɂ�����e�����̓��������̃`���[�g�̑̍ق𐮂���
    With chart2.Chart
        .ChartType = xlBarStacked100     '�`���[�g�̎�ނ�100% �ςݏグ���_�ɐݒ�
        .SetSourceData Range(Cells(lnr + 3 * sr + 18, lnc + 3), Cells(lnr + 3 * sr + 18 + counter, lnc + sr + 3))     '�`���[�g�̃f�[�^�͈͂��w��
        .HasTitle = True     '�`���[�g�̃^�C�g����ǉ�
        .ChartTitle.Text = "�e���җ��v���ɂ�����e�����̓�������"     '�`���[�g�̃^�C�g����ݒ�
        '�`���[�g�̃^�C�g���̃t�H���g�̐ݒ�
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14      '�t�H���g�T�C�Y
            .Bold = msoFalse        '�����𖳌�
            .Italic = msoFalse      '�C�^���b�N�𖳌�
            .Name = "Meiryo UI"      '����
        End With
        '�`���[�g�̖}���L���ɐݒ�
        .HasLegend = True
        '�`���[�g�̖}��̃t�H���g�̐ݒ�
        With .Legend.Format.TextFrame2.TextRange.Font
            .Size = 8      '�t�H���g�T�C�Y
            .Bold = msoFalse        '�����𖳌�
            .Italic = msoFalse      '�C�^���b�N�𖳌�
            .Name = "Meiryo UI"      '����
        End With
        '�`���[�g�̗v�f�̊Ԋu��0�ɐݒ�
        .ChartGroups(1).GapWidth = 0
        '�`���[�g��Y��(����)�̐ݒ�
        With .Axes(xlValue)
            .HasTitle = True      '�����x����L��
            .AxisTitle.Text = "��������"      '�����x����
            '�����x�����̃t�H���g�̐ݒ�
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      '�t�H���g�T�C�Y
                .Bold = msoFalse        '�����𖳌�
                .Italic = msoFalse      '�C�^���b�N�𖳌�
                .Name = "Meiryo UI"      '����
            End With
            .MajorGridlines.Delete      '�ڐ�����𖳌�
        End With
        '�`���[�g��X��(�c��)�̐ݒ�
        With .Axes(xlCategory)
            .HasTitle = True      '�����x����L��
            .AxisTitle.Text = "���җ��v��"      '�����x����
            '�����x�����̃t�H���g�̐ݒ�
            With .AxisTitle.Format.TextFrame2.TextRange.Font
                .Size = 10      '�t�H���g�T�C�Y
                .Bold = msoFalse        '�����𖳌�
                .Italic = msoFalse      '�C�^���b�N�𖳌�
                .Name = "Meiryo UI"      '����
            End With
            .MajorGridlines.Delete      '�ڐ�����𖳌�
            .CategoryNames = labelRange.Value       'X��(�c��)�̖ڐ���̃f�[�^�͈͂�ݒ�
            .TickLabels.NumberFormat = "0.0000"      '�����x���̏����_�ȉ��\������
        End With
    End With
    
    addWs.Cells(1, 1).Select     '�Z��A1��I��
    Application.CutCopyMode = False     '�R�s�[�̖���
    Application.EnableEvents = True    '�C�x���g���J�n
    Application.ScreenUpdating = True  '��ʍX�V���J�n
    Application.Calculation = xlCalculationAutomatic  '�v�Z������
    Unload ufm���S��_�ŏ����U�t�����e�B�A      '���[�U�[�t�H�[�������
    MsgBox "�������������܂����B", vbInformation      '�I�����b�Z�[�W
End Sub
'___________________________________________________________________________________________________
'___________________________________________________________________________________________________

'�N���A�{�^��
Private Sub cmdC_Click()
    Dim ctrls As Control         '�e�L�X�g�{�b�N�X
    '���ׂẴe�L�X�g�{�b�N�X����ɂ���
    For Each ctrls In Controls
        If TypeName(ctrls) = "TextBox" Then _
            ctrls.Value = ""
    Next ctrls
End Sub

