VERSION 5.00
Begin VB.Form frmCaseTendBodySetLine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ͼ�����ݱ༭"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmCaseTendBodySetLine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "˵��"
      Height          =   1485
      Left            =   165
      TabIndex        =   17
      Top             =   2400
      Width           =   5115
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1155
         TabIndex        =   4
         Top             =   645
         Width           =   3540
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   2
         Top             =   255
         Width           =   3540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "˵�����±꣩"
         Height          =   180
         Left            =   105
         TabIndex        =   3
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "˵�����ϱ꣩"
         Height          =   180
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   1080
      End
   End
   Begin zl9BodyEditorYDEY.VsfGrid vsf 
      Height          =   1950
      Left            =   165
      TabIndex        =   0
      Top             =   405
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   3440
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5355
      TabIndex        =   7
      Top             =   3630
      Width           =   1100
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "ʱ�������������(&N)"
      Height          =   210
      Left            =   4680
      TabIndex        =   8
      Top             =   150
      Width           =   2010
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5355
      TabIndex        =   6
      Top             =   855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5355
      TabIndex        =   5
      Top             =   420
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Height          =   2415
      Left            =   6780
      TabIndex        =   10
      Top             =   1590
      Width           =   3750
      Begin VB.TextBox txtItem 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   11
         Top             =   195
         Width           =   1320
      End
      Begin VB.ComboBox cboComment 
         Height          =   300
         Left            =   675
         TabIndex        =   14
         Top             =   1980
         Width           =   2610
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "����ѹ"
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   13
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lblComment 
         Caption         =   "˵��"
         Height          =   180
         Left            =   255
         TabIndex        =   12
         Top             =   2040
         Width           =   540
      End
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   735
      TabIndex        =   16
      Top             =   4170
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ʾ��"
      Height          =   180
      Left            =   165
      TabIndex        =   15
      Top             =   4155
      Width           =   540
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "ʱ�䣺2001��11��16�� 4��8ʱ"
      Height          =   180
      Left            =   195
      TabIndex        =   9
      Top             =   165
      Width           =   2430
   End
End
Attribute VB_Name = "frmCaseTendBodySetLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintNowCol As Integer
Private mintMinCol As Integer
Private mintMaxCol As Integer
Private mblnChanged As Boolean
Private mfrmParent As Object
Private mlng����ȼ� As Long
Private mint����Ӧ�� As Integer
Private mblnStart As Boolean
Private Enum mCol
    ��Ŀ = 1
    ��Ŀid
    ��Сֵ
    ���ֵ
    ��λֵ
    �����
    ����
    ���
    ��λ
    δ��˵��
End Enum

Private Enum GraphDataRow
    ���ı�־ = 0
    �������� = 1
    �ϱ�˵�� = 2
    ������־ = 3
    ��λ��־ = 4
    ��Ժ��־ = 5
    ת�Ʊ�־ = 6
    ������־ = 7
    ��Ժ��־ = 8
    ��Ʊ�־ = 9
    ���Ա�־ = 10
    �±�˵�� = 11
    �Ͽ���־ = 12
    ������־ = 13
    ����ʱ�� = 14
    δ��˵�� = 15
End Enum

Public Function ShowEdit(ByVal frmParent As Object, ByVal intNowCol As Integer, ByVal intMinCol As Integer, ByVal intMaxCol As Long, ByVal lng����ȼ� As Long, ByVal int����Ӧ�� As Integer) As Boolean
    
    mblnChanged = False
    mblnStart = True
    
    mint����Ӧ�� = int����Ӧ��
    
    If intNowCol = -1 Then Exit Function
    
    mintNowCol = intNowCol
    mintMinCol = intMinCol
    mintMaxCol = intMaxCol
    
    mlng����ȼ� = lng����ȼ�
    
    Set mfrmParent = frmParent
    
    Call InitData
    Call LoadNowData
    
'    vsf.SetFocus
    
    Me.Show 1
    
    ShowEdit = mblnChanged
    
End Function

'Private Sub chk_Click(Index As Integer)
'
'    If chk(0).Value = 1 Then
'        vsf.EditMode(mCol.����) = 0
'        vsf.EditMode(mCol.��λ) = 0
'        vsf.EditMode(mCol.���) = 0
'        vsf.Cell(flexcpText, 1, mCol.����, vsf.Rows - 1, mCol.��λ) = ""
'    Else
'        vsf.EditMode(mCol.����) = 1
'    End If
'End Sub

'Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        zlCommFun.PressKey vbKeyTab
'    End If
'End Sub

Private Sub chkContinue_Click()
    vsf.SetFocus
    vsf.ShowCell vsf.Row, vsf.Col
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim intCol As Integer
    Dim intCount As Integer
    Dim dbValue As Double
    Dim str��ʽ As String
    Dim aryValue() As String
    Dim aryPart() As String     '��λ
    Dim intRewrite As Integer
    
    On Error GoTo ErrHead
    
    With mfrmParent.GetmshScale
        '����ע��˵��
        aryValue = Split(.TextMatrix(GraphDataRow.���ı�־, mintNowCol + .FixedCols), ";")
        intRewrite = Val(aryValue(0))
        If Trim(txt(0).Text) <> "" Or Trim(txt(1).Text) <> "" Then
            '�������ݣ��൱�����ӻ��޸Ĳ���
            Select Case intRewrite
            Case 0
                aryValue(0) = 2
            Case 1
                aryValue(0) = 3
            Case 2
                aryValue(0) = 2
            Case 3
                aryValue(0) = 3
            Case 4
                aryValue(0) = 3
            End Select
        Else
            'û�����ݣ��൱��ɾ������
            Select Case intRewrite
            Case 0
                aryValue(0) = 0
            Case 1
                aryValue(0) = 4
            Case 2
                aryValue(0) = 0
            Case 3
                aryValue(0) = 4
            Case 4
                aryValue(0) = 4
            End Select
        End If
        .TextMatrix(GraphDataRow.���ı�־, mintNowCol + .FixedCols) = Join(aryValue, ";")
        .TextMatrix(GraphDataRow.�ϱ�˵��, mintNowCol + .FixedCols) = Trim(txt(0).Text)
        .TextMatrix(GraphDataRow.�±�˵��, mintNowCol + .FixedCols) = Trim(txt(1).Text)
'        .TextMatrix(GraphDataRow.�Ͽ���־, mintNowCol + .FixedCols) = chk(0).Value

        '������������
        For intCount = 1 To vsf.Rows - 1
        
            aryValue = Split(.TextMatrix(GraphDataRow.���ı�־, mintNowCol + .FixedCols), ";")
            intRewrite = Val(aryValue(intCount))
            
            If Trim(vsf.TextMatrix(intCount, mCol.����)) <> "" Or Trim(vsf.TextMatrix(intCount, mCol.���)) <> "" Or Trim(vsf.TextMatrix(intCount, mCol.δ��˵��)) <> "" Then
                '�������ݣ��൱�����ӻ��޸Ĳ���
                Select Case intRewrite
                Case 0
                    aryValue(intCount) = 2
                Case 1
                    aryValue(intCount) = 3
                Case 2
                    aryValue(intCount) = 2
                Case 3
                    aryValue(intCount) = 3
                Case 4
                    aryValue(intCount) = 3
                End Select
            Else
                'û�����ݣ��൱��ɾ������
                Select Case intRewrite
                Case 0
                    aryValue(intCount) = 0
                Case 1
                    aryValue(intCount) = 4
                Case 2
                    aryValue(intCount) = 0
                Case 3
                    aryValue(intCount) = 4
                Case 4
                    aryValue(intCount) = 4
                End Select
            End If
            
            .TextMatrix(GraphDataRow.���ı�־, mintNowCol + .FixedCols) = Join(aryValue, ";")
            
            aryValue = Split(.TextMatrix(GraphDataRow.��������, mintNowCol + .FixedCols), ";")
            If Trim(vsf.TextMatrix(intCount, mCol.����)) <> "" Then
            
                dbValue = ((Val(vsf.TextMatrix(intCount, mCol.���ֵ)) - Val(vsf.TextMatrix(intCount, mCol.����))) / Val(vsf.TextMatrix(intCount, mCol.��λֵ)) + Val(vsf.TextMatrix(intCount, mCol.�����)) - 1) * .ROWHEIGHT(1)
                aryValue(intCount) = dbValue
                
                If Trim(vsf.TextMatrix(intCount, mCol.���)) <> "" Then
                    dbValue = ((Val(vsf.TextMatrix(intCount, mCol.���ֵ)) - Val(vsf.TextMatrix(intCount, mCol.���))) / Val(vsf.TextMatrix(intCount, mCol.��λֵ)) + Val(vsf.TextMatrix(intCount, mCol.�����)) - 1) * .ROWHEIGHT(1)
                    aryValue(intCount) = aryValue(intCount) & "," & dbValue
                End If
            Else
                aryValue(intCount) = ""
            End If
            .TextMatrix(GraphDataRow.��������, mintNowCol + .FixedCols) = Join(aryValue, ";")
            
            '�������µ�"����"
            If Trim(vsf.TextMatrix(intCount, mCol.����)) = "����" And vsf.RowData(intCount) = 1 Then
                aryValue = Split(.TextMatrix(GraphDataRow.δ��˵��, mintNowCol + .FixedCols), ";")
                aryValue(intCount) = "����"
                .TextMatrix(GraphDataRow.δ��˵��, mintNowCol + .FixedCols) = Join(aryValue, ";")
            Else
                aryValue = Split(.TextMatrix(GraphDataRow.δ��˵��, mintNowCol + .FixedCols), ";")
                aryValue(intCount) = vsf.TextMatrix(intCount, mCol.δ��˵��)
                .TextMatrix(GraphDataRow.δ��˵��, mintNowCol + .FixedCols) = Join(aryValue, ";")
            End If
            
            Select Case Val(vsf.RowData(intCount))
            Case 1, 2, 3            '.TextMatrix(GraphDataRow.��λ��־, mintNowCol + .FixedCols)
                str��ʽ = ""
                If Trim(vsf.TextMatrix(intCount, mCol.��λ)) <> "" Then
                    str��ʽ = Trim(vsf.TextMatrix(intCount, mCol.��λ))
                Else
                    If Val(vsf.RowData(intCount)) = 1 Then str��ʽ = "Ҹ��"
                End If
                
                '��֯��λ
                aryPart = Split(.TextMatrix(GraphDataRow.��λ��־, mintNowCol + .FixedCols), ";")
                aryPart(intCount) = str��ʽ
                .TextMatrix(GraphDataRow.��λ��־, mintNowCol + .FixedCols) = Join(aryPart, ";")
                
            End Select
        Next

    End With
    
    '�����ϼ��������ͼ�δ���
    Call mfrmParent.DrawPaper
    Call mfrmParent.DrawGraph
    
    mblnChanged = True
    
    If chkContinue.Value = 0 Then
        Unload Me
        Exit Sub
    End If
    
    If mintNowCol < mintMaxCol Then
        mintNowCol = mintNowCol + 1
    Else
        'If MsgBox("�Ѿ��ﵽ�����±����ʱ�䣬�Ƿ��������룿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Unload Me
        Exit Sub
'        Else
'            intNowCol = intMinCol
'        End If
    End If
    
    Call LoadNowData
    
    vsf.SetFocus
    vsf.ShowCell 1, 2
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHead

    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "��Ŀ", 1080, 1
        .NewColumn "��Ŀid", 0, 1
        .NewColumn "��Сֵ", 0, 1
        .NewColumn "���ֵ", 0, 1
        .NewColumn "��λֵ", 0, 1
        .NewColumn "�����", 0, 1
        .NewColumn "����", 900, 1, , 1
        .NewColumn "���", 900, 1
        .NewColumn "��λ", 750, 1
        .NewColumn "δ��˵��", 1080, 1, "...", 1
        
        .FixedCols = 2
                
        .Body.ColHidden(mCol.��Сֵ) = True
        .Body.ColHidden(mCol.���ֵ) = True
        .Body.ColHidden(mCol.��λֵ) = True
        .Body.ColHidden(mCol.�����) = True
        .Body.ColHidden(mCol.��Ŀid) = True
        
        .Body.WordWrap = True
    End With

    InitData = True
    
    Exit Function
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Private Sub txtComment_GotFocus()
'    zlControl.TxtSelAll cboComment
'    zlCommFun.OpenIme True
'End Sub
'
'Private Sub txtComment_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc(";") Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
'    If KeyAscii = 13 Then
'        KeyAscii = 0
'        cmdOK.SetFocus
'    End If
'End Sub
'
'Private Sub txtComment_LostFocus()
'    cboComment.Text = Replace(Me.cboComment.Text, "'", "")
'    zlCommFun.OpenIme
'End Sub

Private Sub Form_Activate()
    
    If mblnStart = False Then Exit Sub
    mblnStart = False
    
    vsf.Col = mCol.����
    vsf.SetFocus
    
End Sub

'Private Sub txt_Change(Index As Integer)
'
'    Select Case txt(Index).Text
'    Case "�ܲ�", "δ��", "���", "���"
'
'        chk(0).Value = 1
'
'    Case Else
'
'        chk(0).Value = 0
'
'    End Select
'
'End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    Call zlControl.TxtSelAll(txt(Index))
    
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme True
    End Select
    
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)

    
End Sub



Private Sub txtItem_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtItem(Index)
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If KeyAscii = Asc("'") Or KeyAscii = Asc(";") Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    txtItem(Index).Text = Replace(Me.txtItem(Index).Text, "'", "")
End Sub

Private Sub txtItem_Validate(Index As Integer, Cancel As Boolean)
    Dim aryPara() As String
    
    On Error GoTo ErrHead
    
    '��ȡָ����Ŀ���壺���ֵ����Сֵ����λֵ�������
    If Trim(txtItem(Index).Text) = "" Then Exit Sub
    
    If Not MouseInRect(cmdCanc.hWnd) Then
        If IsNumeric(Trim(Me.txtItem(Index).Text)) = False Then
            MsgBox "��Ŀ��" & Me.lblItem(Index).Tag & "����ֵ����Ϊ���֣�", vbExclamation, gstrSysName
            Cancel = True: Exit Sub
        End If
        aryPara = Split(Me.txtItem(Index).Tag, ";")
        If Format(Me.txtItem(Index).Text) > Val(aryPara(0)) Or Format(Me.txtItem(Index).Text) < Val(aryPara(1)) Then
            MsgBox "��Ŀ��" & Me.lblItem(Index).Tag & "����ֵ��������Χ��" & aryPara(1) & "��" & aryPara(0), vbExclamation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadNowData()
    
    'װ��ָ����������
    Dim aryValue() As String
    Dim aryNote() As String
    Dim aryPara() As String
    Dim dtNow As Date
    Dim dtNowTmp As Date
    Dim lngHourBegin As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intCount As Integer
    Dim intCol As Long
    Dim dblValue As Double
    Dim strTmp As String
    
    On Error GoTo ErrHead
    
    lngHourBegin = Val(zlDatabase.GetPara(67, glngSys, , 4))

    aryValue = Split(mfrmParent.GetPicScale.Tag, ";")
    
    strTmp = GetCurveDateTime(mintNowCol + 1, CDate(aryValue(0)), lngHourBegin)
    If strTmp <> "" Then
        strTmp = "ʱ�䣺" & Format(Split(strTmp, ",")(0), "yyyy-MM-dd") & " " & Format(Split(strTmp, ",")(0), "HHʱmm��") & "��" & Format(Split(strTmp, ",")(1), "HHʱmm��")
    End If
    lblTime.Caption = strTmp
    
    vsf.Rows = 2
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    vsf.RowData(1) = 0
    With mfrmParent.GetmshScale
        
        aryValue = Split(.TextMatrix(GraphDataRow.��������, mintNowCol + .FixedCols), ";")
        aryNote = Split(.TextMatrix(GraphDataRow.δ��˵��, mintNowCol + .FixedCols), ";")
        
        For intCount = 0 To .FixedCols - 1
            
            If Val(vsf.RowData(vsf.Rows - 1)) <> 0 Then vsf.Rows = vsf.Rows + 1
            
            vsf.RowData(vsf.Rows - 1) = .ColData(intCount)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��Ŀ) = .TextMatrix(0, intCount)
            
            aryPara = Split(mfrmParent.GetpicLine(intCount).Tag, ";")
            'RS!���ֵ & ";" & RS!��Сֵ & ";" & RS!��λֵ & ";" & RS!�����
            
            vsf.TextMatrix(vsf.Rows - 1, mCol.��Сֵ) = aryPara(1)
            vsf.TextMatrix(vsf.Rows - 1, mCol.���ֵ) = aryPara(0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��λֵ) = aryPara(2)
            vsf.TextMatrix(vsf.Rows - 1, mCol.�����) = aryPara(3)
            
            If Trim(aryValue(intCount + 1)) <> "" Then
                For intCol = 0 To UBound(Split(aryValue(intCount + 1), ","))
                    
                    dblValue = aryPara(0) - (Val(Split(aryValue(intCount + 1), ",")(intCol)) / .ROWHEIGHT(1) - aryPara(3) + 1) * aryPara(2)
                    If InStr(.TextMatrix(0, intCount), "����") > 0 Then
                        dblValue = Format(dblValue, "0.00")
                    Else
                        dblValue = Format(dblValue, "0")
                    End If
                    
                    If intCol = 0 Then
                        If aryNote(intCount + 1) = "����" And CStr(Val(dblValue)) = "0" Then
                            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = "����"
                        Else
                            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = dblValue
                        End If
                    Else
                        vsf.TextMatrix(vsf.Rows - 1, mCol.���) = dblValue
                    End If
                Next
                
            End If
            
            If aryNote(intCount + 1) = "����" And CStr(Val(dblValue)) = "0" Then
                vsf.TextMatrix(vsf.Rows - 1, mCol.δ��˵��) = ""
            Else
                vsf.TextMatrix(vsf.Rows - 1, mCol.δ��˵��) = aryNote(intCount + 1)
            End If
            
            Select Case Val(vsf.RowData(vsf.Rows - 1))
            Case 1, 2, 3
                vsf.TextMatrix(vsf.Rows - 1, mCol.��λ) = Split(.TextMatrix(GraphDataRow.��λ��־, mintNowCol + .FixedCols), ";")(intCount + 1)
            End Select
        Next
        
        txt(0).Text = .TextMatrix(GraphDataRow.�ϱ�˵��, mintNowCol + .FixedCols)
        txt(1).Text = .TextMatrix(GraphDataRow.�±�˵��, mintNowCol + .FixedCols)
'        chk(0).Value = Val(.TextMatrix(GraphDataRow.�Ͽ���־, mintNowCol + .FixedCols))
    End With
    
   
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.����
    
        Call vsf.Body.AutoSize(mCol.����, mCol.����)
        
    End Select
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    On Error Resume Next

    vsf.ComboList(mCol.��λ) = ""
    vsf.EditMode(mCol.��λ) = 0
    vsf.ComboList(mCol.���) = ""
    vsf.EditMode(mCol.���) = 0
    vsf.ComboList(mCol.��λ) = ""
    
    Select Case Val(vsf.RowData(NewRow))
    Case -1
        
    Case 0
        
        vsf.ComboList(mCol.��λ) = "|���|���|����|�ܲ�|δ��"
    
    Case 1
        vsf.ComboList(mCol.��λ) = "����|Ҹ��|����"
        vsf.EditMode(mCol.��λ) = 1
        
        vsf.ComboList(mCol.���) = ""
        vsf.EditMode(mCol.���) = 1
    Case 2
        vsf.ComboList(mCol.��λ) = " |����"
        vsf.EditMode(mCol.��λ) = 1
        
        If mint����Ӧ�� = 2 Then
            vsf.ComboList(mCol.���) = ""
            vsf.EditMode(mCol.���) = 1
        End If
    Case 3
        vsf.ComboList(mCol.��λ) = "��������|������"
        vsf.EditMode(mCol.��λ) = 1
    
    End Select
    
'    If chk(0).Value = 1 Then
'        vsf.EditMode(mCol.����) = 0
'        vsf.EditMode(mCol.��λ) = 0
'        vsf.EditMode(mCol.���) = 0
'    End If
    
    Dim strTmp As String
    
    If vsf.TextMatrix(NewRow, mCol.��Сֵ) <> "" Or vsf.TextMatrix(NewRow, mCol.��Сֵ) <> "" Then
        strTmp = "��Χ��" & vsf.TextMatrix(NewRow, mCol.��Сֵ) & "��" & vsf.TextMatrix(NewRow, mCol.���ֵ) & " "
    End If
    
    Select Case Val(vsf.RowData(NewRow))
    Case 1
        strTmp = strTmp & "��Ǳ�ʾ�����µ��¶ȣ���λΪ�����µĲ�λ��"
    Case 2
        strTmp = strTmp & "��Ǳ�ʾ���ʵ�ֵ����������ͬʱ�ż�¼����"
    End Select
    lblPrompt.Caption = strTmp
    
    If Val(vsf.RowData(NewRow)) = 0 Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case mCol.���, mCol.����, mCol.δ��˵��
        vsf.TextMatrix(Row, Col) = ""
    End Select
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim intCount As Integer 'δ��˵�������ݵĸ���,ֻ��һ��ʱ�Զ��������հ������
    Dim intRow As Integer, intRows As Integer
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    Select Case Col
    Case mCol.δ��˵��
        
        strSQL = "Select ����,����,RowNum As ID,1 As ĩ�� From ��������˵��"
        If ShowGrdSelectDialog(Me, vsf, "����,3000,0,0", Me.Name & "\��������˵��", "�������ѡ��һ��δ��¼˵����", strSQL, rs, 4500, 4500, False, 2) Then
            vsf.EditText = zlCommFun.NVL(rs("����").Value)
            vsf.Cell(flexcpData, Row, Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(Row, Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, mCol.����) = ""
            vsf.TextMatrix(vsf.Row, mCol.���) = ""
            vsf.TextMatrix(vsf.Row, mCol.��λ) = ""
            
            '����������ߵ�δ������Ϊ��,ֱ�Ӹ���
            intRows = vsf.Rows - 1
            For intRow = 1 To intRows
                If vsf.TextMatrix(intRow, mCol.δ��˵��) = "" And vsf.TextMatrix(intRow, mCol.����) = "" Then
                    intCount = intCount + 1
                End If
            Next
            'ʣ�µ���Ŀ���������Ƕ�Ϊ�������
            If intCount = intRows - 1 Then
                For intRow = 1 To intRows
                    If vsf.TextMatrix(intRow, mCol.δ��˵��) = "" Then
                        vsf.TextMatrix(intRow, mCol.δ��˵��) = zlCommFun.NVL(rs("����").Value)
                    End If
                Next
            End If
        End If
    End Select
    
End Sub

Private Sub vsf_ChangeEdit()
    Select Case vsf.Col
    Case mCol.����
        If Val(vsf.RowData(vsf.Row)) <> 0 Then
            vsf.TextMatrix(vsf.Row, mCol.����) = vsf.EditText
            Call vsf.Body.AutoSize(mCol.����, mCol.����)
            
            If vsf.TextMatrix(vsf.Row, mCol.����) <> "" Then vsf.TextMatrix(vsf.Row, mCol.δ��˵��) = ""
        End If
    Case mCol.���
        
        vsf.TextMatrix(vsf.Row, mCol.δ��˵��) = ""
        
    Case mCol.δ��˵��
        vsf.TextMatrix(vsf.Row, mCol.δ��˵��) = vsf.EditText
        If vsf.TextMatrix(vsf.Row, mCol.δ��˵��) <> "" Then
            vsf.TextMatrix(vsf.Row, mCol.����) = ""
            vsf.TextMatrix(vsf.Row, mCol.���) = ""
            vsf.TextMatrix(vsf.Row, mCol.��λ) = ""
        End If
        
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        
        If Col = mCol.δ��˵�� Then
            
            vsf.Cell(flexcpData, Row, Col) = vsf.EditText
            vsf.TextMatrix(Row, Col) = vsf.EditText
            
        End If
        
        If Row = vsf.Rows - 1 Then cmdOK.SetFocus
        
    End If
End Sub


Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn And Row = vsf.Rows - 1 Then
        cmdOK.SetFocus
    End If
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(vsf.RowData(Row)) <> 0 Then
            If Col <> mCol.δ��˵�� Then
                If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
            Else
                If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
            End If
        End If
    End If
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(vsf.RowData(Row)) <> 0 Then
'            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Select Case Col
    Case mCol.����
        GoTo CheckPoint
    Case mCol.���
        
        Select Case Val(vsf.RowData(Row))
        Case 1
            GoTo CheckPoint
        Case 2
            GoTo CheckPoint
        End Select
        
    End Select
    
    Exit Sub
    
CheckPoint:
    If Trim(vsf.EditText) <> "" Then
        '���²����
        If vsf.RowData(Row) <> 1 And (vsf.TextMatrix(Row, mCol.��Сֵ) <> "" Or vsf.TextMatrix(Row, mCol.���ֵ) <> "") Then
            

            If Val(vsf.EditText) < Val(vsf.TextMatrix(Row, mCol.��Сֵ)) Or Val(vsf.EditText) > Val(vsf.TextMatrix(Row, mCol.���ֵ)) Then
'                Cancel = True
                ShowSimpleMsg "��" & vsf.TextMatrix(Row, mCol.��Ŀ) & " ���ķ�ΧӦ�ڣ�" & Val(vsf.TextMatrix(Row, mCol.��Сֵ)) & "��" & Val(vsf.TextMatrix(Row, mCol.���ֵ)) & "��֮�䣡"
            End If

            
        End If

    End If
End Sub



