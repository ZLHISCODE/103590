VERSION 5.00
Begin VB.Form frmDiagSel2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����������"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   Icon            =   "frmDiagSel2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.OptionButton Opt 
      Caption         =   "���������ʱ�Ӽ������Ŀ¼����ȡ(&D)"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1875
      Width           =   4380
   End
   Begin VB.OptionButton Opt 
      Caption         =   "���������ʱ�Ӽ�������Ŀ¼����ȡ(&E)"
      Height          =   255
      Index           =   0
      Left            =   405
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Value           =   -1  'True
      Width           =   4380
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6105
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6075
      TabIndex        =   6
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6075
      TabIndex        =   5
      Top             =   240
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   1275
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   5655
      Begin VB.TextBox txt2 
         Height          =   300
         Left            =   900
         TabIndex        =   4
         Tag             =   "100"
         Top             =   735
         Width           =   4440
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Left            =   915
         TabIndex        =   2
         Tag             =   "100"
         Top             =   315
         Width           =   4425
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "֤��(&S)"
         Height          =   180
         Left            =   195
         TabIndex        =   3
         Top             =   795
         Width           =   630
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&I)"
         Height          =   180
         Left            =   195
         TabIndex        =   1
         Top             =   375
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmDiagSel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LAWLChar = "';`|,"""

Private mblnCancel As Boolean
Public mstrTxt1 As String
Public mstrTxt2 As String
Public mlngID1 As Long
Public mlngID2 As Long
Public mlngIllnID1 As Long

Private strSQL As String
Private rsTmp As New ADODB.Recordset
Private mblnW As Boolean    '�ǲ�����ҽ���,��Ϊ��ҽ,��Ϊ��ҽ

Private mstrReturn As String '��������ѡ������ʼ������

Private i As Long, j As Long

Public Function ShowDiagSel(frmParent As Object, strReturn As String, ByVal blnW As Boolean) As Boolean
    '��ʾ����
    mblnCancel = False
    mblnW = blnW
    Me.txt1.Text = strReturn
    If mblnW = True Then
        Me.txt2.Enabled = False
        Me.txt2.BackColor = Me.BackColor
        Me.lbl2.Enabled = False
    Else
        Me.txt2.Enabled = True
        Me.txt2.BackColor = RGB(255, 255, 255)
        Me.lbl2.Enabled = True
    End If
    Me.Show 1, frmParent
    If mblnCancel = False Then
        '���ظ�ʽ:  ������������;���ID;����ID;֤����;֤ID------>�˴���֤ID�ɱ��� mlngID2 ����
        strReturn = Replace(mstrTxt1, ";", "��") & ";" & mlngID1 & ";" & mlngIllnID1 & ";" & Replace(mstrTxt2, ";", "��") & ";" & mlngID2
        ShowDiagSel = True
    End If
End Function

Private Function LocalCheck�Ƿ�Ƿ�(txt As Control, ByVal strLawlChar As String) As Boolean
    '����:����ǲ��ǰ���strLawlChar����ַ���,����оͷ���Ϊ�����ͷ��ط�
    On Error GoTo ErrHandle
    Dim strSour As String
    
    If TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
        If TypeOf txt Is ComboBox Then
            If txt.Style <> 0 Then
                '����ComboBoxΪѡ��������ֻ����������
                LocalCheck�Ƿ�Ƿ� = True
                Exit Function
            End If
        End If
        strSour = txt.Text
        If Len(strSour) > 0 Then
            For i = 1 To Len(strLawlChar)
                If InStr(strSour, Mid(strLawlChar, i, 1)) > 0 Then
                    txt.SelStart = InStr(strSour, Mid(strLawlChar, i, 1))
                    txt.SelLength = 1
                    MsgBox "�ı�������зǷ��ַ���", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                    Exit Function
                End If
            Next
            If VarType(txt.Tag) = vbLong Or VarType(txt.Tag) = vbInteger Then
                If zlCommFun.ActualLen(strSour) > txt.Tag And txt.Tag > 0 Then
                    MsgBox "����������ı�������", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                End If
            ElseIf VarType(txt.Tag) = vbString And IsNumeric(txt.Tag) Then
                If zlCommFun.ActualLen(strSour) > CLng(txt.Tag) And CLng(txt.Tag) > 0 Then
                    MsgBox "����������ı�������", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                End If
            End If
        End If
    End If
    Exit Function
ErrHandle:
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State <> adStateOpen Then Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdHelp_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub cmdOK_Click()
    zlCommFun.OpenIme
    mblnCancel = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnW Then
        Me.txt2.Enabled = False
        Me.lbl2.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandle
    If KeyCode = 13 And Shift = 0 Then
        If Not TypeOf ActiveControl Is CommandButton Then
            zlCommFun.PressKey vbKeyTab
        End If
    End If
    Exit Sub
ErrHandle:
    If gcnOracle Is Nothing Then Exit Sub
    If gcnOracle.State <> adStateOpen Then Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Me.Opt(0).Value = IIf(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\���ѡ����", "��������Ŀ¼", "1") = "1", True, False)
    Me.Opt(1).Value = Not Me.Opt(0).Value
    Me.txt1.Text = mstrTxt1
    Me.txt1.SelStart = Len(Me.txt1.Text)
    Me.txt2.Text = mstrTxt2
    Me.txt2.SelStart = Len(Me.txt2.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\���ѡ����", "��������Ŀ¼", IIf(Opt(0).Value = True, "1", "0"))
End Sub

Private Sub txt1_Change()
    mstrTxt1 = txt1.Text
End Sub

Private Sub txt1_GotFocus()
    zlControl.TxtSelAll txt1
    zlCommFun.OpenIme True
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
    Dim blnMatching As Boolean
    Dim CurPoint As POINTAPI
    Dim strWidth As String
    
    If InStr("'~|;,.?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        If gcnOracle Is Nothing Then Exit Sub
        If gcnOracle.State <> adStateOpen Then Exit Sub
        
        If Trim(txt1.Text) <> "" Then
            If Asc(Left(txt1.Text, 1)) < 0 Then
                '            Exit Sub
            End If
        End If
        
        blnMatching = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0") = "0", True, False)
        If mblnW = False Then   '��ҽ
            If Opt(0).Value = True Then '����ҽ��������
                strSQL = "(  UPPER(����) like '" & UCase(txt1.Text) & "%' or " & _
                "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                    "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' )"
                
                strSQL = "select id,����,����,����,���� from ��������Ŀ¼ where ���='B' AND " & strSQL
            Else    '�����Ŀ¼
                strSQL = "( UPPER(a.����) like '" & UCase(txt1.Text) & "%' or " & _
                "  UPPER(b.����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                    "  UPPER(b.����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%')"
                
                strSQL = "select a.id,a.����,b.���� from �������Ŀ¼ a,������ϱ��� b where a.���=2 and a.id=b.���id and " & strSQL
            End If
        Else    '��ҽ
            If Opt(0).Value = True Then '��ICD-10��
                strSQL = "( UPPER(����) like '" & UCase(txt1.Text) & "%' or " & _
                "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                    "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%')   " & _
                    "   "
                
                strSQL = "select id,����,����,����,���� from ��������Ŀ¼ where ���='D' AND " & strSQL
            Else    '�����Ŀ¼
                strSQL = "( UPPER(a.����) like '" & UCase(txt1.Text) & "%' or " & _
                "  UPPER(b.����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                    "  UPPER(b.����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%')"
                
                strSQL = "select a.id,a.����,b.���� from �������Ŀ¼ a,������ϱ��� b  where a.���=1 and a.id=b.���id and " & strSQL
            End If
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��ϼ�¼��ѡ����")
        If rsTmp.RecordCount = 1 Then
            txt1.Text = zlCommFun.Nvl(rsTmp!����)
            If Opt(0).Value = True Then
                '˵���Ǽ�������ID
                mlngIllnID1 = rsTmp!ID
            Else
                '˵�������ID
                mlngID1 = rsTmp!ID
            End If
            '����Ǵ��������ȡ�Ļ���ô��Ҫ�Ӽ�����϶����ж������ܴ��ڵļ�������ID
            If Opt(1).Value = True And mlngID1 > 0 Then
                strSQL = "select * from ������϶��� where ���ID=" & mlngID1
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��ϼ�¼��ѡ����")
                If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    mlngIllnID1 = rsTmp!����id
                End If
            End If
        ElseIf rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            '��λѡ����
            CurPoint.x = (txt1.Left) / Screen.TwipsPerPixelX
            CurPoint.y = (txt1.Top + txt1.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen Frame1.hwnd, CurPoint
            
            '��ʼѡ����
            strWidth = "0;1200;" & IIf(txt1.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26 < 1500, "1500", txt1.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26) & ";1200;800"
            strWidth = frmSelectChild.ShowSelectChild(Me, CurPoint.x * Screen.TwipsPerPixelX, CurPoint.y * Screen.TwipsPerPixelY, txt1.Width, Screen.TwipsPerPixelY * 300, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;;;" Then
                Exit Sub
            End If
            '������صĲ���
            txt1.Text = Split(strWidth, ";")(2)
            If IsNumeric(Split(strWidth, ";")(0)) Then
                If Opt(0).Value = True Then
                    '˵���Ǽ�������ID
                    mlngIllnID1 = CLng(Split(strWidth, ";")(0))
                Else
                    '˵�������ID
                    mlngID1 = CLng(Split(strWidth, ";")(0))
                End If
            End If
            '����Ǵ��������ȡ�Ļ���ô��Ҫ�Ӽ�����϶����ж������ܴ��ڵļ�������ID
            If Opt(1).Value = True And mlngID1 > 0 Then
                strSQL = "select * from ������϶��� where ���ID=" & mlngID1
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��ϼ�¼��ѡ����")
                If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    mlngIllnID1 = rsTmp!����id
                End If
            End If
        Else
            KeyAscii = 0
            Beep
            Beep
            Beep
        End If
    End If
End Sub

Private Sub txt1_LostFocus()
    Dim strTmp As String
    strTmp = txt1.Text
    For i = 1 To Len(LAWLChar)
        strTmp = Replace(strTmp, Mid(LAWLChar, i, 1), "")
    Next
    txt1.Text = strTmp
    zlCommFun.OpenIme
End Sub

Private Sub txt2_Change()
    mstrTxt2 = txt2.Text
End Sub

Private Sub txt2_GotFocus()
    zlControl.TxtSelAll txt2
    zlCommFun.OpenIme True
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
    Dim blnMatching As Boolean
    Dim CurPoint As POINTAPI
    Dim strWidth As String
    Dim objParent As Object
    
    If InStr("'~|;,.?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn Then
        If gcnOracle Is Nothing Then Exit Sub
        If gcnOracle.State <> adStateOpen Then Exit Sub
        blnMatching = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0") = "0", True, False)
        
        If Trim(txt2.Text) <> "" Then
            If Asc(Left(txt2.Text, 1)) < 0 Then
                '            Exit Sub
            End If
        End If
        If mblnW Then Exit Sub '��ҽ�˳�
        
        If Opt(0).Value = True Then '����ҽ��������
            strSQL = "( UPPER(����) like '" & UCase(txt2.Text) & "%' or " & _
            "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt2.Text) & "%' or " & _
                "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt2.Text) & "%'" & _
                "   ) "
            strSQL = "select id,����,����,����,���� from ��������Ŀ¼ where ���='Z' AND " & strSQL
        Else '������ϲο�����ȡ
            If mlngID1 < 1 Then Exit Sub
            strSQL = "( UPPER(֤������) like '" & IIf(blnMatching = True, "%", "") & UCase(txt2.Text) & "%' ) "
            strSQL = _
                "SELECT ֤��ID,֤������" & vbCrLf & _
                "  FROM ������ϲο�" & vbCrLf & _
                " WHERE ���ID=" & mlngID1 & vbCrLf & _
                "   AND " & strSQL
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��ϼ�¼��ѡ����")
        If rsTmp.RecordCount = 1 Then
            txt2.Text = zlCommFun.Nvl(rsTmp!����)
            mlngID2 = rsTmp!ID
        ElseIf rsTmp.RecordCount > 1 Then
            rsTmp.MoveFirst
            CurPoint.x = (txt2.Left) / Screen.TwipsPerPixelX
            CurPoint.y = (txt2.Top + txt2.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen Frame1.hwnd, CurPoint
            
            '��ʼ��ѡ�������п�
            If Opt(0).Value = True Then
                strWidth = "0;1200;" & IIf(txt2.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26 < 1500, "1500", txt2.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26) & ";1200;800"
            Else
                strWidth = "0;" & IIf(txt2.Width - Screen.TwipsPerPixelX * 26 < 1500, "1500", txt2.Width - Screen.TwipsPerPixelX * 26)
            End If
            Set objParent = Nothing
            strWidth = frmSelectChild.ShowSelectChild(objParent, CurPoint.x * Screen.TwipsPerPixelX, CurPoint.y * Screen.TwipsPerPixelY, txt2.Width, Screen.TwipsPerPixelY * 300, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;;;" Or Trim(strWidth) = ";" Then
                Exit Sub
            End If
            txt2.Text = Split(strWidth, ";")(2)
            mlngID2 = CLng(Split(strWidth, ";")(0))
        Else
            Beep
            Beep
            Beep
        End If
    End If
End Sub

Private Sub txt2_LostFocus()
    Dim strTmp As String
    strTmp = txt2.Text
    For i = 1 To Len(LAWLChar)
        strTmp = Replace(strTmp, Mid(LAWLChar, i, 1), "")
    Next
    txt2.Text = strTmp
    zlCommFun.OpenIme
End Sub
