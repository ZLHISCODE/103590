VERSION 5.00
Begin VB.Form frmOPSSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ����"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmOPSSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5655
      Begin VB.TextBox txt1 
         Height          =   300
         Left            =   780
         TabIndex        =   2
         Tag             =   "100"
         Top             =   315
         Width           =   4560
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   375
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6045
      TabIndex        =   5
      Top             =   195
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6045
      TabIndex        =   6
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6045
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1410
      Width           =   1100
   End
   Begin VB.OptionButton Opt 
      Caption         =   "(&1)���������ʱ��ICD-9-CM3������������ȡ"
      Height          =   255
      Index           =   0
      Left            =   375
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Value           =   -1  'True
      Width           =   4380
   End
   Begin VB.OptionButton Opt 
      Caption         =   "(&2)���������ʱ��������ĿĿ¼��������Ŀ����ȡ"
      Height          =   255
      Index           =   1
      Left            =   375
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1515
      Width           =   4380
   End
End
Attribute VB_Name = "frmOPSSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LAWLChar = "';`|,"""

Private mblnCancel As Boolean
Private mstrTxt1 As String  '��������
Private mlngID1 As Long     '����ID
Private mlngID2 As Long     '��ĿID
Dim i As Long, j As Long

Private strSQL As String
Private rsTmp As New ADODB.Recordset

Public Function ShowSel(frmParent As Object, strReturn As String) As Boolean
    '��ʾ����
    Dim strTmp As String
    Dim strTmp1 As String
    Dim i As Long
    
    mblnCancel = False
    '������Ĳ������зֽ��Ա�õ���ǰ�����ò����õ���ѡ������,��������ʹ���˱���ѡ��������ǰ�Ĳ���ȥ����
    If Trim(strReturn) <> "" Then
        i = InStr(strReturn, ";")
        If i > 0 Then
            '�ҵ�����
            mstrTxt1 = Left(strReturn, i - 1)
            strTmp = Mid(strReturn, i + 1)
            i = InStr(strTmp, ";")
            If i > 0 Then
                '�ҵ�����ID
                strTmp1 = Left(strTmp, i - 1)
                strTmp = Mid(strTmp, i + 1)
                If IsNumeric(strTmp1) Then
                    mlngID1 = CLng(strTmp1)
                Else
                    mlngID1 = 0
                End If
                '��ĿID
                If IsNumeric(strTmp) Then
                    mlngID2 = CLng(strTmp)
                Else
                    mlngID2 = 0
                End If
            Else
                mlngID1 = 0
                mlngID2 = 0
            End If
        Else
            mstrTxt1 = ""
            mlngID1 = 0
            mlngID2 = 0
        End If
    Else
        mstrTxt1 = ""
        mlngID1 = 0
        mlngID2 = 0
    End If
    
    
    Me.Show 1, frmParent
    If mblnCancel = False Then
        '���ظ�ʽ:  ������������;����ID;��ĿID
        strReturn = Replace(mstrTxt1, ";", "��") & ";" & mlngID1 & ";" & mlngID2
        ShowSel = True
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

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
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

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub cmdHelp_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub cmdOK_Click()
    zlCommFun.OpenIme
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Opt(0).Value = IIf(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\����ѡ����", "ICD-9-CM3��������", "1") = "1", True, False)
    Me.Opt(1).Value = Not Me.Opt(0).Value
    Me.txt1.Text = mstrTxt1
    Me.txt1.SelStart = Len(Me.txt1.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\����ѡ����", "ICD-9-CM3��������", IIf(Opt(0).Value = True, "1", "0"))
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
    Dim strWidth As String
    Dim blnMatching As Boolean
    Dim CurPoint As POINTAPI
    
    If InStr("'~|;,.?", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Trim(txt1.Text) <> "" Then
            If Asc(Left(txt1.Text, 1)) < 0 Then
                Exit Sub
            End If
        End If
        If gcnOracle Is Nothing Then Exit Sub
        If gcnOracle.State <> adStateOpen Then Exit Sub
        
        blnMatching = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0") = "0", True, False)
        If Opt(0).Value = True Then '�Ӽ�������
            strSQL = "( UPPER(����) like '" & UCase(txt1.Text) & "%' or " & _
            "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' or " & _
                "  UPPER(����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%' ) "
            
            strSQL = "select id,����,����,����,���� from ��������Ŀ¼ where ���='S' AND    " & strSQL
        Else    '��������ĿĿ¼
            strSQL = "( UPPER(a.����) like '" & UCase(txt1.Text) & "%' or " & _
            "  UPPER(a.����) like '" & IIf(blnMatching = True, "%", "") & UCase(txt1.Text) & "%')"
            
            strSQL = "SELECT a.id,a.����,a.����  FROM ������ĿĿ¼ a,������Ŀ���� b WHERE a.id=b.������Ŀid AND a.���='F' AND (A.����ʱ�� = to_date('3000-01-01','yyyy-mm-dd') OR A.����ʱ�� IS NULL) AND " & strSQL
        End If
        
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "������Ҫ��¼��ѡ����")
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            '��λѡ����
            CurPoint.x = (txt1.Left) / Screen.TwipsPerPixelX
            CurPoint.y = (txt1.Top + txt1.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen Frame1.hwnd, CurPoint
            If Opt(0).Value = True Then '�Ӽ�������
                '��ʼѡ����
                strWidth = "0;1200;" & IIf(txt1.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26 < 1500, "1500", txt1.Width - 1200 - 1200 - 800 - Screen.TwipsPerPixelX * 26) & ";1200;800"
            Else
                '��ʼѡ����
                strWidth = "0;1200;" & IIf(txt1.Width - 1200 - Screen.TwipsPerPixelX * 26 < 1500, "1500", txt1.Width - 1200 - Screen.TwipsPerPixelX * 26)
            End If
            strWidth = frmSelectChild.ShowSelectChild(Me, CurPoint.x * Screen.TwipsPerPixelX, CurPoint.y * Screen.TwipsPerPixelY, txt1.Width, Screen.TwipsPerPixelY * 300, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;;;" Or Trim(strWidth) = ";;" Then
                Exit Sub
            End If
            '������صĲ���
            txt1.Text = Split(strWidth, ";")(2)
            If IsNumeric(Split(strWidth, ";")(0)) Then
                If Opt(0).Value = True Then '�Ӽ�������
                    mlngID1 = CLng(Trim(Split(strWidth, ";")(0)))
                Else
                    mlngID2 = CLng(Trim(Split(strWidth, ";")(0)))
                End If
            End If
        ElseIf rsTmp.RecordCount = 1 Then
            txt1.Text = zlCommFun.Nvl(rsTmp!����)
            If Opt(0).Value = True Then '�Ӽ�������
                mlngID1 = zlCommFun.Nvl(rsTmp!ID, 0)
            Else
                mlngID2 = zlCommFun.Nvl(rsTmp!ID, 0)
            End If
        End If
    Else
        If InStr(LAWLChar, Chr(KeyAscii)) > 0 Then
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

Private Sub txt1_Change()
    mstrTxt1 = txt1.Text
End Sub
