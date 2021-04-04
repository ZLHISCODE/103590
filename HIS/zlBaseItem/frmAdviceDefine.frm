VERSION 5.00
Begin VB.Form frmAdviceDefine 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�����ݶ���"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmAdviceDefine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdAdd 
      Height          =   270
      Left            =   5175
      Picture         =   "frmAdviceDefine.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "�����ֶ�(ALT+A)"
      Top             =   2325
      Width           =   270
   End
   Begin VB.ComboBox cbo�ֶ� 
      Height          =   300
      Left            =   1020
      TabIndex        =   5
      Top             =   2310
      Width           =   4125
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "���(&K)"
      Height          =   350
      Left            =   2070
      TabIndex        =   7
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4365
      TabIndex        =   9
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3270
      TabIndex        =   8
      Top             =   2865
      Width           =   1100
   End
   Begin VB.TextBox txtAdvice 
      Height          =   1125
      Left            =   1020
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1155
      Width           =   4440
   End
   Begin VB.ComboBox cbo��� 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   825
      Width           =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   -75
      X2              =   5985
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   -165
      X2              =   5895
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   6060
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   -90
      X2              =   5970
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Label lblPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdviceDefine.frx":00D6
      Height          =   645
      Left            =   345
      TabIndex        =   10
      Top             =   75
      Width           =   5040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   180
      Left            =   225
      TabIndex        =   0
      Top             =   885
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ������"
      Height          =   180
      Left            =   225
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֶ���Ŀ"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   2370
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdviceDefine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mblnChange As Boolean
Private mintIndex As Integer
Private mrsField As ADODB.Recordset
Private mrsAdvice As ADODB.Recordset

Public Function ShowMe(frmParent As Object, rsAdvice As ADODB.Recordset) As Boolean
    Set mrsAdvice = Rec.CopyNew(rsAdvice)
    Me.Show 1, frmParent
    If mblnOk Then
        Set rsAdvice = Rec.CopyNew(mrsAdvice)
    End If
    ShowMe = mblnOk
End Function

Private Sub cbo�ֶ�_GotFocus()
    Call zlControl.TxtSelAll(cbo�ֶ�)
End Sub

Private Sub cbo�ֶ�_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    If cbo�ֶ�.Text = "" Then Exit Sub
    txtAdvice.SelText = cbo�ֶ�.Text
    cbo�ֶ�.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim strMsg As String
    
    If Trim(txtAdvice.Text) = "" Then
        MsgBox "û�����ݡ�", vbInformation, gstrSysName
    Else
        strMsg = CheckAdvice(txtAdvice.Text)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
        Else
            MsgBox "ҽ��������д��ȷ��", vbInformation, gstrSysName
        End If
    End If
    txtAdvice.SetFocus
End Sub

Private Function CheckAdvice(ByVal strText As String) As String
'���ܣ����ҽ�������Ƿ���ȷ
'���أ�������Ϣ
'      strPreview=Ԥ��ҽ������Ч��
    Dim intLeft As Integer, intRight As Integer
    Dim strTmp As String, strPar As String
    Dim strMsg As String, i As Long
    Dim objVBA As Object, strEval As String
    Dim objScript As New clsScript
    
    If Trim(strText) = "" Then Exit Function
    If zlCommFun.ActualLen(strText) > txtAdvice.MaxLength Then
        strMsg = "ҽ����������̫����ֻ���� " & txtAdvice.MaxLength & " ���ַ��� " & txtAdvice.MaxLength \ 2 & " �����֡�"
        GoTo EndLine
    End If
        
    '���������
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) = "[" Then
            intLeft = intLeft + 1
        ElseIf Mid(strText, i, 1) = "]" Then
            intRight = intRight + 1
            If intLeft <> intRight Then
                strMsg = """[""��""]""���Ų���ԡ�"
                GoTo EndLine
            End If
        End If
    Next
    If intLeft = 0 And intRight = 0 Then Exit Function
    If intLeft <> intRight Then
        strMsg = """[""��""]""���Ų���ԡ�"
        GoTo EndLine
    End If
    
    '����ֶ�����
    strTmp = strText
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Trim(Left(strTmp, InStr(strTmp, "]") - 1))
                        
        If strPar = "" Then
            strMsg = """[]""����֮��û����д�ֶ�����"
            GoTo EndLine
        End If
        
        For i = 0 To cbo�ֶ�.ListCount - 1
            If cbo�ֶ�.List(i) = "[" & strPar & "]" Then Exit For
        Next
        If i > cbo�ֶ�.ListCount - 1 Then
            strMsg = "ʹ���˲����ڵ�""[" & strPar & "]""�ֶΡ�"
            GoTo EndLine
        End If
    Loop
    
    'ִ�в���
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    If objVBA Is Nothing Then
        strMsg = "Microsoft Script Control δ��ȷ��װ(msscript.ocx)������ִ�м�顣�����°�װ�ͻ��˳���"
        GoTo EndLine
    End If
    Err.Clear: On Error GoTo 0
    objVBA.Language = "VBScript"
    objVBA.AddObject "clsScript", objScript, True
    strEval = Replace(strText, "[", """")
    strEval = Replace(strEval, "]", """")
    On Error Resume Next
    Call objVBA.Eval(strEval)
    If objVBA.Error.Number <> 0 Then
        strMsg = objVBA.Error.Description
        objVBA.Error.Clear
    End If
EndLine:
    CheckAdvice = strMsg
End Function

Private Sub cmdOK_Click()
    If Not UpdateAdvice Then
        txtAdvice.SetFocus: Exit Sub
    End If
    mrsAdvice.Filter = 0
    mblnChange = False
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbAltMask Then
        Call cmdAdd_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    
    mblnOk = False
    
    '��ʼ����ͬ�����õ��ֶ�����
    On Error GoTo ErrHandle
    Set mrsField = New ADODB.Recordset
    mrsField.Fields.Append "���", adVarChar, 4
    mrsField.Fields.Append "�ֶ�", adVarChar, 2000
    mrsField.CursorLocation = adUseClient
    mrsField.LockType = adLockBatchOptimistic
    mrsField.CursorType = adOpenStatic
    mrsField.Open
'    Set mrsField.ActionConnection = Nothing
    mrsField.AddNew: mrsField!��� = "����": mrsField!�ֶ� = "[��ʼʱ��],[ҽ������]" '�������ֶ���Ŀ
    mrsField.AddNew: mrsField!��� = "����": mrsField!�ֶ� = "[������Ŀ],[����],[����],[����Ƶ��],[Ӣ��Ƶ��],[ִ��ʱ��]"
    mrsField.AddNew: mrsField!��� = "4": mrsField!�ֶ� = "[��������],[���],[����]"
    mrsField.AddNew: mrsField!��� = "5": mrsField!�ֶ� = "[������],[ͨ����],[��Ʒ��],[Ӣ����],[���],[����],[����],[����],[����Ƶ��],[Ӣ��Ƶ��],[ִ��ʱ��],[��ҩ;��]"
    mrsField.AddNew: mrsField!��� = "6": mrsField!�ֶ� = "[������],[ͨ����],[��Ʒ��],[Ӣ����],[���],[����],[����],[����],[����Ƶ��],[Ӣ��Ƶ��],[ִ��ʱ��],[��ҩ;��]"
    mrsField.AddNew: mrsField!��� = "8": mrsField!�ֶ� = "[����],[�䷽���],[����Ƶ��],[Ӣ��Ƶ��],[ִ��ʱ��],[�÷�],[�巨]"
    mrsField.AddNew: mrsField!��� = "C": mrsField!�ֶ� = "[������Ŀ],[����걾],[�ɼ�����]"
    mrsField.AddNew: mrsField!��� = "D": mrsField!�ֶ� = "[�����Ŀ],[��鲿λ]"
    mrsField.AddNew: mrsField!��� = "F": mrsField!�ֶ� = "[����ʱ��],[��Ҫ����],[��������],[������]"
    mrsField.AddNew: mrsField!��� = "K": mrsField!�ֶ� = "[��Ѫʱ��],[��Ѫ��Ŀ],[��Ѫ;��],[Ѫ��],[RH],[ִ�з���]"
    mrsField.UpdateBatch
    
    '����������𣬲�����:
    '7-�в�ҩ:���ܵ�����ҽ��
    '9-����:�ǵ���������Ŀ
    'G-����:���ܵ�����ҽ��
    gstrSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('7','9','G') Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do While Not rsTmp.EOF
        cbo���.AddItem rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Loop
    cbo���.ListIndex = 0
    
    mintIndex = cbo���.ListIndex
    mblnChange = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function UpdateAdvice() As Boolean
    Dim strMsg As String
    
    strMsg = CheckAdvice(txtAdvice.Text)
    If strMsg <> "" Then
        Call zlControl.CboSetIndex(cbo���.hwnd, mintIndex)
        MsgBox strMsg, vbInformation, gstrSysName
        txtAdvice.SetFocus: Exit Function
    End If
    mrsAdvice.Filter = "�������='" & Left(cbo���.List(mintIndex), 1) & "'"
    If mrsAdvice.EOF Then
        If Trim(txtAdvice.Text) <> "" Then 'ԭ��û���ݵ������
            mrsAdvice.AddNew
            mrsAdvice!������� = Left(cbo���.List(mintIndex), 1)
            mrsAdvice!ҽ������ = txtAdvice.Text
            mrsAdvice.Update
            mblnChange = True
        End If
    Else
        If Trim(txtAdvice.Text) = "" Then 'ԭ�������ݵ������δ����
            Call zlControl.CboSetIndex(cbo���.hwnd, mintIndex)
            MsgBox "��ǰ����ҽ������û�����á�", vbInformation, gstrSysName
            txtAdvice.SetFocus: Exit Function
        ElseIf mrsAdvice!ҽ������ <> txtAdvice.Text Then
            mrsAdvice!ҽ������ = txtAdvice.Text
            mrsAdvice.Update
            mblnChange = True
        End If
    End If
    txtAdvice.Tag = ""
    UpdateAdvice = True
End Function

Private Sub cbo���_Click()
    Dim arrField As Variant, i As Long
    
    '1.��鲢���µ�ǰ����ҽ������
    '------------------------------
    If Visible And txtAdvice.Tag = "1" Then
        If Not UpdateAdvice Then Exit Sub
    End If
    '2.��ʾ���л���������ҽ������
    '------------------------------
    mintIndex = cbo���.ListIndex
    
    '��ʾ�����ֶ��б�
    cbo�ֶ�.Clear
    
    mrsField.Filter = "���='����'"
    Do While Not mrsField.EOF
        arrField = Split(mrsField!�ֶ�, ",")
        For i = 0 To UBound(arrField)
            cbo�ֶ�.AddItem arrField(i)
        Next
        mrsField.MoveNext
    Loop
    
    mrsField.Filter = "���='" & Left(cbo���.Text, 1) & "'"
    If mrsField.EOF Then
        mrsField.Filter = "���='����'"
    End If
    arrField = Split(mrsField!�ֶ�, ",")
    For i = 0 To UBound(arrField)
        cbo�ֶ�.AddItem arrField(i)
    Next
    
    lblPrompt.Caption = "��ѡ������Ŀ����ֶ��ʹ����VBScript���ݵı��ʽ��ҽ�����ݽ�����ϣ��ֶ�����ʹ�÷�����""[]""�����ʾ�������ֶ���ȡֵ��Ϊ�ַ�����"
    '��ʾ��ǰ���õ�ҽ������
    mrsAdvice.Filter = "�������='" & Left(cbo���.Text, 1) & "'"
    If Not mrsAdvice.EOF Then
        txtAdvice.Text = mrsAdvice!ҽ������
        If mrsAdvice!������� = "D" Then
            lblPrompt.Caption = lblPrompt.Caption & "���ڲ������[��鲿λ]ָ""�걾+����""��"
        End If
    Else
        txtAdvice.Text = ""
    End If
    txtAdvice.Tag = ""
           
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Or txtAdvice.Tag = "1" Then
        If MsgBox("����˳����ᶪʧ�����ı�����ݣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub

Private Sub txtAdvice_Change()
    txtAdvice.Tag = "1"
End Sub
