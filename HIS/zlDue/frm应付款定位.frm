VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmӦ���λ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ���λ����"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frmӦ���λ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ListView lvwDept 
      Height          =   1290
      Left            =   435
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   2275
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1650
      MaxLength       =   100
      TabIndex        =   2
      Top             =   450
      Width           =   2355
   End
   Begin VB.Frame fraTemp 
      Height          =   75
      Left            =   -300
      TabIndex        =   9
      Top             =   1800
      Width           =   5505
   End
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   240
      Left            =   3720
      TabIndex        =   6
      Top             =   1380
      Width           =   255
   End
   Begin VB.OptionButton opt��λ 
      Caption         =   "�����ݺŶ�λ(&N)"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.OptionButton opt��λ 
      Caption         =   "��ҩƷ��Ӧ�̶�λ(&S)"
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3090
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1950
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2355
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ⵥ�ݺ�(&M)"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   1
      Top             =   540
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ��(&U)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   810
      TabIndex        =   4
      Top             =   1440
      Width           =   810
   End
End
Attribute VB_Name = "frmӦ���λ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mstr���ݺ� As String
Dim mstr��Ӧ��ID As String
Dim msngDownX As Single
Dim msngDownY As Single
Private mstrPrivs As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Enabled = True Then
            If StrIsValid(txtEdit(lngIndex).Text, txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                Exit Sub
            End If
            
            Select Case lngIndex
                Case 0
                    mstr��Ӧ��ID = txtEdit(lngIndex).Tag
                Case 1
                    mstr���ݺ� = UCase(Trim(txtEdit(lngIndex).Text))
            End Select
        End If
    Next
    
    If mstr���ݺ� = "" And mstr��Ӧ��ID = "" Then
        MsgBox "�����붨λ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd�ϼ�_Click()
    Dim rs��Ӧ�� As New ADODB.Recordset
    Dim strȨ�� As String
    strȨ�� = " and (ĩ��<>1 or ( ĩ��=1 " & zl_��ȡվ������ & "  and " & Get����Ȩ��(gstrPrivs) & "))"
        
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,ĩ��,����,����,���� " & _
        "   From ��Ӧ�� " & _
        "   Where nvl(����ʱ��,to_date('3000-01-01','yyyy-MM-dd'))=to_date('3000-01-01','yyyy-MM-dd') " & strȨ�� & _
        "   start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
        "   order by level,ID"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rs��Ӧ��, gstrSQL, Me.Caption
    
    txtEdit(0).SetFocus
    If rs��Ӧ��.EOF Then
        rs��Ӧ��.Close
        Exit Sub
    End If
    With frm��Ӧ��ѡ��
        Me.Tag = .SelDept(mstrPrivs)
        If Me.Tag <> "" Then
            txtEdit(0).Tag = Left(Me.Tag, InStr(Me.Tag, ",") - 1)
            txtEdit(0).Text = Mid(Me.Tag, InStr(Me.Tag, ",") + 1)
        End If
    End With
    Unload frm��Ӧ��ѡ��
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function Get��λ����(ByVal strPrivs As String, str���ݺ� As String, str��Ӧ��ID As String) As Boolean
    mstrPrivs = strPrivs
    frmӦ���λ.Show vbModal, frm�嵥����
    
    Get��λ���� = mblnOK
    If mblnOK = True Then
        str���ݺ� = mstr���ݺ�
        str��Ӧ��ID = mstr��Ӧ��ID
    End If
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Private Sub lvwDept_DblClick()
    If lvwDept.HitTest(msngDownX, msngDownY) Is Nothing Then Exit Sub
    txtEdit(0).Tag = Mid(lvwDept.SelectedItem.Key, 2)
    txtEdit(0).Text = lvwDept.SelectedItem.SubItems(1)
    cmdOK.SetFocus
    lvwDept.Visible = False
End Sub

Private Sub lvwDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Not (lvwDept.SelectedItem Is Nothing) Then
        txtEdit(0).Tag = Mid(lvwDept.SelectedItem.Key, 2)
        txtEdit(0).Text = lvwDept.SelectedItem.SubItems(1)
        cmdOK.SetFocus
        lvwDept.Visible = False
    ElseIf KeyCode = 27 Then
        txtEdit(0).SetFocus
        lvwDept.Visible = False
    End If
End Sub

Private Sub lvwDept_LostFocus()
    lvwDept.Visible = False
End Sub

Private Sub lvwDept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    msngDownX = X
    msngDownY = Y
End Sub

Private Sub opt��λ_Click(Index As Integer)
    txtEdit(0).Enabled = opt��λ(0).Value
    lbl(0).Enabled = opt��λ(0).Value
    cmd�ϼ�.Enabled = opt��λ(0).Value
    
    txtEdit(1).Enabled = opt��λ(1).Value
    lbl(1).Enabled = opt��λ(1).Value
    
    txtEdit(Index).SetFocus
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rstTemp As New ADODB.Recordset, strSQL As String, ltmDept As ListItem
    Dim strȨ�� As String, strKey As String
    
    strȨ�� = " and " & Get����Ȩ��(gstrPrivs)
    On Error GoTo errHandle
    If Index = 0 And KeyAscii = 13 Then
        strKey = GetMatchingSting(txtEdit(0).Text, False)
        'by lesfeng 2009-12-2 �����Ż�
        If IsNumeric(txtEdit(0).Text) Then
            strSQL = "" & _
                "   Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
                "   Where ĩ��=1 " & zl_��ȡվ������ & "  " & _
                "        And ���� Like [1] " & strȨ��
        Else
            strSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� Where ĩ��=1 " & zl_��ȡվ������ & " And (���� Like [1] Or ���� Like [1]) " & strȨ��
        End If
        Set rstTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strKey)
        
        If rstTemp.EOF Then
            MsgBox "ָ���Ĺ�Ӧ�̲����ڣ����������롣", vbInformation, Me.Caption
            txtEdit(0).SetFocus
        ElseIf rstTemp.RecordCount > 1 Then
            lvwDept.ListItems.Clear
            While Not rstTemp.EOF
                Set ltmDept = lvwDept.ListItems.Add(, "D" & rstTemp!ID, rstTemp!����)
                ltmDept.ListSubItems.Add , , rstTemp!����
                rstTemp.MoveNext
            Wend
            Set lvwDept.SelectedItem = lvwDept.ListItems(1)
            lvwDept.ColumnHeaders(1).Width = 1000
            lvwDept.ColumnHeaders(2).Width = lvwDept.Width - 1300
            lvwDept.Visible = True
            lvwDept.SetFocus
        Else
            txtEdit(0).Tag = rstTemp!ID
            txtEdit(0).Text = rstTemp!����
            cmdOK.SetFocus
        End If
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Dim intYear  As Integer, strYear As String
    If IsNumeric(txtEdit(Index)) And txtEdit(Index).Text <> "" And Index = 1 Then
        If Len(txtEdit(1).Text) < 8 And Len(txtEdit(1)) > 0 Then
            txtEdit(1).Text = UCase(LTrim(txtEdit(1).Text))
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            txtEdit(1).Text = strYear & String(7 - Len(txtEdit(1).Text), "0") & txtEdit(1).Text
        End If
    End If
    txtEdit(1).Text = UCase(txtEdit(1).Text)
End Sub
