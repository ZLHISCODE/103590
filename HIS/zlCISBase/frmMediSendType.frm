VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMediSendType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҩƷ��ҩ������������"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7560
   Icon            =   "frmMediSendType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7560
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   7
      ToolTipText     =   "��������룬���ƻ���룡"
      Top             =   1620
      Width           =   2265
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�в�ҩ"
      Height          =   210
      Index           =   2
      Left            =   3360
      TabIndex        =   5
      Top             =   1265
      Width           =   1035
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�г�ҩ"
      Height          =   210
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Top             =   1265
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "��ҩ"
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   1265
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4485
      Left            =   120
      TabIndex        =   8
      Tag             =   "1000"
      Top             =   1995
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7911
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   6240
      TabIndex        =   10
      Top             =   6840
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   4920
      TabIndex        =   9
      Top             =   6840
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   6600
      Width           =   7380
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "&S"
      Height          =   300
      Left            =   3270
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1620
      Width           =   285
   End
   Begin VB.ComboBox cbo��ҩ���� 
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Text            =   "cbo��ҩ����"
      Top             =   840
      Width           =   2265
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   885
      Width           =   2265
   End
   Begin VB.TextBox txtInput 
      Height          =   300
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   6
      ToolTipText     =   "��������룬���ƻ���룡"
      Top             =   1620
      Width           =   1995
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   7500
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediSendType.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "����"
      Height          =   180
      Left            =   4560
      TabIndex        =   18
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ҩƷ����"
      Height          =   180
      Left            =   360
      TabIndex        =   17
      Top             =   1275
      Width           =   720
   End
   Begin VB.Label lbl��ҩ���� 
      AutoSize        =   -1  'True
      Caption         =   "��ҩ����"
      Height          =   180
      Left            =   360
      TabIndex        =   15
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "������Ϣ"
      Height          =   180
      Left            =   360
      TabIndex        =   14
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lbl���෽ʽ 
      AutoSize        =   -1  'True
      Caption         =   "���෽ʽ"
      Height          =   180
      Left            =   4200
      TabIndex        =   13
      Top             =   945
      Width           =   720
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmMediSendType.frx":2294
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��������������趨ҩƷ���Զ��巢ҩ���ͣ��ڲ��ŷ�ҩʱͨ����ҩ���Ϳ��ٽ��з�ҩ������"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   885
      TabIndex        =   0
      Top             =   150
      Width           =   4605
   End
End
Attribute VB_Name = "frmMediSendType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrFindStyle As String
Private mrsTemp As ADODB.Recordset
Private mstrFind As String          '��¼�ϴβ�ѯ�����

Private Enum MediType
    ҩƷ���� = 0
    ҩƷƷ��
    ҩƷ���
    ҩƷ����
    ��ҩ;��
End Enum

Private Sub GetSelect(ByVal intType As Integer, ByVal strInput As String, Optional BlnFind As Boolean = False)
    Dim objNode As Node
    Dim rsTemp As ADODB.Recordset
    Dim strID As String
    Dim strSql As String
    
    On Error GoTo errHandle
    If BlnFind = False Then
        tvwClass.Nodes.Clear
        Set objNode = tvwClass.Nodes.Add(, , "Root", "����", 1)
    End If
    Set mrsTemp = Nothing
    Select Case intType
        Case MediType.ҩƷ����
            gstrSql = "Select Level As ��, ID, �ϼ�id, ����, ����, ����, Decode(����, 1, '��ҩ', Decode(����, 2, '�г�ҩ', '�в�ҩ')) As ���� " & _
                " From ���Ʒ���Ŀ¼ " & _
                " Where ����ʱ�� Is Null "
            If strInput <> "" Then
                gstrSql = gstrSql & " And (���� Like [1] Or ���� Like [2] Or ���� Like [2]) "
            End If
            
            If chk����(0).Value = 1 And chk����(1).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And ���� In (1, 2, 3) "
            ElseIf chk����(0).Value = 1 And chk����(1).Value = 1 Then
                gstrSql = gstrSql & " And ���� In (1, 2) "
            ElseIf chk����(0).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And ���� In (1, 3) "
            ElseIf chk����(1).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And ���� In (2, 3) "
            ElseIf chk����(0).Value = 1 Then
                gstrSql = gstrSql & " And ���� = 1 "
            ElseIf chk����(1).Value = 1 Then
                gstrSql = gstrSql & " And ���� = 2 "
            ElseIf chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And ���� = 3 "
            End If
                
            gstrSql = gstrSql & " Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id " & _
                " Order By ���Ʒ���Ŀ¼.����, Level, ���� "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            If chk����(0).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_��ҩ", "��ҩ", 1)
            If chk����(1).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", 1)
            If chk����(2).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", 1)
    
            Do While Not rsTemp.EOF
                If rsTemp!�� = 1 Then
                    If InStr(1, "," & strID & ",", "," & rsTemp!ID & ",") = 0 Then
                        strID = IIf(strID = "", "", strID & ",") & rsTemp!ID
                    End If
                    Set objNode = tvwClass.Nodes.Add("_" & rsTemp!����, 4, "_" & rsTemp!ID, "[" & rsTemp!���� & "]" & rsTemp!����, 1)
'                    objNode.Expanded = True
                Else
                    If InStr(1, "," & strID & ",", "," & rsTemp!�ϼ�ID & ",") > 0 Then
                        If InStr(1, "," & strID & ",", "," & rsTemp!ID & ",") = 0 Then
                            strID = IIf(strID = "", "", strID & ",") & rsTemp!ID
                        End If
                        Set objNode = tvwClass.Nodes.Add("_" & rsTemp!�ϼ�ID, 4, "_" & rsTemp!ID, "[" & rsTemp!���� & "]" & rsTemp!����, 1)
                    Else
                        Set objNode = tvwClass.Nodes.Add("_" & rsTemp!����, 4, "_" & rsTemp!ID, "[" & rsTemp!���� & "]" & rsTemp!����, 1)
                    End If
                End If
                rsTemp.MoveNext
            Loop
        Case MediType.ҩƷƷ��
            gstrSql = "Select Distinct a.Id, a.����, a.����, a.���, Decode(a.���, '5', '��ҩ', Decode(a.���, '6', '�г�ҩ', '�в�ҩ')) As ���� " & _
                " From ������ĿĿ¼ A, ������Ŀ���� B " & _
                " Where a.Id = b.������Ŀid And a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') "
                      
            If strInput <> "" Then
                gstrSql = gstrSql & " And (a.���� Like [1] Or b.���� Like [2] Or b.���� Like [2]) "
            End If
            
            If chk����(0).Value = 1 And chk����(1).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('5', '6', '7') "
            ElseIf chk����(0).Value = 1 And chk����(1).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('5', '6') "
            ElseIf chk����(0).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('5', '7') "
            ElseIf chk����(1).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('6', '7') "
            ElseIf chk����(0).Value = 1 Then
                gstrSql = gstrSql & " And a.��� = '5' "
            ElseIf chk����(1).Value = 1 Then
                gstrSql = gstrSql & " And a.��� = '6' "
            ElseIf chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� = '7' "
            End If
                
            gstrSql = gstrSql & " Order By a.���, a.����, a.Id, a.���� "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            If chk����(0).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_��ҩ", "��ҩ", 1)
            If chk����(1).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", 1)
            If chk����(2).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", 1)
            
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("_" & rsTemp!����, 4, "_" & rsTemp!ID, rsTemp!����, 1)

                rsTemp.MoveNext
            Loop
        Case MediType.ҩƷ���
            gstrSql = "Select Distinct a.Id,  '[' || a.���� || ']' || a.���� || '(' || a.��� || ')' As ����, a.���, Decode(a.���, '5', '��ҩ', Decode(a.���, '6', '�г�ҩ', '�в�ҩ')) As ���� " & _
                " From �շ���ĿĿ¼ A, �շ���Ŀ���� B " & _
                " Where a.Id = b.�շ�ϸĿid And a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') "
            
            If strInput <> "" Then
                gstrSql = gstrSql & " And (a.���� Like [1] Or b.���� Like [2] Or b.���� Like [2]) "
            End If
            
            If chk����(0).Value = 1 And chk����(1).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('5', '6', '7') "
            ElseIf chk����(0).Value = 1 And chk����(1).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('5', '6') "
            ElseIf chk����(0).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('5', '7') "
            ElseIf chk����(1).Value = 1 And chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� In ('6', '7') "
            ElseIf chk����(0).Value = 1 Then
                gstrSql = gstrSql & " And a.��� = '5' "
            ElseIf chk����(1).Value = 1 Then
                gstrSql = gstrSql & " And a.��� = '6' "
            ElseIf chk����(2).Value = 1 Then
                gstrSql = gstrSql & " And a.��� = '7' "
            End If
            
            gstrSql = gstrSql & " Order By a.���, '[' || a.���� || ']' || a.���� || '(' || a.��� || ')', a.Id "
            
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            If chk����(0).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_��ҩ", "��ҩ", 1)
            If chk����(1).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_�г�ҩ", "�г�ҩ", 1)
            If chk����(2).Value = 1 Then Set objNode = tvwClass.Nodes.Add("Root", 4, "_�в�ҩ", "�в�ҩ", 1)
            
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("_" & rsTemp!����, 4, "_" & rsTemp!ID, rsTemp!����, 1)

                rsTemp.MoveNext
            Loop
        Case MediType.ҩƷ����
            gstrSql = "Select ����, ����, ���� From ҩƷ���� "
            
            If strInput <> "" Then
                gstrSql = gstrSql & " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2]) "
            End If
            
            gstrSql = gstrSql & " Order By ���� "
            
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
                            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
            
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("Root", 4, "_" & rsTemp!����, rsTemp!����, 1)

                rsTemp.MoveNext
            Loop
        Case MediType.��ҩ;��
            gstrSql = "Select Distinct a.Id, a.����, a.���� " & _
                " From ������ĿĿ¼ A, ������Ŀ���� B " & _
                " Where a.Id = b.������Ŀid And a.��� = 'E' And a.�������� = '2' And a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') "

            If strInput <> "" Then
                gstrSql = gstrSql & " And (a.���� Like [1] Or b.���� Like [2] Or b.���� Like [2]) "
            End If

            gstrSql = gstrSql & " Order By a.���� "
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetSelect", UCase(strInput) & "%", mstrFindStyle & UCase(strInput) & "%")
            
            If BlnFind = True Then
                Set mrsTemp = rsTemp
                Exit Sub
            End If
            If rsTemp Is Nothing Then Exit Sub
            If rsTemp.RecordCount = 0 Then Exit Sub
    
            Do While Not rsTemp.EOF
                Set objNode = tvwClass.Nodes.Add("Root", 4, "_" & rsTemp!ID, rsTemp!����, 1)

                rsTemp.MoveNext
            Loop
    End Select
    
    tvwClass.Nodes("Root").Selected = True
    tvwClass.Nodes("Root").Expanded = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo��ҩ����_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub cbo����_Click()
    With cbo����
        If .ListIndex <> Val(.Tag) Then
            .Tag = .ListIndex
            
            txtInput.Text = ""
            txtInput.Tag = ""
            
            Call GetSelect(Val(cbo����.Tag), Trim(txtInput.Text))
        End If
    End With
End Sub



Private Sub chk����_Click(Index As Integer)
    If chk����(0).Value = 0 And chk����(1).Value = 0 And chk����(2).Value = 0 Then
        chk����(Index).Value = 1
        Exit Sub
    End If
    
    Call GetSelect(Val(cbo����.Tag), Trim(txtInput.Text))
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    Dim str���� As String
    Dim int���� As Integer
    Dim str���� As String
    Dim str��ҩ���� As String
    Dim lngCount As Long
    
    If chk����(0).Value = 1 And chk����(1).Value = 1 And chk����(2).Value = 1 Then
        str���� = "5,6,7"
    ElseIf chk����(0).Value = 1 And chk����(1).Value = 1 Then
        str���� = "5,6"
    ElseIf chk����(0).Value = 1 And chk����(2).Value = 1 Then
        str���� = "5,7"
    ElseIf chk����(1).Value = 1 And chk����(2).Value = 1 Then
        str���� = "6,7"
    ElseIf chk����(0).Value = 1 Then
        str���� = "5"
    ElseIf chk����(1).Value = 1 Then
        str���� = "6"
    ElseIf chk����(2).Value = 1 Then
        str���� = "7"
    End If
    
    int���� = Val(cbo����.Tag)
    
    For lngCount = 1 To tvwClass.Nodes.Count
        If tvwClass.Nodes(lngCount).Key <> "Root" And _
            tvwClass.Nodes(lngCount).Key <> "_�г�ҩ" And _
            tvwClass.Nodes(lngCount).Key <> "_�в�ҩ" And _
            tvwClass.Nodes(lngCount).Key <> "_��ҩ" And _
            tvwClass.Nodes(lngCount).Checked Then
            str���� = IIf(str���� = "", "", str���� & ",") & Mid(tvwClass.Nodes(lngCount).Key, 2)
        End If
    Next
    
    If Trim(str����) = "" Then
        MsgBox "�����б���ѡ�������ࡣ", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    str��ҩ���� = Trim(cbo��ҩ����.Text)
    
    If str��ҩ���� = "" Then
        If MsgBox("��û��ѡ��ҩ���ͣ��������Ӧ��ҩƷ�ķ�ҩ���ͣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    gstrSql = "Zl_ҩƷ���_��ҩ����("
    'ҩƷ����
    gstrSql = gstrSql & "'" & str���� & "'" & ","
    '���෽ʽ
    gstrSql = gstrSql & int���� & ","
    '��������
    gstrSql = gstrSql & "'" & str���� & "'" & ","
    '��ҩ����
    gstrSql = gstrSql & IIf(str��ҩ���� = "", "Null", "'" & str��ҩ���� & "'")
    gstrSql = gstrSql & ")"
    
    On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSql, "���淢ҩ����")
    
    MsgBox "����ɹ���", vbExclamation, gstrSysName
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Call GetSelect(Val(cbo����.Tag), Trim(txtInput.Text))
End Sub
Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select ���� From ��ҩ���� Order By ����"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "ȡ��ҩ����")
    
    With cbo��ҩ����
        .Clear
        Do While Not rsData.EOF
            cbo��ҩ����.AddItem rsData.Fields(0).Value
            rsData.MoveNext
        Loop
    End With
    
    With cbo����
        .Clear
        .AddItem "0-ҩƷ����"
        .AddItem "1-ҩƷƷ��"
        .AddItem "2-ҩƷ���"
        .AddItem "3-ҩƷ����"
        .AddItem "4-��ҩ;��"
        
        .ListIndex = 0
        .Tag = 0
    End With
    
    mstrFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    
    Call GetSelect(Val(cbo����.Tag), Trim(txtInput.Text))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtFind.Text) <> "" Then
        Call GetSelect(Val(cbo����.Tag), Trim(txtFind.Text), True)
        
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrFind = ""
    Set mrsTemp = Nothing
End Sub

Private Sub tvwClass_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvwClass, Node, Node.Checked
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim objItem As Node
    
    If KeyAscii = vbKeyReturn And Trim(txtFind.Text) <> "" Then
        zlControl.TxtSelAll txtFind
        If mstrFind <> Trim(txtFind.Text) Then    '�Ѿ��������
            mstrFind = Trim(txtFind.Text)
            Call GetSelect(Val(cbo����.Tag), UCase(Trim(txtFind.Text)), True)
            If mrsTemp.RecordCount > 0 Then
                For Each objItem In tvwClass.Nodes
                    If objItem.Key = "_" & mrsTemp!ID Then
                        objItem.Selected = True
                        Exit For
                    End If
                Next
                mrsTemp.MoveNext
            Else
                MsgBox "û���ҵ�����Ҫ�����ݣ�", vbInformation, gstrSysName
                txtFind.SetFocus
                zlControl.TxtSelAll txtFind
            End If
        Else
            If Not mrsTemp.EOF Then
                mrsTemp.MoveNext
                If Not mrsTemp.EOF Then
                    For Each objItem In tvwClass.Nodes
                    If objItem.Key = "_" & mrsTemp!ID Then
                        objItem.Selected = True
                        Exit For
                    End If
                Next
                End If
            ElseIf mrsTemp.EOF Then
                mrsTemp.MoveFirst
                MsgBox "�Ѳ�ѯ�����", vbInformation, gstrSysName
                If Not mrsTemp.EOF Then
                    For Each objItem In tvwClass.Nodes
                    If objItem.Key = "_" & mrsTemp!ID Then
                        objItem.Selected = True
                        Exit For
                    End If
                Next
                End If
            End If
        End If
    End If
End Sub

Private Sub txtInput_GotFocus()
    zlControl.TxtSelAll txtInput
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call GetSelect(Val(cbo����.Tag), Trim(txtInput.Text))
    End If
End Sub


