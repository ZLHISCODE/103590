VERSION 5.00
Begin VB.Form frmSelectReceiver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ռ���ѡ��"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmSelectReceiver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra 
      Height          =   75
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   1433
      Width           =   8145
   End
   Begin VB.CommandButton cmdFind 
      Height          =   315
      Left            =   5700
      Picture         =   "frmSelectReceiver.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1980
      Width           =   390
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   720
      TabIndex        =   20
      Top             =   1995
      Width           =   4965
   End
   Begin VB.OptionButton optPick 
      Caption         =   "ָ������(&S)"
      Height          =   195
      Index           =   7
      Left            =   4275
      TabIndex        =   19
      Top             =   1645
      Width           =   1305
   End
   Begin VB.OptionButton optPick 
      Caption         =   "��Ա����(&X)"
      Height          =   195
      Index           =   5
      Left            =   270
      TabIndex        =   6
      Top             =   1101
      Width           =   1590
   End
   Begin VB.OptionButton optPick 
      Caption         =   "������Ա(&N)"
      Height          =   195
      Index           =   4
      Left            =   2205
      TabIndex        =   18
      Top             =   1645
      Width           =   1365
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   330
      TabIndex        =   17
      Top             =   6435
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3675
      TabIndex        =   16
      Top             =   6435
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4990
      TabIndex        =   15
      Top             =   6435
      Width           =   1100
   End
   Begin VB.OptionButton optPick 
      Caption         =   "ָ����Ա(&I)"
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   7
      Top             =   1645
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.Frame fra 
      Height          =   3945
      Index           =   1
      Left            =   165
      TabIndex        =   8
      Top             =   2400
      Width           =   6045
      Begin VB.CommandButton cmdFunc 
         Caption         =   ">>"
         Height          =   350
         Index           =   3
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         ToolTipText     =   "ȫ������"
         Top             =   540
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "&>"
         Height          =   350
         Index           =   2
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         ToolTipText     =   "��������"
         Top             =   915
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "&<"
         Height          =   350
         Index           =   1
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         ToolTipText     =   "�����Ƴ�"
         Top             =   2160
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "<<"
         Height          =   350
         Index           =   0
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         ToolTipText     =   "ȫ���Ƴ�"
         Top             =   2580
         Width           =   540
      End
      Begin VB.ListBox lst 
         Height          =   3480
         Index           =   1
         ItemData        =   "frmSelectReceiver.frx":685E
         Left            =   3450
         List            =   "frmSelectReceiver.frx":6860
         TabIndex        =   12
         Top             =   270
         Width           =   2385
      End
      Begin VB.ListBox lst 
         Height          =   3480
         Index           =   0
         ItemData        =   "frmSelectReceiver.frx":6862
         Left            =   240
         List            =   "frmSelectReceiver.frx":6864
         TabIndex        =   9
         Top             =   300
         Width           =   2385
      End
   End
   Begin VB.OptionButton optPick 
      Caption         =   "��������Ա(&F)"
      Height          =   195
      Index           =   2
      Left            =   4275
      TabIndex        =   5
      Top             =   769
      Width           =   1485
   End
   Begin VB.OptionButton optPick 
      Caption         =   "��������Ա(&D)"
      Height          =   195
      Index           =   1
      Left            =   2205
      TabIndex        =   4
      Top             =   769
      Width           =   1485
   End
   Begin VB.OptionButton optPick 
      Caption         =   "������Ա(&A)"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   769
      Width           =   1365
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   4365
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -240
      TabIndex        =   0
      Top             =   557
      Width           =   8145
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   270
      TabIndex        =   23
      Top             =   2025
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�û�����ϵͳ(&S)"
      Height          =   180
      Left            =   270
      TabIndex        =   1
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmSelectReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean

Private mstr�ռ��� As String           '�ռ��˵�����

Private mrs��Ա As New ADODB.Recordset '������Ա�嵥
Private mrsϵͳ As New ADODB.Recordset '������ϵͳ

Private mrsUser As New ADODB.Recordset
Private mrsTemp As ADODB.Recordset  '����lst(0)�е���Ϣ
Private mlngOptPick As Long  '���ڴ洢��ǰѡ�������һ��optPick����Ҫ��������optPick(3)��optPick(4)

Private Sub cmbSystem_Click()
    Dim strOwner As String
    On Error GoTo ErrH
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")
    
    gstrSQL = "Select A.���� As ���ű��, B.����, D.�û���" & vbNewLine & _
            "From " & strOwner & ".���ű� A, " & strOwner & ".������Ա C, " & strOwner & ".�ϻ���Ա�� D, " & strOwner & ".��Ա�� B" & vbNewLine & _
            "Where A.ID = C.����id And B.ID = C.��Աid And C.��Աid = D.��Աid And C.ȱʡ = 1 And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) " & vbNewLine & _
            "Order By B.����"
    Set mrs��Ա = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Call optPick_Click(0)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strOwner As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If Trim(txt����.Text) = "" Then Exit Sub
    lst(0).Clear
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")
    
    If optPick(3).Value = True Then
    
        gstrSQL = "select DISTINCT B.����,D.�û��� " & _
                  " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                  strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D " & _
                  "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) and C.ȱʡ=1 " & _
                  " And Upper(B.����) Like [1] order by B.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(Trim(txt����.Text)) & "%")
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("����") & "(" & rsTemp("�û���") & ")"
            rsTemp.MoveNext
        Loop
    ElseIf optPick(4).Value = True Then
        gstrSQL = "select DISTINCT B.����,D.�û��� " & _
                      " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                      strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D,V$session S " & _
                      "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) and C.ȱʡ=1 AND D.�û���=S.USERNAME " & _
                      " And Upper(B.����) Like [1] order by B.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(Trim(txt����.Text)) & "%")
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("����") & "(" & rsTemp("�û���") & ")"
            rsTemp.MoveNext
        Loop
    ElseIf optPick(7).Value = True Then
        gstrSQL = "Select Distinct A.����,A.���� From " & strOwner & ".���ű� A Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                  " And Upper(A.����) Like [1] order by A.����,A.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(Trim(txt����.Text)) & "%")
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("����") & "-" & rsTemp("����")
            rsTemp.MoveNext
        Loop
        
    End If
    If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    lst(0).SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Err = 0
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim intPos  As Long
    Dim strTemp As String
    Dim strOwner As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")

'    mstr�û� = ""
'    mstr���� = ""
    mstr�ռ��� = ""
    
    Set mrsUser = Rec.CopyNew(Nothing, True, , Array("�û���", adVarWChar, 30, Empty, "����", adVarWChar, 30, Empty, "�ռ���", adVarWChar, 30, Empty))

    
    If optPick(3).Value = True Or optPick(4).Value = True Then
        
        '�����б��õ���Ա����
        For i = 0 To lst(1).ListCount - 1
            If lst(1).List(i) <> "" Then
                'ȥ�����ߵ�����
                mrsUser.AddNew
                intPos = InStr(lst(1).List(i), "(")
                strTemp = Mid(lst(1).List(i), intPos + 1)
                strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
                mrsUser.Fields("�û���") = strTemp
                '����ǰΪ�û�����
                strTemp = Mid(lst(1).List(i), 1, intPos - 1)
                mstr�ռ��� = mstr�ռ��� & strTemp & ","
                mrsUser.Fields("����") = strTemp
                mrsUser.Fields("�ռ���") = strTemp
            End If
        Next
        If mstr�ռ��� <> "" Then
            mstr�ռ��� = Mid(mstr�ռ���, 1, Len(mstr�ռ���) - 1)
        End If
    ElseIf optPick(5).Value = True Then
        '��Ա����:�Էֺŷָ�
        For i = 0 To lst(1).ListCount - 1
            mstr�ռ��� = mstr�ռ��� & lst(1).List(i) & ";"
        Next
        If mstr�ռ��� <> "" Then
           
            gstrSQL = "Select Distinct B.����, D.�û���" & vbNewLine & _
                    "From " & strOwner & ".��Ա����˵�� E, " & strOwner & ".�ϻ���Ա�� D, " & strOwner & ".��Ա�� B" & vbNewLine & _
                    "Where B.ID = E.��Աid And B.ID = D.��Աid And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) And Instr('" & mstr�ռ��� & "', E.��Ա����) > 0"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("�û���") = rsTemp.Fields("�û���")
                mrsUser.Fields("����") = rsTemp.Fields("����")
                rsTemp.MoveNext
            Loop
            mstr�ռ��� = "[" & Mid(mstr�ռ���, 1, Len(mstr�ռ���) - 1) & "]"
            
        End If
    ElseIf optPick(7).Value = True Then
        For i = 0 To lst(1).ListCount - 1
            mstr�ռ��� = mstr�ռ��� & lst(1).List(i) & ";"
        Next
        If mstr�ռ��� <> "" Then
            
            gstrSQL = "select DISTINCT B.����,D.�û��� " & _
                      " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                      strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D " & _
                      "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) " & _
                      "  And Instr('" & mstr�ռ��� & "', A.����||'-'||A.���� ) > 0" & _
                      " order by B.����"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("�û���") = rsTemp.Fields("�û���")
                mrsUser.Fields("����") = rsTemp.Fields("����")
                rsTemp.MoveNext
            Loop
            mstr�ռ��� = "{" & Mid(mstr�ռ���, 1, Len(mstr�ռ���) - 1) & "}"
            
        End If
    Else
        If optPick(2).Value = True Then
        '�����ݿ��еõ���Ա����
            mstr�ռ��� = "��������Ա"
            mrs��Ա.Filter = "���ű��='" & gstrDeptCode & "'"
        ElseIf optPick(1).Value = True Then
            mstr�ռ��� = "��������Ա"
            If gstrDeptCode = "" Then
                mrs��Ա.Filter = "���ű��='��'"
            Else
                mrs��Ա.Filter = "���ű�� like '" & gstrDeptCode & "%'"
            End If
        Else
            mstr�ռ��� = "������Ա"
            mrs��Ա.Filter = 0
        End If
        Do Until mrs��Ա.EOF
            mrsUser.AddNew
            mrsUser.Fields("�ռ���") = mstr�ռ���
            mrsUser.Fields("�û���") = mrs��Ա("�û���")
            mrsUser.Fields("����") = mrs��Ա("����")
            
            mrs��Ա.MoveNext
        Loop
    End If
        
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdFunc_Click(Index As Integer)
    '����ָ����Ա��ѡ��
    Dim strTemp As String
    Dim i, j, k As Long
    Dim lngRownum As Long '���ڴ洢���ݵ�rownum
    
    Select Case Index
        Case 0 '������ȫ��
            If lst(1).ListCount = 0 Then Exit Sub
            lst(0).Clear
            lst(1).Clear
            mrsTemp.Filter = ""
            Do Until mrsTemp.EOF
                lst(0).AddItem mrsTemp("����")
                mrsTemp!λ�� = 0
                mrsTemp.MoveNext
            Loop
            If lst(0).ListCount > 0 And lst(1).ListIndex < 0 Then lst(0).ListIndex = 0
        Case 1 '������һ��
            i = lst(1).ListIndex
            If i < 0 Then Exit Sub
            '��ȡҪ����lst(0)��λ��
            mrsTemp.Filter = "����='" & lst(1).List(i) & "'"
            
            If mrsTemp.RecordCount = 1 Then
                lngRownum = mrsTemp!rownum
                mrsTemp.Filter = "λ�� = 0"
                If mrsTemp.RecordCount = 0 Then
                    lst(0).AddItem lst(1).List(i)
                Else
                    For j = 1 To mrsTemp.RecordCount
                        If mrsTemp!rownum > lngRownum Then
                            '��߼�һ������
                            For k = 0 To lst(0).ListCount - 1
                                If lst(0).List(k) = mrsTemp!���� Then
                                    Exit For
                                End If
                            Next
                            lst(0).AddItem lst(1).List(i), k
                            Exit For
                        End If
                        '���������ݶ��������ˣ���û���ҵ����ݣ���˵��Ҫ��ӵ����ݵ�rownum�����ģ�Ӧ��ӵ�lst(0)���
                        If j = mrsTemp.RecordCount Then
                            '��߼�һ������
                            lst(0).AddItem lst(1).List(i)
                        End If
                        mrsTemp.MoveNext
                    Next
                End If
                '��λ�ռ����������¼��������λ����Ϊ0
                mrsTemp.Filter = "rownum = " & lngRownum
                mrsTemp!λ�� = 0
            End If
            '�ұ߼�һ������
            lst(1).RemoveItem i
            If i > lst(1).ListCount - 1 Then
                lst(1).ListIndex = lst(1).ListCount - 1
            Else
                lst(1).ListIndex = i
            End If
            lst(0).ListIndex = lst(0).NewIndex
        Case 2 '������һ��
            i = lst(0).ListIndex
            If i < 0 Then Exit Sub
            mrsTemp.Filter = "���� = '" & lst(0).List(i) & "'"
            mrsTemp!λ�� = 1
            mrsTemp.Filter = "λ�� = 0"
            strTemp = lst(0).List(i)
            lst(0).RemoveItem lst(0).ListIndex
            If i > lst(0).ListCount - 1 Then
                lst(0).ListIndex = lst(0).ListCount - 1
            Else
                lst(0).ListIndex = i
            End If
            lst(1).AddItem strTemp
            lst(1).ListIndex = lst(1).NewIndex
        Case 3 '������ȫ��
            If lst(0).ListCount = 0 Then Exit Sub
            lst(0).Clear
            lst(1).Clear
            mrsTemp.Filter = ""
            Do Until mrsTemp.EOF
                lst(1).AddItem mrsTemp("����")
                mrsTemp!λ�� = 1
                mrsTemp.MoveNext
            Loop
            If lst(1).ListIndex < 0 And lst(1).ListCount > 0 Then lst(1).ListIndex = 0
    End Select
End Sub

Private Sub Form_Load()
    '����ǰѡ��Ϊ��Ա����ʱ�����ṩ��������
    If optPick(5).Value = True Then
        cmdFind.Enabled = False
        txt����.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsTemp = Nothing
End Sub

Private Sub lst_DblClick(Index As Integer)
    If Index = 0 Then
        cmdFunc_Click 2
    Else
        cmdFunc_Click 1
    End If
End Sub

Private Sub lst_GotFocus(Index As Integer)
    cmdFunc(2).Default = True
End Sub

Private Sub lst_LostFocus(Index As Integer)
    cmdOK.Default = True
End Sub

Private Sub optPick_Click(Index As Integer)
    If mrs��Ա.State = 0 Then Exit Sub
    Dim strOwner As String
    Dim var�ռ��� As Variant, strTmp As String, i As Integer
    Dim lngOptPick As Long

    Dim blnList As Boolean
    On Error GoTo errHandle
    mrsϵͳ.Filter = "���=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrsϵͳ("������")
    
    blnList = optPick(3).Value Or optPick(4).Value
    fra(1).Enabled = blnList
    lst(0).Enabled = blnList
    lst(1).Enabled = blnList
    cmdFunc(0).Enabled = blnList
    cmdFunc(1).Enabled = blnList
    cmdFunc(2).Enabled = blnList
    cmdFunc(3).Enabled = blnList
    
    cmdFind.Enabled = False
    txt����.Enabled = False
    txt����.Text = ""
    
    '����Ҫ�б�
    lst(0).Clear

    Set mrsTemp = New ADODB.Recordset
    
    If blnList = True Then
        If optPick(3).Value = True Then
            '��������Ա��ѡȡ������λ���ֶ������ж�������Ӧ�÷���lst(0)����lst(1)�У�0��ʾ��lst(0)�У�1��ʾ��lst(1)��
            gstrSQL = "select rownum,���� || '(' || �û��� || ')' ����, 0 λ�� from (select DISTINCT B.����,D.�û��� " & _
                      " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                      strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D " & _
                      "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') " & _
                      "Or B.����ʱ�� Is Null) and C.ȱʡ=1 order by B.����)"
            lngOptPick = 3
        Else
            '��������Ա��ѡȡ������λ���ֶ������ж�������Ӧ�÷���lst(0)����lst(1)�У�0��ʾ��lst(0)�У�1��ʾ��lst(1)��
            gstrSQL = "select rownum,���� || '(' || �û��� || ')' ����, 0 λ�� from (select DISTINCT B.����,D.�û��� " & _
                      " from " & strOwner & ".���ű� A," & strOwner & ".��Ա�� B," & _
                      strOwner & ".������Ա C," & strOwner & ".�ϻ���Ա�� D,V$session S " & _
                      "  where A.ID=C.����ID and B.ID=C.��ԱID and C.��ԱID=D.��ԱID And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') " & _
                      "Or B.����ʱ�� Is Null) and C.ȱʡ=1 AND D.�û���=S.USERNAME order by B.����)"
            lngOptPick = 4
        End If
        Call zlDatabase.OpenRecordset(mrsTemp, gstrSQL, Me.Caption, adOpenStatic, adLockBatchOptimistic)
        
        '��ֻ��ָ����Ա��������Ա֮�������ȥ�Ļ��������lst(1)�б�ֻ�޸�lst(0)�б�
        If mlngOptPick = 3 Or mlngOptPick = 4 Then
            For i = 0 To lst(1).ListCount - 1
                mrsTemp.Filter = "���� = '" & lst(1).List(i) & "'"
                If mrsTemp.RecordCount <> 0 Then
                    mrsTemp!λ�� = 1
                End If
            Next
        Else
            lst(1).Clear
            If Not mrsUser Is Nothing Then
                If mrsUser.State = adStateOpen Then
                    If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
                    Do Until mrsUser.EOF
                        mrsTemp.Filter = "���� = '" & mrsUser.Fields("����") & "(" & mrsUser.Fields("�û���") & ")" & "'"
                        If mrsTemp.RecordCount <> 0 Then
                            mrsTemp!λ�� = 1
                        End If
                        lst(1).AddItem mrsUser.Fields("����") & "(" & mrsUser.Fields("�û���") & ")"
                        mrsUser.MoveNext
                    Loop
                End If
            End If
        End If
        mrsTemp.Filter = "λ�� = 0"
        Do Until mrsTemp.EOF
            lst(0).AddItem mrsTemp("����")
            mrsTemp.MoveNext
        Loop
        
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
        cmdFind.Enabled = True
        txt����.Enabled = True
        mlngOptPick = lngOptPick
    Else
        mlngOptPick = 0
    End If
    
    If optPick(5).Value = True Then
        lst(1).Clear
        fra(1).Enabled = True
        lst(0).Enabled = True
        lst(1).Enabled = True
        cmdFunc(0).Enabled = True
        cmdFunc(1).Enabled = True
        cmdFunc(2).Enabled = True
        cmdFunc(3).Enabled = True
        
        '����Ա������ѡȡ������λ���ֶ������ж�������Ӧ�÷���lst(0)����lst(1)�У�0��ʾ��lst(0)�У�1��ʾ��lst(1)��
        gstrSQL = "select rownum,����,����, 0 λ�� from (Select ����,���� From " & strOwner & ".��Ա���ʷ���)"
        Call zlDatabase.OpenRecordset(mrsTemp, gstrSQL, Me.Caption, adOpenStatic, adLockBatchOptimistic)
        
        If InStr(mstr�ռ���, "]") > 0 And InStr(mstr�ռ���, "[") > 0 Then
            strTmp = Mid(mstr�ռ���, 2, Len(mstr�ռ���) - 2)
            If InStr(strTmp, ";") > 0 Then
                var�ռ��� = Split(strTmp, ";")
                For i = LBound(var�ռ���) To UBound(var�ռ���)
                    mrsTemp.Filter = "���� = '" & var�ռ���(i) & "'"
                    mrsTemp!λ�� = 1
                    lst(1).AddItem var�ռ���(i)
                Next
            Else
                mrsTemp.Filter = "���� = '" & strTmp & "'"
                mrsTemp!λ�� = 1
                lst(1).AddItem strTmp
            End If
        End If
        
        '�����Ա���ʷ��ൽlst(0)��
        mrsTemp.Filter = "λ�� = 0"
        Do Until mrsTemp.EOF
            lst(0).AddItem mrsTemp("����")
            mrsTemp.MoveNext
        Loop
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    End If
    
    If optPick(7).Value = True Then
        lst(1).Clear
        fra(1).Enabled = True
        lst(0).Enabled = True
        lst(1).Enabled = True
        cmdFunc(0).Enabled = True
        cmdFunc(1).Enabled = True
        cmdFunc(2).Enabled = True
        cmdFunc(3).Enabled = True
        cmdFind.Enabled = True
        txt����.Enabled = True
        
        '��ָ��������ѡȡ������λ���ֶ������ж�������Ӧ�÷���lst(0)����lst(1)�У�0��ʾ��lst(0)�У�1��ʾ��lst(1)��
        gstrSQL = "select rownum,���� || '-' || ���� ����, 0 λ�� from (Select Distinct A.����,A.���� From " & strOwner & _
                ".���ű� A Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                " order by A.����,A.����)"
        Call zlDatabase.OpenRecordset(mrsTemp, gstrSQL, Me.Caption, adOpenStatic, adLockBatchOptimistic)

        If InStr(mstr�ռ���, "}") > 0 And InStr(mstr�ռ���, "{") > 0 Then
            strTmp = Mid(mstr�ռ���, 2, Len(mstr�ռ���) - 2)
            If InStr(strTmp, ";") > 0 Then
                var�ռ��� = Split(strTmp, ";")
                For i = LBound(var�ռ���) To UBound(var�ռ���)
                    mrsTemp.Filter = "���� = '" & var�ռ���(i) & "'"
                    mrsTemp!λ�� = 1
                    lst(1).AddItem var�ռ���(i)
                Next
            Else
                mrsTemp.Filter = "���� = '" & strTmp & "'"
                mrsTemp!λ�� = 1
                lst(1).AddItem strTmp
            End If
        End If

        '�������б�lst(0)��
        mrsTemp.Filter = "λ�� = 0"
        Do Until mrsTemp.EOF
            lst(0).AddItem mrsTemp("����")
            mrsTemp.MoveNext
        Loop
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function Get�ռ���(str�ռ��� As String, rsUser As ADODB.Recordset) As Boolean
    
    Dim var�ռ��� As Variant, strTmp As String, i As Integer
    On Error GoTo errHandle
    mblnOK = False
    mstr�ռ��� = str�ռ���
    
    Set mrsUser = rsUser
    '-----------------------------------
    '���ݴ����Ĳ���������ʾ
    lst(1).Clear
    Select Case str�ռ���
        Case "������Ա"
            optPick(0).Value = True
        Case "��������Ա"
            optPick(1).Value = True
        Case "��������Ա"
            optPick(2).Value = True
        Case Else
            If mlngOptPick = 3 Or str�ռ��� = "" Then
                optPick(3).Value = True
            ElseIf mlngOptPick = 4 Then
                optPick(4).Value = True
            End If
            If InStr(str�ռ���, "[") > 0 And InStr(str�ռ���, "]") > 0 Then
                optPick(5).Value = True
            ElseIf InStr(str�ռ���, "{") > 0 And InStr(str�ռ���, "}") > 0 Then
                optPick(7).Value = True
            End If
            
            If optPick(5).Value = True Or optPick(7).Value = True Then
                '��������,ָ������
                lst(1).Clear
                strTmp = Mid(str�ռ���, 2, Len(str�ռ���) - 2)
                If InStr(strTmp, ";") > 0 Then
                    var�ռ��� = Split(strTmp, ";")
                    For i = 0 To UBound(var�ռ���)
                        lst(1).AddItem var�ռ���(i)
                    Next
                Else
                    lst(1).AddItem strTmp
                End If
            Else
                If Not rsUser Is Nothing Then
                    If rsUser.State = adStateOpen Then
                        If rsUser.RecordCount > 0 Then rsUser.MoveFirst
                        Do Until rsUser.EOF
                            lst(1).AddItem rsUser.Fields("����") & "(" & rsUser.Fields("�û���") & ")"
                            rsUser.MoveNext
                        Loop
                    End If
                End If
            End If
            If lst(1).ListCount > 0 Then lst(1).ListIndex = 0
    End Select
    
    '�õ�ϵͳ
    gstrSQL = "select A.���,A.���� ||'��'||A.���||'��' as ����,A.������ from zlsystems A, (select owner from all_tables where " & _
               " table_name in ('���ű�','��Ա��','������Ա','�ϻ���Ա��') " & _
               " group by owner " & _
               " having count(table_name)=4) B " & _
               " Where A.������ = B.owner"
    Call zlDatabase.OpenRecordset(mrsϵͳ, gstrSQL, Me.Caption)
    
    If mrsϵͳ.EOF Then
        MsgBox "�㲻����ѡ���ռ��˵�Ȩ�ޣ�����ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Function
    End If
    cmbSystem.Clear
    Do Until mrsϵͳ.EOF
        cmbSystem.AddItem mrsϵͳ("����")
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrsϵͳ("���")
        mrsϵͳ.MoveNext
    Loop
    
    If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
    If cmbSystem.ListCount = 1 Then cmbSystem.Enabled = False
    
    'ͨ��cmbSystem��ѡ���Ѿ��õ���Ա�嵥
    
    frmSelectReceiver.Show vbModal
    Get�ռ��� = mblnOK
    If mblnOK = True Then
        str�ռ��� = mstr�ռ���
        Set rsUser = mrsUser
    End If
    If mrs��Ա.State = 1 Then mrs��Ա.Close
    Set mrs��Ա = Nothing
    If mrsϵͳ.State = 1 Then mrsϵͳ.Close
    Set mrsϵͳ = Nothing
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub txt����_GotFocus()
    cmdFind.Default = True
End Sub

Private Sub txt����_LostFocus()
    cmdOK.Default = True
End Sub
