VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXWRelateImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ӱ��"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   Icon            =   "frmXWRelateImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdRepair 
      Caption         =   "�޸�����״̬(&R)"
      Height          =   350
      Left            =   7275
      TabIndex        =   20
      Top             =   6120
      Width           =   1845
   End
   Begin VB.Frame frmFilter 
      Caption         =   "��������"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   11655
      Begin VB.ComboBox cboModality 
         Height          =   300
         ItemData        =   "frmXWRelateImage.frx":038A
         Left            =   960
         List            =   "frmXWRelateImage.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1057
         Width           =   1600
      End
      Begin VB.TextBox txtStudyNo 
         Height          =   300
         Left            =   5280
         TabIndex        =   14
         Top             =   1057
         Width           =   1600
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   9840
         TabIndex        =   13
         Top             =   1057
         Width           =   1600
      End
      Begin VB.Frame frmTime 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11295
         Begin MSComCtl2.DTPicker dtpStart 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   7560
            TabIndex        =   12
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   112328707
            CurrentDate     =   40833
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   9720
            TabIndex        =   11
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   112328707
            CurrentDate     =   40833
         End
         Begin VB.OptionButton optDays 
            Caption         =   "                   ��"
            Height          =   180
            Index           =   6
            Left            =   7200
            TabIndex        =   19
            Top             =   240
            Width           =   2295
         End
         Begin VB.OptionButton optDays 
            Caption         =   "1��"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "3��"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "5��"
            Height          =   180
            Index           =   3
            Left            =   3480
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "7��"
            Height          =   180
            Index           =   4
            Left            =   4800
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "����"
            Height          =   180
            Index           =   5
            Left            =   6000
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optDays 
            Caption         =   "2��"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Ӱ�����"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "����ID��"
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "��  ����"
         Height          =   255
         Left            =   8880
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ ��(&O)"
      Height          =   350
      Left            =   9465
      TabIndex        =   2
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   350
      Left            =   10665
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwUnMatched 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmXWRelateImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pstrStudyDate As String

Private mlngStudyID As Long         '������PACS�ؼ���
Private mlngOrderID As Long         '������ҽ��ID
Private mblnMatch As Boolean        '��������ȡ��������True--������False--ȡ������
Private mblnOpenDB As Boolean       '�����ڴ򿪵����ݿ����ӣ��رմ���ʱҪ�ر�
Private mstrModality As String      'Ĭ�Ϲ�����Ӱ�����

Private mrsUnMatchData As ADODB.Recordset

Private Sub cboModality_Click()
    
    If mblnMatch = True Then '����ͼ��
        If cboModality.ListIndex < 0 Then Exit Sub
        
        Call subFillUnMatched
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mblnMatch And Not lvwUnMatched.SelectedItem Is Nothing Then
        mlngStudyID = Val(Mid(lvwUnMatched.SelectedItem.Key, 2))
        pstrStudyDate = lvwUnMatched.SelectedItem.SubItems(5)
    ElseIf mblnMatch = False And Not lvwUnMatched.SelectedItem Is Nothing Then
        mlngStudyID = lvwUnMatched.SelectedItem.SubItems(1)
        pstrStudyDate = lvwUnMatched.SelectedItem.SubItems(4)
    End If
    Unload Me
End Sub

Public Function zlShowMe(frmParent As Form, lngOrderID As Long, blnMatch As Boolean, strModality As String) As Long
''--------------------------------------------
''���ܣ� ��ʾδƥ���ͼ���¼
''������frmParent --�����壻
''      lngOrderID -- ҽ��ID ��
''      blnMatch --��������ȡ��������True--������False--ȡ������
''      strModality -- ��Ҫ����ͼ���Ӱ�����
''���أ���Ҫƥ���ҽ��ID
''--------------------------------------------
    On Error GoTo err
    
    mblnMatch = blnMatch
    mlngOrderID = lngOrderID
    mstrModality = strModality
    
    '�ж����ݿ��Ƿ��Ѿ����ӣ����û�����ӣ��������
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            mblnOpenDB = True
        End If
    End If
    
    mlngStudyID = 0
    
    If mblnMatch Then
        optDays(3).value = True
        Call subQueryUnmatched
        Call subFillUnMatched
        Call FillModality
    Else
        Call subFillMatched
    End If
    
    Me.Show 1, frmParent
    
    zlShowMe = mlngStudyID
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subQueryUnmatched()
''--------------------------------------------
''���ܣ� ��ѯδƥ���ͼ���¼
''��������
''���أ���
''--------------------------------------------
    Dim strSql As String
    Dim dtNow As Date
    Dim i As Integer
    
    On Error GoTo err
    
    With lvwUnMatched
        If .ColumnHeaders.Count <= 0 Then
            With .ColumnHeaders
                .Clear
                .Add , , "����", 1000
                .Add , , "�Ա�", 600
                .Add , , "��������", 1200
                .Add , , "����", 600
                .Add , , "����ID", 1000
                .Add , , "�������", 1200
                .Add , , "���ʱ��", 1000
                .Add , , "�������", 1000
                .Add , , "Ӱ�����", 1000
                .Add , , "�����Ŀ", 2200
                .Add , , "ͼ������", 800
            End With
            .ListItems.Add , , "Temp"
        End If
    End With
    
    On Error GoTo err
    
    dtNow = zlDatabase.Currentdate
    For i = 0 To 5
        If optDays(i).value = True Then
            Select Case i
                Case 0
                    dtpStart.value = dtNow
                    dtpEnd.value = dtNow
                Case 1
                    dtpStart.value = DateAdd("d", -1, dtNow)
                    dtpEnd.value = dtNow
                Case 2
                    dtpStart.value = DateAdd("d", -2, dtNow)
                    dtpEnd.value = dtNow
                Case 3
                    dtpStart.value = DateAdd("d", -4, dtNow)
                    dtpEnd.value = dtNow
                Case 4
                    dtpStart.value = DateAdd("d", -6, dtNow)
                    dtpEnd.value = dtNow
                Case 5
                    dtpStart.value = DateAdd("d", -14, dtNow)
                    dtpEnd.value = dtNow
            End Select
        End If
    Next i
    
    strSql = "select F_PAT_NAME as ����,F_PAT_NO as ����ID,F_SEX as �Ա�,F_STU_BIRTH as ��������,F_STU_ID as �������, " _
            & "F_STU_NO as ҽ��ID,F_STU_UID as ���UID,F_AGE as ����,F_STU_DATE as �������,F_STU_TIME as ���ʱ��, " _
            & " F_STU_SUSPICION as �������,F_MODALITY as Ӱ�����,F_STU_PLACE as �����Ŀ,F_COUNT_IMG as ͼ������ from V_OEM_STUDY_UNMATCHED " _
            & " where F_MATCHED_FLAG = 0 and F_STU_DATE between '" & Format(dtpStart, "yyyy.mm.dd 00:00") & "' and '" & Format(dtpEnd, "yyyy.mm.dd 23:59") & "'"
    Set mrsUnMatchData = gcnXWDBServer.Execute(strSql)
Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subFillUnMatched()
''--------------------------------------------
''���ܣ� ���δƥ���ͼ���¼
''��������
''���أ���
''--------------------------------------------
    Dim strFilter As String
    Dim tmpItem As ListItem
    
    On Error GoTo err
    
    '���ù�������
    strFilter = ""
    If cboModality.ListIndex >= 0 Then
        strFilter = IIf(strFilter = "", "", strFilter & " and ") & "Ӱ����� = '" & Split(cboModality.Text, "-")(0) & "'"
    End If
    
    If txtName.Text <> "" Then
        strFilter = IIf(strFilter = "", "", strFilter & " and ") & "���� = '" & txtName.Text & "'"
    End If
    
    If txtStudyNo.Text <> "" Then
        strFilter = IIf(strFilter = "", "", strFilter & " and ") & "����ID = '" & txtStudyNo.Text & "'"
    End If
    
    mrsUnMatchData.Filter = strFilter
    
    lvwUnMatched.ListItems.Clear
    
    If Not mrsUnMatchData.EOF Then
        Do While Not mrsUnMatchData.EOF
            Set tmpItem = lvwUnMatched.ListItems.Add(, "_" & mrsUnMatchData("�������"), Nvl(mrsUnMatchData("����")))
            With tmpItem
                .SubItems(1) = Nvl(mrsUnMatchData("�Ա�"))
                .SubItems(2) = Nvl(mrsUnMatchData("��������"))
                .SubItems(3) = Nvl(mrsUnMatchData("����"))
                .SubItems(4) = Nvl(mrsUnMatchData("����ID"))
                .SubItems(5) = Nvl(mrsUnMatchData("�������"))
                .SubItems(6) = Nvl(mrsUnMatchData("���ʱ��"))
                .SubItems(7) = Nvl(mrsUnMatchData("�������"))
                .SubItems(8) = Nvl(mrsUnMatchData("Ӱ�����"))
                .SubItems(9) = Nvl(mrsUnMatchData("�����Ŀ"))
                .SubItems(10) = Nvl(mrsUnMatchData("ͼ������"))

            End With
            mrsUnMatchData.MoveNext
        Loop
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub subFillMatched()
''--------------------------------------------
''���ܣ� �����ƥ���ͼ���¼
''��������
''���أ���
''--------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim tmpItem As ListItem
    
    On Error GoTo err
    
    With lvwUnMatched
        If .ColumnHeaders.Count <= 0 Then
            With .ColumnHeaders
                .Clear
                .Add , , "Ӱ�����", 2000
                .Add , , "����", 1000
                .Add , , "���к�", 4000
                .Add , , "˵��", 4000
                .Add , , "�ɼ�ʱ��", 2000
            End With
            .ListItems.Add , , "Temp"
        End If
        .ListItems.Clear
    End With
    
    strSql = "select F_SER_ID as SERIES����,F_STU_ID as Study����,F_SER_UID as ����UID,F_SER_DATE as ��������,F_SER_TIME as ����ʱ��, " _
                & " F_SER_CONTEXT as ��������,F_MODALITY as Ӱ������,F_STU_NO as ҽ��ID from V_OEM_SERIES where F_STU_NO ='" & mlngOrderID _
                & "' order by F_STU_ID ,F_SER_ID"
    Set rsTemp = gcnXWDBServer.Execute(strSql)
    
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
             Set tmpItem = lvwUnMatched.ListItems.Add(, "_" & rsTemp!SERIES����, rsTemp!Ӱ������)
            With tmpItem
                .SubItems(1) = Nvl(rsTemp("Study����"))
                .SubItems(2) = Nvl(rsTemp("����UID"))
                .SubItems(3) = Nvl(rsTemp("��������"))
                .SubItems(4) = Nvl(rsTemp("��������"), date)
            End With
            rsTemp.MoveNext
        Loop
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRepair_Click()
On Error GoTo errHandle
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    '����ҽ��ID��ѯxwpacs���Ѿ����ڵ�ͼ��������
    strSql = "select F_STU_ID as �������, F_STU_NO as ҽ��ID, F_STU_UID as ���UID, F_STU_DATE as �������, F_STU_TIME as ���ʱ�� " _
            & " from V_OEM_STUDY_UNMATCHED " _
            & " where F_STU_NO = '" & mlngOrderID & "'"
    
    Set rsData = gcnXWDBServer.Execute(strSql)
    
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "ͼ�����״̬�޸�ʧ�ܣ���Ӱ���������δƥ�䵽�ü����Ϣ��", vbOKOnly, "��ʾ"
        Exit Sub
    End If


    '���������洢����"b_XINWANGInterface.PacsStatusChange"������ͼ��
    strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsStatusChange(1," & mlngOrderID & ",null,null,to_date('" _
                & Now & "','YYYY.MM.DD'),null,null)"
    zlDatabase.ExecuteProcedure strSql, "����ͼ��"
    
    mlngStudyID = -1
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpEnd_Change()
    If dtpStart.value > dtpEnd.value Then
        dtpEnd.value = dtpStart.value
    End If
    Call optDays_Click(6)
End Sub

Private Sub dtpEnd_GotFocus()
    optDays(6).value = True
End Sub

Private Sub dtpStart_Change()
    If dtpStart.value > dtpEnd.value Then
        dtpStart.value = dtpEnd.value
    End If
    Call optDays_Click(6)
End Sub

Private Sub dtpStart_GotFocus()
    optDays(6).value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '������ڹ����д򿪵����ݿ����ӣ����˳�ʱ�ر�����
    If mblnOpenDB = True Then
        Call XWDBServerClose
    End If
End Sub

Private Sub FillModality()
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select ����,���� from Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "Ӱ�������")
    
    cboModality.Clear
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!���� & "-" & rsTemp!����
        If rsTemp!���� = mstrModality Then cboModality.ListIndex = cboModality.ListCount - 1
        rsTemp.MoveNext
    Loop
    
    If cboModality.ListIndex = -1 Then
        If cboModality.ListCount >= 1 Then
            cboModality.ListIndex = 1
        End If
    End If
End Sub

Private Sub optDays_Click(Index As Integer)
    Call subQueryUnmatched
    Call subFillUnMatched
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    '�ǻس������ѯ
    Call subFillUnMatched
End Sub

Private Sub txtStudyNo_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    '�ǻس������ѯ
    Call subFillUnMatched
End Sub
