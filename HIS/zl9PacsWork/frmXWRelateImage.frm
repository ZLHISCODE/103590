VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmXWRelateImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ӱ��"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   Icon            =   "frmXWRelateImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdRepair 
      Caption         =   "�޸�����״̬(&R)"
      Height          =   350
      Left            =   7275
      TabIndex        =   20
      Top             =   6600
      Width           =   1845
   End
   Begin VB.Frame frmFilter 
      Caption         =   "��������"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   11655
      Begin VB.CommandButton cmdQuery 
         Caption         =   "�� ѯ(&Q)"
         Height          =   350
         Left            =   10440
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboModality 
         Height          =   300
         ItemData        =   "frmXWRelateImage.frx":038A
         Left            =   960
         List            =   "frmXWRelateImage.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtStudyNo 
         Height          =   300
         Left            =   3240
         TabIndex        =   14
         Top             =   960
         Width           =   1600
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Width           =   1600
      End
      Begin VB.Frame frmTime 
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11415
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
            Left            =   8160
            TabIndex        =   12
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   155779075
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
            Left            =   9960
            TabIndex        =   11
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   155779075
            CurrentDate     =   40833
         End
         Begin VB.OptionButton optDays 
            Caption         =   "                 ��"
            Height          =   180
            Index           =   6
            Left            =   7800
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton optDays 
            Caption         =   "1��"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "3��"
            Height          =   180
            Index           =   2
            Left            =   2040
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "5��"
            Height          =   180
            Index           =   3
            Left            =   3000
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "7��"
            Height          =   180
            Index           =   4
            Left            =   3960
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optDays 
            Caption         =   "����"
            Height          =   180
            Index           =   5
            Left            =   4920
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optDays 
            Caption         =   "2��"
            Height          =   180
            Index           =   1
            Left            =   1080
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
         Top             =   1005
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "����ID"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   1000
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "��  ����"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Top             =   1005
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ ��(&O)"
      Height          =   350
      Left            =   9465
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   350
      Left            =   10665
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwUnMatched 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
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
   Begin MSComctlLib.ListView lvwSeries 
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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


Private mlngOrderID As Long         '������ҽ��ID
Private mblnMatch As Boolean        '��������ȡ��������True--������False--ȡ������
Private mblnOpenDB As Boolean       '�����ڴ򿪵����ݿ����ӣ��رմ���ʱҪ�ر�
Private mstrModality As String      'Ĭ�Ϲ�����Ӱ�����
Private mlngReleationState As Long  '-1�޸�������0�ɹ�������1δ������2����ʧ��

Private mrsUnMatchData As ADODB.Recordset

Private Sub cboModality_Click()
On Error GoTo errHandle
    If mblnMatch = True Then '����ͼ��
        If cboModality.ListIndex < 0 Then Exit Sub
        
        Call subFillUnMatched
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    mlngReleationState = 1
    Unload Me
End Sub

Private Function IsCheckSeries() As Boolean
'��ȡ�����Ƿ�����˹�ѡ
    Dim i As Long
    
    IsCheckSeries = False
    
    For i = 1 To lvwSeries.ListItems.Count
        If lvwSeries.ListItems(i).Checked = True Then
            IsCheckSeries = True
            Exit Function
        End If
    Next i
End Function

Private Sub CmdOK_Click()
On Error GoTo errHandle
    Dim blnOpenDb As Boolean
    
    '������ݿ���δ�򿪳ɹ������˳�����
    If gcnXWDBServer.State <> adStateOpen Then
        MsgBox "PACS���ݿ�������������ӣ��ò��������ܼ�����", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If IsCheckSeries = False Then
        MsgBoxD Me, "�빴ѡ��Ҫ�����������Ϣ��", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If mblnMatch = True Then
        '����Ӱ��
        mlngReleationState = IIf(ReleationImages(mlngOrderID) = True, 0, 2)
        
    ElseIf mblnMatch = False Then
        'ȡ������
        mlngReleationState = IIf(CancelReleation(mlngOrderID) = True, 0, 2)
    End If
    
    
    Unload Me
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ReleationImages(ByVal lngOrderID As Long) As Boolean
'����ͼ��
    Dim lngStudyId As Long
    Dim strSeriesIds As String
    Dim strUnCheckIds As String
    Dim lngSelectState As Long
    Dim lngSourceStudyId As Long
    Dim strSql As String
    
    Dim rsTemp As ADODB.Recordset
    Dim rsOrderInfo As ADODB.Recordset
    Dim strStudyDate As String
    Dim strStudyUID As String
    
    lngSelectState = GetSelectData(lngStudyId, strSeriesIds, strUnCheckIds, strStudyUID)
    
    If lngSelectState = 0 Then
        'û��ѡ���κ����ݣ���ֱ���˳�
        MsgBoxD Me, "û��ѡ����Ҫ�����Ĺ������ݡ�", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    lngSourceStudyId = 0
    
    '��ѯҽ��ID�Ƿ��ж�Ӧ��ͼ�������Ϣ
    strSql = "select f_stu_id from v_oem_study_unmatched where f_stu_no='" & lngOrderID & "'"
    
    Set rsTemp = gcnXWDBServer.Execute(strSql)
    If rsTemp.RecordCount > 0 Then lngSourceStudyId = Val(Nvl(rsTemp!f_stu_id))
    
    
    strStudyDate = Trim(lvwUnMatched.SelectedItem.SubItems(5)) & " " & Trim(lvwUnMatched.SelectedItem.SubItems(6))
    
    If lngSelectState = 2 Then
        '�����������
        ReleationImages = ReleationStudy(lngOrderID, lngStudyId, strStudyDate, strStudyUID, IIf(lngSourceStudyId = 0, True, False))
    Else
        '������������
        
        '�ȶ�û�н���ѡ������н��м����
        If XWUnmatchSeries(lngStudyId, strUnCheckIds) <> 0 Then
            Exit Function
        End If
        
        '��ѡ�е����н��й�������
        ReleationImages = ReleationStudy(lngOrderID, lngStudyId, strStudyDate, strStudyUID, IIf(lngSourceStudyId = 0, True, False))
    End If
    
End Function

Private Function ReleationStudy(ByVal lngOrderID As Long, ByVal lngStudyId As Long, _
    ByVal strStudyDate As String, ByVal strStudyUID As String, Optional ByVal blnUpdateRis As Boolean = True) As Boolean
'--------------------------------------------
'���ܣ� �������
'������ lngOrderID -- ҽ��ID
'       lngStudyId -- PACS�еļ������
'       strStudyDate -- PACS�еļ������
'       strStudyUID -- PACS�еļ��UID
'       blnUpdateRis -- ��ѡ�������Ƿ����RIS����
'���أ�
'--------------------------------------------
    Dim strSql As String
    Dim rsOrderInfo As ADODB.Recordset
    
    ReleationStudy = False
    
    '��ѯ������������ҽ����Ϣ
    strSql = "Select b.����ID,b.�����,b.סԺ��,b.������ as ����,b.����,b.�Ա�,b.����,To_char(b.��������,'yyyymmdd') As ��������, " _
                & " c.Ӣ���� as ƴ����,c.Ӱ�����,c.����,a.������Դ,a.ִ�п���ID,d.���� As ִ�п���,a.����ʱ��,a.��ʼִ��ʱ�� " _
                & " From ����ҽ����¼ a,������Ϣ b,Ӱ�����¼ c,���ű� d  " _
                & " Where a.����Id = b.����ID And a.Id = c.ҽ��ID And a.ִ�п���ID =d.Id  and a.Id = [1]"
    Set rsOrderInfo = zlDatabase.OpenSQLRecord(strSql, "��ѯ�����Ϣ", lngOrderID)
    
    If rsOrderInfo.RecordCount > 0 Then
        '���������洢���̡�P_OEM_MATCHING_RIS��������ͼ��
                
        strSql = "P_OEM_MATCHING_RIS(" & lngStudyId & ",'" & lngOrderID & "','" & rsOrderInfo!����ID & "','" & Nvl(rsOrderInfo!�����, 0) _
                & "','" & Nvl(rsOrderInfo!סԺ��, 0) & "','" & Nvl(rsOrderInfo!����, 0) & "','" & Nvl(rsOrderInfo!����) & "','" _
                & Nvl(rsOrderInfo!�Ա�) & "','" & Nvl(rsOrderInfo!����, 0) & "','" & Nvl(rsOrderInfo!��������) & "','" & Nvl(rsOrderInfo!ƴ����) _
                & "','" & Nvl(rsOrderInfo!Ӱ�����) & "','" & rsOrderInfo!���� & "'," & Nvl(rsOrderInfo!������Դ, 3) & "," & Nvl(rsOrderInfo!ִ�п���ID) _
                & ",'" & Nvl(rsOrderInfo!ִ�п���) & "','','')"
                
        gcnXWDBServer.Execute strSql
        
        If blnUpdateRis = True Then
            '���������洢����"b_XINWANGInterface.PacsStatusChange"������ͼ��
            strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") _
                & "b_XINWANGInterface.PacsStatusChange(1," & lngOrderID & ",'" & Nvl(rsOrderInfo!Ӱ�����) _
                & "','" & rsOrderInfo!���� & "',to_date('" & Trim(strStudyDate) _
                & "','yyyy-mm-dd hh24:mi:ss'),null,null,'" & strStudyUID & "')"
                        
            zlDatabase.ExecuteProcedure strSql, "����ͼ��"
        End If
    End If
        
    ReleationStudy = True
End Function


Private Function GetSelectData(ByRef lngStudyId As Long, ByRef strSeriesIds As String, _
    ByRef strUnCheckIds As String, ByRef strStudyUID As String) As Long
'--------------------------------------------
'���ܣ� ��ȡ���������ѡ�����,�ӽ������νṹ����ȡPACS���ݵ���Ϣ
'������ lngStudyId -- V_OEM_STUDY_UNMATCHED�е�F_STU_ID=�������
'       strSeriesIds --��ѡ�е�����ID��V_OEM_SERIES��F_SER_ID=SERIES�����������
'       strUnCheckIds --δѡ�е�����ID��V_OEM_SERIES��F_SER_ID=SERIES�����������
'       strStudyUID -- V_OEM_STUDY_UNMATCHED�е�F_STU_UID=���UID
'���أ�0-δѡ��,1-ѡ�񲿷�����,2-ѡ���������
'--------------------------------------------
    Dim i As Long
    Dim lngSelectState As Long
    
    lngSelectState = 0
    
    lngStudyId = 0
    strSeriesIds = ""
    strUnCheckIds = ""
    
    For i = 1 To lvwSeries.ListItems.Count
        If lvwSeries.ListItems(i).Checked Then
            If strSeriesIds <> "" Then strSeriesIds = strSeriesIds & ","
            strSeriesIds = strSeriesIds & Val(Mid(lvwSeries.ListItems(i).Key, 2))
        Else
            If strUnCheckIds <> "" Then strUnCheckIds = strUnCheckIds & ","
            strUnCheckIds = strUnCheckIds & Val(Mid(lvwSeries.ListItems(i).Key, 2))
            
            lngSelectState = 1
        End If
    Next i
    
    '�ж������Ƿ�ȫѡ
    If lngSelectState <> 1 And strSeriesIds <> "" Then lngSelectState = 2
    
    lngStudyId = Val(Mid(lvwUnMatched.SelectedItem.Key, 2))
    strStudyUID = lvwUnMatched.SelectedItem.SubItems(11)
    
    GetSelectData = lngSelectState

End Function

Private Function CancelReleation(ByVal lngOrderID As Long) As Boolean
'ȡ����������
    Dim lngStudyId As Long
    Dim strSeriesIds As String
    Dim lngSelectState As Long
    Dim strUnCheckIds As String
    Dim strStudyUID As String
    
    '�ж��Ƿ�ѡ�����������У����ѡ�����������У���ֱ�����鼶����й�������
    CancelReleation = False
    
    lngSelectState = GetSelectData(lngStudyId, strSeriesIds, strUnCheckIds, strStudyUID)
    
    If lngSelectState = 0 Then
        'û��ѡ���κ����ݣ���ֱ���˳�
        MsgBoxD Me, "û��ѡ����Ҫ�����Ĺ������ݡ�", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    If lngSelectState = 2 Then
        '�����鼶��Ĺ�������
        CancelReleation = IIf(XWUnmatchImage(lngOrderID, lngStudyId) = 0, True, False)
    Else
        '�������м���Ĺ�������
        CancelReleation = IIf(XWUnmatchSeries(lngStudyId, strSeriesIds) = 0, True, False)
    End If
    
End Function



Public Function zlShowMe(frmParent As Form, lngOrderID As Long, blnMatch As Boolean, strModality As String) As Long
''--------------------------------------------
''���ܣ� ��ʾδƥ���ͼ���¼
''������frmParent --�����壻
''      lngOrderID -- ҽ��ID ��
''      blnMatch --��������ȡ��������True--������False--ȡ������
''      strModality -- ��Ҫ����ͼ���Ӱ�����
''���أ�-1�޸�������0�ɹ���1ʧ��
''--------------------------------------------
    On Error GoTo err
    
    mblnMatch = blnMatch
    mlngOrderID = lngOrderID
    mstrModality = strModality
    
    mlngReleationState = 2
    
    '�ж����ݿ��Ƿ��Ѿ����ӣ����û�����ӣ��������
    If gcnXWDBServer.State <> adStateOpen Then
        If XWDBServerOpen = 0 Then
            mblnOpenDB = True
        End If
    End If
    
    
    InitSeriesList
    InitStudyList
    
    frmFilter.Visible = blnMatch
    
    
    If mblnMatch Then
        '����Ӱ��
        optDays(3).value = True
        
        Call subQueryUnmatched
        Call subFillUnMatched
        
        Call FillModality
                
        lvwSeries.Height = 1695
        Me.Caption = "����Ӱ��"
    Else
        'ȡ������
        'Call subFillMatched
        Call subQueryCurStudy(mlngOrderID)
        Call FillStudyData(mrsUnMatchData)
        
        lvwSeries.Height = 3375
        Me.Caption = "����ȡ��"
    End If
    
    Me.Show 1, frmParent
    
    zlShowMe = mlngReleationState
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subQuerySeries(ByVal strStudyId As String)
'��ѯ����Ӧ��������Ϣ
    Dim strSql As String
    Dim rsSeries As ADODB.Recordset
        
    strSql = "select F_SER_ID as ����ID,F_STU_ID as ���ID,F_SER_DATE as ��������, F_SER_TIME as ����ʱ��," & _
            " F_SER_NO as ���к�,  F_SER_CONTEXT as ��������,F_MODALITY as �豸����," & _
            " F_SER_PLACE as ���в�λ,F_COUNT_IMG as ͼ������  from v_oem_series where F_STU_ID=" & strStudyId & " order by ���к� "
    Set rsSeries = gcnXWDBServer.Execute(strSql)
    
    Call FillSeriesData(rsSeries)
End Sub

Private Sub InitSeriesList()
'��ʼ�������б�
    Dim tmpItem As ListItem
    
    With lvwSeries
        .ListItems.Clear
        
        '���δ��ʼ���У�����г�ʼ��
        If .ColumnHeaders.Count <= 0 Then
            With .ColumnHeaders
                .Clear
                .Add , , "����ID", 1000
                .Add , , "���к�", 1000
                .Add , , "��������", 1200
                .Add , , "����ʱ��", 1200
                .Add , , "��������", 1600
                .Add , , "�豸����", 1000
                .Add , , "���в�λ", 1200
                .Add , , "ͼ������", 1000
            End With
        End If
    End With
End Sub

Private Sub FillSeriesData(rsSeries As ADODB.Recordset)
'���������
    Dim tmpItem As ListItem
    
    lvwSeries.ListItems.Clear
    
    If Not rsSeries.EOF Then
        Do While Not rsSeries.EOF
            Set tmpItem = lvwSeries.ListItems.Add(, "_" & rsSeries("����ID"), Nvl(rsSeries("����ID")))
            With tmpItem
                .SubItems(1) = Nvl(rsSeries("���к�"))
                .SubItems(2) = Nvl(rsSeries("��������"))
                .SubItems(3) = Nvl(rsSeries("����ʱ��"))
                .SubItems(4) = Nvl(rsSeries("��������"))
                .SubItems(5) = Nvl(rsSeries("�豸����"))
                .SubItems(6) = Nvl(rsSeries("���в�λ"))
                .SubItems(7) = Nvl(rsSeries("ͼ������"))
                '.Checked = True
            End With
            rsSeries.MoveNext
        Loop
    End If
End Sub

Private Sub InitStudyList()
'��ʼ������б�
    Dim tmpItem As ListItem
    
    With lvwUnMatched
        .ListItems.Clear
        '���δ��ʼ���У�����г�ʼ��
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
                .Add , , "���UID", 1000
            End With
        End If
    End With
End Sub


Private Sub FillStudyData(rsData As ADODB.Recordset)
'���������
    Dim tmpItem As ListItem
    
    lvwUnMatched.ListItems.Clear
    lvwSeries.ListItems.Clear
    
    If Not rsData.EOF Then
        Do While Not rsData.EOF
            Set tmpItem = lvwUnMatched.ListItems.Add(, "_" & rsData("�������"), Nvl(rsData("����")))
            With tmpItem
                .SubItems(1) = Nvl(rsData("�Ա�"))
                .SubItems(2) = Nvl(rsData("��������"))
                .SubItems(3) = Nvl(rsData("����"))
                .SubItems(4) = Nvl(rsData("����ID"))
                .SubItems(5) = Nvl(rsData("�������"))
                .SubItems(6) = Nvl(rsData("���ʱ��"))
                .SubItems(7) = Nvl(rsData("�������"))
                .SubItems(8) = Nvl(rsData("Ӱ�����"))
                .SubItems(9) = Nvl(rsData("�����Ŀ"))
                .SubItems(10) = Nvl(rsData("ͼ������"))
                .SubItems(11) = Nvl(rsData("���UID"))

            End With
            rsData.MoveNext
        Loop
    End If
    
    If lvwUnMatched.ListItems.Count > 0 Then
        lvwUnMatched.ListItems(1).Selected = True
        
        Call lvwUnMatched_Click
    End If
End Sub

Private Sub subQueryCurStudy(ByVal lngOrderID As Long)
'��ѯ��ǰ�����Ϣ
    Dim strSql As String
    Dim tmpItem As ListItem
    
    
        
    strSql = "select F_PAT_NAME as ����,F_PAT_NO as ����ID,F_SEX as �Ա�,F_STU_BIRTH as ��������,F_STU_ID as �������, " _
            & "F_STU_NO as ҽ��ID,F_STU_UID as ���UID,F_AGE as ����,F_STU_DATE as �������,F_STU_TIME as ���ʱ��, " _
            & " F_STU_SUSPICION as �������,F_MODALITY as Ӱ�����,F_STU_PLACE as �����Ŀ,F_COUNT_IMG as ͼ������ from V_OEM_STUDY_UNMATCHED " _
            & " where F_Stu_No='" & lngOrderID & "'"
    
    Set mrsUnMatchData = gcnXWDBServer.Execute(strSql)
End Sub



Private Sub subQueryUnmatched()
''--------------------------------------------
''���ܣ� ��ѯδƥ���ͼ���¼
''��������
''���أ���
''--------------------------------------------
    Dim strSql As String
    Dim dtNow As Date
    Dim i As Integer
    Dim tmpItem As ListItem
    
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
            & " where F_MATCHED_FLAG = 0 and F_STU_DATE between '" & Format(dtpStart, "yyyy.mm.dd 00:00") & "' and '" & Format(dtpEnd, "yyyy.mm.dd 23:59") & "' order by F_PAT_NAME, F_STU_NO"
    
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
    
    Call FillStudyData(mrsUnMatchData)
    
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
    Dim strStudyUID As String

    '����ҽ��ID��ѯxwpacs���Ѿ����ڵ�ͼ��������
    strSql = "select F_STU_ID as �������, F_STU_NO as ҽ��ID, F_STU_UID as ���UID, F_STU_DATE as �������, F_STU_TIME as ���ʱ�� " _
            & " from V_OEM_STUDY_UNMATCHED " _
            & " where F_STU_NO = '" & mlngOrderID & "'"
    
    Set rsData = gcnXWDBServer.Execute(strSql)
    
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "ͼ�����״̬�޸�ʧ�ܣ���Ӱ���������δƥ�䵽�ü����Ϣ��", vbOKOnly, "��ʾ"
        Exit Sub
    End If

    strStudyUID = Nvl(rsData!���UID)
    '���������洢����"b_XINWANGInterface.PacsStatusChange"������ͼ��
    strSql = IIf(Trim(gstrOracleOwner) <> "", gstrOracleOwner & ".", "") & "b_XINWANGInterface.PacsStatusChange(1," _
        & mlngOrderID & ",null,null,to_date('" & Now & "','YYYY.MM.DD'),null,null,'" & strStudyUID & "')"
    zlDatabase.ExecuteProcedure strSql, "����ͼ��"
    
    mlngReleationState = -1
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Command1_Click()
On Error GoTo errHandle
    '�ǻس������ѯ
    Call subQueryUnmatched
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpEnd_Change()
On Error GoTo errHandle
    If dtpStart.value > dtpEnd.value Then
        dtpEnd.value = dtpStart.value
    End If
    
    Call optDays_Click(6)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpEnd_GotFocus()
    optDays(6).value = True
End Sub

Private Sub dtpStart_Change()
On Error GoTo errHandle
    If dtpStart.value > dtpEnd.value Then
        dtpStart.value = dtpEnd.value
    End If
    Call optDays_Click(6)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
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

Private Sub lvwUnMatched_Click()
'��ѯ����Ӧ��������Ϣ
On Error GoTo errHandle
    Dim strStudyId As String
    
    If lvwUnMatched.SelectedItem Is Nothing Then Exit Sub
    
    strStudyId = Mid(lvwUnMatched.SelectedItem.Key, 2)
    
    Call subQuerySeries(strStudyId)

Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optDays_Click(Index As Integer)
On Error GoTo errHandle
    Call subQueryUnmatched
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    
    '�ǻس������ѯ
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtStudyNo_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
    If KeyAscii <> 13 Then Exit Sub
    
    '�ǻس������ѯ
    Call subFillUnMatched
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
