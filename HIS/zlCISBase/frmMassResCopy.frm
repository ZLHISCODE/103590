VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmMassResCopy 
   Caption         =   "�ʿ�Ʒ����"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11160
   Icon            =   "frmMassResCopy.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11160
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   6030
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   8070
      _Version        =   589884
      _ExtentX        =   14235
      _ExtentY        =   10636
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   5355
      Left            =   8190
      ScaleHeight     =   5355
      ScaleWidth      =   2925
      TabIndex        =   1
      Top             =   15
      Width           =   2925
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   3285
         Width           =   2760
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   2400
         Width           =   2760
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   1020
         Width           =   2760
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ȫ��(&C)"
         Height          =   350
         Index           =   1
         Left            =   1485
         TabIndex        =   13
         Top             =   1905
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Index           =   0
         Left            =   375
         TabIndex        =   12
         Top             =   1905
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ر�(&X)"
         Height          =   350
         Left            =   1485
         TabIndex        =   10
         Top             =   4305
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ�ϸ���(&O)"
         Height          =   350
         Left            =   150
         TabIndex        =   9
         Top             =   4305
         Width           =   1320
      End
      Begin VB.OptionButton optValue 
         Caption         =   "��ԭ����������ֵ(&2)"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   8
         Top             =   3975
         Value           =   -1  'True
         Width           =   2310
      End
      Begin VB.OptionButton optValue 
         Caption         =   "��ԭ����Ԥ�����ֵ(&1)"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   7
         Top             =   3720
         Width           =   2310
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "��ԭ����˳��һ��(&D)"
         Height          =   350
         Left            =   375
         TabIndex        =   5
         Top             =   2775
         Width           =   2220
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "����ѡ����ʿ�Ʒ(&S)"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   4
         Top             =   1650
         Width           =   2070
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "����ǰ�����ʿ�Ʒ(&U)"
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   3
         Top             =   1395
         Value           =   1  'Checked
         Width           =   2070
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "������ʹ�����ڷ�Χ:"
         Height          =   180
         Left            =   165
         TabIndex        =   17
         Top             =   2565
         Width           =   1710
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         Caption         =   "�б���ʾ��Χ��ѡ��:"
         Height          =   180
         Left            =   165
         TabIndex        =   11
         Top             =   1140
         Width           =   1710
      End
      Begin VB.Label lblValues 
         AutoSize        =   -1  'True
         Caption         =   "��ֵ�����ŵ�Ԥ�����ֵ:"
         Height          =   180
         Left            =   150
         TabIndex        =   6
         Top             =   3450
         Width           =   2070
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "    ��Ҫ���������������ʿ�Ʒʱʹ�ñ����ܣ��ɿ��ٽ��������ſ���Ʒ���̳�ԭ���ŵļ����Ŀ�Ϳ���ֵ��"
         Height          =   720
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   2700
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   165
         Picture         =   "frmMassResCopy.frx":058A
         Top             =   60
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   270
      Top             =   5385
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
            Picture         =   "frmMassResCopy.frx":0B14
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMassResCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ѡ�� = 0: ����: ID: ����: ԭ����: ԭ��ʼ����: ԭ��������: ������: �¿�ʼ����: �½�������
End Enum

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Public Function zlRefList() As Long
    '���ܣ�ˢ��װ���嵥
    Dim rsTemp As New ADODB.Recordset
    gstrSql = "Select R.ID, R.����, R.���� As ԭ����, R.��ʼ���� As ԭ��ʼ����, R.�������� As ԭ��������," & vbNewLine & _
            "       D.���� || '-' || D.���� As ����" & vbNewLine & _
            "From �����ʿ�Ʒ R, �������� D" & vbNewLine & _
            "Where R.����id = D.ID"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr("")): rptItem.HasCheckbox = True: rptItem.Checked = False
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !ID)
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !ԭ����)
            rptRcd.AddItem CStr("" & !ԭ��ʼ����)
            rptRcd.AddItem CStr("" & !ԭ��������)
            rptRcd.AddItem CStr("")
            rptRcd.AddItem CStr("")
            rptRcd.AddItem CStr("")
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    
    zlRefList = Me.rptList.Records.Count
    Call chkShow_Click(0)
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------

Private Sub chkShow_Click(Index As Integer)
    Dim strNow As String
    Dim blnShow As Boolean
    
    strNow = Format(Now(), "yyyy-MM-dd")
    
    For Each rptRcd In Me.rptList.Records
        blnShow = True
        If Me.chkShow(0).Value = vbChecked Then
            If strNow < rptRcd.Item(mCol.ԭ��ʼ����).Value Or strNow > rptRcd.Item(mCol.ԭ��������).Value Then blnShow = False
        End If
        If Me.chkShow(1).Value = vbChecked Then
            If rptRcd.Item(mCol.ѡ��).Checked = False Then blnShow = False
        End If
        rptRcd.Visible = blnShow
    Next
    Me.rptList.Populate
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDate_Click()
    For Each rptRcd In Me.rptList.Records
        rptRcd.Item(mCol.�¿�ʼ����).Value = Format(DateAdd("m", 12, rptRcd.Item(mCol.ԭ��ʼ����).Value), "yyyy-MM-dd")
        rptRcd.Item(mCol.�½�������).Value = Format(DateAdd("m", 12, rptRcd.Item(mCol.ԭ��������).Value), "yyyy-MM-dd")
    Next
    Me.rptList.Populate
End Sub

Private Sub cmdOK_Click()
    Dim strCopies As String
    strCopies = ""
    For Each rptRcd In Me.rptList.Records
        If rptRcd.Visible And rptRcd.Item(mCol.ѡ��).Checked Then
            If Trim(rptRcd.Item(mCol.������).Value) = "" Then
                MsgBox "ѡ���Ƶ��ʿ�Ʒ��" & vbCrLf & rptRcd.Item(mCol.����).Value & " δ���������ţ�", vbInformation, gstrSysName
                Set Me.rptList.FocusedRow = rptRcd: Exit Sub
            End If
            If Trim(rptRcd.Item(mCol.�¿�ʼ����).Value) = "" Then
                MsgBox "ѡ���Ƶ��ʿ�Ʒ��" & vbCrLf & rptRcd.Item(mCol.����).Value & " δ�����¿�ʼ���ڣ�", vbInformation, gstrSysName
                Set Me.rptList.FocusedRow = rptRcd: Exit Sub
            End If
            If Trim(rptRcd.Item(mCol.�½�������).Value) = "" Then
                MsgBox "ѡ���Ƶ��ʿ�Ʒ��" & vbCrLf & rptRcd.Item(mCol.����).Value & " δ�����½������ڣ�", vbInformation, gstrSysName
                Set Me.rptList.FocusedRow = rptRcd: Exit Sub
            End If
            strCopies = strCopies & "|" & rptRcd.Item(mCol.ID).Value
            strCopies = strCopies & ";" & rptRcd.Item(mCol.������).Value
            strCopies = strCopies & ";" & rptRcd.Item(mCol.�¿�ʼ����).Value
            strCopies = strCopies & ";" & rptRcd.Item(mCol.�½�������).Value
        End If
    Next
    If strCopies = "" Then MsgBox "��δѡ���Ƶ��ʿ�Ʒ��", vbInformation, gstrSysName: Exit Sub
    gstrSql = "Zl_�����ʿ�Ʒ_Copy(" & IIf(Me.optValue(0).Value, 0, 1) & ",'" & Mid(strCopies, 2) & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    For Each rptRcd In Me.rptList.Records
        rptRcd.Item(mCol.ѡ��).HasCheckbox = True
        rptRcd.Item(mCol.ѡ��).Checked = (Index = 0)
    Next
    Call chkShow_Click(0)
End Sub

Private Sub Form_Load()
    With Me.rptList
        .AutoColumnSizing = True
        Set rptCol = .Columns.Add(mCol.ѡ��, "", 18, False): rptCol.Editable = True: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.����, "����", 120, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 150, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.ԭ����, "ԭ����", 66, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.ԭ��ʼ����, "ԭ��ʼ����", 66, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.ԭ��������, "ԭ��������", 66, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.������, "������", 60, False): rptCol.Editable = True: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�¿�ʼ����, "�¿�ʼ����", 66, False): rptCol.Editable = True: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.�½�������, "�½�������", 66, False): rptCol.Editable = True: rptCol.Groupable = False
        
        .AllowEdit = True
        .EditOnClick = True
        .FocusSubItems = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        .GroupsOrder.Add .Columns.Find(mCol.����)
        .GroupsOrder(0).SortAscending = True
    End With
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)

    Call zlRefList
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.picBack
        .Left = Me.ScaleWidth - .Width
        .Height = Me.ScaleHeight
    End With
    With Me.rptList
        .Left = 0: .Width = Me.picBack.Left
        .Top = 0: .Height = Me.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub rptList_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim dtInput As Date
    If Trim(Item.Value) = "" Then Item.Value = CStr(""): Exit Sub
    Select Case Column.Index
    Case mCol.������
        Item.Value = UCase(Item.Value)
        For lngCount = 1 To Len(Item.Value)
            If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(Item.Value, lngCount, 1)) = 0 Then
                MsgBox "���Ű�����������ַ�", vbInformation, gstrSysName
                Item.Value = CStr(""): Exit Sub
            End If
        Next
        If Len(Item.Value) > 10 Then
            MsgBox "����̫����", vbInformation, gstrSysName
            Item.Value = CStr(""): Exit Sub
        End If
        If Item.Value = Row.Record.Item(mCol.ԭ����).Value Then
            MsgBox "�����Ų��ܺ�ԭ������ͬ��", vbInformation, gstrSysName
            Item.Value = CStr(""): Exit Sub
        End If
    Case mCol.�¿�ʼ����, mCol.�½�������
        Err = 0: On Error Resume Next
        dtInput = CDate(Item.Value)
        If Err <> 0 Then
            MsgBox "���������ڸ�ʽ��", vbInformation, gstrSysName
            Item.Value = CStr(""): Exit Sub
        End If
        
        Err = 0: On Error GoTo 0
        Item.Value = Format(dtInput, "yyyy-MM-dd")
        
        If Row.Record.Item(mCol.�¿�ʼ����).Value <> CStr("") And Row.Record.Item(mCol.�½�������).Value <> CStr("") Then
            If Row.Record.Item(mCol.�¿�ʼ����).Value >= Row.Record.Item(mCol.�½�������).Value Then
                MsgBox "�µ����ڷ�Χ����(�������ڱ�����ڿ�ʼ����)��", vbInformation, gstrSysName
                Item.Value = CStr(""): Exit Sub
            End If
        End If
    End Select
End Sub
