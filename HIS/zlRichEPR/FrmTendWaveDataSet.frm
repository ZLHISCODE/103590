VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTendWaveDataSet 
   Caption         =   "����ͬ����������"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9300
   Icon            =   "FrmTendWaveDataSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9300
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkItem 
      Caption         =   "ѡ��ͬ������Ŀ(����Ϊѡ��ͬ������Ŀ)"
      Height          =   345
      Left            =   5175
      TabIndex        =   10
      Top             =   5820
      Width           =   3735
   End
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   4410
      ScaleHeight     =   2835
      ScaleWidth      =   4185
      TabIndex        =   11
      Top             =   1770
      Width           =   4185
      Begin MSComctlLib.ListView lvwItem 
         Height          =   2475
         Left            =   0
         TabIndex        =   12
         Tag             =   "10"
         Top             =   0
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   4366
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
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ȫ��(&E)"
         Height          =   350
         Index           =   2
         Left            =   1080
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2475
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Index           =   3
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2475
         Width           =   1100
      End
   End
   Begin VB.CheckBox ChkAll 
      Caption         =   "ȫԺͨ��"
      Height          =   270
      Left            =   7035
      TabIndex        =   15
      Top             =   4905
      Width           =   1515
   End
   Begin VB.PictureBox PicDept 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   2145
      ScaleHeight     =   2940
      ScaleWidth      =   4185
      TabIndex        =   16
      Top             =   3045
      Width           =   4185
      Begin MSComctlLib.ListView lvwDept 
         Height          =   2475
         Left            =   0
         TabIndex        =   17
         Tag             =   "10"
         Top             =   75
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   4366
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ȫ��(&E)"
         Height          =   350
         Index           =   1
         Left            =   1095
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Index           =   0
         Left            =   15
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2550
         Width           =   1100
      End
   End
   Begin VB.PictureBox picAge 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   305
      Left            =   1005
      Picture         =   "FrmTendWaveDataSet.frx":076A
      ScaleHeight     =   300
      ScaleWidth      =   4020
      TabIndex        =   3
      Top             =   480
      Width           =   4020
      Begin VB.ComboBox cboAge 
         Height          =   300
         Index           =   2
         Left            =   2415
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   0
         Width           =   1065
      End
      Begin VB.TextBox txtAge 
         Height          =   300
         Index           =   1
         Left            =   3480
         MaxLength       =   3
         TabIndex        =   8
         Top             =   0
         Width           =   525
      End
      Begin VB.ComboBox cboAge 
         Height          =   300
         Index           =   1
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   810
      End
      Begin VB.TextBox txtAge 
         Height          =   300
         Index           =   0
         Left            =   1065
         MaxLength       =   3
         TabIndex        =   5
         Top             =   0
         Width           =   510
      End
      Begin VB.ComboBox cboAge 
         Height          =   300
         Index           =   0
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1755
      Index           =   0
      Left            =   360
      ScaleHeight     =   1755
      ScaleWidth      =   3510
      TabIndex        =   0
      Top             =   1590
      Width           =   3510
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2040
         Left            =   0
         TabIndex        =   1
         Top             =   -15
         Width           =   1995
         _Version        =   589884
         _ExtentX        =   3519
         _ExtentY        =   3598
         _StockProps     =   0
         BorderStyle     =   2
         ShowGroupBox    =   -1  'True
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   6255
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "FrmTendWaveDataSet.frx":91A0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13494
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsList 
      Left            =   4860
      Top             =   3585
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTendWaveDataSet.frx":9A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTendWaveDataSet.frx":10294
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTendWaveDataSet.frx":103EE
            Key             =   "User"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   540
      Left            =   1275
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2265
      Visible         =   0   'False
      Width           =   1095
      _cx             =   1931
      _cy             =   952
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3675
      Index           =   1
      Left            =   4185
      ScaleHeight     =   3675
      ScaleWidth      =   3510
      TabIndex        =   21
      Top             =   450
      Width           =   3510
      Begin VB.ComboBox cboNursGrade 
         Height          =   300
         Left            =   825
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2040
      End
      Begin XtremeSuiteControls.TaskPanel tkpMain 
         Height          =   3420
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   3045
         _Version        =   589884
         _ExtentX        =   5371
         _ExtentY        =   6032
         _StockProps     =   64
         VisualTheme     =   5
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   180
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "FrmTendWaveDataSet.frx":10548
      Left            =   1080
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "FrmTendWaveDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnChange As Boolean

Private Enum mCol
    ͼ�� = 0
    ���䷶Χ
    ����ȼ�
    ��Ŀ��Ϣ
    ���ÿ���
End Enum

Private Type Type_ItemDate
    strAgeFilter As String
    intNursGrade As Integer
    strItems As String
    strDept As String
End Type

Private T_ItemDate As Type_ItemDate

Private Sub cboAge_Click(Index As Integer)
    Dim intData As Integer
    Select Case Index
        Case 0
            If cboAge(1).ListCount > 0 And cboAge(1).ListIndex >= 0 Then intData = Val(cboAge(1).ItemData(cboAge(1).ListIndex))
            If cboAge(Index).ListIndex = 0 Or cboAge(Index).ListIndex = 1 Then
                With cboAge(1)
                    .Clear
                    .AddItem ""
                    .AddItem "����": .ItemData(.NewIndex) = 1
                    
                    txtAge(0).Enabled = mblnChange
                    cboAge(1).Enabled = mblnChange
                    txtAge(0).BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
                    cboAge(1).BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
                    
                    If intData = 1 Then
                        .ListIndex = 1
                    Else
                        .ListIndex = 0
                    End If
                End With
                
            ElseIf cboAge(Index).ListIndex = 2 Or cboAge(Index).ListIndex = 3 Then
                With cboAge(1)
                    .Clear
                    .AddItem ""
                    .AddItem "����": .ItemData(.NewIndex) = 2
                    
                    txtAge(0).Enabled = mblnChange
                    cboAge(1).Enabled = mblnChange
                    txtAge(0).BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
                    cboAge(1).BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
                
                    If intData = 2 Then
                        .ListIndex = 1
                    Else
                        .ListIndex = 0
                    End If
                End With
            Else
                txtAge(0).Enabled = False
                txtAge(0).BackColor = &H8000000F
                With cboAge(1)
                    .Enabled = False
                    .BackColor = &H8000000F
                    .Clear
                    .AddItem ""
                    .ListIndex = 0
                End With
            End If
        Case 1
            If cboAge(2).ListCount > 0 And cboAge(2).ListIndex >= 0 Then intData = Val(cboAge(2).ItemData(cboAge(2).ListIndex))
            If cboAge(Index).ItemData(cboAge(Index).ListIndex) = 1 Then
                With cboAge(2)
                    .Clear
                    .AddItem "����": .ItemData(.NewIndex) = 1
                    .AddItem "���ڵ���": .ItemData(.NewIndex) = 2
                    .Enabled = mblnChange
                    .BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
                    If intData = 1 Then
                        .ListIndex = 0
                    ElseIf intData = 2 Then
                        .ListIndex = 1
                    End If
                End With
            ElseIf cboAge(Index).ItemData(cboAge(Index).ListIndex) = 2 Then
                With cboAge(2)
                    .Clear
                    .AddItem "С��": .ItemData(.NewIndex) = 1
                    .AddItem "С�ڵ���": .ItemData(.NewIndex) = 2
                    .Enabled = mblnChange
                    .BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
                    If intData = 1 Then
                        .ListIndex = 0
                    ElseIf intData = 2 Then
                        .ListIndex = 1
                    End If
                End With
            Else
                With cboAge(2)
                    .Enabled = False
                    .BackColor = &H8000000F
                    .Clear
                    .AddItem ""
                    .ListIndex = 0
                End With
            End If
        Case 2
            If cboAge(Index).ItemData(cboAge(Index).ListIndex) = 1 Or cboAge(Index).ItemData(cboAge(Index).ListIndex) = 2 Then
                txtAge(1).Enabled = mblnChange
                txtAge(1).BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
            Else
                txtAge(1).Enabled = False
                txtAge(1).BackColor = &H8000000F
                txtAge(1).Text = ""
            End If
    End Select
    
    Call GetFilter
End Sub

Private Sub cboAge_KeyPress(Index As Integer, KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboAge(Index).hWnd, KeyAscii)
End Sub

Private Sub cboNursGrade_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim intNursGrade As Integer
    Dim objItem As ListItem
    Dim strItem As String, arrItem() As String
    Dim lngIndex As Long, lngCount As Long
    
    intNursGrade = cboNursGrade.ItemData(cboNursGrade.ListIndex)
    If Val(cboNursGrade.Tag) = intNursGrade + 1 Then Exit Sub
    cboNursGrade.Tag = intNursGrade + 1
    
    '��¼��֮ǰ����Ŀ��Ϣ
    strItem = ""
    For lngIndex = 1 To lvwItem.ListItems.Count
        If lvwItem.ListItems(lngIndex).Checked = True Then
            strItem = strItem & ";" & Val(lvwItem.ListItems(lngIndex).Text)
        End If
    Next lngIndex
    strItem = Mid(strItem, 2)
    '���ÿ���
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "_���", "��Ŀ���", 900
        .Add , "_����", "��Ŀ����", 2000
        Me.lvwItem.ListItems.Clear
    End With
    
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "SELECT A.��Ŀ���,A.��Ŀ����,A.Ӧ�÷�ʽ  FROM �����¼��Ŀ A,���¼�¼��Ŀ B" & vbNewLine & _
        "WHERE A.��Ŀ���=B.��Ŀ��� AND NVL(A.Ӧ�÷�ʽ,0)<>0 AND NVL(A.��Ŀ����,0)=0 AND NVL(A.����ȼ�,3)>=[1]" & vbNewLine & _
        "ORDER BY ��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intNursGrade)
    lvwItem.Enabled = (rsTemp.RecordCount > 0)
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !��Ŀ���, !��Ŀ���)
            objItem.SubItems(Me.lvwItem.ColumnHeaders("_����").Index - 1) = !��Ŀ����
            objItem.Tag = Val(NVL(!Ӧ�÷�ʽ))
            If InStr(1, ";" & strItem & ";", ";" & !��Ŀ��� & ";") > 0 Then objItem.Checked = True
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(vsfPrint, rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = "����ͬ����Ŀ�����嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cboNursGrade_KeyPress(KeyAscii As Integer)
    zlControl.CboMatchIndex cboNursGrade.hWnd, KeyAscii
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As Object
    Dim strNursGrade As String, strAge As String
    
    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(2)
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_Edit_NewItem '����
            Call ExecuteCommand("��������")
        Case conMenu_Edit_Append '��������
            Call ExecuteCommand("��������")
        Case conMenu_Edit_Modify '�޸�
            Call ExecuteCommand("�޸�����")
        Case conMenu_Edit_Delete 'ɾ��
            Call ExecuteCommand("ɾ������")
        Case conMenu_Edit_Transf_Save '��������
            strNursGrade = cboNursGrade.Text
            strAge = GetFilter(1)
            Call ExecuteCommand("��������", strNursGrade, strAge)
        Case conMenu_Edit_Transf_Cancle 'ȡ��
            If rptList.Records.Count > 0 Then
                If Not rptList.FocusedRow Is Nothing Then
                    If Not rptList.FocusedRow.GroupRow Then
                        strNursGrade = rptList.FocusedRow.Record(mCol.����ȼ�).Value
                        strAge = rptList.FocusedRow.Record(mCol.���䷶Χ).Record.Tag
                    End If
                End If
            End If
            Call ExecuteCommand("ȡ������", strNursGrade, strAge)
        Case conMenu_View_Refresh 'ˢ������
            strNursGrade = "-1-���л���"
            If rptList.Records.Count > 0 Then
                If Not rptList.FocusedRow Is Nothing Then
                    If Not rptList.FocusedRow.GroupRow Then
                        strNursGrade = rptList.FocusedRow.Record(mCol.����ȼ�).Value
                        strAge = rptList.FocusedRow.Record(mCol.���䷶Χ).Record.Tag
                    End If
                End If
            End If
            Call ExecuteCommand("ˢ������", strNursGrade, strAge)
        Case conMenu_View_ToolBar_Button
            cbsMain(2).Visible = Not cbsMain(2).Visible
            cbsMain.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
        
            For Each cbrControl In cbsMain(2).Controls
                If cbrControl.Type = xtpControlButton Then
                    cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
            cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size      '��ͼ��
    
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            cbsMain.RecalcLayout
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsMain.RecalcLayout
        Case conMenu_Help_Help
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, 1)
        Case conMenu_Help_About
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
        Case conMenu_Help_Web_Home
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hWnd)
        Case conMenu_Help_Web_Mail
            Call zlMailTo(Me.hWnd)
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    picPane(0).Enabled = Not mblnChange
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (rptList.Records.Count > 0)
    Case conMenu_Edit_NewItem '����
        Control.Enabled = Not mblnChange
    Case conMenu_Edit_Append '��������
        Control.Enabled = Not mblnChange
    Case conMenu_Edit_Modify '�޸�
        Control.Enabled = Not mblnChange And rptList.Records.Count > 0
        If Control.Enabled = True Then
            If rptList.FocusedRow Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not rptList.FocusedRow.GroupRow
            End If
        End If
    Case conMenu_Edit_Delete 'ɾ��
        Control.Enabled = Not mblnChange And rptList.Records.Count > 0
        If Control.Enabled = True Then
            If rptList.FocusedRow Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not rptList.FocusedRow.GroupRow
            End If
        End If
    Case conMenu_Edit_Transf_Save '��������
        Control.Enabled = mblnChange
    Case conMenu_Edit_Transf_Cancle 'ȡ��
        Control.Enabled = mblnChange
    Case conMenu_View_Refresh
        Control.Enabled = Not mblnChange
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsMain(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub ChkAll_Click()
    Dim lngIndex As Long, lngCount As Long
    Dim blnEnable As Boolean
    Dim strTmp As String, arrTmp() As String
    blnEnable = Not (ChkAll.Value = 1) And mblnChange
    
    lvwDept.Enabled = blnEnable
    lvwDept.BackColor = IIf(blnEnable = True, &H80000005, &H8000000F)
    cmdSelect(0).Enabled = blnEnable And (lvwDept.ListItems.Count > 0)
    cmdSelect(1).Enabled = blnEnable And (lvwDept.ListItems.Count > 0)
    
    If mblnChange = True Then
        If ChkAll.Value = 1 Then
            For lngIndex = 1 To lvwDept.ListItems.Count
                lvwDept.ListItems(lngIndex).Checked = False
            Next lngIndex
        Else
            If rptList.Tag = "�޸�" Then
                If rptList.Records.Count <= 0 Then Exit Sub
                If rptList.FocusedRow Is Nothing Then Exit Sub
                If rptList.FocusedRow.GroupRow = True Then Exit Sub
                strTmp = rptList.FocusedRow.Record(mCol.���ÿ���).Value
                If strTmp = "" Then strTmp = "-1"
                If Val(strTmp) = -1 Then Exit Sub
                arrTmp = Split(strTmp, ";")
                For lngCount = 0 To UBound(arrTmp)
                   For lngIndex = 1 To lvwDept.ListItems.Count
                       If Val(arrTmp(lngCount)) = Val(Mid(lvwDept.ListItems(lngIndex).Key, 2)) Then
                           lvwDept.ListItems(lngIndex).Checked = True
                           Exit For
                       End If
                   Next lngIndex
                Next lngCount
            End If
        End If
    End If
End Sub

Private Sub chkItem_Click()
'����: ��Ŀѡ���л� (��ѡ��ͬ������Ŀ����ѡ��ͬ������Ŀ)
    Dim intCheck As Integer
    Dim lngIndex As Long
    Dim strItem As String
    
    intCheck = chkItem.Value
    If Val(chkItem.Tag) = intCheck Then Exit Sub
    chkItem.Tag = intCheck
    '��ȷ��Ŀǰѡ������Щ��Ŀ
    For lngIndex = 1 To lvwItem.ListItems.Count
        If lvwItem.ListItems(lngIndex).Checked = True Then
            strItem = strItem & "," & Val(lvwItem.ListItems(lngIndex).Text)
        End If
    Next lngIndex
    strItem = Mid(strItem, 2)
    
    If strItem = "" Then Exit Sub
    For lngIndex = 1 To lvwItem.ListItems.Count
        If InStr(1, "," & strItem & ",", "," & Val(lvwItem.ListItems(lngIndex).Text) & ",") = 0 Then
            lvwItem.ListItems(lngIndex).Checked = True
        Else
            lvwItem.ListItems(lngIndex).Checked = False
        End If
    Next lngIndex
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim objItem As ListItem
    Select Case Index
        Case 0, 1
            If lvwDept.ListItems.Count <= 0 Then Exit Sub
            For Each objItem In Me.lvwDept.ListItems
                objItem.Checked = (Index = 0)
            Next
        Case 2, 3
            If lvwItem.ListItems.Count <= 0 Then Exit Sub
            For Each objItem In Me.lvwItem.ListItems
                objItem.Checked = (Index = 3)
            Next
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Call ExecuteCommand("��ʼ����")
    Call ExecuteCommand("��ȡ����")
End Sub


Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    Dim blnAdd As Boolean
    
    Dim arrNurs() As Variant
    If UBound(varParam) < 1 Then
        arrNurs = Array("-1-���л���", "")
    Else
        arrNurs = varParam
    End If
    Select Case strCommand
        Case "��ʼ����"
            rptList.Tag = ""
            mblnChange = False
            Call InitCommandBar
            Call InitDockPannel
            Call CreateToolBox
            Call InitData
            Call RestoreWinState(Me, App.ProductName)
        Case "��ȡ����"
            rptList.Tag = ""
            Call ReadData(arrNurs)
            Call RefreshStateInfo
        Case "ˢ������"
            rptList.Tag = ""
            mblnChange = False
            Call InitData
            Call ReadData(arrNurs)
            Call RefreshStateInfo
        Case "��������"
            blnAdd = rptList.Tag = "��������"
            If Not SaveData Then Exit Function
            Call ExecuteCommand("ˢ������", varParam(0), varParam(1))
            If blnAdd Then Call ExecuteCommand("��������")
        Case "ȡ������"
            rptList.Tag = ""
            mblnChange = False
            Call rptList_SelectionChanged
            If picPane(0).Enabled And picPane(0).Visible Then picPane(0).SetFocus
        Case "��������"
            rptList.Tag = "����"
            mblnChange = True
            Call ClearControl
            If cboAge(0).Enabled And cboAge(0).Visible Then cboAge(0).SetFocus
        Case "��������"
            rptList.Tag = "��������"
            mblnChange = True
            Call ClearControl
            If cboAge(0).Enabled And cboAge(0).Visible Then cboAge(0).SetFocus
        Case "�޸�����"
            mblnChange = True
            rptList.Tag = "�޸�"
            Call ClearControl(False)
            If cboAge(0).Enabled And cboAge(0).Visible Then cboAge(0).SetFocus
        Case "ɾ������"
            rptList.Tag = "ɾ��"
            If Not DeleteData Then Exit Function
            Call ExecuteCommand("ˢ������")
    End Select
End Function

Private Sub InitData()
    Dim rsTemp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim rptCol As ReportColumn
    On Error GoTo errHand
    With rptList
        .Records.DeleteAll: .Columns.DeleteAll
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 20, False)
        rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
    
        Set rptCol = .Columns.Add(mCol.���䷶Χ, "���䷶Χ", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����ȼ�, "����ȼ�", 100, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��Ŀ��Ϣ, "��ͬ������Ŀ", 300, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.���ÿ���, "���ÿ���", 300, False): rptCol.Editable = False: rptCol.Groupable = False
       
        '.SetImageList ilsList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GridLineColor = RGB(225, 225, 225)
            .NoItemsText = "û�п���ʾ����Ϣ..."
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶�����ȼ��б��⵽����,�����з���..."
        End With
        .PreviewMode = True
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.����ȼ�)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(mCol.���䷶Χ)
    End With
    
    '�����
    With cboAge(0)
        .Clear
        .AddItem "С��"
        .AddItem "С�ڵ���"
        .AddItem "����"
        .AddItem "���ڵ���"
    End With
    '����ȼ�
    With cboNursGrade
        .Tag = "-1"
        .Clear
        .AddItem "-1-���л���": .ItemData(.NewIndex) = -1
        .AddItem "0-�ؼ�����": .ItemData(.NewIndex) = 0
        .AddItem "1-һ������": .ItemData(.NewIndex) = 1
        .AddItem "2-��������": .ItemData(.NewIndex) = 2
        .AddItem "3-��������": .ItemData(.NewIndex) = 3
        .ListIndex = 0
    End With
    
    '���ÿ���
    With Me.lvwDept.ColumnHeaders
        .Clear
        .Add , "_����", "����", 900
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 800
    End With
    With Me.lvwDept
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
        .Sorted = True
        .ListItems.Clear
    End With

    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.Id, d.����, d.����, d.����" & _
            " From ���ű� d, ��������˵�� m" & _
            " Where d.Id = m.����id  And m.�������� = '�ٴ�' And m.������� In (2, 3)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    lvwDept.Enabled = (rsTemp.RecordCount > 0)
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwDept.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwDept.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwDept.ColumnHeaders("_����").Index - 1) = "" & !����
            .MoveNext
        Loop
    End With
    
    Call ClearControl
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReadData(varParam() As Variant)
'���ܣ���ȡ������Ϣ
    Dim rsTemp As New ADODB.Recordset
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim lngRow As Long
    
    On Error GoTo errHand
    gstrSQL = "SELECT ����ȼ�,���䷶Χ,������Ŀ,���ÿ��� FROM ����ͬ����Ŀ"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "����ͬ����Ŀ")
    With rsTemp
        rptList.Records.DeleteAll
        Do While Not .EOF
            Set rptRcd = rptList.Records.Add()
            Set rptItem = rptRcd.AddItem("")
            'rptItem.Icon =
            Set rptItem = rptRcd.AddItem(Replace(NVL(rsTemp!���䷶Χ), vbTab, ""))
            rptItem.Record.Tag = NVL(rsTemp!���䷶Χ)
            rptRcd.AddItem NursGradeSwitch(NVL(rsTemp!����ȼ�, -1))
            rptRcd.AddItem NursItemSwitch(NVL(rsTemp!������Ŀ))
            rptRcd.AddItem DeptSwitch(NVL(rsTemp!���ÿ���, -1))
        .MoveNext
        Loop
        rptList.Populate
    End With
    On Error Resume Next
    With rptList
        For lngRow = 0 To rptList.Rows.Count - 1
            If Not rptList.Rows(lngRow).GroupRow Then
                If rptList.Rows(lngRow).Record(mCol.����ȼ�).Value = CStr(varParam(0)) _
                    And rptList.Rows(lngRow).Record(mCol.���䷶Χ).Record.Tag = CStr(varParam(1)) Then
                    rptList.FocusedRow = rptList.Rows(lngRow)
                    rptList.FocusedRow.Selected = True
                    Exit For
                End If
            End If
        Next
    End With
    If rptList.FocusedRow Is Nothing Then
        If rptList.Records.Count > 0 Then
            Set rptList.FocusedRow = rptList.Rows(1)
            rptList.FocusedRow.Selected = True
        End If
    End If
    
    If picPane(0).Enabled And picPane(0).Visible Then picPane(0).SetFocus
    If Not UCase(Me.ActiveControl.Name) = "RPTLIST" Then
        If rptList.Visible And rptList.Records.Count Then rptList.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearControl(Optional ByVal blnClear As Boolean = True)
    Dim lngIndex As Long
    
    If blnClear = True Then
        '�ؼ���������
        cboAge(0).ListIndex = -1
        txtAge(0).Text = ""
        
        If cboNursGrade.ListCount > 0 Then cboNursGrade.ListIndex = 0
        For lngIndex = 1 To lvwItem.ListItems.Count
            lvwItem.ListItems(lngIndex).Checked = False
        Next lngIndex
        chkItem.Value = 0: chkItem.Tag = ""
        
        For lngIndex = 1 To lvwDept.ListItems.Count
            lvwDept.ListItems(lngIndex).Checked = False
        Next lngIndex
        If ChkAll.Value <> 1 Then
            ChkAll.Value = 1
        Else
            Call ChkAll_Click
        End If
    End If
    
    '�ռ��Ƿ��������
    picPane(0).Enabled = Not mblnChange
    cboAge(0).Enabled = mblnChange
    cboAge(0).BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
    txtAge(0).Enabled = mblnChange
    txtAge(0).BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
    Call cboAge_Click(0)
    
    cboNursGrade.Enabled = mblnChange
    cboNursGrade.BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
    
    chkItem.Enabled = mblnChange
    picItem.Enabled = mblnChange
    lvwItem.Enabled = mblnChange
    lvwItem.BackColor = IIf(mblnChange = True, &H80000005, &H8000000F)
    cmdSelect(2).Enabled = mblnChange And (lvwItem.ListItems.Count > 0)
    cmdSelect(3).Enabled = mblnChange And (lvwItem.ListItems.Count > 0)
    ChkAll.Enabled = mblnChange
    If blnClear = False Then Call ChkAll_Click
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = True

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)

    '------------------------------------------------------------------------------------------------------------------
    '�༭
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "��������(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Append, "��������(&A)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸�����(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ������(&D)")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "�������(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������(&R)")

    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")

    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)


    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)

    '------------------------------------------------------------------------------------------------------------------
    '����������:������������

    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�", False)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")

    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings

        .Add 0, vbKeyF5, conMenu_View_Refresh           'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help              '����

        .Add FCONTROL, vbKeyP, conMenu_File_Print       '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem     '����
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify      '�޸�
        .Add FCONTROL, vbKeyD, conMenu_Edit_Delete     'ɾ��
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save '����
        .Add FCONTROL, vbKeyC, conMenu_Edit_Transf_Cancle 'ȡ��
    End With

End Function


Private Sub InitDockPannel()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim objPane As Pane
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    dkpMain.SetCommandBars cbsMain
    
    Set objPane = dkpMain.CreatePane(1, 300, 100, DockLeftOf, Nothing)
    objPane.Title = "�嵥"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(2, 100, 100, DockRightOf, Nothing)
    objPane.Title = "��ϸ"
    objPane.Options = PaneNoCaption
End Sub

Private Function CreateToolBox() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    Dim objGrp As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim objIlsItem As Object
    
    Call tkpMain.SetImageList(ilsList)
    
    Set objGrp = tkpMain.Groups.Add(0, "���䷶Χ(��)")
    objGrp.Expandable = False
    
    Set objItem = objGrp.Items.Add(0, "����(��)��", xtpTaskItemTypeControl)
    Call objGrp.Items.Add(0, "˵  ����", xtpTaskItemTypeText)
    Set tkpMain.Groups(1).Items(1).Control = picAge
    
    Set objGrp = tkpMain.Groups.Add(1, "���û���ȼ�")
    objGrp.Expandable = False
    Call objGrp.Items.Add(1, "", xtpTaskItemTypeControl)
    Set tkpMain.Groups(2).Items(1).Control = cboNursGrade
    
    Call tkpMain.SetImageList(ilsList)
    Set objGrp = tkpMain.Groups.Add(2, "��Ŀѡ��")
    objGrp.Expandable = False
    Call objGrp.Items.Add(1, "", xtpTaskItemTypeControl)
    Set tkpMain.Groups(3).Items(1).Control = chkItem
    Call objGrp.Items.Add(2, "", xtpTaskItemTypeControl)
    Set tkpMain.Groups(3).Items(2).Control = picItem
    
    Set objGrp = tkpMain.Groups.Add(3, "���ÿ���")
    objGrp.Expandable = False
    Call objGrp.Items.Add(1, "", xtpTaskItemTypeControl)
    Set tkpMain.Groups(4).Items(1).Control = ChkAll
    Call objGrp.Items.Add(2, "", xtpTaskItemTypeControl)
    Set tkpMain.Groups(4).Items(2).Control = PicDept
    
    
    tkpMain.Animation = xtpTaskPanelAnimationNo
    tkpMain.Behaviour = xtpTaskPanelBehaviourExplorer
    tkpMain.HotTrackStyle = xtpTaskPanelHighlightItem
    tkpMain.VisualTheme = xtpTaskPanelThemeOffice2003Plain
    tkpMain.SetGroupInnerMargins 0, 1, 1, 1
    
    tkpMain.AllowDrag = False
    tkpMain.SelectItemOnFocus = False

    tkpMain.Groups(1).Expanded = True
    
    CreateToolBox = True
    
End Function

Private Sub Form_Resize()
    On Error Resume Next

    Call SetPaneRange(dkpMain, 2, 305, 15, 305, Me.ScaleHeight)

    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwItem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lngIndex As Long, lngItemNO As Long
    If Val(Item.Text) = 4 Or Val(Item.Text) = 5 Then
        '����ѹ������ѹ����ͬ��
        If Val(Item.Text) = 4 Then
            lngItemNO = 5
        Else
            lngItemNO = 4
        End If
        
        For lngIndex = 1 To lvwItem.ListItems.Count
            If Val(lvwItem.ListItems(lngIndex).Text) = lngItemNO And lvwItem.ListItems(lngIndex).Checked <> Item.Checked Then
                lvwItem.ListItems(lngIndex).Checked = Item.Checked
            End If
        Next lngIndex
    ElseIf Val(Item.Text) = -1 And Val(Item.Tag) = 2 Then
        '���ʺ���������ʱ������ͬ����������ͬ��
        For lngIndex = 1 To lvwItem.ListItems.Count
            If Val(lvwItem.ListItems(lngIndex).Text) = 2 Then
                If chkItem.Value = 0 Then '��ͬ����ѡ����
                    If Item.Checked = False And lvwItem.ListItems(lngIndex).Checked = True Then
                        lvwItem.ListItems(lngIndex).Checked = False
                    End If
                Else 'ͬ����Ŀ��ѡ����
                    If Item.Checked = True And lvwItem.ListItems(lngIndex).Checked = False Then
                        lvwItem.ListItems(lngIndex).Checked = True
                    End If
                End If
                Exit For
            End If
        Next lngIndex
    ElseIf Val(Item.Text) = 2 Then
        '���ʺ���������ʱ������ͬ����������ͬ��
        For lngIndex = 1 To lvwItem.ListItems.Count
            If Val(lvwItem.ListItems(lngIndex).Text) = -1 And Val(lvwItem.ListItems(lngIndex).Tag) = 2 Then
                If chkItem.Value = 0 Then '��ͬ����ѡ����
                    If Item.Checked = True And lvwItem.ListItems(lngIndex).Checked = False Then
                        lvwItem.ListItems(lngIndex).Checked = True
                    End If
                Else 'ͬ����Ŀ��ѡ����
                    If Item.Checked = False And lvwItem.ListItems(lngIndex).Checked = True Then
                        lvwItem.ListItems(lngIndex).Checked = False
                    End If
                End If
                Exit For
            End If
        Next lngIndex
    End If
End Sub

Private Sub PicDept_GotFocus()
    If lvwDept.Enabled And lvwDept.Visible Then lvwDept.SetFocus
End Sub

Private Sub PicDept_Resize()
    On Error Resume Next
    lvwDept.Width = PicDept.ScaleWidth - lvwDept.Left
End Sub

Private Sub picItem_GotFocus()
    If lvwItem.Enabled And lvwItem.Visible Then lvwItem.SetFocus
End Sub

Private Sub picItem_Resize()
    On Error Resume Next
    lvwItem.Width = picItem.ScaleWidth - lvwItem.Left
End Sub

Private Sub picPane_GotFocus(Index As Integer)
    If Index = 0 Then
        If rptList.Visible And rptList.Records.Count > 0 Then rptList.SetFocus
    ElseIf Index = 1 Then
        If mblnChange = False Then If picPane(0).Enabled And picPane(0).Visible Then picPane(0).SetFocus
    End If
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    If Index = 0 Then
        With rptList
            .Top = 0
            .Left = 0
            .Width = picPane(Index).ScaleWidth
            .Height = picPane(Index).ScaleHeight
        End With
    Else
        With tkpMain
            .Top = 0
            .Left = 0
            .Width = picPane(Index).ScaleWidth
            .Height = picPane(Index).ScaleHeight
        End With
    End If
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not (ROW Is Nothing) Then
        If Not ROW.GroupRow Then
            Call ExecuteCommand("�޸�����")
        End If
    End If
End Sub

Private Sub rptList_SelectionChanged()
    Dim strTmp As String, arrTmp() As String
    Dim lngCount As Long, lngIndex As Long
    
    Call ClearControl
    If rptList.FocusedRow Is Nothing Then Exit Sub
    With rptList.FocusedRow
        If Not .GroupRow Then
            strTmp = .Record(mCol.���䷶Χ).Record.Tag
            arrTmp = Split(strTmp, vbTab)
            For lngCount = 0 To UBound(arrTmp)
                Select Case lngCount
                    Case 0
                        Call zlControl.CboLocate(cboAge(0), arrTmp(0))
                    Case 1
                        txtAge(0).Text = Val(arrTmp(1))
                    Case 2
                        Call zlControl.CboLocate(cboAge(1), arrTmp(2))
                    Case 3
                        Call zlControl.CboLocate(cboAge(2), arrTmp(3))
                    Case 4
                        txtAge(1).Text = Val(arrTmp(4))
                End Select
            Next lngCount
            strTmp = .Record(mCol.����ȼ�).Value
            Call zlControl.CboLocate(cboNursGrade, Val(strTmp), True)
            
            strTmp = .Record(mCol.��Ŀ��Ϣ).Value
            arrTmp = Split(strTmp, ";")
            For lngCount = 0 To UBound(arrTmp)
                For lngIndex = 1 To lvwItem.ListItems.Count
                    If Val(arrTmp(lngCount)) = Val(lvwItem.ListItems(lngIndex).Text) Then
                        lvwItem.ListItems(lngIndex).Checked = True
                        Exit For
                    End If
                Next lngIndex
            Next lngCount
             
            strTmp = .Record(mCol.���ÿ���).Value
            If strTmp = "" Then strTmp = "-1"
            arrTmp = Split(strTmp, ";")
            If Not Val(strTmp) = -1 Then
                For lngCount = 0 To UBound(arrTmp)
                   For lngIndex = 1 To lvwDept.ListItems.Count
                       If Val(arrTmp(lngCount)) = Val(Mid(lvwDept.ListItems(lngIndex).Key, 2)) Then
                           lvwDept.ListItems(lngIndex).Checked = True
                           Exit For
                       End If
                   Next lngIndex
                Next lngCount
                If ChkAll.Value <> 0 Then
                    ChkAll.Value = 0
                Else
                    Call ChkAll_Click
                End If
            Else
                If ChkAll.Value <> 1 Then
                    ChkAll.Value = 1
                Else
                    Call ChkAll_Click
                End If
            End If
        End If
    End With
End Sub

Private Sub tkpMain_GotFocus()
    If mblnChange = False Then Call picPane_GotFocus(1)
End Sub

Private Sub txtAge_Change(Index As Integer)
    Call GetFilter
End Sub

Private Sub txtAge_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtAge(Index)
End Sub

Private Sub txtAge_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Exit Sub
    If InStr(1, ",1,2,3,4,5,6,7,8,9,0," & Chr(8) & ",", "," & Chr(KeyAscii) & ",") = 0 Then KeyAscii = 0
End Sub

Private Function GetFilter(Optional ByVal intMode As Integer = 0) As String
'����:��װ����������Ϣ
    Dim strFilter As String
    If intMode = 0 Then
        strFilter = "����" & cboAge(0).Text & txtAge(0).Text & cboAge(1).Text & cboAge(2).Text & txtAge(1).Text
         If tkpMain.Groups.Count >= 1 Then
            If tkpMain.Groups(1).Items.Count >= 2 Then
                tkpMain.Groups(1).Items(2).Caption = "˵  ����" & strFilter
            End If
         End If
    Else
        strFilter = cboAge(0).Text & vbTab & txtAge(0).Text
        If cboAge(1).ListIndex = 1 And cboAge(2).ListIndex >= 0 Then
            strFilter = cboAge(0).Text & vbTab & txtAge(0).Text & vbTab & cboAge(1).Text & vbTab & cboAge(2).Text & vbTab & txtAge(1).Text
            If cboAge(0).Text Like "С��*" And cboAge(1).Text = "����" And cboAge(2).Text Like "����*" Then
                strFilter = cboAge(2).Text & vbTab & txtAge(1).Text & vbTab & cboAge(1).Text & vbTab & cboAge(0).Text & vbTab & txtAge(0).Text
            End If
        End If
    End If
    
    GetFilter = strFilter
End Function

Private Function IsValid() As Boolean
'���ܣ�������ݵĺϷ���Ч��
    Dim lngIndex As Long, lngRow As Long
    Dim blnCheck As Boolean
    Dim strInfo As String
    Dim intNursGrade As Integer, strAgeFilter As String
    Dim arrAge() As String, arrAge1() As String
    
    '��һ��:������ݲ�����
    If cboAge(0).ListIndex = -1 Then
        MsgBox "���Ƚ������䷶Χ�������ã�", vbInformation, gstrSysName
        If cboAge(0).Enabled And cboAge(0).Visible Then cboAge(0).SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtAge(0).Text) Then
        MsgBox "���䷶Χ�����������ʽ���ǺϷ�����,���飡", vbInformation, gstrSysName
        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
        Exit Function
    End If
    
    If cboAge(1).ListIndex = 1 And cboAge(2).ListIndex >= 0 Then
        If Not IsNumeric(txtAge(1).Text) Then
            MsgBox "���䷶Χ�����������ʽ���ǺϷ�����,���飡", vbInformation, gstrSysName
            If txtAge(1).Enabled And txtAge(1).Visible Then txtAge(1).SetFocus
            Exit Function
        End If
        
        '������������Ƿ񽻲�
        If Mid(cboAge(0).Text, 1, 2) = Mid(cboAge(1).Text, 1, 2) Then
            MsgBox "���䷶Χ�����������棬����ͬʱ����С�ڻ���ڿ�ͷ�ĵ�ʽ��������������", vbInformation, gstrSysName
            If cboAge(1).Enabled And cboAge(1).Visible Then cboAge(1).SetFocus
            Exit Function
        End If
        
        strInfo = "[����" & cboAge(0).Text & txtAge(0).Text & cboAge(1).Text & cboAge(2).Text & txtAge(1).Text & "]"
        If Mid(cboAge(0).Text, 1, 2) Like "С��*" Then
            If cboAge(1).Text = "����" Then '����һ����ʽ�϶��Ǵ��ڻ���ڵ���
                If Val(txtAge(0).Text) < Val(txtAge(1).Text) Or (Val(txtAge(0).Text) = Val(txtAge(1).Text) _
                        And Not (cboAge(0).Text = "С�ڵ���" And cboAge(2).Text = "���ڵ���")) Then
                    MsgBox "���䷶Χ�������ʽ" & strInfo & "������������", vbInformation, gstrSysName
                    If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                    Exit Function
                End If
            End If
            If cboAge(1).Text = "����" Then
                If Not Val(txtAge(0).Text) < Val(txtAge(1).Text) Then
                    MsgBox "���䷶Χ�������ʽ��ϵΪ[����]ʱ�����䲻�ܽ��棬��������" & vbCrLf & _
                        "���ʽ" & strInfo, vbInformation, gstrSysName
                    If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                    Exit Function
                End If
            End If
        End If
        If Mid(cboAge(0).Text, 1, 2) Like "����*" Then
            If cboAge(1).Text = "����" Then '����һ����ʽ�϶���С�ڻ�С�ڵ���
                If Val(txtAge(0).Text) > Val(txtAge(1).Text) Or (Val(txtAge(0).Text) = Val(txtAge(1).Text) _
                        And Not (cboAge(0).Text = "���ڵ���" And cboAge(2).Text = "С�ڵ���")) Then
                    MsgBox "���䷶Χ�������ʽ" & strInfo & "������������", vbInformation, gstrSysName
                    If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                    Exit Function
                End If
            End If
            If cboAge(1).Text = "����" Then
                If Not Val(txtAge(0).Text) > Val(txtAge(1).Text) Then
                    MsgBox "���䷶Χ�������ʽ��ϵΪ[����]ʱ�����䲻�ܽ��棬��������" & vbCrLf & _
                        "���ʽ" & strInfo, vbInformation, gstrSysName
                    If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    If cboNursGrade.ListIndex = -1 Then
        MsgBox "��ѡ���Ӧ�Ļ���ȼ���", vbInformation, gstrSysName
        If cboNursGrade.Enabled And cboNursGrade.Visible Then cboNursGrade.SetFocus
        Exit Function
    End If
    
    If ChkAll.Value <> 1 Then
        For lngIndex = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(lngIndex).Checked = True Then
                blnCheck = True
                Exit For
            End If
        Next
        If blnCheck = False Then
            MsgBox "������Ҫѡ��һ����Ч�����ÿ��ң�", vbInformation, gstrSysName
            If lvwDept.Enabled And lvwDept.Visible Then lvwDept.SetFocus
            Exit Function
        End If
    End If
    
    '�ڶ�����������õ�ͬ�������Ƿ��Ѿ��ظ�,ͬһ����ȼ�������β��ܽ��桢�ظ�
    With rptList
        If .Records.Count <= 0 Then
            IsValid = True
            Exit Function
        End If
        strAgeFilter = GetFilter(1) '��ȡ���ν�Ҫ�������������
        arrAge1 = Split(strAgeFilter, vbTab)
        For lngRow = 0 To .Rows.Count - 1
            If Not .Rows(lngRow).GroupRow Then
                If rptList.Tag = "�޸�" Then
                    If rptList.FocusedRow.Index = lngRow Then GoTo GoNext
                End If
                intNursGrade = Val(.Rows(lngRow).Record(mCol.����ȼ�).Value)
                strAgeFilter = .Rows(lngRow).Record(mCol.���䷶Χ).Record.Tag
                '�ȼ�黤��ȼ��Ƿ��Ѿ����ڣ����ڵĻ��ڼ������������Ƿ��ظ�����
                If intNursGrade = cboNursGrade.ItemData(cboNursGrade.ListIndex) And strAgeFilter <> "" Then
                    '���������ѱ������������
                    arrAge = Split(strAgeFilter, vbTab)
                    If UBound(arrAge) > 0 And UBound(arrAge) <= 2 Then
                        strAgeFilter = arrAge(0) & vbTab & arrAge(1)
                    ElseIf UBound(arrAge) <= 0 Then
                        strAgeFilter = ""
                    Else
                        If UBound(arrAge) < 4 Then arrAge = Split(strAgeFilter & vbTab, vbTab)
                        If arrAge(0) Like "С��*" And arrAge(2) = "����" And arrAge(3) Like "����*" Then
                            strAgeFilter = arrAge(3) & vbTab & arrAge(4) & vbTab & arrAge(2) & vbTab & arrAge(0) & vbTab & arrAge(1)
                        End If
                    End If
                    If strAgeFilter <> "" Then
                        strInfo = "����ǰ���õ����䷶Χ�������ʽ������" & Replace(GetFilter(1), vbTab, "") & "������ȼ���" & cboNursGrade.Text & "��" & vbCrLf & _
                                    "����ʷ�����䷶Χ�������ʽ������" & Replace(strAgeFilter, vbTab, "") & "������ȼ���" & NursGradeSwitch(intNursGrade) & "�����ڽ���,���������ã�"
                        arrAge = Split(strAgeFilter, vbTab)
                
                        '���ֶ��������һ���ߵ����
                        If UBound(arrAge) > 3 And UBound(arrAge1) > 3 Then
                            '��������ֻ�ܴ���һ��,����ؽ�����
                            If arrAge(2) = "����" And arrAge1(2) = "����" Then
                                MsgBox strInfo, vbInformation, gstrSysName
                                If cboAge(1).Enabled And cboAge(1).Visible Then cboAge(1).SetFocus
                                Exit Function
                            End If
                            '��鲢�������Ƿ񽻲�,�磺��ʷ����:����>=0��������<=3 ,��������:����>=3��������<7,�������ʹ��ڽ���
                            If arrAge(2) = "����" And arrAge1(2) = "����" Then
                                If Val(arrAge(4)) > Val(arrAge1(1)) Or (Val(arrAge(4)) = Val(arrAge1(1)) And arrAge(3) Like "*����" And arrAge1(0) Like "*����") Then
                                    MsgBox strInfo, vbInformation, gstrSysName
                                    If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                    Exit Function
                                End If
                            End If
                            '�����ʷ�����в��ң������ǻ��ߣ���ô�����������Сֵ����С�ڻ�С�ڵ�����ʷ����Сֵ�����ֵ������ڻ���ڵ�����ʷ�����ֵ��
                            '����ʷ����:����>=3��������<=7 ,��ô����������Ϊ����<3��������>7��
                            If arrAge(2) = "����" And arrAge1(2) = "����" Then
                                If arrAge1(0) Like "С��*" Then
                                    If Val(arrAge1(1)) > Val(arrAge(1)) Or (Val(arrAge1(1)) = Val(arrAge(1)) And arrAge1(0) Like "*����" And arrAge(0) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                        Exit Function
                                    End If
                                    If Val(arrAge1(4)) < Val(arrAge(4)) Or (Val(arrAge1(4)) = Val(arrAge(4)) And arrAge1(3) Like "*����" And arrAge(3) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(1).Enabled And txtAge(1).Visible Then txtAge(1).SetFocus
                                        Exit Function
                                    End If
                                End If
                                If arrAge1(0) Like "����*" Then
                                    If Val(arrAge1(1)) < Val(arrAge(4)) Or (Val(arrAge1(1)) = Val(arrAge(4)) And arrAge1(0) Like "*����" And arrAge(3) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                        Exit Function
                                    End If
                                    If Val(arrAge1(4)) > Val(arrAge(1)) Or (Val(arrAge1(4)) = Val(arrAge(1)) And arrAge1(3) Like "*����" And arrAge(0) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(1).Enabled And txtAge(1).Visible Then txtAge(1).SetFocus
                                        Exit Function
                                    End If
                                End If
                            End If
                            '�����ʷ�����л��ߣ������ǲ��ң���ô�������䷶Χ��������ʷ���䷶Χ�м䡣
                            '����ʷ����:����<3��������>7��,��ô����������������>=3��������<=7�����Χ��
                            If arrAge(2) = "����" And arrAge1(2) = "����" Then
                                If arrAge(0) Like "С��*" Then
                                    If Val(arrAge1(1)) < Val(arrAge(1)) Or (Val(arrAge1(1)) = Val(arrAge(1)) And arrAge1(0) Like "*����" And arrAge(0) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                        Exit Function
                                    End If
                                    If Val(arrAge1(4)) > Val(arrAge(4)) Or (Val(arrAge1(4)) = Val(arrAge(4)) And arrAge1(3) Like "*����" And arrAge(3) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(1).Enabled And txtAge(1).Visible Then txtAge(1).SetFocus
                                        Exit Function
                                    End If
                                End If
                                If arrAge(0) Like "����*" Then
                                    If Val(arrAge1(1)) < Val(arrAge(4)) Or (Val(arrAge1(1)) = Val(arrAge(4)) And arrAge1(0) Like "*����" And arrAge(3) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                        Exit Function
                                    End If
                                    If Val(arrAge1(4)) > Val(arrAge(1)) Or (Val(arrAge1(4)) = Val(arrAge(1)) And arrAge1(3) Like "*����" And arrAge(0) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(1).Enabled And txtAge(1).Visible Then txtAge(1).SetFocus
                                        Exit Function
                                    End If
                                End If
                            End If
                        '��ʷ���ݰ������һ���ߣ���������ֻ��һ������
                        ElseIf UBound(arrAge) > 3 And UBound(arrAge1) < 3 Then
                            If arrAge(2) = "����" Then
                                MsgBox strInfo, vbInformation, gstrSysName
                                If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                Exit Function
                            End If
                            If arrAge(2) = "����" Then
                                If arrAge1(0) Like "С��*" Then
                                    If Val(arrAge1(1)) > Val(arrAge(1)) Or (Val(arrAge1(1)) = Val(arrAge(1)) And arrAge1(0) Like "*����" And arrAge(0) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                        Exit Function
                                    End If
                                End If
                                If arrAge1(0) Like "����*" Then
                                    If Val(arrAge1(1)) < Val(arrAge(4)) Or (Val(arrAge1(1)) = Val(arrAge(4)) And arrAge1(0) Like "*����" And arrAge(3) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                        Exit Function
                                    End If
                                End If
                            End If
                        '�������ݰ������һ���ߣ���ʷ����ֻ��һ������
                        ElseIf UBound(arrAge) < 3 And UBound(arrAge1) > 3 Then
                            If arrAge1(2) = "����" Then
                                MsgBox strInfo, vbInformation, gstrSysName
                                If cboAge(1).Enabled And cboAge(1).Visible Then cboAge(1).SetFocus
                                Exit Function
                            End If
                            If arrAge1(2) = "����" Then
                                If arrAge(0) Like "С��*" Then
                                    If Val(arrAge1(1)) < Val(arrAge(1)) Or (Val(arrAge1(1)) = Val(arrAge(1)) And arrAge1(0) Like "*����" And arrAge(0) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                        Exit Function
                                    End If
                                End If
                                If arrAge(0) Like "����*" Then
                                    If Val(arrAge1(4)) > Val(arrAge(1)) Or (Val(arrAge1(4)) = Val(arrAge(1)) And arrAge1(3) Like "*����" And arrAge(0) Like "*����") Then
                                        MsgBox strInfo, vbInformation, gstrSysName
                                        If txtAge(1).Enabled And txtAge(1).Visible Then txtAge(1).SetFocus
                                        Exit Function
                                    End If
                                End If
                            End If
                        '��������ʷ���ݶ�ֻ��һ�����������
                        Else
                            '��������ʷ���ݵĵ�ʽ���ܽ���,�磺������С�ڻ����
                            If Mid(arrAge(0), 1, 2) = Mid(arrAge1(0), 1, 2) Then
                                MsgBox strInfo, vbInformation, gstrSysName
                                If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                Exit Function
                            End If
                            If arrAge1(0) Like "С��*" Then
                                If Val(arrAge1(1)) > Val(arrAge(1)) Or (Val(arrAge1(1)) = Val(arrAge(1)) And arrAge1(0) Like "*����" And arrAge(0) Like "*����") Then
                                    MsgBox strInfo, vbInformation, gstrSysName
                                    If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                    Exit Function
                                End If
                            End If
                            If arrAge1(0) Like "����*" Then
                                If Val(arrAge1(1)) < Val(arrAge(1)) Or (Val(arrAge1(1)) = Val(arrAge(1)) And arrAge1(0) Like "*����" And arrAge(0) Like "*����") Then
                                    MsgBox strInfo, vbInformation, gstrSysName
                                    If txtAge(0).Enabled And txtAge(0).Visible Then txtAge(0).SetFocus
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
GoNext:
        Next lngRow
    End With
    
    IsValid = True
End Function

Private Function SaveData() As Boolean
'���ܣ���������ͬ��������Ϣ
    Dim lngIndex As Long
    On Error GoTo errHand
    
    '��һ�����������ݺϷ��Լ��
    If Not IsValid Then Exit Function
    '�ڶ�����������ݱ���
    
    '��ȡ��������
    T_ItemDate.strAgeFilter = GetFilter(1)
    T_ItemDate.intNursGrade = cboNursGrade.ItemData(cboNursGrade.ListIndex)
    '��Ŀ
    T_ItemDate.strItems = ""
    For lngIndex = 1 To lvwItem.ListItems.Count
        If lvwItem.ListItems(lngIndex).Checked = IIf(chkItem.Value = 0, True, False) Then
            T_ItemDate.strItems = T_ItemDate.strItems & ";" & lvwItem.ListItems(lngIndex).Text
        End If
    Next
    T_ItemDate.strItems = Mid(T_ItemDate.strItems, 2)
    '����
    T_ItemDate.strDept = -1
    If ChkAll.Value <> 1 Then
        T_ItemDate.strDept = ""
        For lngIndex = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(lngIndex).Checked = True Then
                T_ItemDate.strDept = T_ItemDate.strDept & ";" & Mid(lvwDept.ListItems(lngIndex).Key, 2)
            End If
        Next
        T_ItemDate.strDept = Mid(T_ItemDate.strDept, 2)
    End If
    If rptList.Tag = "�޸�" And rptList.Records.Count > 0 Then
        If Not rptList.FocusedRow Is Nothing Then
            If Not rptList.FocusedRow.GroupRow Then
                If Not DeleteData Then Exit Function
            End If
        End If
    End If
    '��������
    gstrSQL = "zl_����ͬ����Ŀ_Update("
'    ����ȼ�_IN ����ͬ����Ŀ.����ȼ�%TYPE,
    gstrSQL = gstrSQL & T_ItemDate.intNursGrade & ",'"
'    ���䷶Χ_IN ����ͬ����Ŀ.���䷶Χ%TYPE,
    gstrSQL = gstrSQL & T_ItemDate.strAgeFilter & "','"
'    ������Ŀ_IN ����ͬ����Ŀ.������Ŀ%TYPE,
    gstrSQL = gstrSQL & T_ItemDate.strItems & "','"
'    ���ÿ���_IN ����ͬ����Ŀ.���ÿ���%TYPE
    gstrSQL = gstrSQL & T_ItemDate.strDept & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "zl_����ͬ����Ŀ_Update")
    SaveData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DeleteData() As Boolean
    On Error GoTo errHand
    gstrSQL = "zl_����ͬ����Ŀ_Delete(" & Val(rptList.FocusedRow.Record(mCol.����ȼ�).Value) & ",'" & rptList.FocusedRow.Record(mCol.���䷶Χ).Record.Tag & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "zl_����ͬ����Ŀ_Delete")
    
    DeleteData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NursGradeSwitch(ByVal intNursGrade As Integer) As String
'����: ��ȡ����ȼ�
    Dim strTmp As String
    Select Case intNursGrade
        Case 0
            strTmp = "0-�ؼ�����"
        Case 1
            strTmp = "1-һ������"
        Case 2
            strTmp = "2-��������"
        Case 3
            strTmp = "3-��������"
        Case Else
            strTmp = "-1-���л���"
    End Select
    NursGradeSwitch = strTmp
End Function

Private Function NursItemSwitch(ByVal strItems As String) As String
'����:��ȡ��Ŀ��Ϣ
    Dim lngIndex As Long, lngCount As Long
    Dim strTmp As String
    
    Dim arrItems() As String
    If strItems = "" Then Exit Function
    arrItems = Split(strItems, ";")
    For lngCount = 0 To UBound(arrItems)
        For lngIndex = 1 To lvwItem.ListItems.Count
            If Val(arrItems(lngCount)) = Val(lvwItem.ListItems(lngIndex).Text) Then
                strTmp = strTmp & ";" & Val(arrItems(lngCount)) & "-" & lvwItem.ListItems(lngIndex).SubItems(1)
                Exit For
            End If
        Next lngIndex
    Next lngCount
    
    strTmp = Mid(strTmp, 2)
    NursItemSwitch = strTmp
End Function

Private Function DeptSwitch(ByVal strDept As String) As String
'���ܣ���ȡ������Ϣ
    Dim lngIndex As Long, lngCount As Long
    Dim strTmp As String
    
    Dim arrItems() As String
    If strDept = "" Then strDept = "-1"
    If Val(strDept) = -1 Then
        DeptSwitch = "-1-ȫԺͨ��"
        Exit Function
    End If
    arrItems = Split(strDept, ";")
    For lngCount = 0 To UBound(arrItems)
        For lngIndex = 1 To lvwDept.ListItems.Count
            If Val(arrItems(lngCount)) = Val(Mid(lvwDept.ListItems(lngIndex).Key, 2)) Then
                strTmp = strTmp & ";" & Val(arrItems(lngCount)) & "-" & lvwDept.ListItems(lngIndex).SubItems(1)
                Exit For
            End If
        Next lngIndex
    Next lngCount
    
    strTmp = Mid(strTmp, 2)
    DeptSwitch = strTmp
End Function

Private Sub RefreshStateInfo()
'------------------------------------------------------------------------------------------------------------------
'���ܣ�ˢ��״̬����ʾ��Ϣ
'-----------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngCount As Long
    
    For lngRow = 0 To rptList.Rows.Count - 1
        '������=0��Ϊ���������࣬������ͳ��
        If Not rptList.Rows(lngRow).GroupRow Then
            lngCount = lngCount + 1
        End If
    Next lngRow
    
    stbThis.Panels(2).Text = "���� " & lngCount & " ������ͬ��������Ϣ��"
End Sub

