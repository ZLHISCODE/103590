VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmQCTodayList 
   Caption         =   "�����ʿع���"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10470
   Icon            =   "frmQCTodayList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10470
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picLeft 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   90
      ScaleHeight     =   6750
      ScaleWidth      =   6060
      TabIndex        =   3
      Top             =   570
      Width           =   6060
      Begin VB.Frame fraNS 
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   -225
         MousePointer    =   7  'Size N S
         TabIndex        =   7
         Top             =   2430
         Width           =   3360
      End
      Begin VB.PictureBox PicList 
         BorderStyle     =   0  'None
         Height          =   4200
         Left            =   345
         ScaleHeight     =   4200
         ScaleWidth      =   5610
         TabIndex        =   4
         Top             =   120
         Width           =   5610
         Begin XtremeReportControl.ReportControl rptList 
            Height          =   3210
            Left            =   90
            TabIndex        =   5
            Top             =   90
            Width           =   5280
            _Version        =   589884
            _ExtentX        =   9313
            _ExtentY        =   5662
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgList 
            Height          =   900
            Left            =   0
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   3300
            Visible         =   0   'False
            Width           =   1080
            _cx             =   1905
            _cy             =   1587
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
         Begin MSComctlLib.ImageList imgList 
            Left            =   1245
            Top             =   3510
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmQCTodayList.frx":058A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmQCTodayList.frx":0B24
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
         Height          =   2295
         Left            =   720
         TabIndex        =   8
         Top             =   4605
         Width           =   4305
         _cx             =   7594
         _cy             =   4048
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         WordWrap        =   0   'False
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
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   5550
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   2115
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1845
   End
   Begin MSComCtl2.DTPicker dtp���� 
      Height          =   300
      Left            =   7860
      TabIndex        =   0
      Top             =   15
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   98566147
      CurrentDate     =   39110
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   7380
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCTodayList.frx":10BE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13388
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   90
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQCTodayList.frx":1950
      Left            =   840
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmQCTodayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ͼ�� = 0: �걾ID: �걾��:  ����id: ��������: �ʿ�Ʒid: �ʿ�Ʒ: ����: ˮƽ: ����
End Enum
Private Enum mColL
    ͼ�� = 0: ID: ������: Ӣ����: ���: ��ֵ: SD: ��λ: ���: ȡֵ����: ���ý��: ��Ŀid: �ʿ�Ʒid: ��ʼ����: ��������: ԭʼ���: �鵵��: ���
End Enum
Const conPane_List = 201
Const conPane_LJ = 202
Const conPane_Report = 203

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�

Private mfrmLJ As frmQCChartLJ
Private mfrmReport As frmQCTodayReport

Private mintEditState As Integer    '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�ʿؼ�¼�༭,2-����༭
Private mstrDate As String          '����
Private mlngRecord As Long          '����id
Private mlngResult As Long          '���id

Private mblnAllDev As Boolean      '�Ƿ�߱���������Ȩ�ޣ�����ֻ�ܴ������ŵ�����
Private mlngEditWidth As Long, mlngEditHeight As Long   '�༭����ĸ߶ȺͿ��

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrCustom As CommandBarControlCustom
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim RptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long
Dim mblnEdit As Boolean '�Ƿ�༭��

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Public Function zlRefList(Optional lngRecord As Long, Optional lngResult As Long) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ���ļ�¼
    Dim rsTemp As New ADODB.Recordset
    Dim strLists As String, strValue As String
    
    mstrDate = Format(Me.dtp����.Value, "yyyy-MM-dd")
    If Me.cbo����.Tag <> "" Then Exit Function
    Err = 0: On Error GoTo ErrHand
                
   gstrSql = " Select l.�걾id, l.�걾��� as �걾��, l.����id,m.���� as ����, x.���� as �ʿ�Ʒ, l.�ʿ�Ʒid, x.����, x.ˮƽ, l.���Դ��� as ���� " & vbNewLine & _
            " From �����ʿؼ�¼ l, �����ʿ�Ʒ x,�������� m " & vbNewLine & _
            " Where (l.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([1], 'yyyy-mm-dd') + 1 - 1 / 86400) And" & vbNewLine & _
            " l.����id = m.id And l.�ʿ�Ʒid = X.ID "
    
    '����
    If Me.cbo����.ItemData(Me.cbo����.ListIndex) > 0 Then
        gstrSql = gstrSql & " and  L.����id = [2] "
    End If
    gstrSql = gstrSql & " Order by L.�걾���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstrDate, Me.cbo����.ItemData(Me.cbo����.ListIndex))
    
    Me.rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
'          Select Case Val("" & !���)
'            Case 1
'                Set RptItem = rptRcd.AddItem("1"): RptItem.Icon = 0
'            Case 2
'                Set RptItem = rptRcd.AddItem("2"): RptItem.Icon = 1
'            Case Else
                Set RptItem = rptRcd.AddItem("")
'            End Select
            rptRcd.AddItem CLng(!�걾ID)
            rptRcd.AddItem CStr("" & !�걾��)
            rptRcd.AddItem CStr("" & !����id)
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !�ʿ�Ʒid)
            rptRcd.AddItem CStr("" & !�ʿ�Ʒ)
            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !ˮƽ)
            rptRcd.AddItem CStr("" & !����)
            .MoveNext
        Loop
    End With
    Me.rptList.Populate
    
    If lngRecord <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.�걾ID).Value) = lngRecord And lngRecord <> 0 Then
                    Set Me.rptList.FocusedRow = rptRow: If mlngResult = 0 Then Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    Call rptList_SelectionChanged
    zlRefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "����" & Me.rptList.Records.Count & "����¼"
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    'If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = Format(mstrDate, "yyyy��MM��dd��") & "�ʿؼ�¼�嵥"
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

Private Sub cbo����_Click()
    Dim rsTmp As New ADODB.Recordset
    
'    gstrSql = "Select id ,���� , ���� From �������� a Where ʹ��С��id = [1] order by ���� "
    gstrSql = "Select ID, ����, ����" & vbNewLine & _
            " From �������� A" & vbNewLine & _
            " Where ʹ��С��id = [1] And" & vbNewLine & _
            "      A.ID In (Select Distinct D.ID" & vbNewLine & _
            "               From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
            "               Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [2] And C.����id = D.ID)" & vbNewLine & _
            " Order By ����"


    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, Me.cbo����.ItemData(Me.cbo����.ListIndex), UserInfo.ID)
    Me.cbo����.Clear
    If InStr(1, mstrPrivs, "���п���") > 0 Then
        Me.cbo����.AddItem "��������"
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = 0
    End If
    
    Do Until rsTmp.EOF
        Me.cbo����.AddItem rsTmp("����")
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    Me.cbo����.ListIndex = 0

    
End Sub

Private Sub cbo����_Click()
    Call zlRefList
End Sub

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long, strInfo As String
    
    '------------------------------------
    Select Case Control.ID
'    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_Save
        lngRetuId = 0
        Select Case mintEditState   '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�ʿؼ�¼�༭,2-����༭
        Case 1
            lngRetuId = zlEditSave()
            If lngRetuId <> 0 Then
                mlngRecord = lngRetuId ':  Call zlRefList(mlngRecord)
                
                mintEditState = 0: Me.dtp����.Enabled = True
                Me.cbo����.Enabled = True: Me.cbo����.Enabled = True: Me.PicList.Enabled = True
                vfgRecord.Editable = flexEDKbd
                mblnEdit = False
                vfgRecord.SelectionMode = flexSelectionByRow
                mlngResult = 0: Call vfgRecord_RowColChange
            End If
            
        Case 2:
            lngRetuId = mfrmReport.zlEditSave()
            If lngRetuId <> 0 Then
                mlngResult = lngRetuId:  Call zlRefList(mlngRecord, mlngResult)
                mintEditState = 0: Me.dtp����.Enabled = True: Me.picLeft.Enabled = True: Me.rptList.SetFocus
            End If
        End Select
    Case conMenu_Edit_Untread:
        Select Case mintEditState   '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�ʿؼ�¼�༭,2-����༭
        Case 1
            If mblnEdit Then
                If MsgBox("�Ƿ�����������޸ģ�", vbInformation + vbOKCancel, Me.Caption) = vbCancel Then
                    Exit Sub
                Else
                    With vfgRecord
                        For lngRetuId = .FixedRows To .Rows - 1
                            If .TextMatrix(lngRetuId, mColL.���) <> .TextMatrix(lngRetuId, mColL.ԭʼ���) Then
                                .TextMatrix(lngRetuId, mColL.���) = .TextMatrix(lngRetuId, mColL.ԭʼ���)
                            End If
                        Next
                    End With
                End If
                mblnEdit = False
            End If
            vfgRecord.SelectionMode = flexSelectionByRow
            Me.PicList.Enabled = True
        Case 2: Call mfrmReport.zlEditCancel
             Me.picLeft.Enabled = True
        End Select
        
        mintEditState = 0: Me.dtp����.Enabled = True
        Me.cbo����.Enabled = True: Me.cbo����.Enabled = True
        
    Case conMenu_Edit_NewItem
        
        If frmQCTodayRecord.ZlEditStart(True, mlngRecord, mstrDate, mblnAllDev) <> 0 Then
            Call zlRefList
        End If
    Case conMenu_Edit_Modify

        mintEditState = 1: Me.dtp����.Enabled = False: Me.PicList.Enabled = False
        Me.cbo����.Enabled = False: Me.cbo����.Enabled = False
        vfgRecord.Editable = flexEDKbdMouse
        vfgRecord.SelectionMode = flexSelectionFree
    Case conMenu_Edit_Delete
        If mlngRecord = 0 Then Exit Sub
        With Me.rptList
            strInfo = "���Ҫɾ���ñ걾���ʿصǼǣ���ԭΪ��ͨ�걾��" & vbCrLf
            strInfo = strInfo & vbCrLf & "    �� �� �ţ�" & .FocusedRow.Record(mCol.�걾��).Value
            strInfo = strInfo & vbCrLf & "    ����������" & .FocusedRow.Record(mCol.��������).Value
            
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_�����ʿؼ�¼_Edit(3," & mlngRecord & ")"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

            Err = 0: On Error GoTo 0
            mlngRecord = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                If .Rows(lngRetuId + 1).GroupRow = False Then mlngRecord = .Rows(lngRetuId + 1).Record(mCol.�걾ID).Value
            ElseIf lngRetuId > 0 Then
                If .Rows(lngRetuId - 1).GroupRow = False Then mlngRecord = .Rows(lngRetuId - 1).Record(mCol.�걾ID).Value
            End If
            Call Me.zlRefList(mlngRecord)
        End With
    Case conMenu_Edit_Adjust                '��дʧ�ر���
        If mfrmReport.ZlEditStart(mlngResult) Then
            mintEditState = 2: Me.dtp����.Enabled = False: Me.picLeft.Enabled = False
        End If
    Case conMenu_Edit_Archive
        With Me.rptList
            strInfo = strInfo & vbCrLf & "    �� �� �ţ�" & .FocusedRow.Record(mCol.�걾��).Value
            strInfo = strInfo & vbCrLf & "    ����������" & .FocusedRow.Record(mCol.��������).Value
            strInfo = strInfo & vbCrLf & "    ������Ŀ��" & Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.������)
        End With
        
        If Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.�鵵��) = "" Then
            strInfo = "���Ҫ����ǰʧ�ر���鵵��" & vbCrLf & strInfo
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_�����ʿر���_Archive(" & mlngResult & ",0)"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.�鵵��) = UserInfo.����
        Else
            strInfo = "��ʧ�ر����Ѿ��鵵�����ȡ���鵵��" & vbCrLf & strInfo
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_�����ʿر���_Archive(" & mlngResult & ",1)"
            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Me.vfgRecord.TextMatrix(Me.vfgRecord.Row, mColL.�鵵��) = ""
        End If
        Call Me.zlRefList(mlngRecord, mlngResult)
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh        'ˢ��
        Call zlRefList(mlngRecord)
    
    Case conMenu_Tool_Analyse '����
        With Me.vfgRecord
            lngRetuId = Val("" & .TextMatrix(.Row, mColL.��Ŀid))
        End With
        If lngRetuId <= 0 Then
            MsgBox "��ѡ��һ����Ŀ��ʹ�ô˹��ܣ�", vbInformation, Me.Caption
            Exit Sub
        End If
        
        With Me.rptList
            If frmQCCompute.ShowMe(Me, _
                .FocusedRow.Record(mCol.����id).Value, lngRetuId, _
                CDate(mstrDate), .FocusedRow.Record(mCol.�ʿ�Ʒid).Value) Then
                Call Me.zlRefList(mlngRecord)
            End If
        End With
    Case conMenu_Tool_Define '��ֵ
        With Me.vfgRecord
            lngRetuId = Val("" & .TextMatrix(.Row, mColL.��Ŀid))
        End With
        If lngRetuId <= 0 Then
            MsgBox "��ѡ��һ����Ŀ��ʹ�ô˹��ܣ�", vbInformation, Me.Caption
            Exit Sub
        End If
        With Me.rptList
            If .FocusedRow Is Nothing Then
                MsgBox "��ѡ��һ���ʿ�Ʒ����ʹ�ô˹��ܣ�", vbInformation, Me.Caption
                Exit Sub
            End If
            If frmQCRedefine.ShowMe(Me, _
                .FocusedRow.Record(mCol.����id).Value, lngRetuId, _
                CDate(mstrDate), .FocusedRow.Record(mCol.�ʿ�Ʒid).Value) Then
                Call Me.zlRefList(mlngRecord)
            End If
        End With
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else

            If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub

            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread: Control.Enabled = (mintEditState <> 0)
    
    Case conMenu_Edit_NewItem: Control.Enabled = (InStr(1, mstrPrivs, "�Ǽ�") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "�Ǽ�") > 0 And mintEditState = 0 And mlngRecord <> 0)
        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.�鵵��).Value) = "")
    Case conMenu_Edit_Adjust
        Control.Enabled = (InStr(1, mstrPrivs, "����") > 0 And mintEditState = 0 And mlngResult <> 0)
        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Val(Me.rptList.FocusedRow.Record(mCol.ͼ��).Value) <> 0)
'        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.�鵵��).Value) = "")
    Case conMenu_Edit_Archive
        Control.Enabled = (InStr(1, mstrPrivs, "�鵵") > 0 And mintEditState = 0 And mlngResult <> 0)
        If Control.Enabled = False Then Exit Sub
'        Control.Enabled = (Trim(Me.rptList.FocusedRow.Record(mCol.������).Value) <> "")
    
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    
    Case conMenu_Tool_Analyse
        Control.Enabled = (InStr(1, mstrPrivs, "����") > 0 And mintEditState = 0 And mlngResult <> 0)
    Case conMenu_Tool_Define
        Control.Enabled = (InStr(1, mstrPrivs, "��ֵ") > 0 And mintEditState = 0 And mlngResult <> 0)
    End Select
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picLeft.hWnd
    Case conPane_LJ
        Item.Handle = mfrmLJ.hWnd
    Case conPane_Report
        Item.Handle = mfrmReport.hWnd
    End Select
End Sub

Private Sub dtp����_Change()
    Call zlRefList
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    
'    mstrPrivs = gstrPrivs
     '�����м�վҪֱ�����������������һ�½ű�
    gstrPrivs = GetPrivFunc(100, 1210)
    mstrPrivs = gstrPrivs
    mblnAllDev = IIf(InStr(1, mstrPrivs, "���п���") = 0, False, True)
    Me.cbo����.Tag = "��ˢ��"
    Me.dtp����.Value = Date: Me.dtp����.MaxDate = Date
    mstrDate = Format(Date, "yyyy-MM-dd"): mlngRecord = 0: mlngResult = 0
    mintEditState = 0
    
    mlngEditWidth = Me.picLeft.Width
'    mlngEditHeight = frmQCTodayRecord.Height
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�Ǽ�(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "����(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵(&T)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    cbrMenuBar.ID = xtpControlPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "ʧ�ؼ���(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "���¶�ֵ(&N)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "����")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "����")
    cbrCustom.Handle = Me.cbo����.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "����")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "����")
    cbrCustom.Handle = Me.cbo����.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��������")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "��������")
    cbrCustom.Handle = Me.dtp����.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("Y"), conMenu_Tool_Analyse
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbsThis, glngSys, glngModul, mstrPrivs)
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�Ǽ�"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "ʧ�ؼ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "���¶�ֵ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Set mfrmLJ = New frmQCChartLJ
    Set mfrmReport = New frmQCTodayReport
    
    Dim panThis As Pane, panSub As Pane
    Set panThis = dkpMan.CreatePane(conPane_List, 350, 400, DockLeftOf, Nothing)
    panThis.Title = "�����ʿؼ�¼"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_LJ, 400, 600, DockRightOf, Nothing)
    panThis.Title = "�����ʿ�ͼ��"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panSub = dkpMan.CreatePane(conPane_Report, 400, 200, DockBottomOf, panThis)
    panSub.Title = "����ʧ�ر���"
    panSub.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        .SetImageList Me.imgList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 1024)   '������������֮ǰ���ã�������Ч
        .AllowColumnRemove = False
        .AllowEdit = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False):  rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.�걾ID, "�걾ID", 0, False):  rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�걾��, "�걾��", 62, False):  rptCol.Groupable = False: .SortOrder.Add rptCol
        
        Set rptCol = .Columns.Add(mCol.����id, "����id", 0, False):  rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��������, "��������", 120, True):  rptCol.Groupable = True: rptCol.Visible = False: .GroupsOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.�ʿ�Ʒid, "�ʿ�Ʒid", 0, False):  rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�ʿ�Ʒ, "�ʿ�Ʒ", 160, True):  rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����, "����", 160, True):   rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.ˮƽ, "ˮƽ", 30, False):  rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.����, "����", 30, False):  rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        .Populate
    End With
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    
    'װ���������
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand
    
    If InStr(1, mstrPrivs, "���п���") > 0 Then
        gstrSql = " Select Distinct b.Id, b.���� , b.���� As ���� From �������� a ,���ű� b,�����ʿ�Ʒ c " & _
                  "Where a.ʹ��С��ID = b.ID and a.id = c.����id order by b.���� "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName)
        
    Else

        gstrSql = "Select Distinct B.ID, B.����, B.���� As ����" & vbNewLine & _
                " From �������� A, ���ű� B, �����ʿ�Ʒ C" & vbNewLine & _
                " Where A.ʹ��С��id = B.ID And A.ID = C.����id And" & vbNewLine & _
                "      A.ʹ��С��id In (Select Distinct D.ʹ��С��id" & vbNewLine & _
                "                   From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
                "                   Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [1] And C.����id = D.ID)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, UserInfo.ID)
    End If
    
    Me.cbo����.Clear
    If InStr(1, mstrPrivs, "���п���") > 0 Then
        Me.cbo����.AddItem "���п���"
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = 0
    End If
    Do Until rsTemp.EOF
        Me.cbo����.AddItem rsTemp("����") & "-" & rsTemp("����")
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = rsTemp("Id")
        rsTemp.MoveNext
    Loop
    If Me.cbo����.ListCount = 0 Then MsgBox "��δ�������ʹ��С������ã�", vbInformation, gstrSysName: Unload Me: Exit Sub
    Me.cbo����.ListIndex = 0
    If Me.cbo����.ListCount = 1 Then Me.cbo����.Enabled = False
    Me.cbo����.Tag = ""
    '����װ��
    Call zlRefList
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim panKind As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panKind = Me.dkpMan.FindPane(conPane_List)
'    panKind.MinTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
'    panKind.MaxTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
'    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panKind.MinTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, mlngEditHeight / Screen.TwipsPerPixelY
    panKind.MaxTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
    Me.dkpMan.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmLJ
    Unload mfrmReport
    Set mfrmLJ = Nothing
    Set mfrmReport = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub fraNS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.fraNS.Top = Me.fraNS.Top + y
        Me.PicList.Height = Me.PicList.Height + y
        Me.vfgRecord.Top = Me.vfgRecord.Top + y
        Me.vfgRecord.Height = Me.vfgRecord.Height - y
    End If
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With Me.fraNS
        .Left = Me.picLeft.ScaleLeft: .Width = Me.picLeft.ScaleWidth - .Left
    End With
    With Me.vfgRecord
        .Left = Me.picLeft.ScaleLeft: .Width = Me.picLeft.ScaleWidth - .Left
        .Top = Me.fraNS.Top + Me.fraNS.Height
        .Height = Me.picLeft.ScaleHeight - .Top
    End With
    With Me.PicList
        .Left = Me.picLeft.ScaleLeft: .Width = Me.picLeft.ScaleWidth - .Left
        .Top = Me.picLeft.ScaleTop
        .Height = Me.picLeft.ScaleHeight - Me.vfgRecord.Height - Me.fraNS.Height
    End With
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = Me.PicList.ScaleLeft: .Width = Me.PicList.ScaleWidth - .Left
        .Top = Me.PicList.ScaleTop
        .Height = Me.picLeft.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_SelectionChanged()
    Dim lng�ʿ�Ʒid As Long, str���� As String

    str���� = Format(dtp����.Value, "yyyy-MM-dd")
    If Me.rptList.FocusedRow Is Nothing Then
        mlngRecord = 0: mlngResult = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngRecord = 0: mlngResult = 0
    Else
        mlngRecord = Me.rptList.FocusedRow.Record.Item(mCol.�걾ID).Value
        lng�ʿ�Ʒid = Me.rptList.FocusedRow.Record.Item(mCol.�ʿ�Ʒid).Value
'        mlngResult = Me.rptList.FocusedRow.Record.Item(mCol.���ID).Value
    End If

    Call LoadRecord(str����, lng�ʿ�Ʒid, mlngRecord)
    
End Sub


Private Sub LoadRecord(ByVal str���� As String, lng�ʿ�Ʒid As Long, lng�걾ID)
    '��ʾ�ʿر걾��ϸ
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    On Error GoTo errH

    With Me.vfgRecord
        .Redraw = flexRDNone
        .Clear
        .Cols = 18
        .Rows = .FixedRows
        .TextMatrix(0, mColL.ͼ��) = "": .ColWidth(mColL.ͼ��) = 180: .FixedAlignment(mColL.ͼ��) = flexAlignGeneralCenter
        .TextMatrix(0, mColL.ID) = "": .ColWidth(mColL.ID) = 0: .ColHidden(mColL.ID) = True
        .TextMatrix(0, mColL.������) = "������": .ColWidth(mColL.������) = 1500: .FixedAlignment(mColL.������) = flexAlignLeftCenter
        .TextMatrix(0, mColL.Ӣ����) = "Ӣ����": .ColWidth(mColL.Ӣ����) = 800: .FixedAlignment(mColL.Ӣ����) = flexAlignLeftCenter
        .TextMatrix(0, mColL.���) = "���ֵ": .ColWidth(mColL.���) = 900: .FixedAlignment(mColL.���) = flexAlignRightCenter
        .TextMatrix(0, mColL.��ֵ) = "��ֵ": .ColWidth(mColL.��ֵ) = 800: .FixedAlignment(mColL.��ֵ) = flexAlignRightCenter
        .TextMatrix(0, mColL.SD) = "SD": .ColWidth(mColL.SD) = 800: .FixedAlignment(mColL.SD) = flexAlignRightCenter
        .TextMatrix(0, mColL.��λ) = "��λ": .ColWidth(mColL.��λ) = 900: .FixedAlignment(mColL.��λ) = flexAlignLeftCenter
        .TextMatrix(0, mColL.���) = "���": .ColWidth(mColL.���) = 0: .ColHidden(mColL.���) = True
        .TextMatrix(0, mColL.ȡֵ����) = "ȡֵ����": .ColWidth(mColL.ȡֵ����) = 0: .ColHidden(mColL.ȡֵ����) = True
        .TextMatrix(0, mColL.���ý��) = "���ý��": .ColWidth(mColL.���ý��) = 0: .ColHidden(mColL.���ý��) = True
        .TextMatrix(0, mColL.��Ŀid) = "��Ŀid": .ColWidth(mColL.��Ŀid) = 0: .ColHidden(mColL.��Ŀid) = True
        .TextMatrix(0, mColL.�ʿ�Ʒid) = "�ʿ�Ʒid": .ColWidth(mColL.�ʿ�Ʒid) = 0: .ColHidden(mColL.�ʿ�Ʒid) = True
        .TextMatrix(0, mColL.��ʼ����) = "��ʼ����": .ColWidth(mColL.��ʼ����) = 0: .ColHidden(mColL.��ʼ����) = True
        .TextMatrix(0, mColL.��������) = "��������": .ColWidth(mColL.��������) = 0: .ColHidden(mColL.��������) = True
        .TextMatrix(0, mColL.ԭʼ���) = "ԭʼ���": .ColWidth(mColL.ԭʼ���) = 0: .ColHidden(mColL.ԭʼ���) = True
        .TextMatrix(0, mColL.�鵵��) = "�鵵��": .ColWidth(mColL.�鵵��) = 0: .ColHidden(mColL.�鵵��) = True
        .TextMatrix(0, mColL.���) = "���": .ColWidth(mColL.���) = 0: .ColHidden(mColL.���) = True
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        strsql = "Select r.id,r.������Ŀid, Nvl(f.���, 0) As ���, i.������, i.Ӣ����, r.������, x.��ֵ, x.Sd, i.��λ," & vbNewLine & _
                "            Decode(p.�������, 3, p.ȡֵ����, '') As ȡֵ����, decode(p.�������,Null,i.����,p.�������) As ���,Nvl(r.���ý��, 0) As ���ý��," & vbNewLine & _
                "            x.�ʿ�Ʒid,x.��ʼ����,x.��������,F.�鵵��,F.������ " & vbNewLine & _
                "From ������ͨ��� r, �����ʿر��� f," & vbNewLine & _
                "        (Select x.�ʿ�Ʒid,x.��Ŀid, x.��ֵ, x.Sd,x.��ʼ����,nvl(x.��������,M.��������) as �������� " & vbNewLine & _
                "            From �����ʿؾ�ֵ x,�����ʿ�Ʒ M " & vbNewLine & _
                "            Where x.�ʿ�Ʒid=M.id And x.�ʿ�Ʒid = [2] And To_Date([1], 'yyyy-mm-dd') Between x.��ʼ���� And Nvl(x.��������, Sysdate)) x," & vbNewLine & _
                "        ����������Ŀ i, ������Ŀ p" & vbNewLine & _
                "Where Nvl(r.���ý��, 0) = 0 And r.Id = f.���id(+) And r.����걾id = [3] And r.������Ŀid = x.��Ŀid(+) And" & vbNewLine & _
                "           r.������Ŀid = i.Id And r.������Ŀid = p.������Ŀid" & vbNewLine & _
                "Order By decode(p.�������,Null,i.����,p.�������)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, str����, lng�ʿ�Ʒid, lng�걾ID)
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mColL.ID) = Val("" & rsTmp!ID)
            .TextMatrix(.Rows - 1, mColL.������) = Trim("" & rsTmp!������)
            .TextMatrix(.Rows - 1, mColL.Ӣ����) = Trim("" & rsTmp!Ӣ����)
            .TextMatrix(.Rows - 1, mColL.���) = IIf(Left(Trim("" & rsTmp!������), 1) = ".", "0" & Trim("" & rsTmp!������), Trim("" & rsTmp!������))
            .TextMatrix(.Rows - 1, mColL.ԭʼ���) = .TextMatrix(.Rows - 1, mColL.���)
            .TextMatrix(.Rows - 1, mColL.��ֵ) = IIf(Left(Trim("" & rsTmp!��ֵ), 1) = ".", "0" & Trim("" & rsTmp!��ֵ), Trim("" & rsTmp!��ֵ))
            .TextMatrix(.Rows - 1, mColL.SD) = IIf(Left(Trim("" & rsTmp!SD), 1) = ".", "0" & Trim("" & rsTmp!SD), Trim("" & rsTmp!SD))
            .TextMatrix(.Rows - 1, mColL.��λ) = Trim("" & rsTmp!��λ)
            .TextMatrix(.Rows - 1, mColL.���) = Trim("" & rsTmp!���)
            .TextMatrix(.Rows - 1, mColL.ȡֵ����) = Trim("" & rsTmp!ȡֵ����)
            .TextMatrix(.Rows - 1, mColL.���ý��) = Trim("" & rsTmp!���ý��)
            .TextMatrix(.Rows - 1, mColL.��Ŀid) = Val("" & rsTmp!������Ŀid)
            .TextMatrix(.Rows - 1, mColL.�ʿ�Ʒid) = Val("" & rsTmp!�ʿ�Ʒid)
            .TextMatrix(.Rows - 1, mColL.��ʼ����) = Trim(Format("" & rsTmp!��ʼ����, "yyyy-MM-dd"))
            .TextMatrix(.Rows - 1, mColL.��������) = Trim(Format("" & rsTmp!��������, "yyyy-MM-dd"))
            .TextMatrix(.Rows - 1, mColL.�鵵��) = Trim("" & rsTmp!�鵵��)
            .TextMatrix(.Rows - 1, mColL.���) = Trim("" & rsTmp!���)
            If rsTmp!��� <> 0 Then
                .Cell(flexcpBackColor, .Rows - 1, mColL.���) = &HC0C0FF
                .Cell(flexcpFontBold, .Rows - 1, mColL.���) = True
            End If
            rsTmp.MoveNext
        Loop
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then .Row = .FixedRows: .Col = 0
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vfgRecord_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vfgRecord
        If mblnEdit = False Then
            If Trim(.TextMatrix(Row, mColL.���)) <> Trim(.TextMatrix(Row, mColL.ԭʼ���)) Then
                mblnEdit = True
            End If
        End If
    End With
End Sub

Private Sub vfgRecord_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mintEditState <> 1 Then
        Cancel = True
        Exit Sub
    End If
    If Col <> mColL.��� Then
        Cancel = True
        Exit Sub
    End If
    If Row < vfgRecord.FixedRows Then
        Cancel = True
        Exit Sub
    End If
    
End Sub

Private Sub vfgRecord_DblClick()
   If mlngRecord = 0 Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub vfgRecord_RowColChange()
    Dim lng��ĿID As Long, lng�ʿ�Ʒid As Long, str�ʿ��ڼ� As String, str��ʼ���� As String, str�������� As String
    If Me.cbo����.Tag <> "" Then Exit Sub
    If mintEditState <> 0 Then Exit Sub
    With vfgRecord
        If mlngResult <> Val(.TextMatrix(.Row, mColL.ID)) And Val(.TextMatrix(.Row, mColL.ID)) <> 0 Then
            mlngResult = Val(.TextMatrix(.Row, mColL.ID))
            lng��ĿID = Val(.TextMatrix(.Row, mColL.��Ŀid))
            lng�ʿ�Ʒid = Val(.TextMatrix(.Row, mColL.�ʿ�Ʒid))
            
            str��ʼ���� = Format(dtp����.Value, "yyyy-MM") & "-01"
            str�������� = Format(DateAdd("m", 1, CDate(Format(dtp����.Value, "yyyy-MM") & "-01")) - 1, "yyyy-MM-dd")
            
            str�ʿ��ڼ� = lng�ʿ�Ʒid & "=" & .TextMatrix(.Row, mColL.��ʼ����) & "," & .TextMatrix(.Row, mColL.��������)
                    
            Call mfrmReport.zlRefresh(mlngResult)
            Call mfrmLJ.zlRefresh(CStr(lng�ʿ�Ʒid), lng��ĿID, str��ʼ����, str��������, str�ʿ��ڼ�)
            On Error Resume Next
            .SetFocus
            .Select .Row, mColL.���
            
        End If
    End With
End Sub

Private Function zlEditSave() As Long
    '�����޸Ľ��
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim lng����ID As Long, int�Ա� As Integer, str����  As String
    Dim strItem As String, intRow As Integer, lng��ĿID As Long
    
    If mblnEdit = False Then Exit Function
    strItem = ""
    With Me.vfgRecord
        For intRow = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(intRow, mColL.���)) <> Trim(.TextMatrix(intRow, mColL.ԭʼ���)) Then
                lng��ĿID = Val(.TextMatrix(intRow, mColL.��Ŀid))
                If lng��ĿID <> 0 Then
                     .TextMatrix(intRow, mColL.ԭʼ���) = Trim(.TextMatrix(intRow, mColL.���))
                    strItem = strItem & "|" & lng��ĿID & "^" & Trim(.TextMatrix(intRow, mColL.���))
                End If
            End If
        Next
    End With
    If strItem <> "" Then
        strItem = Mid(strItem, 2)
        
        strsql = "Select ����id,�걾����,Decode(�Ա�,'��',1,'Ů',2,0) as �Ա�,to_char(��������,'yyyy-MM-dd') as ���� From ����걾��¼ where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngRecord)
        Do Until rsTmp.EOF
            '����걾id_In ,����id_In ,�걾����_In,�Ա�_in,��������_in,����ָ��_in(��ĿID^ֵ|������) ,[΢����_in],[ø���id_in]
            strsql = "Zl_������ͨ���_Batchupdate(" & mlngRecord & "," & rsTmp!����id & ",'" & rsTmp!�걾���� & "'," & rsTmp!�Ա� & _
                     IIf(Trim("" & rsTmp!����) = "", ",Null", ",To_Date('" & rsTmp!���� & "','yyyy-MM-dd')") & ",'" & strItem & "')"
            zlDatabase.ExecuteProcedure strsql, Me.Caption
            rsTmp.MoveNext
        Loop
        zlEditSave = mlngRecord
    End If
    
    
End Function


