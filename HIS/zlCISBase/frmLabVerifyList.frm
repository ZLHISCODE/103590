VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmLabVerifyList 
   Caption         =   "��������"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   12600
   Icon            =   "frmLabVerifyList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   12600
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7410
      Left            =   4500
      ScaleHeight     =   7410
      ScaleWidth      =   7995
      TabIndex        =   26
      Top             =   105
      Width           =   8000
      Begin VB.Frame fraRule 
         Caption         =   "������"
         Height          =   2300
         Left            =   105
         TabIndex        =   33
         Top             =   60
         Width           =   7800
         Begin VB.ComboBox cboValid 
            Height          =   300
            Left            =   4005
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   570
            Width           =   3645
         End
         Begin VB.TextBox txt��ע 
            Height          =   900
            Left            =   4005
            MaxLength       =   200
            TabIndex        =   7
            Top             =   1260
            Width           =   3645
         End
         Begin VB.TextBox txt��Ŀ 
            Height          =   300
            Left            =   690
            TabIndex        =   4
            ToolTipText     =   "��DEL�������Ŀ"
            Top             =   915
            Width           =   2500
         End
         Begin VB.CommandButton cmd��Ŀ 
            Caption         =   "��"
            Height          =   300
            Left            =   3195
            TabIndex        =   42
            Top             =   915
            Width           =   300
         End
         Begin VB.TextBox txt���� 
            Height          =   285
            Left            =   4005
            MaxLength       =   3
            TabIndex        =   1
            Top             =   240
            Width           =   3660
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   4005
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   915
            Width           =   3660
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   690
            MaxLength       =   30
            TabIndex        =   2
            Top             =   578
            Width           =   2800
         End
         Begin VB.TextBox txtInfo 
            Height          =   900
            Left            =   690
            MaxLength       =   200
            TabIndex        =   6
            Top             =   1275
            Width           =   2820
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            ItemData        =   "frmLabVerifyList.frx":6852
            Left            =   690
            List            =   "frmLabVerifyList.frx":6854
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2820
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ч"
            Height          =   180
            Left            =   3585
            TabIndex        =   45
            Top             =   630
            Width           =   360
         End
         Begin VB.Label Label14 
            Caption         =   "��ע"
            Height          =   165
            Left            =   3600
            TabIndex        =   43
            Top             =   1275
            Width           =   420
         End
         Begin VB.Label Label10 
            Caption         =   "����"
            Height          =   165
            Left            =   3585
            TabIndex        =   41
            Top             =   285
            Width           =   780
         End
         Begin VB.Label Label1 
            Caption         =   "����"
            Height          =   165
            Left            =   3585
            TabIndex        =   38
            Top             =   975
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "����"
            Height          =   165
            Left            =   270
            TabIndex        =   37
            Top             =   630
            Width           =   420
         End
         Begin VB.Label Label3 
            Caption         =   "��ʾ"
            Height          =   165
            Left            =   285
            TabIndex        =   36
            Top             =   1275
            Width           =   420
         End
         Begin VB.Label Label11 
            Caption         =   "��Ŀ"
            Height          =   165
            Left            =   285
            TabIndex        =   35
            Top             =   975
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "����"
            Height          =   165
            Left            =   270
            TabIndex        =   34
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.Frame fraWhere 
         Caption         =   "��������"
         Height          =   1365
         Left            =   90
         TabIndex        =   27
         Top             =   2490
         Width           =   7800
         Begin VB.CheckBox chk���� 
            Alignment       =   1  'Right Justify
            Caption         =   "����"
            Height          =   195
            Left            =   6790
            TabIndex        =   15
            Top             =   615
            Width           =   800
         End
         Begin VB.ComboBox cbo���䵥λ 
            Height          =   300
            Left            =   4740
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   195
            Width           =   800
         End
         Begin VB.ComboBox cbo�Ա� 
            Height          =   300
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   195
            Width           =   1200
         End
         Begin VB.TextBox txt�������� 
            Height          =   285
            Left            =   2850
            MaxLength       =   9
            TabIndex        =   9
            Top             =   195
            Width           =   800
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   577
            Width           =   2600
         End
         Begin VB.ComboBox cbo�������� 
            Height          =   300
            Left            =   4740
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   577
            Width           =   1800
         End
         Begin VB.TextBox txt��� 
            Height          =   285
            Left            =   1065
            MaxLength       =   500
            TabIndex        =   16
            ToolTipText     =   "��DEL��������"
            Top             =   960
            Width           =   6555
         End
         Begin VB.TextBox txt�������� 
            Height          =   285
            Left            =   3855
            MaxLength       =   9
            TabIndex        =   10
            Top             =   195
            Width           =   800
         End
         Begin VB.CheckBox chk��ֹ 
            Alignment       =   1  'Right Justify
            Caption         =   "���Ϲ����ֹ���"
            Height          =   195
            Left            =   5820
            TabIndex        =   12
            Top             =   225
            Width           =   1770
         End
         Begin VB.Label Label4 
            Caption         =   "�Ա�"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   255
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "����          ��"
            Height          =   165
            Left            =   2415
            TabIndex        =   31
            Top             =   255
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "�ͼ����"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   630
            Width           =   780
         End
         Begin VB.Label Label7 
            Caption         =   "��������"
            Height          =   225
            Left            =   3915
            TabIndex        =   29
            Top             =   645
            Width           =   810
         End
         Begin VB.Label Label13 
            Caption         =   "�ٴ����"
            Height          =   165
            Left            =   240
            TabIndex        =   28
            Top             =   1005
            Width           =   780
         End
      End
      Begin VB.TextBox txtRule 
         Height          =   2625
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   4335
         Width           =   3880
      End
      Begin VB.CommandButton cmdSetEspecial 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   6780
         TabIndex        =   20
         Top             =   7005
         Width           =   1100
      End
      Begin VB.TextBox txtEspecial 
         Height          =   2610
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   4335
         Width           =   3880
      End
      Begin VB.CommandButton cmdSetRule 
         Caption         =   "�༭(&E)"
         Height          =   350
         Left            =   105
         TabIndex        =   17
         Top             =   6990
         Width           =   1100
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "AND"
         Height          =   315
         Left            =   3480
         TabIndex        =   18
         Top             =   3945
         Width           =   720
      End
      Begin VB.OptionButton optOr 
         Caption         =   "OR"
         Height          =   315
         Left            =   4170
         TabIndex        =   19
         Top             =   3945
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "��ͨ���� "
         Height          =   225
         Left            =   2490
         TabIndex        =   40
         Top             =   4005
         Width           =   810
      End
      Begin VB.Label Label9 
         Caption         =   "�������"
         Height          =   225
         Left            =   4830
         TabIndex        =   39
         Top             =   4005
         Width           =   870
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   900
      Left            =   180
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5625
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
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   255
      ScaleHeight     =   5295
      ScaleWidth      =   3540
      TabIndex        =   24
      Top             =   735
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4590
         Left            =   15
         TabIndex        =   0
         Top             =   60
         Width           =   3390
         _Version        =   589884
         _ExtentX        =   5980
         _ExtentY        =   8096
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   90
         Top             =   4800
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
               Picture         =   "frmLabVerifyList.frx":6856
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   7605
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabVerifyList.frx":2E630
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17145
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
      Left            =   465
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabVerifyList.frx":2EEC2
      Left            =   2100
      Top             =   165
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabVerifyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    ͼ�� = 0: ����: ID: ����: ����: ��Ŀ: ��Ŀid: ����: ����Id: ����ID: ��������: �Ա�: ��������: ��������: ���䵥λ: ���: ����: �������: �����ϵ: ��ʾ��Ϣ: ��Ч: ���: ��ע
End Enum
Const conPane_List = 201
Const conPane_Edit = 202

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�

Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim mLngEditWidth As Long

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim mlngItemID As Long
Dim mintEditState As Integer '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�༭״̬
Dim mstrMatch As String
Private mstr��Ŀ As String

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cmdSetEspecial_Click()
    Dim strRule As String
    Dim lng������ĿID As Long, lng����ID As Long
    
    strRule = Trim(Me.txtEspecial)
    lng������ĿID = Val(Me.txt��Ŀ.Tag)
    lng����ID = Val(Me.cbo����.ItemData(Me.cbo����.ListIndex))

    Me.txtEspecial = frmLabVerifyEspecial.DefFormula(lng������ĿID, lng����ID, strRule, Me)

End Sub

Private Sub cmdSetRule_Click()
    Dim strRule As String
    Dim lng������ĿID As Long, lng����ID As Long
    
    strRule = Trim(Me.txtRule)
    lng������ĿID = Val(Me.txt��Ŀ.Tag)
    lng����ID = Val(Me.cbo����.ItemData(Me.cbo����.ListIndex))

    Me.txtRule = frmLabVerifySet.DefFormula(lng������ĿID, lng����ID, strRule, Me)
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.picList.hWnd
    Case conPane_Edit
        Item.Handle = Me.picEdit.hWnd
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRetuId As Long

    '------------------------------------
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me

    Case conMenu_Edit_Save
        lngRetuId = EditSave()
        If lngRetuId <> 0 Then
            mlngItemID = lngRetuId: Call RefList(mlngItemID)
            mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
        End If
    Case conMenu_Edit_Untread
        Call EditCancel: Call RefList(mlngItemID)
        mintEditState = 0: Me.picList.Enabled = True: Me.rptList.SetFocus
    Case conMenu_Edit_NewItem
        If EditStart(True, mlngItemID) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select
    Case conMenu_Edit_Modify
        If mlngItemID = 0 Then Exit Sub
        If EditStart(False, mlngItemID) = False Then Exit Sub
        mintEditState = 1: Me.picList.Enabled = False
        Me.dkpMan.FindPane(conPane_Edit).Select

    Case conMenu_Edit_Delete
        Dim strMsg As String
        With Me.rptList
            strMsg = "���ɾ���ù�����" & vbCrLf & "����" & .FocusedRow.Record(mCol.����).Value
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSql = "Zl_������˹���_Edit(3," & mlngItemID & ")"

            Err = 0: On Error GoTo ErrHand
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Err = 0: On Error GoTo 0
            mlngItemID = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                lngRetuId = lngRetuId + 1
            ElseIf lngRetuId > 0 Then
                lngRetuId = lngRetuId - 1
            End If
            If .Rows(lngRetuId).GroupRow = False Then mlngItemID = .Rows(lngRetuId - 1).Record(mCol.ID).Value
            Call RefList(mlngItemID)
        End With

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
    Case conMenu_View_Refresh
        Call RefList(mlngItemID)

    Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
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
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0 And mintEditState = 0)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (mintEditState <> 0)
    Case conMenu_Edit_NewItem
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0 And mintEditState = 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        Control.Enabled = (InStr(1, mstrPrivs, "��ɾ��") > 0 And mintEditState = 0)
        If Control.Enabled Then Control.Enabled = mlngItemID <> 0
        If Control.Enabled Then Control.Enabled = Not Me.rptList.FocusedRow.GroupRow
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Find, conMenu_View_Refresh, conMenu_View_Option: Control.Enabled = (mintEditState = 0)
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs

    mintEditState = 0
    mlngItemID = 0
    mstr��Ŀ = ""

    mstrMatch = gstrMatch

    mLngEditWidth = picEdit.Width

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlCommFun.GetPubIcons
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
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
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

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next


    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Dim panThis As Pane

    Set panThis = dkpMan.CreatePane(conPane_List, 450, 580, DockLeftOf, Nothing)
    panThis.Title = "�����б�"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panThis = dkpMan.CreatePane(conPane_Edit, 550, 580, DockRightOf, Nothing)
    panThis.Title = "����༭"
    panThis.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True

    '-----------------------------------------------------
    With Me.rptList
        .AutoColumnSizing = (Screen.Width / Screen.TwipsPerPixelX > 800)   '������������֮ǰ���ã�������Ч
        Set rptCol = .Columns.Add(mCol.ͼ��, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.����, "����", 70, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 60, True): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ŀ, "��Ŀ", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ŀid, "��ĿID", 120, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.����Id, "����ID", 70, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����ID, "����ID", 70, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��������, "��������", 80, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�Ա�, "�Ա�", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��������, "��������", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��������, "��������", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���䵥λ, "���䵥λ", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        'Set rptCol = .Columns.Add(mCol.����id, "����id", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.����, "����", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�������, "�������", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�����ϵ, "�����ϵ", 60, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.��ʾ��Ϣ, "��ʾ��Ϣ", 60, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.��Ч, "��Ч", 60, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.���, "���", 60, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False

        .SetImageList Me.imgList
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
    End With
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '����װ��

    Call RefSelect
    Call RefList(0)
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub

    Set panThis = Me.dkpMan.FindPane(conPane_Edit)

    panThis.MinTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize 0, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height

End Sub

Private Sub picEdit_Resize()
    Me.cmdSetRule.Left = Me.txtRule.Left
    Me.cmdSetRule.Top = Me.picEdit.ScaleHeight - Me.cmdSetRule.Height - 45
    Me.cmdSetEspecial.Left = Me.txtEspecial.Left + Me.txtEspecial.Width - Me.cmdSetEspecial.Width
    Me.cmdSetEspecial.Top = Me.cmdSetRule.Top
    
    With Me.txtRule
        .Height = Me.picEdit.ScaleHeight - .Top - Me.cmdSetRule.Height - 45
        Me.txtEspecial.Height = .Height
    End With
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = Me.picList.ScaleLeft: .Width = Me.picList.ScaleWidth - .Left
        .Top = Me.picList.ScaleTop: .Height = Me.picList.ScaleHeight - .Top
    End With
End Sub

Private Sub rptList_SelectionChanged()

    With Me.rptList
        If .FocusedRow Is Nothing Then
            mlngItemID = 0
        ElseIf .FocusedRow.GroupRow = True Then
            mlngItemID = 0
        Else
            mlngItemID = Val(.FocusedRow.Record.Item(mCol.ID).Value)
        End If
        Call RefRule(mlngItemID)
    End With
End Sub

Private Sub txt��Ŀ_GotFocus()
    Call zlControl.TxtSelAll(txt��Ŀ)
End Sub

Private Sub txt��Ŀ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mintEditState = 1 Then
        '���ԭ��������
        mstr��Ŀ = ""
        Me.txt��Ŀ.Tag = ""
    End If
End Sub

Private Sub txt��Ŀ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset

    If KeyAscii = vbKeyReturn Then
        If Me.txt��Ŀ <> mstr��Ŀ Then
            Set rsTmp = Select��Ŀ(Trim(Me.txt��Ŀ))
            If rsTmp Is Nothing Then
                Me.txt��Ŀ = mstr��Ŀ
            Else
                Me.txt��Ŀ = rsTmp("����") & "(" & rsTmp("����") & ")": Me.txt��Ŀ.Tag = rsTmp("ID"): mstr��Ŀ = Me.txt��Ŀ
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        Exit Sub

    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ŀ_LostFocus()
    If Me.txt��Ŀ <> mstr��Ŀ Then Me.txt��Ŀ = mstr��Ŀ
End Sub

Private Sub cmd��Ŀ_Click()
    Dim rsTmp As ADODB.Recordset

    Set rsTmp = Select��Ŀ
    If Not rsTmp Is Nothing Then
        Me.txt��Ŀ = rsTmp("����") & "(" & rsTmp("����") & ")": Me.txt��Ŀ.Tag = rsTmp("ID"): mstr��Ŀ = Me.txt��Ŀ
    End If
End Sub


Private Sub txt���_GotFocus()
    Call zlControl.TxtSelAll(txt���)
End Sub

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Private Sub RefRule(ByVal lngItem As Long)
    '���ܣ�ˢ�¹���
    Dim intIndex As Integer
    If rptList.FocusedRow Is Nothing Then
        Exit Sub
    End If
    If Not rptList.FocusedRow.GroupRow Then
        With rptList.FocusedRow.Record
            '- ������
            Me.txt���� = .Item(mCol.����).Value
            Me.txtName = .Item(mCol.����).Value
            Me.txtInfo = .Item(mCol.��ʾ��Ϣ).Value
            Me.txt��ע = .Item(mCol.��ע).Value

            If .Item(mCol.����).Value = "X-δ����" Then
                cbo����.ListIndex = 0
            Else
                For intIndex = 0 To cbo����.ListCount - 1
                    If cbo����.List(intIndex) = .Item(mCol.����).Value Then
                        cbo����.ListIndex = intIndex: Exit For
                    End If
                Next
            End If

            If Val(.Item(mCol.����Id).Value) = 0 Then
                cbo����.ListIndex = 0
            Else
                For intIndex = 0 To cbo����.ListCount - 1
                    If cbo����.ItemData(intIndex) = Val(.Item(mCol.����Id).Value) Then
                        cbo����.ListIndex = intIndex: Exit For
                    End If
                Next
            End If

            txt��Ŀ = .Item(mCol.��Ŀ).Value: txt��Ŀ.Tag = Val(.Item(mCol.��Ŀid).Value): mstr��Ŀ = txt��Ŀ
            cboValid.ListIndex = Val(.Item(mCol.��Ч).Value)

            '--��������
            cbo�Ա�.ListIndex = 0
            For intIndex = 0 To cbo�Ա�.ListCount - 1
                If cbo�Ա�.List(intIndex) = .Item(mCol.�Ա�).Value Then
                    cbo�Ա�.ListIndex = intIndex: Exit For
                End If
            Next

            txt�������� = .Item(mCol.��������).Value
            txt�������� = .Item(mCol.��������).Value

            cbo���䵥λ.ListIndex = 0
            For intIndex = 0 To cbo���䵥λ.ListCount - 1
                If cbo���䵥λ.List(intIndex) = .Item(mCol.���䵥λ).Value Then
                    cbo���䵥λ.ListIndex = intIndex: Exit For
                End If
            Next

            chk��ֹ = Val(.Item(mCol.���).Value)

            cbo����.ListIndex = 0
            For intIndex = 0 To cbo����.ListCount - 1
                If cbo����.ItemData(intIndex) = .Item(mCol.����ID).Value Then
                    cbo����.ListIndex = intIndex: Exit For
                End If
            Next

            cbo��������.ListIndex = 0
            For intIndex = 0 To cbo��������.ListCount - 1
                If Val(cbo��������.List(intIndex)) = Val(.Item(mCol.��������).Value) Then
                    cbo��������.ListIndex = intIndex: Exit For
                End If
            Next

            txt��� = .Item(mCol.���).Value

            '--����
            txtRule = .Item(mCol.����).Value
            txtEspecial = .Item(mCol.�������).Value
            If UCase(.Item(mCol.�����ϵ).Value) = "AND" Then
                Me.optAnd = True
                Me.optOr = False
            Else
                Me.optAnd = False
                Me.optOr = True
            End If
        End With
    End If
End Sub

Private Function RefList(ByVal lngItemID As Long) As Long
    '���ܣ�ˢ���б�
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    strSql = "Select  Nvl(A.����, 'X-δ����') As ����,A.ID, A.����, A.����,C.����||'('||C.���� || ')' As ��Ŀ, A.��Ŀid, B.����||'('||B.����|| ')' As ����," & vbNewLine & _
            "       A.����id, A.����id, A.��������, A.�Ա�, A.��������, A.��������, A.���䵥λ, A.���, A.����, A.�������, A.�����ϵ, A.��ʾ��Ϣ, A.��Ч, A.���," & vbNewLine & _
            "       A.��ע" & vbNewLine & _
            "From ������ĿĿ¼ C, �������� B, ������˹��� A" & vbNewLine & _
            "Where A.����id = B.ID(+) And A.��Ŀid = C.ID(+) " & vbNewLine & _
            "Order By A.����, A.����"
    Err = 0: On Error GoTo ErrHand
    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Me.rptList.Records.DeleteAll
    With rsRecord
        Do While Not .EOF

            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem("0"): rptItem.Icon = 0

            rptRcd.AddItem CStr("" & !����)
            rptRcd.AddItem CStr("" & !ID)
            Set rptItem = rptRcd.AddItem(CStr("" & !����)): rptItem.SortPriority = Val(("" & !����))
            If Val("" & !ID) = 0 Then
                rptRcd.AddItem CStr("...�÷�����û�й���...")
            Else
                rptRcd.AddItem CStr("" & !����)
            End If
            rptRcd.AddItem IIf(CStr("" & !��Ŀ) = "()", "", CStr("" & !��Ŀ))
            rptRcd.AddItem CStr("" & !��Ŀid)
            rptRcd.AddItem IIf(CStr("" & !����) = "()", "", CStr("" & !����))
            rptRcd.AddItem CStr("" & !����Id)
            rptRcd.AddItem CStr("" & !����ID)
            rptRcd.AddItem CStr("" & !��������)
            rptRcd.AddItem CStr("" & !�Ա�)
            rptRcd.AddItem CStr("" & !��������)
            rptRcd.AddItem CStr("" & !��������)
            rptRcd.AddItem CStr("" & !���䵥λ)
            rptRcd.AddItem CStr("" & !���)
            rptRcd.AddItem Tran��ʾ��ʽ(CStr("" & !����))
            rptRcd.AddItem Tran��ʾ��ʽ(CStr("" & !�������))
            rptRcd.AddItem CStr("" & !�����ϵ)
            rptRcd.AddItem CStr("" & !��ʾ��Ϣ)
            rptRcd.AddItem CStr("" & !��Ч)
            rptRcd.AddItem CStr("" & !���)
            rptRcd.AddItem CStr("" & !��ע)

            .MoveNext
        Loop
    End With

    With Me.rptList
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(mCol.����)
        .GroupsOrder(0).SortAscending = True
        .Populate
    End With

    If lngItemID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngItemID Then
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    Call rptList_SelectionChanged

    RefList = Me.rptList.Records.Count
    Me.stbThis.Panels(2).Text = "����" & Me.rptList.Records.Count & "����Ŀ"
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    RefList = Me.rptList.Records.Count
End Function

Private Sub RefSelect()
    'ˢ��ѡ����Ŀ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    '--����
    On Error GoTo ErrHand
    cbo����.Clear
    cbo����.AddItem ""
    strSql = "Select ����,���� From ���������� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Do Until rsTmp.EOF
        cbo����.AddItem "" & rsTmp.Fields("����") & "-" & rsTmp.Fields("����")
        rsTmp.MoveNext
    Loop
    cbo����.ListIndex = 0

    '����
    cbo����.Clear
    cbo����.AddItem ""
    strSql = "Select ID, ����, ���� From �������� Order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Do Until rsTmp.EOF
        cbo����.AddItem rsTmp.Fields("����") & "(" & rsTmp.Fields("����") & ")"
        cbo����.ItemData(cbo����.ListCount - 1) = rsTmp.Fields("ID")
        rsTmp.MoveNext
    Loop
    cbo����.ListIndex = 0

    '�Ա�
    cbo�Ա�.Clear
    cbo�Ա�.AddItem ""
    cbo�Ա�.AddItem "��"
    cbo�Ա�.AddItem "Ů"
    cbo�Ա�.ListIndex = 0

    '���䵥λ
    cbo���䵥λ.Clear
    cbo���䵥λ.AddItem ""
    cbo���䵥λ.AddItem "Сʱ"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0

    '�ͼ����
    cbo����.Clear
    cbo����.AddItem ""
    strSql = "Select A.ID, A.����, A.����, B.������� From ��������˵�� B, ���ű� A Where A.ID = B.����id And B.�������� = '�ٴ�'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Do Until rsTmp.EOF
        cbo����.AddItem rsTmp.Fields("����") & "(" & rsTmp.Fields("����") & ")"
        cbo����.ItemData(cbo����.ListCount - 1) = rsTmp.Fields("ID")
        rsTmp.MoveNext
    Loop
    cbo����.ListIndex = 0

    '��������
    cbo��������.Clear
    cbo��������.AddItem ""
    cbo��������.AddItem "1-����"
    cbo��������.AddItem "2-סԺ"
    
    '��Ч��
    cboValid.Clear
    cboValid.AddItem "0-��ֹʹ�øù���"
    cboValid.AddItem "1-���ʱʹ�øù���"
    cboValid.AddItem "2-���������ʱʹ�øù���"
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function EditSave() As Long
    '��������
    Dim strSql As String
    Dim lngNewId As Long, rsTmp As ADODB.Recordset
    'һ�����ݼ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: EditSave = 0: Exit Function
    End If
    If Val(Me.txt����.Text) > Val(String(Me.txt����.MaxLength, "9")) Then
        MsgBox "����̫��", vbInformation, gstrSysName
        Me.txt����.SetFocus: EditSave = 0: Exit Function
    End If
    
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txtName.SetFocus: EditSave = 0: Exit Function
    End If
    
    Err = 0: On Error GoTo ErrHand
    If Trim(Me.cbo����.List(cbo����.ListIndex)) <> "" Then
        strSql = "select ����,���� From ���������� where ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Mid(Trim(Me.cbo����.List(cbo����.ListIndex)), 1, InStr(Trim(Me.cbo����.List(cbo����.ListIndex)), "-") - 1))
        If rsTmp.RecordCount <= 0 Then
            MsgBox Trim(Me.cbo����.List(cbo����.ListIndex)) & "�ѱ�������ɾ����������ѡ��", vbInformation, gstrSysName
            Me.cbo����.SetFocus: EditSave = 0: Exit Function
        End If
    End If

    '���ݱ��������֯
    strSql = "'" & Replace(Trim(Me.txt����.Text), "'", "") & "','" & Replace(Trim(Me.txtName.Text), "'", "") & "','" & Me.cbo����.List(Me.cbo����.ListIndex) & "'" & _
              "," & IIf(Val(txt��Ŀ.Tag) = 0, "Null", Val(txt��Ŀ.Tag)) & "," & _
              IIf(Val(Me.cbo����.ItemData(Me.cbo����.ListIndex)) = 0, "Null", Val(Me.cbo����.ItemData(Me.cbo����.ListIndex))) & "," & _
              IIf(Val(Me.cbo����.ItemData(Me.cbo����.ListIndex)) = 0, "Null", Val(Me.cbo����.ItemData(Me.cbo����.ListIndex))) & "," & _
              IIf(Val(Me.cbo��������.List(Me.cbo��������.ListIndex)) = 0, "Null", "'" & Val(Me.cbo��������.List(Me.cbo��������.ListIndex)) & "'") & "," & _
              IIf(Trim(Me.cbo�Ա�.List(Me.cbo�Ա�.ListIndex)) = "", "Null", "'" & Me.cbo�Ա�.List(Me.cbo�Ա�.ListIndex) & "'") & "," & _
              IIf(Val(txt��������) = 0, "Null", "'" & Val(txt��������) & "'") & "," & _
              IIf(Val(txt��������) = 0, "Null", "'" & Val(txt��������) & "'") & "," & _
              IIf(Trim(Me.cbo���䵥λ.List(Me.cbo���䵥λ.ListIndex)) = "", "Null", "'" & Me.cbo���䵥λ.List(Me.cbo���䵥λ.ListIndex) & "'") & "," & _
              "'" & Replace(Trim(txt���), "'", "") & "'," & _
              IIf(Trim(txtRule) = "", "Null", "'" & Tran���湫ʽ(Replace(Trim(txtRule), "'", "''") & "'")) & "," & _
              IIf(Trim(txtEspecial) = "", "Null", "'" & Tran���湫ʽ(Replace(Trim(txtEspecial), "'", "''") & "'")) & "," & _
              IIf(optAnd, "'AND'", "'OR'") & "," & _
              IIf(Trim(txtInfo) = "", "Null", "'" & Replace(Trim(txtInfo), "'", "''") & "'") & "," & _
              IIf(chk����.Value = 1, "'1'", "'0'") & "," & _
              IIf(chk��ֹ.Value = 1, "'1'", "'0'") & ",'" & _
              Val(cboValid.List(cboValid.ListIndex)) & "'," & _
              IIf(Trim(txt��ע) = "", "Null", "'" & Replace(Trim(txt��ע), "'", "''") & "'")

    lngNewId = mlngItemID
    If Me.picEdit.Tag = "����" Then
        lngNewId = zlDatabase.GetNextId("������˹���")
        strSql = "Zl_������˹���_Edit(1," & lngNewId & "," & strSql & ")"
    Else
        strSql = "Zl_������˹���_Edit(2," & lngNewId & "," & strSql & ")"
    End If

    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    'Call SQLTest(App.ProductName, Me.Caption, strSQL): gcnOracle.Execute strSQL, , adCmdStoredProc: Call SQLTest

    If Me.picEdit.Tag = "����" Then mlngItemID = lngNewId

    Me.picEdit.Tag = "":    mintEditState = 0
    Me.picList.Enabled = True: Me.picEdit.Enabled = False: Me.picEdit.BackColor = &H8000000F

    EditSave = mlngItemID: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    EditSave = 0: Exit Function
End Function

Private Function EditStart(blnAdd As Boolean, lngItemID As Long) As Boolean
    '��ʼ�༭
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHand
    If blnAdd Then
        '����
        strSql = "Select Max(����) as ���� From ������˹���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Me.txt���� = zlCommFun.IncStr(IIf(Trim("" & rsTmp.Fields("����")) = "", "000", Trim("" & rsTmp.Fields("����"))))
        Else
            Me.txt���� = "001"
        End If
        Me.txtName = ""
        If Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
        Me.txt��Ŀ = "": Me.txt��Ŀ.Tag = "": mstr��Ŀ = ""
        If Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
        cboValid.ListIndex = 1
        Me.txtInfo = ""
        Me.txt��ע = ""

        '��������
        If Me.cbo�Ա�.ListCount > 0 Then Me.cbo�Ա�.ListIndex = 0
        Me.txt�������� = "": Me.txt�������� = ""
        If Me.cbo���䵥λ.ListCount > 0 Then Me.cbo���䵥λ.ListIndex = 0
        chk��ֹ.Value = 0
        If Me.cbo����.ListCount > 0 Then Me.cbo����.ListIndex = 0
        If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
        chk����.Value = 0
        Me.txt��� = ""
        '����
        Me.optOr.Value = True
        Me.txtRule = "": Me.txtEspecial = ""

    End If
    picEdit.Tag = IIf(blnAdd, "����", "�޸�")
    picList.Enabled = False
    picEdit.Enabled = True: Me.picEdit.BackColor = RGB(250, 250, 250)
    mintEditState = 1
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub EditCancel()
    'ȡ��
    picList.Enabled = True
    picEdit.Enabled = False: Me.picEdit.BackColor = &H8000000F
    mintEditState = 0

End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub

    '-------------------------------------------------
    '�������ݱ��
    If zlControl.RPTCopyToVSF(Me.rptList, Me.vfgList) Is Nothing Then Exit Sub
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "������˹���"
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

Private Function Select��Ŀ(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer
    
    On Error GoTo ErrHand
    If Len(strName) = 0 Then
        '������Ŀ
        strSql = "Select 0 As ĩ��, ID, �ϼ�id, ����, ����" & vbNewLine & _
                "From ���Ʒ���Ŀ¼ A" & vbNewLine & _
                "Where ���� = 5" & vbNewLine & _
                "Start With �ϼ�id is null" & vbNewLine & _
                "Connect By Prior A.id = A.�ϼ�ID" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Distinct 1 As ĩ��, ID, ����id, ����, ���� From ������ĿĿ¼ Where Nvl(����Ӧ��, 0) = 1 And ��� = 'C'"
        Set Select��Ŀ = zlDatabase.ShowSelect(Me, strSql, 2, "������Ŀ", , , , , True)
    Else
        'ָ����Ŀ
        strSQLItem = " From ������Ŀ���� B,������ĿĿ¼ A" & _
            " Where A.ID=B.������ĿID And Nvl(A.����Ӧ��, 0) = 1 And A.��� = 'C'" & _
            " And (Upper(A.����) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%')"
'
'        strSQL = "Select distinct  0 As ĩ��, ID, �ϼ�id, ����, ����" & vbNewLine & _
'                "From ���Ʒ���Ŀ¼ A" & vbNewLine & _
'                "Where ���� = 5" & vbNewLine & _
'                "Start With ID In (Select A.����id " & strSQLItem & ")" & vbNewLine & _
'                "Connect By Prior A.�ϼ�id = A.ID" & vbNewLine & _
'                "Union All" & vbNewLine & _
'                "Select Distinct 1 As ĩ��, A.ID, A.����id, A.����, A.���� " & strSQLItem
                
        strSql = "Select Distinct  A.ID, A.����id, A.����, A.���� " & strSQLItem
        Set Select��Ŀ = zlDatabase.ShowSelect(Me, strSql, 0, "������Ŀ", , , , , True)
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Select���(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer

    If Len(strName) = 0 Then
        '������Ŀ
        strSql = "Select Distinct 0 As ĩ��, ID, �ϼ�id, to_char(���) as ����, ����" & vbNewLine & _
                "From ����������� A" & vbNewLine & _
                "Where ���='D'" & vbNewLine & _
                "Start With �ϼ�ID IS NULL " & vbNewLine & _
                "Connect By Prior A.id = A.�ϼ�ID" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Distinct 1 As ĩ��, ID, ����id, ����, ���� From ��������Ŀ¼ Where ���='D'"
    Else
        'ָ����Ŀ
        strSQLItem = " From ��������Ŀ¼ A" & _
            " Where A.��� = 'D'" & _
            " And (Upper(A.����) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(A.����) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = "Select Distinct 0 As ĩ��, ID, �ϼ�id,to_char(���) as ����, ����" & vbNewLine & _
                "From ����������� A" & vbNewLine & _
                "Where ���='D'" & vbNewLine & _
                "Start With ID In (Select A.����id " & strSQLItem & ")" & vbNewLine & _
                "Connect By Prior A.�ϼ�id = A.ID" & vbNewLine & _
                "Union All" & vbNewLine & _
                "Select Distinct 1 As ĩ��, A.ID, A.����id, A.����, A.���� " & strSQLItem
    End If
    Set Select��� = zlDatabase.ShowSelect(Me, strSql, 2, "����", , , , , True)

End Function

Private Function Tran���湫ʽ(ByVal str��ʾ��ʽ As String) As String
    '����ʾ��ʽתΪ���湫ʽ
    Dim strItem As String, strTmp As String, strLast As String
    Dim rsGS As ADODB.Recordset, lngLength As Long
    strItem = "": strTmp = ""
    On Error GoTo ErrHand
    If str��ʾ��ʽ <> "" Then
        Do While str��ʾ��ʽ Like "*[[]*[]]*"
            strTmp = strTmp & Mid(str��ʾ��ʽ, 1, InStr(str��ʾ��ʽ, "[") - 1)
            lngLength = InStr(str��ʾ��ʽ, "]") - InStr(str��ʾ��ʽ, "[") - 1
            strItem = Mid(str��ʾ��ʽ, InStr(str��ʾ��ʽ, "[") + 1, lngLength)
            If InStr(strItem, "_") > 0 Then
                strItem = Mid(strItem, 1, InStr(strItem, "_") - 1)
            End If
            If InStr(strItem, "�ϴ�.") > 0 Then
                strLast = "�ϴ�."
                strItem = Replace(strItem, "�ϴ�.", "")
            ElseIf InStr(strItem, "���.") > 0 Then
                strLast = "���."
                strItem = Replace(strItem, "���.", "")
            Else
                strLast = ""
            End If
            gstrSql = "Select ID,Ӣ���� From ����������Ŀ  Where (id=[1] or ����=[2]) "
            Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem), strItem)

            Do Until rsGS.EOF
                strTmp = strTmp & "[" & strLast & Val("" & rsGS.Fields("ID")) & "]"
                rsGS.MoveNext
            Loop
            str��ʾ��ʽ = Mid(str��ʾ��ʽ, InStr(str��ʾ��ʽ, "]") + 1)
        Loop
        strTmp = strTmp & Mid(str��ʾ��ʽ, InStr(str��ʾ��ʽ, "]") + 1)
        strTmp = Replace(strTmp, "{D:©����}", "{D:1}")
        strTmp = Replace(strTmp, "{D:������}", "{D:2}")
        strTmp = Replace(strTmp, "{D:©�������}", "{D:3}")
        Tran���湫ʽ = strTmp
        
        
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Tran��ʾ��ʽ(ByVal str���湫ʽ As String) As String
    '�����湫ʽתΪ��ʾ��ʽ
    Dim strTmp As String, strItem As String, strLast As String
    Dim rsGS As ADODB.Recordset, lngLength As Long
    On Error GoTo ErrHand
    Do While str���湫ʽ Like "*[[]*[]]*"
        strTmp = strTmp & Mid(str���湫ʽ, 1, InStr(str���湫ʽ, "[") - 1)
        lngLength = InStr(str���湫ʽ, "]") - InStr(str���湫ʽ, "[") - 1
        strItem = Mid(str���湫ʽ, InStr(str���湫ʽ, "[") + 1, lngLength)
        If InStr(strItem, "�ϴ�.") > 0 Then
            strLast = "�ϴ�."
            strItem = Replace(strItem, "�ϴ�.", "")
        ElseIf InStr(strItem, "���.") > 0 Then
            strLast = "���."
            strItem = Replace(strItem, "���.", "")
        Else
            strLast = ""
        End If
        gstrSql = "Select ID,Ӣ����,���� From ����������Ŀ Where id=[1] "
        Set rsGS = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(strItem))
        Do Until rsGS.EOF
            If Trim("" & rsGS.Fields("Ӣ����")) <> "" Then
                strTmp = strTmp & "[" & strLast & rsGS.Fields("����") & "_" & Trim("" & rsGS.Fields("Ӣ����")) & "]"
            Else
                strTmp = strTmp & "[" & strLast & Val(strItem) & "]"
            End If
            rsGS.MoveNext
        Loop
        str���湫ʽ = Mid(str���湫ʽ, InStr(str���湫ʽ, "]") + 1)
    Loop
    strTmp = strTmp & Mid(str���湫ʽ, InStr(str���湫ʽ, "]") + 1)
    strTmp = Replace(strTmp, "{D:1}", "{D:©����}")
    strTmp = Replace(strTmp, "{D:2}", "{D:������}")
    strTmp = Replace(strTmp, "{D:3}", "{D:©�������}")
    Tran��ʾ��ʽ = strTmp
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
