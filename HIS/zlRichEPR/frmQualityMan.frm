VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmQualityMan 
   Caption         =   "�����������"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   9615
   Icon            =   "frmQualityMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9615
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   1095
      Left            =   3330
      TabIndex        =   23
      Top             =   810
      Width           =   1005
      _Version        =   589884
      _ExtentX        =   1773
      _ExtentY        =   1931
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   2565
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityMan.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityMan.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQualityMan.frx":0EBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2085
      Left            =   225
      ScaleHeight     =   2085
      ScaleWidth      =   2325
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2325
      Begin VB.Label lblZZXD 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   22
         Top             =   1755
         Width           =   2580
      End
      Begin VB.Label lblSXWC 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   21
         Top             =   1485
         Width           =   2580
      End
      Begin VB.Label lblSXCS 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   20
         Top             =   1215
         Width           =   2580
      End
      Begin VB.Label lblZZSX 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   19
         Top             =   945
         Width           =   2580
      End
      Begin VB.Label lblRYRS 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   18
         Top             =   675
         Width           =   2580
      End
      Begin VB.Label lblBM 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   945
         TabIndex        =   17
         Top             =   405
         Width           =   2580
      End
      Begin VB.Label lblMC 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   945
         TabIndex        =   16
         Top             =   135
         Width           =   2580
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "�����޶�:"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   1755
         Width           =   870
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "��д���:"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   1485
         Width           =   870
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "��д��ʱ:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "������д:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   945
         Width           =   870
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "���ұ���:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   405
         Width           =   870
      End
      Begin VB.Label lblName1 
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   135
         Width           =   870
      End
   End
   Begin VB.PictureBox picDate 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   180
      ScaleHeight     =   1140
      ScaleWidth      =   2325
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   2325
      Begin VB.CommandButton cmdSearch 
         Caption         =   "����ͳ��(&R)"
         Height          =   350
         Left            =   450
         TabIndex        =   2
         Top             =   720
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   300
         Left            =   450
         TabIndex        =   1
         Top             =   390
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   69009411
         CurrentDate     =   38683
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   300
         Left            =   450
         TabIndex        =   0
         Top             =   45
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   69009411
         CurrentDate     =   38683
      End
      Begin VB.Label lblDateFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   105
         Width           =   180
      End
      Begin VB.Label lblDateTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   450
         Width           =   180
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5790
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQualityMan.frx":1258
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11880
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
   Begin XtremeSuiteControls.TaskPanel tplThis 
      Height          =   4425
      Left            =   45
      TabIndex        =   7
      Top             =   90
      Width           =   2805
      _Version        =   589884
      _ExtentX        =   4948
      _ExtentY        =   7805
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgסԺ���� 
      Height          =   1110
      Left            =   6030
      TabIndex        =   24
      Top             =   135
      Width           =   1410
      _cx             =   2487
      _cy             =   1958
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vfg���ﲡ�� 
      Height          =   1110
      Left            =   4500
      TabIndex        =   25
      Top             =   135
      Width           =   1410
      _cx             =   2487
      _cy             =   1958
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vfg������ 
      Height          =   1110
      Left            =   7560
      TabIndex        =   26
      Top             =   135
      Width           =   1410
      _cx             =   2487
      _cy             =   1958
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VB.Image imgBG 
      Height          =   2295
      Left            =   7290
      Picture         =   "frmQualityMan.frx":1AEA
      Top             =   3465
      Visible         =   0   'False
      Width           =   2265
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   3240
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmQualityMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ID = 0: ����: ����: ������д: ��д��ʱ: �����: �����޶�: ��Ժ: ת��: ������Ժ: ת��: תԺ: ����
End Enum

Private Enum Enum��������
    ���ﲡ�� = 1
    סԺ���� = 2
    ������ = 4
End Enum
Private mvar�������� As Enum��������

Private Const ID_ViewFile = 802           '�鿴�ļ�
Private Const ID_ViewPati = 803           '�鿴����

Private cbp�ļ� As CommandBarPopup      '�ļ��˵�
Private cbp��ͼ As CommandBarPopup      '��ͼ�˵�
Private cbp���� As CommandBarPopup      '�����˵�
Private mfrmQualityViewFile As New frmQualityViewFile
Private mfrmQualityViewPati As New frmQualityViewPati

Private Bar���� As CommandBar           '���ù�����
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '       strSubhead����ӡ�ĸ�����
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Select Case mvar��������
    Case ���ﲡ��
        Set objPrint.Body = Me.vfg���ﲡ��
        objPrint.Title.Text = "���ﲡ���������"
    Case סԺ����
        Set objPrint.Body = Me.vfgסԺ����
        objPrint.Title.Text = "סԺ�����������"
    Case ������
        Set objPrint.Body = Me.vfg������
        objPrint.Title.Text = "�������������"
    End Select
    objPrint.Title.Font.Name = "����"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("ͳ��ʱ��:" & Format(Me.dtpDateFrom.Value, "yyyy-MM-dd") & " �� " & Format(Me.dtpDateTo.Value, "yyyy-MM-dd"))
    Call objPrint.UnderAppRows.Add(objAppRow)
    
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

Private Sub FillGrid(ByVal strFrom As String, ByVal strTo As String)
    '�������
    Dim Rs As ADODB.Recordset, i As Long, lngCount(1 To 10) As Long
    Select Case mvar��������
    Case ���ﲡ��
        gstrSQL = "Select d.ID, d.����, d.����, l.������д, l.��д��ʱ, l.����� " & _
            " From ���ű� d," & _
            "          (Select ����id, Sum(Decode(���ʱ��, Null, 1, 0)) As ������д, " & _
            "                             Sum(Decode(���ʱ��, Null, Decode(Sign(Sysdate - ����ʱ�� - 1), 1, 1, 0), 0)) As ��д��ʱ, " & _
            "                             Sum(Decode(���ʱ��, Null, 0, 1)) As ����� " & _
            "              From ���Ӳ�����¼ " & _
            "              Where �������� = 1 And ����ʱ�� Between [1] And [2] " & _
            "              Group By ����id) l " & _
            " Where D.ID = L.����id " & _
            " Order By d.����"
        Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
        Call InitGrid(mvar��������)
        Me.vfg���ﲡ��.Rows = 2 + Rs.RecordCount
        stbThis.Panels(2).Text = "���ƣ�" & Rs.RecordCount & "����¼��"
        i = 1
        Do While Not Rs.EOF
            With Me.vfg���ﲡ��
                .TextMatrix(i, mCol.ID) = NVL(Rs("ID"))
                .TextMatrix(i, mCol.����) = NVL(Rs("����"))
                .TextMatrix(i, mCol.����) = NVL(Rs("����"))
                .TextMatrix(i, mCol.������д) = IIf(NVL(Rs("������д")) = 0, "", NVL(Rs("������д"))): lngCount(1) = lngCount(1) + Val(.TextMatrix(i, mCol.������д))
                .TextMatrix(i, mCol.��д��ʱ) = IIf(NVL(Rs("��д��ʱ")) = 0, "", NVL(Rs("��д��ʱ"))): lngCount(2) = lngCount(2) + Val(.TextMatrix(i, mCol.��д��ʱ))
                .TextMatrix(i, mCol.�����) = IIf(NVL(Rs("�����")) = 0, "", NVL(Rs("�����"))): lngCount(3) = lngCount(3) + Val(.TextMatrix(i, mCol.�����))
            End With
            Rs.MoveNext
            i = i + 1
        Loop
        With Me.vfg���ﲡ��
            .TextMatrix(i, mCol.����) = "�ϼ�"
            .TextMatrix(i, mCol.����) = ""
            .TextMatrix(i, mCol.������д) = lngCount(1)
            .TextMatrix(i, mCol.��д��ʱ) = lngCount(2)
            .TextMatrix(i, mCol.�����) = lngCount(3)
        End With
        Rs.Close
        Set Rs = Nothing
        If Me.vfg���ﲡ��.Rows > 1 Then Me.vfg���ﲡ��.Row = 1
    Case סԺ����
        gstrSQL = "Select d.ID, d.����, d.����, l.������д, l.��д��ʱ, l.�����, l.�����޶�, p.��Ժ, t.ת��, e.������Ժ, e.����, e.תԺ, i.ת�� " & _
            " From ���ű� d, (Select ��Ժ����id, Count(*) As ��Ժ From ������ҳ Where ��Ժ���� Between [1] And [2] Group By ��Ժ����id) p, " & _
            "          (Select ����id, Sum(Decode(���ʱ��, Null, 1, 0)) As ������д, " & _
            "                             Sum(Decode(���ʱ��, Null, Decode(Sign(Sysdate - ����ʱ�� - 1), 1, 1, 0), 0)) As ��д��ʱ, " & _
            "                             Sum(Decode(���ʱ��, Null, 0, 1)) As �����, " & _
            "                             Sum(Decode(���ʱ��, Null, 0, Decode(NVL(ǩ������, 0), 0, 1, 0))) As �����޶� " & _
            "              From ���Ӳ�����¼ " & _
            "              Where  �������� = 2 And ����ʱ�� Between [1] And [2] " & _
            "              Group By ����id) l, " & _
            "          (Select b.Id, Count((a.����id)) As ת�� " & _
            "              From ���˱䶯��¼ a, ���ű� b " & _
            "              Where a.��ʼʱ�� Between [1] And [2] And a.��ʼԭ�� = 3 And Nvl(���Ӵ�λ, 0) = 0 And b.Id = a.����id " & _
            "              Group By b.Id) t, " & _
            "          (Select a.Id, Sum(Decode(b.��Ժ��ʽ, '����', 1, 0)) As ������Ժ, Sum(Decode(b.��Ժ��ʽ, '����', 1, 0)) As ����, " & _
            "                             Sum(Decode(b.��Ժ��ʽ, 'תԺ', 1, 0)) As תԺ " & _
            "              From ���ű� a, ������ҳ b " & _
            "              Where b.��Ժ���� Between [1] And [2] And b.��Ժ����id = a.Id " & _
            "              Group By a.Id) e, " & _
            "          (Select b.Id, Count((a.����id)) As ת�� " & _
            "              From ���˱䶯��¼ a, ���ű� b " & _
            "              Where a.��ֹʱ�� Between [1] And [2] And a.��ֹԭ�� = 3 And Nvl(���Ӵ�λ, 0) = 0 And b.Id = a.����id " & _
            "              Group By b.Id) i " & _
            " Where d.Id = p.��Ժ����id And p.��Ժ����id = l.����id(+) And t.Id = l.����id And e.Id = l.����id And i.Id = l.����id " & _
            " Order By d.����"
        Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
        Call InitGrid(mvar��������)
        Me.vfgסԺ����.Rows = 3 + Rs.RecordCount
        stbThis.Panels(2).Text = "���ƣ�" & Rs.RecordCount & "����¼��"
        i = 2
        Do While Not Rs.EOF
            With Me.vfgסԺ����
                .TextMatrix(i, mCol.ID) = NVL(Rs("ID"))
                .TextMatrix(i, mCol.����) = NVL(Rs("����"))
                .TextMatrix(i, mCol.����) = NVL(Rs("����"))
                .TextMatrix(i, mCol.������д) = IIf(NVL(Rs("������д")) = 0, "", NVL(Rs("������д"))): lngCount(1) = lngCount(1) + Val(.TextMatrix(i, mCol.������д))
                .TextMatrix(i, mCol.��д��ʱ) = IIf(NVL(Rs("��д��ʱ")) = 0, "", NVL(Rs("��д��ʱ"))): lngCount(2) = lngCount(2) + Val(.TextMatrix(i, mCol.��д��ʱ))
                .TextMatrix(i, mCol.�����) = IIf(NVL(Rs("�����")) = 0, "", NVL(Rs("�����"))): lngCount(3) = lngCount(3) + Val(.TextMatrix(i, mCol.�����))
                .TextMatrix(i, mCol.�����޶�) = IIf(NVL(Rs("�����޶�")) = 0, "", NVL(Rs("�����޶�"))): lngCount(4) = lngCount(4) + Val(.TextMatrix(i, mCol.�����޶�))
                .TextMatrix(i, mCol.��Ժ) = IIf(NVL(Rs("��Ժ")) = 0, "", NVL(Rs("��Ժ"))): lngCount(5) = lngCount(5) + Val(.TextMatrix(i, mCol.��Ժ))
                .TextMatrix(i, mCol.ת��) = IIf(NVL(Rs("ת��")) = 0, "", NVL(Rs("ת��"))): lngCount(6) = lngCount(6) + Val(.TextMatrix(i, mCol.ת��))
                .TextMatrix(i, mCol.������Ժ) = IIf(NVL(Rs("������Ժ")) = 0, "", NVL(Rs("������Ժ"))): lngCount(7) = lngCount(7) + Val(.TextMatrix(i, mCol.������Ժ))
                .TextMatrix(i, mCol.ת��) = IIf(NVL(Rs("ת��")) = 0, "", NVL(Rs("ת��"))): lngCount(8) = lngCount(8) + Val(.TextMatrix(i, mCol.ת��))
                .TextMatrix(i, mCol.תԺ) = IIf(NVL(Rs("תԺ")) = 0, "", NVL(Rs("תԺ"))): lngCount(9) = lngCount(9) + Val(.TextMatrix(i, mCol.תԺ))
                .TextMatrix(i, mCol.����) = IIf(NVL(Rs("����")) = 0, "", NVL(Rs("����"))): lngCount(10) = lngCount(10) + Val(.TextMatrix(i, mCol.����))
            End With
            Rs.MoveNext
            i = i + 1
        Loop
        With Me.vfgסԺ����
            .TextMatrix(i, mCol.����) = "�ϼ�"
            .TextMatrix(i, mCol.����) = ""
            .TextMatrix(i, mCol.������д) = lngCount(1)
            .TextMatrix(i, mCol.��д��ʱ) = lngCount(2)
            .TextMatrix(i, mCol.�����) = lngCount(3)
            .TextMatrix(i, mCol.�����޶�) = lngCount(4)
            .TextMatrix(i, mCol.��Ժ) = lngCount(5)
            .TextMatrix(i, mCol.ת��) = lngCount(6)
            .TextMatrix(i, mCol.������Ժ) = lngCount(7)
            .TextMatrix(i, mCol.ת��) = lngCount(8)
            .TextMatrix(i, mCol.תԺ) = lngCount(9)
            .TextMatrix(i, mCol.����) = lngCount(10)
            .Cell(flexcpFontBold, i, mCol.����) = True
            .Cell(flexcpFontBold, i, mCol.����) = True
            .Cell(flexcpFontBold, i, mCol.������д) = True
            .Cell(flexcpFontBold, i, mCol.��д��ʱ) = True
            .Cell(flexcpFontBold, i, mCol.�����) = True
            .Cell(flexcpFontBold, i, mCol.�����޶�) = True
            .Cell(flexcpFontBold, i, mCol.��Ժ) = True
            .Cell(flexcpFontBold, i, mCol.ת��) = True
            .Cell(flexcpFontBold, i, mCol.������Ժ) = True
            .Cell(flexcpFontBold, i, mCol.ת��) = True
            .Cell(flexcpFontBold, i, mCol.תԺ) = True
            .Cell(flexcpFontBold, i, mCol.����) = True
        End With
        Rs.Close
        Set Rs = Nothing
        If Me.vfgסԺ����.Rows > 2 Then Me.vfgסԺ����.Row = 2: Call vfgסԺ����_RowColChange
    Case ������
        gstrSQL = "Select d.ID, d.����, d.����, l.������д, l.��д��ʱ, l.�����, l.�����޶�, p.��Ժ, t.ת��, e.������Ժ, e.����, e.תԺ, i.ת�� " & _
            " From ���ű� d, (Select ��Ժ����id, Count(*) As ��Ժ From ������ҳ Where ��Ժ���� Between [1] And [2] Group By ��Ժ����id) p, " & _
            "          (Select ����id, Sum(Decode(���ʱ��, Null, 1, 0)) As ������д, " & _
            "                             Sum(Decode(���ʱ��, Null, Decode(Sign(Sysdate - ����ʱ�� - 1), 1, 1, 0), 0)) As ��д��ʱ, " & _
            "                             Sum(Decode(���ʱ��, Null, 0, 1)) As �����, " & _
            "                             Sum(Decode(���ʱ��, Null, 0, Decode(NVL(ǩ������, 0), 0, 1, 0))) As �����޶� " & _
            "              From ���Ӳ�����¼ " & _
            "              Where  �������� = 4 And ����ʱ�� Between [1] And [2] " & _
            "              Group By ����id) l, " & _
            "          (Select b.Id, Count((a.����id)) As ת�� " & _
            "              From ���˱䶯��¼ a, ���ű� b " & _
            "              Where a.��ʼʱ�� Between [1] And [2] And a.��ʼԭ�� = 3 And Nvl(���Ӵ�λ, 0) = 0 And b.Id = a.����id " & _
            "              Group By b.Id) t, " & _
            "          (Select a.Id, Sum(Decode(b.��Ժ��ʽ, '����', 1, 0)) As ������Ժ, Sum(Decode(b.��Ժ��ʽ, '����', 1, 0)) As ����, " & _
            "                             Sum(Decode(b.��Ժ��ʽ, 'תԺ', 1, 0)) As תԺ " & _
            "              From ���ű� a, ������ҳ b " & _
            "              Where b.��Ժ���� Between [1] And [2] And b.��Ժ����id = a.Id " & _
            "              Group By a.Id) e, " & _
            "          (Select b.Id, Count((a.����id)) As ת�� " & _
            "              From ���˱䶯��¼ a, ���ű� b " & _
            "              Where a.��ֹʱ�� Between [1] And [2] And a.��ֹԭ�� = 3 And Nvl(���Ӵ�λ, 0) = 0 And b.Id = a.����id " & _
            "              Group By b.Id) i " & _
            " Where d.Id = p.��Ժ����id And p.��Ժ����id = l.����id(+) And t.Id = l.����id And e.Id = l.����id And i.Id = l.����id " & _
            " Order By d.����"
        Set Rs = OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"))
        Call InitGrid(mvar��������)
        Me.vfg������.Rows = 3 + Rs.RecordCount
        stbThis.Panels(2).Text = "���ƣ�" & Rs.RecordCount & "����¼��"
        i = 2
        Do While Not Rs.EOF
            With Me.vfg������
                .TextMatrix(i, mCol.ID) = NVL(Rs("ID"))
                .TextMatrix(i, mCol.����) = NVL(Rs("����"))
                .TextMatrix(i, mCol.����) = NVL(Rs("����"))
                .TextMatrix(i, mCol.������д) = IIf(NVL(Rs("������д")) = 0, "", NVL(Rs("������д"))): lngCount(1) = lngCount(1) + Val(.TextMatrix(i, mCol.������д))
                .TextMatrix(i, mCol.��д��ʱ) = IIf(NVL(Rs("��д��ʱ")) = 0, "", NVL(Rs("��д��ʱ"))): lngCount(2) = lngCount(2) + Val(.TextMatrix(i, mCol.��д��ʱ))
                .TextMatrix(i, mCol.�����) = IIf(NVL(Rs("�����")) = 0, "", NVL(Rs("�����"))): lngCount(3) = lngCount(3) + Val(.TextMatrix(i, mCol.�����))
                .TextMatrix(i, mCol.�����޶�) = IIf(NVL(Rs("�����޶�")) = 0, "", NVL(Rs("�����޶�"))): lngCount(4) = lngCount(4) + Val(.TextMatrix(i, mCol.�����޶�))
                .TextMatrix(i, mCol.��Ժ) = IIf(NVL(Rs("��Ժ")) = 0, "", NVL(Rs("��Ժ"))): lngCount(5) = lngCount(5) + Val(.TextMatrix(i, mCol.��Ժ))
                .TextMatrix(i, mCol.ת��) = IIf(NVL(Rs("ת��")) = 0, "", NVL(Rs("ת��"))): lngCount(6) = lngCount(6) + Val(.TextMatrix(i, mCol.ת��))
                .TextMatrix(i, mCol.������Ժ) = IIf(NVL(Rs("������Ժ")) = 0, "", NVL(Rs("������Ժ"))): lngCount(7) = lngCount(7) + Val(.TextMatrix(i, mCol.������Ժ))
                .TextMatrix(i, mCol.ת��) = IIf(NVL(Rs("ת��")) = 0, "", NVL(Rs("ת��"))): lngCount(8) = lngCount(8) + Val(.TextMatrix(i, mCol.ת��))
                .TextMatrix(i, mCol.תԺ) = IIf(NVL(Rs("תԺ")) = 0, "", NVL(Rs("תԺ"))): lngCount(9) = lngCount(9) + Val(.TextMatrix(i, mCol.תԺ))
                .TextMatrix(i, mCol.����) = IIf(NVL(Rs("����")) = 0, "", NVL(Rs("����"))): lngCount(10) = lngCount(10) + Val(.TextMatrix(i, mCol.����))
            End With
            Rs.MoveNext
            i = i + 1
        Loop
        With Me.vfg������
            .TextMatrix(i, mCol.����) = "�ϼ�"
            .TextMatrix(i, mCol.����) = ""
            .TextMatrix(i, mCol.������д) = lngCount(1)
            .TextMatrix(i, mCol.��д��ʱ) = lngCount(2)
            .TextMatrix(i, mCol.�����) = lngCount(3)
            .TextMatrix(i, mCol.�����޶�) = lngCount(4)
            .TextMatrix(i, mCol.��Ժ) = lngCount(5)
            .TextMatrix(i, mCol.ת��) = lngCount(6)
            .TextMatrix(i, mCol.������Ժ) = lngCount(7)
            .TextMatrix(i, mCol.ת��) = lngCount(8)
            .TextMatrix(i, mCol.תԺ) = lngCount(9)
            .TextMatrix(i, mCol.����) = lngCount(10)
            .Cell(flexcpFontBold, i, mCol.����) = True
            .Cell(flexcpFontBold, i, mCol.����) = True
            .Cell(flexcpFontBold, i, mCol.������д) = True
            .Cell(flexcpFontBold, i, mCol.��д��ʱ) = True
            .Cell(flexcpFontBold, i, mCol.�����) = True
            .Cell(flexcpFontBold, i, mCol.�����޶�) = True
            .Cell(flexcpFontBold, i, mCol.��Ժ) = True
            .Cell(flexcpFontBold, i, mCol.ת��) = True
            .Cell(flexcpFontBold, i, mCol.������Ժ) = True
            .Cell(flexcpFontBold, i, mCol.ת��) = True
            .Cell(flexcpFontBold, i, mCol.תԺ) = True
            .Cell(flexcpFontBold, i, mCol.����) = True
        End With
        Rs.Close
        Set Rs = Nothing
        If Me.vfg������.Rows > 2 Then Me.vfg������.Row = 2
    End Select
End Sub

Private Sub InitGrid(ByVal �������� As Enum��������)
    Dim i As Long
    Select Case ��������
    Case ���ﲡ��
        With Me.vfg���ﲡ��
            .Clear
            .Rows = 1
            .FixedRows = 1
            .Cols = 6
            .RowHeightMin = 300
            .WallPaper = imgBG.Picture
            .WallPaperAlignment = flexPicAlignRightBottom
            
    '        .BackColorAlternate = RGB(240, 240, 255)
            .BackColorSel = RGB(125, 125, 255)
            .ForeColorSel = vbWhite
            .Sort = flexSortCustom
            
            .TextMatrix(0, mCol.����) = "����"
            .TextMatrix(0, mCol.����) = "����"
            .TextMatrix(0, mCol.������д) = "������д������"
            .TextMatrix(0, mCol.��д��ʱ) = "���г�ʱ��д��"
            .TextMatrix(0, mCol.�����) = "����ɲ�����"
            
            For i = 0 To 5
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
            Next
            
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.����) = 600
            .ColWidth(mCol.����) = 2500
            .ColWidth(mCol.������д) = 1600
            .ColWidth(mCol.��д��ʱ) = 1600
            .ColWidth(mCol.�����) = 1600
        End With
    Case ������
        With Me.vfg������
            .Clear
            .Rows = 3
            .FixedRows = 2
            .Cols = 13
            .RowHeightMin = 300
            .WallPaper = imgBG.Picture
            .WallPaperAlignment = flexPicAlignRightBottom
            
    '        .BackColorAlternate = RGB(240, 240, 255)
            .BackColorSel = RGB(125, 125, 255)
            .ForeColorSel = vbWhite
            .Sort = flexSortCustom
            
            .TextMatrix(0, mCol.����) = "סԺ����"
            .TextMatrix(0, mCol.����) = "סԺ����"
            .TextMatrix(0, mCol.������д) = "������д�������ݣ�"
            .TextMatrix(0, mCol.��д��ʱ) = "������д�������ݣ�"
            .TextMatrix(0, mCol.�����) = "����ɲ������ݣ�"
            .TextMatrix(0, mCol.�����޶�) = "����ɲ������ݣ�"
            .TextMatrix(0, mCol.��Ժ) = "�����˴�"
            .TextMatrix(0, mCol.ת��) = "�����˴�"
            .TextMatrix(0, mCol.������Ժ) = "�����˴�"
            .TextMatrix(0, mCol.ת��) = "�����˴�"
            .TextMatrix(0, mCol.תԺ) = "�����˴�"
            .TextMatrix(0, mCol.����) = "�����˴�"
            .TextMatrix(1, mCol.����) = "����"
            .TextMatrix(1, mCol.����) = "����"
            .TextMatrix(1, mCol.������д) = "����"
            .TextMatrix(1, mCol.��д��ʱ) = "��ʱ24Сʱ"
            .TextMatrix(1, mCol.�����) = "����"
            .TextMatrix(1, mCol.�����޶�) = "�����޶�"
            .TextMatrix(1, mCol.��Ժ) = "��Ժ"
            .TextMatrix(1, mCol.ת��) = "����ת��"
            .TextMatrix(1, mCol.������Ժ) = "������Ժ"
            .TextMatrix(1, mCol.ת��) = "ת��"
            .TextMatrix(1, mCol.תԺ) = "תԺ"
            .TextMatrix(1, mCol.����) = "����"
            
            .MergeRow(0) = True
            .MergeCells = flexMergeRestrictRows
            
            For i = 0 To 12
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next
            
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.����) = 600
            .ColWidth(mCol.����) = 1600
            .ColWidth(mCol.������д) = 600
            .ColWidth(mCol.��д��ʱ) = 1200
            .ColWidth(mCol.�����) = 600
            .ColWidth(mCol.�����޶�) = 1200
            .ColWidth(mCol.��Ժ) = 900
            .ColWidth(mCol.ת��) = 900
            .ColWidth(mCol.������Ժ) = 900
            .ColWidth(mCol.ת��) = 600
            .ColWidth(mCol.תԺ) = 600
            .ColWidth(mCol.����) = 600
        End With
    Case סԺ����
        With Me.vfgסԺ����
            .Clear
            .Rows = 3
            .FixedRows = 2
            .Cols = 13
            .RowHeightMin = 300
            .WallPaper = imgBG.Picture
            .WallPaperAlignment = flexPicAlignRightBottom
            
    '        .BackColorAlternate = RGB(240, 240, 255)
            .BackColorSel = RGB(125, 125, 255)
            .ForeColorSel = vbWhite
            .Sort = flexSortCustom
            
            .TextMatrix(0, mCol.����) = "סԺ����"
            .TextMatrix(0, mCol.����) = "סԺ����"
            .TextMatrix(0, mCol.������д) = "������д�������ݣ�"
            .TextMatrix(0, mCol.��д��ʱ) = "������д�������ݣ�"
            .TextMatrix(0, mCol.�����) = "����ɲ������ݣ�"
            .TextMatrix(0, mCol.�����޶�) = "����ɲ������ݣ�"
            .TextMatrix(0, mCol.��Ժ) = "�����˴�"
            .TextMatrix(0, mCol.ת��) = "�����˴�"
            .TextMatrix(0, mCol.������Ժ) = "�����˴�"
            .TextMatrix(0, mCol.ת��) = "�����˴�"
            .TextMatrix(0, mCol.תԺ) = "�����˴�"
            .TextMatrix(0, mCol.����) = "�����˴�"
            .TextMatrix(1, mCol.����) = "����"
            .TextMatrix(1, mCol.����) = "����"
            .TextMatrix(1, mCol.������д) = "����"
            .TextMatrix(1, mCol.��д��ʱ) = "��ʱ24Сʱ"
            .TextMatrix(1, mCol.�����) = "����"
            .TextMatrix(1, mCol.�����޶�) = "�����޶�"
            .TextMatrix(1, mCol.��Ժ) = "��Ժ"
            .TextMatrix(1, mCol.ת��) = "����ת��"
            .TextMatrix(1, mCol.������Ժ) = "������Ժ"
            .TextMatrix(1, mCol.ת��) = "ת��"
            .TextMatrix(1, mCol.תԺ) = "תԺ"
            .TextMatrix(1, mCol.����) = "����"
            
            .MergeRow(0) = True
            .MergeCells = flexMergeRestrictRows
            
            For i = 0 To 12
                .Cell(flexcpAlignment, 0, i) = flexAlignCenterCenter
                .Cell(flexcpAlignment, 1, i) = flexAlignCenterCenter
            Next
            
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.����) = 600
            .ColWidth(mCol.����) = 1600
            .ColWidth(mCol.������д) = 600
            .ColWidth(mCol.��д��ʱ) = 1200
            .ColWidth(mCol.�����) = 600
            .ColWidth(mCol.�����޶�) = 1200
            .ColWidth(mCol.��Ժ) = 900
            .ColWidth(mCol.ת��) = 900
            .ColWidth(mCol.������Ժ) = 900
            .ColWidth(mCol.ת��) = 600
            .ColWidth(mCol.תԺ) = 600
            .ColWidth(mCol.����) = 600
        End With
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_Exit: Unload Me
    
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
        Call cmdSearch_Click
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case ID_ViewFile
        If mvar�������� = סԺ���� Then
            If Me.vfgסԺ����.Row = 0 Or Me.vfgסԺ����.Row = Me.vfgסԺ����.Rows - 1 Then Exit Sub
            mfrmQualityViewFile.ShowMe mvar��������, Me, Me.vfgסԺ����.TextMatrix(Me.vfgסԺ����.Row, mCol.ID), Me.vfgסԺ����.TextMatrix(Me.vfgסԺ����.Row, mCol.����), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        ElseIf mvar�������� = ���ﲡ�� Then
            If Me.vfg���ﲡ��.Row = 0 Or Me.vfg���ﲡ��.Row = Me.vfg���ﲡ��.Rows - 1 Then Exit Sub
            mfrmQualityViewFile.ShowMe mvar��������, Me, Me.vfg���ﲡ��.TextMatrix(Me.vfg���ﲡ��.Row, mCol.ID), Me.vfg���ﲡ��.TextMatrix(Me.vfg���ﲡ��.Row, mCol.����), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        Else
            If Me.vfg������.Row = 0 Or Me.vfg������.Row = Me.vfg������.Rows - 1 Then Exit Sub
            mfrmQualityViewFile.ShowMe mvar��������, Me, Me.vfg������.TextMatrix(Me.vfg������.Row, mCol.ID), Me.vfg������.TextMatrix(Me.vfg������.Row, mCol.����), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        End If
    Case ID_ViewPati
        If mvar�������� = סԺ���� Then
            If Me.vfgסԺ����.Row = 0 Or Me.vfgסԺ����.Row = Me.vfgסԺ����.Rows - 1 Then Exit Sub
            mfrmQualityViewPati.ShowMe mvar��������, Me, Me.vfgסԺ����.TextMatrix(Me.vfgסԺ����.Row, mCol.ID), Me.vfgסԺ����.TextMatrix(Me.vfgסԺ����.Row, mCol.����), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        ElseIf mvar�������� = ���ﲡ�� Then
            If Me.vfg���ﲡ��.Row = 0 Or Me.vfg���ﲡ��.Row = Me.vfg���ﲡ��.Rows - 1 Then Exit Sub
            mfrmQualityViewPati.ShowMe mvar��������, Me, Me.vfg���ﲡ��.TextMatrix(Me.vfg���ﲡ��.Row, mCol.ID), Me.vfg���ﲡ��.TextMatrix(Me.vfg���ﲡ��.Row, mCol.����), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        Else
            If Me.vfg������.Row = 0 Or Me.vfg������.Row = Me.vfg������.Rows - 1 Then Exit Sub
            mfrmQualityViewPati.ShowMe mvar��������, Me, Me.vfg������.TextMatrix(Me.vfg������.Row, mCol.ID), Me.vfg������.TextMatrix(Me.vfg������.Row, mCol.����), _
            Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
        End If
    End Select
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.cbsThis.GetClientRect Left, Top, Right, Bottom
    With Me.tplThis
        .Left = Left: .Width = 3050
        .Top = Top: .Height = Bottom - Top - stbThis.Height
    End With
    Me.tbcThis.Move Me.tplThis.Width, Top, Right - Left - Me.tplThis.Width, Me.tplThis.Height
'    vfg���ﲡ��.Move 0, 0, tbcThis.Width, tbcThis.Height
'    vfgסԺ����.Move 0, 0, tbcThis.Width, tbcThis.Height
'    vfg������.Move 0, 0, tbcThis.Width, tbcThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vfgסԺ����.Rows <> 0)
    Case conMenu_View_Jump '��ת
        If Me.tbcThis.Selected.Index + 1 <= Me.tbcThis.ItemCount - 1 Then
            Me.tbcThis.Item(Me.tbcThis.Selected.Index + 1).Selected = True
        Else
            Me.tbcThis.Item(0).Selected = True
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case ID_ViewFile:
        Select Case mvar��������
        Case ���ﲡ��
            Control.Enabled = Me.vfg���ﲡ��.Row > 0 And Me.vfg���ﲡ��.Row < Me.vfg���ﲡ��.Rows - 1
        Case סԺ����
            Control.Enabled = Me.vfgסԺ����.Row > 1 And Me.vfgסԺ����.Row < Me.vfgסԺ����.Rows - 1
        Case ������
            Control.Enabled = Me.vfg������.Row > 1 And Me.vfg������.Row < Me.vfg������.Rows - 1
        End Select
    Case ID_ViewPati:
        Select Case mvar��������
        Case ���ﲡ��
            Control.Enabled = False 'Me.vfg���ﲡ��.Row > 0 And Me.vfg���ﲡ��.Row < Me.vfg���ﲡ��.Rows - 1
        Case סԺ����
            Control.Enabled = Me.vfgסԺ����.Row > 1 And Me.vfgסԺ����.Row < Me.vfgסԺ����.Rows - 1
        Case ������
            Control.Enabled = Me.vfg������.Row > 1 And Me.vfg������.Row < Me.vfg������.Rows - 1
        End Select
    End Select
End Sub

Private Sub cmdSearch_Click()
    FillGrid Format(Me.dtpDateFrom.Value, "yyyy-MM-dd"), Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
End Sub

Private Sub dtpDateFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    Dim cbpPopup As CommandBarPopup                     '��ʱ����
    Dim cbpPopupSub As CommandBarPopup                  '��ʱ����
    Dim objControl As CommandBarControl                 '�������ؼ�
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�
    Dim Combo As CommandBarComboBox                     '������������ؼ�
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
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
    cbsThis.ActiveMenuBar.Title = "�˵���"
    
    Set cbp�ļ� = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ļ�(&F)")
    With cbp�ļ�.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "������&Excel")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With
    
    Set cbp��ͼ = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbp��ͼ.ID = conMenu_ViewPopup
    With cbp��ͼ.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set objControl = .Add(xtpControlButton, ID_ViewFile, "�ļ��б���ͼ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ViewPati, "����Ժ������ͼ")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
    End With
    
    Set cbp���� = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbp����.ID = conMenu_HelpPopup
    With cbp����.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    Set Bar���� = cbsThis.Add("���ù�����", xtpBarTop)
    With Bar����.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ViewFile, "�ļ��б���ͼ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_ViewPati, "����Ժ������ͼ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.BeginGroup = True
    End With
    For Each objControl In Bar����.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next
    
    '�ȼ���
    cbsThis.KeyBindings.Add FCONTROL, Asc("Q"), conMenu_File_Exit
    cbsThis.KeyBindings.Add FCONTROL, Asc("P"), conMenu_File_Print
    cbsThis.KeyBindings.Add 0, vbKeyF5, conMenu_View_Refresh
    cbsThis.KeyBindings.Add 0, vbKeyF6, conMenu_View_Jump  '��ת
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet '��ӡ����
        .AddHiddenCommand conMenu_File_Excel '�����Excel
        .AddHiddenCommand conMenu_View_Jump '��ת
    End With

    Set Group = tplThis.Groups.Add(0, "ͳ������")
    Group.Special = True
    Set Item = Group.Items.Add(0, "��д���ڷ�Χ:", xtpTaskItemTypeText)
    Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
    Set Item.Control = Me.picDate
    Me.picDate.BackColor = Item.BackColor
        
    Set Group = tplThis.Groups.Add(0, "��������")
    Group.Tooltip = "���������б�"
    Group.Items.Add conMenu_File_Excel, "������Excel", xtpTaskItemTypeLink, 1
    Group.Items.Add conMenu_File_Preview, "��ӡԤ��", xtpTaskItemTypeLink, 2
    Group.Items.Add conMenu_File_Print, "��ӡ...", xtpTaskItemTypeLink, 3
    
    Set Group = tplThis.Groups.Add(0, "ͳ����Ϣ")
    Group.Tooltip = "ͳ�ƽ������"
    Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
    Set Item.Control = Me.picInfo
    Me.picInfo.BackColor = Item.BackColor
    
    With Me.tbcThis
        .RemoveAll
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = False
            .ShowIcons = True
        End With
        .InsertItem(0, "���ﲡ��", vfg���ﲡ��.hWnd, 0).Tag = "���ﲡ��"
        .InsertItem(1, "סԺ����", vfgסԺ����.hWnd, 0).Tag = "סԺ����"
        .InsertItem(2, "������", vfg������.hWnd, 0).Tag = "������"
    End With
    
    tplThis.SetImageList imlTaskPanelIcons
    Call RestoreWinState(Me)
    
    '-----------------------------------------------------
    '��������װ��:
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select Sysdate From Dual"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
    With Me.dtpDateTo
        .Value = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd")
        .MaxDate = .Value: .MinDate = Format("1990-01-01", "yyyy-MM-dd")
    End With
    With Me.dtpDateFrom
        .Value = Me.dtpDateTo.Value - 7
        .MaxDate = Me.dtpDateTo.MaxDate: .MinDate = Me.dtpDateTo.MinDate
    End With
    rsTemp.Close
    Set rsTemp = Nothing
    
    mvar�������� = ���ﲡ��
    
    Call cmdSearch_Click        '�����ʼ����
End Sub

Private Sub dtpDateTo_Validate(Cancel As Boolean)
    Me.dtpDateFrom.MaxDate = Me.dtpDateTo.Value
    If Me.dtpDateFrom.Value > Me.dtpDateFrom.MaxDate Then Me.dtpDateFrom.Value = Me.dtpDateFrom.MaxDate
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsThis_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me)
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If tbcThis.Tag <> "" Then Exit Sub
    Select Case Item.Tag
    Case "���ﲡ��"
        mvar�������� = ���ﲡ��
    Case "סԺ����"
        mvar�������� = סԺ����
    Case "������"
        mvar�������� = ������
    End Select
    Call cmdSearch_Click
End Sub

Private Sub tplThis_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Select Case Item.ID
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    End Select
End Sub

Private Sub vfg������_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tbcThis.Tag = "Moving"
    tbcThis.Item(0).Selected = True
    tbcThis.Item(2).Selected = True
    tbcThis.Tag = ""
End Sub

Private Sub vfg���ﲡ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tbcThis.Tag = "Moving"
    tbcThis.Item(2).Selected = True
    tbcThis.Item(0).Selected = True
    tbcThis.Tag = ""
End Sub

Private Sub vfgסԺ����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tbcThis.Tag = "Moving"
    tbcThis.Item(0).Selected = True
    tbcThis.Item(1).Selected = True
    tbcThis.Tag = ""
End Sub

Private Sub vfgסԺ����_RowColChange()
    Dim i As Long
    With vfgסԺ����
        i = .Row
        If i > 1 And i < .Rows - 1 Then
            lblMC = .TextMatrix(i, 1)
            lblBM = .TextMatrix(i, 0)
            lblRYRS = Val(.TextMatrix(i, 6)) & " ��"
            lblZZSX = Val(.TextMatrix(i, 2)) & " ��"
            lblSXCS = Val(.TextMatrix(i, 3)) & " ��"
            lblSXWC = Val(.TextMatrix(i, 4)) & " ��"
            lblZZXD = Val(.TextMatrix(i, 5)) & " ��"
        Else
            lblMC = "-"
            lblBM = "-"
            lblRYRS = "-" & " ��"
            lblZZSX = "-" & " ��"
            lblSXCS = "-" & " ��"
            lblSXWC = "-" & " ��"
            lblZZXD = "-" & " ��"
        End If
    End With
    If mvar�������� = ���ﲡ�� Then
        lblZZXD.Visible = False
    Else
        lblZZXD.Visible = True
    End If
End Sub
