VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathTableOut 
   BorderStyle     =   0  'None
   Caption         =   "�����ٴ�·����"
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10020
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraPath 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   380
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10815
      Begin VB.ComboBox cboPath 
         Height          =   300
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "·������"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblInDiag 
         BackColor       =   &H8000000E&
         Caption         =   "������ϣ�"
         Height          =   255
         Left            =   10080
         TabIndex        =   11
         Top             =   120
         Width           =   4995
      End
      Begin VB.Label lblOutDate 
         BackColor       =   &H8000000E&
         Caption         =   "����ʱ�䣺3000-01-01 00:01"
         Height          =   255
         Left            =   7620
         TabIndex        =   10
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblInDate 
         BackColor       =   &H8000000E&
         Caption         =   "����ʱ�䣺3000-01-01 00:00"
         Height          =   255
         Left            =   5160
         TabIndex        =   9
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblInPep 
         BackColor       =   &H8000000E&
         Caption         =   "�����ˣ�***"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
   End
   Begin zlCISPath.UCAdviceList UCAdvice 
      Height          =   1455
      Left            =   0
      TabIndex        =   4
      Top             =   5760
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   2566
   End
   Begin VB.Frame fraline 
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   5640
      Width           =   8175
   End
   Begin MSComctlLib.ImageList imgCharacter 
      Left            =   8280
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":0000
            Key             =   "�Ѿ�ִ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":059A
            Key             =   "��δִ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":0B34
            Key             =   "ȡ��ִ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":10CE
            Key             =   "����ִ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":1668
            Key             =   "��ǰִ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":1C02
            Key             =   "�Ӻ�ִ��"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   8400
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   225
      Begin VB.Image imgMore 
         Height          =   225
         Left            =   0
         Picture         =   "frmPathTableOut.frx":219C
         Top             =   0
         Width           =   225
      End
   End
   Begin MSComctlLib.ImageList imgFlow 
      Left            =   8280
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":259D
            Key             =   "node"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":26E4
            Key             =   "currnode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":2833
            Key             =   "multnode"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":29B5
            Key             =   "currmultnode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":2B7B
            Key             =   "arrow"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":2FFE
            Key             =   "arrowlate"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":3479
            Key             =   "arrow_Branch"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTableOut.frx":3899
            Key             =   "arrowlate_Branch"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPath 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "˫���鿴·����Ŀ����"
      Top             =   2400
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
      Appearance      =   0
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTableOut.frx":3CBD
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsFlow 
      Height          =   1920
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "˫���鿴·���׶ζ���"
      Top             =   390
      Width           =   8175
      _cx             =   14420
      _cy             =   3387
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483634
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   1800
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTableOut.frx":3DF8
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   101
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPathPrint 
      Height          =   3105
      Index           =   0
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "˫���鿴·����Ŀ����"
      Top             =   -99999
      Visible         =   0   'False
      Width           =   8175
      _cx             =   14420
      _cy             =   5477
      Appearance      =   0
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   3
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathTableOut.frx":3E69
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbsSub 
      Left            =   8880
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathTableOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)     'Ҫ��鿴����
Public Event Activate()                                                         '���Ѽ���ʱ
Public Event RequestRefresh(ByVal lngPathState As Long)                         'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String)                             'Ҫ�����������״̬������
Private Const C_Exe = "��"                                                      '��
Private Const CON_SmallFontSize As Long = 9                                     'С����
Private Const CON_BigFontSize As Long = 12                                      '������
Private Const CON_PathOutItemColor As Long = &HC0FFFF                           '·������Ŀ��ǳ��ɫ
Private Const CON_PathOutItemColorBlue As Long = &HFAEADA                       '�ݴ�·������Ŀ,ǳ��ɫ��ʶ
Private Const C_UnExe = "��"

Private Enum EFixedRow
    R0�׶��� = 0
    R1���� = 1
    R2���� = 2
End Enum

Private Enum PatiType
    pt���� = 0
    pt���� = 1
    pt���� = 2
    ptת�� = 3
    ptԤԼ = 4
    pt���� = 5
    pt�Ŷӽк� = 6
End Enum

Private mfrmParent          As Object
Private mcbsMain            As Object
Private mobjPublicPACS      As Object

Private mPP                 As TYPE_PATH_Pati
Private mPati               As TYPE_Pati
Private mcolReason          As Collection

Private mbln����ִ�л���    As Boolean                  '�Ƿ�����·��ִ�л���
Private mblnUnChange        As Boolean                  '�����õ�Ԫ��仯�¼���ˢ�µ�Ԫ������
Private mblnInOverScope     As Boolean                  '���˵�ǰִ�������Ƿ��ڱ�׼����ʱ�䷶Χ���������·����
Private mlngFontSize        As Long                     '���������С
Private mlngPathCount As Long   '����סԺ��·����

Private Sub SetUnImport()
'���ܣ�����δ����ʱ��״̬����Ϣ
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .ForeColorSel = vbBlack
        .TextMatrix(0, 0) = "  �ò���δ���������ٴ�·����"
    End With
    Call ClearPathItem
End Sub

Private Sub SetImportFalse()
'���ܣ����õ����˵��������ٴ�·��ʧ��ʱ��״̬����Ϣ
    With vsFlow
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 5000
        .TextMatrix(0, 0) = "  �ò��˲�����·������������" & vbCrLf & "  ԭ��" & mPP.δ����ԭ��
        .AutoSize 0
        .ForeColorSel = &HC0&
        If .Visible And .Enabled Then .SetFocus
    End With
    Call ClearPathItem
End Sub

Private Sub ClearPathItem(Optional blnImported As Boolean)
'���ܣ�������û�п��õ������ٴ�·��ʱ���·������Ŀ
    With vsPath
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 0
        .Cols = 0
        If blnImported Then
            .Rows = 1
            .Cols = 1
            .TextMatrix(0, 0) = vbCrLf & "  �ò��˻�û������·����Ŀ��"
            .Select 0, 0
            .CellAlignment = flexAlignLeftTop
        End If
    End With
End Sub

Private Sub cboPath_Click()
    If cboPath.ListIndex >= 0 Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬, False, , Val(cboPath.ItemData(cboPath.ListIndex)))
    End If
End Sub

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlExecuteCommandBars(Control)
End Sub

Private Sub cbsSub_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Call zlPopupCommandBars(CommandBar)
End Sub

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If fraPath.Visible Then
        fraPath.Top = lngTop
        fraPath.Left = lngLeft
        fraPath.Width = lngRight - lngLeft
        lngTop = fraPath.Top + fraPath.Height
    End If
    vsFlow.Left = lngLeft
    vsFlow.Top = lngTop
    vsFlow.Width = lngRight - lngLeft
    vsFlow.Height = 1140

    If vsPath.FixedRows = 0 And vsPath.Rows = 0 Then  'û�е���·��ʱ
        vsFlow.Height = Me.Height
        vsPath.Visible = False
        UCAdvice.Visible = False
        fraline.Visible = False
    Else
        If vsPath.Visible = False Then vsPath.Visible = True
        If UCAdvice.Visible = False Then UCAdvice.Visible = True
        If fraline.Visible = False Then fraline.Visible = True
        
        With vsPath
            .Top = lngTop + vsFlow.Height
            .Width = lngRight - lngLeft
            If lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30 > 0 Then
                .Height = lngBottom - lngTop - vsFlow.Height - IIf(UCAdvice.Visible, UCAdvice.Height + fraline.Height, 0) - 30
            Else
                .Height = lngBottom - lngTop - vsFlow.Height
            End If

            If .FixedRows = 0 And .Rows = 1 Then             'û��������Ŀ
                .ColWidth(0) = .Width - 30
                .RowHeight(0) = .Height
            End If
            fraline.Top = .Top + .Height
            fraline.Width = .Width

            UCAdvice.Top = fraline.Top + fraline.Height
            UCAdvice.Width = .Width
        End With
    End If

    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub fraline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        If vsPath.Height + Y < 1000 Or vsPath.Height - Y < 500 Then Exit Sub
        If UCAdvice.Height + Y < 250 Or UCAdvice.Height - Y < 500 Then Exit Sub

        If fraMore.Visible Then fraMore.Visible = False

        fraline.Top = fraline.Top + Y
        vsPath.Height = vsPath.Height + Y
        UCAdvice.Top = UCAdvice.Top + Y
        UCAdvice.Height = UCAdvice.Height - Y
    End If
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call zlUpdateCommandBars(Control)
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    Call InitCbsSubBar
End Sub

Private Sub Form_Resize()
    Call cbsSub_Resize
End Sub

Private Sub LoadPathFlow()
'���ܣ����ݲ��˵����·�������·��������Ϣ������
    Dim strSql As String, i As Long, j As Long, lngCurCol As Long
    Dim rsTmp As ADODB.Recordset, lngDayMin As Long, lngDayMax As Long
    Dim lng�������� As Long
    Dim lng��� As Long
    Dim str��׼����ʱ�� As String
    
    On Error GoTo errH
    
    With vsFlow
        .Clear
        .Rows = 1: .Cols = 1
        .ForeColorSel = vbBlack
        mblnInOverScope = False

        strSql = " Select a.ID,a.���� �׶���,Decode(a.��������, Null, 0, 1) ����,b.����,b.���� ·����,b.���°汾,c.��׼����ʱ��" & _
                 " From ����·���׶� a,����·��Ŀ¼ b,����·���汾 c " & _
                 " Where a.·��id = [1] And a.�汾�� = [2] And a.·��id=b.id And a.��ID is null And b.id = c.·��id And a.�汾�� = c.�汾�� " & _
                 " Order by a.���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.·��ID, mPP.�汾��)
        str��׼����ʱ�� = NVL(rsTmp!��׼����ʱ��)
        If rsTmp.RecordCount > 0 Then
            .Rows = 1
            .Cols = rsTmp.RecordCount * 2                         '��һ��Ϊ·��������ͷΪ�׶���-1
            .Select 0, 0
            .RowHeight(0) = 1100

            '��һ����ʾ·������
            .ColWidth(0) = 2800
            If mPP.����·��״̬ > 0 Then
                .TextMatrix(0, 0) = rsTmp!·���� & ""

                If mPP.����·��״̬ = 3 Then
                    .Cell(flexcpForeColor, 0, 0) = vbRed
                End If
            Else
                .TextMatrix(0, 0) = rsTmp!·���� & ""
            End If
            
            If mPP.��ǰ���� > 0 And mPP.����·��״̬ = 1 Then
            
                '��ȡ��׼����ʱ��
                If InStr(str��׼����ʱ��, "-") > 0 Then
                    j = Split(str��׼����ʱ��, "-")(1)
                    lngDayMin = Val(Split(str��׼����ʱ��, "-")(0))
                    lngDayMax = j
                Else
                    j = Val(str��׼����ʱ��)                                'С�ڵ���n������
                    lngDayMin = 1
                    lngDayMax = j
                End If

                lng�������� = GetMustDayOut(mPP.����·��ID, mPP.��ǰ����)

                i = Format(lng�������� / j * 100, "0")
                If i = 100 And lng�������� <> j Then
                    i = 99
                End If
                
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "���ȣ�" & i & "%"

                If lng�������� > lngDayMax Then
                    mblnInOverScope = True
                Else
                    mblnInOverScope = Between(lng��������, lngDayMin, lngDayMax)
                End If
            End If
            
            If mPP.����·��״̬ > 0 Then
                .TextMatrix(0, 0) = .TextMatrix(0, 0) & vbCrLf & "״̬��" & IIf(mPP.����·��״̬ = 1, "ִ����", IIf(mPP.����·��״̬ = 2, "���", "�����˳�"))
            End If
            .Cell(flexcpTextStyle, 0, 0) = 3

            For i = 1 To .Cols Step 2
                .TextMatrix(0, i) = " " & rsTmp!�׶��� & " "            '���ñ߾�
                .ColAlignment(i) = flexAlignCenterCenter

                .ColWidth(i) = 1750
                .Col = i
                .PicturesOver = True
                .CellPictureAlignment = flexPicAlignLeftCenter
                If mPP.��ǰ�׶�ID = rsTmp!ID Or mPP.�׶θ�ID = rsTmp!ID Or (mPP.��ǰ�׶�ID = 0 And i = 1 And mPP.����·��״̬ = 1) Then
                    lngCurCol = i
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!���� = 1, "currmultnode", "currnode")).Picture
                    Call .ShowCell(0, i)
                Else
                    .CellPicture = imgFlow.ListImages(IIf(rsTmp!���� = 1, "multnode", "node")).Picture
                End If
                .ColData(i) = Val(rsTmp!ID)

                rsTmp.MoveNext

                '��ͷ
                If i < .Cols - 1 Then
                    .ColWidth(i + 1) = 550
                    .Col = i + 1
                    .CellPictureAlignment = flexPicAlignCenterCenter
                    .CellPicture = imgFlow.ListImages(IIf(i + 1 > lngCurCol And lngCurCol <> 0 Or mPP.����·��״̬ > 1, "arrowlate", "arrow")).Picture
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Getִ�н������ͼ��(ByVal lngִ�н������ As Long) As Long
'���ܣ�����ִ�н�����ʷ��ض�Ӧ��ͼ�����
'1-�Ѿ�ִ�У�2-��δִ�У�3-ȡ��ִ�У�4-����ִ�У�5-��ǰִ�У�6-�Ӻ�ִ��
    Dim lngIdx As Long
    Select Case lngִ�н������
        Case 1
            lngIdx = imgCharacter.ListImages("�Ѿ�ִ��").Index
        Case 2
            lngIdx = imgCharacter.ListImages("��δִ��").Index
        Case 3
            lngIdx = imgCharacter.ListImages("ȡ��ִ��").Index
        Case 4
            lngIdx = imgCharacter.ListImages("����ִ��").Index
        Case 5
            lngIdx = imgCharacter.ListImages("��ǰִ��").Index
        Case 6
            lngIdx = imgCharacter.ListImages("�Ӻ�ִ��").Index
    End Select
    Getִ�н������ͼ�� = lngIdx
End Function

Private Sub LoadPathItem()
'���ܣ����ز��������ɵ�·����Ŀ
    Dim strSql As String, strOldType As String, str������� As String
    Dim lngRow As Long, lngCol As Long, i As Long, j As Long, arrtmp As Variant, strTmp As String
    Dim rsTmp As ADODB.Recordset
    Dim CPos As New Collection  'ÿ���������ʼ��
    Dim lngPreRow As Long, lngPreCol As Long, lngDayRow As Long
    Dim rsSort As Recordset
    Dim str����ԭ�� As String

    With vsPath
        lngPreRow = -1
        lngPreCol = -1
        If .Row >= .FixedRows Then lngPreRow = .Row
        If .Col >= .FixedCols Then lngPreCol = .Col

        '1)���ಿ��
        .Redraw = flexRDNone
        mblnUnChange = True
        .Clear
        .Rows = 3: .FixedRows = 3
        .Cols = 1: .FixedCols = 1
        mblnUnChange = False
        .MergeCol(0) = True
        .MergeRow(0) = True

        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 0, .FixedRows - 1, 0) = "ʱ��׶�"
        On Error GoTo errH
        
        '��ȡ��������ƺ͸���
        strSql = " Select ����, Max(����) As ����,100 as ���" & vbNewLine & _
                 " From (Select Count(a.Id) As ����, a.����, a.�׶�id, a.����" & vbNewLine & _
                 "       From ��������·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1]" & vbNewLine & _
                 "       Group By a.����, a.����, a.�׶�id)" & vbNewLine & _
                 " Group By ����"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        Set rsTmp = zlDatabase.CopyNewRec(rsTmp)

        '��ȡ���
        strSql = " Select ����, ���" & vbNewLine & _
                 " From (Select ����, ���, Row_Number() Over(Partition By ���� Order By 1) As Top" & vbNewLine & _
                 "       From (Select a.���, a.���� As ����" & vbNewLine & _
                 "              From ����·������ A, ��������·��ִ�� B, ����·����Ŀ C" & vbNewLine & _
                 "              Where a.���� = c.���� And b.·����¼id = [1] And b.��Ŀid = c.Id And c.·��id = a.·��id And c.�汾�� = a.�汾��" & vbNewLine & _
                 "              Union" & vbNewLine & _
                 "              Select a.���, a.���� As ����" & vbNewLine & _
                 "              From ����·������ A, ��������·��ִ�� B, ����·���׶� C" & vbNewLine & _
                 "              Where a.���� = b.���� And b.�׶�id + 0 = c.Id And b.·����¼id = [1] And b.��Ŀid Is Null And a.·��id = c.·��id And" & vbNewLine & _
                 "                    a.�汾�� = c.�汾��))" & vbNewLine & _
                 " Where Top = 1" & vbNewLine & _
                 " Order By ���"
        Set rsSort = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        '����
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                rsSort.Filter = "����='" & rsTmp!���� & "'"
                If rsSort.RecordCount > 0 Then
                    rsTmp!��� = Val(rsSort!��� & "")
                    rsTmp.Update
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Sort = "���"
            rsTmp.MoveFirst
        End If
        
        For i = 1 To rsTmp.RecordCount
            CPos.Add .Rows, "T" & rsTmp!����
            .Rows = .Rows + rsTmp!����
            For j = 1 To rsTmp!����
                .TextMatrix(.Rows - j, .FixedCols - 1) = rsTmp!����
            Next
            rsTmp.MoveNext
        Next

        '2)ʱ��׶β���
        '�׶�����ʱ�� NVL(c.���,b.���) ��Ϊ�˴����÷�֧������������⣬ȡֵb.��� ����Ϊ��������Ҫ��ʾ�ǵڼ�����֧��
        strSql = "Select a.�׶�id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, To_Char(a.����, 'day') ����, b.���� As �׶���, b.���, b.˵��, b.��id,Decode(g.·��id,b.·��id,1,0) as ����" & vbNewLine & _
                 "From (Select a.�׶�id, a.����, a.����,a.·����¼id" & vbNewLine & _
                 "       From ��������·��ִ�� A" & vbNewLine & _
                 "       Where a.·����¼id = [1]" & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����,a.·����¼id) A, ����·���׶� B,����·���׶� C,��������·�� G" & vbNewLine & _
                 "Where a.�׶�id = b.Id And b.��id=c.id(+) And g.id=A.·����¼ID " & vbNewLine & _
                 "Order By ����,����, NVL(c.���,b.���)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .Cols = .Cols + rsTmp.RecordCount
        
        For i = 1 To rsTmp.RecordCount
            .ColWidth(i) = 2800
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColData(i) = Val("" & rsTmp!�׶�ID)
            If IsNull(rsTmp!��ID) Then
                .TextMatrix(EFixedRow.R0�׶���, i) = Replace(rsTmp!�׶���, vbLf, vbCrLf)                                'Ϊ�˴�ӡʱ��������(vbLfʱ������ʾ������)
            Else
                .TextMatrix(EFixedRow.R0�׶���, i) = Replace(rsTmp!�׶���, vbLf, vbCrLf) & ",��֧:" & NVL(rsTmp!˵��, rsTmp!���)
            End If
            .TextMatrix(EFixedRow.R1����, i) = "��" & rsTmp!���� & "��"
            .Cell(flexcpData, EFixedRow.R1����, i) = rsTmp!����
            .TextMatrix(EFixedRow.R2����, i) = rsTmp!���� & "(" & rsTmp!���� & ")"
            .Cell(flexcpData, EFixedRow.R2����, i) = rsTmp!���� & ""
            
            If rsTmp!���� = mPP.��ǰ���� Then
                mPP.��ǰ���� = rsTmp!����
            End If
            rsTmp.MoveNext
        Next

        For i = 1 To mcolReason.count
            mcolReason.Remove 1                 'ɾ���ֲ���������(�����ƶ���,���¼���ʱ��Ҫ��ձ���ԭ��)
        Next i
        
        '3)·����Ŀ����
        strSql = " Select a.Id, Nvl(b.ͼ��id, a.ͼ��id) ͼ��id, a.����, To_Char(a.����, 'yyyy-mm-dd') ����, a.����, a.�׶�id, Nvl(a.��Ŀ���, b.��Ŀ���) As ��Ŀ���," & vbNewLine & _
                 " Nvl(b.��Ŀ����, a.��Ŀ����) ��Ŀ����, a.��Ŀid, Decode(a.ִ����, Null, 0, 1) ִ��״̬, Nvl(b.ִ�з�ʽ, 1) ִ�з�ʽ, a.���ԭ��, c.���� As ����ԭ��," & vbNewLine & _
                 " Nvl(b.��Ŀ���, a.��Ŀ���) As ��Ŀ���, a.ִ�н��, d.·��id " & vbNewLine & _
                 " From ��������·��ִ�� A, ����·����Ŀ B, ������쳣��ԭ�� C, ����·���׶� D" & vbNewLine & _
                 " Where a.·����¼id = [1] And a.��Ŀid = b.Id(+) And a.����ԭ�� = c.����(+) And a.�׶�id + 0 = d.Id" & vbNewLine & _
                 " Order By a.����,����,��Ŀ���"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        
        For lngCol = .FixedCols To .Cols - 1
            rsTmp.Filter = "�׶�ID='" & .ColData(lngCol) & "' And ����=" & Val(Replace(.TextMatrix(EFixedRow.R1����, lngCol), "��", ""))
            strOldType = ""

            Do While Not rsTmp.EOF
                If strOldType <> rsTmp!���� Then
                    lngRow = CPos("T" & rsTmp!����)
                    strOldType = rsTmp!����
                End If

                If mbln����ִ�л��� Then
                    .TextMatrix(lngRow, lngCol) = IIf(rsTmp!ִ�з�ʽ = 0, "", IIf(rsTmp!ִ��״̬ = 0, C_UnExe, C_Exe)) & rsTmp!��Ŀ����
                Else
                    .TextMatrix(lngRow, lngCol) = "" & rsTmp!��Ŀ����   'ҽ��������Ӻ󣬻�δ����·������Ŀǰˢ�£���Ŀ����Ϊ��
                End If
                '����������֯��ʽ ID|��ĿID|��Ŀ���
                '·������Ŀ��ĿidΪ��
                .Cell(flexcpData, lngRow, lngCol) = Val(rsTmp!ID) & "|" & Val("" & rsTmp!��ĿID) & "|" & Val("" & rsTmp!��Ŀ���)

                If IsNull(rsTmp!��ĿID) Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = CON_PathOutItemColor               '·������Ŀ��ǳ��ɫ
                    mcolReason.Add "����˵����" & rsTmp!���ԭ�� & vbCrLf & "����ԭ��" & rsTmp!����ԭ��, "C" & rsTmp!ID
                    If rsTmp!����ԭ�� & "" <> "" Or rsTmp!���ԭ�� & "" <> "" Then
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "����ԭ��" & rsTmp!����ԭ�� & vbCrLf & "����˵����" & rsTmp!���ԭ��
                    End If
                ElseIf Val(NVL(rsTmp!ִ�з�ʽ)) = 1 Then                                        '�������ɵģ�δ����
                    If Not IsNull(rsTmp!����ԭ��) Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED                       'ǳ��ɫ
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "����ԭ��" & rsTmp!����ԭ��
                    End If
                ElseIf rsTmp!ִ�з�ʽ = 3 Then                                                  '��ѡ�����ɫ
                    .Cell(flexcpForeColor, lngRow, lngCol) = &HC00000
                    If Not IsNull(rsTmp!����ԭ��) Then                                          '��ҩ·����Ŀ�ı���ԭ��
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HE0EFED                       'ǳ��ɫ
                        .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & vbCrLf & "����ԭ��" & rsTmp!����ԭ��
                    End If
                End If

                If InStr(rsTmp!��Ŀ���, "|") > 0 And Not IsNull(rsTmp!ִ�н��) Then
                    i = Val(Mid(rsTmp!��Ŀ���, InStr(rsTmp!��Ŀ���, rsTmp!ִ�н��) + Len(rsTmp!ִ�н��) + 1, 1))
                    If i > 0 Then i = Getִ�н������ͼ��(i)
                Else
                    i = 0
                End If

                If Not IsNull(rsTmp!ͼ��ID) Or i > 0 Then
                    .Cell(flexcpPictureAlignment, lngRow, lngCol) = flexPicAlignRightCenter    ' flexPicAlignLeftCenter
                    If i > 0 Then
                        .Cell(flexcpPicture, lngRow, lngCol) = imgCharacter.ListImages(i).Picture
                    Else
                        .Cell(flexcpPicture, lngRow, lngCol) = GetPathIcon(rsTmp!ͼ��ID)
                    End If
                End If

                lngRow = lngRow + 1
                rsTmp.MoveNext
            Loop
        Next

        '4)��ʾ������Ϣ
        If .Rows = .FixedRows And .Cols = .FixedCols Then
            Call ClearPathItem(True)
            .BackColorSel = vbWhite
            .ForeColorSel = vbBlack
        Else
            .BackColorSel = &H8000000D
            .ForeColorSel = &H8000000E
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            lngDayRow = .FixedRows - 2
            .TextMatrix(lngRow, .FixedCols - 1) = "�������"
            .Cell(flexcpBackColor, lngRow, 0) = .BackColorFixed         '&HEFF0E0      '&HD0EFFF
            Call .CellBorderRange(.Rows - 1, 0, .Rows - 1, .Cols - 1, vbBlack, 0, 1, 0, 0, 0, 0)

            strSql = " Select a.�׶�id, a.����, a.�������, a.����˵��, a.������,a.����ʱ��, c.���� As ����ԭ��, a.���������, Nvl(a.ʱ�����, 0) ʱ�����" & vbNewLine & _
                     " From ��������·������ A, ��������·������ B, ������쳣��ԭ�� C" & vbNewLine & _
                     " Where a.·����¼id = b.·����¼id(+) And a.�׶�ID=B.�׶�ID(+) And a.����=b.����(+) And a.·����¼id = [1] And b.����ԭ�� = c.����(+)"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
            For lngCol = .FixedCols To .Cols - 1
                .Cell(flexcpBackColor, lngRow, lngCol) = &HEDF8FF

                rsTmp.Filter = "�׶�ID='" & .ColData(lngCol) & "' And ����=" & Val(Replace(.TextMatrix(EFixedRow.R1����, lngCol), "��", ""))
                str����ԭ�� = ""
                For j = 1 To rsTmp.RecordCount
                    '��ȡ�������ԭ��
                    str����ԭ�� = str����ԭ�� & rsTmp!����ԭ�� & "��"
                    If j = rsTmp.RecordCount Then
                        str����ԭ�� = Mid(str����ԭ��, 1, Len(str����ԭ��) - 1)
                        If InStr(rsTmp!����˵��, vbCrLf) = 0 Or IsNull(rsTmp!����˵��) Then
                            strTmp = "" & rsTmp!����˵��
                        Else
                            arrtmp = Split(rsTmp!����˵��, vbCrLf)
                            strTmp = ""
                            For i = 0 To UBound(arrtmp)
                                strTmp = strTmp & vbCrLf & Space(4) & (i + 1) & "." & arrtmp(i)
                            Next
                        End If
                        strTmp = strTmp & vbCrLf & "�� �� �ˣ�" & rsTmp!������
                        If rsTmp!������� = 1 Then
                            str������� = "����"
                        ElseIf mPP.����·��״̬ = 3 And lngCol = .Cols - 1 Then
                            str������� = "������˳�" & vbCrLf & "����ԭ��" & str����ԭ�� & vbCrLf & "�� �� �ˣ�" & rsTmp!���������

                        ElseIf mPP.����·��״̬ = 2 And lngCol = .Cols - 1 Then
                            str������� = "��������" & vbCrLf & "����ԭ��" & str����ԭ��
                        Else
                            str������� = "��������" & vbCrLf & "����ԭ��" & str����ԭ��
                            If Not IsNull(rsTmp!���������) Then str������� = str������� & vbCrLf & "�� �� �ˣ�" & rsTmp!���������
                        End If

                        .TextMatrix(lngRow, lngCol) = "���������" & str������� & vbCrLf & "����˵����" & strTmp
                        If rsTmp!������� = -1 Then
                            .Cell(flexcpForeColor, lngRow, lngCol) = vbRed     '�����ú�ɫ��ʾ
                        End If

                        If rsTmp!ʱ����� = 1 Or rsTmp!ʱ����� = 2 Then
                            '��ǰ
                            .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "��"
                            .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                        ElseIf rsTmp!ʱ����� = -1 Then    '�Ӻ�
                            .TextMatrix(lngDayRow, lngCol) = .TextMatrix(lngDayRow, lngCol) & "��"
                            .Cell(flexcpForeColor, lngDayRow, lngCol) = &H80FF&
                        End If
                    End If
                    rsTmp.MoveNext
                Next
                If rsTmp.RecordCount = 0 Then
                    .TextMatrix(lngRow, lngCol) = ""
                End If
            Next
        End If
        
        '5)�����������
        
        
    
        .Redraw = True
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч

        If lngPreRow <> -1 And lngPreCol <> -1 And lngPreRow <= .Rows - 1 And lngPreCol <= .Cols - 1 Then
            .Select lngPreRow, lngPreCol
        Else
            .Select .FixedRows, .FixedCols
        End If
    End With

    Exit Sub
errH:
    vsPath.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Get����·����Ϣ(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, ByVal lng����ID As Long, Optional ByVal lng·����¼ID As Long)
'���ܣ���ȡ���˵������ٴ�·����Ϣ
'������lng·����¼ID=��һ�������ж���·��ʱ��ˢ��ָ��·����¼ID��·����
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH
    'һ�ξ���ֻ֧��һ��·�������ܿ���
    '��ǰ�׶�Ϊ0��ʾ��δ���ɹ�·��
    strSql = " Select a.ID,a.·��ID,c.·��ID as ԭ·��ID,a.�汾��,a.״̬,a.��ǰ�׶�ID,a.��ǰ����," & _
             " b.���� as δ����ԭ��,c.��ID,e.���� as ·������,a.������,a.����ʱ��,a.����ʱ��" & _
             " From ��������·�� A,������쳣��ԭ�� B,����·���׶� C,����·���׶� D,����·��Ŀ¼ E" & _
             " Where a.����ID = [1] And a.����ID = [2] And a.·��ID=e.id And a.δ����ԭ�� = b.����(+) And a.��ǰ�׶�ID = c.ID(+) And a.ǰһ�׶�ID=d.id(+)" & _
             IIf(lng·����¼ID <> 0, " And a.ID=[3] ", "") & _
             " Order By a.����ʱ�� Desc"                                                                         'ȡ���һ�ε����·����֧��һ�ξ�����·����
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng����ID, lng·����¼ID)
    If rsTmp.RecordCount > 0 Then
        mPP.ԭ·��ID = Val("" & rsTmp!ԭ·��ID)
        mPP.·��ID = rsTmp!·��ID
        mPP.�汾�� = rsTmp!�汾��
        mPP.����·��ID = rsTmp!ID
        mPP.����·��״̬ = rsTmp!״̬
        mPP.��ǰ�׶�ID = Val("" & rsTmp!��ǰ�׶�ID)
        mPP.�׶θ�ID = Val("" & rsTmp!��ID)
        mPP.��ǰ���� = Val("" & rsTmp!��ǰ����)
        mPP.��ǰ���� = "0"                                      '��LoadPathItem�и�ֵ
        mPP.δ����ԭ�� = "" & rsTmp!δ����ԭ��
        
        If lng·����¼ID = 0 Then mlngPathCount = rsTmp.RecordCount
    Else
        mPP.ԭ·��ID = 0
        mPP.·��ID = 0
        mPP.�汾�� = 0
        mPP.����·��ID = 0
        mPP.����·��״̬ = -1
        mPP.��ǰ�׶�ID = 0
        mPP.�׶θ�ID = 0
        mPP.��ǰ���� = 0
        mPP.��ǰ���� = "0"
        mPP.δ����ԭ�� = ""
        mlngPathCount = 0
    End If

    If mlngPathCount > 1 Then
        fraPath.Visible = True
        lblInDiag.Caption = "�������:" & Get�������(mPP.����·��ID)
        lblInPep.Caption = "������:" & rsTmp!������
        lblInDate.Caption = "����ʱ��:" & Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:mm")
        lblOutDate.Caption = "����ʱ��:" & Format(rsTmp!����ʱ��, "YYYY-MM-DD HH:mm")
        If lng·����¼ID = 0 Then
            cboPath.Clear
            Do While Not rsTmp.EOF
                cboPath.AddItem rsTmp!·������ & ""
                cboPath.ItemData(cboPath.NewIndex) = rsTmp!ID & ""
                rsTmp.MoveNext
            Loop
            zlControl.CboSetIndex cboPath.Hwnd, 0
        End If
    Else
        fraPath.Visible = False
        cboPath.Clear
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get�������(ByVal lng·����¼ID As Long) As String
'���ܣ�ȡ������ϵ�����
    Dim strSql As String, rsTmp As Recordset

    On Error GoTo errH
    strSql = " Select B.������� From ��������·�� A,������ϼ�¼ B Where " & _
             " a.����id = b.����id And a.�Һ�id = b.��ҳid  and a.������� = b.������� And " & _
             " a.�����Դ = b.��¼��Դ And NVL(a.����id,0) = NVL(b.����id,0) And NVL(a.���id,0) = NVL(b.���id,0) And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "Get�������", lng·����¼ID)
    If rsTmp.RecordCount > 0 Then Get������� = rsTmp!������� & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlPrintOutPut(ByVal bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'���ܣ������ٴ�·������ӡ
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel,4-�����PDF
'      blnIsSetup-��ʾ������ӡ�������д�ӡǰ����
'      ��bytStyle=4ʱ����Ҫ����strPDFFile=PDF���Ĭ��·��,�����ļ�������׺
    Call FuncPathTableOutput(bytStyle, blnIsSetup, strPDFFile, strDeviceName)
End Sub

Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, ByVal str�Һ�NO As String, ByVal lng����ID As Long, _
                          ByVal int����״̬ As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal blnForceRefresh As Boolean = True, _
                          Optional ByVal lng·����¼ID As Long) As Long
'������lng·����¼ID=��һ�������ж���·��ʱ��ˢ��ָ��·����¼ID��·����
'      blnForceRefresh=True δ�л�����ˢ��ʱҲ����ˢ�£�����ˢ��
    Dim objControl As CommandBarControl
    Dim strPrePati As String

    strPrePati = mPati.����ID & "_" & mPati.�Һ�ID
    If strPrePati = lng����ID & "_" & lng�Һ�ID And lng����ID <> 0 And Not blnForceRefresh Then Exit Function       '����֮ǰ��Ԫ��λ�ò���

    mPati.����ID = lng����ID
    mPati.�Һ�ID = lng�Һ�ID
    mPati.�Һ�NO = str�Һ�NO
    mPati.����ID = lng����ID
    mPati.����״̬ = int����״̬

    Set mcolReason = New Collection
    Call Get����·����Ϣ(lng����ID, lng�Һ�ID, lng����ID, lng·����¼ID)

    If mPP.����·��ID = 0 Then
        Call SetUnImport
    Else
        If mPP.����·��״̬ = 0 Then
            Call SetImportFalse
        Else
            Call LoadPathFlow
            Call LoadPathItem
        End If
    End If
    Call Form_Resize                                '����·�����̱��Ƿ��й������������߶�
End Function

Private Sub InitCbsSubBar()
    Dim objBar As CommandBar

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOffice2003
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True         '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False     'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    Set cbsSub.Icons = zlCommFun.GetPubIcons
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False

    Set objBar = cbsSub.Add("�ڲ�������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 24, 24
    objBar.Visible = False              'ֻ���ڲ�����ʱ����ʾ(zlDefCommandBars)
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object)
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim lngStart As Long, i As Long

    mbln����ִ�л��� = Val(zlDatabase.GetPara("�Ƿ�����·��ִ�л���", glngSys, P����·��Ӧ��, 1))

    Set mfrmParent = frmParent

    If cbsMain Is Nothing Then Exit Sub

    Set mcbsMain = cbsMain
    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '�ļ��˵�
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Find(, conMenu_File_Excel)
        objControl.Caption = "�����&Excel(ҽʦ��)��"
        Set objControl = .Find(, conMenu_File_Print)
        objControl.Caption = "��ӡ·����(ҽʦ��)(&P)"
        Set objControl = .Add(xtpControlButton, conMenu_File_Print_PatiPath, "��ӡ·����(���߰�)(&Q)", objControl.Index + 1)
        objControl.IconId = conMenu_File_Print
    End With

    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "·��(&E)", objPopup.Index + 1, False)
    objPopup.ID = conMenu_EditPopup
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "����·��(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ������")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����·����Ŀ(&C)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����������Ŀ(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "��������ҽ��")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "���·������Ŀ")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�·������Ŀ")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ȡ����������(&X)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ȡ����ǰ��Ŀ(&V)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive, "��Ŀִ�еǼ�(&E)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnArchive, "ȡ��ִ�еǼ�(&Z)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "����ִ�еǼ�(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "����ȡ��ִ��(&F)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����(&D)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�޸�����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Clear, "ȡ������")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "���·��(&O)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ȡ�����")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogModi, "�޸ĳ����ǼǱ�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "����")
        objControl.IconId = conMenu_Manage_Up
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "����")
        objControl.IconId = conMenu_Manage_Down
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objPopup.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "��׼·���ο�")
        objControl.BeginGroup = True
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Select, "�鿴��������")
        objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "�鿴�����ǼǱ�")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)", objControl.Index + 1)
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        '
        Set objControl = .Add(xtpControlButton, conMenu_Edit_View, "�鿴��Ŀ����(&A)")
    End With

    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objPopup Is Nothing Then
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objPopup = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objPopup.Index, False)
        objPopup.ID = conMenu_ToolPopup
    End If

    '����������
    '-----------------------------------------------------
    lngStart = 0
    Set cbrToolBar = cbsMain(2)
    For Each objControl In cbrToolBar.Controls    '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = cbrToolBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    lngStart = objControl.Index + 1

    If lngStart <> 0 Then
        With cbrToolBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Import, "����", lngStart)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", objControl.Index + 1)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "����", objControl.Index + 1)
            objControl.ToolTipText = "�������ɿ�ѡ���ɵ�·����Ŀ"

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Merge, "ִ��", objControl.Index + 1)
            objControl.BeginGroup = True
            objControl.ToolTipText = "����ִ��·����Ŀ"
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "����", objControl.Index + 1)
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "���", objControl.Index + 1)

            Set objControl = .Add(xtpControlButton, conMenu_Edit_Up, "����", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Up
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Down, "����", objControl.Index + 1)
            objControl.IconId = conMenu_Manage_Down

            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "����", objControl.Index + 1)
            objPopup.BeginGroup = True
            objPopup.IconId = conMenu_Manage_Report
            objPopup.ToolTipText = "���ı���"
        End With
    End If
End Sub

Private Sub FuncPatiPathPrint()
'���ܣ�������߰��ٴ�·��
    Dim WordApp As Object       'Word.Application
    Dim WordDoc As Object       'Word.Document
    Dim strSql As String
    Dim rsTmp As Recordset
    Dim strFileName As String, strFilePath As String
    Dim lngRetu As Long, strInfo As String

    If vsPath.FixedRows < 3 Then
        MsgBox "�ò��˻�δ��������·����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    '���·��
    Screen.MousePointer = 11
    strSql = "Select �ļ��� from ����·���ļ� where ·��ID=[1] And ���=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.·��ID)
    If rsTmp.RecordCount > 0 Then
        strFileName = rsTmp!�ļ��� & ""
        strFilePath = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & strFileName
        If gobjFile.FileExists(strFilePath) Then gobjFile.DeleteFile strFilePath, True
        '�����ݿ���BLOB���ݶ���������ʱ�ļ�Ŀ¼��
        strFilePath = Sys.ReadLob(glngSys, 26, mPP.·��ID & "," & strFileName, strFilePath)
        If Not gobjFile.FileExists(strFilePath) Then
            MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Sub
        End If
    Else
        Screen.MousePointer = 0
        MsgBox "��·����û�����ö�Ӧ�������ٴ�·����(���߰�),�뵽�����ٴ�·�����������á�", vbInformation, gstrSysName
        Exit Sub
    End If

    Set WordApp = CreateObject("Word.Application")
    If WordApp Is Nothing Then
        MsgBox "�밲װMicrosoft Office Word��", vbInformation, gstrSysName
        Exit Sub
    End If

    Set WordDoc = WordApp.Documents.Open(strFilePath)      '��RTF�ĵ�
    WordDoc.PrintPreview
    WordApp.Visible = True
    WordApp.ScreenUpdating = True
    WordApp.Activate
    Screen.MousePointer = 0

    '��¼��ӡ��Ϣ
    Call zlDatabase.ExecuteProcedure("Zl_���Ӳ�����ӡ_Insert(" & mPP.����·��ID & ",12," & mPati.����ID & "," & mPati.�Һ�ID & ",'" & UserInfo.���� & "')", "��ӡ���߰�·����")
    '��ӡ��ǿ�����¼�����ʾ��Ϣ��������ʾ��Ϣ
    Call LoadPathFlow
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathTableOutput(bytStyle As Byte, Optional ByVal blnIsSetup As Boolean, Optional ByVal strPDFFile As String, Optional ByVal strDeviceName As String)
'���ܣ���������ٴ�·����
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel,4-�����PDF
'      blnIsSetup-������ӡ�����д�ӡǰ����
'      strPDFFile=PDF���Ĭ��·��
'      strDeviceName=ָ����ӡ������
    Dim rsTmp As ADODB.Recordset
    Dim vsBody As VSFlexGrid
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngColor As Long, bytR As Byte
    Dim strSql As String
    Dim rsSQLTmp As ADODB.Recordset
    Dim strDisease As String            '�������
    Dim strStandardDate As String       '��׼����ʱ��
    Dim i As Long, j As Long
    Dim strTitle As String
    Dim strTmp As String
    Dim lngDefDay As Long

    strSql = " Select a.����id, a.�Һ�id, b.����id, b.���id, b.�������, c.��׼����ʱ��" & vbNewLine & _
             " From ��������·�� A, ������ϼ�¼ B, ����·���汾 C" & vbNewLine & _
             " Where a.����id = b.����id And a.�Һ�id = b.��ҳid And a.������� = b.�������" & vbNewLine & _
             " And a.�����Դ = b.��¼��Դ And c.·��id = a.·��id And c.�汾�� = a.�汾�� And" & vbNewLine & _
             " b.��ϴ��� = 1 And a.����id = [1] And a.�Һ�id = [2] And a.ID=[3]"
    mblnUnChange = True
    If vsPath.FixedRows < 3 Then
        '���PDF���������·�����ˣ���ֱ���˳�����ʾ
        If bytStyle = 4 Then Exit Sub
        '������ӡ����ʾ
        If blnIsSetup Then Exit Sub
        MsgBox "�ò��˻�δ��������·����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error GoTo errH
    Set rsTmp = GetPatiInfoOut(mPati.����ID, mPati.�Һ�ID)
    Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.����ID, mPati.�Һ�ID, mPP.����·��ID)

    If rsSQLTmp.RecordCount > 0 Then
        strDisease = rsSQLTmp!������� & ""
        strStandardDate = rsSQLTmp!��׼����ʱ�� & ""
    Else
        strDisease = ""
        strStandardDate = ""
    End If
    '��ͷ
    If InStr(vsFlow.TextMatrix(0, 0), vbCrLf) > 0 Then
        strTitle = Mid(vsFlow.TextMatrix(0, 0), 1, InStr(vsFlow.TextMatrix(0, 0), vbCrLf) - 1)
    Else
        strTitle = vsFlow.TextMatrix(0, 0)
    End If
    objOut.Title.Text = strTitle & vbCrLf & "�����ٴ�·����"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 20
    objOut.Title.Font.Bold = True

    '����
    strSql = " Select a.������� From ������ϼ�¼ A" & vbNewLine & _
             " Where a.����id = [1] And a.��ҳid = [2] And a.��¼��Դ = 3 And a.������� In (1, 11) Order By a.��ϴ���"
    Set rsSQLTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPati.����ID, mPati.�Һ�ID)
    If rsSQLTmp.RecordCount > 0 Then
        strTmp = rsSQLTmp!������� & ""
        strTmp = Mid(strTmp, InStr(strTmp, ")") + 1) & Mid(strTmp, 1, InStr(strTmp, ")"))
    Else
        strTmp = ""
    End If
   
    Set objRow = New zlTabAppRow
    objRow.Add "���ö��󣺵�һ���Ϊ " & strTmp
    objOut.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "����������" & rsTmp!���� & " �Ա�" & rsTmp!�Ա� & " ���䣺" & rsTmp!���� & " ����ţ�" & rsTmp!�����
    objOut.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��������:" & Format(rsTmp!����ʱ��, "yyyy��MM��dd��")
    objRow.Add "��ɾ�������:" & Format(rsTmp!���ʱ��, "yyyy��MM��dd��")
    objRow.Add "��׼����ʱ�䣺" & IIf(InStr(strStandardDate, "-") > 0, "", "��") & strStandardDate & "��"
    objOut.UnderAppRows.Add objRow
    objOut.AppFont.Size = 12
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow

    'ҳ��
    objOut.Footer = ";��[ҳ��]ҳ����[ҳ��]ҳ;"
    objOut.PageFooter = 5

    '����
    strTmp = zlDatabase.GetPara("·������ӡ����", glngSys, P����·��Ӧ��, "0")
    If strTmp = "1" Then
        Set vsBody = FuncConvertPathTable
    Else
        Set vsBody = vsPath
    End If

    '���
    With vsBody
        .Redraw = flexRDNone
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "ҽ��ǩ��"
        .RowHeight(.Rows - 1) = 440

        'Ĭ�ϴ�ӡ����
        lngDefDay = Val(zlDatabase.GetPara("·����ÿҳ��ӡ������", glngSys, P����·��Ӧ��, "2"))
        objOut.PageCols = lngDefDay + .FixedCols
        '�����������ʱ���������
        If (.Cols - 1) Mod lngDefDay <> 0 Then
           .Cols = .Cols + (lngDefDay - ((.Cols - 1) Mod lngDefDay))
        End If
        '��ӡ���ת��
        Call FuncPathTableChange(vsBody, lngDefDay)

        '�ƻ��ϲ�������,��ӡ�����жԺϲ����е�������
        For i = .FixedCols To .Cols - 1
            If i Mod 2 = 0 Then
                .TextMatrix(R0�׶���, i) = .TextMatrix(R0�׶���, i) & vbTab
            End If
        Next
        .Redraw = flexRDDirect
        '�п�����Ӧ
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч

        objOut.FixCol = vsBody.FixedCols
        objOut.FixRow = vsBody.FixedRows
        Set objOut.Body = vsBody

        'ָ����ӡ��
        If strDeviceName <> "" Then SaveSetting "ZLSOFT", "����ģ��\zl9PrintMode\Default", "DeviceName", strDeviceName
        If bytStyle = 1 Or bytStyle = 4 Then
            If bytStyle = 4 Then
                bytR = 4
                objOut.Privileged = True '���Ӳ������� ���ڹ�������Zl9PrintMode�ڲ�������ӡȨ�޼��
            Else
                If Not blnIsSetup Then
                    bytR = zlPrintAsk(objOut)
                Else
                    bytR = 1
                End If
            End If
            Me.Refresh

            If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR, strPDFFile
            '��ӡ���˲�����ӡ��¼
            strSql = "zl_���Ӳ�����ӡ_insert(" & mPP.����·��ID & ",11," & mPati.����ID & "," & mPati.�Һ�ID & ",'" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Else
            zlPrintOrView1Grd objOut, bytStyle
        End If
        mblnUnChange = False
        '�ָ�����ʼ״̬
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)    'vsPath�䶯�����¼���

        If vsPathPrint.UBound = 1 Then Unload vsPathPrint(1)
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Edit_Import, conMenu_Edit_Untread, conMenu_Edit_ImportMerge, conMenu_Edit_UnImportMerge
            If InStr(GetInsidePrivs(P����·��Ӧ��), ";����·��;") = 0 Then blnVisible = False
        Case conMenu_Edit_Send, conMenu_Edit_Append, conMenu_Edit_Delete, conMenu_Edit_Blankoff, conMenu_Edit_SendBack
            If InStr(GetInsidePrivs(P����·��Ӧ��), ";����·��;") = 0 Then blnVisible = False
            If Control.ID = conMenu_Edit_SendBack And blnVisible Then
                blnVisible = Not InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ���´�;") = 0
            End If
        Case conMenu_Edit_Surplus, conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down
            If InStr(GetInsidePrivs(P����·��Ӧ��), ";·������Ŀ;") = 0 Then blnVisible = False
        Case conMenu_Edit_Archive, conMenu_Edit_UnArchive, conMenu_Edit_Merge, conMenu_Edit_DeleteParent
            If InStr(GetInsidePrivs(P����·��Ӧ��), ";ִ��·��;") = 0 Or mbln����ִ�л��� = False Then blnVisible = False
            '����·��ִ�л���ʱ�����ó��Ϻ͵�ǰ���ϲ�һ��ʱ,���ز˵���ť
        Case conMenu_Edit_Audit, conMenu_Edit_Reuse, conMenu_Edit_Clear
            If InStr(GetInsidePrivs(P����·��Ӧ��), ";�׶�����;") = 0 Then blnVisible = False
        Case conMenu_Edit_Stop, conMenu_Edit_ClearUp
            If InStr(GetInsidePrivs(P����·��Ӧ��), ";����·��;") = 0 Then blnVisible = False
        Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView
            If Control.ID = conMenu_Edit_OutLogModi Then
                If InStr(GetInsidePrivs(P����·��Ӧ��), ";����·��;") = 0 Then blnVisible = False
            End If
            If blnVisible Then blnVisible = CheckPathOutLog
        Case conMenu_Edit_Compend
            '���浯��(����ӡ),���ı���
            If InStr(GetInsidePrivs(pסԺҽ���´�), ";�������;") = 0 Then blnVisible = False
    End Select
    Control.Visible = blnVisible
    Control.Category = "���ж�"
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveItem As Boolean
    Dim lng��ĿID As Long

    If vsPath.Redraw = flexRDNone Then Exit Sub

    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub

    With vsPath
        blnHaveItem = .Row > .FixedRows - 1 And .FixedRows <> 0 And .Col > .FixedCols - 1   '.FixedRows=0ʱ��ֻ��һ����ʾ��Ϣ
    End With
    Select Case Control.ID
        '0.���
    Case conMenu_File_PrintSet, conMenu_File_Print, conMenu_File_Preview, conMenu_File_Excel, conMenu_File_Print_PatiPath
        Control.Enabled = mPP.����·��ID <> 0
        '1.����
        '-----------------------------------------
    Case conMenu_Edit_Import    '����·��
        Control.Enabled = (mPati.����״̬ = pt���� Or mPati.����״̬ = pt����) And mPati.����ID <> 0 And cboPath.ListIndex <= 0
    Case conMenu_Edit_Untread   'ȡ������(���ڵ�һ������ʱ��ȡ������)
        Control.Enabled = (mPati.����״̬ = pt���� Or mPati.����״̬ = pt����) And mPP.����·��ID <> 0 And (mPP.����·��״̬ = 0 Or mPP.����·��״̬ = 1) And vsPath.Cols <= vsPath.FixedCols + 1
    Case conMenu_Edit_Select      '�鿴��������
        Control.Enabled = mPP.����·��ID <> 0
        '2.����
        '-----------------------------------------
    Case conMenu_Edit_Send      '����·��
        Control.Enabled = (mPati.����״̬ = pt���� Or mPati.����״̬ = pt����) And mPP.����·��ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Append    '��������
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Blankoff  'ȡ����������
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Delete, conMenu_Edit_SendBack   'ȡ��·����Ŀ,��������ҽ��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Col > 0 Then
                    Control.Enabled = (.ColData(.Col) = mPP.��ǰ�׶�ID And .Col = .Cols - 1)
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Surplus   '���·������Ŀ
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Modify, conMenu_Edit_Up, conMenu_Edit_Down     '�޸�·������Ŀ
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" And vsPath.Row <> vsPath.Rows - 1 Then
                lng��ĿID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '·������ĿΪ0
                Control.Enabled = lng��ĿID = 0
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_View      '�鿴��Ŀ����
        Control.Enabled = blnHaveItem
        If Control.Enabled Then
            If vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) <> "" Then
                lng��ĿID = Split(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col), "|")(1)    '·������ĿΪ0
                Control.Enabled = lng��ĿID <> 0
            End If
        End If
        '3.ִ��
        '-----------------------------------------
    Case conMenu_Edit_Archive, conMenu_Edit_UnArchive   '������Ŀִ��(�����һ�ε��в���) 'ȡ��ִ��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then
            With vsPath
                If .TextMatrix(.Row, .Col) <> "" And .Row <> .Rows - 1 And .Row >= .FixedRows Then
                    Control.Enabled = (.ColData(.Col) = mPP.��ǰ�׶�ID And .Col = .Cols - 1)
                Else
                    Control.Enabled = False
                End If
            End With
        End If
    Case conMenu_Edit_Merge, conMenu_Edit_DeleteParent    '����ִ��,����ȡ��ִ��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        '4.����
        '-----------------------------------------
    Case conMenu_Edit_Audit     '�׶�����
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Reuse     '�޸�����
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
    Case conMenu_Edit_Clear     'ȡ������
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        '5.���
        '-----------------------------------------
    Case conMenu_Edit_Stop      '���·��
        Control.Enabled = mPP.��ǰ�׶�ID <> 0 And mPP.����·��״̬ = 1
        If Control.Enabled Then    '��ǰ���������׼����ʱ�䷶Χ�����������������
            Control.Enabled = mblnInOverScope And vsPath.TextMatrix(vsPath.Rows - 1, vsPath.Cols - 1) <> ""
        End If
    Case conMenu_Edit_OutLogModi, conMenu_Edit_OutLogView   '�����ǼǱ�
        Control.Enabled = (mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3)     '2-������ɣ�3-�������
    Case conMenu_Edit_ClearUp   'ȡ�����
        If mPP.����·��״̬ = 3 Then
            Control.Caption = "ȡ���˳�"
        Else
            Control.Caption = "ȡ�����"
        End If
        Control.Enabled = (mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3) And cboPath.ListIndex <= 0    '2-������ɣ�3-�������
    Case conMenu_Edit_Compend    '�鿴����
        With vsPath
            Control.Enabled = blnHaveItem
            If Control.Enabled Then Control.Enabled = .Cell(flexcpData, .Row, .Col) <> ""
        End With
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rsTmp As ADODB.Recordset, str�������� As String
    Dim blnDo As Boolean
    Dim strTmp As String

    Select Case Control.ID
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Print
        Call FuncPathTableOutput(1)
    Case conMenu_File_Preview
        Call FuncPathTableOutput(2)
    Case conMenu_File_Excel
        Call FuncPathTableOutput(3)
    Case conMenu_File_Print_PatiPath
        '��ӡ���߰�·����
        Call FuncPatiPathPrint
        '1.����
        '-----------------------------------------
    Case conMenu_Edit_Import                '����·��
        Call FuncImport(, True)
    Case conMenu_Edit_Untread               'ȡ������
        Call FuncUnImport
    Case conMenu_Edit_Select                '�鿴��������
        Call frmEvaluateOut.ShowMe(mfrmParent, 0, 0, mPati, mPP)
        '2.����
        '-----------------------------------------
    Case conMenu_Edit_Send                  '����·��
        Call FuncSendItem
    Case conMenu_Edit_Append                '��������
        Call FuncSendItemApend
    Case conMenu_Edit_Delete                'ȡ�������ɵ���Ŀ
        Call FuncDelItem
    Case conMenu_Edit_Blankoff              'ȡ����������
        Call FuncDelAllItem
    Case conMenu_Edit_SendBack              '��������ҽ��
        Call FuncReSendItem
    Case conMenu_Edit_Surplus               '���·������Ŀ
        Call FuncAppendItem(0)
    Case conMenu_Edit_Modify                '�޸�·������Ŀ
        Call FuncAppendItemModify
        '3.ִ��
        '-----------------------------------------
    Case conMenu_Edit_Archive               'ִ��·��
        Call FuncExecuteItem
    Case conMenu_Edit_Merge                 '����ִ��
        Call FuncExecuteAll
    Case conMenu_Edit_UnArchive             'ȡ��ִ��
        Call FuncExecuteItemCancel
    Case conMenu_Edit_DeleteParent          '����ȡ��ִ��
        Call FuncExecuteAllCancel
        '4.����
        '-----------------------------------------
    Case conMenu_Edit_Audit                 '����
        Call FuncEvaluate
    Case conMenu_Edit_Reuse                 '�޸�����
        Call FuncReEvaluate
    Case conMenu_Edit_Clear                 'ȡ������
        Call FuncEvaluateCancel
        '5.���
        '-----------------------------------------
    Case conMenu_Edit_Stop                  '���·��
        Call FuncOver
    Case conMenu_Edit_ClearUp               'ȡ�����
        Call FuncOverCancel
    Case conMenu_Edit_OutLogModi    '�޸ĳ����ǼǱ�
        Call OutLogModi
    Case conMenu_Edit_OutLogView   '�鿴�����ǼǱ�
        Call frmPathOutLogOut.ShowMe(mfrmParent, mPati.����ID, mPati.�Һ�ID, 1, Nothing, mPP.·��ID, mPP.����·��ID)
        '6.���ƣ�����
        '-----------------------------------------
    Case conMenu_Edit_Up                    '1-����
        Call MovePathItem(1)
    Case conMenu_Edit_Down                  '-1-����
        Call MovePathItem(-1)
        '7.����
        '-----------------------------------------
    Case conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 10  '�������10������
        If InStr(Control.Parameter, ":") > 0 Then
            Call FuncViewReport(Split(Control.Parameter, ":")(0), Split(Control.Parameter, ":")(1))
        End If
    Case conMenu_Edit_View                  '��ʾ·����Ŀ�������Ϣ
        Call vsPath_DblClick
    Case conMenu_View_StPath                '�鿴��׼·���ο�
        Set rsTmp = GetPatiDiagnose(mPati.����ID, mPati.�Һ�ID, 2)  '��ȡ��Ҫ���
        If rsTmp.RecordCount <> 0 Then
            str�������� = rsTmp!����
        End If
        Call frmStPathList.ShowMe(mfrmParent, str��������, 1)
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
'���ܣ�����˵���ĵ����˵�
    Dim objControl As CommandBarControl
    Dim rsTmp As ADODB.Recordset, i As Long, j As Long
    Dim rsTmpPacs As Recordset

    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_Edit_Compend
             With CommandBar.Controls
                .DeleteAll
                With vsPath
                    Set rsTmp = GetReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                    Set rsTmpPacs = GetPACSReportOfPath(Val(Split(.Cell(flexcpData, .Row, .Col), "|")(0)))
                End With

                If rsTmp.RecordCount = 0 And rsTmpPacs.RecordCount = 0 Then
                     .Add xtpControlButton, conMenu_Edit_Compend * 10 + 1, "�ޱ����δ��д"
                Else
                    For i = 1 To rsTmp.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i, rsTmp!�������� & "(&" & i & ")")
                        objControl.Parameter = rsTmp!ID & ":" & rsTmp!ҽ��id
                        rsTmp.MoveNext
                    Next
                    i = rsTmp.RecordCount
                    For j = 1 To rsTmpPacs.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10 + i + j, rsTmpPacs!�ĵ����� & "(&" & i + j & ")")
                        objControl.Parameter = rsTmpPacs!����ID & ":" & rsTmpPacs!ҽ��id
                        rsTmpPacs.MoveNext
                    Next
                End If
            End With
    End Select
End Sub

Private Function GetReportOfPath(ByVal lng·��ִ��ID As Long) As ADODB.Recordset
'���ܣ���ȡ·����Ӧ�ı�������
    Dim strSql As String

    strSql = " Select d.id, d.��������,c.ҽ��Id" & vbNewLine & _
             " From ��������·��ִ�� A, ��������·��ҽ�� B, ����ҽ������ C, ���Ӳ�����¼ D" & vbNewLine & _
             " Where a.Id = [1] And a.Id = b.·��ִ��id And b.����ҽ��id = c.ҽ��Id And c.����id = d.Id"
    On Error GoTo errH
    Set GetReportOfPath = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ִ��ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPACSReportOfPath(ByVal lng·��ִ��ID As Long) As ADODB.Recordset
'���ܣ���ȡ·����Ӧ�ı�������
    Dim strSql As String
    Dim strIDs As String
    Dim rsTmp As Recordset

    strSql = " Select b.����ҽ��id" & vbNewLine & _
             " From ��������·��ִ�� A, ��������·��ҽ�� B" & vbNewLine & _
             " Where a.Id = [1] And a.Id = b.·��ִ��id "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ִ��ID)
    Do While Not rsTmp.EOF
        strIDs = strIDs & "," & rsTmp!����ҽ��id
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If strIDs <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Set GetPACSReportOfPath = mobjPublicPACS.zlDocGetListWithAdvice(strIDs)
    Else
        Set rsTmp = New Recordset
        rsTmp.Fields.Append "ID", adInteger, 1
        rsTmp.CursorLocation = adUseClient
        rsTmp.LockType = adLockOptimistic
        rsTmp.CursorType = adOpenStatic
        rsTmp.Open
        Set GetPACSReportOfPath = rsTmp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CreateObjectPacs(objPublicPACS As Object) As Boolean
    If objPublicPACS Is Nothing Then
        On Error Resume Next
        Set objPublicPACS = CreateObject("zlPublicPACS.clsPublicPACS")
        Err.Clear: On Error GoTo 0
        If Not objPublicPACS Is Nothing Then
            Call objPublicPACS.InitInterface(gcnOracle, UserInfo.����)
        End If
        If objPublicPACS Is Nothing Then
            MsgBox "PACS��������δ�����ɹ���", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreateObjectPacs = True
End Function

Private Sub FuncOver()
'���ܣ����·��
    Dim strSql As String, blnOK As Boolean, lngValue As Long
    Dim colSQL As New Collection, blnTrans As Boolean, i As Long
    Dim str����� As String
    Dim rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long

    On Error GoTo errH

    str����� = UserInfo.����

    If MsgBox("��ȷ��Ҫ��ɵ�ǰ���˵������ٴ�·����?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If

    lngPPStatus = mPP.����·��״̬
    
    If CheckPathOutLogOut Then
        blnOK = frmPathOutLogOut.ShowMe(mfrmParent, mPati.����ID, mPati.�Һ�ID, 0, colSQL, mPP.·��ID, mPP.����·��ID)
        If blnOK = False Then
            lngValue = Val(zlDatabase.GetPara("������д�����ǼǱ�", glngSys, P����·��Ӧ��, "0"))
            If lngValue = 1 Then
                MsgBox "�������·��ǰ������д�����ǼǱ���ȡ������д��·����ɲ���δִ�С�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If

    strSql = "Zl_��������·������_Update(" & mPP.����·��ID & ")"
    gcnOracle.BeginTrans: blnTrans = True
        Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·�����")
    gcnOracle.CommitTrans: blnTrans = False

    Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)

    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncOverCancel()
'���ܣ�ȡ��·�������
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngPPStatus As Long

    On Error GoTo errH
    '��ɺ�û���������õ�ҽ��������ȡ��
    strSql = " Select 1 From ��������·�� A, ����ҽ����¼ B, ���˹Һż�¼ C, ��������·����¼ D " & vbNewLine & _
             " Where a.ID = d.·����¼ID And d.�Һ�ID = C.ID And  B.�Һŵ� = c.No And" & vbNewLine & _
             "       b.����ʱ�� > Trunc(a.����ʱ��, 'MI') And b.ҽ��״̬ Not In (-1, 4) And Nvl(b.Ӥ��, 0) = 0 And a.Id = [1] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "·����ɺ��Ѳ������µ�ҽ������ɾ�������Ϻ��ٽ���ȡ��������", vbInformation, gstrSysName
        Exit Sub
    End If

    If mPP.����·��״̬ = 3 Then
        If MsgBox("��ǰ·���Ǳ�����Զ���ɵģ�ȡ����������������ͬʱɾ��������ȡ����������ȷ��Ҫ������?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If

    lngPPStatus = mPP.����·��״̬

    strSql = "Zl_��������·������_Delete(" & mPP.����·��ID & "," & mPP.����·��״̬ & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·�����")
    Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)

    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItem()
'���ܣ�ִ��·����Ŀ
    Dim lngִ��ID As Long, lng��ĿID As Long
    Dim rsTmp As ADODB.Recordset, strSql As String

    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng��ĿID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)    '·������ĿΪ0
    End With

    strSql = "Select 1 From ��������·��ִ�� Where ID = [1] And ִ��ʱ�� is Not Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����Ŀ��ִ�С�", vbInformation, gstrSysName
        Exit Sub
    End If

    If frmPathExecute.ShowMe(mfrmParent, 1, mPati, mPP, lngִ��ID, 0, , 1) Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteAll()
'���ܣ�����ִ��
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, 0, , 1) Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncExecuteItemCancel()
'���ܣ�ȡ��·����Ŀ��ִ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim lngִ��ID As Long
    Dim blnTip As Boolean

    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With

    strSql = "Select 1 From ��������·��ִ�� Where ID = [1] And ִ��ʱ�� is Null"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����Ŀ��δִ�С�", vbInformation, gstrSysName
        Exit Sub
    End If

    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, Val(vsPath.ColData(vsPath.Col)), CDate(vsPath.Cell(flexcpData, EFixedRow.R2����, vsPath.Col)))
    If rsTmp.RecordCount > 0 Then
        'ǿ��ȡ�������������Ȩ��
        If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ�����������ȡ��ִ�С�" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call FuncEvaluateCancel(False, False)
        Else
            Exit Sub
        End If
    Else
        blnTip = True
    End If

    If blnTip Then
        If MsgBox("��ȷ��Ҫȡ��[" & vsPath.TextMatrix(vsPath.Row, vsPath.Col) & "]��ִ����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If

    strSql = "Zl_��������·��ִ��_Delete(" & lngִ��ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·����Ŀ")
    Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FuncExecuteAllCancel(Optional blnRefresh As Boolean = True) As Boolean
'���ܣ�����ȡ��·����Ŀ��ִ��
'˵����ҽ��վ������ʱ�����ҽ�������ߵ�ִ�еǼ����
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnDo As Boolean

    On Error GoTo errH

    If blnRefresh = True Then
        strSql = " Select 1 From ��������·��ִ�� A,����·����Ŀ B Where A.·����¼ID = [1] And A.�׶�ID = [2] And A.���� = [3] And A.��ĿID=B.ID(+) " & _
                 " And A.ִ��ʱ�� is Not Null And Rownum<2 "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
        If rsTmp.RecordCount = 0 Then
            MsgBox "��ǰ��������ҽ��ִ�еǼǵ��κ���Ŀ��", vbInformation, gstrSysName
            FuncExecuteAllCancel = True
            Exit Function
        End If
    End If

    '�������ڼ��
    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then
        'ǿ��ȡ�������������Ȩ��
        If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ�����������ȡ��ִ�С�" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call FuncEvaluateCancel(False, True)
        Else
            Exit Function
        End If
    End If

    blnDo = frmPathExecute.ShowMe(mfrmParent, 2, mPati, mPP, 0, 0, False, 1)
    If blnDo And blnRefresh Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    FuncExecuteAllCancel = blnDo
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function FuncSendItem(Optional ByRef blnIsCancel As Boolean, Optional ByVal lngType As Long) As Boolean
'���ܣ�ִ������·��
'������blnIsCancel��û��·��������ʱ���û��Ƿ�ȡ����������true=ȡ��
'     lngType:1-ҽ���༭������ã��������󲻼������ɣ���Ϊҽ���༭���治���ٵ���ҽ���༭��
    Dim rsTmp As ADODB.Recordset
    Dim lng���� As Long, lngʱ����� As Long, lng�������� As Long
    Dim lng�׶�ID As Long
    Dim lngPPStatus As Long
    Dim strTmp As String
    Dim strSql As String
    Dim strDate As String
    Dim strPhase As String
    Dim strMsg As String
    Dim blnDo As Boolean
    Dim blnIsNext As Boolean
    Dim blnEvaluate As Boolean
    Dim blnRefresh As Boolean

    On Error GoTo errH

    If mPP.��ǰ���� = 0 Then '��һ��
        strSql = " Select To_number(Trunc(Sysdate)-Trunc(a.��ʼʱ��)+1) as �������� " & _
                 " From ��������·�� a,����·��Ŀ¼ b Where a.ID = [1] And a.·��id = b.id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
        lng���� = rsTmp!��������
        lngʱ����� = 2
    Else
        '2.��ǰδ�����������������µ�
        strSql = "Select ʱ����� From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))

        If rsTmp.RecordCount = 0 Then
            If InStr(GetInsidePrivs(P����·��Ӧ��), ";�׶�����;") = 0 Then
                MsgBox "�ò�����" & mPP.��ǰ���� & "��û�н������������ܽ��к���������", vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("�ò�����" & mPP.��ǰ���� & "��û�н���������������������" & vbCrLf & "������Ҫ��������������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    '����ǰ���ȼ��ִ�еǼ����
                    If Not CheckPathIsExecuted() Then
                        Exit Function
                    End If

                    If frmEvaluateOut.ShowMe(mfrmParent, 1, 1, mPati, mPP) = False Then
                        Exit Function
                    Else
                        lngPPStatus = mPP.����·��״̬
                        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)

                        '�����󣬿��ܽ������˳�·�������Ը��������е�״̬�����ж��Ƿ�Ҫ��������,�˳�������򲻼�������
                        If mPP.����·��״̬ <> 1 Or lngType = 1 Then
                            Exit Function
                        End If

                        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
                        strSql = "Select ʱ����� From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
                        If rsTmp.RecordCount <> 0 Then
                            lngʱ����� = Val("" & rsTmp!ʱ�����): blnEvaluate = True
                        Else
                            Exit Function
                        End If
                        blnIsNext = True
                    End If
                Else
                    Exit Function
                End If
            End If
        Else
            lngʱ����� = Val("" & rsTmp!ʱ�����): blnEvaluate = True
        End If

        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
        If lngʱ����� = 0 Then
            If mPP.��ǰ���� = strDate Then
                lng�������� = GetMustDayOut(mPP.����·��ID, mPP.��ǰ����)
                'a.������컹�������׶Σ��������������׶Σ����������ǵ���
                If CheckSameDayOfPhaseOut(mPP.��ǰ�׶�ID, lng��������) Then
                    lng���� = mPP.��ǰ����
                Else
                    MsgBox "�ò��˵���û���������õĽ׶ο������ɡ�", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mPP.��ǰ���� < strDate Then          '
                lng���� = DateDiff("D", mPP.��ǰ����, strDate) + IIf(mPP.��ǰ���� = 0, 0, mPP.��ǰ����)
            Else                                        'c.��ǰ���ɺ����׶�
                Exit Function
            End If
        ElseIf lngʱ����� = 1 Then                     '��һ�׶���ǰ������(ʱ�䲻�䣬ͬһ�����ɶ���׶ε�����)
            lng���� = mPP.��ǰ����
        ElseIf lngʱ����� = 2 Then                     '��һ�׶���ǰ������
            If mPP.��ǰ���� = strDate Then
                MsgBox "��һ�׶�����Ϊ����һ�׶���ǰ�����족,����û���������õĽ׶ο������ɡ�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            lng���� = DateDiff("D", mPP.��ǰ����, strDate) + 1
        Else                                            '��һ�׶��Ӻ�(������ǰ�׶�)
            If mPP.��ǰ���� = strDate Then
                MsgBox "�ò����ڽ����·�������ɡ�", vbInformation, gstrSysName
                Exit Function
            End If
            lng���� = DateDiff("D", mPP.��ǰ����, strDate) + 1
        End If
    End If

    If frmPathSendOut.ShowMe(mfrmParent, 0, mPati, mPP, mPP.��ǰ�׶�ID, lng����, 0, 0, lngʱ�����, blnDo) Then
        FuncSendItem = True
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncSendItemApend()
'���ܣ���������·��
'      �����·��������ʱ��������
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long
    Dim strTmp As String
    Dim strDate As String

    On Error GoTo errH
    strSql = "Select Max(ID) as ID From ��������·��ִ�� Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
    If IsNull(rsTmp!ID) Then
        MsgBox "�ò����ڽ����·����û�����ɡ�", vbInformation, gstrSysName
        Exit Sub
    End If

    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P����·��Ӧ��), ";�׶�����;") = 0 Then
            MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ����������������ٲ���������Ŀ��", vbInformation, gstrSysName
            Exit Sub
        Else
            'ȡ������
            If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ����������ܲ���������Ŀ��" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        End If
    End If
    
    If frmPathSendOut.ShowMe(mfrmParent, 1, mPati, mPP, mPP.��ǰ�׶�ID, mPP.��ǰ����) Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncReSendItem()
'���ܣ���������·����Ŀ��ҽ��
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lngִ��ID As Long, lng��ĿID As Long, blnMust As Boolean, lng���� As Long

    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng��ĿID = Val(Split(.Cell(flexcpData, .Row, .Col), "|")(1))
    End With
    If lng��ĿID = 0 Then
        MsgBox "Ҫ��������·������Ŀ����ȡ������Ŀ�����ɺ�������ӡ�", vbInformation, gstrSysName
        Exit Sub
    End If

    '1.�Ѿ�ִ�еĲ�������������;�Ѿ����͵�ҽ��������������
    strSql = "Select a.ִ��ʱ��, c.ҽ��״̬" & vbNewLine & _
            "From ��������·��ҽ�� B, ����ҽ����¼ C, ��������·��ִ�� A" & vbNewLine & _
            "Where a.Id = b.·��ִ��id And b.����ҽ��id = c.Id And a.Id = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        If Not IsNull(rsTmp!ִ��ʱ��) And mbln����ִ�л��� Then
            If rsTmp.RecordCount > 0 Then
                MsgBox "����Ŀ��ִ�У������������ɡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        rsTmp.Filter = "ҽ��״̬=8"
        If rsTmp.RecordCount > 0 Then
            MsgBox "����Ŀ��Ӧ��ҽ���Ѿ�������Ч�������Ϻ���ִ�д˲�����", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        MsgBox "����Ŀ����ҽ������Ŀ�������������ɡ�", vbInformation, gstrSysName
        Exit Sub
    End If

    If frmPathSendOut.ShowMe(mfrmParent, 3, mPati, mPP, mPP.��ǰ�׶�ID, mPP.��ǰ����, lng��ĿID, lngִ��ID) Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncDelPhaseItem()
'���ܣ�ǿ��ɾ�����һ�����е�ִ����Ŀ(���ڲ���ʱ�������)
    Dim strSql As String
    Dim lngִ��ID As Long
    Dim i As Long

    On Error GoTo errH
    With vsPath
        For i = .FixedRows To .Rows - 2     '���һ��������
            If .TextMatrix(i, .Cols - 1) <> "" Then
                lngִ��ID = Split(.Cell(flexcpData, i, .Cols - 1), "|")(0)
                strSql = "Zl_��������·������_Delete(" & lngִ��ID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·����Ŀ")
            End If
        Next
    End With
    Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    Exit Sub
 Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FuncDelAllItem(Optional ByVal blnRefresh As Boolean = True, Optional ByVal blnPrompt As Boolean = True) As Boolean
'���ܣ�����ȡ���������ɵ�����·����Ŀ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim i As Long, strIDs As String, strIDSQL As String, blnTrans As Boolean
    Dim strNewIDs As String
    Dim blnExecuted As Boolean
    Dim dat����ʱ�� As Date
    Dim lng���� As Long
    
    If blnPrompt Then
        If MsgBox("ȡ�����ɽ�ɾ��·����Ŀ��Ӧ��ҽ���Ͳ����ļ���" & vbCrLf & "��ȷʵҪȡ���������ɵ�����·����Ŀ��?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    On Error GoTo errH
    
    strSql = "Select A.ID,A.ִ��ʱ�� From ��������·��ִ�� A,����·����Ŀ B Where A.·����¼ID = [1] And A.�׶�ID = [2] And A.���� = [3] and A.��ĿID=B.ID(+) "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����)
       
    Do While Not rsTmp.EOF
        If blnExecuted = False Then
            If Not IsNull(rsTmp!ִ��ʱ��) Then
                blnExecuted = True
            End If
        End If
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If blnExecuted Then
        '���ж�Ȩ�ޣ�����ʾ��ǿ��ȡ��
        If FuncExecuteAllCancel(False) = False Then
            Exit Function
        End If
    End If
    
    strSql = "Select ����ʱ�� from ��������·�� Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    dat����ʱ�� = Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm:ss")
    '����Ƿ�������
    If mbln����ִ�л��� = False Or Not blnExecuted Then
        strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            'ǿ��ȡ�������������Ȩ��
            MsgBox "�������ɵ���Ŀ��������ȡ������֮ǰ���Զ�ȡ��������", vbInformation, gstrSysName
            Call FuncEvaluateCancel(False, False)
        End If
    End If
    
    strIDSQL = "(Select Column_value From Table(f_Str2List([1])))"
    '2.���ҽ��
    '���ǵ������ɵĳ���������ȡ��·����Ŀ�������Ƿ��ͣ�
    '�ǵ������ɵĳ�������У�Ե�δ���ϣ�������ȡ����δУ�Եģ�ȡ��ʱ�Զ�ɾ����Ӧ��ҽ����

    strSql = "Select /*+ Rule*/ distinct A.·��ִ��id" & vbNewLine & _
             "From ��������·��ҽ�� A, ��������·��ҽ�� B" & vbNewLine & _
             "Where a.·��ִ��id In " & strIDSQL & " And a.����ҽ��id = b.����ҽ��id And b.·��ִ��id <> a.·��ִ��id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs)
    If rsTmp.RecordCount = 0 Then
        strNewIDs = strIDs
        'û�зǵ��յĳ���
    Else
      '����ǰ�����˳������ǲ���ȥ����ֻ��鵱���
        strNewIDs = "," & strIDs & ","
        For i = 1 To rsTmp.RecordCount
            If InStr(strNewIDs, "," & rsTmp!·��ִ��id & ",") > 0 Then
                strNewIDs = Replace(strNewIDs, "," & rsTmp!·��ִ��id & ",", ",")
            End If
            rsTmp.MoveNext
        Next
        If strNewIDs = "," Then
            strNewIDs = ""
        Else
            strNewIDs = Mid(strNewIDs, 2, Len(strNewIDs) - 2)
        End If
    End If
    
    If strNewIDs <> "" Then
        '��ʹ��ֹͣ��ҽ��Ҳ������ɾ������59������Ϊ����ʱ��δ��ȷ����
        strSql = "Select /*+ Rule*/ C.ҽ������ From ��������·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id In " & strIDSQL & _
                 " And b.����ҽ��id = c.Id And c.ҽ��״̬ > 1 And c.ҽ��״̬ <> 4 And rownum<2 And to_date(to_char(c.����ʱ�� +59/24/60/60,'yyyy-mm-dd hh24:mi:ss'),'yyyy-mm-dd hh24:mi:ss') >[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewIDs, dat����ʱ��)
        If rsTmp.RecordCount > 0 Then
            strIDs = ""
            For i = 1 To rsTmp.RecordCount
                If i > 10 Then strIDs = strIDs & "......": Exit For
                strIDs = strIDs & vbNewLine & rsTmp!ҽ������
                rsTmp.MoveNext
            Next
            MsgBox "��ǰ���ɵ���Ŀ�����ѷ��͵�δ���ϵ�ҽ����" & strIDs & vbNewLine & "��������ҽ������ִ��ȡ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '3.��鲡��
    strSql = "Select /*+ Rule*/ 1 From ���Ӳ�����¼ Where ·��ִ��id In " & strIDSQL & _
             " And (���ʱ�� is not null or ��ӡ�� is not null) And rownum<2  And ����ʱ�� >[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strIDs, dat����ʱ��)
    If rsTmp.RecordCount > 0 Then
        MsgBox "��ǰ���ɵ���Ŀ��Ӧ�Ĳ�����ǩ�����Ѵ�ӡ����������ȡ����", vbInformation, gstrSysName
        Exit Function
    End If
        
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(Split(strIDs, ","))
        strSql = "Zl_��������·������_Delete(" & Split(strIDs, ",")(i) & ",0)"
        Call zlDatabase.ExecuteProcedure(strSql, "ȡ������·����Ŀ")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    FuncDelAllItem = True

    If blnRefresh Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncDelItem()
'���ܣ�ȡ�����ɵ�ǰѡ���δִ�е�·����Ŀ
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lngִ��ID As Long, lng��ĿID As Long, blnMust As Boolean, lng���� As Long
    Dim blnCancel As Boolean, strReason As String, blnTrans As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long

    With vsPath
        If .Cell(flexcpBackColor, .Row, .Col) = &HE0EFED Then
            MsgBox "����ĿΪ�������ɵ�û�����ɵ���Ŀ������ȡ�����ɡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
        lng��ĿID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
    End With

    If mbln����ִ�л��� Then
        '�Ѿ�ִ�еĲ�����ȡ��
        strSql = "Select 1 " & vbNewLine & _
                "From ��������·��ִ�� A, ����·����Ŀ B" & vbNewLine & _
                "Where a.��Ŀid = b.Id(+) And a.Id = [1] And a.ִ��ʱ�� Is Not Null"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "����Ŀ��ִ�У�����ȡ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '1.���·����Ŀ
    strSql = "Select b.ִ�з�ʽ,a.���� From ��������·��ִ�� a, ����·����Ŀ b Where a.��ĿID = b.ID And a.ID = [1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then '��ʱ��Ŀ������ȡ��
        lng���� = Val("" & rsTmp!����)
        If rsTmp!ִ�з�ʽ = 1 Then
            blnMust = True
        ElseIf rsTmp!ִ�з�ʽ = 2 Or rsTmp!ִ�з�ʽ = 4 Then  '����һ�λ����һ��
            strSql = "Select ��ʼ����,�������� From ����·���׶� Where ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.��ǰ�׶�ID)
            If Not IsNull(rsTmp!��ʼ����) Then
                If Not IsNull(rsTmp!��������) Then
                    blnMust = (lng���� = Val("" & rsTmp!��������))    '�Ƿ����һ��
                    If blnMust Then '�жϸ���Ŀ֮ǰ��û��ִ�й�(·������Ŀ����)
                        strSql = "Select 1 From ��������·��ִ�� Where ·����¼ID = [1] And �׶�ID = [2] And ��ĿID = [3] And ����<[4] And rownum<2"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, lng��ĿID, lng����)
                        If rsTmp.RecordCount > 0 Then blnMust = False
                    End If
                Else
                    blnMust = True  '����
                End If
            End If
        End If
    End If

    '2.���ҽ��
    If lngִ��ID <> 0 Then
        '��ʹ��ֹͣ��ҽ��Ҳ������ɾ������59������Ϊ����ʱ��δ��ȷ����
        strSql = "Select /*+ Rule*/ C.ҽ������ From ��������·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id=[1] " & _
                 " And b.����ҽ��id = c.Id And c.ҽ��״̬ > 1 And c.ҽ��״̬ <> 4 And rownum<2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
        If rsTmp.RecordCount > 0 Then
            MsgBox "��ǰ���ɵ���Ŀ�����ѷ��͵�δ���ϵ�ҽ����" & rsTmp!ҽ������ & vbNewLine & "��������ҽ������ִ��ȡ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    '3.�������ɵ���Ŀ��д����ԭ��
    If blnMust Then
        'ȡ���������ɵ���Ŀʱѡ�����ԭ��
        strSql = "Select b.���� as ����,a.���� as ID,a.����,a.����,a.���� From ������쳣��ԭ�� a,������쳣��ԭ�� b" & _
                " Where a.����=1 And a.ĩ��=1 And a.�ϼ�=b.���� And b.ĩ��=0 " & _
                " Order by ����,a.����"
        vPoint = zlControl.GetCoordPos(vsPath.Hwnd, vsPath.CellLeft, vsPath.CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "������쳣��ԭ��", True, , , True, True, True, _
                 vPoint.X, vPoint.Y, vsPath.RowHeight(vsPath.Row), blnCancel, False, True)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "ϵͳû�г�ʼ������쳣��ԭ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            strReason = rsTmp!ID
        End If
    End If
    '4.��鲡��
    strSql = "Select 1 From ���Ӳ�����¼ Where ·��ִ��id = [1] And (���ʱ�� is not null or ��ӡ�� is not null) And rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngִ��ID)
    If rsTmp.RecordCount > 0 Then
        MsgBox "����Ŀ��Ӧ�Ĳ�����ǩ�����Ѵ�ӡ������ȡ����", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsPath
        If MsgBox("ȷʵҪȡ��·����Ŀ""" & .TextMatrix(.Row, .Col) & """��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    End With
    If Not mbln����ִ�л��� Then
        '�ж��Ƿ��Ѿ�����
        strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            'ǿ��ȡ�������������Ȩ��
            MsgBox "�������ɵ���Ŀ��������ȡ������֮ǰ���Զ�ȡ��������", vbInformation, gstrSysName
            Call FuncEvaluateCancel(False, False)
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    If strReason <> "" Then
        strSql = "Zl_��������·������_Update(" & lngִ��ID & ",'" & vsPath.TextMatrix(vsPath.Row, 0) & "',Null,NULL,NULL,NULL,'" & strReason & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "�޸�·����Ŀ")
    End If
    strSql = "Zl_��������·������_Delete(" & lngִ��ID & "," & IIf(strReason <> "", "2", "0") & ")"

    Call zlDatabase.ExecuteProcedure(strSql, "ȡ��·����Ŀ")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAppendItemModify()
'���ܣ��޸�·������Ŀ
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim lngִ��ID As Long

    With vsPath
        lngִ��ID = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
    End With

    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P����·��Ӧ��), ";�׶�����;") = 0 Then
            MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ������������������޸�·������Ŀ��", vbInformation, gstrSysName
            Exit Sub
        Else
            'ȡ������
            If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ������������޸�·������Ŀ��" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Sub
            End If
        End If
    End If

    If frmPathAppendOut.ShowMe(mfrmParent, mPati, mPP, "", 2, "", lngִ��ID) Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncAppendItem(ByVal bytUseType As Byte, Optional ByVal strItemType As String, Optional ByVal strAdviceIDs As String, _
                                Optional ByVal lngִ��ID As Long, Optional ByVal datDate As Date) As Boolean
'���ܣ����·������Ŀ(ͨ��clsDockPath�еĽӿڿ��Ÿ�ҽ����������)
'������bytUseType=0-ֱ�����,1-ҽ���¿�ʱ���
'       strItemType=ҽ���ӿڵ���ʱ���루�������һ����Ŀ�ķ��ࣩ
'       strAdviceIDs=ҽ���ӿڵ���ʱ����,ҽ�����
'       datDate =ҽ���Ŀ�ʼִ�����ڣ�ͬһ��·����ҽ����
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim DatCur As Date
    Dim blnRefresh As Boolean

    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then
        If InStr(GetInsidePrivs(P����·��Ӧ��), ";�׶�����;") = 0 Then
            MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ��������������������·������Ŀ��", vbInformation, gstrSysName
            Exit Function
        Else
            'ȡ������
            If MsgBox("�ò�����" & mPP.��ǰ���� & "�ѽ���������������ȡ��������������·������Ŀ��" & vbCrLf & vbCrLf & "������Ҫȡ��������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call FuncEvaluateCancel(False, False)
            Else
                Exit Function
            End If
        End If
    Else
        'δ���������Ƿ��ǵ���δ���������ǵĻ�����ʾ�Ƿ�Ҫ��ӵ����һ���׶�
        If bytUseType = 0 Then
            DatCur = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
            If DatCur <> Format(mPP.��ǰ����, "yyyy-MM-dd") Then
                If MsgBox("��Ҫ���·������Ŀ��""" & mPP.��ǰ���� & """?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Call FuncSendItem
                    Exit Function
                End If
            End If
        End If
    End If

    If bytUseType = 0 Then
        With vsPath
            If .Row > 0 And .Row < .Rows - 2 Then strItemType = .TextMatrix(.Row, .FixedCols - 1) '���һ����"·������"
        End With
    End If
    If frmPathAppendOut.ShowMe(mfrmParent, mPati, mPP, strItemType, bytUseType, strAdviceIDs, lngִ��ID, datDate) Or blnRefresh Then
        FuncAppendItem = True
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncImport(Optional ByVal blnask As Boolean, Optional blnImport As Boolean) As Boolean
'���ܣ�����·��

    If frmPathImportOut.ShowMe(mfrmParent, mPati, blnImport) Then
        FuncImport = True
        
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
        RaiseEvent RequestRefresh(mPP.����·��״̬)
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncUnImport(Optional ByVal blnPrompt As Boolean = True)
'���ܣ�ȡ������,δ����·��ʱ��ȡ������
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, blnTrans As Boolean
    Dim str����� As String
    Dim lngPPStatus As Long

    '�ȼ���Ƿ���ȡ��·����Ȩ��
    If InStr(GetInsidePrivs(P����·��Ӧ��), ";ȡ������;") = 0 Then
        str����� = zlDatabase.UserIdentify(Me, "û��ȡ������Ȩ����Ҫ��ˡ�", glngSys, P����·��Ӧ��, "ȡ������")
        If str����� = "" Then Exit Sub
    Else
        str����� = UserInfo.����
    End If
    strSql = "Select 1 From ��������·��ִ�� Where ·����¼ID = [1] And rownum<2"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If rsTmp.RecordCount > 0 Then

        If MsgBox("��ǰ�׶ε�·����Ŀ�����ɣ����Ƚ�����ȡ�����ɲ�����" & vbCrLf & "��ȷʵҪȡ���ò����ѵ�����ٴ�·����?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If

        'Ҫˢ�£��Ա���ȡ·����Ϣ(��ǰ�׶ε�)
        If FuncDelAllItem(True, False) Then
            Call FuncUnImport(False)    '���µ��ã��ٴμ��
        End If
        Exit Sub
    ElseIf blnPrompt Then
        If MsgBox("��ȷʵҪȡ���ò����ѵ�����ٴ�·����?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If

    lngPPStatus = mPP.����·��״̬
    
    gcnOracle.BeginTrans: blnTrans = True
    strSql = "Zl_��������·������_Delete(" & mPP.����·��ID & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ������")
    '����ȡ�������¼
    strSql = "Zl_��������·��ȡ��_Insert(" & mPati.����ID & "," & mPati.�Һ�ID & ",'" & UserInfo.���� & "','" & str����� & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ������")
    gcnOracle.CommitTrans: blnTrans = False
    Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function FuncEvaluateCancel(Optional ByVal blnPrompt As Boolean = True, Optional ByVal blnRefresh As Boolean = True) As Boolean
'���ܣ�ȡ������,δ����ʱ����ȡ����������Զ�������ֻ��ȡ��������
'������blnPrompt=�Ƿ񵯳�ѯ����ʾ
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim lngPPStatus As Long

    On Error GoTo errH

    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò�����" & mPP.��ǰ���� & "��û�н���������", vbInformation, gstrSysName
        Exit Function
    End If

    If blnPrompt Then
        If MsgBox("��ȷ��Ҫȡ����" & mPP.��ǰ���� & "���������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
    End If
    lngPPStatus = mPP.����·��״̬

    strSql = "Zl_��������·������_Delete(" & mPP.����·��ID & ", " & mPP.��ǰ�׶�ID & ",To_Date('" & mPP.��ǰ���� & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSql, "ȡ������")
    FuncEvaluateCancel = True
    If blnRefresh Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncEvaluate()
'���ܣ��׶�����,ֻ�ܶԵ�ǰ�׶ε����һ����ִ���˵Ľ���
    Dim strSql As String, rsTmp As ADODB.Recordset, i As Long
    Dim strTmp As String
    Dim bln��¼ As Boolean
    Dim blnRefresh As Boolean
    Dim strDate  As String
    Dim lngPPStatus As Long

    '1.�������Ĳ��������� '�ѽ����Ĳ���������(�����˲˵���)
    '2.ֻ�ܶ����һ��ִ�еļ�¼��������(������������ָ��ģ�����������������ɴ���·����û�ж������������ִ������)����Ϊ����Ϊ��������ܽ���·��
    '3.����ý׶ε�������Ŀ��ִ�к��������
    On Error GoTo errH

    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount > 0 Then
        MsgBox "�ò�����" & mPP.��ǰ���� & "�ѽ�����������", vbInformation, gstrSysName
        Exit Sub
    End If

    'ִ�еǼǼ��
    If Not CheckPathIsExecuted(blnRefresh) Then
        'ǿ��ˢ��
        If blnRefresh Then
            Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
        End If
        Exit Sub
    End If

    lngPPStatus = mPP.����·��״̬

    If frmEvaluateOut.ShowMe(mfrmParent, 1, 1, mPati, mPP, , , , , , bln��¼) Or bln��¼ Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If

    If mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3 Then RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncReEvaluate()
'���ܣ��޸���������������˺����׶ε���Ŀ�������޸��������Ϊ�����������������ڱ���Ĵ洢�������жϡ�
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim bln��¼ As Boolean
    Dim lng�׶�ID As Long
    Dim strSysDate As String
    Dim lng���� As Long

    On Error GoTo errH


    strSql = "Select 1 From ��������·������ Where ·����¼ID = [1] And �׶�ID = [2] And ���� = [3]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ò����ڵ�ǰ�׶λ�û�н���������", vbInformation, gstrSysName
        Exit Sub
    End If

    If frmEvaluateOut.ShowMe(mfrmParent, 1, 2, mPati, mPP, , , , , , bln��¼) Or bln��¼ Then
        Call zlRefresh(mPati.����ID, mPati.�Һ�ID, mPati.�Һ�NO, mPati.����ID, mPati.����״̬)
    End If

    If mPP.����·��״̬ = 2 Or mPP.����·��״̬ = 3 Then RaiseEvent RequestRefresh(mPP.����·��״̬)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncViewReport(ByVal str����ID As String, ByVal lngҽ��ID As Long)
'���ܣ����ı���

    '���ж��Ƿ���Լ�������
    If IsNumeric(str����ID) Then
        If CheckEPRReport(Val(str����ID), lngҽ��ID) = 2 Then
            If InStr(GetInsidePrivs(pסԺҽ���´�), "����δ��ɱ���") > 0 Then
                MsgBox "ע�⣺��ҽ���ı��滹û����ʽǩ����", vbInformation, gstrSysName
            Else
                MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û��Ȩ�޲�����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        RaiseEvent ViewEPRReport(Val(str����ID), False)
    Else
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(0, str����ID, Val(zlDatabase.GetPara("�Զ���Ǳ������״̬", glngSys, pסԺҽ���´�, "1")) = 1, mfrmParent)
    End If
End Sub

Public Function CheckEPRReport(ByVal lng����ID As Long, ByVal lngҽ��ID As Long) As Integer
'���ܣ�����Ӧ��Ŀ�ı�����д���
'������lng·��ִ��ID=��������·��ִ�м�¼�е�ID
'      lng����ID=���ر��没��ID
'���أ�
'      1-��������д���(��ǩ��,�����޶���ǩ��,����ִ�����)
'      2-����δ��д���(δǩ��,���޶���δǩ��,��δִ�����)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String

    On Error GoTo errH

    '��鱨��ִ�й���(5-���;6-�������)��״̬(1-���)
    '���鱨���ǹ������ɼ���ʽ����ģ����ɼ���ʽ����Ϊ����δ�������ͼ�¼
    strSql = " Select 2 as ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ�� From ����ҽ������ Where ҽ��ID=[1]" & _
            " Union ALL" & _
            " Select ����,ҽ��ID,ִ�й���,ִ��״̬,����ʱ��" & _
            " From (" & _
                " Select 1 as ����,B.ҽ��ID,B.ִ�й���,B.ִ��״̬,B.����ʱ�� From ����ҽ����¼ A,����ҽ������ B" & _
                " Where A.ID=B.ҽ��ID And A.���ID=(" & _
                    " Select A.ID From ����ҽ����¼ A,������ĿĿ¼ B Where A.ID=[1] And A.������ĿID=B.ID And A.�������='E' And B.��������='6')" & _
                " Order by A.���" & _
            " ) Where Rownum=1" & _
            " Order by ����,����ʱ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lngҽ��ID)
    If NVL(rsTmp!ִ�й���, 0) >= 5 Or NVL(rsTmp!ִ��״̬, 0) = 1 Then
        CheckEPRReport = 1
    Else
        CheckEPRReport = 2
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mcolReason = Nothing
    SaveWinState Me, App.ProductName
    Set mobjPublicPACS = Nothing
End Sub

Private Sub imgMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String, lngId As Long, i As Long
    Dim strSql As String, rsTmp As ADODB.Recordset

    lngId = fraMore.Tag
    If lngId = 0 Then
        Call zlCommFun.ShowTipInfo(0, strInfo)
    Else
        strSql = "Select " & IIf(mbln����ִ�л���, "A.ִ�н��,A.ִ��˵��,A.ִ����,to_char(A.ִ��ʱ��,'yyyy-mm-dd hh24:mi') as ִ��ʱ��,", "") & _
                " A.�Ǽ���,to_char(A.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi') as �Ǽ�ʱ�� From ��������·��ִ�� A,����·����Ŀ B Where A.��ĿID=B.ID(+) And A.ID = [1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        If rsTmp.RecordCount > 0 Then
            With rsTmp
                For i = 0 To .Fields.count - 1
                    strInfo = strInfo & .Fields(i).Name & "��" & .Fields(i).Value & vbCrLf
                Next
            End With
            Call zlCommFun.ShowTipInfo(fraMore.Hwnd, strInfo, True)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsFlow_DblClick()
    Dim lngPhaseID As Long
    If mPP.����·��״̬ = 0 And mPP.����·��ID <> 0 Then   '����ʧ��
        Call frmEvaluateOut.ShowMe(mfrmParent, 0, 0, mPati, mPP)
    Else
        lngPhaseID = Val(vsFlow.ColData(vsFlow.Col))
        If lngPhaseID <> 0 Then
            Call frmPathSendOut.ShowMe(mfrmParent, 2, mPati, mPP, lngPhaseID, 0)
        ElseIf vsFlow.Col = 0 And mPP.·��ID <> 0 Then
            Call frmPathDefinition.ShowMe(mfrmParent, mPP.·��ID, 1)
        Else
            If vsFlow.Col = vsFlow.Cols - 2 And gstrDBUser = "ZLHIS" Then
                vsFlow.Editable = flexEDKbdMouse
            End If
        End If
    End If
End Sub

Private Sub vsFlow_LostFocus()
    If Not (mPP.����·��״̬ = 0 And mPP.����·��ID <> 0) Then
        vsFlow.ForeColorSel = vsFlow.CellForeColor
    End If
End Sub

Private Sub vsFlow_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ����ڲ���ʱǿ��ɾ�����һ�����Ŀ(ѡ�����һ����ͷ������DELA)
    Dim strPass As String, i As Long

    If vsFlow.Col = vsFlow.Cols - 2 Then
        strPass = UCase(vsFlow.EditText)
        vsFlow.EditText = ""
        If strPass = "DELA" Then
            If MsgBox("��ȷ��Ҫɾ�����һ���������Ŀ��", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
            Call FuncDelPhaseItem
        End If
        vsFlow.Editable = flexEDNone
    End If
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Or OldCol <> NewCol Then
        If fraMore.Visible Then fraMore.Visible = False

        If NewRow <> -1 And NewCol <> -1 And mblnUnChange = False Then
            '��ʾ·����Ŀ���ɵ�ҽ���嵥
            Dim strTmp As String

            strTmp = vsPath.Cell(flexcpData, NewRow, NewCol)
            If InStr(strTmp, "|") > 0 Then
                Call UCAdvice.ShowAdvice(1, "", Val("" & Split(strTmp, "|")(0)), , , , , 1)
            Else
                Call UCAdvice.ShowAdvice(1, "", 0, , , , , 1)
            End If
        Else
            Call UCAdvice.ShowAdvice(1, "", 0, , , , , 1)
        End If
    End If
End Sub

Private Sub vsPath_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45
End Sub

Private Sub vsPath_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If fraMore.Visible Then fraMore.Visible = False
End Sub

Private Sub vsPath_DblClick()
    Dim lng��ĿID As Long

    With vsPath
        If Trim(.TextMatrix(.Row, .Col)) <> "" And .Cell(flexcpData, .Row, .Col) <> "" And .Row <> .Rows - 1 Then
            lng��ĿID = Split(.Cell(flexcpData, .Row, .Col), "|")(1)
            If lng��ĿID <> 0 Then
                Call frmPathItemEditOut.ShowView(mfrmParent, lng��ĿID)
            End If
        End If
    End With
End Sub

Private Sub vsPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngId As Long, lngRow As Long, lngCol As Long, lngItemID As Long
    
    With vsPath
        If .MouseCol >= .FixedCols And .MouseRow >= .FixedRows Then
            lngRow = .MouseRow: lngCol = .MouseCol
            If .Cell(flexcpData, lngRow, lngCol) <> "" And lngRow <> .Rows - 1 Then
                lngId = Split(.Cell(flexcpData, lngRow, lngCol), "|")(0)
                lngItemID = Split(.Cell(flexcpData, lngRow, lngCol), "|")(1)
                If lngItemID = 0 Then
                    .ToolTipText = ""
                    Call zlCommFun.ShowTipInfo(.Hwnd, mcolReason("C" & lngId), True)      '·������Ŀ�����ԭ��
                Else
                    If .ToolTipText = "" Then .ToolTipText = "˫���鿴·����Ŀ����"
                    Call zlCommFun.ShowTipInfo(.Hwnd, "")
                End If
            Else
                .ToolTipText = ""
            End If

            If lngId = 0 Then
                If imgMore.Visible Then fraMore.Visible = False
                fraMore.Tag = ""
            Else
                If lngRow = .Row And lngCol = .Col Then
                    fraMore.BackColor = .BackColorSel
                Else
                    fraMore.BackColor = .BackColor
                End If

                fraMore.Tag = lngId
                If fraMore.Visible = False Then fraMore.Visible = True
                fraMore.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - imgMore.Height - 30
                fraMore.Left = .Left + .ColPos(lngCol) + .ColWidth(lngCol) - imgMore.Width - 30
            End If
        Else
            If fraMore.Visible Then fraMore.Visible = False
        End If
    End With
End Sub

Private Sub vsPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Dim lng��ĿID As Long

    '��ʾ�༭�˵����������
    If Button = 2 Then
        If mcbsMain Is Nothing Then Exit Sub
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����·��������ͼ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    mlngFontSize = IIf(bytSize = 0, CON_SmallFontSize, CON_BigFontSize)

    vsFlow.Font.Size = mlngFontSize
    vsFlow.Redraw = flexRDDirect

    Call Grid.SetFontSize(vsPath, mlngFontSize)
    If vsPath.FixedRows > 1 Then vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45 '��ҪDraw֮�����Ч

    Call UCAdvice.SetVsAdviceFontSize(mlngFontSize)
End Sub

Private Sub MovePathItem(ByVal lngWay As Long)
'����:��ǰ��Ԫ��ѡ��·������Ŀʱ������·������Ŀ�������ƶ�
'����:lngWay=1����һ��,-1����һ��(�൱����һ������һ��)
    Dim lngId       As Long
    Dim lngItemNum  As Long
    Dim arrSQL()    As Variant
    Dim i           As Integer
    Dim blnTran     As Boolean
    Dim blnDo As Boolean, blnFind As Boolean
    Dim lngRow As Long, lngCol As Long

    blnDo = True: blnFind = False

    With vsPath
        Do While blnDo
            If .TextMatrix(.Row, .FixedCols - 1) <> .TextMatrix(.Row - lngWay, .FixedCols - 1) Or .Cell(flexcpData, .Row - lngWay, .Col) = "" Then
                MsgBox "��Ŀ����:" & .TextMatrix(.Row, .Col) & vbCrLf & _
                       "�Ѵ��ڡ�" & .TextMatrix(.Row, .FixedCols - 1) & "�������" & IIf(lngWay > 0, "��һ��", "���һ��"), vbInformation, gstrSysName
                blnDo = False: blnFind = False: Exit Do
            Else
                lngRow = .Row - lngWay: lngCol = .Col
                blnFind = True: Exit Do
            End If
        Loop
        '������Ŀ���
        If blnFind Then
            arrSQL = Array()

            lngId = Split(.Cell(flexcpData, .Row, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_��������·�����_update(" & lngId & "," & lngItemNum & ")"

            lngId = Split(.Cell(flexcpData, .Row - lngWay, .Col), "|")(0)
            lngItemNum = Split(.Cell(flexcpData, .Row, .Col), "|")(2)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_��������·�����_update(" & lngId & "," & lngItemNum & ")"

            On Error GoTo errH
            gcnOracle.BeginTrans: blnTran = True
            For i = LBound(arrSQL) To UBound(arrSQL)
                zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
            Next i
            gcnOracle.CommitTrans: blnTran = False

            Call ClearPathItem(True)
            Call LoadPathItem

            '�����ƶ�
            .Row = lngRow: .Col = lngCol
            .ShowCell lngRow, lngCol
        End If
    End With
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FuncConvertPathTable() As VSFlexGrid
'����:ת���ٴ�·����,����Ӧ������Ĵ�ӡ���� 78233
'����:ת�����·����
    Dim lngFirstSameCol As Long
    Dim lngLastRow As Long
    Dim i As Long, j As Long, k As Long

    Grid.CopyTo vsPath, vsPathPrint(0)
    vsPath.Redraw = flexRDNone
    vsPathPrint(0).MergeCol(0) = True
    vsPathPrint(0).MergeRow(0) = True

    With vsPathPrint(0)
        'һ���׶δ�ӡһ�У�����ӡ����������
        For i = 2 To .Cols - 1
            If .TextMatrix(R0�׶���, i) = .TextMatrix(R0�׶���, i - 1) Then
            '��������Ϊͬһ�׶�Ҫ�ϲ�
                For j = 1 To i - 1
                    If .TextMatrix(R0�׶���, i) = .TextMatrix(R0�׶���, j) Then
                        lngFirstSameCol = j '�ҵ��뵱ǰ����ͬһ�׶ε�����
                        Exit For
                    End If
                Next
                'j-��ǰ��,i-��ǰ��
                For j = R2���� + 1 To .Rows - 1
                    If .TextMatrix(j, i) <> "" And .TextMatrix(j, 0) <> "�������" Then
                        k = 0
                        For k = R2���� + 1 To .Rows - 1
                            If .TextMatrix(j, i) = .TextMatrix(k, lngFirstSameCol) Then
                                Exit For
                            End If
                        Next
                        If k = .Rows Then
                            '����
                            k = 0
                            lngLastRow = 0
                            For k = R2���� + 1 To .Rows - 1
                                If .TextMatrix(j, 0) = .TextMatrix(k, 0) Then
                                    If .TextMatrix(k, lngFirstSameCol) = "" Then
                                        'ͬ�����¿�������
                                        .TextMatrix(k, lngFirstSameCol) = .TextMatrix(j, i)
                                        Exit For
                                    End If
                                    lngLastRow = k
                                End If
                            Next
                            If k = .Rows Then
                                'ͬ�������һ������һ��
                                .AddItem "", lngLastRow + 1
                                .TextMatrix(lngLastRow + 1, lngFirstSameCol) = .TextMatrix(j, i)
                                .TextMatrix(lngLastRow + 1, 0) = .TextMatrix(j, 0)
                                .RowHeight(lngLastRow + 1) = .RowHeight(j)
                            End If
                        Else
                            'ǰһ�д��ڣ����Բ�����
                        End If
                    End If
                Next
                '���ɾ����
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        'ɾ����
        For i = .Cols - 1 To 0 Step -1
            If .ColHidden(i) = True And .ColWidth(i) = 0 Then
                '���һ��ֱ��ɾ��
                If i = .Cols - 1 Then
                    .Cols = .Cols - 1
                Else
                    '������ǰ��
                    For k = i + 1 To .Cols - 1
                        For j = 0 To .Rows - 1
                            .TextMatrix(j, k - 1) = .TextMatrix(j, k)
                        Next
                    Next
                    .Cols = .Cols - 1
                End If
            End If
        Next
        '�������ں�����
        .RowHidden(R1����) = True: .RowHidden(R2����) = True
        .RowHeight(R1����) = 0: .RowHeight(R2����) = 0
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч
    End With
    Set FuncConvertPathTable = vsPathPrint(0)
End Function

Private Function CheckPathIsExecuted(Optional ByRef blnRefresh As Boolean) As Boolean
'-------------------------------------------------------------------------------------------
'����:��鵱ǰ�׶��Ƿ����δִ�е�·����Ŀ
'������=1 ���ɻ��ڵ���,=2 ����ʱ�����
'���أ�F-����δ���ִ�еǼǵ�·����Ŀ,���������ɻ�����
'     T-������δ���ִ�еǼǵ�·����Ŀ\�����ִ�еǼ�������������ɻ�����
'˵����1.��ʿ����ʱ,��ǰδִ��ʱ�����������µĽ׶�,ҽ������ʱ,��ǰδִ�в����������µĽ׶�
'      2.����ҽ��Ҫ��ǰ���ɺ����׶� mbln���ò�����=trueʱ,ҽ��վ����ʱ������Ƿ����ִ�еǼǣ������������ڼ��
'      3.��ʿվû����������,��Ҫÿ�����ɶ�Ҫ���ǰһ�ε�ִ�еǼ����
'-------------------------------------------------------------------------------------------
    Dim blnHave As Boolean          '������ִ�л��ڵļ��
    Dim blnReturn As Boolean
    Dim blnExePath As Boolean
    Dim blnUnExe As Boolean         '���ڱ��û��ִ��·��Ȩ���Ҵ��ڲ���Աִ�е�·����Ŀʱ,��Ҫ�����û���ʾ
    Dim strSql As String
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    blnHave = True                  'Ĭ�ϼ��ִ�еǼ����
    blnExePath = InStr(GetInsidePrivs(P����·��Ӧ��), ";ִ��·��;") > 0
    blnReturn = True
    blnHave = mbln����ִ�л���

    If blnHave Then
        strSql = "Select Nvl(b.��Ŀ����,a.��Ŀ����) ��Ŀ���� From ��������·��ִ�� a,����·����Ŀ b " & vbNewLine & _
                        "Where a.��Ŀid=b.id(+) And a.·����¼ID = [1] And a.�׶�ID = [2] And a.���� = [3] And a.ִ��ʱ�� Is null "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID, mPP.��ǰ�׶�ID, CDate(mPP.��ǰ����))
        If rsTmp.RecordCount > 0 Then
            'ҽ�������������ڼ��,����ʱ�����
            If rsTmp.RecordCount > 0 Then
                Call FuncGetRSTipInfo(rsTmp, "��Ŀ����", strTmp)
                If blnExePath Then
                    If MsgBox("�ò��˻���δִ�е���Ŀ:" & vbCrLf & strTmp & vbCrLf & "������ִ�С�������Ҫ����ִ�в�����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        If frmPathExecute.ShowMe(mfrmParent, 0, mPati, mPP, 0, 0, , 1) Then
                            blnRefresh = True
                        Else
                            blnReturn = False
                        End If
                    Else
                        blnReturn = False
                    End If
                Else
                    blnUnExe = True: blnReturn = False
                End If
            End If

            If blnUnExe Then
                'û��ִ��·��Ȩ���Ҵ��ڲ���Աִ�е�·����Ŀʱ , ��Ҫ�����û���ʾ
                MsgBox "�ò��˻���δִ�е���Ŀ��" & vbCrLf & strTmp & vbCrLf & "����ִ�к���ܼ�����", vbInformation, gstrSysName
            End If
        End If
    End If
    CheckPathIsExecuted = blnReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncGetRSTipInfo(ByVal rsTmp As ADODB.Recordset, ByVal strFieldName As String, ByRef strTipInfo As String)
'����:ѭ����ȡ��¼���м�Ҫ��Ϣ
    Dim i As Long

    strTipInfo = ""
    For i = 1 To rsTmp.RecordCount
        strTipInfo = IIf(i = 1, "", strTipInfo & vbCrLf) & rsTmp.Fields(strFieldName)
        If Len(strTipInfo) > 500 Then strTipInfo = strTipInfo & "��": Exit For
        rsTmp.MoveNext
    Next
End Sub

Private Sub GetPathCurrPhase(ByVal bytType As Byte, ByRef lng�׶�ID As Long, ByRef lng���� As Long, Optional ByRef strDate As String)
'--------------------------------------------------
'����:��ȡ����ִ�еǼǻ�����ȡ��ִ�еǼǵĵ�ǰ�׶�
'����:bytType =1 ����ִ��,=2 ����ȡ��ִ��
'--------------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errH

    If bytType = 1 Then
        strSql = " Select *" & vbNewLine & _
                 " From (Select Distinct a.�׶�id, a.����, a.����, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��" & vbNewLine & _
                 "       From ��������·��ִ�� A, ����·����Ŀ B" & vbNewLine & _
                 "       Where a.��Ŀid = b.Id(+) And a.·����¼id = [1] " & _
                 "             And a.ִ��ʱ�� Is Null" & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����" & vbNewLine & _
                 "       Order By Min(a.�Ǽ�ʱ��))" & vbNewLine & _
                 " Where Rownum < 2"
    Else
        strSql = " Select *" & vbNewLine & _
                 " From (Select Distinct a.�׶�id, a.����, a.����, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��" & vbNewLine & _
                 "       From ��������·��ִ�� A, ����·����Ŀ B" & vbNewLine & _
                 "       Where a.��Ŀid = b.Id(+) And a.·����¼id = [1] " & _
                 "             And a.ִ��ʱ�� Is Not Null" & vbNewLine & _
                 "       Group By a.�׶�id, a.����, a.����" & vbNewLine & _
                 "       Order By Min(a.�Ǽ�ʱ��) Desc )  " & vbNewLine & _
                 " Where Rownum < 2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mPP.����·��ID)
    If rsTmp.RecordCount > 0 Then
        lng�׶�ID = Val(NVL(rsTmp!�׶�ID))
        strDate = rsTmp!���� & ""
        lng���� = Val(NVL(rsTmp!����))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathTableChange(ByRef vsBody As VSFlexGrid, ByVal lngPageCOL As Long, Optional vsHead As VSFlexGrid)
'����:����ӡ��ת���ɹ̶���,���ڴ�ӡ���
'��Ҫ�������: ���׶��и߳�����ӡ��Ч��ΧʱҪ����һҳ��������ǰ�׶�ʣ����
'              ÿһ�׶ε������Զ������м��,�޳��հ��С�
'����: ����:vsBody��ӡ����
'      ���:lngPageCOL ��ӡ����(�����̶���)
    Dim lngRow As Long
    Dim lngCol As Long

    On Error Resume Next
    
    Load vsPathPrint(1)
    
    Err.Clear: On Error GoTo 0

    With vsPathPrint(1)
        '���
        .Rows = 0
        .Cols = 0

        If lngPageCOL = 0 Then Exit Sub
        If (vsBody.Cols - vsBody.FixedCols) Mod lngPageCOL <> 0 Then Exit Sub
        
        .Rows = ((vsBody.Cols - vsBody.FixedCols) / lngPageCOL) * vsBody.Rows
        .Cols = vsBody.FixedCols + lngPageCOL
        .FixedCols = vsBody.FixedCols
        .FixedRows = vsBody.FixedRows

        '��vsBody�����ݸ��Ƶ�vsPathPrint(1)
        '�̶���
        For lngCol = 0 To .FixedCols
            lngRow = 0
            Do
                '��ԭ���ǹ̶���ת���ɷǹ̶���ʱ��Ҫ�����Ǳ��ڴ�ӡ����ʶ��
                If lngRow Mod vsBody.Rows < vsBody.FixedRows And lngRow >= vsBody.FixedRows And lngCol = 0 Then
                    .RowData(lngRow) = UCase("FIXEDROW")
                End If
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next
        '�ǹ̶���
        For lngCol = .FixedCols To (.FixedCols + lngPageCOL) - 1
            lngRow = 0
            Do
                Call FuncPathCellCopy(vsBody, vsPathPrint(1), lngRow Mod vsBody.Rows, (lngPageCOL * (lngRow \ vsBody.Rows)) + lngCol, lngRow, lngCol)
                lngRow = lngRow + 1
            Loop While lngRow <> .Rows
        Next

        '��ն��ж��ǿհ׵���
        For lngRow = 0 To .Rows - 1
            For lngCol = 1 To lngPageCOL
                If .RowData(lngRow) = UCase("FIXEDROW") Then
                    .Cell(flexcpAlignment, lngRow, lngCol, lngRow, .Cols - 1) = flexAlignCenterCenter
                    Exit For
                ElseIf .TextMatrix(lngRow, 0) = "ҽ��ǩ��" Then
                    Exit For
                ElseIf .TextMatrix(lngRow, lngCol) <> "" Then
                    Exit For
                ElseIf lngCol = lngPageCOL Then
                    '��¼��Ҫɾ���Ŀհ���
                   .RemoveItem lngRow
                   lngRow = lngRow - 1  'ɾ��һ��,��һ���������
                End If
            Next
            If lngRow = .Rows - 1 Then Exit For
        Next
        '��ʾ�����
        .MergeCol(0) = True
        '�趨���壬���
        .FontSize = IIf(mlngFontSize = 0, CON_SmallFontSize, mlngFontSize) '·������������ӡmlngFontSize=0
        '��ʾ���
        .Cell(flexcpAlignment, 0, .FixedCols, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack
    End With
    Set vsBody = vsPathPrint(1)
End Sub

Private Sub FuncPathCellCopy(ByRef vsSource As VSFlexGrid, ByRef vsCopy As VSFlexGrid, _
        ByVal lngSourRow As Long, ByVal lngSourCol As Long, ByVal lngCopyRow As Long, ByVal lngCopyCol As Long)
'����:���Ƶ�Ԫ��
'������vsSource-��Copy�ı�
'      vsCopy-copy��ı�
'      lngSourRow ,lngSourCol ��Copy�ı���Ӧ�к���
'      lngCopyRow��lngCopyCol Copy�����Ӧ�к���
    With vsCopy
        .Cell(flexcpText, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpText, lngSourRow, lngSourCol)
        .Cell(flexcpAlignment, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpAlignment, lngSourRow, lngSourCol)
        .Cell(flexcpBackColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpBackColor, lngSourRow, lngSourCol)
        .Cell(flexcpForeColor, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpForeColor, lngSourRow, lngSourCol)
        .Cell(flexcpPicture, lngCopyRow, lngCopyCol) = vsSource.Cell(flexcpPicture, lngSourRow, lngSourCol)
    End With
End Sub

Private Sub OutLogModi()
    Dim colSQL As New Collection, i As Long, blnTrans As Boolean

    Call frmPathOutLogOut.ShowMe(mfrmParent, mPati.����ID, mPati.�Һ�ID, 2, colSQL, mPP.·��ID, mPP.����·��ID)

    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    'ִ�г����ǼǱ��SQL
    For i = 1 To colSQL.count
        Call zlDatabase.ExecuteProcedure(colSQL("C" & i), "�޸ĳ����ǼǱ�")
    Next
    gcnOracle.CommitTrans: blnTrans = False

    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
