VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmMediPriceBatch 
   Caption         =   "����ִ�е���"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "frmMediPriceBatch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8415
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic��ʾ 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   8175
      TabIndex        =   1
      Top             =   840
      Width           =   8175
      Begin VB.Label Label1 
         Caption         =   "��������Ϊִ���˵��ۣ�����û����Ч�ĵ��ۼ�¼��ͨ���������С�����ִ�е��ۡ����ܿ���ʹ�����б��е��ּ�������Ч��"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8055
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
      Height          =   1575
      Left            =   1080
      TabIndex        =   0
      Top             =   2640
      Width           =   6255
      _cx             =   11033
      _cy             =   2778
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
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
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   1320
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMediPriceBatch.frx":6852
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMediPriceBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mconMenuEecute As Integer = 101
Private Const mconMenuExit As Integer = 102

'�Ӳ�������ȡҩƷ�۸�С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintҩ�ⵥλ As Integer

Private Enum Mcolumn
    mintcol��� = 0
    mintcolid = 1
    mintcolҩƷid = 2
    mintcol����
    mintcol����
    mintcol���
    mintcolԭ��
    mintcol�ּ�
    mintcol������
    mintcolִ������
    mintcol����ϵ��
    mintcolҩ���װ
    mintcolCOUNT = 12
End Enum

Public Sub ShowMe(ByVal objfrm As frmMediLists)
    Me.Show vbModal, objfrm
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case 101
            Call ExecuteSave
        Case 102
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    pic��ʾ.Move lngLeft, lngTop, lngRight - lngLeft
    vsfDetails.Move lngLeft, pic��ʾ.Height + pic��ʾ.Top, lngRight - lngLeft, lngBottom - lngTop - pic��ʾ.Height
End Sub

Private Sub Form_Load()
    Me.Width = 12000
    Me.Height = 8000
    
    '�ж��Ƿ���ҩ�ⵥλ��ʾ
    mintҩ�ⵥλ = Val(zlDatabase.GetPara(29, glngSys))
    
    mintPriceDigit = GetDigit(1, 2, IIf(mintҩ�ⵥλ = 0, 1, 4))
    
    Call InitComandBars
    Call InitVsf
    Call setVSF
    Call getData
End Sub

Private Sub InitComandBars()
    '��ʼ���������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgPicture.Icons
    
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenuEecute, "����ִ�е���")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, mconMenuExit, "�˳�")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
    End With
    
    cbsMain.Item(1).Delete
End Sub

Private Sub InitVsf()
    '��ʼ�����λ�úʹ�С
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    pic��ʾ.Move lngLeft, lngTop, lngRight - lngLeft
    vsfDetails.Move lngLeft, pic��ʾ.Height + pic��ʾ.Top, lngRight - lngLeft, lngBottom - lngTop - pic��ʾ.Height
End Sub

Private Sub setVSF()
    
    With vsfDetails
        .Editable = flexEDNone
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
'        .AllowSelection = True '���ܶ�ѡ��Ԫ��
        .AllowBigSelection = True '����ѡ��
        .SelectionMode = flexSelectionByRow '����ѡ��
        .Rows = 1
        
        .Cols = Mcolumn.mintcolCOUNT
        
        VsfGridColFormat vsfDetails, Mcolumn.mintcol���, "���", 600, flexAlignCenterCenter, "���"
        VsfGridColFormat vsfDetails, Mcolumn.mintcolid, "id", 1000, flexAlignRightCenter, "id"
        VsfGridColFormat vsfDetails, Mcolumn.mintcolҩƷid, "ҩƷid", 1000, flexAlignRightCenter, "ҩƷid"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol����, "����", 2000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol���, "���", 1500, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol������, "������", 1000, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfDetails, Mcolumn.mintcolִ������, "��Ч����", 2000, flexAlignLeftCenter, "��Ч����"
        VsfGridColFormat vsfDetails, Mcolumn.mintcolԭ��, "ԭ��", 1000, flexAlignRightCenter, "ԭ��"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol�ּ�, "�ּ�", 1000, flexAlignRightCenter, "�ּ�"
        VsfGridColFormat vsfDetails, Mcolumn.mintcol����ϵ��, "����ϵ��", 1000, flexAlignRightCenter, "����ϵ��"
        VsfGridColFormat vsfDetails, Mcolumn.mintcolҩ���װ, "ҩ���װ", 1000, flexAlignRightCenter, "ҩ���װ"
        
        .ColHidden(Mcolumn.mintcolid) = True
        .ColHidden(Mcolumn.mintcolҩƷid) = True
        .ColHidden(Mcolumn.mintcol����ϵ��) = True
        .ColHidden(Mcolumn.mintcolҩ���װ) = True
        
    End With
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Private Sub getData()
    '��ȡ���ݵĹ���
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSql = "Select Distinct n.Id, i.Id As ҩƷid, i.����, i.����, i.���, n.������, n.ִ������, n.��ֹ����, n.ԭ��, n.�ּ�, i.���㵥λ, p.ҩ�ⵥλ, p.����ϵ��, p.ҩ���װ" & _
               " From �շ���ĿĿ¼ I, �շѼ�Ŀ N, ҩƷ��� P" & _
               " Where i.Id = n.�շ�ϸĿid And i.Id = p.ҩƷid And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) And" & _
                   " n.�䶯ԭ�� = 0 And Sysdate>n.ִ������" & _
                GetPriceClassString("N") & _
               " Order By n.id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    If rsTemp Is Nothing Then
        Exit Sub
    Else
        Call setColumn(rsTemp)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub setColumn(ByVal rsRecord As ADODB.Recordset)
    Dim i As Integer
    Dim dblPrice As Double
    
    With vsfDetails
        .Rows = rsRecord.RecordCount + 1
        For i = 1 To rsRecord.RecordCount
            If mintҩ�ⵥλ = 1 Then
                dblPrice = IIf(IsNull(rsRecord!ҩ���װ), 0, rsRecord!ҩ���װ)
            Else
                dblPrice = 1
            End If
            
            .TextMatrix(i, Mcolumn.mintcol���) = i
            .TextMatrix(i, Mcolumn.mintcolid) = rsRecord!ID
            .TextMatrix(i, Mcolumn.mintcolҩƷid) = rsRecord!ҩƷid
            .TextMatrix(i, Mcolumn.mintcol����) = rsRecord!����
            .TextMatrix(i, Mcolumn.mintcol����) = rsRecord!����
            .TextMatrix(i, Mcolumn.mintcol���) = rsRecord!���
            .TextMatrix(i, Mcolumn.mintcol������) = IIf(IsNull(rsRecord!������), "", rsRecord!������)
            .TextMatrix(i, Mcolumn.mintcolִ������) = Format(rsRecord!ִ������, "yyyy-mm-dd hh:mm:ss")

            .TextMatrix(i, Mcolumn.mintcolԭ��) = FormatEx(rsRecord!ԭ�� * dblPrice, mintPriceDigit, , True)
            .TextMatrix(i, Mcolumn.mintcol�ּ�) = FormatEx(rsRecord!�ּ� * dblPrice, mintPriceDigit, , True)

            .TextMatrix(i, Mcolumn.mintcol����ϵ��) = IIf(IsNull(rsRecord!����ϵ��), 0, rsRecord!����ϵ��)
            .TextMatrix(i, Mcolumn.mintcolҩ���װ) = IIf(IsNull(rsRecord!ҩ���װ), 0, rsRecord!ҩ���װ)
            .RowHeight(i) = 350
            rsRecord.MoveNext
        Next
    End With
End Sub

Private Sub ExecuteSave()
    'ִ����������
    Dim i As Integer
    On Error GoTo ErrHand
    
    If vsfDetails.Rows <= 1 Then Exit Sub
    For i = 1 To vsfDetails.Rows - 1
        gstrSql = ""
        gstrSql = "Zl_ҩƷ�շ���¼_Adjust(" & vsfDetails.TextMatrix(i, Mcolumn.mintcolid) & ")"
        zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    Next
    MsgBox "����ִ�е��۳ɹ�,���������ּ���Ч��", vbInformation, gstrSysName
    Call getData
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



