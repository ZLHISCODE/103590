VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmStoreSpace 
   Caption         =   "�ⷿ��λ����"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11475
   Icon            =   "frmStoreSpace.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11475
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   575
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   13575
      TabIndex        =   2
      Top             =   720
      Width           =   13575
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   $"frmStoreSpace.frx":6852
         Height          =   360
         Left            =   600
         TabIndex        =   3
         Top             =   150
         Width           =   10170
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   0
         Picture         =   "frmStoreSpace.frx":68C9
         Top             =   0
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList imgP 
      Left            =   3840
      Top             =   240
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
            Picture         =   "frmStoreSpace.frx":7193
            Key             =   "pic"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDetails 
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   9795
      TabIndex        =   0
      Top             =   1680
      Width           =   9855
      Begin VB.ComboBox cboRoom 
         Height          =   300
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   2400
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
         Height          =   1245
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   6615
         _cx             =   11668
         _cy             =   2196
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmStoreSpace.frx":D9F5
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfNoStock 
         Height          =   1245
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Visible         =   0   'False
         Width           =   6615
         _cx             =   11668
         _cy             =   2196
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmStoreSpace.frx":DACD
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7125
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStoreSpace.frx":DBC1
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15161
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Bindings        =   "frmStoreSpace.frx":E453
      Left            =   2400
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmStoreSpace.frx":E467
   End
End
Attribute VB_Name = "frmStoreSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MCONMENU_Edit_ADJUST = 100    'δ����ⷿ�Ļ�λ����
Private Const MCONMENU_Edit_ADD = 101
Private Const MCONMENU_Edit_UPDATE = 102
Private Const MCONMENU_Edit_DELETE = 103
Private Const MCONMENU_Edit_HELP = 104
Private Const MCONMENU_Edit_EXIT = 105
Private Const MCONMENU_Edit_ADJUSTSAVE = 106    'δ����ⷿ�Ļ�λ������
Private Const MCONMENU_Edit_ADJUSTEXIT = 107    '�˳�δ����ⷿ�Ļ�λ����

Private mobjPopup As CommandBar
Private mobjControl As CommandBarControl

Private Const mlngBorderColor As Long = &H0&    'ѡ���б߿���ɫ
Private Const mlngNoneBorderColor As Long = &HE0E0E0    ' ûѡ���б߿���ɫ
Private Const mconlngCanColColor As Long = &HE7CFBA    '���޸�����ɫΪ����ɫ
Private mlng�ⷿID As Long
Private mblnNoStock As Boolean  '�Ƿ����û�з����ⷿ�Ļ�λ true-����;false-������
Private mint�༭ģʽ As Integer '0-������ɾ��ģʽ��1-����δ����ⷿ�Ļ�λģʽ


Private Sub GetNoStockDetail()
    '��ѯδ����ⷿ�Ļ�λ
    Dim rsTemp As ADODB.Recordset

    On Error GoTo err

    gstrSql = "Select id,����,����,����,nvl(�ⷿid,0) as �ⷿid,��ע From ҩƷ�ⷿ��λ Where �ⷿid is null Order by ���� "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "GetNoStockDetail")
    
    mblnNoStock = Not rsTemp.EOF
    
    vsfNoStock.Rows = 1
    vsfNoStock.RowHeight(0) = 400
    Do While Not rsTemp.EOF
        With vsfNoStock
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTemp!ID
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("��λ")) = rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsTemp!����, "")
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = NVL(rsTemp!��ע, "")
            
            .RowHeight(.Rows - 1) = 300
        End With

        rsTemp.MoveNext
    Loop
    
'    vsfNoStock.Cell(flexcpBackColor, 1, vsfNoStock.ColIndex("ѡ��"), vsfNoStock.Rows - 1) = mconlngCanColColor

    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitComandBars()
    '��ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    
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
    Me.cbsMain.Icons = Me.imgPublic.Icons
    
    '�Ҽ��˵�
'    Set mobjPopup = cbsMain.Add("Popup", xtpBarPopup)
'    With mobjPopup.Controls
'        Set mobjControl = .Add(xtpControlButton, MCONMENU_Edit_ADJUST, "δ�����λ����")
'    End With
  
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADJUST, "δ�����λ����")
        cbrControlMain.Visible = mblnNoStock
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADD, "����")
        cbrControlMain.BeginGroup = mblnNoStock
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_UPDATE, "�޸�")
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_DELETE, "ɾ��")
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADJUSTSAVE, "�������")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_ADJUSTEXIT, "�˳�����")
        cbrControlMain.Visible = False
        
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_HELP, "����")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_Edit_EXIT, "�˳�")
    End With
    
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    cbsMain.Item(1).Visible = False
End Sub

Private Sub TBFunc_Add()
    If frmStoreSpaceCard.ShowMe(1, Val(cboRoom.ItemData(cboRoom.ListIndex)), 0, Me) = True Then
        Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
    End If
End Sub

Private Sub TBFunc_SetNoStock(ByVal blnBegin As Boolean)
    'δ����ⷿ�Ļ�λ����
    'blnBegin��true-��ʼ����false-��������
    Dim objPopup As CommandBarControl
    
    mint�༭ģʽ = IIf(blnBegin, 1, 0)

    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADD, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_UPDATE, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_DELETE, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADJUST, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = Not blnBegin And mblnNoStock
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADJUSTSAVE, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = blnBegin
    
    Set objPopup = Me.cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_Edit_ADJUSTEXIT, , True)
    If Not objPopup Is Nothing Then objPopup.Visible = blnBegin

    vsfDetails.Visible = Not blnBegin
    vsfNoStock.Visible = blnBegin
    
    If blnBegin Then
        lblComment.Caption = "˵����ѡ���б��еĻ�λ���䵽ָ���ⷿ��˫����ѡ���н���ѡ���ȡ��ѡ��"
    Else
        If mblnNoStock Then
            lblComment.Caption = "˵����1.δ���䵽�ⷿ�Ļ�λ��ѡ��˵���δ�����λ���� 2.˫�����еĻ�λ���б༭ 3.��DEL��ɾ����ǰѡ��Ļ�λ"
        Else
            lblComment.Caption = "˵����1.˫�����еĻ�λ���б༭ 2.��DEL��ɾ����ǰѡ��Ļ�λ"
        End If
    End If
End Sub

Private Sub TBFunc_Update()
    If vsfDetails.Rows = 1 Then Exit Sub
    If vsfDetails.Row < 1 Then Exit Sub
    If vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("����")) = "" Then Exit Sub
    
    If frmStoreSpaceCard.ShowMe(2, Val(cboRoom.ItemData(cboRoom.ListIndex)), Val(vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id"))), Me) Then
        Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
    End If
End Sub

Private Sub cboRoom_Click()
    err = 0: On Error GoTo ErrHand
    
    If mlng�ⷿID = cboRoom.ItemData(cboRoom.ListIndex) Then Exit Sub
    mlng�ⷿID = cboRoom.ItemData(cboRoom.ListIndex)
    
    Call GetDetails(mlng�ⷿID)
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case MCONMENU_Edit_ADJUST  'δ�����λ����
            Call TBFunc_SetNoStock(True)
        Case MCONMENU_Edit_ADD '����
            Call TBFunc_Add
        Case MCONMENU_Edit_UPDATE '�޸�
            Call TBFunc_Update
        Case MCONMENU_Edit_DELETE 'ɾ��
            Call TBFunc_SetDelete
        Case MCONMENU_Edit_ADJUSTSAVE 'δ�����λ������
            Call SetNoStock
        Case MCONMENU_Edit_ADJUSTEXIT '�˳�δ�����λ����
            Call TBFunc_SetNoStock(False)
            
            If cboRoom.ListIndex <> -1 Then
                Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
            End If
        Case MCONMENU_Edit_EXIT '�˳�
            Unload Me
        Case MCONMENU_Edit_HELP '����
            Call TBFunc_SetHelp
    End Select
End Sub

Private Sub TBFunc_SetHelp()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub TBFunc_SetDelete()
    'ɾ����λ
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    
    If vsfDetails.Rows = 1 Then Exit Sub
    If vsfDetails.Row < 1 Then Exit Sub
    If vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("����")) = "" Then Exit Sub
    
    On Error GoTo errH
    
    gstrSql = "Select 1 From ҩƷ��λ���� Where ��λid = [1] And Rownum < 2"
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "SetDelete", vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id")))
    
    If Not rsData.EOF Then
        strMsg = "��ѡ���Ļ�λ�������˴洢ҩƷ���Ƿ�ɾ����"
    Else
        strMsg = "�Ƿ�ɾ��ѡ���Ļ�λ��"
    End If
    
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        gstrSql = "Zl_ҩƷ�ⷿ��λ_Delete("
        'id_In In ҩƷ�ⷿ��λ.id%Type
        gstrSql = gstrSql & vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id"))
        gstrSql = gstrSql & ")"
        Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        MsgBox "ɾ���ɹ���", vbInformation, gstrSysName
        
        vsfDetails.RemoveItem vsfDetails.Row
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub





Private Sub SetNoStock()
    Dim intRow As Integer
    Dim lng�ⷿID As Long
    Dim objCol As New Collection
    
    On Error GoTo err

    With vsfNoStock
        If .Rows <= 1 Then Exit Sub
        If cboRoom.ListIndex = -1 Then Exit Sub
        
        lng�ⷿID = Val(cboRoom.ItemData(cboRoom.ListIndex))

        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, .ColIndex("ѡ��")) = "��" Then
                '�޸�
                gstrSql = "Zl_ҩƷ�ⷿ��λ_Update("
                'ID
                gstrSql = gstrSql & Val(.TextMatrix(intRow, .ColIndex("id")))
                '����
                gstrSql = gstrSql & ",'" & .TextMatrix(intRow, .ColIndex("����")) & "'"
                '����_In   In ҩƷ�ⷿ��λ.����%Type,
                gstrSql = gstrSql & ",'" & .TextMatrix(intRow, .ColIndex("��λ")) & "'"
                '����_In   In ҩƷ�ⷿ��λ.����%Type,
                gstrSql = gstrSql & ",'" & .TextMatrix(intRow, .ColIndex("����")) & "'"
                '�ⷿid_In In ҩƷ�ⷿ��λ.�ⷿid%Type
                gstrSql = gstrSql & "," & lng�ⷿID
                '��ע_In In ҩƷ�ⷿ��λ.��ע%Type
                gstrSql = gstrSql & "," & IIf(.TextMatrix(intRow, .ColIndex("��ע")) = "", "null", "'" & .TextMatrix(intRow, .ColIndex("��ע")) & "'")
                gstrSql = gstrSql & ")"
                
                objCol.Add gstrSql, "_" & objCol.Count + 1
            End If
        Next
    End With

    If objCol.Count = 0 Then
        MsgBox "δѡ���λ�����ܱ��棡", vbInformation, gstrSysName
        Exit Sub
    End If
        
    For intRow = 1 To objCol.Count
        Call zldatabase.ExecuteProcedure(objCol(intRow), "��λ����")
    Next

    MsgBox "��λ����ɹ���", vbInformation, gstrSysName
    
    Call GetNoStockDetail
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picCondition.Move lngLeft, lngTop, lngRight - lngLeft

    Me.picDetails.Move lngLeft, picCondition.Top + picCondition.Height + 50, lngRight - lngLeft, _
        Me.ScaleHeight - Me.picCondition.Top - Me.picCondition.Height - stbThis.Height - 50
End Sub


Private Sub Form_Load()
    Dim rsData As ADODB.Recordset
    
    mint�༭ģʽ = 0
    
    Call GetNoStockDetail
    Call InitComandBars
    Call GetStock
    
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
    
    If mblnNoStock Then
        lblComment.Caption = "˵����1.δ���䵽�ⷿ�Ļ�λ��ѡ��˵���δ�����λ���� 2.˫�����еĻ�λ���б༭ 3.��DEL��ɾ����ǰѡ��Ļ�λ"
    Else
        lblComment.Caption = "˵����1.˫�����еĻ�λ���б༭ 2.��DEL��ɾ����ǰѡ��Ļ�λ"
    End If
End Sub



Private Sub GetStock()
    '��ȡ�ⷿ��Ϣ
    Dim rsTemp As ADODB.Recordset

    On Error GoTo err

    If InStr(1, gstrPrivs, "���пⷿ") > 0 Then
        gstrSql = "Select distinct b.Id, b.����, b.����,b.����" & vbNewLine & _
                "From ��������˵�� A, ���ű� B" & vbNewLine & _
                "Where a.����id = b.Id And (a.�������� Like '%ҩ��' Or a.�������� Like '%ҩ��' Or a.�������� = '�Ƽ���')  order by b.����"
    Else
        gstrSql = "Select distinct b.Id, b.����, b.����, b.����" & vbNewLine & _
                "From ������Ա A, ���ű� B, ��������˵�� C" & vbNewLine & _
                "Where c.����id = b.Id And a.����id = b.Id And (c.�������� Like '%ҩ��' Or c.�������� Like '%ҩ��' Or c.�������� = '�Ƽ���') And a.��Աid = [1] order by b.����"
    End If
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "����", UserInfo.ID)
    
    cboRoom.Clear
    With rsTemp
        Do While Not rsTemp.EOF
            cboRoom.AddItem !���� & "-" & !����
            cboRoom.ItemData(Me.cboRoom.NewIndex) = !ID
            .MoveNext
        Loop
    End With

    If cboRoom.ListCount <= 0 Then
        MsgBox "δ���ÿⷿ��ǰ��Ա�����ڿⷿ���޷����ÿⷿ��λ��", vbExclamation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    Me.cboRoom.ListIndex = 0

    Exit Sub
err:

If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetDetails(ByVal lngDeptID As Long)
    '��ȡ����ⷿ��Ӧ��λ
    Dim rsTemp As ADODB.Recordset

    On Error GoTo err
    
    vsfDetails.RowHeight(0) = 400
    
    If lngDeptID = 0 Then
        vsfDetails.Rows = 1
        vsfDetails.Rows = 2
        Exit Sub
    End If

    gstrSql = "Select id,����,����,����,nvl(�ⷿid,0) as �ⷿid,��ע From ҩƷ�ⷿ��λ Where �ⷿid = [1] Order by ���� "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "��λ", lngDeptID)

    vsfDetails.Rows = 1
    Do While Not rsTemp.EOF
        With vsfDetails
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTemp!ID
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("�ⷿid")) = rsTemp!�ⷿid
            .TextMatrix(.Rows - 1, .ColIndex("��λ")) = rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsTemp!����, "")
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = NVL(rsTemp!��ע, "")
            
            .RowHeight(.Rows - 1) = 300
        End With

        rsTemp.MoveNext
    Loop

    Exit Sub
err:

If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picDetails_Resize()
    vsfDetails.Move 0, cboRoom.Top + cboRoom.Height + 100, picDetails.Width - 20, picDetails.Height - cboRoom.Top - cboRoom.Height - 100
    vsfNoStock.Move vsfDetails.Left, vsfDetails.Top, vsfDetails.Width, vsfDetails.Height
End Sub





Private Sub vsfDetails_DblClick()
    If vsfDetails.Rows = 1 Then Exit Sub
    If vsfDetails.MouseRow < 1 Then Exit Sub
    If vsfDetails.TextMatrix(vsfDetails.MouseRow, vsfDetails.ColIndex("����")) = "" Then Exit Sub
    
    frmStoreSpaceCard.ShowMe 2, Val(cboRoom.ItemData(cboRoom.ListIndex)), Val(vsfDetails.TextMatrix(vsfDetails.Row, vsfDetails.ColIndex("id"))), Me
    Call GetDetails(Val(cboRoom.ItemData(cboRoom.ListIndex)))
End Sub

Private Sub vsfDetails_EnterCell()
    '������ѡ�б߿�
    Dim intRow As Integer
    
    With vsfDetails
        If .Rows <> 1 Then
            For intRow = 0 To .Rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("����"), .Row, .ColIndex("��ע"), mlngBorderColor, 0, 2, 0, 2, 0, 2
            .CellBorderRange .Row, .ColIndex("����"), .Row, .ColIndex("����"), mlngBorderColor, 2, 2, 0, 2, 0, 0
            .CellBorderRange .Row, .ColIndex("��ע"), .Row, .ColIndex("��ע"), mlngBorderColor, 0, 2, 2, 2, 0, 0
        End If
    End With
End Sub

Private Sub vsfDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfDetails
        If KeyCode = vbKeyDelete Then
            If .Rows = 1 Then Exit Sub
            If .Row < 1 Then Exit Sub
            
            Call TBFunc_SetDelete
        End If
    End With
End Sub

Private Sub vsfDetails_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    With vsfDetails
'        If Val(.TextMatrix(.Row, .ColIndex("�ⷿid"))) = 0 And .ColHidden(.ColIndex("����")) = False Then
'            If .Col = .ColIndex("����") Then
'                mobjPopup.ShowPopup
'            End If
'        End If
    End With
End Sub


Private Sub vsfNoStock_DblClick()
    With vsfNoStock
        If .Row < 1 Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .Col = .ColIndex("ѡ��") Then
            If .TextMatrix(.Row, .Col) = "��" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "��"
            End If
        End If
    End With
End Sub

Private Sub vsfNoStock_EnterCell()
    '������ѡ�б߿�
    Dim intRow As Integer
    
    With vsfNoStock
        If .Rows <> 1 Then
            For intRow = 1 To .Rows - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, mlngNoneBorderColor, 0, 0, 0, 0, 0, 0
            Next
            
            .CellBorderRange .Row, .ColIndex("����"), .Row, .ColIndex("��ע"), mlngBorderColor, 0, 2, 0, 2, 0, 2
            .CellBorderRange .Row, .ColIndex("����"), .Row, .ColIndex("����"), mlngBorderColor, 2, 2, 0, 2, 0, 0
            .CellBorderRange .Row, .ColIndex("��ע"), .Row, .ColIndex("��ע"), mlngBorderColor, 0, 2, 2, 2, 0, 0
        End If
    End With
End Sub


