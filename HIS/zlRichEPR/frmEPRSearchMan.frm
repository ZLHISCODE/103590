VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRSearchMan 
   Caption         =   "���˲�������"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "frmEPRSearchMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7725
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRSearchMan.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13838
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
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   3720
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   5280
      _cx             =   9313
      _cy             =   6562
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   120
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRSearchMan.frx":0E1C
      Left            =   930
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRSearchMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ID = 0: ����ID: ��ҳID: ������Դ: ����: �Ա�: ����: סԺ����: ��������: ��������: ����: ������Ա: �������: ����ת��: �༭��ʽ: ��ӡ��: ��ӡʱ��
End Enum

Const conPane_Content = 1
Const conPane_Search = 2

'�������
Private mlngDeptId As Long          'ָ����Ĭ����д����id
Private mlngFileID As Long          'ָ�����ҵĲ����ļ�id��0-δָ��; >0,ֻ�����ض��Ĳ����ļ���ͨ�����ڲ����༭�еĲ��Ҳ���;
Private mbytKind As Byte            'Ҫ����ҵ��ļ����ࣺ0-��ʾ�ٴ���д�Ĳ����������Ͳ����ļ�������ͬ
Private mlngSelectId As Long        'ѡ�񷵻ص���ʷ����id

Private mblnPrivacy As Boolean      '�Ƿ���˽����
Private WithEvents mfrmContent As frmDockEPRContent
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mfrmSearch As frmEPRSearchTerms
Attribute mfrmSearch.VB_VarHelpID = -1
Private WithEvents mfrmPrint As frmPrintPreview
Attribute mfrmPrint.VB_VarHelpID = -1
Private mObjTabEprView As cTableEPR
Private mstrPrivs As String
Private mfrmParent As Object        '���������


Public Function ShowSearchFile(frmParent As Form, ByVal lngFileID As Long, Optional lngDeptId As Long) As Long
    '���ܣ�����ָ���Ķ����ļ���������ѡ�����ʷ������¼id�����ڲ����༭�еĲ��Ҳ���
    '���أ�����ѡ�����ʷ������¼id
    'ע�⣺�������´��ڷ�ʽ����
    Dim rsTemp As New ADODB.Recordset
    mlngDeptId = lngDeptId
    mlngFileID = lngFileID
    
    If lngFileID = 0 Then
        MsgBox "����ָ�����ҵ��ļ���", vbExclamation, gstrSysName
        Unload Me: ShowSearchFile = 0: Exit Function
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord("Select ����, ���� From �����ļ��б� Where Id = [1]", frmParent.Caption, lngFileID)
    If rsTemp.RecordCount <= 0 Then
        MsgBox "ָ���ļ����嶪ʧ���޷�������ʷ�ļ���", vbExclamation, gstrSysName
        Unload Me: Exit Function
    End If
    mbytKind = Val("" & rsTemp!����)
    Me.Caption = "���� - " & rsTemp!����
    mlngSelectId = 0
    Me.Show vbModal, frmParent
    ShowSearchFile = mlngSelectId
    Unload Me
End Function

Public Sub ShowSearchReport(frmParent As Form, Optional ByVal lngDeptId As Long)
    '���ܣ���������ָ�����ҵ����Ʊ���
    Set mfrmParent = frmParent
    mlngDeptId = lngDeptId
    mbytKind = cpr���Ʊ���
    mlngFileID = 0
    Me.Caption = "���Ʊ������"
    Me.Show vbModeless, frmParent
End Sub

Public Sub ShowSearchClinic(frmParent As Form, Optional ByVal lngDeptId As Long)
    '���ܣ���������ָ�����ҵ��ٴ�����
    mlngDeptId = lngDeptId
    mbytKind = 0
    mlngFileID = 0
    Me.Caption = "���˲�������"
    Me.Show vbModeless, frmParent
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim strItemKey As String, blnMoved As Boolean, i As Integer, strPrint As String
Dim cbrControl As CommandBarControl
    
    blnMoved = vfgThis.TextMatrix(Me.vfgThis.ROW, mCol.����ת��) = "1" 'ת�����󱸱�
    
    Select Case Control.ID
    Case conMenu_File_Open
        If blnMoved Then
            MsgBox "�ò��˵ı��������Ѿ�ת���������ݿ⣬�����������", vbInformation, gstrSysName
            Exit Sub
        Else
            If vfgThis.TextMatrix(vfgThis.ROW, mCol.�༭��ʽ) <> 0 Then Exit Sub
            If mlngFileID <= 0 Then
                Dim frmOpen As New frmEPRView
                frmOpen.ShowMe Me, CLng(Me.vfgThis.Cell(flexcpText, Me.vfgThis.ROW, mCol.ID)), True
            Else
                mlngSelectId = Me.vfgThis.TextMatrix(Me.vfgThis.ROW, mCol.ID): Me.Hide
            End If
        End If
    Case conMenu_File_Preview, conMenu_File_Print
        If vfgThis.TextMatrix(vfgThis.ROW, mCol.�༭��ʽ) = 0 Then
            '�������ļ���ͬ���ﲡ����ӡ��������ͬҳ��ϲ���ӡ����
            Set mfrmPrint = New frmPrintPreview
            mfrmPrint.DoMultiDocPreview Me, cpr���ﲡ��, , , , , CLng(Me.vfgThis.Cell(flexcpText, Me.vfgThis.ROW, mCol.ID)), (Control.ID = conMenu_File_Print), , , blnMoved
            Unload mfrmPrint: Set mfrmPrint = Nothing 'ByZT:����Load��δ��ʾ��û����Ϊ�رյ������VB�����Զ�Unload
        ElseIf vfgThis.TextMatrix(vfgThis.ROW, mCol.�༭��ʽ) = 1 Then
            Set mObjTabEprView = New cTableEPR
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, CLng(vfgThis.TextMatrix(vfgThis.ROW, mCol.ID)), False
            mObjTabEprView.zlPrintDoc Me, IIf(Control.ID = conMenu_File_Print, False, True)
            Set mObjTabEprView = Nothing
        ElseIf vfgThis.TextMatrix(vfgThis.ROW, mCol.�༭��ʽ) = 2 And Control.ID = conMenu_File_Print Then
            On Error GoTo errHand
            strPrint = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "PrintName", "")
            Dim objInfection As Object
            Set objInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "��Ⱦ�����濨", True)
            If Not objInfection Is Nothing Then
                objInfection.Init gcnOracle, glngSys
            End If
            objInfection.PrintDoc Me, CLng(vfgThis.TextMatrix(vfgThis.ROW, mCol.����ID)), CLng(vfgThis.TextMatrix(vfgThis.ROW, mCol.��ҳID)), CLng(vfgThis.TextMatrix(vfgThis.ROW, mCol.ID)), strPrint
        End If
    Case conMenu_File_Print * 100
        strItemKey = frmEPRSearchPrint.ShowMe(Me, vfgThis, strPrint)
        For i = 0 To UBound(Split(strItemKey, "|"))
            If Split(Split(strItemKey, "|")(i), ",")(0) = 0 Then
                Set mfrmPrint = New frmPrintPreview
                mfrmPrint.DoMultiDocPreview Me, cpr���ﲡ��, , , , , Val(Split(Split(strItemKey, "|")(i), ",")(1)), True, , True, blnMoved, , strPrint
                Unload mfrmPrint: Set mfrmPrint = Nothing 'ByZT:����Load��δ��ʾ��û����Ϊ�رյ������VB�����Զ�Unload
            Else
                Set mObjTabEprView = New cTableEPR
                mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, Val(Split(Split(strItemKey, "|")(i), ",")(1)), False
                mObjTabEprView.zlPrintDoc Me, False
                Set mObjTabEprView = Nothing
            End If
        Next
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_BatPrint '�嵥��ӡ
        Call zlRptPrint(1)
    Case conMenu_File_Exit
        If mlngFileID <= 0 Then
            Unload Me
        Else
            mlngSelectId = 0: Me.Hide
        End If
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh:
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)

    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        'ִ�з�������ǰģ��ı���
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "ID=" & vfgThis.TextMatrix(vfgThis.ROW, mCol.ID), "����ID=" & vfgThis.TextMatrix(vfgThis.ROW, mCol.����ID), "��ҳID=" & vfgThis.TextMatrix(vfgThis.ROW, mCol.��ҳID))
        End If
    End Select
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    Me.vfgThis.Move lngScaleLeft, lngScaleTop, lngScaleRight - lngScaleLeft, lngScaleBottom - lngScaleTop
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open
        Control.Enabled = (Val(Me.vfgThis.TextMatrix(Me.vfgThis.ROW, mCol.ID)) <> 0)
        If Control.Enabled Then Control.Enabled = vfgThis.TextMatrix(vfgThis.ROW, mCol.�༭��ʽ) <> 2
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Print * 100
        Dim strKind As String
        Dim strRight As String
        strKind = Me.vfgThis.TextMatrix(Me.vfgThis.ROW, mCol.��������)
        If strKind = "����֤���뱨��" Then
            strRight = "����֤��"
        ElseIf strKind = "֪���ļ�" Then
            strRight = "֪���ļ�"
        ElseIf strKind = "������" Then
            strRight = "������"
        End If
        Control.Enabled = (Val(Me.vfgThis.TextMatrix(Me.vfgThis.ROW, mCol.ID)) <> 0)
        If Control.Enabled Then Control.Enabled = InStr(1, mstrPrivs, "��ӡ���")
        If Control.Enabled Then Control.Enabled = InStr(1, mstrPrivs, strRight)
        If Control.ID = conMenu_File_Preview Then
            If Control.Enabled Then Control.Enabled = vfgThis.TextMatrix(vfgThis.ROW, mCol.�༭��ʽ) <> 2
        End If
    Case conMenu_File_Excel, conMenu_File_BatPrint
        Control.Enabled = (Val(Me.vfgThis.TextMatrix(Me.vfgThis.ROW, mCol.ID)) <> 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Content
        Item.Handle = mfrmContent.hWnd
    Case conPane_Search
        Item.Handle = mfrmSearch.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim lngCount As Long

    mstrPrivs = GetPrivFunc(glngSys, 1273)
    If mlngFileID <= 0 Then
        Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Else
        Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    End If
    
    '������˽
    mblnPrivacy = (InStr(gstrPrivsEpr, ";������˽����;") = 0)
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
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        If mlngFileID <= 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��(&O)��"): cbrControl.BeginGroup = True
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "ѡ�񸲸ǵ�ǰ�༭�ļ�(&O)��"): cbrControl.BeginGroup = True
        End If
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print * 100, "ѡ���ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "�嵥��ӡ(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
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
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        If mlngFileID <= 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��"): cbrControl.BeginGroup = True
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "ѡ�񸲸ǵ�ǰ�༭�ļ�"): cbrControl.BeginGroup = True
        End If
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next

    '��ȡ��������ģ��ı���:��Ϊ��һ���Զ�ȡ,ȫ�ֱ�������
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)

    '-----------------------------------------------------
    '��������
    Dim panContent As Pane, panSearch As Pane
    If mfrmContent Is Nothing Then Set mfrmContent = New frmDockEPRContent
    If mfrmSearch Is Nothing Then Set mfrmSearch = New frmEPRSearchTerms
    mfrmSearch.mlngDeptId = mlngDeptId
    mfrmSearch.mbytKind = mbytKind
    mfrmSearch.mlngFileID = mlngFileID
    mfrmSearch.mstrPrivs = mstrPrivs
    
    Set panContent = dkpMan.CreatePane(conPane_Content, 400, 150, DockBottomOf, Nothing)
    panContent.Title = "�ļ�����"
    panContent.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set panSearch = dkpMan.CreatePane(conPane_Search, 280, 400, DockLeftOf, Nothing)
    panSearch.Title = "��������"
    panSearch.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = False
    '-----------------------------------------------------
    
    If mlngFileID = 0 Then Call RestoreWinState(Me, App.ProductName)
    With Me.vfgThis
        .Rows = .FixedRows: .Cols = 17
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.����ID) = 0: .ColWidth(mCol.��ҳID) = 0
        .ColWidth(mCol.������Դ) = 600: .ColWidth(mCol.����) = 800: .ColWidth(mCol.�Ա�) = 500: .ColWidth(mCol.����) = 800: .ColWidth(mCol.סԺ����) = 800
        .ColWidth(mCol.��������) = IIf(mbytKind = cpr���Ʊ���, 0, 1200): .ColWidth(mCol.��������) = 1600
        .ColWidth(mCol.����) = 1100: .ColWidth(mCol.������Ա) = 700: .ColWidth(mCol.�������) = 1600
        .ColWidth(mCol.����ת��) = 0: .ColWidth(mCol.�༭��ʽ) = 0: .ColWidth(mCol.��ӡ��) = 700
        .ColWidth(mCol.��ӡʱ��) = 1600
        
        .TextMatrix(0, mCol.������Դ) = "��Դ": .TextMatrix(0, mCol.����) = "����": .TextMatrix(0, mCol.�Ա�) = "�Ա�"
        .TextMatrix(0, mCol.��������) = "����": .TextMatrix(0, mCol.��������) = "�ļ�����": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.סԺ����) = "סԺ����"
        .TextMatrix(0, mCol.����) = "����": .TextMatrix(0, mCol.������Ա) = "��д��": .TextMatrix(0, mCol.�������) = "��д����"
        .TextMatrix(0, mCol.��ӡ��) = "��ӡ��": .TextMatrix(0, mCol.��ӡʱ��) = "��ӡʱ��"
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftCenter
            .ColHidden(lngCount) = IIf(.ColWidth(lngCount) = 0, True, False)
        Next
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmContent
    Unload mfrmSearch
    Set mfrmContent = Nothing
    Set mfrmSearch = Nothing
    If mlngFileID = 0 Then Call SaveWinState(Me, App.ProductName)
    Set mfrmParent = Nothing
    Set mfrmPrint = Nothing
End Sub

Private Sub mfrmContent_DblClick()
    Call vfgThis_DblClick
End Sub

Private Sub mfrmPrint_PrintEpr(ByVal lngRecordId As Long)
    Event_AfterPrinted lngRecordId
End Sub

Private Sub mfrmSearch_SearchClick(rsResult As ADODB.Recordset)
Dim lngCount As Long
    With Me.vfgThis
        .Redraw = flexRDNone
        Set .DataSource = rsResult
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.����ID) = 0: .ColWidth(mCol.��ҳID) = 0
        .ColWidth(mCol.������Դ) = 600: .ColWidth(mCol.����) = 800: .ColWidth(mCol.�Ա�) = 500: .ColWidth(mCol.����) = 800: .ColWidth(mCol.סԺ����) = 800
        .ColWidth(mCol.��������) = IIf(mbytKind = cpr���Ʊ���, 0, 1200): .ColWidth(mCol.��������) = 1600
        .ColWidth(mCol.����) = 1100: .ColWidth(mCol.������Ա) = 700: .ColWidth(mCol.�������) = 1600
        .ColWidth(mCol.����ת��) = 0: .ColWidth(mCol.�༭��ʽ) = 0: .ColWidth(mCol.��ӡ��) = 700
        .ColWidth(mCol.��ӡʱ��) = 1600
        
        .TextMatrix(0, mCol.������Դ) = "��Դ": .TextMatrix(0, mCol.����) = "����": .TextMatrix(0, mCol.�Ա�) = "�Ա�"
        .TextMatrix(0, mCol.��������) = "����": .TextMatrix(0, mCol.��������) = "�ļ�����": .TextMatrix(0, mCol.����) = "����"
        .TextMatrix(0, mCol.סԺ����) = "סԺ����"
        .TextMatrix(0, mCol.����) = "����": .TextMatrix(0, mCol.������Ա) = "��д��": .TextMatrix(0, mCol.�������) = "��д����"
        .TextMatrix(0, mCol.��ӡ��) = "��ӡ��": .TextMatrix(0, mCol.��ӡʱ��) = "��ӡʱ��"
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftCenter
            .ColHidden(lngCount) = IIf(.ColWidth(lngCount) = 0, True, False)
        Next
        For lngCount = .FixedRows To .Rows - 1
            If mblnPrivacy Then .TextMatrix(lngCount, mCol.����) = "******"
            Select Case Val(.TextMatrix(lngCount, mCol.������Դ))
            Case 1: .TextMatrix(lngCount, mCol.������Դ) = "����"
            Case 2: .TextMatrix(lngCount, mCol.������Դ) = "סԺ"
            Case 3: .TextMatrix(lngCount, mCol.������Դ) = "����"
            Case 4: .TextMatrix(lngCount, mCol.������Դ) = "���"
            End Select
            Select Case Val(.TextMatrix(lngCount, mCol.��������))
            Case 1: .TextMatrix(lngCount, mCol.��������) = "���ﲡ��"
            Case 2: .TextMatrix(lngCount, mCol.��������) = "סԺ����"
            Case 4: .TextMatrix(lngCount, mCol.��������) = "������"
            Case 5: .TextMatrix(lngCount, mCol.��������) = "����֤���뱨��"
            Case 6: .TextMatrix(lngCount, mCol.��������) = "֪���ļ�"
            Case 7: .TextMatrix(lngCount, mCol.��������) = "���Ʊ���"
            End Select
            .TextMatrix(lngCount, mCol.�������) = Format(.TextMatrix(lngCount, mCol.�������), "yyyy-MM-dd hh:mm")
        Next
        If .Rows > .FixedRows Then .ROW = .FixedRows
        .Redraw = flexRDDirect
        .Tag = ""
    End With
    Call vfgThis_RowColChange
    If rsResult.RecordCount > 0 Then
        Me.stbThis.Panels(2).Text = "�����ҵ� " & rsResult.RecordCount & "�ݲ���"
    Else
        Me.stbThis.Panels(2).Text = "û�з��������Ĳ���"
    End If
End Sub

Private Sub vfgThis_DblClick()
Dim cbrControl As CommandBarControl
    With Me.vfgThis
        If .MouseRow = 0 Then Exit Sub
        If Val(.TextMatrix(.ROW, mCol.ID)) = 0 Then Exit Sub
    End With
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_File_Open)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub vfgThis_RowColChange()
    Dim lngRecordId As Long
    
    Err = 0: On Error Resume Next
    With Me.vfgThis
        If .Cols < mCol.ID + 1 Then Exit Sub
        lngRecordId = Val(.TextMatrix(.ROW, mCol.ID))
    End With
    Err = 0: On Error GoTo 0
    If Me.Tag = "" And (Val(Me.vfgThis.Tag) <> Me.vfgThis.ROW) Then
        Call mfrmContent.zlRefresh(lngRecordId, "", True, Val(vfgThis.TextMatrix(vfgThis.ROW, mCol.����ת��)) = 1, , Val(vfgThis.TextMatrix(vfgThis.ROW, mCol.�༭��ʽ)))
        Me.vfgThis.Tag = Me.vfgThis.ROW
    End If
End Sub

'----------------------------------------------------
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgThis
    objPrint.Title.Text = IIf(mbytKind = cpr���Ʊ���, "��������嵥", "���������嵥")
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objPrint.UnderAppRows.Add(objAppRow)
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
End Sub
Public Sub Event_AfterPrinted(lngRecordId As Long)
    If Not mfrmParent Is Nothing Then '���ش�ӡ״̬
        If InStr(mfrmParent.Caption, "���Ʊ������") > 0 Then
            Call mfrmParent.Event_AfterPrinted(lngRecordId)
        End If
    End If
End Sub
