VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmMarkMapMan 
   Caption         =   "�������ͼ����"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   Icon            =   "frmFigureMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picVBar_S 
      BackColor       =   &H00808080&
      Height          =   4080
      Left            =   3495
      MouseIcon       =   "frmFigureMan.frx":058A
      MousePointer    =   99  'Custom
      ScaleHeight     =   4080
      ScaleWidth      =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   990
      Width           =   40
   End
   Begin MSComctlLib.ImageList imlTool 
      Left            =   855
      Top             =   225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":06DC
            Key             =   ""
            Object.Tag             =   "301"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":0C76
            Key             =   ""
            Object.Tag             =   "302"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":1210
            Key             =   ""
            Object.Tag             =   "303"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":17AA
            Key             =   ""
            Object.Tag             =   "304"
         EndProperty
      EndProperty
   End
   Begin zlRichEPR.ucCanvas Canvas 
      Height          =   1140
      Left            =   4005
      TabIndex        =   2
      Top             =   1260
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
   End
   Begin MSComctlLib.ImageList imglvw 
      Left            =   2850
      Top             =   4695
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
            Picture         =   "frmFigureMan.frx":1D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFigureMan.frx":85A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   4230
      Left            =   15
      TabIndex        =   0
      Top             =   930
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   7461
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglvw"
      SmallIcons      =   "imglvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   5850
      Top             =   585
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"frmFigureMan.frx":EE08
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5565
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3043
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1235
            MinWidth        =   1235
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Left            =   210
      Top             =   210
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmMarkMapMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ID_FitSize = 301                      '�ʺϳߴ�
Private Const ID_ActualSize = 302                   'ʵ�ʳߴ�
Private Const ID_ZoomIn = 303                       '�Ŵ�
Private Const ID_ZoomOut = 304                      '��С

'���弶����
Private mlngScaleLeft As Long, mlngScaleTop As Long, mlngScaleRight As Long, mlngScaleBottom As Long  '�ͻ�����Ĵ�С
Private mstrPrivs As String

'��ʱ����
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem

'################################################################################################################
'-- λͼ����
Private WithEvents DIBFilter As cDIBFilter      ' DIB �˾�����(24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Private WithEvents DIBDither As cDIBDither      ' DIB ��������(1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Private DIBPal               As New cDIBPal     ' DIB ��ɫ����� (1, 4, 8 bpp)
Private DIBSave              As New cDIBSave    ' Save ���� (BMP)  (1, 4, 8, 24 bpp)
Private DIBbpp               As Byte            ' ��ǰ��ɫ���
Private WithEvents cPicEditor As cPictureEditor     ' ͼƬ�༭����
Attribute cPicEditor.VB_VarHelpID = -1
Private m_LastFilename As String                    ' ���򿪵�ͼƬλ��
Private m_Temp As String                            ' ��ʱ�ļ�·��
Private m_AppID As Long
'-- GDI+
Private m_GDIpToken         As Long         ' ���ڹر� GDI+
'ɨ�躯������
Private Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hWnd As Long, ByVal wPixTypes As Integer) As Integer
Private Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hWnd As Long) As Integer

Private Sub Canvas_Crop()
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    With Canvas.DIB
        stbThis.Panels(3).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
    End With
End Sub

Private Sub Canvas_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = vbRightButton Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        
        Set Popup = Me.cbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "ͼƬ(&I)��"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, ID_FitSize, "�ʺϳߴ�(&F)"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, ID_ActualSize, "ʵ�ʳߴ�(&S)")
            Set cbrControl = .Add(xtpControlButton, ID_ZoomIn, "�Ŵ�(&Z)")
            Set cbrControl = .Add(xtpControlButton, ID_ZoomOut, "��С(&O)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub DoZoomMenu(Index As Integer)
    Select Case Index
        Case 0 '-- Zoom +
            Canvas.Zoom = Canvas.Zoom + IIf(Canvas.Zoom < 25, 1, 0)
            Canvas.FitMode = False
        Case 1 '-- Zoom -
            Canvas.Zoom = Canvas.Zoom - IIf(Canvas.Zoom > 1, 1, 0)
            Canvas.FitMode = False
        Case 2 '-- 1 : 1
            Canvas.Zoom = 1
            Canvas.FitMode = False
        Case 3 '-- Best fit
            Canvas.FitMode = True
    End Select
    Call Canvas.Resize
    stbThis.Panels(4).Text = Format(Canvas.Zoom, "0%")
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strItemKey As String
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    Case ID_FitSize
        DoZoomMenu 3
    Case ID_ActualSize
        DoZoomMenu 2
    Case ID_ZoomIn
        DoZoomMenu 0
    Case ID_ZoomOut
        DoZoomMenu 1
    Case conMenu_Edit_NewItem
        strItemKey = frmMarkMapEdit.ShowMe(Me, True)
        If strItemKey <> "" Then Call zlRefLists(strItemKey)
    Case conMenu_Edit_Modify
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        
        Dim strFileName1 As String, strTmp1 As String, oDIB As New cDIB, bSuccess1 As Boolean
        
        strFileName1 = m_Temp & "\R" & m_AppID & ".jpg"
        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName1, [ImageJPEG], 90)         '100%��ͼƬ���������޸�ʱ����ʧ����
        Call oDIB.CreateFromStdPicture(pvGetStdPicture(strFileName1, bSuccess1), DIBPal, DIBDither)
        
        strItemKey = frmMarkMapEdit.ShowMe(Me, False, Mid(Me.lvwList.SelectedItem.Key, 2), oDIB)
        If strItemKey <> "" Then Call zlRefLists(strItemKey): Me.Canvas.Resize
    Case conMenu_Edit_Delete
        With Me.lvwList
            If .SelectedItem Is Nothing Then Exit Sub
            If MsgBox("���ɾ���ñ��ͼ��" & vbCrLf & "����" & .SelectedItem.Text, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            gstrSQL = "zl_�������ͼ��_delete('" & Mid(.SelectedItem.Key, 2) & "')"
            Err = 0: On Error GoTo errHand
            Call SQLTest(App.ProductName, Me.Caption, gstrSQL): gcnOracle.Execute gstrSQL, , adCmdStoredProc: Call SQLTest
            Call .ListItems.Remove(.SelectedItem.Key)
            If Not .SelectedItem Is Nothing Then
                Call lvwList_ItemClick(.SelectedItem)
            Else
                stbThis.Panels(3).Text = ""
                Set Canvas.DIB = New cDIB
                Canvas.Resize
            End If
            Me.stbThis.Panels(2).Text = "ʣ��" & .ListItems.Count & "���ͼ"
            If .Visible And .Enabled Then .SetFocus
        End With
        Exit Sub
errHand:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
        Exit Sub
        
    Case conMenu_Edit_MarkMap
        Dim strFileName As String, bSuccess As Boolean, strTmp As String
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        dlgThis.InitDir = m_LastFilename
        dlgThis.CancelError = True
        On Error GoTo LL
        dlgThis.ShowOpen
        strFileName = dlgThis.Filename
        If gobjFSO.FileExists(strFileName) Then
            If MsgBox("ע�⣺ѡ����ͼƬ�����ǵ�ǰͼƬ���Ƿ������", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        If Trim(strFileName) <> "" Then
            '-- Create DIB
            DoEvents
            Call pvSetDIBPicture(pvGetStdPicture(strFileName, bSuccess))
            
            If (bSuccess) Then
                m_LastFilename = strFileName
                stbThis.Panels(3).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
                stbThis.Panels(4).Text = Format(Canvas.Zoom, "0%")
            End If
        End If
        strFileName = m_Temp & "\R" & m_AppID & ".jpg"
        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName, [ImageJPEG], 90)         '90%��ͼƬ����������ѹ��
        
        Dim arySql() As String, lngSql As Long
        If gobjFSO.FileExists(strFileName) Then
            If zlBlobSql(0, Mid(Me.lvwList.SelectedItem.Key, 2), strFileName, arySql()) = False Then
                MsgBox "���ͼ�α���ʧ��", vbExclamation, gstrSysName
                Exit Sub
            End If
            gobjFSO.DeleteFile strFileName  'ɾ����ʱ�ļ�
        End If
        
        'ִ�б���
        Err = 0: On Error GoTo ErrMap
        gcnOracle.BeginTrans
        For lngSql = LBound(arySql) To UBound(arySql)
            Call SQLTest(App.ProductName, Me.Caption, arySql(lngSql))
            gcnOracle.Execute arySql(lngSql), , adCmdStoredProc
            Call SQLTest
        Next
        gcnOracle.CommitTrans
        Exit Sub
ErrMap:
        gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
        Exit Sub
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
        If Me.lvwList.SelectedItem Is Nothing Then
            strItemKey = ""
        Else
            strItemKey = Mid(Me.lvwList.SelectedItem.Key, 2)
        End If
        Call zlRefLists(strItemKey)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
LL:
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Call Me.cbsThis.GetClientRect(mlngScaleLeft, mlngScaleTop, mlngScaleRight, mlngScaleBottom)
    Call Form_Resize
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
            Control.Visible = Not (InStr(1, mstrPrivs, "��ɾ��") = 0 And InStr(1, mstrPrivs, "ͼƬ����") = 0)
        End Select
    Else
        Err = 0: On Error Resume Next
        Select Case Control.ID
        Case ID_FitSize
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
            Control.Checked = (Canvas.FitMode = True)
        Case ID_ActualSize
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
        Case ID_ZoomIn
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
        Case ID_ZoomOut
            Control.Enabled = (Canvas.DIB.hDIB <> 0)
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = Not (Me.lvwList.ListItems.Count = 0)
        Case conMenu_Edit_NewItem
            Control.Visible = Not (InStr(1, mstrPrivs, "��ɾ��") = 0)
        Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_MarkMap
            Control.Visible = Not (InStr(1, mstrPrivs, "��ɾ��") = 0)
            Control.Enabled = Not (Me.lvwList.SelectedItem Is Nothing)
        Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        End Select
    End If
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'ͼƬ�������
    m_LastFilename = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", App.Path)
    
    Dim GpInput As GdiplusStartupInput
    '-- ���� GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("���� GDI+ �����޷�����ͼƬ���룡���� GDI+ DLL �Ƿ���ڻ����𻵣�", vbInformation + vbOKOnly)
        Call Unload(Me)
        Exit Sub
    End If
    
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    m_AppID = Me.hWnd
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    Set cPicEditor = New cPictureEditor
    
    Canvas.FitMode = True
    
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
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
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "ͼƬ(&I)��"): cbrControl.BeginGroup = True
        
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, ID_FitSize, "�ʺϳߴ�(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ID_ActualSize, "ʵ�ʳߴ�(&S)")
        Set cbrControl = .Add(xtpControlButton, ID_ZoomIn, "�Ŵ�(&Z)")
        Set cbrControl = .Add(xtpControlButton, ID_ZoomOut, "��С(&O)")
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
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("I"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, vbKeySubtract, ID_ZoomOut
        .Add 0, vbKeyAdd, ID_ZoomIn
        .Add FCONTROL, vbKeyF, ID_FitSize
        .Add FCONTROL, vbKeyS, ID_ActualSize
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
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "ͼƬ")
        cbrControl.BeginGroup = True
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '����Ԫ����̬����
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 650
        .Add , "_����", "����", 1000
    End With
    With Me.lvwList
        .ColumnHeaders("_����").Position = 1
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    '-----------------------------------------------------
    '���븽��ͼ��
    Me.cbsThis.AddImageList Me.imlTool
    '����ָ�
    Call Me.cbsThis.GetClientRect(mlngScaleLeft, mlngScaleTop, mlngScaleRight, mlngScaleBottom)
    Call RestoreWinState(Me, App.ProductName)
    Me.picVBar_S.BackColor = Me.BackColor
    
    '-----------------------------------------------------
    '����װ��
    Call zlRefLists
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.picVBar_S
        .Top = mlngScaleTop: .Height = mlngScaleBottom - mlngScaleTop
        If .Left < 2000 + mlngScaleLeft Then .Left = 2000 + mlngScaleLeft
        If .Left > mlngScaleRight - mlngScaleLeft - 2000 Then .Left = mlngScaleRight - mlngScaleLeft - 2000
    End With
    With Me.lvwList
        .Left = mlngScaleLeft: .Width = Me.picVBar_S.Left - .Left
        .Top = mlngScaleTop: .Height = mlngScaleBottom - .Top
    End With
    With Me.Canvas
        .Top = Me.lvwList.Top: .Height = mlngScaleBottom - .Top
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width: .Width = mlngScaleRight - .Left
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", m_LastFilename

    Me.Canvas.DIB.Destroy
    LockWindowUpdate 0
    UpdateWindow Me.hWnd
    ' Unload the GDI+ Dll
    Call mGdIpEx.GdiplusShutdown(m_GDIpToken)

    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    Set cPicEditor = Nothing
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgMark_DblClick()
    Set cbrControl = Me.cbsThis.FindControl(xtpControlButton, conMenu_Edit_MarkMap)
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwList.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwList.SortOrder = IIf(Me.lvwList.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwList.SortKey = ColumnHeader.Index - 1
        Me.lvwList.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwList_DblClick()
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub lvwList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim strTemp As String, bSuccess As Boolean
    
    If lvwList.Tag = Item Then Exit Sub
    Screen.MousePointer = vbHourglass
    stbThis.Panels(3).Text = ""
    Set Canvas.DIB = New cDIB
    strTemp = zlBlobRead(0, Mid(Item.Key, 2))
    If Len(strTemp) > 0 Then
        Call pvSetDIBPicture(pvGetStdPicture(strTemp, bSuccess))
        If bSuccess Then
            stbThis.Panels(3).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
        Else
            If Err <> 0 Then MsgBox "ͼƬ���ܱ��𻵣�", vbExclamation, gstrSysName
        End If
        Kill strTemp
    End If
    Canvas.Resize
    lvwList.Tag = Item
    Screen.MousePointer = vbNormal
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvwList_DblClick
End Sub

Private Sub lvwList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�����˵�����
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Or Me.cbsThis.ActiveMenuBar.Controls(2).Visible <> True Then Exit Sub
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        If cbrControl.Visible Then
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
            cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        End If
    Next
    Set cbrControl = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_File_Print, "��ӡ")
    cbrControl.BeginGroup = True
    
    cbrPopupBar.ShowPopup
End Sub

Private Sub picVBar_S_DblClick()
    If lvwList.ListItems.Count > 0 Then
        picVBar_S.Left = lvwList.ListItems(1).Width + Screen.TwipsPerPixelX * 4
        Call Form_Resize
    End If
End Sub

Private Sub picVBar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.picVBar_S.Left = Me.picVBar_S.Left + X
End Sub

Private Sub picVBar_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Public Sub zlRefLists(Optional strKeyCode As String)
    '---------------------------------------------
    '��д�б�
    '---------------------------------------------
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select ����,����,���� From �������ͼ�� Order By ����"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !����, !����)
            objItem.SubItems(Me.lvwList.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwList.ColumnHeaders("_����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.Icon = 1: objItem.SmallIcon = objItem.Icon
            If !���� = strKeyCode Then objItem.Selected = True
            .MoveNext
        Loop
    End With
    If Me.lvwList.ListItems.Count > 0 Then
        If Me.lvwList.SelectedItem Is Nothing Then Me.lvwList.ListItems(1).Selected = True
        Me.lvwList.SelectedItem.EnsureVisible
        Call lvwList_ItemClick(Me.lvwList.SelectedItem)
        Me.stbThis.Panels(2).Text = "����" & Me.lvwList.ListItems.Count & "���ͼ"
    Else
        stbThis.Panels(3).Text = ""
        Set Canvas.DIB = New cDIB
        Canvas.Resize
        Me.stbThis.Panels(2).Text = ""
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    If Me.lvwList.ListItems.Count = 0 Then Exit Sub
    
    Err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwList
    objPrint.Title.Text = "�������ͼ�嵥"
    objPrint.BelowAppItems.Add "��ӡʱ��:" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ������غ���
'################################################################################################################
Private Function pvGetStdPicture(ByVal sFileName As String, bSuccess As Boolean) As StdPicture
    On Error Resume Next
    If (pvGetExt(sFileName) = "png" Or pvGetExt(sFileName) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGdIpEx.LoadPictureEx(sFileName)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFileName)
    End If
    
    '-- Is there an image ?
    bSuccess = Not (pvGetStdPicture Is Nothing)
    
    If (bSuccess = False) Then
        '-- Nothing loaded
        Call MsgBox("����ͼƬʱ�����������", vbExclamation)
    End If

    On Error GoTo 0
End Function
    
Private Sub pvSetDIBPicture(Image As StdPicture)
  Static lstW As Long
  Static lstH As Long

    If (Not Picture Is Nothing) Then

        '-- Save last DIB dimensions
        lstW = Canvas.DIB.Width
        lstH = Canvas.DIB.Height
        
        '-- Clear palette
        Call DIBPal.Clear
        
        '-- Create 32bpp DIB section from std. picture.
        '   Case source <=8bpp, palette saved in DIBPal, palette indexes in DIBDither.
        '   Return value: source color depth / 0 = Err.
        DIBbpp = Canvas.DIB.CreateFromStdPicture(Image, DIBPal, DIBDither)
        
        '-- Select current depth mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Remove Crop rectangle and resize canvas
        Call Canvas.RemoveCropRectangle
        With Canvas.DIB
            If (lstW <> .Width Or lstH <> .Height) Then
                Call Canvas.Resize
              Else
                Call Canvas.Repaint
            End If
        End With
        
        '-- Show image info: Size + bpp
'        stbThis.Panels(3).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
        stbThis.Panels(4).Text = Format(Canvas.Zoom, "0%")
    End If
End Sub

Private Sub pvSetPalMode(ByVal bpp As Long)
  Dim lIdxNew As Long
  Dim lIdxOld As Long
    
    Select Case bpp
        Case 1  '-- 2 colors / Black and White
            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
        Case 4  '-- 16 colors / 16 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
        Case 8  '-- 256 colors / 256 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
        Case 24 '-- True color
            lIdxNew = 8
        Case Else
            Exit Sub
    End Select
End Sub

Private Function pvGetExt(ByVal sFileName As String) As String
    pvGetExt = Mid(sFileName, Len(sFileName) - 2)
End Function

