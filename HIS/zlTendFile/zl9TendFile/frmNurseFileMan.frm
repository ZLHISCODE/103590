VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNurseFileMan 
   Caption         =   "�����ļ�����"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10335
   Icon            =   "frmNurseFileMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10335
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   5025
      Left            =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   660
      Width           =   8400
      _Version        =   589884
      _ExtentX        =   14817
      _ExtentY        =   8864
      _StockProps     =   0
      BorderStyle     =   1
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1590
      Top             =   0
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
            Picture         =   "frmNurseFileMan.frx":5B62
            Key             =   "���µ�"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNurseFileMan.frx":6274
            Key             =   "��¼��"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic�鵵 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1080
      Picture         =   "frmNurseFileMan.frx":6986
      ScaleHeight     =   345
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   90
      Width           =   375
   End
   Begin VB.PictureBox pic���� 
      Appearance      =   0  'Flat
      BackColor       =   &H00EAFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2340
      ScaleHeight     =   225
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   150
      Width           =   1365
      Begin VB.ComboBox cbo���� 
         BackColor       =   &H00EAFFFF&
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   -30
         Width           =   1425
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmNurseFileMan.frx":7088
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15319
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   540
      Left            =   8460
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1260
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
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   510
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNurseFileMan.frx":791A
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   60
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmNurseFileMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrSQL As String
Private mstrPrivs As String
Private mblnSaved As Boolean            '���뱾ģ����Ƿ񱣴������
Private mblnAuto As Boolean
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintӤ�� As Integer
Private mlng����ID As Long
Private mstr���� As String
Private mlngFormat As Long
Private mblnPigeonhole As Boolean       '�鵵
Private mblnFileEnd As Boolean          '�ļ�����
Private mblnPrintMerge As Boolean       '�ϲ���ӡ
Private mintNORule As Integer           '�����ļ�ҳ�����:סԺ�ڼ�ͳһ���ʱ�����������ļ�Ϊ�ϲ���ӡ
Private mblnBlowup As Boolean
Private mblnPrint As Boolean            '��¼���ļ��Ƿ��Ѿ���ӡ
Private mblnSheetFile As Boolean        '�Ƿ���ڼ�¼���ļ�
Private Enum COLDef
    c_ͼ��
    c_�ļ�����
    c_�ļ���Դ
    c_��ʼʱ��
    c_����ʱ��
    c_�����¼��
    c_�����¼��ID
    c_������
    c_����ʱ��
    c_����
End Enum

'�󶨿�ݼ�ʱ,IDֵ������޷������͵�ȡֵ��Χ���޷���,Ҳ����0-65535
Private Const conMenu_Add As Long = 32761
Private Const conMenu_Modify As Long = 32762
Private Const conMenu_Delete As Long = 32763
Private Const conMenu_FileEnd As Long = 32764
Private Const conMenu_FileRestore As Long = 32765
Private Const conMenu_PrintMerge As Long = 32766
Private Const conMenu_PrintSingle As Long = 32767
Private Const conMenu_EndTime As Long = 32768
Private Const conMenu_FormatChange As Long = 32769 '��ʽ���
Private Const conMenu_RetryPage As Long = 32770 '��¼��ҳ������

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
          (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Function ShowEditor(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intBaby As Integer, ByVal strPrivs As String, _
    Optional ByVal blnAuto As Boolean = False, Optional ByVal bytSize As Byte = 0, Optional lngFormat As Long = 0, Optional lng��� As Long = 0) As Boolean
    'blnAuto:�ֹ�����ļ�����Ϊ��,�����鷢�����ļ�ʱ�Զ����������ļ�ʱ��Ϊ��
    On Error Resume Next
    
    mblnAuto = blnAuto
    mlngFormat = 0
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mintӤ�� = intBaby
    mstrPrivs = strPrivs
    mblnSaved = False
    mblnBlowup = (bytSize = 1)
    mblnPrint = False
    mblnSheetFile = False
    
    Me.Show 1
    If mblnSaved = True Then
        lng��� = mintӤ��
        lngFormat = mlngFormat
    End If
    ShowEditor = mblnSaved
End Function

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С(����ģ���Ѿ����ص���)
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Control
    Dim bytFontSize As Byte
    
    bytFontSize = IIf(mblnBlowup = True, 12, 9)
    
    Me.FontSize = bytFontSize
    
    'ReportControl
    Set CtlFont = rptList.PaintManager.CaptionFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set rptList.PaintManager.CaptionFont = CtlFont
    
    Set CtlFont = rptList.PaintManager.TextFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set rptList.PaintManager.TextFont = CtlFont
    'CommandBars
    Set CtlFont = cbsMain.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsMain.Options.Font = CtlFont
End Sub

Private Sub CheckForm()
'���ܣ���鴰���Ƿ��Ѿ�����
    Dim lngAns As Long
    On Error Resume Next
    lngAns = FindWindow(vbNullString, "�����ļ�����")
    lngAns = IsWindow(lngAns)
    If lngAns <> 0 Then Unload Me
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objTool As CommandBar
    Dim objMenu As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    
    'cbsMain
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
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
    cbsMain.Icons = imgPublic.Icons
    
    '�˵���
    cbsMain.ActiveMenuBar.Title = "�˵���"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Add, "����(&A)"): objControl.IconId = 1
        Set objControl = .Add(xtpControlButton, conMenu_Modify, "�޸�(&M)"): objControl.IconId = 2
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "ɾ��(&D)"): objControl.IconId = 3
        Set objControl = .Add(xtpControlButton, conMenu_FileEnd, "��ǽ���(&E)"): objControl.IconId = 4: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_FileRestore, "������¼(&C)"): objControl.IconId = 5
        Set objControl = .Add(xtpControlButton, conMenu_PrintMerge, "�ϲ���ӡ(&G)"): objControl.IconId = 6: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_PrintSingle, "������ӡ(&L)"): objControl.IconId = 7
        Set objControl = .Add(xtpControlButton, conMenu_FormatChange, "��ʽ�滻(&F)"): objControl.IconId = 9: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_RetryPage, "ҳ������(&R)"): objControl.IconId = 9023: objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        objControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): objControl.BeginGroup = True
    End With
    '���ӹ鵵��־
    Set objCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Option, "�鵵")
    objCustom.Handle = Me.pic�鵵.hWnd
    objCustom.Flags = xtpFlagRightAlign
    cbsMain(1).EnableDocking xtpFlagHideWrap + xtpFlagStretched

    '����������
    '-----------------------------------------------------
    Set objTool = cbsMain.Add("������", xtpBarTop)      '����
    objTool.EnableDocking xtpFlagStretched
    objTool.Closeable = False
    With objTool.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Add, "����"): objControl.IconId = 1: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Modify, "�޸�"): objControl.IconId = 2: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Delete, "ɾ��"): objControl.IconId = 3: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_FileEnd, "����"): objControl.IconId = 4: objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True: objControl.ToolTipText = "��ǵ�ǰ�ļ�����"
        Set objControl = .Add(xtpControlButton, conMenu_FileRestore, "ȡ��"): objControl.IconId = 5: objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ȡ����ǰ�ļ��Ľ�����־"
        Set objControl = .Add(xtpControlButton, conMenu_EndTime, "ʱ��"): objControl.IconId = 8: objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "�趨�ļ��Ľ���ʱ��"
        Set objControl = .Add(xtpControlButton, conMenu_PrintMerge, "�ϲ�"): objControl.IconId = 6: objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True: objControl.ToolTipText = "ָ����ʽ��ͬ�������ļ�Ϊ�ϲ���ӡ"
        Set objControl = .Add(xtpControlButton, conMenu_PrintSingle, "����"): objControl.IconId = 7: objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "��ǰ�ļ��趨Ϊ������ӡ"
        Set objControl = .Add(xtpControlButton, conMenu_FormatChange, "�滻"): objControl.IconId = 9: objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "��ǰ�ļ���ʽ�滻Ϊ�µĸ�ʽ": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_RetryPage, "����"): objControl.IconId = 9023: objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "�����������м�¼���ļ���ҳ��": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.Style = xtpButtonIconAndCaption: objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): objControl.Style = xtpButtonIconAndCaption
    End With
    '���⴦��
    '-----------------------------------------------------
    '�������Ҳಡ��������ѡ��
    cbo����.FontSize = BlowUp(9)
    cbo����.Width = BlowUp(1425)
    pic����.Height = cbo����.Height - 45
    pic����.Width = cbo����.Left + cbo����.Width
    With objTool.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Find, "����")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "����")
        objCustom.Handle = Me.pic����.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Add               '���ӻ����ļ�
        .Add 0, vbKeyDelete, conMenu_Delete              'ɾ�������ļ�
        .Add 0, vbKeyF1, conMenu_Help_Help               '����
        .Add 0, vbKeyF5, conMenu_View_Refresh
    End With
End Sub

Private Sub cbo����_Click()
    On Error GoTo ErrHand
    Dim objItem As ReportRecordItem
    Dim objRecord As ReportRecord
    Dim objRpt As ReportControl
    Dim rsTemp As New ADODB.Recordset
    
    mintӤ�� = Val(cbo����.ItemData(cbo����.ListIndex))
    '��ʾָ�����˵Ļ����ļ��б�
    mstrSQL = " Select A.ID,A.�ļ�����, B.���� AS �ļ���Դ,B.����,A.��ʼʱ��,A.����ʱ��,A.������,A.����ʱ��,A.�鵵��,C.�ļ����� AS �����ļ�,C.ID AS �����ļ�ID,B.���� " & _
              " From ���˻����ļ� A,�����ļ��б� B,���˻����ļ� C" & _
              " Where A.��ʽID=B.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3] And A.����ID=C.ID(+)" & _
              " Order by B.����,A.��ʼʱ��"
    Call SQLDIY(mstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ʾָ�����˵Ļ����ļ��б�", mlng����ID, mlng��ҳID, Val(cbo����.ItemData(cbo����.ListIndex)))
    
    mblnPigeonhole = False
    mblnSheetFile = False
    rptList.Records.DeleteAll
    With rsTemp
        If .RecordCount <> 0 Then
            mblnPigeonhole = (NVL(!�鵵��) <> "")
        End If

        '����¼���뱨��ؼ�
        Do While Not .EOF
            Set objRecord = Me.rptList.Records.Add()
            objRecord.Tag = CStr(!ID)
            Set objItem = objRecord.AddItem(""): objItem.Icon = IIf(Val(NVL(!����)) = 1, -1, Val(NVL(!����))) + 1
            Set objItem = objRecord.AddItem(CStr(!�ļ�����))
            objItem.Caption = CStr(!�ļ�����)
            Set objItem = objRecord.AddItem(CStr(!�ļ���Դ))
            objItem.Caption = CStr(!�ļ���Դ)
            Set objItem = objRecord.AddItem(CStr(Format(!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")))
            objItem.Caption = CStr(Format(!��ʼʱ��, "yyyy-MM-dd HH:mm:ss"))
            Set objItem = objRecord.AddItem(CStr(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss")))
            objItem.Caption = CStr(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
            Set objItem = objRecord.AddItem(CStr(NVL(!�����ļ�)))
            objItem.Caption = CStr(NVL(!�����ļ�))
            Set objItem = objRecord.AddItem(Val(NVL(!�����ļ�ID)))
            objItem.Caption = CStr(NVL(!�����ļ�ID))
            Set objItem = objRecord.AddItem(CStr(!������))
            objItem.Caption = CStr(!������)
            Set objItem = objRecord.AddItem(CStr(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss")))
            objItem.Caption = CStr(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"))
            Set objItem = objRecord.AddItem(NVL(!����, 0))
            objItem.Caption = NVL(!����, 0)
            If mblnSheetFile = False And InStr(1, ",-1,1," & "," & Val(NVL(!����)) & ",") = 0 Then
                mblnSheetFile = True
            End If
            .MoveNext
        Loop
    End With
    rptList.Populate
    
    Exit Sub
ErrHand:
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
    
    objPrint.Title.Text = "�����ļ��嵥"
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

Private Sub LoadPati()
    Dim strName As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '��ȡ���˵�ǰ����
    mstrSQL = " Select B.ID,B.����" & _
              " From ������ҳ A,���ű� B" & _
              " Where A.����ID=[1] And A.��ҳID=[2] And A.��Ժ����ID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���˵�ǰ����", mlng����ID, mlng��ҳID)
    mlng����ID = rsTemp!ID
    mstr���� = rsTemp!����
    
    '��ȡ�������
    mstrSQL = "" & _
            "SELECT A.����ID,0 AS ���,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա� FROM ������Ϣ A,������ҳ B WHERE A.����ID=B.����ID And B.����ID=[1] AND B.��ҳID=[2] " & vbNewLine & _
            "UNION" & vbNewLine & _
            "SELECT ����ID,���,Ӥ������ AS ����,Ӥ���Ա� AS �Ա� FROM ������������¼ WHERE ����ID=[1] AND ��ҳID=[2]" & vbNewLine & _
            "ORDER BY ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ�������", mlng����ID, mlng��ҳID)
    
    With rsTemp
        cbo����.Clear
        Do While Not .EOF
            If !��� = 0 Then strName = !����
            cbo����.AddItem IIf(IsNull(!����), strName & "֮��" & !���, !����)
            cbo����.ItemData(cbo����.NewIndex) = !���
            If mintӤ�� = !��� Then cbo����.ListIndex = .AbsolutePosition - 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRpt()
    Dim objCol As ReportColumn
    With rptList
        .Columns.DeleteAll
        Set objCol = .Columns.Add(c_ͼ��, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_�ļ�����, "��¼������", BlowUp(120), True)
        Set objCol = .Columns.Add(c_�ļ���Դ, "�ļ���Դ", BlowUp(120), True)
        Set objCol = .Columns.Add(c_��ʼʱ��, "��ʼʱ��", BlowUp(130), True)
        Set objCol = .Columns.Add(c_����ʱ��, "����ʱ��", BlowUp(130), True)
        Set objCol = .Columns.Add(c_�����¼��, "�����¼��", BlowUp(120), True)
        Set objCol = .Columns.Add(c_�����¼��ID, "�����¼��ID", 0, True): objCol.Visible = False
        Set objCol = .Columns.Add(c_������, "������", BlowUp(70), True)
        Set objCol = .Columns.Add(c_����ʱ��, "����ʱ��", BlowUp(130), True)
        Set objCol = .Columns.Add(c_����, "����", 0, False): objCol.Visible = False
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Sortable = True
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            '.HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û���ļ�..."
        End With
        .TabStop = True
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgList
    End With
End Sub

'�ؼ��¼�##############################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngFormat As Long
    Dim cbrControl As Object
    Dim DBeginTime As Date
    Dim lngFileID As Long
    Dim blnTrans As Boolean
    Dim ArrSQL()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    blnTrans = False
    '����״̬�����
    Select Case Control.ID
        Case conMenu_Add
            If Val(Split(EprIsCommit(mlng����ID, mlng��ҳID), "|")(0)) = 0 Then
                MsgBox "�ò��˵Ĳ������ύ���[״̬��" & gstrMecState & "]����������ļ�����ȡ���������ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
        Case conMenu_Delete
            If Val(Split(EprIsCommit(mlng����ID, mlng��ҳID), "|")(1)) = 0 Then
                MsgBox "�ò��˵Ĳ������ύ���[״̬��" & gstrMecState & "]������ɾ���ļ�����ȡ���������ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
        Case conMenu_Modify, conMenu_FormatChange, conMenu_RetryPage, conMenu_FileEnd, conMenu_FileRestore, _
            conMenu_EndTime, conMenu_PrintMerge, conMenu_PrintSingle
            If Val(Split(EprIsCommit(mlng����ID, mlng��ҳID), "|")(2)) = 0 Then
                MsgBox "�ò��˵Ĳ������ύ���[״̬��" & gstrMecState & "]�������޸��ļ�����ȡ���������ԣ�", vbInformation, gstrSysName
                Exit Sub
            End If
    End Select
    Select Case Control.ID
        Case conMenu_File_PrintSet

            Call zl9PrintMode.zlPrintSet

        Case conMenu_File_Preview

            Call zlRptPrint(2)

        Case conMenu_File_Print

            Call zlRptPrint(1)

        Case conMenu_File_Excel

            Call zlRptPrint(3)

        Case conMenu_Add
            '���ݸ�ʽID��λ�ļ�
            mlngFormat = 0
            If frmNurseFileEdit.ShowEditor(mlng����ID, mlng��ҳID, Me.cbo����.ItemData(Me.cbo����.ListIndex), mlng����ID, mstr����, mlngFormat, lngFormat) Then
                mblnSaved = True
                Call cbo����_Click
            End If
            mlngFormat = lngFormat
        Case conMenu_Modify
            '���ݸ�ʽID��λ�ļ�
            mlngFormat = Val(rptList.FocusedRow.Record.Tag)
            If frmNurseFileEdit.ShowEditor(mlng����ID, mlng��ҳID, Me.cbo����.ItemData(Me.cbo����.ListIndex), mlng����ID, mstr����, mlngFormat, lngFormat) Then
                mblnSaved = True
                Call cbo����_Click
            End If
            mlngFormat = lngFormat
        Case conMenu_Delete
            If rptList.FocusedRow Is Nothing Then Exit Sub
            
            If Val(rptList.FocusedRow.Record.Item(c_����).Value) = -1 Then
                gstrSQL = "SELECT A.ID,B.��ʼʱ��" & _
                    " FROM ���˻������� A,���˻����ļ� B,���˻�����ϸ C" & _
                    " WHERE B.ID=[1] And A.�ļ�ID=B.ID And A.id=C.��¼id And RowNum<2"
            Else
                gstrSQL = "SELECT A.ID,B.��ʼʱ��" & _
                    " FROM ���˻������� A,���˻����ӡ C,���˻����ļ� B,���˻�����ϸ D" & _
                    " WHERE B.ID=[1] And A.�ļ�ID=B.ID and A.�ļ�ID=C.�ļ�ID And A.id=D.��¼id And A.ID=C.��¼ID And RowNum<2"
            End If
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ��������", Val(rptList.FocusedRow.Record.Tag))
            If rsTemp.RecordCount > 0 Then
                MsgBox "���ļ��Ѿ������������ݲ�����ɾ��,���飡", vbInformation, gstrSysName
                Exit Sub
            End If
            
            If Val(rptList.FocusedRow.Record.Item(c_����).Value) = -1 Then
                DBeginTime = CDate(rptList.FocusedRow.Record(c_��ʼʱ��).Value)
                gstrSQL = " Select A.ID,A.��ʼʱ��" & _
                    " From ���˻����ļ� A,�����ļ��б� B" & _
                    " Where A.��ʽID=B.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3] And B.����=-1 order by A.��ʼʱ�� DESC"
                Call SQLDIY(gstrSQL)
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��Ѷ������µ�", mlng����ID, mlng��ҳID, mintӤ��)
                rsTemp.Filter = "��ʼʱ��> '" & CStr(DBeginTime) & "'"
                If rsTemp.RecordCount > 0 Then
                    MsgBox "���ļ�֮�󻹴������������µ��ļ�,�ļ�ֻ�ܴӺ���ǰɾ��,���飡", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            If MsgBox("��ȷ��Ҫɾ��" & rptList.FocusedRow.Record.Item(c_�ļ�����).Caption & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            'If MsgBox("���ļ����еĻ�������Ҳ��һ��ɾ�������ٴ�ȷ���Ƿ�ɾ����", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            ArrSQL = Array()
            gstrSQL = "ZL_���˻����ļ�_DELETE(" & Val(rptList.FocusedRow.Record.Tag) & ")"
            ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
            ArrSQL(UBound(ArrSQL)) = gstrSQL
            
            If Val(rptList.FocusedRow.Record.Item(c_����).Value) = -1 Then
                rsTemp.Filter = "��ʼʱ��< '" & CStr(DBeginTime) & "'"
                rsTemp.Sort = "��ʼʱ�� DESC"
                If rsTemp.RecordCount > 0 Then
                    'ȡ����һ���µ��ļ��Ľ���ʱ��
                    gstrSQL = "ZL_���˻����ļ�_STATE(" & Val(rsTemp!ID) & ",1,NULL)"
                    ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
                    ArrSQL(UBound(ArrSQL)) = gstrSQL
                End If
            End If
            
            'ɾ�������¼��ʱ��������ļ�ҳ��˳������Ҫ������ļ�֮����ļ�ҳ��
            '�˴��������ļ����ںϲ��������(��Ϊɾ���Ѿ����ƣ������ļ�������ںϲ���Ϣ����ɾ��)
            If InStr(1, ",-1,1,", "," & Val(rptList.FocusedRow.Record.Item(c_����).Value) & ",") = 0 And mintNORule <> 0 Then
                
                gstrSQL = " Select id " & vbNewLine & _
                    " From (" & vbNewLine & _
                    "   With ���˻����ļ�_F1 As" & vbNewLine & _
                    "   (Select a.Id, a.����id, ��ʼʱ��, ����ʱ��" & vbNewLine & _
                    "   From ���˻����ļ� a, �����ļ��б� b" & vbNewLine & _
                    "   Where a.��ʽid = b.Id And b.���� = 3 And b.���� <> 1 And b.���� <> -1 And a.����id = [1] And a.��ҳid = [2] And Nvl(a.Ӥ��, 0) = [3])" & vbNewLine & _
                    "   Select Id" & vbNewLine & _
                    "   From (Select Id, ��ʼʱ��, ����ʱ��" & vbNewLine & _
                    "       From ���˻����ļ�_F1 a" & vbNewLine & _
                    "       Where Not Exists (Select 1 From ���˻����ļ�_F1 Where a.Id = ����id))" & vbNewLine & _
                    "   Where id<>[4] And (��ʼʱ��>[5] OR (��ʼʱ��=[5] And ����ʱ��>[6])) " & vbNewLine & _
                    "   Order by ��ʼʱ��)"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ļ�֮��Ļ����ļ�", mlng����ID, mlng��ҳID, Val(cbo����.ItemData(cbo����.ListIndex)), Val(rptList.FocusedRow.Record.Tag), _
                    CDate(Format(rptList.FocusedRow.Record.Item(c_��ʼʱ��).Value, "YYYY-MM-DD HH:mm:ss")), CDate(Format(rptList.FocusedRow.Record.Item(c_����ʱ��).Value, "YYYY-MM-DD HH:mm:ss")))
                If rsTemp.RecordCount > 0 Then
                    gstrSQL = "Zl_���˻����ӡ_Batchretrypage(" & rsTemp!ID & ",'1;0')"
                    ReDim Preserve ArrSQL(UBound(ArrSQL) + 1)
                    ArrSQL(UBound(ArrSQL)) = gstrSQL
                End If
            End If
            
            If UBound(ArrSQL) > 0 Then gcnOracle.BeginTrans: blnTrans = True
            For lngLoop = 0 To UBound(ArrSQL)
                If CStr(ArrSQL(lngLoop)) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(ArrSQL(lngLoop)), "�ļ�ɾ��")
            Next
            If UBound(ArrSQL) > 0 Then gcnOracle.CommitTrans: blnTrans = False
            mblnSaved = True
            Call cbo����_Click
        Case conMenu_FormatChange '�ļ���ʽ���
            lngFormat = Val(rptList.FocusedRow.Record.Tag)
            If frmNurseFileChange.ShowEditor(Me, Val(rptList.FocusedRow.Record.Tag)) Then
                mblnSaved = True
                Call cbo����_Click
                mlngFormat = lngFormat
            End If
        Case conMenu_RetryPage '��¼��ҳ������
            lngFileID = 0
            If MsgBox("ҳ������������ݲ���:�����ļ�ҳ�����,�Ե�ǰ�������еļ�¼���ļ�ҳ���������," & _
                "���Ҵ˲������������¼���ļ��Ĵ�ӡ��Ϣ��" & vbCrLf & "�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                For lngLoop = 0 To rptList.Records.Count - 1
                    If InStr(1, ",-1,1,", "," & Val(rptList.Records(lngLoop).Item(c_����).Value) & ",") = 0 Then
                        lngFileID = Val(rptList.Records(lngLoop).Tag)
                        Exit For
                    End If
                Next lngLoop
                If lngFileID = 0 Then Exit Sub
                Screen.MousePointer = 11
                gstrSQL = "Zl_���˻����ӡ_Batchretrypage(" & lngFileID & ",NULL,1)"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "ҳ������")
                mblnSaved = True
                Screen.MousePointer = 0
                MsgBox "ҳ��������ɣ�", vbInformation, gstrSysName
            End If
        Case conMenu_FileEnd
            gstrSQL = "ZL_���˻����ļ�_STATE(" & Val(rptList.FocusedRow.Record.Tag) & ",1,sysdate)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "����ļ�����")
            Call cbo����_Click
        Case conMenu_FileRestore
            gstrSQL = "ZL_���˻����ļ�_STATE(" & Val(rptList.FocusedRow.Record.Tag) & ",1,NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "ȡ���ļ�����")
            Call cbo����_Click
        Case conMenu_EndTime
            If frmNurseFileEndTime.ShowEditor(Val(rptList.FocusedRow.Record.Tag)) Then cbo����_Click
        Case conMenu_PrintMerge
            If frmNurseFileMerge.ShowEditor(Val(rptList.FocusedRow.Record.Tag)) Then
                mblnSaved = True
                cbo����_Click
            End If
        Case conMenu_PrintSingle
            gcnOracle.BeginTrans
            blnTrans = True
            gstrSQL = "ZL_���˻����ļ�_STATE(" & Val(rptList.FocusedRow.Record.Tag) & ",2,NULL)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "ȡ���ϲ���ӡ")
            gstrSQL = "Zl_���˻����ӡ_Batchretrypage(" & Val(rptList.FocusedRow.Record.Item(c_�����¼��ID).Value) & ",'1;1')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "ҳ������")
            gcnOracle.CommitTrans
            blnTrans = False
            mblnSaved = True
            Call cbo����_Click
        Case conMenu_View_ToolBar_Button

            cbsMain(2).Visible = Not cbsMain(2).Visible
            cbsMain.RecalcLayout

        Case conMenu_View_ToolBar_Text

            For Each cbrControl In cbsMain(2).Controls
                If cbrControl.Type = xtpControlButton Then
                    cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next

            cbsMain.RecalcLayout

        Case conMenu_View_StatusBar

            stbThis.Visible = Not stbThis.Visible
            cbsMain.RecalcLayout

        Case conMenu_View_Refresh
            Call cbo����_Click

        Case conMenu_Help_Help

            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))

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
            Exit Sub
    End Select

    cbsMain.RecalcLayout

    Exit Sub

ErrHand:
    Screen.MousePointer = 0
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Resize()
    Call Form_Resize
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (rptList.Records.Count > 0)
    Case conMenu_Add
        Control.Enabled = Not mblnPigeonhole
    Case conMenu_Modify, conMenu_Delete
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                'ֻ���޸ġ�ɾ���Լ��������ļ�
'                Control.Enabled = (rptList.FocusedRow.Record.Item(c_������).Value = gstrUserName) And Not mblnPigeonhole
                Control.Enabled = Not mblnPigeonhole
                If Control.ID = conMenu_Delete Then
                    Control.Visible = InStr(1, mstrPrivs, "�����ļ�ɾ��") <> 0
                    Control.Enabled = Control.Enabled And Not ISPrintMerge And Control.Visible
                End If
            End If
        End If
    Case conMenu_FileEnd
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnFileEnd And Not mblnPigeonhole And (Val(rptList.FocusedRow.Record.Item(c_����).Value) >= 0)
            End If
        End If
    Case conMenu_FormatChange '�ļ���ʽ���
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not ISPrintMerge And Not mblnPrint And Not mblnPigeonhole And (Val(rptList.FocusedRow.Record.Item(c_����).Value) <> 1)
            End If
        End If
    Case conMenu_RetryPage 'ҳ������(δ�鵵�����Ҵ��ڼ�¼���ļ�)
        Control.Enabled = Not mblnPigeonhole And mblnSheetFile
    Case conMenu_FileRestore, conMenu_EndTime
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = mblnFileEnd And Not mblnPigeonhole And (Val(rptList.FocusedRow.Record.Item(c_����).Value) >= 0)
            End If
        End If
    Case conMenu_PrintMerge
        Control.Enabled = False
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnPigeonhole And Not mblnPrintMerge And (rptList.FocusedRow.Record.Item(c_ͼ��).Icon > 0)
            End If
        End If
    Case conMenu_PrintSingle
        If rptList.FocusedRow Is Nothing Then
            Control.Enabled = False
        Else
            If rptList.FocusedRow.Record Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Not mblnPigeonhole And mblnPrintMerge And (rptList.FocusedRow.Record.Item(c_ͼ��).Icon > 0)
            End If
        End If
    Case conMenu_View_Option    '�鵵��־
        Control.Visible = mblnPigeonhole
    Case conMenu_View_ToolBar_Button
        Control.Checked = Me.cbsMain(2).Visible
    Case conMenu_View_ToolBar_Text
        Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar
        Control.Checked = Me.stbThis.Visible
    End Select
End Sub


Private Sub Form_Load()
    Dim blnSel As Boolean
    mintNORule = zlDatabase.GetPara("�����ļ�ҳ�����", glngSys, 1255, 0)
    Call RestoreWinState(Me, App.ProductName)
    Call MainDefCommandBar
    Call InitRpt
    Call LoadPati
    
    If rptList.Records.Count = 0 Then
        '���û���ļ��򵯳��Զ������ļ�
        Dim Control As XtremeCommandBars.ICommandBarControl
        Dim objControl As CommandBarControl, objParent As CommandBarPopup
        
        Set objParent = cbsMain.Item(1).Controls.Item(2)        'ҽ��ҵ��
        For Each objControl In objParent.CommandBar.Controls
            If objControl.ID = conMenu_Add Then blnSel = True: Exit For
        Next
        If blnSel Then objControl.Execute
        
        If mblnAuto Then
            Unload Me
            Exit Sub
        End If
    End If
    
    Call ReSetFontSize
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With rptList
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight
        .Height = lngBottom - lngTop - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If rptList.Records.Count = 0 Then Exit Sub
    If rptList.FocusedRow Is Nothing Then Exit Sub
    If mblnPigeonhole Then Exit Sub
    
    Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Modify))
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptList_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub rptList_SelectionChanged()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mblnPrint = False
    If rptList.Records.Count = 0 Then Exit Sub
    If rptList.FocusedRow Is Nothing Then Exit Sub
    
    mblnFileEnd = (rptList.FocusedRow.Record.Item(c_����ʱ��).Caption <> "")
    mblnPrintMerge = (rptList.FocusedRow.Record.Item(c_�����¼��).Caption <> "")
    

    If InStr(1, ",-1,1,", "," & Val(rptList.FocusedRow.Record.Item(c_����).Value) & ",") = 0 Then
        gstrSQL = "SELECT �ļ�ID from ���˻����ӡ where �ļ�ID=[1] And ��ӡҳ�� Is Not NULL And RowNum<2"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ��ӡ", Val(rptList.FocusedRow.Record.Tag))
        mblnPrint = (rsTemp.RecordCount > 0)
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ISPrintMerge() As Boolean
'���ܣ�����ļ��Ƿ�������ļ��ϲ���ӡ
    Dim lngRow As Long
    
    If rptList.Records.Count = 0 Then Exit Function
    If rptList.FocusedRow Is Nothing Then Exit Function
    
    If mblnPrintMerge = True Then
        ISPrintMerge = True
        Exit Function
    Else
        With rptList
            For lngRow = 0 To .Records.Count - 1
                If Val(.Records(lngRow).Item(c_�����¼��ID).Value) = Val(.FocusedRow.Record.Tag) And _
                    .Records(lngRow).Index <> .FocusedRow.Index Then
                    ISPrintMerge = True
                    Exit Function
                End If
            Next lngRow
        End With
    End If
    ISPrintMerge = False
End Function
