VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmImportFile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "������Ŀ"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   ForeColor       =   &H00FF0000&
   Icon            =   "frmImportFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10425
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   5145
      TabIndex        =   6
      Top             =   900
      Width           =   5175
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   2565
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4905
         _cx             =   8652
         _cy             =   4524
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
         BackColorSel    =   4227072
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmImportFile.frx":6852
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
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "��"
      Height          =   300
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   280
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Left            =   300
      MousePointer    =   7  'Size N S
      ScaleHeight     =   354.167
      ScaleMode       =   0  'User
      ScaleWidth      =   4215
      TabIndex        =   0
      Top             =   3990
      Width           =   4215
      Begin VB.Label lblCollect 
         BackColor       =   &H80000005&
         Caption         =   "�е�������ʾ"
         Height          =   180
         Left            =   45
         TabIndex        =   8
         Top             =   45
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList imgError 
      Left            =   5145
      Top             =   75
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
            Picture         =   "frmImportFile.frx":68C7
            Key             =   "error"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportFile.frx":D129
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   6090
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   765
      Left            =   75
      TabIndex        =   3
      Top             =   4425
      Width           =   5055
      _cx             =   8916
      _cy             =   1349
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
      BackColorSel    =   8454016
      ForeColorSel    =   16744576
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImportFile.frx":1398B
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
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   750
      Left            =   7305
      TabIndex        =   5
      Top             =   1110
      Width           =   1500
      _Version        =   589884
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   64
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��  ��(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   810
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgPicture 
      Bindings        =   "frmImportFile.frx":13A00
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmImportFile.frx":13A14
   End
End
Attribute VB_Name = "frmImportFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbyt���뷽ʽ, mbyt�������, mbyt�ϼ�����, mbyt����, mbyt����, mbytNumAndKind, mbytNameAndKind As Byte
Private mbytҩƷ���, mbytҩƷ����, mbytƷ�ֱ���, mbytƷ������, mbyt������, mbytҩƷ���, mbyt����, mbyt����, mbyt��λ, mbyt����ϵ��, mbyt���, mbyt�۸�, mbytЧ��, mbyt������Ŀ, mbyt����, mbyt�������, mbyt��������, mbyt��Ӧ��, mbyt����, mbytƷ��Ψһ, mbyt���Ψһ As Byte
'�����(����,0��ѡ1��ѡ,0��ʾ1����|...)
Private Const MSTRMEDICAL      As String = "���,0,0|�ϼ�����,0,0|����,0,0|����,0,0||���,0,0|����,0,0|Ʒ�ֱ���,0,0|Ʒ������,0,0|������,0,0|ҩƷ���,0,0|������,0,0|����,0,0|������λ,0,0|�ۼ۵�λ,0,0|�ۼۻ���ϵ��,0,0|���ﵥλ,0,0|���ﻻ��ϵ��,0,0|סԺ��λ,0,0|סԺ����ϵ��,0,0|ҩ�ⵥλ,0,0|ҩ�⻻��ϵ��,0,0|" & _
                                        "�Ƿ���,0,0|�ɱ���,0,0|�ۼ�,0,0|������Ŀ,0,0|סԺ�ɷ����,0,0|����ɷ����,0,0|�������,1,0|ҩ�����,1,0|ҩ������,1,0|Ч��(��),1,0|��Ӧ������,1,0|��Ӧ�����֤��,1,0|��Ӧ�����֤Ч��,1,0|"
'�е�������ʾ(����;��ʾ|...)
Private Const MSTRCOMMENT      As String = "���;���ֻ��������ҩ���г�ҩ���в�ҩ������Ϊ��|�ϼ�����;�ϼ�������ձ������е����ݣ����Ʒ���Ŀ¼.���ƣ���ʽ����\�ָ��������� ����һ������\������Ϊ�ձ�ʾû���ϼ�|����;���벻��Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���|����;���Ʋ��ܺ��зǷ��ַ����磺�����ţ�����Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���||" & _
                                        "���;���ֻ��������ҩ���г�ҩ���в�ҩ������Ϊ��|����;������ձ������е����ݣ����Ʒ���Ŀ¼.���� ��ʽ����\�ָ��������� ����һ������\����������Ϊ��|Ʒ�ֱ���;Ʒ�ֱ��벻��Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���|Ʒ������;Ʒ�����Ʋ��ܺ��зǷ��ַ����磺�����ţ�����Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���|" & _
                                        "������;�����벻��Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���|ҩƷ���;ҩƷ����ܺ��зǷ��ַ����磺�����ţ�����Ϊ�գ����Ȳ��ܳ������ݿ��ֶγ���|������;�������ֶγ��Ȳ��ܳ������ݿ��ֶ���Ƴ��ȣ����ܺ��зǷ��ַ����磺�����ţ�����Ϊ��|����;���Ͳ������ݿ�����������ݣ�����Ϊ�գ����ܺ��зǷ��ַ������Ȳ��ܳ������ݿ��ֶ���Ƴ���|" & _
                                        "������λ;��λ����Ϊ�գ����Ȳ��ܳ������ݿ��ֶ���Ƴ���|�ۼ۵�λ;|�ۼۻ���ϵ��;����ϵ������Ϊ�գ���λ����ϵ�������Ҷ�>0����λ��ͬ����ϵ��������ͬ�Ҷ�������|���ﵥλ;|���ﻻ��ϵ��;|סԺ��λ;|סԺ����ϵ��;|ҩ�ⵥλ;|ҩ�⻻��ϵ��;|�Ƿ���;Ϊ��Ĭ��Ϊ���ۣ����̡���ʾʱ��|�ɱ���;�۸��ֶ�ֻ���������ͣ����Ȳ��ܳ���������þ���|" & _
                                        "�ۼ�;|������Ŀ;������Ŀ����Ϊ�գ�ֻ�������ݿ�����������Ŀ|סԺ�ɷ����;���㷽ʽ��0-���Է���,1-���ɷ���,2-һ����ʹ��,3-�����һ������Ч,4-�������������Ч,5-�������������Ч|����ɷ����;|�������;�������0�Ϳ�-�������ڲ��ˣ�1-���2-סԺ��3-�����סԺ|" & _
                                        "ҩ�����;Ϊ�ձ�ʾ�����������̡���ʾ����|ҩ������;ҩ�����ʱҩ�����ܷ���|Ч��(��);Ч�ڱ�����������ֻ���ǲ�С��0������|��Ӧ������;��Ӧ�̲������ݿ�����������ݣ�����Ϊ��|��Ӧ�����֤��;|��Ӧ�����֤Ч��;¼�����ڵı���������ڸ�ʽ��2015-10-10����2015/10/10����2015.10.10|"

'ҩƷ��ϸ������(����1|ѡ��1,ѡ��2,ѡ��3;����2|ѡ��1,ѡ��2,ѡ��3)
Private Const mstrҩƷ��ϸ As String = "���|����ҩ,�г�ҩ,�в�ҩ;�Ƿ���|��;סԺ�ɷ����|0-���Է���,1-���ɷ���,2-һ����ʹ��,3-�����һ������Ч,4-�������������Ч,5-�������������Ч;" & _
                                                            "����ɷ����|0-���Է���,1-���ɷ���,2-һ����ʹ��,3-�����һ������Ч,4-�������������Ч,5-�������������Ч;�������|0-�������ڲ���,1-����,2-סԺ,3-�����סԺ;" & _
                                                            "ҩ�����|��;ҩ������|��"
Private Const MCONTOOLMODE     As Integer = 100  'Excel����
Private Const MCONTOOLOUTPUT   As Integer = 101  '����Excel
Private Const MCONTOOLCHECK    As Integer = 102  'У��
Private Const MCONTOOLSAVE     As Integer = 103  '����
Private Const MCONTOOLEXIT     As Integer = 104  '�˳�
Private Const MCONTOOLCHECKSET As Integer = 107  '�������
Private Const MCONTOOLCOLSET   As Integer = 109  '������
Private mstrType               As String         '��ʾ�ķ�����ͷ
Private mstrMedi               As String         '��ʾ����ϸ��ͷ
Private mstrTypeMsg            As String         '��������������Ϣ
Private mstrMediMsg            As String         '��ϸ�����������Ϣ
Private mintType               As Integer        '�����ļ����� 1-���ã�2-ҩƷ��3-����
Private mlngModule             As Long           'ģ���
Private mobjXLS As Object
Private mobjWB As Object
Private mobjWS As Object


Private Sub InitComandbar()
    '��ʼ��������
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
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
    Me.cbsMain.Icons = Me.imgPicture.Icons
    
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
        
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLMODE, "����Excel����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLOUTPUT, "����Excel")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECKSET, "�������")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCOLSET, "������")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECK, "У��")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLSAVE, "����")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLEXIT, "�˳�")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
    End With
    cbsMain.Item(1).Delete
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case MCONTOOLMODE     'Excel����
            Call ProduceStyleBook
        Case MCONTOOLOUTPUT   '����Excel
            Call OutPutFile
        Case MCONTOOLCHECKSET '��������
            frmImportFileCondition.ShowMe Me, mlngModule
            If vsfList.Rows > 1 Then
                If TabControl.Selected.Caption = "����" Then Call CheckKind
                If TabControl.Selected.Caption = "��ϸ" Then Call CheckƷ��: Call Check���
            End If
        Case MCONTOOLCOLSET   '������
            Call SetCols
        Case MCONTOOLCHECK    'У��
            Call FS.ShowFlash("����У������,���Ժ� ...", Me)
            Me.MousePointer = vbHourglass
            If TabControl.Selected.Caption = "����" Then
                Call CheckKind
                Call GetColumns("����")
                Call CheckKind
                Call GetColumns("����")
            Else
                Call CheckƷ��
                Call Check���
                Call GetColumns("��ϸ")
                Call CheckƷ��
                Call Check���
                Call GetColumns("��ϸ")
            End If
            Me.MousePointer = vbDefault
            Call FS.StopFlash
        Case MCONTOOLSAVE     '����
            Call SaveCard
        Case MCONTOOLEXIT     '�˳�
            Unload Me
    End Select
End Sub

Private Sub OutPutFile()
    '��������ļ�
    Dim strFileName As String
    Dim i As Long
    Dim j As Long
    Dim arrType As Variant
    Dim arrMedi As Variant
    Dim intNum As Integer
    Dim blnFinished As Boolean
    
    On Error GoTo ErrHand
    
    arrType = Split(mstrTypeMsg, "|")
    arrMedi = Split(mstrMediMsg, "|")
    
    If mobjXLS Is Nothing Then Call InitExcel
    mobjXLS.SheetsInNewWorkbook = 1  '���½��Ĺ�����������Ϊ1
    mobjXLS.Workbooks.Add          '����һ��������
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "ҩƷ����"  '�޸Ĺ���������
    mobjXLS.Sheets.Add , mobjXLS.Sheets("ҩƷ����") '���ӵڶ����������ڵ�һ��֮��
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "ҩƷ��ϸ"
    
    mobjXLS.Sheets("ҩƷ����").Select     'ѡ�й�����<ҩƷ����>
    mobjXLS.Columns("A:L").NumberFormatLocal = "@"   '�����ı���ʽ
    'ѭ��д������
    For i = LBound(arrType) To UBound(arrType) - 1
        For j = 0 To UBound(Split(arrType(i), ";")) - 1
            mobjXLS.Cells(i + 1, j + 1) = Split(arrType(i), ";")(j)
            '����Excel��ע
            If i = 0 Then
                For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")) - 1
                    If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(0) = Split(arrType(i), ";")(j) Then
                        If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1) <> "" Then
                            mobjXLS.ActiveSheet.Cells(i + 1, j + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1)
                        End If
                        Exit For
                    End If
                Next
            End If
        Next
    Next
    Call SetExcel("����", j, i)
    
    mobjXLS.Sheets("ҩƷ��ϸ").Select     'ѡ�й�����<ҩƷ��ϸ>
    mobjXLS.Columns("A:AE").NumberFormatLocal = "@"   '�����ı���ʽ
    'ѭ��д������
    For i = LBound(arrMedi) To UBound(arrMedi) - 1
        For j = 0 To UBound(Split(arrMedi(i), ";")) - 1
            mobjXLS.Cells(i + 1, j + 1) = Split(arrMedi(i), ";")(j)
            '����Excel��ע
            If i = 0 Then
                For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(1), "|")) - 1
                    If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(0) = Split(arrMedi(i), ";")(j) Then
                        If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1) <> "" Then
                            mobjXLS.ActiveSheet.Cells(i + 1, j + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1)
                        End If
                        Exit For
                    End If
                Next
            End If
        Next
    Next
    Call SetExcel("��ϸ", j, i)
    mobjXLS.Sheets("ҩƷ����").Select
    
    With dlgOpenFile
        .CancelError = True
        .FileName = ""
        .Filter = "*.xlsx|*.xlsx|*.xls|*.xls"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjXLS.ActiveWorkbook.SaveAs strFileName
            blnFinished = True
        End If
    End With
    
ErrHand:
    mobjXLS.Quit
    If blnFinished Then
        MsgBox "�����ɹ���", vbInformation, gstrSysName
    Else
        MsgBox "����ʧ�ܣ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdFile_Click()
'��ȡ�ļ�����ȡ����
    On Error GoTo ErrHand
    
    dlgOpenFile.FileName = ""
    dlgOpenFile.Filter = "*.xlsx|*.xlsx|*.xls|*.xls"
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        txtFile.Text = dlgOpenFile.FileName
    Else
        GoTo ErrHand
    End If
    
    If txtFile.Text <> "" Then
        DoEvents
        Call FS.ShowFlash("���ڼ�������,���Ժ� ...", Me)
        Me.MousePointer = vbHourglass
        '��ȡ����
        Call GetExcelData
        Me.MousePointer = vbDefault
        Call FS.StopFlash
    End If
    
    Exit Sub
ErrHand:
    Exit Sub
End Sub

Private Sub ParseParameter()
'������������ȡУ�鷽ʽ
    Dim arryPara As Variant
    Dim strPara  As String
    
    '���뷽ʽ/���|�ϼ�����|����|����|��������Ψһ���|���ơ�����ϼ�����Ψһ���|���|����|Ʒ�ֱ���|Ʒ������|������|ҩƷ���|���غϷ��Լ��|����|������λ���|������λ������|��ۼ��|�۸���|Ч��|������Ŀ|����/סԺ����|�������|��������|��Ӧ��|���ڸ�ʽ|Ʒ��Ψһ�Լ��|���Ψһ�Լ��
    '(0������ʾ1�����ֹ/0��ʾ1��ֹ|....)
    strPara = zlDatabase.GetPara("�����ļ���鷽ʽ", glngSys, mlngModule, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    mbyt���뷽ʽ = Mid(strPara, 1, 1)
    strPara = Mid(strPara, 3)
    arryPara = Split(strPara, "|")
    '����
    mbyt������� = arryPara(0)
    mbyt�ϼ����� = arryPara(1)
    mbyt���� = arryPara(2)
    mbyt���� = arryPara(3)
    mbytNumAndKind = arryPara(4)
    mbytNameAndKind = arryPara(5)
    '��ϸ
    mbytҩƷ��� = arryPara(6)
    mbytҩƷ���� = arryPara(7)
    mbytƷ�ֱ��� = arryPara(8)
    mbytƷ������ = arryPara(9)
    mbyt������ = arryPara(10)
    mbytҩƷ��� = arryPara(11)
    mbyt���� = arryPara(12)
    mbyt���� = arryPara(13)
    mbyt��λ = arryPara(14)
    mbyt����ϵ�� = arryPara(15)
    mbyt��� = arryPara(16)
    mbyt�۸� = arryPara(17)
    mbytЧ�� = arryPara(18)
    mbyt������Ŀ = arryPara(19)
    mbyt���� = arryPara(20)
    mbyt������� = arryPara(21)
    mbyt�������� = arryPara(22)
    mbyt��Ӧ�� = arryPara(23)
    mbyt���� = arryPara(24)
    mbytƷ��Ψһ = arryPara(25)
    mbyt���Ψһ = arryPara(26)
End Sub

Private Sub ProduceStyleBook()
'���ɵ����ⲿ�ļ��ı�׼XLS�ļ�����
    Dim arrTypeCols As Variant
    Dim arrMediCols As Variant
    Dim blnFinished As Boolean
    Dim strFileName As String
    Dim strMedi     As String
    Dim intNum      As Integer
    Dim i           As Integer
    
    On Error GoTo ErrHand
    
    strMedi = zlDatabase.GetPara("�е���ʾ����", glngSys, mlngModule, MSTRMEDICAL)
    arrTypeCols = Split(Split(strMedi, "||")(0) & "|", "|")
    arrMediCols = Split(Split(strMedi, "||")(1), "|")
    
    If mobjXLS Is Nothing Then Call InitExcel
    mobjXLS.SheetsInNewWorkbook = 1  '���½��Ĺ�����������Ϊ1
    mobjXLS.Workbooks.Add          '����һ��������
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "ҩƷ����"  '�޸Ĺ���������
    mobjXLS.Sheets.Add , mobjXLS.Sheets("ҩƷ����") '���ӵڶ����������ڵ�һ��֮��
    mobjXLS.Sheets(mobjXLS.Sheets.Count).Name = "ҩƷ��ϸ"
    
    mobjXLS.Sheets("ҩƷ����").Select     'ѡ�й�����<ҩƷ����>
    'ѭ��д������
    For i = LBound(arrTypeCols) To UBound(arrTypeCols) - 1
        mobjXLS.Cells(1, i + 1) = Split(arrTypeCols(i), ",")(0)
        '����Excel��ע
        For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")) - 1
            If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(0) = Split(arrTypeCols(i), ",")(0) Then
                If Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1) <> "" Then
                    mobjXLS.ActiveSheet.Cells(1, i + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(0) & "|", "|")(intNum), ";")(1)
                End If
                Exit For
            End If
        Next
    Next
    Call SetExcel("����", i, 1)
    
    mobjXLS.Sheets("ҩƷ��ϸ").Select     'ѡ�й�����<ҩƷ��ϸ>
    'ѭ��д������
    For i = LBound(arrMediCols) To UBound(arrMediCols) - 1
        mobjXLS.Cells(1, i + 1) = Split(arrMediCols(i), ",")(0)
        For intNum = 0 To UBound(Split(Split(MSTRCOMMENT, "||")(1), "|")) - 1
            If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(0) = Split(arrMediCols(i), ",")(0) Then
                If Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1) <> "" Then
                    mobjXLS.ActiveSheet.Cells(1, i + 1).AddComment Split(Split(Split(MSTRCOMMENT, "||")(1), "|")(intNum), ";")(1)
                End If
                Exit For
            End If
        Next
    Next
    Call SetExcel("��ϸ", i, 1)
    mobjXLS.Sheets("ҩƷ����").Select
    
    With dlgOpenFile
        .CancelError = False
        .FileName = ""
        .Filter = "*.xlsx|*.xlsx|*.xls|*.xls"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjXLS.ActiveWorkbook.SaveAs strFileName
            blnFinished = True
        End If
    End With
ErrHand:
    mobjXLS.Quit
    If blnFinished Then
        MsgBox "��׼�ļ������Ѿ����ɣ�", vbInformation, gstrSysName
    Else
        MsgBox "��׼�ļ���������ʧ�ܣ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub SetExcel(ByVal strType As String, ByVal intCol As Integer, ByVal intRow As Integer)
'����Excel����
    Dim intCount As Integer
    Dim strMedi  As String
    Dim strStart As String
    Dim strEnd   As String
    Dim strFileColumn As String
    Dim lngCol As Long
    Dim strArr��ϸ������() As String
    Dim strArr��ϸ����() As String
    Dim strArr��������() As String
    Dim arrMediCols() As String
    Dim i As Integer, n As Integer
    
    On Error GoTo ErrHand
    
    strMedi = zlDatabase.GetPara("�е���ʾ����", glngSys, mlngModule, MSTRMEDICAL)
    
    With mobjXLS
        If strType = "����" Then
            For intCount = 0 To intCol - 1
                .ActiveCell(1, intCount + 1).HorizontalAlignment = 3  '��ͷ�ı����ж���
                If Split(Split(Split(strMedi, "||")(0), "|")(intCount), ",")(1) = 0 Then
                    .ActiveSheet.Cells(1, intCount + 1).Interior.Color = &H80FF80  '��ɫΪ������ʾ��Ŀ
                End If
            Next
            '���ñ߿�
            If intCount < 27 Then
                strEnd = Chr(intCount - 1 + 65) & intRow
            Else
                strEnd = "A" & Chr(intCount - 27 + 65) & intRow
            End If
            .Range("A1", strEnd).Borders.Weight = 2
        Else
            For intCount = 0 To intCol - 1
                .ActiveCell(1, intCount + 1).HorizontalAlignment = 3  '��ͷ�ı����ж���
                If Split(Split(Split(strMedi, "||")(1), "|")(intCount), ",")(1) = 0 Then
                    .ActiveSheet.Cells(1, intCount + 1).Interior.Color = &H80FF80  '��ɫΪ������ʾ��Ŀ
                End If
            Next
            '���ñ߿�
            If intCount < 27 Then
                strEnd = Chr(intCount - 1 + 65) & intRow
            Else
                strEnd = "A" & Chr(intCount - 27 + 65) & intRow
            End If
            .Range("A1", strEnd).Borders.Weight = 2
        End If
        
        .Rows("1:1").Select           'ѡ�е�һ��
        .Selection.Font.Bold = True   '��Ϊ����
        .Selection.Font.Size = 11     '���������С
        .Columns.ColumnWidth = 16     '�п�
        .ActiveWindow.SplitRow = 1    '�̶���
        .ActiveWindow.FreezePanes = True
        .ActiveSheet.Rows(1).RowHeight = 25  '�и�
        .ActiveSheet.Rows(1).Insert   '����һ��
        .Cells(1).Value = " ˵������ɫΪ������ʾ��Ŀ�����ʱ��ע��鿴��ע��"
'        .Range("A1:C1").Select        '�ϲ�
'        .Range("A1:C1").Merge
        .Range("A3").Select
        
        'ҩƷ����������
        If strType = "����" Then
            .Sheets("ҩƷ����").Select
            .Columns("A:L").NumberFormatLocal = "@"   '�����ı���ʽ
            arrMediCols = Split(Split(strMedi, "||")(0), "|")
            strFileColumn = ""
            For i = LBound(arrMediCols) To UBound(arrMediCols)
                strFileColumn = strFileColumn & "|" & Trim(Split(arrMediCols(i), ",")(0)) & "," & i + 1
            Next
            strFileColumn = Mid(strFileColumn, 2)
            strArr�������� = Split(strFileColumn, "|")
            For i = LBound(strArr��������) To UBound(strArr��������)
                If Split(strArr��������(i), ",")(0) = "���" Then
                    lngCol = Split(strArr��������(i), ",")(1)
                    .Columns(lngCol).Select
                    With mobjXLS.Selection.Validation
                        .Delete
                        .Add 3, 1, 1, "����ҩ,�г�ҩ,�в�ҩ"
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .InputTitle = ""
                        .ErrorTitle = "��������"
                        .InputMessage = ""
                        .ErrorMessage = "����������б���ѡ��"
                        .IMEMode = 0
                        .ShowInput = True
                        .ShowError = True
                    End With

                    .Rows("1:2").Select
                    With mobjXLS.Selection.Validation
                        .Delete
                    End With
                End If
            Next
        End If
        'ҩƷ��ϸ������
        If strType = "��ϸ" Then
            arrMediCols = Split(Split(strMedi, "||")(1), "|")
            strFileColumn = ""
            For i = LBound(arrMediCols) To UBound(arrMediCols) - 1
                strFileColumn = strFileColumn & "|" & Trim(Split(arrMediCols(i), ",")(0)) & "," & i + 1
            Next
            strFileColumn = Mid(strFileColumn, 2)
            strArr��ϸ���� = Split(strFileColumn, "|")
            strArr��ϸ������ = Split(mstrҩƷ��ϸ, ";")
            For i = LBound(strArr��ϸ����) To UBound(strArr��ϸ����)
                For n = LBound(strArr��ϸ������) To UBound(strArr��ϸ������)
                    If Split(strArr��ϸ����(i), ",")(0) = Split(strArr��ϸ������(n), "|")(0) Then
                        .Sheets("ҩƷ��ϸ").Select
                        .Columns("A:AE").NumberFormatLocal = "@"   '�����ı���ʽ
                        lngCol = Split(strArr��ϸ����(i), ",")(1)
                        .Columns(lngCol).Select
                        With mobjXLS.Selection.Validation
                            .Delete
                            .Add 3, 1, 1, Split(strArr��ϸ������(n), "|")(1)
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .InputTitle = ""
                            .ErrorTitle = "��������"
                            .InputMessage = ""
                            .ErrorMessage = "����������б���ѡ��"
                            .IMEMode = 0
                            .ShowInput = True
                            .ShowError = True
                        End With
                        
                        .Rows("1:2").Select
                        With mobjXLS.Selection.Validation
                            .Delete
                        End With
                    End If
                Next
            Next
        End If
        .Range("A3").Select
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveCard()
    '��������
    Dim cbrControl As CommandBarControl
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo ErrHand
    
    If vsfList.Rows = 1 Then Exit Sub
    '�жϵ��뷽ʽ
    With vsfError
        If .Rows > 1 Then
            For lngRow = 1 To .Rows - 1
                If mbyt���뷽ʽ = 1 Then
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        MsgBox "���ܴ����κβ��ϸ�����ݣ���������", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        If MsgBox("�����ڲ��ϸ����ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End With
    '����
    If TabControl.Selected.Caption = "����" Then
        Call SaveType
    Else
        Call SaveMedi
    End If
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = False
    
    Exit Sub
ErrHand:
    If Not mobjWB Is Nothing Then
        mobjWB.Close
    End If
    Set mobjWB = Nothing
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim cbrControl As CommandBarControl
    Dim rsTemp     As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    mlngModule = glngModul
    Me.Height = 600 * 15
    Me.Width = 800 * 15
    lblCollect.Caption = ""
    
    Call InitComandbar
    Call InitTabControl
    Call GetColumnHead
    Call InitVsf
        
    If vsfList.Rows <= 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    Exit Sub
ErrHand:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub InitVsf()
    '��ʼ��vsf���ؼ�
    With vsfList
        .Rows = 1
        .Cols = 16
        .Editable = flexEDNone
        .ExplorerBar = flexExNone   '�в�֧��������϶�
    End With
    
    With vsfError
        .Rows = 1
        .Cols = 4
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 300
        .TextMatrix(0, 1) = "����λ��"
        .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "��������"
        .ColWidth(2) = 2000
        .TextMatrix(0, 3) = "����ԭ��"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .ScrollBars = flexScrollBarBoth
        '.ExtendLastCol = True '���һ�������
        .ExplorerBar = flexExNone   '�в�֧��������϶�
        .AllowUserResizing = flexResizeNone
    End With
End Sub

Private Sub InitExcel()
    '��ʼ��Excel���
    Set mobjXLS = CreateObject("Excel.Application")
    mobjXLS.DisplayAlerts = False
End Sub

Private Function GetExcelData() As Boolean
    '��ȡexcel������ݣ����������ַ�����ʽ���浽
    '����true-�д��� ����false-û�д���
    Dim strFileColumn As String    '�ļ���������
    Dim blnNotNullRow As Boolean   '�������ǲ��ǿ���
    Dim cbrControl    As CommandBarControl
    Dim lngRow        As Long
    Dim lngCol        As Long
    Dim rsTemp        As Recordset
    Dim strSql        As String
    Dim i             As Integer
    
    On Error GoTo ErrHand
    
    vsfList.Clear
    lblCollect.Caption = ""
    vsfError.Rows = 1

    If txtFile.Text = "" Then Exit Function
    
    If mobjXLS Is Nothing Then Call InitExcel
    Set mobjWB = mobjXLS.Workbooks.Open(txtFile.Text)
    
    '��������ͷ
    Set mobjWS = mobjWB.Sheets(1)
    If mobjWS Is Nothing Then Exit Function
    For lngCol = 1 To mobjWS.UsedRange.Columns.Count
        strFileColumn = strFileColumn & Trim(mobjWS.UsedRange.Cells(2, lngCol)) & "|"
    Next
    For lngCol = 0 To UBound(Split(mstrType, "|")) - 1
        If InStr(1, "|" & strFileColumn, "|" & Split(mstrType, "|")(lngCol) & "|") = 0 Then
            vsfError.Rows = vsfError.Rows + 1
            vsfError.TextMatrix(vsfError.Rows - 1, 1) = "����ҳ"
            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ͷ����"
            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ͷ��" & Split(mstrType, "|")(lngCol) & "��Ϊ��ʾ�У�Ȼ�����ļ��в����ڸ��У�������Ҫ�����Excel�ļ���"
            GetExcelData = True
        End If
    Next
    '�����ϸ��ͷ
    strFileColumn = ""
    Set mobjWS = Nothing
    Set mobjWS = mobjWB.Sheets(2)
    If mobjWS Is Nothing Then Exit Function
    For lngCol = 1 To mobjWS.UsedRange.Columns.Count
        strFileColumn = strFileColumn & Trim(mobjWS.UsedRange.Cells(2, lngCol)) & "|"
    Next
    For lngCol = 0 To UBound(Split(mstrMedi, "|")) - 1
        If InStr(1, "|" & strFileColumn, "|" & Split(mstrMedi, "|")(lngCol) & "|") = 0 Then
            vsfError.Rows = vsfError.Rows + 1
            vsfError.TextMatrix(vsfError.Rows - 1, 1) = "��ϸҳ"
            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ͷ����"
            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ͷ��" & Split(mstrMedi, "|")(lngCol) & "��Ϊ��ʾ�У�Ȼ�����ļ��в����ڸ��У�������Ҫ�����Excel�ļ���"
            GetExcelData = True
        End If
    Next
    
    If GetExcelData = True Then Exit Function
    
    '��ӷ�������
    Set mobjWS = Nothing
    Set mobjWS = mobjWB.Sheets(1)
    If mobjWS Is Nothing Then Exit Function
    With mobjWS.UsedRange
        vsfList.Redraw = flexRDNone
        vsfList.Cols = UBound(Split(mstrType, "|")) + 1
        vsfList.Rows = 1
        
        For lngCol = 1 To UBound(Split(mstrType, "|"))
            vsfList.ColKey(lngCol) = Split(mstrType, "|")(lngCol - 1)
            vsfList.TextMatrix(0, lngCol) = Split(mstrType, "|")(lngCol - 1)
        Next
        
        For i = 1 To vsfList.Cols - 1
            For lngCol = 1 To .Columns.Count
                If vsfList.ColKey(i) = Trim(.Cells(2, lngCol)) Then
                    For lngRow = 3 To .Rows.Count
                        If i = 1 Then
                            vsfList.Rows = vsfList.Rows + 1
                        End If
                        vsfList.TextMatrix(lngRow - 2, i) = Trim(.Cells(lngRow, lngCol))
                    Next
                    Exit For
                End If
            Next
        Next
    End With
    Call GetColumns("����")
    
    '�����ϸ����
    vsfList.Clear
    Set mobjWS = Nothing
    Set mobjWS = mobjWB.Sheets(2)
    If mobjWS Is Nothing Then Exit Function
    With mobjWS.UsedRange
        vsfList.Redraw = flexRDNone
        vsfList.Cols = UBound(Split(mstrMedi, "|")) + 1
        vsfList.Rows = 1
        
        For lngCol = 1 To UBound(Split(mstrMedi, "|"))
            vsfList.ColKey(lngCol) = Split(mstrMedi, "|")(lngCol - 1)
            vsfList.TextMatrix(0, lngCol) = Split(mstrMedi, "|")(lngCol - 1)
        Next
        
        For i = 1 To vsfList.Cols - 1
            For lngCol = 1 To .Columns.Count
                If vsfList.ColKey(i) = Trim(.Cells(2, lngCol)) Then
                    For lngRow = 3 To .Rows.Count
                        If i = 1 Then
                            vsfList.Rows = vsfList.Rows + 1
                        End If
                        vsfList.TextMatrix(lngRow - 2, i) = Trim(.Cells(lngRow, lngCol))
                    Next
                    Exit For
                End If
            Next
        Next
    End With
    Call GetColumns("��ϸ")
    
    If mstrTypeMsg <> "" And TabControl.Selected.Caption = "����" Then
        Call SetColumns("����")
        If vsfList.Rows > 1 Then Call CheckKind
    ElseIf mstrMediMsg <> "" And TabControl.Selected.Caption = "��ϸ" Then
        Call SetColumns("��ϸ")
        If vsfList.Rows > 1 Then Call CheckƷ��: Call Check���
    End If
    
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    mobjXLS.Quit
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetColumns(ByVal strType As String)
    '��ҩƷ��Ϣ��ʾ����������
    Dim blnNotNullRow As Boolean
    Dim cbrControl    As CommandBarControl
    Dim rsTemp        As Recordset
    Dim lngRow        As Long
    Dim lngCol        As Long
    Dim strSql        As String
    
    With vsfList
        .Tag = "1"
        .Clear
        .Redraw = flexRDDirect
        If strType = "����" Then
            .Rows = UBound(Split(mstrTypeMsg, "|"))
            .Cols = UBound(Split(Split(mstrTypeMsg, "|")(0), ";")) + 1
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    If lngRow = 0 Then
                        .ColKey(lngCol) = Split(Split(mstrTypeMsg, "|")(lngRow), ";")(lngCol - 1)
                    End If
                    .TextMatrix(lngRow, lngCol) = Split(Split(mstrTypeMsg, "|")(lngRow), ";")(lngCol - 1)
                Next
            Next
        ElseIf strType = "��ϸ" Then
            strSql = "select ����,���� from ҩƷ���ľ��� where ���=1 and ���� in(1, 2) and ��λ=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "")
            .Rows = UBound(Split(mstrMediMsg, "|"))
            .Cols = UBound(Split(Split(mstrMediMsg, "|")(0), ";")) + 1
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    If lngRow = 0 Then
                        .ColKey(lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                    End If
                    Select Case .ColKey(lngCol)
                        Case "�ɱ���", "�ۼ�"
                            rsTemp.Filter = ""
                            rsTemp.Filter = "����=" & IIf(.ColKey(lngCol) = "�ɱ���", 1, 2)
                            If IsNumeric(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)) Then
                                .TextMatrix(lngRow, lngCol) = zlStr.FormatEx(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1), Val(rsTemp!����), , True)
                            Else
                                .TextMatrix(lngRow, lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                            End If
                        Case "��Ӧ�����֤Ч��"
                            If IsNumeric(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)) Then
                                .TextMatrix(lngRow, lngCol) = TranNumToDate(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1))
                            Else
                                .TextMatrix(lngRow, lngCol) = FormatDate(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1))
                            End If
                        Case "�������"
                            If IsNumeric(Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)) Then
                                Select Case Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                                    Case 0
                                        .TextMatrix(lngRow, lngCol) = "0-�������ڲ���"
                                    Case 1
                                        .TextMatrix(lngRow, lngCol) = "1-����"
                                    Case 2
                                        .TextMatrix(lngRow, lngCol) = "2-סԺ"
                                    Case 3
                                        .TextMatrix(lngRow, lngCol) = "3-�����סԺ"
                                End Select
                            Else
                                .TextMatrix(lngRow, lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                            End If
                        Case Else
                            .TextMatrix(lngRow, lngCol) = Split(Split(mstrMediMsg, "|")(lngRow), ";")(lngCol - 1)
                    End Select
                Next
            Next
        End If
        
        '������ɾ��
        blnNotNullRow = True
        For lngRow = .Rows - 1 To 1 Step -1
            For lngCol = 1 To .Cols - 1
                If .TextMatrix(lngRow, lngCol) <> "" Then
                    blnNotNullRow = False
                End If
            Next
            '����ǿ��н���ɾ��
            If blnNotNullRow = True Then vsfList.RemoveItem lngRow
        Next
        
        If .Rows <= 1 Then
            Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
            cbrControl.Enabled = False
            Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
            cbrControl.Enabled = False
            Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
            cbrControl.Enabled = False
        End If
        
        Call setVSF
        .Tag = "0"
    End With
End Sub

Private Function FormatDate(ByVal StrDate As String) As String
    '���ܣ���ʽ�����ڣ������÷ֺ�(-)�ָ������ڸ�ʽ
    Dim strYear  As String
    Dim strMonth As String
    Dim strDay   As String
    
    If LenB(StrConv(StrDate, vbFromUnicode)) >= 8 Then
        If InStr(1, StrDate, ".") > 0 Or InStr(1, StrDate, "/") > 0 Or InStr(1, StrDate, "-") > 0 Then
            StrDate = Replace(StrDate, ".", "")
            StrDate = Replace(StrDate, "/", "")
            StrDate = Replace(StrDate, "-", "")
        End If
        strYear = Mid(StrDate, 1, 4)
        If LenB(StrConv(StrDate, vbFromUnicode)) < 8 Then
            strMonth = Mid(StrDate, 5, 1)
        Else
            strMonth = Mid(StrDate, 5, 2)
        End If
        If LenB(StrConv(StrDate, vbFromUnicode)) < 8 Then
            strDay = Mid(StrDate, 6, 1)
        Else
            strDay = Mid(StrDate, 7, 2)
        End If
        If IsNumeric(strYear) = True And IsNumeric(strMonth) = True And IsNumeric(strDay) = True Then
            FormatDate = strYear & "-" & IIf(strMonth < 10, "0" & strMonth, strMonth) & "-" & IIf(strDay < 10, "0" & strDay, strDay)
        Else
            FormatDate = StrDate
        End If
    Else
        FormatDate = StrDate
    End If
End Function

Public Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
    'ת����ֵΪ����
    Dim strYear  As String
    Dim strMonth As String
    Dim strDay   As String
    Dim StrDate  As String
    
    TranNumToDate = ""
    If LenB(StrConv(strNum, vbFromUnicode)) < 4 Or LenB(StrConv(strNum, vbFromUnicode)) > 8 Then Exit Function
    
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    StrDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(StrDate) Then Exit Function
    
    StrDate = Format(StrDate, "yyyy-mm-dd")
    If blnDec Then StrDate = DateAdd("d", -1, Format(StrDate, "yyyy-mm-dd"))
    TranNumToDate = StrDate
End Function

Private Function GetColumnPostation(ByVal strColumn As String) As Integer
    '��ȡ��λ�ú��ж��Ƿ����
    '���� strcolumn-���������
    '����ֵ :���ش�����λ�� 0-û���ҵ� >0�ҵ���
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            If .TextMatrix(0, lngCol) = strColumn Then
                GetColumnPostation = lngCol
                Exit Function
            End If
        Next
        GetColumnPostation = 0
    End With
End Function

Private Sub CheckKind()
'���������ݺϷ���
    Dim cbrControl As CommandBarControl
    Dim rsTemp  As Recordset
    Dim lngRow  As Long
    Dim lngCol  As Long
    Dim strTemp As String
    Dim strSql  As String
    Dim j       As Integer
    Dim rs���� As Recordset
    Dim strSqls As String
    
    On Error GoTo ErrHand
    
    Call ParseParameter
    
    strSql = "Select ID,Decode(Substr(����, 1, 1), '0', Substr(����, 2), ����) As ����,����,�ϼ�ID,���� From ���Ʒ���Ŀ¼ Where ���� in (1,2,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���Ʒ���Ŀ¼")
    
    vsfError.Rows = 1
    With vsfList
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack '�����óɺ�ɫ
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow '����б�
            '���
            If GetColumnPostation("���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "����ҩ" And Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "�г�ҩ" And Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "�в�ҩ" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("���"), lngRow, .ColIndex("���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�������ֻ��������ҩ���г�ҩ���в�ҩ��"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("���"), lngRow, .ColIndex("���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������в���Ϊ�գ�"
                End If
            End If
            '�ϼ�����
            If GetColumnPostation("�ϼ�����") > 0 And GetColumnPostation("���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("�ϼ�����"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("�ϼ�����"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�ϼ�����"), lngRow, .ColIndex("�ϼ�����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ϼ����� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ϼ����ơ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ϼ����ơ��в����зǷ��ַ���"
                    Else
                        If GetTypeID(.TextMatrix(lngRow, .ColIndex("�ϼ�����")), .TextMatrix(lngRow, .ColIndex("���"))) = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("�ϼ�����"), lngRow, .ColIndex("�ϼ�����")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�ϼ����� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ϼ����ơ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���Ʒ���Ŀ¼�в��������Ϊ��" & .TextMatrix(lngRow, .ColIndex("���")) & "��������Ϊ��" & .TextMatrix(lngRow, .ColIndex("�ϼ�����")) & "�����"
                        End If
                    End If
                End If
            End If
            '����
            If GetColumnPostation("����") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("����"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����롿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����롿�в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("����"))), vbFromUnicode)) > rsTemp("����").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����롿��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����롿���ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("����").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����롿��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����롿�в���Ϊ�գ�"
                End If
            End If
            '����
            If GetColumnPostation("����") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("����"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ơ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ơ��в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("����"))), vbFromUnicode)) > rsTemp("����").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ơ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ơ����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("����").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ơ���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ơ��в���Ϊ�գ�"
                End If
            End If
            '��������Ψһ
            If GetColumnPostation("���") > 0 And GetColumnPostation("����") > 0 Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("���")) = .TextMatrix(j, .ColIndex("���")) And .TextMatrix(lngRow, .ColIndex("����")) = .TextMatrix(j, .ColIndex("����")) And .TextMatrix(lngRow, .ColIndex("����")) <> "" And .TextMatrix(lngRow, .ColIndex("���")) <> "" Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����롿��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������ǰ���Ѵ������Ϊ��" & Trim(.TextMatrix(lngRow, .ColIndex("���"))) & "��������Ϊ��" & Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "�������ݣ����飡"
                        End If
                    Next
                End If
                If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "'") = 0 And Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "" Then
                    rsTemp.Filter = ""
                    rsTemp.Filter = "����=" & Switch(Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "����ҩ", 1, Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "�г�ҩ", 2, Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "�в�ҩ", 3) & " and ����='" & IIf(Mid(Trim(.TextMatrix(lngRow, .ColIndex("����"))), 1, 1) = 0, Mid(Trim(.TextMatrix(lngRow, .ColIndex("����"))), 2), Trim(.TextMatrix(lngRow, .ColIndex("����")))) & "'"
                    If rsTemp.RecordCount > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����롿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������Ŀ�����ݿ����������ݳ�ͻ�����" & Trim(.TextMatrix(lngRow, .ColIndex("���"))) & "�����Ѵ��ڱ��롾" & Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "��"
                    End If
                End If
            End If
            '���ơ�����ϼ�����Ψһ
            If GetColumnPostation("���") > 0 And GetColumnPostation("�ϼ�����") > 0 And GetColumnPostation("����") > 0 Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("���")) = .TextMatrix(j, .ColIndex("���")) And .TextMatrix(lngRow, .ColIndex("�ϼ�����")) = .TextMatrix(j, .ColIndex("�ϼ�����")) And .TextMatrix(lngRow, .ColIndex("����")) = .TextMatrix(j, .ColIndex("����")) And .TextMatrix(lngRow, .ColIndex("���")) <> "" And .TextMatrix(lngRow, .ColIndex("����")) <> "" Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ơ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������ǰ���Ѵ������" & Trim(.TextMatrix(lngRow, .ColIndex("���"))) & "�����ϼ����ơ�" & Trim(.TextMatrix(lngRow, .ColIndex("�ϼ�����"))) & "�������ơ�" & Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "�������ݣ����飡"
                        End If
                    Next
                End If
                If GetTypeID(.TextMatrix(lngRow, .ColIndex("�ϼ�����")), .TextMatrix(lngRow, .ColIndex("���"))) >= 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("�ϼ�����"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "'") = 0 Then
                    strSqls = "Select ID,Decode(Substr(����, 1, 1), '0', Substr(����, 2), ����) As ����,����,�ϼ�ID,���� From ���Ʒ���Ŀ¼ Where ���� in (1,2,3) and ����=[1] and �ϼ�ID" & IIf(GetTypeID(.TextMatrix(lngRow, .ColIndex("�ϼ�����")), .TextMatrix(lngRow, .ColIndex("���"))) = 0, " is null", "=" & GetTypeID(.TextMatrix(lngRow, .ColIndex("�ϼ�����")), .TextMatrix(lngRow, .ColIndex("���"))))
                    Set rs���� = zlDatabase.OpenSQLRecord(strSqls, "������ĿĿ¼", Trim(.TextMatrix(lngRow, .ColIndex("����"))))
                    If rs����.RecordCount > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNameAndKind = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ơ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������Ŀ�����ݿ����������ݳ�ͻ�����" & Trim(.TextMatrix(lngRow, .ColIndex("���"))) & "�����ϼ����ơ�" & Trim(.TextMatrix(lngRow, .ColIndex("�ϼ�����"))) & "�����Ѵ������ơ�" & Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "��"
                    End If
                End If
            End If
        Next
    End With
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt���뷽ʽ = 0 Then
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    Else
                        cbrControl.Enabled = True
                    End If
                Next
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
    
    If vsfList.Rows > 1 Then
        vsfList.Row = 1: vsfList.Col = 1
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CheckƷ��()
'�����ϸ���ݣ�Ʒ�֣��Ϸ���
    Dim cbrControl As CommandBarControl
    Dim rsTemp  As Recordset
    Dim rs����  As Recordset
    Dim lngRow  As Long
    Dim lngCol  As Long
    Dim strTemp As String
    Dim strSql  As String
    Dim j       As Integer
    Dim rs���� As Recordset
    Dim strSqls As String

    On Error GoTo ErrHand
    
    Call ParseParameter
    
    strSql = "Select ���,����ID,ID,����,����,���㵥λ From ������ĿĿ¼ Where ��� In ('5', '6', '7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "������ĿĿ¼")
    Set rs���� = zlDatabase.OpenSQLRecord("select ���� from ҩƷ����", "ҩƷ����")
    
    vsfError.Rows = 1
    With vsfList
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack '�����óɺ�ɫ
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow '����б�
            '���
            If GetColumnPostation("���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "" Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("���"), lngRow, .ColIndex("���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������в���Ϊ�գ�"
                Else
                    If Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "����ҩ" And Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "�г�ҩ" And Trim(.TextMatrix(lngRow, .ColIndex("���"))) <> "�в�ҩ" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("���"), lngRow, .ColIndex("���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�������ֻ��������ҩ���г�ҩ���в�ҩ��"
                    End If
                End If
            End If
            '����
            If GetColumnPostation("����") > 0 And GetColumnPostation("���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("����"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytҩƷ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ࡿ��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ࡿ�в����зǷ��ַ���"
                    Else
                        If GetTypeID(.TextMatrix(lngRow, .ColIndex("����")), .TextMatrix(lngRow, .ColIndex("���"))) = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytҩƷ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ࡿ��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���Ʒ���Ŀ¼�в��������Ϊ��" & .TextMatrix(lngRow, .ColIndex("���")) & "��������Ϊ��" & .TextMatrix(lngRow, .ColIndex("����")) & "�����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytҩƷ���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ࡿ��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ࡿ�в���Ϊ�գ�"
                End If
            End If
            'Ʒ�ֱ���
            If GetColumnPostation("Ʒ�ֱ���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ�ֱ���"), lngRow, .ColIndex("Ʒ�ֱ���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ�ֱ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�ֱ��롿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʒ�ֱ��롿�в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))), vbFromUnicode)) > rsTemp("����").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ�ֱ���"), lngRow, .ColIndex("Ʒ�ֱ���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ�ֱ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�ֱ��롿��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʒ�ֱ��롿���ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("����").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ�ֱ���"), lngRow, .ColIndex("Ʒ�ֱ���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ�ֱ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�ֱ��롿��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʒ�ֱ��롿�в���Ϊ�գ�"
                End If
            End If
            'Ʒ������
            If GetColumnPostation("Ʒ������") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ������"), lngRow, .ColIndex("Ʒ������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ������ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�����ơ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʒ�����ơ��в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))), vbFromUnicode)) > rsTemp("����").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ������"), lngRow, .ColIndex("Ʒ������")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ������ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�����ơ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʒ�����ơ����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("����").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ������"), lngRow, .ColIndex("Ʒ������")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ������ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�����ơ���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ʒ�����ơ��в���Ϊ�գ�"
                End If
            End If
            '����
            If GetColumnPostation("����") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("����"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����͡���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����͡��в����зǷ��ַ���"
                    Else
                        rs����.Filter = ""
                        rs����.Filter = "����='" & Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "'"
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("����"))), vbFromUnicode)) > rs����("����").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����͡���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����͡����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rs����("����").DefinedSize & "����"
                        Else
                            If rs����.RecordCount = 0 Then
                                .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                                vsfError.Rows = vsfError.Rows + 1
                                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����͡���"
                                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����͡�ֻ�������ݿ�����������ݣ����͡�" & Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "�������ڣ�"
                            End If
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("����"), lngRow, .ColIndex("����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����͡���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����͡��в���Ϊ�գ�"
                End If
            End If
            '������λ
            If GetColumnPostation("������λ") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("������λ"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("������λ"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("������λ"), lngRow, .ColIndex("������λ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�������λ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������λ���в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("������λ"))), vbFromUnicode)) > rsTemp("���㵥λ").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("������λ"), lngRow, .ColIndex("������λ")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�������λ����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������λ�����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("���㵥λ").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("������λ"), lngRow, .ColIndex("������λ")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�������λ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������λ���в���Ϊ�գ�"
                End If
            End If
            'Ʒ��Ψһ��
            If GetColumnPostation("���") > 0 And GetColumnPostation("����") > 0 And GetColumnPostation("Ʒ������") Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("���")) = .TextMatrix(j, .ColIndex("���")) And .TextMatrix(lngRow, .ColIndex("����")) = .TextMatrix(j, .ColIndex("����")) And .TextMatrix(lngRow, .ColIndex("Ʒ������")) <> .TextMatrix(j, .ColIndex("Ʒ������")) And .TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���")) = .TextMatrix(j, .ColIndex("Ʒ�ֱ���")) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ������"), lngRow, .ColIndex("Ʒ������")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ��Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�����ơ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�ڡ�" & j & "���к͵ڡ�" & lngRow & "���У�ͬ���ͬ���ࡢͬƷ�ֱ��룬Ʒ�����ơ�" & Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))) & "������ͬ�����飡"
                        ElseIf .TextMatrix(lngRow, .ColIndex("���")) = .TextMatrix(j, .ColIndex("���")) And .TextMatrix(lngRow, .ColIndex("����")) = .TextMatrix(j, .ColIndex("����")) And .TextMatrix(lngRow, .ColIndex("Ʒ������")) = .TextMatrix(j, .ColIndex("Ʒ������")) And .TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���")) <> .TextMatrix(j, .ColIndex("Ʒ�ֱ���")) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ�ֱ���"), lngRow, .ColIndex("Ʒ�ֱ���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ��Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�ֱ��롿��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�ڡ�" & j & "���к͵ڡ�" & lngRow & "���У�ͬ���ͬ���ࡢͬƷ�����ơ�Ʒ�ֱ��롾" & Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))) & "������ͬ�����飡"
                        End If
                    Next
                End If
            End If
            '�����з����½������ļ��
            If GetColumnPostation("���") > 0 And GetColumnPostation("����") > 0 And GetColumnPostation("Ʒ������") > 0 And GetColumnPostation("Ʒ�ֱ���") > 0 Then
                If GetTypeID(Trim(.TextMatrix(lngRow, .ColIndex("����"))), .TextMatrix(lngRow, .ColIndex("���"))) > 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("����"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))), "'") = 0 Then
                    rsTemp.Filter = ""
                    rsTemp.Filter = "����='" & Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))) & "'"
                    If rsTemp.RecordCount > 0 Then '������ݿ���ڽ���¼��ġ�Ʒ�ֱ��롿���ͼ�����¼������¡�Ʒ�ֱ��롿��Ʒ�����ơ��������Ƿ�һ��
                        strSqls = "Select ���,����ID,ID,����,����,���㵥λ From ������ĿĿ¼ Where ��� In ('5', '6', '7') and ����ID=[1] and ����=[2] and ����=[3] "
                        Set rs���� = zlDatabase.OpenSQLRecord(strSqls, "������ĿĿ¼", GetTypeID(Trim(.TextMatrix(lngRow, .ColIndex("����"))), .TextMatrix(lngRow, .ColIndex("���"))), Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))), Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))))
                        If rs����.RecordCount = 0 Then '������ݿ���ڽ���¼��ġ�Ʒ�ֱ��롿���ҽ���¼������¡�Ʒ�ֱ��롿��Ʒ�����ơ������в�һ��
                            .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ�ֱ���"), lngRow, .ColIndex("Ʒ�ֱ���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ��Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�ֱ��롿��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������Ŀ�����ݿ����������ݳ�ͻ�����롾" & Trim(.TextMatrix(lngRow, .ColIndex("Ʒ�ֱ���"))) & "���Ѵ��ڣ�"
                        End If
                    Else '������ݿⲻ���ڽ���¼��ġ�Ʒ�ֱ��롿���ͼ�����¼������¡�Ʒ�����ơ��������Ƿ�һ��
                        strSqls = "Select ���,����ID,ID,����,����,���㵥λ From ������ĿĿ¼ Where ��� In ('5', '6', '7') and ����ID=[1] and ����=[2]  "
                        Set rs���� = zlDatabase.OpenSQLRecord(strSqls, "������ĿĿ¼", GetTypeID(Trim(.TextMatrix(lngRow, .ColIndex("����"))), .TextMatrix(lngRow, .ColIndex("���"))), Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))))
                        If rs����.RecordCount > 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("Ʒ������"), lngRow, .ColIndex("Ʒ������")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytƷ��Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ʒ�����ơ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������Ŀ�����ݿ������������Ƿ��ͻ�����" & Trim(.TextMatrix(lngRow, .ColIndex("���"))) & "���ͷ��ࡾ" & Trim(.TextMatrix(lngRow, .ColIndex("����"))) & "���£�Ʒ�֡�" & Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))) & "���Ѵ��ڣ�"
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt���뷽ʽ = 0 Then
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    Else
                        cbrControl.Enabled = True
                    End If
                Next
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
    If vsfList.Rows > 1 Then
        vsfList.Row = 1: vsfList.Col = 1
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Check���()
'�����ϸ���ݣ���񣩺Ϸ���
    Dim cbrControl As CommandBarControl
    Dim rs������Ŀ As Recordset
    Dim rs��Ӧ��   As Recordset
    Dim rsTemp     As Recordset
    Dim rs����     As Recordset
    Dim lngRow     As Long
    Dim lngCol     As Long
    Dim strTemp    As String
    Dim strSql     As String
    Dim j          As Integer
    Dim rs���� As Recordset
    Dim strSqls As String
    
    On Error GoTo ErrHand
    
    Call ParseParameter
    
    strSql = "Select a.���, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, b.����ϵ��, b.���ﵥλ, b.�����װ, b.סԺ��λ, b.סԺ��װ, b.ҩ�ⵥλ, b.ҩ���װ," & vbNewLine & _
             "b.���Ч��, b.סԺ�ɷ����, b.ҩ�����, b.ҩ������, b.�ɱ���, b.��ͬ��λid, b.����ɷ����" & vbNewLine & _
             "From �շ���ĿĿ¼ A, ҩƷ��� B Where a.Id = b.ҩƷid And a.��� In ('5', '6', '7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "ҩƷ���")
    Set rs���� = zlDatabase.OpenSQLRecord("select ���,����,��λ,���� from ҩƷ���ľ��� where ���=1", "�۸񾫶�")
    Set rs������Ŀ = zlDatabase.OpenSQLRecord("Select ID,����,���� From ������Ŀ Where ĩ�� = 1", "������Ŀ")
    Set rs��Ӧ�� = zlDatabase.OpenSQLRecord("Select ID,����,����,���֤��,���֤Ч�� From ��Ӧ��", "��Ӧ��")
    
    With vsfList
        For lngRow = 1 To .Rows - 1
            '������
            If GetColumnPostation("������") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("������"))) <> "" Then
                      If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("������"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������롿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�������롿�в����зǷ��ַ���"
                    Else
                        If lngRow > 1 Then
                            For j = lngRow - 1 To 1 Step -1
                                If .TextMatrix(lngRow, .ColIndex("������")) = .TextMatrix(j, .ColIndex("������")) Then
                                    .Cell(flexcpForeColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = vbRed
                                    vsfError.Rows = vsfError.Rows + 1
                                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������롿��"
                                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������ǰ���Ѵ��ڹ�����Ϊ��" & Trim(.TextMatrix(lngRow, .ColIndex("������"))) & "�������ݣ����飡"
                                End If
                            Next
                        End If
                        rsTemp.Filter = ""
                        rsTemp.Filter = "����='" & Trim(.TextMatrix(lngRow, .ColIndex("������"))) & "'"
                        If rsTemp.RecordCount > 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������롿��"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������Ŀ�����ݿ����������ݳ�ͻ�����롾" & Trim(.TextMatrix(lngRow, .ColIndex("������"))) & "���Ѵ��ڣ�"
                        Else
                            If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("������"))), vbFromUnicode)) > rsTemp("����").DefinedSize Then
                                .Cell(flexcpForeColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = vbRed
                                vsfError.Rows = vsfError.Rows + 1
                                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������롿��"
                                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�������롿���ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("����").DefinedSize & "����"
                            End If
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������롿��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�������롿�в���Ϊ�գ�"
                End If
            End If
            'ҩƷ���
            If GetColumnPostation("ҩƷ���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ���"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ���"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩƷ���"), lngRow, .ColIndex("ҩƷ���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytҩƷ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩƷ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩƷ����в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ���"))), vbFromUnicode)) > rsTemp("���").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("ҩƷ���"), lngRow, .ColIndex("ҩƷ���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytҩƷ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩƷ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩƷ������ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("���").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("ҩƷ���"), lngRow, .ColIndex("ҩƷ���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytҩƷ��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩƷ�����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩƷ����в���Ϊ�գ�"
                End If
            End If
            '������
            If GetColumnPostation("������") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("������"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("������"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������̡���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�������̡��в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("������"))), vbFromUnicode)) > rsTemp("����").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("������"), lngRow, .ColIndex("������")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С������̡���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�������̡����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("����").DefinedSize & "����"
                        End If
                    End If
                End If
            End If
            '�ۼ۵�λ
            If GetColumnPostation("�ۼ۵�λ") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("�ۼ۵�λ"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("�ۼ۵�λ"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼ۵�λ"), lngRow, .ColIndex("�ۼ۵�λ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼ۵�λ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼ۵�λ���в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("�ۼ۵�λ"))), vbFromUnicode)) > rsTemp("���㵥λ").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼ۵�λ"), lngRow, .ColIndex("�ۼ۵�λ")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼ۵�λ����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼ۵�λ�����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("���㵥λ").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼ۵�λ"), lngRow, .ColIndex("�ۼ۵�λ")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼ۵�λ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼ۵�λ���в���Ϊ�գ�"
                End If
            End If
            '�ۼۻ���ϵ��
            If GetColumnPostation("�ۼۻ���ϵ��") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("�ۼۻ���ϵ��"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("�ۼۻ���ϵ��")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("�ۼۻ���ϵ��")))) <= 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼۻ���ϵ��"), lngRow, .ColIndex("�ۼۻ���ϵ��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼۻ���ϵ������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼۻ���ϵ������ֻ��������0-9�����>0��"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼۻ���ϵ��"), lngRow, .ColIndex("�ۼۻ���ϵ��")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼۻ���ϵ������"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼۻ���ϵ�����в���Ϊ�գ�"
                End If
            End If
            '���ﵥλ
            If GetColumnPostation("���ﵥλ") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("���ﵥλ"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("���ﵥλ"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("���ﵥλ"), lngRow, .ColIndex("���ﵥλ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ﵥλ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ﵥλ���в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("���ﵥλ"))), vbFromUnicode)) > rsTemp("���ﵥλ").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("���ﵥλ"), lngRow, .ColIndex("���ﵥλ")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ﵥλ����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ﵥλ�����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("���ﵥλ").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("���ﵥλ"), lngRow, .ColIndex("���ﵥλ")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ﵥλ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ﵥλ���в���Ϊ�գ�"
                End If
            End If
            '���ﻻ��ϵ��
            If GetColumnPostation("���ﻻ��ϵ��") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("���ﻻ��ϵ��"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("���ﻻ��ϵ��")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("���ﻻ��ϵ��")))) <= "0" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("���ﻻ��ϵ��"), lngRow, .ColIndex("���ﻻ��ϵ��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ﻻ��ϵ������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ﻻ��ϵ������ֻ��������0-9�����>0��"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("���ﻻ��ϵ��"), lngRow, .ColIndex("���ﻻ��ϵ��")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����ﻻ��ϵ������"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����ﻻ��ϵ�����в���Ϊ�գ�"
                End If
            End If
            'סԺ��λ
            If GetColumnPostation("סԺ��λ") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("סԺ��λ"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("סԺ��λ"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ��λ"), lngRow, .ColIndex("סԺ��λ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ��λ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��סԺ��λ���в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("סԺ��λ"))), vbFromUnicode)) > rsTemp("סԺ��λ").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ��λ"), lngRow, .ColIndex("סԺ��λ")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ��λ����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��סԺ��λ�����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("סԺ��λ").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ��λ"), lngRow, .ColIndex("סԺ��λ")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ��λ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��סԺ��λ���в���Ϊ�գ�"
                End If
            End If
            'סԺ����ϵ��
            If GetColumnPostation("סԺ����ϵ��") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("סԺ����ϵ��"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("סԺ����ϵ��")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("סԺ����ϵ��")))) <= "0" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ����ϵ��"), lngRow, .ColIndex("סԺ����ϵ��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ����ϵ������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��סԺ����ϵ������ֻ��������0-9�����>0��"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ����ϵ��"), lngRow, .ColIndex("סԺ����ϵ��")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ����ϵ������"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��סԺ����ϵ�����в���Ϊ�գ�"
                End If
            End If
            'ҩ�ⵥλ
            If GetColumnPostation("ҩ�ⵥλ") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�ⵥλ"), lngRow, .ColIndex("ҩ�ⵥλ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ�ⵥλ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩ�ⵥλ���в����зǷ��ַ���"
                    Else
                        If LenB(StrConv(Trim(.TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ"))), vbFromUnicode)) > rsTemp("ҩ�ⵥλ").DefinedSize Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�ⵥλ"), lngRow, .ColIndex("ҩ�ⵥλ")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ�ⵥλ����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩ�ⵥλ�����ֶγ��Ȳ��ܳ������ݿ��ֶγ��ȡ�" & rsTemp("ҩ�ⵥλ").DefinedSize & "����"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�ⵥλ"), lngRow, .ColIndex("ҩ�ⵥλ")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��λ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ�ⵥλ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩ�ⵥλ���в���Ϊ�գ�"
                End If
            End If
            'ҩ�⻻��ϵ��
            If GetColumnPostation("ҩ�⻻��ϵ��") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("ҩ�⻻��ϵ��"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("ҩ�⻻��ϵ��")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("ҩ�⻻��ϵ��")))) <= "0" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�⻻��ϵ��"), lngRow, .ColIndex("ҩ�⻻��ϵ��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ�⻻��ϵ������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩ�⻻��ϵ������ֻ��������0-9�����>0��"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�⻻��ϵ��"), lngRow, .ColIndex("ҩ�⻻��ϵ��")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ�⻻��ϵ������"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩ�⻻��ϵ�����в���Ϊ�գ�"
                End If
            End If
            '���סԺ��ҩ�ⵥλ��ͬ������ϵ���Ƚ�
            If GetColumnPostation("���ﵥλ") > 0 And GetColumnPostation("סԺ��λ") > 0 And GetColumnPostation("���ﻻ��ϵ��") > 0 And GetColumnPostation("סԺ����ϵ��") > 0 And GetColumnPostation("ҩ�ⵥλ") > 0 And GetColumnPostation("ҩ�⻻��ϵ��") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("���ﵥλ"))) = Trim(.TextMatrix(lngRow, .ColIndex("סԺ��λ"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("���ﻻ��ϵ��"))) <> Trim(.TextMatrix(lngRow, .ColIndex("סԺ����ϵ��"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ����ϵ��"), lngRow, .ColIndex("סԺ����ϵ��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ����ϵ������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���סԺ��λ��ͬ������ϵ��������ͬ��"
                    End If
                End If
                If Trim(.TextMatrix(lngRow, .ColIndex("���ﵥλ"))) = Trim(.TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("���ﻻ��ϵ��"))) <> Trim(.TextMatrix(lngRow, .ColIndex("ҩ�⻻��ϵ��"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�⻻��ϵ��"), lngRow, .ColIndex("ҩ�⻻��ϵ��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ�⻻��ϵ������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ҩ�ⵥλ��ͬ������ϵ��������ͬ��"
                    End If
                End If
            End If
            '���סԺ��ҩ�⻻��ϵ����ͬ����λ�Ƚ�
            If GetColumnPostation("���ﵥλ") > 0 And GetColumnPostation("סԺ��λ") > 0 And GetColumnPostation("���ﻻ��ϵ��") > 0 And GetColumnPostation("סԺ����ϵ��") > 0 And GetColumnPostation("ҩ�ⵥλ") > 0 And GetColumnPostation("ҩ�⻻��ϵ��") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("���ﻻ��ϵ��"))) = Trim(.TextMatrix(lngRow, .ColIndex("סԺ����ϵ��"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("���ﵥλ"))) <> Trim(.TextMatrix(lngRow, .ColIndex("סԺ��λ"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ��λ"), lngRow, .ColIndex("סԺ��λ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ��λ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���סԺ����ϵ����ͬ����λ������ͬ��"
                    End If
                End If
                If Trim(.TextMatrix(lngRow, .ColIndex("���ﻻ��ϵ��"))) = Trim(.TextMatrix(lngRow, .ColIndex("ҩ�⻻��ϵ��"))) Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("���ﵥλ"))) <> Trim(.TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�ⵥλ"), lngRow, .ColIndex("ҩ�ⵥλ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt����ϵ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ�ⵥλ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ҩ�⻻��ϵ����ͬ����λ������ͬ��"
                    End If
                End If
            End If
            '�Ƿ���
            If GetColumnPostation("�Ƿ���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("�Ƿ���"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("�Ƿ���"))) <> "��" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�Ƿ���"), lngRow, .ColIndex("�Ƿ���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��Ƿ��ۡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���Ƿ��ۡ���ֻ���ǡ��̡���գ�"
                    End If
                End If
            End If
            '�ɱ���
            If GetColumnPostation("�ɱ���") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���")))) < 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ���"), lngRow, .ColIndex("�ɱ���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�۸� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ��ۡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ���ʹ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ��ۡ���ֻ������������Ҳ�С��0��"
                    Else
                        rs����.Filter = ""
                        rs����.Filter = "����=1 and ��λ=1"
                        If LenB(StrConv(Mid(Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���"))), InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("�ɱ���"))), ".") + 1), vbFromUnicode)) > rs����!���� Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ���"), lngRow, .ColIndex("�ɱ���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�۸� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ��ۡ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ��ۡ����ֶξ��Ȳ��ܳ���������þ��ȣ�"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("�ɱ���"), lngRow, .ColIndex("�ɱ���")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�۸� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ɱ��ۡ���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ɱ��ۡ��в���Ϊ�գ�"
                End If
            End If
            '�ۼ�
            If GetColumnPostation("�ۼ�") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("�ۼ�"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("�ۼ�")))) Or Val(Trim(.TextMatrix(lngRow, .ColIndex("�ۼ�")))) < 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼ�"), lngRow, .ColIndex("�ۼ�")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�۸� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼۡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ���ʹ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼۡ���ֻ������������Ҳ�С��0��"
                    Else
                        rs����.Filter = ""
                        rs����.Filter = "����=2 and ��λ=1"
                        If LenB(StrConv(Mid(Trim(.TextMatrix(lngRow, .ColIndex("�ۼ�"))), InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("�ۼ�"))), ".") + 1), vbFromUnicode)) > rs����!���� Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼ�"), lngRow, .ColIndex("�ۼ�")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�۸� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼۡ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼۡ����ֶξ��Ȳ��ܳ���������þ��ȣ�"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("�ۼ�"), lngRow, .ColIndex("�ۼ�")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�۸� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С��ۼۡ���"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "���ۼۡ��в���Ϊ�գ�"
                End If
            End If
            'Ч��(��)
            If GetColumnPostation("Ч��(��)") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("Ч��(��)"))) <> "" Then
                    If Not IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("Ч��(��)")))) Or InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("Ч��(��)"))), ".") > 0 Or Val(Trim(.TextMatrix(lngRow, .ColIndex("Ч��(��)")))) < 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("Ч��(��)"), lngRow, .ColIndex("Ч��(��)")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytЧ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�Ч��(��)����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ���ʹ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��Ч��(��)����ֻ���������Ҳ�С��0��"
                    End If
                End If
            End If
            '������Ŀ
            If GetColumnPostation("������Ŀ") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("������Ŀ"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("������Ŀ"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("������Ŀ"), lngRow, .ColIndex("������Ŀ")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������Ŀ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�������Ŀ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������Ŀ���в����зǷ��ַ���"
                    Else
                        rs������Ŀ.Filter = ""
                        rs������Ŀ.Filter = "����='" & Trim(.TextMatrix(lngRow, .ColIndex("������Ŀ"))) & "'"
                        If rs������Ŀ.RecordCount = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("������Ŀ"), lngRow, .ColIndex("������Ŀ")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������Ŀ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�������Ŀ����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������Ŀ����ֻ�������ݿ�����������Ŀ��������Ŀ��" & Trim(.TextMatrix(lngRow, .ColIndex("������Ŀ"))) & "�������ڣ�"
                        End If
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("������Ŀ"), lngRow, .ColIndex("������Ŀ")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������Ŀ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�������Ŀ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������Ŀ���в���Ϊ�գ�"
                End If
            End If
            'סԺ�ɷ���
            If GetColumnPostation("סԺ�ɷ����") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("סԺ�ɷ����"))) <> "" Then
                    If InStr(1, ",0-���Է���,1-���ɷ���,2-һ����ʹ��,3-�����һ������Ч,4-�������������Ч,5-�������������Ч,", "," & Trim(.TextMatrix(lngRow, .ColIndex("סԺ�ɷ����"))) & ",") = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ�ɷ����"), lngRow, .ColIndex("סԺ�ɷ����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ�ɷ���㡿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��סԺ�ɷ���㡿��ֻ�������з��㷽ʽ�����㷽ʽ��" & Trim(.TextMatrix(lngRow, .ColIndex("סԺ�ɷ����"))) & "�������ڣ�"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("סԺ�ɷ����"), lngRow, .ColIndex("סԺ�ɷ����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�סԺ�ɷ���㡿��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��סԺ�ɷ���㡿�в���Ϊ�գ�"
                End If
            End If
            '����ɷ���
            If GetColumnPostation("����ɷ����") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("����ɷ����"))) <> "" Then
                    If InStr(1, ",0-���Է���,1-���ɷ���,2-һ����ʹ��,3-�����һ������Ч,4-�������������Ч,5-�������������Ч,", "," & Trim(.TextMatrix(lngRow, .ColIndex("����ɷ����"))) & ",") = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("����ɷ����"), lngRow, .ColIndex("����ɷ����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�����ɷ���㡿��"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������ɷ���㡿��ֻ�������з��㷽ʽ�����㷽ʽ��" & Trim(.TextMatrix(lngRow, .ColIndex("����ɷ����"))) & "�������ڣ�"
                    End If
                Else
                    .Cell(flexcpForeColor, lngRow, .ColIndex("����ɷ����"), lngRow, .ColIndex("����ɷ����")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�����ɷ���㡿��"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ֵ����"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������ɷ���㡿�в���Ϊ�գ�"
                End If
            End If
            '�������
            If GetColumnPostation("�������") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("�������"))) <> "" Then
                    If InStr(1, ",0-�������ڲ���,1-����,2-סԺ,3-�����סԺ,", "," & Trim(.TextMatrix(lngRow, .ColIndex("�������"))) & ",") = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("�������"), lngRow, .ColIndex("�������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С����������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "�����������ֻ�������з�����󣡷������" & Trim(.TextMatrix(lngRow, .ColIndex("�������"))) & "�������ڣ�"
                    End If
                End If
            End If
            'ҩ�����
            If GetColumnPostation("ҩ�����") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("ҩ�����"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("ҩ�����"))) <> "��" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ�����"), lngRow, .ColIndex("ҩ�����")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ���������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩ���������ֻ���ǡ��̡���գ�"
                    End If
                End If
            End If
            'ҩ������
            If GetColumnPostation("ҩ������") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("ҩ������"))) <> "" Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("ҩ������"))) <> "��" Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ������"), lngRow, .ColIndex("ҩ������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ����������"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��ҩ����������ֻ���ǡ��̡���գ�"
                    End If
                    If GetColumnPostation("ҩ�����") > 0 Then
                        If Trim(.TextMatrix(lngRow, .ColIndex("ҩ�����"))) = "" Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("ҩ������"), lngRow, .ColIndex("ҩ������")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt�������� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩ����������"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "ҩ�����ʱҩ�����ܷ�����"
                        End If
                    End If
                End If
            End If
            '��Ӧ������
            If GetColumnPostation("��Ӧ������") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("��Ӧ������"))) <> "" Then
                    If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("��Ӧ������"))), "'") > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("��Ӧ������"), lngRow, .ColIndex("��Ӧ������")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��Ӧ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ӧ�����ơ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����Ӧ�����ơ��в����зǷ��ַ���"
                    Else
                        rs��Ӧ��.Filter = ""
                        rs��Ӧ��.Filter = "����='" & Trim(.TextMatrix(lngRow, .ColIndex("��Ӧ������"))) & "'"
                        If rs��Ӧ��.RecordCount = 0 Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("��Ӧ������"), lngRow, .ColIndex("��Ӧ������")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt��Ӧ�� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ӧ�����ơ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "ֵ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����Ӧ�����ơ���ֻ�������ݿ����й�Ӧ�����ƣ���Ӧ�̡�" & Trim(.TextMatrix(lngRow, .ColIndex("��Ӧ������"))) & "�������ڣ�"
                        End If
                    End If
                End If
            End If
            '��Ӧ�����֤Ч��
            If GetColumnPostation("��Ӧ�����֤Ч��") > 0 Then
                If Trim(.TextMatrix(lngRow, .ColIndex("��Ӧ�����֤Ч��"))) <> "" Then
                    If Not IsDate(Trim(.TextMatrix(lngRow, .ColIndex("��Ӧ�����֤Ч��")))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("��Ӧ�����֤Ч��"), lngRow, .ColIndex("��Ӧ�����֤Ч��")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���� = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С���Ӧ�����֤Ч�ڡ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "��ʽ����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "����Ӧ�����֤Ч�ڡ������ڸ�ʽ����"
                    End If
                End If
            End If
            '���Ψһ��
            If GetColumnPostation("���") > 0 And GetColumnPostation("Ʒ������") > 0 And GetColumnPostation("ҩƷ���") > 0 And GetColumnPostation("������") > 0 Then
                If lngRow > 1 Then
                    For j = lngRow - 1 To 1 Step -1
                        If .TextMatrix(lngRow, .ColIndex("���")) = .TextMatrix(j, .ColIndex("���")) And .TextMatrix(lngRow, .ColIndex("Ʒ������")) = .TextMatrix(j, .ColIndex("Ʒ������")) And .TextMatrix(lngRow, .ColIndex("������")) = .TextMatrix(j, .ColIndex("������")) And .TextMatrix(lngRow, .ColIndex("ҩƷ���")) = .TextMatrix(j, .ColIndex("ҩƷ���")) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("ҩƷ���"), lngRow, .ColIndex("ҩƷ���")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩƷ�����"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "��������ǰ���Ѵ������" & Trim(.TextMatrix(lngRow, .ColIndex("���"))) & "����Ʒ�����ơ�" & Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))) & "���������̡�" & Trim(.TextMatrix(lngRow, .ColIndex("������"))) & "����ҩƷ���" & Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ���"))) & "�������ݣ����飡"
                        End If
                    Next
                End If
                If InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ���"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))), "'") = 0 And InStr(1, Trim(.TextMatrix(lngRow, .ColIndex("������"))), "'") = 0 Then
                    strSqls = "Select a.���, a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, b.����ϵ��, b.���ﵥλ, b.�����װ, b.סԺ��λ, b.סԺ��װ, b.ҩ�ⵥλ, b.ҩ���װ," & vbNewLine & _
                             "b.���Ч��, b.סԺ�ɷ����, b.ҩ�����, b.ҩ������, b.�ɱ���, b.��ͬ��λid, b.����ɷ����" & vbNewLine & _
                             "From �շ���ĿĿ¼ A, ҩƷ��� B Where a.Id = b.ҩƷid And a.��� In ('5', '6', '7') and ���=[1] and ����=[2] and ���=[3] and ����" & IIf(Trim(.TextMatrix(lngRow, .ColIndex("������"))) = "", " is null", "='" & Trim(.TextMatrix(lngRow, .ColIndex("������"))) & "'")
                    Set rs���� = zlDatabase.OpenSQLRecord(strSqls, "ҩƷ���", Switch(Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "����ҩ", 5, Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "�г�ҩ", 6, Trim(.TextMatrix(lngRow, .ColIndex("���"))) = "�в�ҩ", 7), Trim(.TextMatrix(lngRow, .ColIndex("Ʒ������"))), Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ���"))))
                    If rs����.RecordCount > 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("ҩƷ���"), lngRow, .ColIndex("ҩƷ���")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt���Ψһ = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "�С�ҩƷ�����"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "Ψһ�Դ���"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "������Ŀ�����ݿ������������Ƿ��ͻ�����" & Trim(.TextMatrix(lngRow, .ColIndex("ҩƷ���"))) & "���Ѵ��ڣ�"
                    End If
                End If
            End If
        Next
    End With
    
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt���뷽ʽ = 0 Then
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    Else
                        cbrControl.Enabled = True
                    End If
                Next
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    
    If vsfList.Rows > 1 Then
        vsfList.Row = 1: vsfList.Col = 1
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub setVSF()
    '�п�ȶ��뷽ʽ����
    Dim cbrControl As CommandBarControl
    Dim lngRow     As Long
    Dim lngCol     As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            Select Case .TextMatrix(0, lngCol)
                Case "�ۼ�", "�ɱ���", "�ۼۻ���ϵ��", "���ﻻ��ϵ��", "סԺ����ϵ��", "ҩ�⻻��ϵ��", "Ч��(��)"
                    .ColAlignment(lngCol) = flexAlignRightCenter
                Case Else
                    .ColAlignment(lngCol) = flexAlignLeftCenter
            End Select
            .ColComboList(lngCol) = ""
        Next
        .FixedRows = 1
        .FixedCols = 1
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '����
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True '�Ӵ�
        .ExplorerBar = flexExNone   '�в�֧��������϶�
        .ColWidth(-1) = 2000
        .ColWidth(0) = 300
        
        If .Rows > 1 Then
            .Editable = flexEDKbdMouse
            If TabControl.Selected.Caption = "����" Then
                .ColComboList(.ColIndex("���")) = "����ҩ|�г�ҩ|�в�ҩ"
            Else
                .ColComboList(.ColIndex("���")) = "����ҩ|�г�ҩ|�в�ҩ"
                .ColComboList(.ColIndex("�Ƿ���")) = " |��"
                .ColComboList(.ColIndex("סԺ�ɷ����")) = "0-���Է���|1-���ɷ���|2-һ����ʹ��|3-�����һ������Ч|4-�������������Ч|5-�������������Ч"
                .ColComboList(.ColIndex("����ɷ����")) = "0-���Է���|1-���ɷ���|2-һ����ʹ��|3-�����һ������Ч|4-�������������Ч|5-�������������Ч"
                If GetColumnPostation("�������") > 0 Then .ColComboList(.ColIndex("�������")) = "0-�������ڲ���|1-����|2-סԺ|3-�����סԺ"
                If GetColumnPostation("ҩ�����") > 0 Then .ColComboList(.ColIndex("ҩ�����")) = " |��"
                If GetColumnPostation("ҩ������") > 0 Then .ColComboList(.ColIndex("ҩ������")) = " |��"
            End If
        End If
    End With
End Sub

Private Sub Form_Resize()
    '�ؼ�λ�ÿ���
    On Error Resume Next
    
    lblFile.Move 110, 600
    txtFile.Move lblFile.Left + lblFile.Width + 20, lblFile.Top - 40, Me.ScaleWidth - (cmdFile.Width + txtFile.Left) - 50
    cmdFile.Move txtFile.Left + txtFile.Width + 20, txtFile.Top - 30
    
    TabControl.Move lblFile.Left - 40, txtFile.Top + txtFile.Height + 50, Me.ScaleWidth - lblFile.Left - 20, ((Me.ScaleHeight - TabControl.Top) / 5) * 3 - 20
    picSplit.Move lblFile.Left, TabControl.Top + TabControl.Height + 20, Me.ScaleWidth - lblFile.Left * 2
    lblCollect.Width = picSplit.Width
    
    vsfError.Move lblFile.Left - 40, picSplit.Top + picSplit.Height + 50, Me.ScaleWidth - lblFile.Left - 20, Me.ScaleHeight - picSplit.Top - picSplit.Height - 120
    
    vsfError.ColWidth(3) = 10000
    If vsfError.Width > 14500 Then
        vsfError.ColWidth(3) = vsfError.Width - 5000
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    Set mobjXLS = Nothing
    mstrType = ""
    mstrMedi = ""
    mstrTypeMsg = ""
    mstrMediMsg = ""
End Sub

Private Sub pic_Resize()
    On Error Resume Next
    vsfList.Move 0, 0, pic.ScaleWidth, pic.ScaleHeight
End Sub

Private Sub lblCollect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y + 20
    End With

    With TabControl
        .Height = picSplit.Top - .Top - 20
    End With
    
    With vsfError
        .Top = picSplit.Top + picSplit.Height + 50
        .Height = ScaleHeight - .Top + 50
    End With
    Me.Refresh
End Sub

Private Sub picSplit_Resize()
    On Error Resume Next
    lblCollect.Move 10, , picSplit.ScaleWidth - 10
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
        Case 0
            If mstrTypeMsg <> "" Then
                Call SetColumns("����")
                If vsfList.Rows > 1 Then Call CheckKind
            End If
        Case 1
            If mstrMediMsg <> "" Then
                Call SetColumns("��ϸ")
                If vsfList.Rows > 1 Then Call CheckƷ��: Call Check���
            End If
    End Select
End Sub

Private Function InitTabControl()
    '��ʼ����ҳ�ؼ�
    With TabControl
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPageSelected
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With
        
        .InsertItem 1, "����", pic.hwnd, 101
        .InsertItem 2, "��ϸ", pic.hwnd, 102
        .Item(0).Selected = True
    End With
End Function

Private Sub SaveType()
'�����������
    Dim lngItemId As Long
    Dim int���   As Integer
    Dim int�ϼ�ID As Long
    Dim str����   As String
    Dim str����   As String
    Dim strSql   As String
    Dim strTemp  As String
    Dim rsTemp   As Recordset
    Dim intCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim arrSql As Variant
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHandle
    arrSql = Array()
    
    Call FS.ShowFlash("���ڱ�������,���Ժ� ...", Me)
    Me.MousePointer = vbHourglass
    With vsfList
        For i = 1 To .Rows - 1
            '���Ǵ����У�������
            For j = 0 To .Cols - 1
                If .Cell(flexcpForeColor, i, j, i, j) = vbRed Then
                    GoTo ErrHand
                End If
            Next
            
            'ID
            If mintType = 2 Then
                lngItemId = sys.NextId("���Ʒ���Ŀ¼")
            End If
            
            '���
            If .TextMatrix(i, .ColIndex("���")) = "����ҩ" Then
                int��� = 1
            ElseIf .TextMatrix(i, .ColIndex("���")) = "�г�ҩ" Then
                int��� = 2
            ElseIf .TextMatrix(i, .ColIndex("���")) = "�в�ҩ" Then
                int��� = 3
            End If
            
            '�ϼ�id
            int�ϼ�ID = GetTypeID(.TextMatrix(i, .ColIndex("�ϼ�����")), .TextMatrix(i, .ColIndex("���")))
            
            '����
            str���� = .TextMatrix(i, .ColIndex("����"))
            
            '����
            str���� = .TextMatrix(i, .ColIndex("����"))
            
            strSql = "zl_���Ʒ���Ŀ¼_insert("
            strSql = strSql & lngItemId & ","
            strSql = strSql & int�ϼ�ID & ","
            strSql = strSql & "'" & str���� & "',"
            strSql = strSql & "'" & str���� & "',"
            strSql = strSql & "'" & zlStr.GetCodeByVB(str����) & "',"
            strSql = strSql & int��� & ","
            strSql = strSql & "0)"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            intCount = intCount + 1
            .TextMatrix(i, 0) = "��"
ErrHand:
        Next
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveType")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Me.MousePointer = vbDefault
    Call FS.StopFlash
    If intCount = vsfList.Rows - 1 And intCount <> 0 Then
        MsgBox "�ɹ����桾���ࡿҳ�������ݣ�", vbInformation, gstrSysName
    ElseIf intCount <> 0 Then
        MsgBox "�ɹ����桾���ࡿҳ" & intCount & "�����ݣ�", vbInformation, gstrSysName
    Else
        MsgBox "�����ࡿҳû�кϸ����ݣ�����ʧ�ܣ�", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveMedi()
'������ϸ����
    Dim rs������Ŀ As Recordset
    Dim rs��Ӧ��   As Recordset
    Dim intCount  As Integer
    Dim strSql    As String
    Dim lngҩ��id As Long
    Dim lngҩƷID As Long
    Dim int���   As Integer
    Dim int����ID As Integer
    Dim str����   As String
    Dim int����   As Integer
    Dim i As Integer
    Dim j As Integer
    Dim str���� As String
    Dim rsTemp As Recordset
    Dim arrSql As Variant
    Dim strTemp As String
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHandle
    arrSql = Array()
    Call FS.ShowFlash("���ڱ�������,���Ժ� ...", Me)
    Me.MousePointer = vbHourglass
    
    Set rs������Ŀ = zlDatabase.OpenSQLRecord("Select ID,����,���� From ������Ŀ Where ĩ�� = 1", "������Ŀ")
    Set rs��Ӧ�� = zlDatabase.OpenSQLRecord("Select ID,����,����,���֤��,���֤Ч�� From ��Ӧ��", "��Ӧ��")
    
    '��ȡƷ�ֱ���
    strSql = "Select ���,id,���� From ������ĿĿ¼ Where ��� In ('5','6','7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
    str���� = ""
    Do While Not rsTemp.EOF
        '��ʽ�����,ID[����]���,ID[����]...
        str���� = str���� & rsTemp!��� & "," & rsTemp!ID & "[" & rsTemp!���� & "]"
        rsTemp.MoveNext
    Loop
    
    With vsfList
        For i = 1 To .Rows - 1
            '���Ǵ����У�������
            For j = 0 To .Cols - 1
                If .Cell(flexcpForeColor, i, j, i, j) = vbRed Then
                    GoTo ErrHand
                End If
            Next
            
            If InStr(1, str����, "[" & .TextMatrix(i, .ColIndex("Ʒ�ֱ���")) & "]") <= 0 Then
                '����Ʒ��
                '���
                If .TextMatrix(i, .ColIndex("���")) = "����ҩ" Then
                    int��� = 1
                    strTemp = "5"
                    strSql = "zl_��ҩƷ��_Insert('5',"
                ElseIf .TextMatrix(i, .ColIndex("���")) = "�г�ҩ" Then
                    int��� = 2
                    strTemp = "6"
                    strSql = "zl_��ҩƷ��_Insert('6',"
                ElseIf .TextMatrix(i, .ColIndex("���")) = "�в�ҩ" Then
                    int��� = 3
                    strTemp = "7"
                    strSql = "zl_��ҩƷ��_Insert('7',"
                End If
                '����ID
                int����ID = GetTypeID(.TextMatrix(i, .ColIndex("����")), .TextMatrix(i, .ColIndex("���")))
                strSql = strSql & int����ID & ","
                'ҩ��ID
                lngҩ��id = sys.NextId("������ĿĿ¼")
                str���� = str���� & strTemp & "," & lngҩ��id & "[" & .TextMatrix(i, .ColIndex("Ʒ�ֱ���")) & "]"
                strSql = strSql & lngҩ��id & ","
                'Ʒ�ֱ���
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("Ʒ�ֱ���")) & "',"
                'Ʒ������
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("Ʒ������")) & "',"
                'ƴ��
                str���� = .TextMatrix(i, .ColIndex("Ʒ������"))
                strSql = strSql & "'" & zlStr.GetCodeByORCL(str����) & "',"
                '���
                strSql = strSql & "'" & zlStr.GetCodeByORCL(str����, True) & "',"
                'Ӣ��
                strSql = strSql & "'',"
                '������λ
                strSql = strSql & "'" & .TextMatrix(i, .ColIndex("������λ")) & "',"
                '����
                strSql = strSql & IIf(strTemp = "7", "", "'" & .TextMatrix(i, .ColIndex("����")) & "',")
                '�������,��ֵ����,��Դ���,��ҩ�ݴ�
                strSql = strSql & "'��ͨҩ','�ռ�','����','��ѡ')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            Else
                '�����ʽ:���,ID[����]...
                '���ر����Ӧ������ID:���,ID
                strTemp = Mid(Mid(str����, 1, InStr(1, str����, "[" & .TextMatrix(i, .ColIndex("Ʒ�ֱ���")) & "]") - 1), InStrRev(Mid(str����, 1, InStr(1, str����, "[" & .TextMatrix(i, .ColIndex("Ʒ�ֱ���")) & "]") - 1), "]") + 1)
                
                '����ID
                lngҩ��id = Split(strTemp, ",")(1)
                '�������
                strTemp = Split(strTemp, ",")(0)
            End If
            
            '������
            strSql = ""
            '���
            If strTemp = "5" Then
                strSql = "zl_��ҩ���_Insert("
            ElseIf strTemp = "6" Then
                strSql = "zl_��ҩ���_Insert("
            Else
                strSql = "zl_��ҩ���_Insert("
            End If
            'ҩ��ID
            strSql = strSql & lngҩ��id & ","
            'ҩƷID
            lngҩƷID = sys.NextId("�շ���ĿĿ¼")
            strSql = strSql & lngҩƷID & ","
            '����
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("������")) & "',"
            '���
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("ҩƷ���")) & "',"
            '������
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("������")) & "',"
            '��Ʒ��,ƴ������,��ʼ���,������,��ʶ��,ҩƷ��Դ,��ע�ĺ�,ע���̱�
            strSql = strSql & "'','','','','','','','',"
            '�ۼ۵�λ
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("�ۼ۵�λ")) & "',"
            '����ϵ��
            strSql = strSql & .TextMatrix(i, .ColIndex("�ۼۻ���ϵ��")) & ","
            '���ﵥλ
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("���ﵥλ")) & "',"
            '����ϵ��
            strSql = strSql & .TextMatrix(i, .ColIndex("���ﻻ��ϵ��")) & ","
            'סԺ��λ
            strSql = strSql & IIf(strTemp = "7", "", "'" & .TextMatrix(i, .ColIndex("סԺ��λ")) & "',")
            'סԺϵ��
            strSql = strSql & IIf(strTemp = "7", "", .TextMatrix(i, .ColIndex("סԺ����ϵ��")) & ",")
            'ҩ�ⵥλ
            strSql = strSql & "'" & .TextMatrix(i, .ColIndex("ҩ�ⵥλ")) & "',"
            'ҩ��ϵ��
            strSql = strSql & .TextMatrix(i, .ColIndex("ҩ�⻻��ϵ��")) & ","
            '���쵥λ,���췧ֵ
            strSql = strSql & "1,null,"
            '�Ƿ���
            If .TextMatrix(i, .ColIndex("�Ƿ���")) = "" Then
                strSql = strSql & "0,"
            ElseIf .TextMatrix(i, .ColIndex("�Ƿ���")) = "��" Then
                strSql = strSql & "1,"
            End If
            'ָ�������ۣ��ɱ���
            strSql = strSql & .TextMatrix(i, .ColIndex("�ɱ���")) & ","
            '����
            strSql = strSql & "100,"
            'ָ�����ۼ�
            strSql = strSql & .TextMatrix(i, .ColIndex("�ۼ�")) & ","
            '�ӳ���
            If .TextMatrix(i, .ColIndex("�ɱ���")) = 0 Then
                strSql = strSql & "100,"
            Else
                strSql = strSql & (Val(.TextMatrix(i, .ColIndex("�ۼ�"))) / Val(.TextMatrix(i, .ColIndex("�ɱ���"))) - 1) * 100 & ","
                If (Val(.TextMatrix(i, .ColIndex("�ۼ�"))) / Val(.TextMatrix(i, .ColIndex("�ɱ���"))) - 1) * 100 > 100 Then
                    strSql = strSql & "100,"
                End If
            End If
            '����ѱ���,ҩ�ۼ���,��������
            strSql = strSql & "null,'','',"
            '�������
            If GetColumnPostation("�������") > 0 Then
                Select Case .TextMatrix(i, .ColIndex("�������"))
                    Case "0-�������ڲ���", ""
                        strSql = strSql & "0,"
                    Case "1-����"
                        strSql = strSql & "1,"
                    Case "2-סԺ"
                        strSql = strSql & "2,"
                    Case "3-�����סԺ"
                        strSql = strSql & "3,"
                End Select
            Else
                strSql = strSql & "3,"
            End If
            'Gmp��֤,�б�ҩƷ,���ηѱ�,
            strSql = strSql & "0,0,0,"
            'סԺ�ɷ����
            Select Case .TextMatrix(i, .ColIndex("סԺ�ɷ����"))
                Case "0-���Է���", ""
                    strSql = strSql & "0,"
                Case "1-���ɷ���"
                    strSql = strSql & "1,"
                Case "2-һ����ʹ��"
                    strSql = strSql & "2,"
                Case "3-�����һ������Ч"
                    strSql = strSql & "3,"
                Case "4-�������������Ч"
                    strSql = strSql & "4,"
                Case "5-�������������Ч"
                    strSql = strSql & "5,"
            End Select
            'ҩ�����
            If GetColumnPostation("ҩ�����") > 0 Then
                If .TextMatrix(i, .ColIndex("ҩ�����")) = "" Then
                    strSql = strSql & "0,"
                ElseIf .TextMatrix(i, .ColIndex("ҩ�����")) = "��" Then
                    strSql = strSql & "1,"
                End If
            Else
                strSql = strSql & "0,"
            End If
            'ҩ������
            If GetColumnPostation("ҩ������") > 0 Then
                If .TextMatrix(i, .ColIndex("ҩ������")) = "" Then
                    strSql = strSql & "0,"
                ElseIf .TextMatrix(i, .ColIndex("ҩ������")) = "��" Then
                    strSql = strSql & "1,"
                End If
            Else
                strSql = strSql & "0,"
            End If
            'Ч��(��)
            If GetColumnPostation("Ч��(��)") > 0 Then
                If .TextMatrix(i, .ColIndex("Ч��(��)")) = "" Then
                    strSql = strSql & "0,"
                Else
                    strSql = strSql & .TextMatrix(i, .ColIndex("Ч��(��)")) & ","
                End If
            Else
                strSql = strSql & "0,"
            End If
            '���������
            strSql = strSql & "0,"
            '�ɱ���
            strSql = strSql & .TextMatrix(i, .ColIndex("�ɱ���")) & ","
            '�ۼ�
            strSql = strSql & .TextMatrix(i, .ColIndex("�ۼ�")) & ","
            '������ĿID
            rs������Ŀ.Filter = ""
            rs������Ŀ.Filter = "����='" & .TextMatrix(i, .ColIndex("������Ŀ")) & "'"
            strSql = strSql & rs������Ŀ!ID & ","
            '��Ӧ�����ƣ���ͬ��λid��
            If GetColumnPostation("��Ӧ������") > 0 Then
                If .TextMatrix(i, .ColIndex("��Ӧ������")) <> "" Then
                    rs��Ӧ��.Filter = ""
                    rs��Ӧ��.Filter = "����='" & .TextMatrix(i, .ColIndex("��Ӧ������")) & "'"
                    strSql = strSql & rs��Ӧ��!ID & ","
                Else
                    strSql = strSql & "null,"
                End If
            Else
                strSql = strSql & "null,"
            End If
            
            If strTemp = "7" Then
                '˵��,��̬����,��ҩ����,��ѡ��,��ֵ˰��,����ҩ��,��ҩ��̬,վ��,�Ƿ񳣱�,������Ŀ
                strSql = strSql & "Null,0,Null,Null,Null,Null,Null,Null,Null,Null,"
            Else
                '˵��,��̬����,��ҩ����,��ѡ��,��ֵ˰��,����ҩ��,վ��,�Ƿ񳣱�,�洢�¶�,�洢����,��ҩ����,�Ƿ�������,����,������Ŀ
                strSql = strSql & "Null,0,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,"
            End If
            '����ɷ����
            Select Case .TextMatrix(i, .ColIndex("����ɷ����"))
                Case "0-���Է���", ""
                    strSql = strSql & "0)"
                Case "1-���ɷ���"
                    strSql = strSql & "1)"
                Case "2-һ����ʹ��"
                    strSql = strSql & "2)"
                Case "3-�����һ������Ч"
                    strSql = strSql & "3)"
                Case "4-�������������Ч"
                    strSql = strSql & "4)"
                Case "5-�������������Ч"
                    strSql = strSql & "5)"
            End Select
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            intCount = intCount + 1
            vsfList.TextMatrix(i, 0) = "��"
ErrHand:
        Next
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Me.MousePointer = vbDefault
    Call FS.StopFlash
    If intCount = vsfList.Rows - 1 And intCount <> 0 Then
        MsgBox "�ɹ����桾��ϸ��ҳ�������ݣ�", vbInformation, gstrSysName
    ElseIf intCount <> 0 Then
        MsgBox "�ɹ����桾��ϸ��ҳ" & intCount & "�����ݣ�", vbInformation, gstrSysName
    Else
        MsgBox "����ϸ��ҳû�кϸ����ݣ�����ʧ�ܣ�", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfError_EnterCell()
    Dim strTemp As String
    Dim lngRow  As Long
    Dim lngCol  As Long
    Dim strCol  As String
    
    With vsfError
        If .Row = 0 Then Exit Sub
        .FocusRect = flexFocusSolid
        If InStr(1, .TextMatrix(.Row, 1), "��") = 0 Then Exit Sub
        If .TextMatrix(.Row, 1) <> "" Then
            strTemp = .TextMatrix(.Row, 1)
            lngRow = Mid(strTemp, 1, InStr(1, strTemp, "��") - 1)
            strCol = Mid(strTemp, InStr(1, strTemp, "��") + 1, InStr(1, strTemp, "��") - InStr(1, strTemp, "��") - 1)
            lngCol = vsfList.ColIndex(strCol)
            If lngRow > vsfList.Rows - 1 Then MsgBox "���������Ѿ���ɾ���ˣ�", vbInformation, gstrSysName: Exit Sub
            vsfList.Row = lngRow
            vsfList.Col = lngCol
            vsfList.ShowCell lngRow, lngCol
        End If
    End With
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '��¼���µĽ�������
    If vsfList.Tag = "1" Then Exit Sub
    If TabControl.Selected.Caption = "����" Then
        Call GetColumns("����")
    Else
        Call GetColumns("��ϸ")
    End If
End Sub

Private Sub vsfList_ChangeEdit()
    Dim cbrControl As CommandBarControl
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = False
'    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
'    cbrControl.Enabled = False
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        If .Rows > 1 And .Row > 0 Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = LenB(StrConv(.EditText, vbFromUnicode))
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim strTemp As String
    Dim intRow  As Integer
    Dim i As Integer
    With vsfList
        If .Row < 1 Then Exit Sub
        strTemp = .Row & "�С�" & .TextMatrix(0, .Col) & "����"
        .FocusRect = flexFocusSolid
    End With
    
    With vsfError
        If .Rows < 2 Then Exit Sub
        i = 0
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) = strTemp Then
                .Row = intRow
                .TopRow = intRow
                i = 1
                Exit For
            End If
        Next
        If i = 0 Then .Row = 0
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cbrControl As CommandBarControl
    
    With vsfList
        If KeyCode = vbKeyDelete Then
            If .Row < 1 Then Exit Sub
            If MsgBox("��ɾ����" & .Row & "�����ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                With vsfList
                    .RemoveItem .Row
                    '��¼���µĽ�������
                    If TabControl.Selected.Caption = "����" Then
                        Call GetColumns("����")
                    Else
                        Call GetColumns("��ϸ")
                    End If
                    If .Rows <= 1 Then
                        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
                        cbrControl.Enabled = False
                        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
                        cbrControl.Enabled = False
                        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
                        cbrControl.Enabled = False
                    End If
                End With
            End If
        End If
        
        If KeyCode = vbKeyReturn And .Rows > 1 Then
            If .Col = .Cols - 1 Then
                If .Row = .Rows - 1 Then .Rows = .Rows + 1: .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Row + 1
                .Col = 1
            Else
                .Col = .Col + 1
            End If
        End If
    End With
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfList
        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then
            Exit Sub
        End If
        Select Case .Col
            Case .ColIndex("�ۼۻ���ϵ��"), .ColIndex("���ﻻ��ϵ��"), .ColIndex("סԺ����ϵ��"), .ColIndex("ҩ�⻻��ϵ��"), .ColIndex("Ч��(��)")
                If InStr(1, "1234567890", Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("����"), .ColIndex("Ʒ�ֱ���"), .ColIndex("������")
                If InStr(1, "1234567890abcdefghijklmnopqrstuvwxyz", LCase(Chr(KeyAscii))) = 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("�ɱ���"), .ColIndex("�ۼ�")
                If InStr(1, "1234567890.", Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
                If Chr(KeyAscii) = "." And InStr(1, .EditText, ".") > 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("��Ӧ�����֤Ч��")
                If InStr(1, "1234567890.-/", Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case .ColIndex("����"), .ColIndex("Ʒ������"), .ColIndex("ҩƷ���"), .ColIndex("������"), .ColIndex("����"), .ColIndex("������λ"), .ColIndex("�ۼ۵�λ"), .ColIndex("���ﵥλ"), .ColIndex("סԺ��λ"), .ColIndex("ҩ�ⵥλ"), .ColIndex("������Ŀ"), .ColIndex("��Ӧ������"), .ColIndex("��Ӧ�����֤��")
                If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                End If
        End Select
    End With
End Sub

Private Sub vsfList_RowColChange()
    If vsfList.Rows > 1 Then Call SetNote
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim cbrControl As CommandBarControl
    Dim strSql     As String
    Dim rsTemp     As Recordset
    
    If mbyt���뷽ʽ = 1 And vsfError.Rows > 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    With vsfList
        Select Case .Col
            Case .ColIndex("�ɱ���"), .ColIndex("�ۼ�")
                strSql = "select ���� from ҩƷ���ľ��� where ���=1 and ����=" & IIf(.Col = .ColIndex("�ɱ���"), 1, 2) & " and ��λ=1"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "")
                .EditText = zlStr.FormatEx(.EditText, Val(rsTemp!����), , True)
            Case .ColIndex("�ۼۻ���ϵ��"), .ColIndex("���ﻻ��ϵ��"), .ColIndex("סԺ����ϵ��"), .ColIndex("ҩ�⻻��ϵ��"), .ColIndex("Ч��(��)")
                .EditText = Val(.EditText)
            Case .ColIndex("��Ӧ�����֤Ч��")
                If IsNumeric(.EditText) Then
                    .EditText = TranNumToDate(.EditText)
                Else
                    .EditText = FormatDate(.EditText)
                End If
            Case .ColIndex("����"), .ColIndex("Ʒ�ֱ���"), .ColIndex("������")
                .EditText = UCase(Trim(.EditText))
        End Select
        .EditText = Trim(.EditText)
    End With
End Sub

Private Sub SetCols()
    Dim strMediColumn As String
    
    Select Case mintType
    Case 1  '�շ�Ŀ¼����
    Case 2   'ҩƷĿ¼����
        strMediColumn = zlDatabase.GetPara("�е���ʾ����", glngSys, mlngModule, MSTRMEDICAL)
    Case 3   '������Ŀ����
    End Select
    
    If Not frmImportFileCols.ShowMe(Me, strMediColumn) Then Exit Sub
    If MsgBox("�����ñ��������Ŀ���Զ��رգ����´δ򿪲Ż���Ч��" & vbCrLf & "�Ƿ�������棿", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
 
    Call zlDatabase.SetPara("�е���ʾ����", strMediColumn, glngSys, mlngModule)
    Call GetColumnHead
    
    Unload Me
End Sub

Public Sub ShowMe(ByVal intType As Integer, ByVal frmParent As Form)
    Call InitExcel
    
    If mobjXLS Is Nothing Then
        err.Clear
        Exit Sub
    End If
    
    mobjXLS.DisplayAlerts = False
    
    mlngModule = glngModul
    mintType = intType
    
    Me.Show 1, frmParent
End Sub

Public Sub GetColumnHead()
'��¼����ʾ�ķ������ϸ����ͷ��Ϣ
    Dim arrTypeColumn As Variant
    Dim arrMediColumn As Variant
    Dim strMedical    As String
    Dim intCol As Integer
    Dim intNum As Integer
    
    mstrType = ""
    mstrMedi = ""
    strMedical = zlDatabase.GetPara("�е���ʾ����", glngSys, mlngModule)
    '����
    arrTypeColumn = Split(Split(strMedical, "||")(0) & "|", "|")
    Do While arrTypeColumn(intCol) <> ""
        If Split(arrTypeColumn(intCol), ",")(2) = 0 Then
            mstrType = mstrType & Split(arrTypeColumn(intCol), ",")(0) & "|"
        End If
        intCol = intCol + 1
    Loop
    '��ϸ
    arrMediColumn = Split(Split(strMedical, "||")(1), "|")
    Do While arrMediColumn(intNum) <> ""
        If Split(arrMediColumn(intNum), ",")(2) = 0 Then
            mstrMedi = mstrMedi & Split(arrMediColumn(intNum), ",")(0) & "|"
        End If
        intNum = intNum + 1
    Loop
End Sub

Private Sub GetColumns(ByVal strType As String)
'��¼�±���з������ϸ��������Ϣ
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        If strType = "����" Then
            mstrTypeMsg = ""
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    mstrTypeMsg = mstrTypeMsg & Trim(.TextMatrix(lngRow, lngCol)) & ";"
                Next
                mstrTypeMsg = mstrTypeMsg & "|"
            Next
        ElseIf strType = "��ϸ" Then
            mstrMediMsg = ""
            For lngRow = 0 To .Rows - 1
                For lngCol = 1 To .Cols - 1
                    mstrMediMsg = mstrMediMsg & Trim(.TextMatrix(lngRow, lngCol)) & ";"
                Next
                mstrMediMsg = mstrMediMsg & "|"
            Next
        End If
    End With
End Sub

Private Function GetTypeID(ByVal strVal As String, ByVal strType As String, Optional ByVal strKind As String) As Long
'��ȡ����ID
    Dim strSql As String
    Dim intType As Integer
    Dim strSecType As String
    Dim rsTemp As Recordset

    On Error GoTo ErrHand
    
    If strType = "����ҩ" Then
        intType = 1
    ElseIf strType = "�г�ҩ" Then
        intType = 2
    ElseIf strType = "�в�ҩ" Then
        intType = 3
    Else
        intType = 0
    End If
    
    '�������Ƿ�ֻ��һ��
    If InStr(1, strVal, "\") > 1 Then
        strType = Mid(strVal, InStrRev(strVal, "\") + 1)
        strSecType = Mid(strVal, 1, InStrRev(strVal, "\") - 1)
    Else
        strType = strVal
        strSecType = ""
    End If

    If strSecType = "" And InStr(1, strVal, "\") = 0 Then
        '����ֻ��һ�������
        strSql = "Select ID,���� " & _
                 "From ���Ʒ���Ŀ¼ " & _
                 "Where ���� = [1] And ����=[2] order by ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetTypeID", strType, intType)
    Else
        strSql = "Select ID,���� From ���Ʒ���Ŀ¼" & vbNewLine & _
                "Where ���� = [1] And �ϼ�id in (Select ID From ���Ʒ���Ŀ¼ Where ���� = [2] And ����=[3] ) order by ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetTypeID", strType, strSecType, intType)
    End If

    If rsTemp.RecordCount > 0 Then
        GetTypeID = rsTemp!ID
    ElseIf rsTemp.RecordCount = 0 Then
        GetTypeID = 0
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetNote()
'������˵��
    Dim arrComent As Variant
    
    With vsfList
        If TabControl.Selected.Caption = "����" Then
            arrComent = Split(Split(MSTRCOMMENT, "||")(0), "|")
            Select Case .Col
                Case .ColIndex("���")
                    lblCollect.Caption = Split(arrComent(0), ";")(1)
                Case .ColIndex("�ϼ�����")
                    lblCollect.Caption = Split(arrComent(1), ";")(1)
                Case .ColIndex("����")
                    lblCollect.Caption = Split(arrComent(2), ";")(1)
                Case .ColIndex("����")
                    lblCollect.Caption = Split(arrComent(3), ";")(1)
                Case Else
                    lblCollect.Caption = ""
            End Select
        Else
            arrComent = Split(Split(MSTRCOMMENT, "||")(1), "|")
            Select Case .Col
                Case .ColIndex("���")
                    lblCollect.Caption = Split(arrComent(0), ";")(1)
                Case .ColIndex("����")
                    lblCollect.Caption = Split(arrComent(1), ";")(1)
                Case .ColIndex("Ʒ�ֱ���")
                    lblCollect.Caption = Split(arrComent(2), ";")(1)
                Case .ColIndex("Ʒ������")
                    lblCollect.Caption = Split(arrComent(3), ";")(1)
                Case .ColIndex("������")
                    lblCollect.Caption = Split(arrComent(4), ";")(1)
                Case .ColIndex("ҩƷ���")
                    lblCollect.Caption = Split(arrComent(5), ";")(1)
                Case .ColIndex("������")
                    lblCollect.Caption = Split(arrComent(6), ";")(1)
                Case .ColIndex("����")
                    lblCollect.Caption = Split(arrComent(7), ";")(1)
                Case .ColIndex("������λ"), .ColIndex("�ۼ۵�λ"), .ColIndex("���ﵥλ"), .ColIndex("סԺ��λ"), .ColIndex("ҩ�ⵥλ")
                    lblCollect.Caption = Split(arrComent(8), ";")(1)
                Case .ColIndex("�ۼۻ���ϵ��"), .ColIndex("���ﻻ��ϵ��"), .ColIndex("סԺ����ϵ��"), .ColIndex("ҩ�⻻��ϵ��")
                    lblCollect.Caption = Split(arrComent(10), ";")(1)
                Case .ColIndex("�Ƿ���")
                    lblCollect.Caption = Split(arrComent(17), ";")(1)
                Case .ColIndex("�ɱ���"), .ColIndex("�ۼ�")
                    lblCollect.Caption = Split(arrComent(18), ";")(1)
                Case .ColIndex("������Ŀ")
                    lblCollect.Caption = Split(arrComent(20), ";")(1)
                Case .ColIndex("סԺ�ɷ����"), .ColIndex("����ɷ����")
                    lblCollect.Caption = Split(arrComent(21), ";")(1)
                Case .ColIndex("�������")
                    lblCollect.Caption = Split(arrComent(23), ";")(1)
                Case .ColIndex("ҩ�����"), .ColIndex("ҩ������")
                    lblCollect.Caption = Split(arrComent(24), ";")(1) & "," & Split(arrComent(25), ";")(1)
                Case .ColIndex("Ч��(��)")
                    lblCollect.Caption = Split(arrComent(26), ";")(1)
                Case .ColIndex("��Ӧ������")
                    lblCollect.Caption = Split(arrComent(27), ";")(1)
                Case .ColIndex("��Ӧ�����֤��")
                    lblCollect.Caption = Split(arrComent(28), ";")(1)
                Case .ColIndex("��Ӧ�����֤Ч��")
                    lblCollect.Caption = Split(arrComent(29), ";")(1)
                Case Else
                    lblCollect.Caption = ""
            End Select
        End If
        lblCollect.ForeColor = &HFF0000
    End With
End Sub
