VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomQueryCfg 
   Caption         =   "��ѯ��������"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomQueryCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5400
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSScriptControlCtl.ScriptControl sctExecute 
      Left            =   4035
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   7560
      Left            =   255
      ScaleHeight     =   7560
      ScaleWidth      =   11415
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   11415
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   7560
         Left            =   6885
         TabIndex        =   3
         Top             =   0
         Width           =   85
         _ExtentX        =   159
         _ExtentY        =   13335
         SplitWidth      =   85
         SplitLevel      =   3
         Con1MinSize     =   3000
         Con2MinSize     =   5000
         Control1Name    =   "picScheme"
         Control2Name    =   "picSchemeCfg"
      End
      Begin VB.PictureBox picScheme 
         BorderStyle     =   0  'None
         Height          =   7560
         Left            =   0
         ScaleHeight     =   7560
         ScaleWidth      =   6885
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   0
         Width           =   6885
         Begin VB.ComboBox cbxDepart 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   120
            Width           =   2385
         End
         Begin zl9PACSWork.ucFlexGrid ufgScheme 
            Height          =   7290
            Left            =   105
            TabIndex        =   19
            Top             =   510
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   12859
            DefaultCols     =   ""
            DisCellColor    =   16777215
            HeadCheckValue  =   1
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontSize    =   10.5
            HeadFontCharset =   134
            HeadFontWeight  =   400
            HeadColor       =   0
            DataFontSize    =   10.5
            DataFontCharset =   134
            DataFontWeight  =   400
            DataColor       =   -2147483640
         End
         Begin VB.Label labDepart 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   180
            Width           =   885
         End
      End
      Begin VB.PictureBox picSchemeCfg 
         BorderStyle     =   0  'None
         Height          =   7560
         Left            =   6970
         ScaleHeight     =   7560
         ScaleWidth      =   4440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   4445
         Begin VB.Frame framSql 
            Caption         =   "������乹��"
            Height          =   2415
            Left            =   255
            TabIndex        =   6
            Top             =   4980
            Width           =   6480
            Begin VB.CommandButton cmdInsertPar 
               Caption         =   "�������(&I)"
               Height          =   375
               Left            =   5085
               TabIndex        =   12
               Top             =   1935
               Width           =   1305
            End
            Begin VB.TextBox txtFilterSql 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1650
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   240
               Width           =   6345
            End
         End
         Begin VB.Frame framInput 
            Caption         =   "¼��������"
            Height          =   3225
            Left            =   165
            TabIndex        =   5
            Top             =   1305
            Width           =   6585
            Begin VB.CommandButton cmdMoveNext 
               Caption         =   "������(&E)"
               Height          =   375
               Left            =   5415
               TabIndex        =   17
               Top             =   2745
               Width           =   1095
            End
            Begin VB.CommandButton cmdMoveLast 
               Caption         =   "������(&L)"
               Height          =   375
               Left            =   4335
               TabIndex        =   16
               Top             =   2745
               Width           =   1095
            End
            Begin VB.CommandButton cmdDelInput 
               Caption         =   "ɾ����(&D)"
               Height          =   375
               Left            =   1185
               TabIndex        =   15
               Top             =   2745
               Width           =   1095
            End
            Begin VB.CommandButton cmdNewInput 
               Caption         =   "������(&N)"
               Height          =   375
               Left            =   105
               TabIndex        =   14
               Top             =   2745
               Width           =   1095
            End
            Begin zl9PACSWork.ucFlexGrid ufgInputCfg 
               Height          =   2430
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   6390
               _ExtentX        =   11271
               _ExtentY        =   4286
               DefaultCols     =   ""
               DisCellColor    =   16777215
               HeadCheckValue  =   1
               IsBtnNextCell   =   0   'False
               IsCopyAdoMode   =   0   'False
               IsEjectConfig   =   -1  'True
               HeadFontSize    =   10.5
               HeadFontCharset =   134
               HeadFontWeight  =   400
               HeadColor       =   0
               DataFontSize    =   10.5
               DataFontCharset =   134
               DataFontWeight  =   400
               DataColor       =   -2147483640
            End
         End
         Begin VB.Frame framBase 
            Caption         =   "������Ϣ����"
            Height          =   795
            Left            =   165
            TabIndex        =   4
            Top             =   165
            Width           =   6570
            Begin VB.TextBox txtSchemeMemo 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4155
               TabIndex        =   9
               Top             =   270
               Width           =   2175
            End
            Begin VB.TextBox txtSchemeName 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   7
               Top             =   270
               Width           =   1935
            End
            Begin VB.Label labSchemeMemo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����˵��:"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3195
               TabIndex        =   10
               Top             =   360
               Width           =   975
            End
            Begin VB.Label labObj 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������:"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   8
               Top             =   360
               Width           =   975
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   8175
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCustomQueryCfg.frx":0AE2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7938
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   480
      Top             =   300
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCustomQueryCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�˵�����ö�ٶ���
Private Enum TMenuType
    mtFile = 1
    mtSave = 2
    mtImport = 3
    mtExport = 4
    mtQuit = 5
    
    mtEdit = 6
    mtNewScheme = 7
    mtModifyScheme = 8
    mtDelScheme = 9
    mtSetDefault = 10
    mtUseScheme = 11
    mtMoveLastScheme = 12
    mtMoveNextScheme = 13
    mtCheckScheme = 14
    mtCancel = 15
    
End Enum

'��ѯ�����ж���
Private Const M_STR_SCHEME_COLS As String = "|Id,hide,key|�������,hide|��ѯ���,hide|��������,hide|��������,read,w2100|�Ƿ�Ĭ��,read,w1000|ʹ��״̬,read,w1000|����˵��,read,w2400|"
Private Const M_STR_SCHEME_CONVERT As String = "|�Ƿ�Ĭ��:0-,1-Ĭ��|ʹ��״̬:0-����,1-����|"

'¼�������ж���
'cbx<,[��ǰ����],[��ǰʱ��],[��ǰ�û�ID],[��ǰ����ID],[��ǰϵͳ���],[��ǰģ����]>
Private Const M_STR_INPUT_COLS As String = "|ID,hide,key|����ID,hide|¼��˳��,hide|¼����Ŀ,w1400|¼������,cbx<0-�ı���,1-���ڿ�,2-ʱ���,3-�����ڿ�,4-������,5-��ѡ��>|" & _
                                            "Ĭ��ֵ,btn,w1200|������Դ,w2400,btn|"
Private Const M_STR_INPUT_CONVERT As String = "|¼������:0-�ı���,1-���ڿ�,2-ʱ���,3-�����ڿ�,4-������,5-��ѡ��|"


Private mblnCurModifyState As Boolean
Public mblnIsChange As Boolean




Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '���ò˵����͹��������
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True                                '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False                            '�����õĲ˵���������
        .UseFadedIcons = False                                  'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True                                 '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True                                '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True                                      '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24                               '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16                              '����Сͼ��ĳߴ�
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '���ÿؼ���ʾ���
        .EnableCustomization False                             '�Ƿ������Զ�������
        Set .Icons = zlCommFun.GetPubIcons                     '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "�ļ�(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "ȡ��(&C)"): cbrControl.IconId = 3565
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtImport, "����(&I)"): cbrControl.IconId = 0: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtExport, "����(&E)"): cbrControl.IconId = 0
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "�༭(&E)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtNewScheme, "����(&N)"): cbrControl.IconId = 4010
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtModifyScheme, "�޸�(&M)"): cbrControl.IconId = 3003
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtDelScheme, "ɾ��(&D)"): cbrControl.IconId = 4008
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSetDefault, "����Ĭ��(&F)"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtUseScheme, "����(&A)"): cbrControl.IconId = 3006
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtMoveLastScheme, "����(&L)"): cbrControl.IconId = 3082: cbrControl.BeginGroup = True
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtMoveNextScheme, "����(&X)"): cbrControl.IconId = 21903
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCheckScheme, "��֤(&V)"): cbrControl.IconId = 3823: cbrControl.BeginGroup = True

    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����", "���淽��"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "ȡ��", "ȡ���޸�"): cbrControl.IconId = 3565
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtNewScheme, "����", "��������"): cbrControl.IconId = 4010: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtModifyScheme, "�޸�", "�޸ķ���"): cbrControl.IconId = 3003
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtDelScheme, "ɾ��", "ɾ������"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSetDefault, "����Ĭ��", "����Ĭ�Ϸ���"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtUseScheme, "����", "���÷���"): cbrControl.IconId = 3006
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtMoveLastScheme, "����", "���Ʒ���"): cbrControl.IconId = 3082: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtMoveNextScheme, "����", "���Ʒ���"): cbrControl.IconId = 21903
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCheckScheme, "��֤", "��֤����"): cbrControl.IconId = 3823: cbrControl.BeginGroup = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�", "�˳�"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
'ִ�н��湦��
Dim strResult As String

On Error GoTo errHandle
    Select Case control.ID
    
        Case TMenuType.mtCancel
            Call CancelScheme   '��������
            
        Case TMenuType.mtNewScheme
            Call NewScheme      '��������
            
        Case TMenuType.mtModifyScheme
            Call ModifyScheme   '�޸ķ���
            
        Case TMenuType.mtDelScheme
            Call DelScheme      'ɾ������
            
        Case TMenuType.mtCheckScheme
            strResult = VerifyScheme   '��֤����
            If strResult <> "" Then
                MsgBoxD Me, strResult, vbOKOnly, Me.Caption
            Else
                MsgBoxD Me, "ͨ����֤��", vbOKOnly, Me.Caption
            End If
            
        Case TMenuType.mtSave
            Call SaveScheme     '���淽��
            
        Case TMenuType.mtSetDefault
            Call DefaultScheme  '����Ĭ�Ϸ���
            
        Case TMenuType.mtUseScheme
            Call UseScheme      '���÷���ʹ��״̬
            
        Case TMenuType.mtMoveLastScheme
            Call MoveLastScheme '���Ʒ���
            
        Case TMenuType.mtMoveNextScheme
            Call MoveNextScheme '���Ʒ���
            
        Case TMenuType.mtQuit
            Call Unload(Me)
            
        Case TMenuType.mtImport
            Call ImportScheme   '���뷽��
            
        Case TMenuType.mtExport
            Call ExportScheme   '��������
            
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(control)
            
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errHandle
    zlMailTo hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'���ܣ����ð�������
On Error GoTo errHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errHandle
    zlHomePage hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    picBack.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ImportScheme()
'���뷽��
    Dim rsData As ADODB.Recordset
    Dim lngOldSchemeId As Long
    Dim lngNewSchemeId As Long
    Dim strSql As String
    Dim strExeSql() As String
    Dim strResult As String
    Dim rsTemp As ADODB.Recordset
    
    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"
    
    dlgFile.Filename = ""
    dlgFile.ShowOpen
    
    If dlgFile.Filename = "" Then Exit Sub
    
    Set rsData = New ADODB.Recordset
    Call rsData.Open(dlgFile.Filename)
    
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "û�п����ڵ�������ݣ������ļ��Ƿ���ȷ��", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    rsData.Sort = "id"
    
    lngOldSchemeId = 0
    rsData.MoveFirst
    ReDim Preserve strExeSql(1)
    
    While Not rsData.EOF
        If lngOldSchemeId <> Val(Nvl(rsData!ID)) Then
            '����Ӱ���ѯ������¼
            
            strSql = "select Ӱ���ѯ����_ID.NextVal as ID from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�·���ID")
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD Me, "���ܻ�ȡ�����ķ���ID��ϵͳ���˳����档", vbOKOnly, Me.Caption
                Exit Sub
            End If
            
            lngNewSchemeId = Val(Nvl(rsTemp!ID))
            
            ReDim Preserve strExeSql(UBound(strExeSql) + 1)
            strExeSql(UBound(strExeSql) - 1) = "zl_Ӱ���ѯ����_��������(" & lngNewSchemeId & ",'" & _
                                                        Nvl(rsData!��������) & "','" & _
                                                        Nvl(rsData!����˵��) & "','" & _
                                                        Nvl(rsData!��ѯ���) & "'," & _
                                                        ufgScheme.ShowingDataRowCount + Val(Nvl(Nvl(rsData!�������))) & "," & _
                                                        0 & ")"
            
            lngOldSchemeId = Val(Nvl(rsData!ID))
        End If
        
        '����Ӱ�񷽰����ü�¼
        ReDim Preserve strExeSql(UBound(strExeSql) + 1)
        strExeSql(UBound(strExeSql) - 1) = "zl_Ӱ���ѯ����_��������(" & lngNewSchemeId & ",'" & _
                                                                Nvl(rsData!¼����Ŀ) & "'," & _
                                                                Val(Nvl(rsData!¼������)) & ",'" & _
                                                                Nvl(rsData!Ĭ��ֵ) & "','" & _
                                                                Nvl(rsData!������Դ) & "'," & _
                                                                Val(Nvl(rsData!¼��˳��)) & ")"
        
        Call rsData.MoveNext
    Wend
    
    
    'д�뷽�������������
    strResult = ExeSqlTrans(strExeSql())
    If strResult <> "" Then
        MsgBoxD Me, "��������ʧ�ܣ�ԭ��Ϊ��" & strResult, vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call LoadSchemeData
    
    MsgBoxD Me, "�ѳɹ�����" & rsData.RecordCount & "�����ݡ�"
End Sub

Private Sub ExportScheme()
'��������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"
    
    dlgFile.Filename = ""
    dlgFile.ShowSave
    
    If dlgFile.Filename = "" Then Exit Sub
    
    strSql = "select a.id, ��������,����˵��,��ѯ���,�������,�Ƿ�Ĭ��,ʹ��״̬,b.id as ����id,¼����Ŀ,¼������,¼��˳��,Ĭ��ֵ,������Դ " & _
            " from Ӱ���ѯ���� a, Ӱ���ѯ���� b where a.id=b.����id and ʹ��״̬=1 order by id, ¼��˳��"
            
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��������")
    If rsData.RecordCount <= 0 Then
        MsgBoxD Me, "û�п����ڵ��������ݣ����鷽�����á�", vbOKOnly, Me.Caption
        Exit Sub
    End If
            
    Call rsData.Save(dlgFile.Filename, adPersistXML)
    
    MsgBoxD Me, "�ѳɹ�����" & rsData.RecordCount & "�����ݡ�"
    
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
'���²˵��Ͱ�ť��ʾ
On Error Resume Next
    Dim blnHasRecord As Boolean
    
    '���û�м�¼����û��ѡ���У��˵��͹������򲻿���
    blnHasRecord = ufgScheme.IsSelectionRow
    
    Select Case control.ID
    
        Case TMenuType.mtSave, TMenuType.mtCancel
            control.Enabled = mblnCurModifyState
            
        Case TMenuType.mtDelScheme, TMenuType.mtModifyScheme, _
            TMenuType.mtMoveLastScheme, TMenuType.mtMoveNextScheme
            control.Enabled = Not mblnCurModifyState And blnHasRecord
            
        Case TMenuType.mtNewScheme
            control.Enabled = Not mblnCurModifyState
            
        Case TMenuType.mtSetDefault
            control.Enabled = Not mblnCurModifyState And blnHasRecord
            
            If blnHasRecord Then
                If control.Parent.type = xtpControlPopup Then
                    control.Caption = IIf(ufgScheme.CurText("�Ƿ�Ĭ��") = "Ĭ��", "ȡ��Ĭ��(&F)", "����Ĭ��(&F)")
                    control.IconId = IIf(ufgScheme.CurText("�Ƿ�Ĭ��") = "Ĭ��", 2616, 3002)
                Else
                    control.Caption = IIf(ufgScheme.CurText("�Ƿ�Ĭ��") = "Ĭ��", "ȡ��Ĭ��", "����Ĭ��")
                    control.IconId = IIf(ufgScheme.CurText("�Ƿ�Ĭ��") = "Ĭ��", 2616, 3002)
                End If
                
                control.Enabled = Not mblnCurModifyState And IIf(ufgScheme.CurText("ʹ��״̬") = "����", True, False)
                
                control.Enabled = Not control.Enabled
                control.Enabled = Not control.Enabled
            End If
            
        Case TMenuType.mtUseScheme
            control.Enabled = Not mblnCurModifyState And blnHasRecord
            
            If blnHasRecord Then
                If control.Parent.type = xtpControlPopup Then
                    control.Caption = IIf(ufgScheme.CurText("ʹ��״̬") = "����", "����(&A)", "����(&A)")
                    control.IconId = IIf(ufgScheme.CurText("ʹ��״̬") = "����", 3006, 3009)
                Else
                    control.Caption = IIf(ufgScheme.CurText("ʹ��״̬") = "����", "����", "����")
                    control.IconId = IIf(ufgScheme.CurText("ʹ��״̬") = "����", 3006, 3009)
                End If
                
                control.Enabled = Not control.Enabled
                control.Enabled = Not control.Enabled
            End If
            
        Case TMenuType.mtCheckScheme
            control.Enabled = blnHasRecord
            
    End Select
End Sub

Private Sub MoveLastScheme()
'���Ʒ���
    Dim strSql As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call ufgScheme.MoveUp(ufgScheme.SelectionRow)
    
    strSql = "zl_Ӱ���ѯ����_�ƶ�����(" & ufgScheme.CurKeyValue & "," & ufgScheme.SelectionRow & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "���Ʒ���")
    
    mblnIsChange = True
End Sub


Private Sub MoveNextScheme()
'���Ʒ���
    Dim strSql As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call ufgScheme.MoveDown(ufgScheme.SelectionRow)
    
    strSql = "zl_Ӱ���ѯ����_�ƶ�����(" & ufgScheme.CurKeyValue & "," & ufgScheme.SelectionRow & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "���Ʒ���")
    
    mblnIsChange = True
End Sub


Private Sub UseScheme()
'���÷���
    Dim strSql As String
    Dim strCurUseState As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    strCurUseState = ufgScheme.CurText("ʹ��״̬")
    
    strSql = "zl_Ӱ���ѯ����_ʹ��״̬(" & ufgScheme.CurKeyValue & "," & IIf(strCurUseState = "����", 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ʹ��״̬����")
    
    ufgScheme.CurText("ʹ��״̬") = IIf(strCurUseState = "����", "����", "����")
    
    mblnIsChange = True
End Sub

Private Sub DefaultScheme()
'����Ĭ�Ϸ���
    Dim strSql As String
    Dim strCurDefaultState As String
    Dim i As Long
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    strCurDefaultState = ufgScheme.CurText("�Ƿ�Ĭ��")
    
    strSql = "zl_Ӱ���ѯ����_����Ĭ��(" & ufgScheme.CurKeyValue & "," & IIf(strCurDefaultState = "Ĭ��", 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "����Ĭ��")
    
    For i = 1 To ufgScheme.GridRows - 1
        ufgScheme.Text(i, "�Ƿ�Ĭ��") = ""
    Next i
    
    ufgScheme.CurText("�Ƿ�Ĭ��") = IIf(strCurDefaultState = "Ĭ��", "", "Ĭ��")
    
    mblnIsChange = True
End Sub

Private Sub NewScheme()
'��������
    ufgScheme.NewRow
    
    Call ConfigFaceEditState(True)
    
    ufgInputCfg.NewRow
    
    ufgInputCfg.CurText("¼����Ŀ") = "��ʼ����"
    ufgInputCfg.CurText("¼������") = "1-���ڿ�"
    
    
    ufgInputCfg.NewRow
    
    ufgInputCfg.CurText("¼����Ŀ") = "��������"
    ufgInputCfg.CurText("¼������") = "1-���ڿ�"
End Sub

Private Sub DelScheme()
'ɾ������
    Dim strSql As String
    
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    strSql = "zl_Ӱ���ѯ����_ɾ������(" & ufgScheme.CurKeyValue & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "ɾ����ѯ����")
    
    Call ufgScheme.DelCurRow(False)
    
    mblnIsChange = True
End Sub

Private Sub ModifyScheme()
'�޸ķ���
    If Not ufgScheme.IsSelectionRow Then
        MsgBoxD Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call ConfigFaceEditState(True)
End Sub

Private Sub SaveScheme()
'���淽��
'��Ҫ�ж����޸�ԭ���ķ������������ķ���
    Dim lngSchemeId As Long
    Dim rsData As ADODB.Recordset
    Dim strSql As String
    Dim i As Long
    Dim strExeSql() As String
    Dim strResult As String
    
    strResult = VerifyScheme
    If strResult <> "" Then
        If MsgBoxD(Me, strResult & vbCrLf & "��Ҫǿ�Ʊ����𣿣���", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    lngSchemeId = Val(ufgScheme.CurKeyValue)
    If lngSchemeId <= 0 Then
        'С�ڻ����0��ʾ�����ķ���
        strSql = "select Ӱ���ѯ����_ID.NextVal as ID from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ȡ�·���ID")
        If rsData.RecordCount <= 0 Then
            MsgBoxD Me, "���ܻ�ȡ�����ķ���ID��ϵͳ���˳����档", vbOKOnly, Me.Caption
            Exit Sub
        End If
        
        lngSchemeId = Val(Nvl(rsData!ID))
        
        ReDim Preserve strExeSql(1)
        strExeSql(0) = "zl_Ӱ���ѯ����_��������(" & lngSchemeId & ",'" & _
                                                    txtSchemeName.Text & "','" & _
                                                    txtSchemeMemo.Text & "','" & _
                                                    txtFilterSql.Text & "'," & _
                                                    ufgScheme.SelectionRow & "," & _
                                                    cbxDepart.ItemData(cbxDepart.ListIndex) & ")"

    Else
        '�޸ĵķ�������
        ReDim Preserve strExeSql(1)
        strExeSql(0) = "zl_Ӱ���ѯ����_�������(" & lngSchemeId & ")"
        
        ReDim Preserve strExeSql(2)
        strExeSql(1) = "zl_Ӱ���ѯ����_���·���(" & lngSchemeId & ",'" & _
                                                    txtSchemeName.Text & "','" & _
                                                    txtSchemeMemo.Text & "','" & _
                                                    txtFilterSql.Text & "')"
        
    End If
    
    For i = 1 To ufgInputCfg.GridRows - 1
        If Not ufgInputCfg.RowHidden(i) Then
            ReDim Preserve strExeSql(UBound(strExeSql) + 1)
            strExeSql(UBound(strExeSql) - 1) = "zl_Ӱ���ѯ����_��������(" & lngSchemeId & ",'" & _
                                                                    ufgInputCfg.Text(i, "¼����Ŀ") & "'," & _
                                                                    Val(ufgInputCfg.Text(i, "¼������")) & ",'" & _
                                                                    ufgInputCfg.Text(i, "Ĭ��ֵ") & "','" & _
                                                                    ufgInputCfg.Text(i, "������Դ") & "'," & _
                                                                    i & ")"
        End If
    Next i
        
    'д�뷽�������������
    strResult = ExeSqlTrans(strExeSql())
    If strResult <> "" Then
        MsgBoxD Me, "��������ʧ�ܣ�ԭ��Ϊ��" & strResult, vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    '���·����б���ʾ
    With ufgScheme
        .CurKeyValue = lngSchemeId
        .CurText("��������") = txtSchemeName.Text
        .CurText("����˵��") = txtSchemeMemo.Text
        .CurText("��ѯ���") = Replace(txtFilterSql.Text, "''", "'")
        .CurText("�Ƿ�Ĭ��") = ""
        .CurText("ʹ��״̬") = "����"
        .CurText("�������") = ufgScheme.SelectionRow
    End With


    Call ConfigFaceEditState(False)
    
    mblnIsChange = True
End Sub


Private Function ExeSqlTrans(ByVal strSql As Variant) As String
    Dim i As Long
    Dim strExeSql As String
    
    ExeSqlTrans = ""
    
    gcnOracle.BeginTrans
    
On Error GoTo errRollback
    For i = 0 To UBound(strSql)
        strExeSql = strSql(i)
        If strExeSql <> "" Then
            Call zlDatabase.ExecuteProcedure(strExeSql, "���淽������")
        End If
    Next i
    
    gcnOracle.CommitTrans
Exit Function
errRollback:
    gcnOracle.RollbackTrans
    ExeSqlTrans = err.Description
End Function

Private Sub CancelScheme()
'ȡ�����������������޸�
    If Not mblnCurModifyState Then Exit Sub
    
    If ufgScheme.CurKeyValue = "" Then
        Call ufgInputCfg.ClearListData
        Call ufgScheme.DelCurRow(False)
        
        
        txtSchemeName.Text = ""
        txtSchemeMemo.Text = ""
        txtFilterSql.Text = ""
    Else
        Call LoadSchemeCfgData(ufgScheme.SelectionRow)
    End If
    
    
    Call ConfigFaceEditState(False)
End Sub


Private Function VerifyScheme() As String
'��֤��ǰ���õķ���
    Dim i As Long
    Dim j As Long
    Dim strInputProNames As String
    Dim strParName As String
    Dim strResult As String
    Dim rsTemp As ADODB.Recordset
    Dim strSqlFrom As String
    Dim strPars(20) As String
    
    VerifyScheme = ""
    
    If Trim(txtSchemeName.Text) = "" Then
        VerifyScheme = "δ��ͨ����֤��ԭ���Ƿ�������Ϊ�ա�"
        Exit Function
    End If
    
    If Trim(txtFilterSql.Text) = "" Then
        VerifyScheme = "δ��ͨ����֤��ԭ���ǹ�����乹��Ϊ�ա�"
        Exit Function
    End If
        
    For i = 1 To ufgInputCfg.GridRows - 1
        strInputProNames = strInputProNames & "[" & ufgInputCfg.Text(i, "¼����Ŀ") & "]"
        
        '��֤Ĭ��ֵ����
        strParName = ufgInputCfg.Text(i, "Ĭ��ֵ")
        If IsParameterFormat(strParName) Then
            strResult = TestParameter(strParName, strInputProNames)
            
            If strResult <> "" Then
                VerifyScheme = "��" & i & "��'Ĭ��ֵ'�е����� " & vbCrLf & strParName & vbCrLf & " δͨ����֤��ԭ�����£�" & strResult
                Exit Function
            End If
        End If
        
        '��֤������Դ����
        strSqlFrom = ufgInputCfg.Text(i, "������Դ")
        strResult = TestSql(strSqlFrom, strInputProNames)
        If strResult <> "" Then
            VerifyScheme = "��" & i & "��'������Դ'�е����� " & vbCrLf & strSqlFrom & vbCrLf & " δͨ����֤��ԭ�����£�" & strResult
            Exit Function
        End If
        
        
    Next i
    
    '��֤�������
    strResult = TestSql(Replace(txtFilterSql.Text, "''", "'"), strInputProNames)
    If strResult <> "" Then
        VerifyScheme = "������乹�� " & vbCrLf & txtFilterSql.Text & vbCrLf & " δͨ����֤��ԭ�����£�" & strResult
        Exit Function
    End If

End Function


Private Function TestSql(ByVal strSqlFrom As String, ByVal strInputProNames As String) As String
On Error GoTo errHandle
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strParName As String
    Dim strPars(20) As String
    Dim strTestSql As String
    
    TestSql = ""
        
    If strSqlFrom = "" Then Exit Function
    
    strTestSql = strSqlFrom
    
    Call GetParameterNames(strTestSql, strPars)
    
    For i = 1 To 20
        strParName = strPars(i)
        If strParName <> "" Then
            TestSql = TestParameter(strParName, strInputProNames)
            If TestSql <> "" Then Exit Function
            
            strTestSql = Replace(strTestSql, strParName, "Null")
            
            '�ָ������������ù��������ʹ��
            strPars(i) = ""
        End If
    Next i
    
    'sql��ѯ������֤
    If Not IsParameterFormat(strTestSql) Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strTestSql, "��֤Sql��ѯ")
    End If
    
Exit Function
errHandle:
    TestSql = err.Description
End Function

Private Function GetParameterNames(ByVal strSqlFrom As String, ByRef strParameters() As String) As Boolean
'�ж�����Դsql����Ƿ��������
    Dim strTemp As String
    Dim lngStart As Long, lngEnd As Long
    Dim lngParCount As Long
    
    strTemp = strSqlFrom
    lngStart = InStr(strTemp, "[")
    lngEnd = InStr(strTemp, "]")
    
    GetParameterNames = False
    
    If lngStart <= 0 Or lngEnd <= 0 Then Exit Function
    
    lngParCount = 0
    
    'ѭ����ȡ���еĲ�������
    While lngStart > 0 And lngEnd > 0
        
        lngParCount = lngParCount + 1
        
        strTemp = Mid(strTemp, lngStart, 1024)
        lngEnd = InStr(strTemp, "]")
        
        strParameters(lngParCount) = Mid(strTemp, 1, lngEnd)
        
        strTemp = Mid(strTemp, lngEnd + 1, 1024)
        
        lngStart = InStr(strTemp, "[")
        lngEnd = InStr(strTemp, "]")
    Wend
       
    GetParameterNames = IIf(lngParCount > 0, True, False)
End Function

Private Function TestParameter(ByVal strParameterName As String, ByVal strInputProNames As String) As String
On Error GoTo errHandle

    TestParameter = ""
        
    If strParameterName = "" Then Exit Function
    If Not IsParameterFormat(strParameterName) Then
        '������ǲ�����ʽ���������ֱ����Ĭ��ֵ���ô��������ֵ������Ĭ��ֵ���õ��ǡ�2012-03-05������û�в��á�[��ǰʱ��]����ʽ
        TestParameter = ""
        Exit Function
    End If
    
    Select Case strParameterName
        Case "[��ǰ����]", "[��ǰʱ��]", "[��ǰ�û�ID]", "[��ǰ����ID]", "[��ǰϵͳ���]", "[��ǰģ����]"
            Exit Function
        Case Else
            '��ȡ�ı����Ӧ��ֵ
             If InStr(strInputProNames, strParameterName) > 0 Then
                Exit Function
            End If
    End Select
    
    '��ǰ��Ĵ����У�����ҵ���Ӧ�Ĳ������ͻ�ֱ�ӽ�ֵ���Ǻ��������أ����ִ�е����˵��û���ҵ���Ӧ���������������Զ���ű��硰[now-1]��
    
    'ִ�нű�����
    Call RunScripting(strParameterName)
Exit Function
errHandle:
    TestParameter = err.Description
End Function


Private Function IsParameterFormat(ByVal strData As String) As Boolean
'�ж������Ƿ�Ϊ��������
    IsParameterFormat = False
    
    If strData = "" Then Exit Function
    If Left(strData, 1) <> "[" Or Right(strData, 1) <> "]" Then Exit Function
    
    IsParameterFormat = True
End Function

Private Function RunScripting(ByVal strScript As String) As String
'ִ��vbs�ű�
    Dim strFormatScript As String

    strFormatScript = Replace(Replace(strScript, "[", ""), "]", "")

On Error GoTo errHandle
    RunScripting = sctExecute.Eval(strFormatScript)
    Exit Function
errHandle:
    strFormatScript = "function return()" & vbCrLf & strFormatScript & " end function"
    Call sctExecute.AddCode(strFormatScript)
    
    RunScripting = sctExecute.Run("return")
End Function


Private Sub LoadSchemeData()
'���뷽��
    Dim strSql As String
    
    strSql = "select id, ��������,����˵��,��ѯ���,�������,�Ƿ�Ĭ��,ʹ��״̬,�������� from Ӱ���ѯ���� where ��������=[1] order by �������"
    Set ufgScheme.AdoData = zlDatabase.OpenSQLRecord(strSql, "��ѯ���˷���", cbxDepart.ItemData(cbxDepart.ListIndex))
    
    Call ufgScheme.RefreshData
    
    If ufgScheme.ShowingDataRowCount > 1 Then
        Call ufgScheme.LocateRow(1)
    End If
End Sub

Private Sub LoadSchemeCfgData(ByVal lngSchemeRowIndex As Long)
'���뷽����������
    Dim strSql As String
    
    txtFilterSql.Text = Replace(ufgScheme.Text(lngSchemeRowIndex, "��ѯ���"), "'", "''")
    
    txtSchemeName.Text = ufgScheme.Text(lngSchemeRowIndex, "��������")
    txtSchemeMemo.Text = ufgScheme.Text(lngSchemeRowIndex, "����˵��")
    
    strSql = "select id, ����ID,¼����Ŀ,¼������,¼��˳��,Ĭ��ֵ,������Դ from Ӱ���ѯ���� where ����Id=[1] order by ¼��˳��"
    Set ufgInputCfg.AdoData = zlDatabase.OpenSQLRecord(strSql, "��ѯ��������", ufgScheme.KeyValue(lngSchemeRowIndex))
    
    Call ufgInputCfg.RefreshData
    
End Sub


Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
    picBack.Left = Left
    picBack.Top = Top
    picBack.Width = Right - Left
    picBack.Height = Bottom - Top - IIf(stbThis.Visible, stbThis.Height, 0)
End Sub



Private Sub cbxDepart_Click()
    Call LoadSchemeData
End Sub

Private Sub cmdDelInput_Click()
'ɾ��¼����Ŀ
On Error GoTo errHandle
    Call ufgInputCfg.DelCurRow(False)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdInsertPar_Click()
'�������
On Error GoTo errHandle
    Dim strPar As String
    Dim frmPar As New frmCustomInsertPar
    
    strPar = frmPar.ShowParameterWindow(ufgInputCfg, True, Me)
    If strPar <> "" Then
        txtFilterSql.SelText = strPar
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdMoveLast_Click()
'����һ��
On Error GoTo errHandle
    Call ufgInputCfg.MoveUp(ufgInputCfg.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdMoveNext_Click()
'����һ��
On Error GoTo errHandle
    Call ufgInputCfg.MoveDown(ufgInputCfg.SelectionRow)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdNewInput_Click()
'����¼����Ŀ
On Error GoTo errHandle
    Call ufgInputCfg.NewRow
    Call ufgInputCfg.DataGrid.EditCell
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    
    
'    InitDebugObject 1290, Me, "zlhis", "HIS"
    
    Call RestoreWinState(Me, App.ProductName)
    
    mblnCurModifyState = False
    mblnIsChange = False
    
    Call InitCommandBars
    Call InitFaceList

    Call ConfigFaceEditState(False)
    
    Call InitDepts
'    Call LoadSchemeData
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    str��Դ = "1,2,3"
    If InStr(gstrPrivs, "���п���") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " and (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null ) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If
   
    cbxDepart.Clear
    cbxDepart.AddItem "����"
    cbxDepart.ItemData(cbxDepart.ListCount - 1) = 0
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ؿ�����Ϣ", CStr("," & str��Դ & ","))
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        cbxDepart.ListIndex = 0
        
        Exit Function
    End If
    
    While Not rsTmp.EOF
        cbxDepart.AddItem Nvl(rsTmp!����)
        cbxDepart.ItemData(cbxDepart.ListCount - 1) = Nvl(rsTmp!ID)
        
        Call rsTmp.MoveNext
    Wend
    
    cbxDepart.ListIndex = 0
End Function


Private Sub InitFaceList()
'��ʼ��������ص������б�

    ufgScheme.IsKeepRows = False
    ufgScheme.IsEjectConfig = False
    ufgScheme.ColNames = M_STR_SCHEME_COLS
    ufgScheme.ColConvertFormat = M_STR_SCHEME_CONVERT
    ufgScheme.ExtendLastCol = True
    ufgScheme.RowHeightMin = 320
    
    ufgInputCfg.IsKeepRows = False
    ufgInputCfg.IsEjectConfig = False
    ufgInputCfg.ColNames = M_STR_INPUT_COLS
    ufgInputCfg.ColConvertFormat = M_STR_INPUT_CONVERT
    ufgInputCfg.ExtendLastCol = True
    ufgInputCfg.RowHeightMin = 320
End Sub


Private Sub ConfigFaceEditState(ByVal blnIsEdit As Boolean)
    mblnCurModifyState = blnIsEdit
    
    txtSchemeName.Locked = Not blnIsEdit
    txtSchemeMemo.Locked = Not blnIsEdit
    
    ufgInputCfg.ReadOnly = Not blnIsEdit
    
    cmdNewInput.Enabled = blnIsEdit
    cmdDelInput.Enabled = blnIsEdit
    cmdMoveLast.Enabled = blnIsEdit
    cmdMoveNext.Enabled = blnIsEdit
    
    txtFilterSql.Locked = Not blnIsEdit
    cmdInsertPar.Enabled = blnIsEdit
    
    cbxDepart.Enabled = Not blnIsEdit
    labDepart.Enabled = Not blnIsEdit
    
    ufgScheme.DataGrid.Enabled = Not blnIsEdit
    
    If blnIsEdit Then
        txtSchemeName.BackColor = &H80000005
        txtSchemeMemo.BackColor = &H80000005
        txtFilterSql.BackColor = &H80000005
        ufgInputCfg.BackColor = &H80000005
        ufgScheme.BackColor = &H8000000F
    Else
        txtSchemeName.BackColor = &H8000000F
        txtSchemeMemo.BackColor = &H8000000F
        txtFilterSql.BackColor = &H8000000F
        ufgInputCfg.BackColor = &H8000000F
        ufgScheme.BackColor = &H80000005
    End If
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'����ж��
On Error GoTo errHandle
'    Unload frmCustomQueryFrom

    Call SaveWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picBack_Resize()
On Error Resume Next
    Call ucSplitter1.RePaint(False)
End Sub


Private Sub picScheme_Resize()
On Error Resume Next
    cbxDepart.Width = picScheme.Width - cbxDepart.Left - 60
    
    ufgScheme.Left = 60
    ufgScheme.Top = cbxDepart.Top + cbxDepart.Height + 60
    ufgScheme.Height = picScheme.Height - ufgScheme.Top
    ufgScheme.Width = picScheme.Width - 60
End Sub

Private Sub picSchemeCfg_Resize()
On Error Resume Next
    framBase.Left = 0
    framBase.Top = 0
    framBase.Width = picSchemeCfg.ScaleWidth

        txtSchemeMemo.Width = framBase.Width - txtSchemeMemo.Left - 120
        
    framInput.Left = 0
    framInput.Top = framBase.Top + framBase.Height + 60
    framInput.Width = picSchemeCfg.ScaleWidth
    framInput.Height = picSchemeCfg.ScaleHeight - framBase.Height - framSql.Height - 120
    
        ufgInputCfg.Left = 60
        ufgInputCfg.Top = 240
        ufgInputCfg.Width = framInput.Width - 120
        ufgInputCfg.Height = framInput.Height - cmdMoveLast.Height - 360
        
        cmdNewInput.Left = ufgInputCfg.Left
        cmdNewInput.Top = ufgInputCfg.Top + ufgInputCfg.Height + 60
        
        cmdDelInput.Left = cmdNewInput.Left + cmdNewInput.Width + 60
        cmdDelInput.Top = cmdNewInput.Top
        
        cmdMoveNext.Left = framInput.Width - 60 - cmdMoveNext.Width
        cmdMoveNext.Top = cmdNewInput.Top
        
        cmdMoveLast.Left = cmdMoveNext.Left - 60 - cmdMoveLast.Width
        cmdMoveLast.Top = cmdNewInput.Top
        
    framSql.Left = 0
    framSql.Top = framInput.Top + framInput.Height + 60
    framSql.Width = picSchemeCfg.ScaleWidth
    
        txtFilterSql.Left = 60
        txtFilterSql.Top = 240
        txtFilterSql.Width = framSql.Width - 120
        
        cmdInsertPar.Left = framSql.Width - 60 - cmdInsertPar.Width
    
End Sub

Private Sub ufgInputCfg_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
'����¼����Ŀ��������Դ����Ĭ��ֵ
On Error GoTo errHandle
    Dim frmSqlFrom As New frmCustomQueryFrom
    Dim strSql As String
    Dim lngCurCol As Long
    
    lngCurCol = ufgInputCfg.GetColIndex("Ĭ��ֵ")
    
    If Col > lngCurCol Then lngCurCol = Col


    strSql = ufgInputCfg.Text(Row, ufgInputCfg.GetColName(lngCurCol))
    
    strSql = frmSqlFrom.ShowSqlFromWindow(strSql, ufgInputCfg, Me)
    
    ufgInputCfg.Text(Row, ufgInputCfg.GetColName(lngCurCol)) = strSql
    
    Unload frmSqlFrom
Exit Sub
errHandle:
    Unload frmSqlFrom
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgInputCfg_OnDblClick()
'    If ufgInputCfg.ShowingDataRowCount < 1 Then Exit Sub
'
'    Call ufgInputCfg_OnCellButtonClick(ufgInputCfg.SelectionRow, 0)
End Sub

Private Sub ufgScheme_OnSelChange()
On Error GoTo errHandle
    If ufgScheme.IsSelectionRow Then
        Call LoadSchemeCfgData(ufgScheme.SelectionRow)
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
