VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�������չ��� 
   Caption         =   "�������չ���"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11250
   Icon            =   "frm�������չ���.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   11250
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   360
      ScaleHeight     =   6855
      ScaleWidth      =   4935
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame fraPatient 
         Caption         =   "����������Ϣ"
         Height          =   2295
         Left            =   360
         TabIndex        =   26
         Top             =   120
         Width           =   3855
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   4
            Top             =   1440
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.TextBox txtEndNo 
            Height          =   300
            Left            =   1200
            TabIndex        =   1
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkOutHp 
            Caption         =   "��Ժ����"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.ComboBox cboOutDept 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   2580
         End
         Begin VB.TextBox txtBeginNo 
            Height          =   300
            Left            =   1200
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   5
            Top             =   1800
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   960
            TabIndex        =   30
            Top             =   1845
            Width           =   180
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   960
            TabIndex        =   29
            Top             =   645
            Width           =   180
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ����"
            Height          =   180
            Left            =   420
            TabIndex        =   28
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label lblNo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��"
            Height          =   180
            Left            =   600
            TabIndex        =   27
            Top             =   285
            Width           =   540
         End
      End
      Begin VB.Frame fraSong 
         Caption         =   "����������Ϣ"
         Height          =   2775
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   3855
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   10
            Top             =   1200
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   7
            Top             =   345
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.TextBox txtAuditingMan 
            Height          =   300
            Left            =   1200
            TabIndex        =   13
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox txtApplyman 
            Height          =   300
            Left            =   1200
            TabIndex        =   12
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CheckBox chkIncept 
            Caption         =   "��������"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkRecord 
            Caption         =   "��Ŀ����"
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1245
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   8
            Top             =   720
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   11
            Top             =   1560
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483646
            CalendarTitleForeColor=   -2147483643
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   110034947
            CurrentDate     =   39777
         End
         Begin VB.Label lblAuditingMan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   600
            TabIndex        =   25
            Top             =   2355
            Width           =   660
         End
         Begin VB.Label lblApplyman 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   600
            TabIndex        =   24
            Top             =   1965
            Width           =   660
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   180
            Index           =   2
            Left            =   960
            TabIndex        =   23
            Top             =   765
            Width           =   300
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   180
            Index           =   3
            Left            =   960
            TabIndex        =   22
            Top             =   1620
            Width           =   300
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         Height          =   300
         Left            =   3360
         TabIndex        =   14
         Top             =   5520
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫѡ(&A)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   840
         TabIndex        =   15
         Top             =   5520
         Width           =   1100
      End
      Begin VB.CommandButton cmdNoAll 
         Caption         =   "ȫ��ѡ(&N)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   16
         Top             =   5520
         Width           =   1100
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   3
         Left            =   3120
         Picture         =   "frm�������չ���.frx":0442
         Top             =   5970
         Width           =   240
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   2
         Left            =   2115
         Picture         =   "frm�������չ���.frx":6C94
         Top             =   5955
         Width           =   240
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   1
         Left            =   1155
         Picture         =   "frm�������չ���.frx":6FD6
         Top             =   5955
         Width           =   240
      End
      Begin VB.Image imgCelNo 
         Height          =   240
         Index           =   0
         Left            =   75
         Picture         =   "frm�������չ���.frx":7318
         Top             =   5955
         Width           =   240
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�� δ���ա���  �ѽ���     �ѱ�Ŀ     �ٽ���"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   135
         TabIndex        =   31
         Top             =   6000
         Width           =   3870
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   7290
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm�������չ���.frx":765A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14764
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
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   6000
      ScaleHeight     =   4815
      ScaleWidth      =   3855
      TabIndex        =   18
      Top             =   120
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   4575
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   3615
         _cx             =   6376
         _cy             =   8070
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
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   280
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
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frm�������չ���.frx":7EEE
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frm�������չ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objToolBar As CommandBar
Private objMenu As CommandBarPopup
Private objPopup As CommandBarPopup
Private objControl As CommandBarControl
Private objCombox As CommandBarComboBox
Private objExtendedBar As CommandBar

Private Const conMenu_Edit_Display = 209            '�鿴����(&C)
Private Const conMenu_View_ToolBar_Visible = 6014    '���ع�����(&V)
Private Const conMenu_View_Choose = 7900

Private mstrPrivs As String            '��ǰʹ����Ȩ�޴�
Private mlngApplyId As Long            '��������
Private mlngTempId As Long
Private mstrApply As String
Private mblnBootUp As Boolean
Private mstrMsg As String
Private mstrListTitle As String
Private mbln����ϵͳ As Boolean

'��������
Private mdtInBeginDate As Date           '��������ʱ�俪ʼ
Private mdtInEndDate As Date             '��������ʱ�����
Private mdtRecBeginDate As Date        '������¼ʱ�俪ʼ
Private mdtRecEndDate As Date          '������¼ʱ�����
Private mdtOutBeginDate As Date        '���˳�Ժʱ�俪ʼ
Private mdtOutEndDate As Date          '���˳�Ժʱ�����

Private mstrNoShowDate As String          '����ʾ��ʷ�ѱ�Ŀδ�Ǽǳ�Ժ����

Private mblnShow As Boolean
Private mintDblClick As Integer
Private mintDelete As Integer
Private mstrFind As String
Private mcllTemp As New Collection

Private mintEditState As Integer
Private mlngModule As Long   'ģ���

Private Function GetInitDept() As Boolean
    '----------------------------------------------------------------------------
    '����:��ȡ����
    '����:���п���,�򷵻�True,���򷵻�False
    '----------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If InStr(mstrPrivs, "���п���") > 0 Then
        gstrSQL = "Select a.id,a.����, a.����" & vbNewLine & _
                "From ���ű� A, ��������˵�� B" & vbNewLine & _
                "Where a.Id = b.����id And b.�������� = '�ٴ�' And (b.������� = 2 Or b.������� = 3) And " & vbNewLine & _
                  Where����ʱ��("A") & zl_��ȡվ������(True, "a")
    Else
        gstrSQL = "Select a.id,a.����, a.����,c.ȱʡ" & vbNewLine & _
                "From ���ű� A, ��������˵�� B ,������Ա C" & vbNewLine & _
                "Where a.Id = b.����id And b.�������� = '�ٴ�' And (b.������� = 2 Or b.������� = 3) And A.ID=c.����id And C.��ԱID=" & UserInfo.ID & vbNewLine & _
                " And " & Where����ʱ��("A") & zl_��ȡվ������(True, "a")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rsTemp.EOF Then
        mstrMsg = "��ǰ�޿����ٴ�����,���鲿�����ã�"
        GetInitDept = False
        Exit Function
    End If
    
    With cboOutDept
        .Clear
        'װ������
        If InStr(mstrPrivs, "���п���") > 0 Then '��Ȩ�޲��������п���
            .AddItem "���в���"
             .ItemData(.NewIndex) = 1
             If mlngApplyId = 0 Then '�״�ˢ�²�ѡ�����в���
                 .ListIndex = .NewIndex
                 mstrApply = "���в���"
             End If
        End If
            
        Do Until rsTemp.EOF
            .AddItem rsTemp!���� & "-" & rsTemp!����
            .ItemData(.NewIndex) = rsTemp!ID
            
            If mlngApplyId = 0 And InStr(mstrPrivs, "���п���") = 0 Then '�״δ򿪴���
                If UserInfo.����ID = rsTemp!ID Then
                    .ListIndex = .NewIndex
                    mstrApply = rsTemp!���� & "-" & rsTemp!����
                End If
            ElseIf mlngApplyId > 0 Then '�ظ�ˢ��
                If rsTemp!ID = mlngApplyId Then
                    .ListIndex = .NewIndex
                    mstrApply = rsTemp!���� & "-" & rsTemp!����
                End If
            End If
            
            rsTemp.MoveNext
        Loop
        cmdAll.Enabled = True
        cmdNoAll.Enabled = True
    End With
    GetInitDept = True
    mblnBootUp = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitMenus()
'��ʼ���˵���������
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    cbsThis.EnableCustomization False
    
    '����������
    Set cbsThis.Icons = zlCommFun.GetPubIcons 'imgIcon.Icons
    With cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    '����˵�
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup  '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel...")
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "�鿴(&C)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "Сͼ��(&D)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Visible, "���ع�����(&V)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Choose, "ѡ����(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�����") '
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): objControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("X"), conMenu_File_Exit
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_Help_About
    End With

    '����������
    Set objToolBar = cbsThis.Add("������", xtpBarTop)
    objToolBar.ShowTextBelowIcons = False
    objToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    With objToolBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objToolBar.Controls
        objControl.STYLE = xtpButtonIconAndCaption
    Next
       
    '���������˵�
    Set objExtendedBar = cbsThis.Add("Popup", xtpBarPopup)
    With objExtendedBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "�鿴(&C)"): objControl.BeginGroup = True
    End With

End Sub

Private Sub InitMenusElectron()
'��ʼ���˵���������-���Ӳ�������
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    cbsThis.EnableCustomization False
    
    '����������
    Set cbsThis.Icons = zlCommFun.GetPubIcons 'imgIcon.Icons
    With cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    '����˵�
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup  '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "��ӡԤ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel...")
        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "������(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "�ܾ����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "���˽���(&1)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "���˾ܾ�(&2)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "�鿴(&C)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "Сͼ��(&D)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Visible, "���ع�����(&V)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Choose, "ѡ����(&C)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "���Ĳ���(&V)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
    End With
    Set objMenu = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�����") '
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."): objControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("X"), conMenu_File_Exit
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_Help_About
    End With

    '����������
    Set objToolBar = cbsThis.Add("������", xtpBarTop)
    objToolBar.ShowTextBelowIcons = False
    objToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
    With objToolBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "������"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "�ܾ����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "���˽���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "���˾ܾ�")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each objControl In objToolBar.Controls
        objControl.STYLE = xtpButtonIconAndCaption
    Next
       
    '���������˵�
    Set objExtendedBar = cbsThis.Add("Popup", xtpBarPopup)
    With objExtendedBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Plan, "������")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "�ܾ����")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "���˽���"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "���˾ܾ�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Display, "�鿴(&C)"): objControl.BeginGroup = True
    End With

End Sub

Private Sub InitDkpMain()
    Dim panThis As Pane
    Dim panfilter As Pane
    Set panfilter = dkpMain.CreatePane(1, 300, 200, DockLeftOf, Nothing)
    panfilter.Title = "����������Ϣ��ѯ����"
'
    panfilter.Options = PaneNoFloatable Or PaneNoCloseable
    
    Set panThis = dkpMain.CreatePane(2, 750, 500, DockRightOf, Nothing)
    panThis.Title = "�������͵Ǽ���Ϣ"
    panThis.Options = PaneNoFloatable Or PaneNoCloseable Or PaneNoHideable ' Or PaneNoCaption
    
    Call GetdkpMain(Me.Caption & "-" & mlngModule, "dkpMain")
    
    Me.dkpMain.SetCommandBars Me.cbsThis
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.LunaColors = True
    Me.dkpMain.Options.HideClient = True
End Sub

Private Function SavedkpMain(ByVal strCaption As String, ByVal strKey As String) As Boolean
    Dim strValue As String
    If dkpMain.FindPane(1).Hidden Then
        strValue = "0"
    Else
        strValue = "1"
    End If
     If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
     zlDatabase.SetPara "���������Ƿ���ʾ", strValue, glngSys, mlngModule
End Function

Private Function GetdkpMain(ByVal strCaption As String, ByVal strKey As String) As Boolean
    Dim strReg As String
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
'        dkpMain.FindPane(1).Hide
        Exit Function
    End If
    strReg = zlDatabase.GetPara("���������Ƿ���ʾ", glngSys, mlngModule, "")
    If strReg = "" Then
        dkpMain.FindPane(1).Hide 'HidePane panfilter
        Exit Function
    End If
    Err = 0: On Error GoTo errHand:
    If strReg = "0" Then
        dkpMain.FindPane(1).Hide
    End If
    GetdkpMain = True
    Exit Function
errHand:
End Function

Private Sub cboOutDept_Click()
    
    If Me.cboOutDept.ListCount = 0 Then Exit Sub
    If Me.cboOutDept.ListIndex = -1 Then Exit Sub
    
    If cboOutDept.ItemData(cboOutDept.ListIndex) = 1 And cboOutDept.Text = "���в���" Then
        mlngApplyId = 0
        cmdAll.Enabled = True
        cmdNoAll.Enabled = True
    Else
        mlngApplyId = cboOutDept.ItemData(cboOutDept.ListIndex)
        cmdAll.Enabled = True
        cmdNoAll.Enabled = True
    End If
    
End Sub
Private Sub cboOutDept_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Call zlControl.CboSetIndex(cboOutDept.hWnd, zlControl.CboMatchIndex(cboOutDept.hWnd, KeyAscii))
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCurrIndex As Long
    Dim strNo As String
    Dim strMsg As String
    Dim rsSQL As ADODB.Recordset
    Dim strSQL As String
    Dim strNote As String
    Dim strNow As String
    
    On Error GoTo errHand
    Call SQLRecord(rsSQL)

    
'   ����ť���ɼ����ݷ�ʽ����ִ��
    If Control.ID <> 0 Then
        If cbsThis.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(0)
            If Me.ActiveControl Is vfgList Then
                vfgList.Redraw = False
                zlRptPrint 0
                vfgList.Redraw = True
                vfgList.Col = 0
                vfgList.ColSel = vfgList.Cols - 1
            End If
        Case conMenu_File_Print
            Call zlRptPrint(1)
            If Me.ActiveControl Is vfgList Then
                vfgList.Redraw = False
                zlRptPrint 1
                vfgList.Redraw = True
                vfgList.Col = 0
                vfgList.ColSel = vfgList.Cols - 1
            End If
        Case conMenu_File_Excel
            If Me.ActiveControl Is vfgList Then
                vfgList.Redraw = False
                zlRptPrint 3
                vfgList.Redraw = True
                vfgList.Col = 0
                vfgList.ColSel = vfgList.Cols - 1
            End If
        Case conMenu_File_Parameter
            frm�������ղ���.�������� Me, mlngModule, mstrPrivs
            mblnShow = IIf(Val(zlDatabase.GetPara("����ʾ��ʷ�ѱ�Ŀδ�Ǽ�", glngSys, mlngModule)) = 1, 1, 0) = 1
            If mblnShow Then Call GetNoSHowData
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
        Case conMenu_Edit_NewItem
           Call SetAdd
        Case conMenu_Edit_Modify
            Call SetModify
        Case conMenu_Edit_Delete
            Call SetDelete
        Case conMenu_Edit_Display
            Call SetDisplay
        Case conMenu_View_ToolBar_Button
            For Each objControl In Me.cbsThis(2).Controls
                If objControl.Type <> xtpControlLabel Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each objControl In Me.cbsThis(2).Controls
                If objControl.Type <> xtpControlLabel Then
                    objControl.STYLE = IIf(objControl.STYLE = xtpButtonCaption, xtpButtonIconAndCaption, xtpButtonCaption)
                End If
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            Call cbsThis_ChangeCaption(Control, "��ͼ��(&D)", "Сͼ��(&D)")
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Visible
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Call cbsThis_ChangeCaption(Control, "���ع�����(&V)", "��ʾ������(&V)")
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_Refresh
            Call SetRefresh
        Case conMenu_View_Choose
            Call SetVfgSelect
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_ReportView      '������ҳ
            Call RecordLook
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Plan   '������
            Call SetAdd
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Refuse '�ܾ����
            With vfgList
                strMsg = "ȷ�Ͼ������²�����?" & vbCrLf & vbCrLf
                strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
                If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                    If frmPubNoteEdit.ShowNoteEdit(Me, "����ܾ�����", strNote) Then
                         If strNow = "" Then strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                         
                         If Val(.TextMatrix(.Row, .ColIndex("�ύID"))) > 0 And Val(.TextMatrix(.Row, .ColIndex("����״ֵ̬"))) = 1 Then
                                strSQL = "zl_�����ύ��¼_Refuse('" & Val(.TextMatrix(.Row, .ColIndex("�ύID"))) & "','" & UserInfo.���� & "',To_Date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),'" & strNote & "')"
                                Call SQLRecordAdd(rsSQL, strSQL)
                         End If
                         Call SQLRecordExecute(rsSQL, Me.Caption)
                         Call SetRefresh
                    End If
                End If
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            Select Case Control.Caption
            Case "���˽���", "���˽���(&1)"
                With vfgList
                    strMsg = "ȷ�ϻ��˽������²�����?" & vbCrLf & vbCrLf
                    strMsg = strMsg & ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
                    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, ParamInfo.ϵͳ����) = vbYes Then
                        strSQL = "zl_�����ύ��¼_UnReceive('" & Val(.TextMatrix(.Row, .ColIndex("�ύID"))) & "')"
                        Call SQLRecordAdd(rsSQL, strSQL)
                        
                        '�Ѿ���װ�˲���,���������ռ�¼
                        If mbln����ϵͳ Then
                            strSQL = "Zl_���Ӳ������ռ�¼_Delete(" & Val(.TextMatrix(.Row, .ColIndex("����ID"))) & "," & Val(.TextMatrix(.Row, .ColIndex("��ҳID"))) & ")"
                            Call SQLRecordAdd(rsSQL, strSQL)
                        End If
                        
                        Call SQLRecordExecute(rsSQL, Me.Caption)
                        
                        Call SetRefresh
                    End If
                    
                End With
            Case "���˾ܾ�", "���˾ܾ�(&2)"
                
                
            End Select
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home:  Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Forum: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_Web_Mail:  Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:     Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            If Control.ID > 401 And Control.ID < 499 Then
            '��ر���ִ��
            Call OpenRpt(Control)
        End If
    End Select
    
    GoTo endHand
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
End Sub

Private Function OpenRpt(ByVal Control As XtremeCommandBars.ICommandBarControl) As Boolean
    '------------------------------------------------------------------------------
    '����:�򿪱���
    '����:Control-ִ�б���Ŀؼ�
    '����:
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim arrData As Variant
    Dim strDept As String
    Dim strTemp As String
    
    arrData = Split(Control.Parameter, ",")
    strTemp = cboOutDept.Text
    If strTemp = "���в���" Then
'        strDept = "���п���"
        Call ReportOpen(gcnOracle, Val(arrData(0)), arrData(1), Me, "��Ժ��ʼ����=" & CDate(Format(dtpBegin(0).Value, "yyyy-MM-dd")), _
                "��Ժ��������=" & CDate(Format(dtpEnd(0).Value, "yyyy-MM-dd")), "��Ժ����=" & "is not null")
    Else
'        strDept = Split(strTemp, "-")(1)
        Call ReportOpen(gcnOracle, Val(arrData(0)), arrData(1), Me, "��Ժ��ʼ����=" & CDate(Format(dtpBegin(0).Value, "yyyy-MM-dd")), _
                "��Ժ��������=" & CDate(Format(dtpEnd(0).Value, "yyyy-MM-dd")), "��Ժ����=" & "=" & mlngApplyId)
    End If

End Function

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = stbThis.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '����Ȩ�޿���
    Dim strAuditingMan As String
    Dim strNo As String
    
    Select Case Control.ID
        Case conMenu_View_ToolBar_Visible
            Call cbsThis_ChangeCaption(Control, "���ع�����(&V)", "��ʾ������(&V)")
    End Select
        
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If

    On Error Resume Next
    If mlngModule = 201 Then '��������
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.vfgList.Rows <> 0)
            Case conMenu_Edit_NewItem
                If (InStr(1, mstrPrivs, ";����;") <> 0) Then
                    Control.Visible = True
                Else
                    Control.Enabled = False
                    Control.Visible = False
                End If
            Case conMenu_Edit_Modify
                If InStr(mstrPrivs, ";�޸�;") = 0 Then
                    Control.Enabled = False
                    Control.Visible = False
                    Exit Sub
                End If
                Call SetVerify_Update(Control)
                If Control.Enabled And Control.Visible Then
                    mintDblClick = 1
                Else
                    mintDblClick = 0
                End If
            Case conMenu_Edit_Delete
                If InStr(mstrPrivs, ";ɾ��;") = 0 Then
                    Control.Enabled = False
                    Control.Visible = False
                    Exit Sub
                End If
                Call SetVerify_Update(Control)
            Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
            Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonCaption)
            Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
            Case conMenu_View_ToolBar_Visible: Control.Checked = Me.cbsThis(2).Visible
            Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
        End Select
    Else '���Ӳ�������
        Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.vfgList.Rows <> 0)
        Case conMenu_Edit_NewItem
            If (InStr(1, mstrPrivs, "����;") <> 0) Then
                Control.Visible = True
            Else
                Control.Enabled = False
                Control.Visible = False
            End If
        Case conMenu_Edit_Modify
            If InStr(mstrPrivs, "����;") = 0 Then
                Control.Enabled = False
                Control.Visible = False
                Exit Sub
            End If
            Call SetVerify_Update(Control)
            If Control.Enabled And Control.Visible Then
                mintDblClick = 1
            Else
                mintDblClick = 0
            End If
        Case conMenu_Edit_Delete
            If InStr(mstrPrivs, "����;") = 0 Then
                Control.Enabled = False
                Control.Visible = False
                Exit Sub
            End If
            Call SetVerify_Update(Control)
        Case conMenu_File_Parameter '����
            Control.Enabled = (InStr(1, mstrPrivs, "��������;") <> 0)
        Case conMenu_Manage_ReportView '���Ĳ���
            Control.Enabled = (InStr(1, mstrPrivs, "���ĵ��Ӳ���;") <> 0)
        Case conMenu_Manage_ReportView      '������ҳ
            
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Plan   '������
            Control.Visible = IsPrivs(mstrPrivs, "������")
            Control.Enabled = Control.Visible
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Manage_Refuse '�ܾ����
            Control.Visible = IsPrivs(mstrPrivs, "�ܾ����")
            
            If Me.vfgList.Rows = 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = (Control.Visible And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("�ύID"))) > 0 And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("����״ֵ̬"))) = 1)
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_Edit_Untread
            Select Case Control.Caption
            Case "���˽���", "���˽���(&1)"
                 Control.Visible = IsPrivs(mstrPrivs, "���˽���")
            Case "���˾ܾ�", "���˾ܾ�(&2)"
                 Control.Visible = False ' IsPrivs(mstrPrivs, "���˾ܾ�")
            End Select
            

            Select Case Control.Caption
                Case "���˽���", "���˽���(&1)"
                
                    If vfgList.ColIndex("����״ֵ̬") = -1 Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Control.Visible And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("�ύID"))) > 0 And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("����״ֵ̬"))) = 10)
                    End If

                Case "���˾ܾ�", "���˾ܾ�(&2)"
                
                    If vfgList.ColIndex("����״ֵ̬") = -1 Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Control.Visible And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("�ύID"))) > 0 And Val(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("����״ֵ̬"))) = 2)
                    End If
            End Select

        Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonCaption)
        Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_ToolBar_Visible: Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
    End If
End Sub

Private Sub cbsThis_ChangeCaption(ByVal Control As XtremeCommandBars.ICommandBarControl, OldCaption As String, NewCaption As String)
    Select Case Control.ID
        Case conMenu_View_ToolBar_Visible
            If Me.cbsThis(2).Visible Then
                If Control.Caption = NewCaption Then Control.Caption = OldCaption
            Else
                If Control.Caption = OldCaption Then Control.Caption = NewCaption
            End If
        Case Else
            If Control.Caption = OldCaption Then
                Control.Caption = NewCaption
            Else
                If Control.Caption = NewCaption Then Control.Caption = OldCaption
            End If
    End Select
End Sub

Private Sub chkIncept_Click()
    dtpBegin(1).Enabled = IIf(chkIncept.Value = 1, True, False)
    dtpEnd(1).Enabled = IIf(chkIncept.Value = 1, True, False)
'    lbldate(1).Enabled = IIf(chkIncept.Value = 1, True, False)
    lblTo(2).Enabled = IIf(chkIncept.Value = 1, True, False)
    If chkIncept.Value = 1 Then
        chkRecord.Value = 0
        dtpBegin(2).Enabled = IIf(chkRecord.Value = 1, True, False)
        dtpEnd(2).Enabled = IIf(chkRecord.Value = 1, True, False)
        lblTo(3).Enabled = IIf(chkRecord.Value = 1, True, False)
    End If
End Sub

Private Sub chkIncept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkOutHp_Click()
    dtpBegin(0).Enabled = IIf(chkOutHp.Value = 1, True, False)
    dtpEnd(0).Enabled = IIf(chkOutHp.Value = 1, True, False)
    lblTo(1).Enabled = IIf(chkOutHp.Value = 1, True, False)
End Sub

Private Sub chkOutHp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chkRecord_Click()
    dtpBegin(2).Enabled = IIf(chkRecord.Value = 1, True, False)
    dtpEnd(2).Enabled = IIf(chkRecord.Value = 1, True, False)
'    lbldate(2).Enabled = IIf(chkRecord.Value = 1, True, False)
    lblTo(3).Enabled = IIf(chkRecord.Value = 1, True, False)
    If chkRecord.Value = 1 Then
        chkIncept.Value = 0
        dtpBegin(1).Enabled = IIf(chkIncept.Value = 1, True, False)
        dtpEnd(1).Enabled = IIf(chkIncept.Value = 1, True, False)
        lblTo(2).Enabled = IIf(chkIncept.Value = 1, True, False)
    End If
End Sub

Private Sub chkRecord_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdAll_Click()
    Dim lngRows As Long
    Dim i As Long
    Dim lngApplyId As Long
    Dim strStatus As String
    Dim strTemp As String
    
    If mlngModule = 201 Then '���Ӳ���
        strTemp = ";����;"
    Else
        strTemp = "����;"
    End If
    
    With vfgList
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                lngApplyId = Val(.TextMatrix(i, .ColIndex("��Ժ����id")))
                strStatus = Trim(.TextMatrix(i, .ColIndex("״̬")))
                If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "δ����" Then
                    mlngTempId = cboOutDept.ItemData(cboOutDept.ListIndex)
                    If lngApplyId = mlngTempId Or mlngTempId = 1 Then
                        .TextMatrix(i, .ColIndex("ѡ��")) = -1
                    Else
                        .TextMatrix(i, .ColIndex("ѡ��")) = 0
                    End If
                Else
                    .TextMatrix(i, .ColIndex("ѡ��")) = 0
                End If
            Next
        End If
    End With
End Sub

Private Sub cmdNoAll_Click()
    Dim lngRows As Long
    Dim i As Long
    Dim strStatus As String
    Dim strTemp As String
    
    If mlngModule = 201 Then '���Ӳ���
        strTemp = ";����;"
    Else
        strTemp = "����;"
    End If
    
    With vfgList
         If .Rows > 1 Then
            lngRows = .Rows - 1
                For i = 1 To lngRows
            
                strStatus = Trim(.TextMatrix(i, .ColIndex("״̬")))
                If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "δ����" Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = 0
                    mlngTempId = 0
'                Else
'                    .TextMatrix(i, .ColIndex("ѡ��")) = 0
                End If
            Next
           
        End If
    End With
End Sub

Private Sub cmdRefresh_Click()
    Call SetRefresh
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picFilter.hWnd
        Case 2
            Item.Handle = picList.hWnd
    End Select
End Sub

Private Sub dtpBegin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtpEnd_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        If mstrMsg = "" Then Unload Me: Exit Sub
        MsgBox mstrMsg, vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    mstrPrivs = gstrPrivs
   ' mstrPrivs = "����;��������;������;���˾ܾ�;���˽���;�ܾ����;���п���;���ĵ��Ӳ���;" '����ʹ��
    If ParamInfo.ģ��� = 201 Then
        mlngModule = 201
        Me.Caption = "�������չ���"
    Else
        mlngModule = ParamInfo.ģ���
        Me.Caption = "���Ӳ�������"
        
        Set rs = GetMedicalExits
        If Not rs.EOF Then
            mbln����ϵͳ = True
        Else
            mbln����ϵͳ = False
        End If
        
    End If
    
    
    mblnBootUp = False
    mlngApplyId = 0
    mlngTempId = 0
    mintDblClick = 0
    mintDelete = 0
    
    If mlngModule = 201 Then   '�������ղ˵�
        Call InitMenus
    Else
        Call InitMenusElectron '���Ӳ������ղ˵�
    End If
    Call InitDkpMain
       
    If Not GetInitDept Then Exit Sub
    
    mblnShow = IIf(Val(zlDatabase.GetPara("����ʾ��ʷ�ѱ�Ŀδ�Ǽ�", glngSys, mlngModule)) = 1, 1, 0) = 1
    If mblnShow Then Call GetNoSHowData
    
    mdtInBeginDate = Format(DateAdd("d", -7, zlDatabase.Currentdate), "yyyy-MM-dd")
    mdtInEndDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtRecBeginDate = "1901-01-01"
    mdtRecEndDate = "1901-01-01"
    mdtOutBeginDate = "1901-01-01"
    mdtOutEndDate = "1901-01-01"

    dtpBegin(0).Value = mdtInBeginDate
    dtpBegin(1).Value = mdtInBeginDate
    dtpBegin(2).Value = mdtInBeginDate
    dtpEnd(0).Value = mdtInEndDate
    dtpEnd(1).Value = mdtInEndDate
    dtpEnd(2).Value = mdtInEndDate
    
    mstrListTitle = "�����������Ͳ��˲������"
    Call SetRefresh
    
    If mlngModule = 201 Then   '�������ղ˵�
        If InStr(mstrPrivs, ";����;") <> 0 Then
            With vfgList
                .Editable = flexEDKbdMouse
            End With
            cmdAll.Visible = True
            cmdNoAll.Visible = True
        Else
            cmdAll.Visible = False
            cmdNoAll.Visible = False
        End If
    Else
        If InStr(mstrPrivs, "����;") <> 0 Then
            With vfgList
                .Editable = flexEDKbdMouse
            End With
            cmdAll.Visible = True
            cmdNoAll.Visible = True
        Else
            cmdAll.Visible = False
            cmdNoAll.Visible = False
        End If

    End If
    
    Call zlDatabase.ShowReportMenu(Me, 300, mlngModule, gstrPrivs)
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Err = 0
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 13840 Then Me.Width = 13840
    If Me.Height < 8660 Then Me.Height = 8660  '9020 -360
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SavedkpMain(Me.Caption & "-" & mlngModule, "dkpMain")
    Call SaveHead(vfgList, 1)
    SaveWinState Me, App.ProductName
    Unload Me
End Sub

Private Sub picFilter_Resize()
    Dim lngWidth As Long
    Dim lngTemp As Long
    On Error Resume Next
    lngWidth = picFilter.ScaleWidth
    
    With fraPatient
        lngTemp = lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngWidth - .Left
        Else
            .Width = 0
        End If
    End With
    
    With fraSong
        lngTemp = lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngWidth - .Left
        Else
            .Width = 0
        End If
    End With
    
    With cboOutDept
        lngTemp = fraPatient.Width - .Left - 200
        If lngTemp > 0 Then
            .Width = lngTemp
        Else
            .Width = 0
        End If
        .SelLength = 0
    End With
    
    With dtpBegin(0)
        lngTemp = fraPatient.Width - .Left - 200
        If lngTemp > 0 Then
            .Width = lngTemp
        Else
            .Width = 0
        End If
        dtpEnd(0).Width = .Width
        dtpBegin(1).Width = .Width
        dtpEnd(1).Width = .Width
        dtpBegin(2).Width = .Width
        dtpEnd(2).Width = .Width
        txtBeginNo.Width = .Width
        txtEndNo.Width = .Width
    End With
    
'    lblApplyman.Top = 4845
    With txtApplyman
        .Top = lblApplyman.Top - 45
        lngTemp = fraPatient.Width - .Left - 200 'lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngTemp 'lngWidth - .Left
        Else
            .Width = 0
        End If
    End With
    
    With txtAuditingMan
        lngTemp = fraPatient.Width - .Left - 200 'lngWidth - .Left
        If lngTemp > 0 Then
            .Width = lngTemp 'lngWidth - .Left
        Else
            .Width = 0
        End If
'        If lngWidth - .Left > 0 Then
'            .Width = lngWidth - .Left
'        Else
'            .Width = 0
'        End If
    End With
    
    With cmdRefresh
        .Top = 5520 '6120
        lngTemp = lngWidth - .Width - 100
        If lngTemp > 0 Then
            .Left = lngTemp
            If .Left - 100 - cmdNoAll.Width > 0 Then
                cmdNoAll.Left = .Left - 100 - cmdNoAll.Width
                If cmdNoAll.Left - 100 - cmdAll.Width > 0 Then
                    cmdAll.Left = cmdNoAll.Left - 100 - cmdAll.Width
                Else
                    cmdAll.Left = 0
                End If
            Else
                cmdNoAll.Left = 0
                cmdAll.Left = 0
            End If
        Else
            .Left = 0
            cmdNoAll.Left = 0
            cmdAll.Left = 0
        End If
    End With
    
End Sub

Private Sub picList_Resize()
    vfgList.Top = picList.ScaleTop
    vfgList.Left = picList.ScaleLeft
    vfgList.Width = picList.ScaleWidth
    vfgList.Height = picList.ScaleHeight
End Sub

Private Sub initVfgList()
    Dim strHead As String
    
    strHead = "���,500,4,1;ѡ��,500,4,1;סԺ��,1500,1,1;����,900,1,0;�Ա�,500,4,0;����,500,7,0;סԺ����,900,7,0;��Ժ����,1200,1,0;" & _
              "��Ժʱ��,1100,1,0;��Ժ����,1200,1,0;��Ժʱ��,1100,1,0;����״̬,1100,1,0;������,900,1,0;������,900,1,0;����ʱ��,1100,1,0;��Ŀʱ��,1100,1,0;��¼ʱ��,1100,1,0;" & _
              "״̬,800,4,0;��������,1100,1,0;��ͥ��ַ,1350,1,0;��Ժ����id,0,7,-1;����ID,0,7,-1;��ҳid,0,7,-1;����״ֵ̬,0,7,-1;�ύID,0,7,-1;�ύ����,0,7,-1"
    Call SetVsFlexGridChangeHead(strHead, vfgList, 1)
End Sub

Private Sub SetInitVfgListFormat(ByVal vsGrid As VSFlexGrid)
    Dim i As Long
    With vsGrid
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
        .ForeColorSel = .CellForeColor
        .ExplorerBar = flexExSortShowAndMove
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub GetListData()
'�������ղ�ѯ
    Dim strSQL As String
    Dim strBillHead As String
    Dim rsTemp As New ADODB.Recordset
    Dim strBeginNo As String
    Dim strEndNo As String
    Dim strApplyMan As String
    Dim strAuditingMan As String
    Dim strNo As String
    Dim strStatus As String
    Dim i As Long
    Dim strFind As String
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngApplyId As Long
    Dim cllTemp As New Collection
    
    strBeginNo = Trim(txtBeginNo.Text)
    strEndNo = Trim(txtEndNo.Text)
    strApplyMan = Trim(txtApplyman.Text)
    strAuditingMan = Trim(txtAuditingMan.Text)
            
    If gstrMatchMethod = "0" Then
        LfPBF = "%"
        RgPbf = "%"
    Else
        LfPBF = ""
        RgPbf = "%"
    End If
        
    If Trim(strBeginNo) <> "" Then
        If InStr(1, strBeginNo, "'") <> 0 Then
            MsgBox "��ʼסԺ���к��зǷ��ַ���", vbInformation, gstrSysName
            If Me.txtBeginNo.Enabled Then Me.txtBeginNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strBeginNo)) > 18 Then
            MsgBox "��ʼסԺ�ų���,���������9�����ֻ�18���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtBeginNo.Enabled Then txtBeginNo.SetFocus
            Exit Sub
        End If
    End If
    If Trim(strEndNo) <> "" Then
        If InStr(1, strEndNo, "'") <> 0 Then
            MsgBox "����סԺ���к��зǷ��ַ���", vbInformation, gstrSysName
            If Me.txtEndNo.Enabled Then Me.txtEndNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(txtEndNo)) > 18 Then
            MsgBox "����סԺ�ų���,���������9�����ֻ�18���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtEndNo.Enabled Then txtEndNo.SetFocus
            Exit Sub
        End If
    End If
    
     If Trim(strApplyMan) <> "" Then
        If InStr(1, strApplyMan, "'") <> 0 Then
            MsgBox "�������к��зǷ��ַ���", vbInformation, gstrSysName
            If Me.txtApplyman.Enabled Then Me.txtApplyman.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strApplyMan)) > 20 Then
            MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtApplyman.Enabled Then txtApplyman.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(strAuditingMan) <> "" Then
       If InStr(1, strAuditingMan, "'") <> 0 Then
           MsgBox "�������к��зǷ��ַ���", vbInformation, gstrSysName
           If Me.txtAuditingMan.Enabled Then Me.txtAuditingMan.SetFocus
           Exit Sub
       End If
       If zlCommFun.ActualLen(Trim(txtEndNo)) > 20 Then
            MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtAuditingMan.Enabled Then txtAuditingMan.SetFocus
            Exit Sub
        End If
    End If
    
    mdtInBeginDate = "1901-01-01"
    mdtInEndDate = "1901-01-01"
    mdtRecBeginDate = "1901-01-01"
    mdtRecEndDate = "1901-01-01"
    mdtOutBeginDate = "1901-01-01"
    mdtOutEndDate = "1901-01-01"
        
    If strEndNo <> "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.סԺ�� Between " & strBeginNo & " And " & strEndNo
        strFind = strFind & " and A.סԺ�� Between  [1]  And [2] "
    End If
    
    If strEndNo <> "" And strBeginNo = "" Then
'        strFind = strFind & " and A.סԺ�� = " & strEndNo
        strFind = strFind & " and A.סԺ�� = [2] "
    End If
    
    If strEndNo = "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.סԺ�� = " & strBeginNo
        strFind = strFind & " and A.סԺ�� = [1] "
    End If
        
    If strApplyMan <> "" Then
'        strFind = strFind & " and A.������ like '" & LfPBF & strApplyMan & RgPbf & "'"
        strFind = strFind & " and A.������ like [3]"
    End If
    If strAuditingMan <> "" Then
'       strFind = strFind & " and A.������ like '" & LfPBF & strAuditingMan & RgPbf & "'"
       strFind = strFind & " and A.������ like [4]"
    End If
    If strBeginNo <> "" Then
        AddArray cllTemp, strBeginNo
    Else
        AddArray cllTemp, "0"
    End If
    If strEndNo <> "" Then
        AddArray cllTemp, strEndNo
    Else
        AddArray cllTemp, "0"
    End If
'    AddArray cllTemp, strBeginNo
'    AddArray cllTemp, strEndNo
    AddArray cllTemp, LfPBF & strApplyMan & RgPbf
    AddArray cllTemp, LfPBF & strAuditingMan & RgPbf
    
    
    If chkIncept.Value = 1 And chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS'))"
'            strFind = strFind & " Or (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
'
            strFind = strFind & " And ((A.����ʱ�� Between [7] And [8]) " _
            & " or (A.��Ŀ���� Between To_Date([9]) And [10])"
            strFind = strFind & " Or (A.��Ժ���� Between [5] And [6])) "
            
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And ((A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')))"
            strFind = strFind & " And ((A.����ʱ�� Between [7] And [8]) " _
            & " or (A.��Ŀ���� Between [9] And [10]))"
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
    ElseIf chkIncept.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.����ʱ�� Between [7] And [8]) "
            strFind = strFind & " Or (A.��Ժ���� Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS') "
            strFind = strFind & " And A.����ʱ�� Between [7] And [8] "
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
    ElseIf chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.��Ŀ���� Between [9] And [10]) "
            strFind = strFind & " Or (A.��Ժ���� Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And (A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
            strFind = strFind & " And (A.��Ŀ���� Between [9] And [10]) "
        End If
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
'    Else
'        strFind = strFind & " And A.����ʱ�� is null and A.��¼ʱ�� is null "
    End If
    
    If chkOutHp.Value = 1 And chkIncept.Value = 0 And chkRecord.Value = 0 Then
'        strFind = strFind & " And (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
        strFind = strFind & " And (A.��Ժ���� Between [5] And [6]) "
        mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
        mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
    End If
    
    AddArray cllTemp, Format(mdtOutBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtOutEndDate, "yyyy-mm-dd") & " 23:59:59 " ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecBeginDate, "yyyy-mm-dd") & " 00:00:00" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"

    
    If mintDelete = 1 Then
        strFind = mstrFind
        Set cllTemp = mcllTemp
    Else
        mstrFind = strFind
        Set mcllTemp = cllTemp
        If chkIncept.Value = 0 And chkRecord.Value = 0 And chkOutHp.Value = 0 Then
            MsgBox "����ѡ��һ����Ժ���ڻ��߽������ڻ��߼�¼����!", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    vfgList.Redraw = False
    Call zlCommFun.ShowFlash("��������������Ӧ�ļ�¼,���Ժ� ...", Me)
'    DoEvents
    Screen.MousePointer = vbHourglass
    
'     And (U.��Ŀ���� is null Or (U.��Ŀ���� is not null and U.��Ժ���� >= to_date('" & strNoShowDate & "','yyyy-mm-dd')))
    mlngTempId = 0
    
    If mlngApplyId = 0 Then
        strBillHead = "" & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, Null As ������, Null As ������, Null As ����ʱ��, Null As ��¼ʱ��," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', 'δ����', '�ѱ�Ŀδ�Ǽ�') As ״̬ " & _
        " From ������ҳ U, ������Ϣ X " & _
        " Where U.����ID = X.����ID And Not Exists" & _
        "      (Select 1 From �������ռ�¼ A Where A.����id = U.����id And A.��ҳid = U.��ҳid) And U.�������� = 0 And U.��ҳID <> 0 And U.��Ժ���� is not null " & _
               IIf(mblnShow, " And (U.��Ŀ���� is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', '�ѽ���', '�ѱ�Ŀ') As ״̬" & _
        " From ������ҳ U, ������Ϣ X,�������ռ�¼ A" & _
        " Where U.����id = X.����id And U.�������� = 0 And U.��ҳID <> 0 And " & _
        "       A.����id = U.����id And A.��ҳID = U.��ҳID"
        strSQL = "" & _
        "   Select distinct A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����,��Ժ����id," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��,A.״̬" & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & strFind & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
    Else
        strBillHead = "" & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, Null As ������, Null As ������, Null As ����ʱ��, Null As ��¼ʱ��," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', 'δ����', '�ѱ�Ŀδ�Ǽ�') As ״̬ " & _
        " From ������ҳ U, ������Ϣ X " & _
        " Where U.����ID = X.����ID And U.��Ժ����id =[11] And Not Exists" & _
        "      (Select 1 From �������ռ�¼ A Where A.����id = U.����id And A.��ҳid = U.��ҳid) And U.�������� = 0 And U.��ҳID <> 0 And U.��Ժ���� is not null " & _
               IIf(mblnShow, " And (U.��Ŀ���� is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', '�ѽ���', '�ѱ�Ŀ') As ״̬" & _
        " From ������ҳ U, ������Ϣ X,�������ռ�¼ A" & _
        " Where U.����id = X.����id And U.�������� = 0 And U.��ҳID <> 0 And U.��Ժ����id =[11] And " & _
        "       A.����id = U.����id And A.��ҳID = U.��ҳID"
        strSQL = "" & _
        "   Select distinct A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����,��Ժ����id," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��,A.״̬" & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & strFind & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
    End If
            
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(cllTemp(1)), CLng(cllTemp(2)), cllTemp(3), cllTemp(4), CDate(cllTemp(5)), _
                        CDate(cllTemp(6)), CDate(cllTemp(7)), CDate(cllTemp(8)), CDate(cllTemp(9)), CDate(cllTemp(10)), mlngApplyId)
    With vfgList
        Call initVfgList
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("���")) = i
                .TextMatrix(i, .ColIndex("סԺ��")) = IIf(IsNull(rsTemp!סԺ��), 0, rsTemp!סԺ��)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("�Ա�")) = IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("סԺ����")) = IIf(IsNull(rsTemp!��סԺ����), "", rsTemp!��סԺ����)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("������")) = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                .TextMatrix(i, .ColIndex("������")) = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                .TextMatrix(i, .ColIndex("����ʱ��")) = IIf(IsNull(rsTemp!����ʱ��), "", Format(rsTemp!����ʱ��, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��Ŀʱ��")) = IIf(IsNull(rsTemp!��Ŀ����), "", Format(rsTemp!��Ŀ����, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��¼ʱ��")) = IIf(IsNull(rsTemp!��¼ʱ��), "", Format(rsTemp!��¼ʱ��, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("״̬")) = IIf(IsNull(rsTemp!״̬), "", rsTemp!״̬)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("��������")) = IIf(IsNull(rsTemp!��������), "", Format(rsTemp!��������, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��ͥ��ַ")) = IIf(IsNull(rsTemp!��ͥ��ַ), "", rsTemp!��ͥ��ַ)
                .TextMatrix(i, .ColIndex("����ID")) = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
                .TextMatrix(i, .ColIndex("��ҳid")) = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
                .TextMatrix(i, .ColIndex("��Ժ����id")) = IIf(IsNull(rsTemp!��Ժ����ID), 0, rsTemp!��Ժ����ID)
                
                rsTemp.MoveNext
                strStatus = Trim(.TextMatrix(i, .ColIndex("״̬")))
                Select Case strStatus
                Case "δ����"
                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(0)
                Case "�ѽ���"
                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(1)
                Case "�ѱ�Ŀ", "�ѱ�Ŀδ�Ǽ�"
                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(2)
                End Select

                If InStr(mstrPrivs, ";����;") <> 0 And strStatus = "δ����" Then
'                    lngApplyId = Val(.TextMatrix(i, .ColIndex("��Ժ����id")))
'                    If mlngTempId = 0 Then mlngTempId = lngApplyId
'                    If lngApplyId <> mlngTempId Then
                        .TextMatrix(i, .ColIndex("ѡ��")) = 0
'                    Else
'                        .TextMatrix(i, .ColIndex("ѡ��")) = -1
'                    End If
                    
                Else
                    .TextMatrix(i, .ColIndex("ѡ��")) = 0
                End If
            Next
        End If


    End With
    
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    Call GetStatusCount
   
    Call SetInitVfgListFormat(vfgList)
    Call RestoreHead(vfgList, 1)
    rsTemp.Close
    vfgList.Redraw = True
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetStatusCount()
    Dim lngδ���� As Long
    Dim lng�ѽ��� As Long
    Dim lng�ѱ�Ŀ As Long
    Dim lngCount As Long
    
    With vfgList
        For lngCount = 1 To .Rows - 1
            Select Case Trim(.TextMatrix(lngCount, .ColIndex("״̬")))
            Case "δ����"
                lngδ���� = lngδ���� + 1
            Case "�ѽ���"
                lng�ѽ��� = lng�ѽ��� + 1
            Case "�ѱ�Ŀ", "�ѱ�Ŀδ�Ǽ�"
                lng�ѱ�Ŀ = lng�ѱ�Ŀ + 1
            End Select
        Next
    End With
    
    stbThis.Panels(2).Text = "ֻ��״̬Ϊδ���ռ������Ժ����ʱ���ܽ���ѡ���������ǰ����" & vfgList.Rows - 1 & "�����˲���������δ����:" & lngδ���� & "�����ѽ���:" & lng�ѽ��� & "�����ѱ�Ŀ:" & lng�ѱ�Ŀ & "����"
End Sub

Private Sub GetListDataElectron()
'���Ӳ������ղ�ѯ
    Dim strSQL As String
    Dim strBillHead As String
    Dim rsTemp As New ADODB.Recordset
    Dim strBeginNo As String
    Dim strEndNo As String
    Dim strApplyMan As String
    Dim strAuditingMan As String
    Dim strNo As String
    Dim strStatus As String
    Dim i As Long
    Dim strFind As String
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngApplyId As Long
    Dim cllTemp As New Collection
    
    strBeginNo = Trim(txtBeginNo.Text)
    strEndNo = Trim(txtEndNo.Text)
    strApplyMan = Trim(txtApplyman.Text)
    strAuditingMan = Trim(txtAuditingMan.Text)
            
    If gstrMatchMethod = "0" Then
        LfPBF = "%"
        RgPbf = "%"
    Else
        LfPBF = ""
        RgPbf = "%"
    End If
        
    If Trim(strBeginNo) <> "" Then
        If InStr(1, strBeginNo, "'") <> 0 Then
            MsgBox "��ʼסԺ���к��зǷ��ַ���", vbInformation, gstrSysName
            If Me.txtBeginNo.Enabled Then Me.txtBeginNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strBeginNo)) > 18 Then
            MsgBox "��ʼסԺ�ų���,���������9�����ֻ�18���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtBeginNo.Enabled Then txtBeginNo.SetFocus
            Exit Sub
        End If
    End If
    If Trim(strEndNo) <> "" Then
        If InStr(1, strEndNo, "'") <> 0 Then
            MsgBox "����סԺ���к��зǷ��ַ���", vbInformation, gstrSysName
            If Me.txtEndNo.Enabled Then Me.txtEndNo.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(txtEndNo)) > 18 Then
            MsgBox "����סԺ�ų���,���������9�����ֻ�18���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtEndNo.Enabled Then txtEndNo.SetFocus
            Exit Sub
        End If
    End If
    
     If Trim(strApplyMan) <> "" Then
        If InStr(1, strApplyMan, "'") <> 0 Then
            MsgBox "�������к��зǷ��ַ���", vbInformation, gstrSysName
            If Me.txtApplyman.Enabled Then Me.txtApplyman.SetFocus
            Exit Sub
        End If
        If zlCommFun.ActualLen(Trim(strApplyMan)) > 20 Then
            MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtApplyman.Enabled Then txtApplyman.SetFocus
            Exit Sub
        End If
    End If
    
    If Trim(strAuditingMan) <> "" Then
       If InStr(1, strAuditingMan, "'") <> 0 Then
           MsgBox "�������к��зǷ��ַ���", vbInformation, gstrSysName
           If Me.txtAuditingMan.Enabled Then Me.txtAuditingMan.SetFocus
           Exit Sub
       End If
       If zlCommFun.ActualLen(Trim(txtEndNo)) > 20 Then
            MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
            If txtAuditingMan.Enabled Then txtAuditingMan.SetFocus
            Exit Sub
        End If
    End If
    
    mdtInBeginDate = "1901-01-01"
    mdtInEndDate = "1901-01-01"
    mdtRecBeginDate = "1901-01-01"
    mdtRecEndDate = "1901-01-01"
    mdtOutBeginDate = "1901-01-01"
    mdtOutEndDate = "1901-01-01"
        
    If strEndNo <> "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.סԺ�� Between " & strBeginNo & " And " & strEndNo
        strFind = strFind & " and A.סԺ�� Between  [1]  And [2] "
    End If
    
    If strEndNo <> "" And strBeginNo = "" Then
'        strFind = strFind & " and A.סԺ�� = " & strEndNo
        strFind = strFind & " and A.סԺ�� = [2] "
    End If
    
    If strEndNo = "" And strBeginNo <> "" Then
'        strFind = strFind & " and A.סԺ�� = " & strBeginNo
        strFind = strFind & " and A.סԺ�� = [1] "
    End If
        
    If strApplyMan <> "" Then
'        strFind = strFind & " and A.������ like '" & LfPBF & strApplyMan & RgPbf & "'"
        strFind = strFind & " and A.������ like [3]"
    End If
    If strAuditingMan <> "" Then
'       strFind = strFind & " and A.������ like '" & LfPBF & strAuditingMan & RgPbf & "'"
       strFind = strFind & " and A.������ like [4]"
    End If
    If strBeginNo <> "" Then
        AddArray cllTemp, strBeginNo
    Else
        AddArray cllTemp, "0"
    End If
    If strEndNo <> "" Then
        AddArray cllTemp, strEndNo
    Else
        AddArray cllTemp, "0"
    End If
'    AddArray cllTemp, strBeginNo
'    AddArray cllTemp, strEndNo
    AddArray cllTemp, LfPBF & strApplyMan & RgPbf
    AddArray cllTemp, LfPBF & strAuditingMan & RgPbf
    
    
    If chkIncept.Value = 1 And chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS'))"
'            strFind = strFind & " Or (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
'
            strFind = strFind & " And ((A.����ʱ�� Between [7] And [8]) " _
            & " or (A.��Ŀ���� Between To_Date([9]) And [10])"
            strFind = strFind & " Or (A.��Ժ���� Between [5] And [6])) "
            
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And ((A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'            & " or (A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2), "yyyy-mm-dd") & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtpEnd(2), "yyyy-mm-dd") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')))"
            strFind = strFind & " And ((A.����ʱ�� Between [7] And [8]) " _
            & " or (A.��Ŀ���� Between [9] And [10]))"
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
    ElseIf chkIncept.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.����ʱ�� Between [7] And [8]) "
            strFind = strFind & " Or (A.��Ժ���� Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And A.����ʱ�� Between To_Date('" & Format(dtpBegin(1), "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(1), "yyyy-mm-dd") & " 23:59:59','yyyy-MM-dd HH24:MI:SS') "
            strFind = strFind & " And A.����ʱ�� Between [7] And [8] "
        End If
        mdtInBeginDate = Format(dtpBegin(1), "yyyy-mm-dd")
        mdtInEndDate = Format(dtpEnd(1), "yyyy-mm-dd")
    ElseIf chkRecord.Value = 1 Then
        If chkOutHp.Value = 1 Then
'            strFind = strFind & " And ((A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
'            strFind = strFind & " Or (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS'))) "
            strFind = strFind & " And ((A.��Ŀ���� Between [9] And [10]) "
            strFind = strFind & " Or (A.��Ժ���� Between [5] And [6])) "
            mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
            mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
        Else
'            strFind = strFind & " And (A.��Ŀ���� Between To_Date('" & Format(dtpBegin(2).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(2).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
            strFind = strFind & " And (A.��Ŀ���� Between [9] And [10]) "
        End If
        mdtRecBeginDate = Format(dtpBegin(2), "yyyy-mm-dd")
        mdtRecEndDate = Format(dtpEnd(2), "yyyy-mm-dd")
'    Else
'        strFind = strFind & " And A.����ʱ�� is null and A.��¼ʱ�� is null "
    End If
    
    If chkOutHp.Value = 1 And chkIncept.Value = 0 And chkRecord.Value = 0 Then
'        strFind = strFind & " And (A.��Ժ���� Between To_Date('" & Format(dtpBegin(0).Value, "yyyy-mm-dd") & " 00:00:00','yyyy-MM-dd HH24:MI:SS') And To_Date('" & Format(dtpEnd(0).Value, "YYYY-mm-dd") & " 23:59:59 ','yyyy-MM-dd HH24:MI:SS')) "
        strFind = strFind & " And (A.��Ժ���� Between [5] And [6]) "
        mdtOutBeginDate = Format(dtpBegin(0), "yyyy-mm-dd")
        mdtOutEndDate = Format(dtpEnd(0), "yyyy-mm-dd")
    End If
    
    AddArray cllTemp, Format(mdtOutBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtOutEndDate, "yyyy-mm-dd") & " 23:59:59 " ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInBeginDate, "yyyy-mm-dd") & " 00:00:00" ' ,yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtInEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecBeginDate, "yyyy-mm-dd") & " 00:00:00" ',yyyy-MM-dd HH24:MI:SS"
    AddArray cllTemp, Format(mdtRecEndDate, "yyyy-mm-dd") & " 23:59:59" ',yyyy-MM-dd HH24:MI:SS"

    
    If mintDelete = 1 Then
        strFind = mstrFind
        Set cllTemp = mcllTemp
    Else
        mstrFind = strFind
        Set mcllTemp = cllTemp
        If chkIncept.Value = 0 And chkRecord.Value = 0 And chkOutHp.Value = 0 Then
            MsgBox "����ѡ��һ����Ժ���ڻ��߽������ڻ��߼�¼����!", vbInformation, gstrSysName
            If chkOutHp.Enabled And chkOutHp.Visible Then
                chkOutHp.SetFocus
            End If
            Exit Sub
        End If
    End If
    vfgList.Redraw = False
    Call zlCommFun.ShowFlash("��������������Ӧ�ļ�¼,���Ժ� ...", Me)
    Screen.MousePointer = vbHourglass
    mlngTempId = 0
    
    If mlngApplyId = 0 Then
        strBillHead = "" & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, Null As ������, Null As ������, Null As ����ʱ��, Null As ��¼ʱ��,Decode(Nvl(U.����״̬,1),1,'�ύ����',10,'���մ���',2,'�ܾ�����',3,'�������',4,'��鷴��',5,'���鵵',6,'�������',13,'���ڳ��',14,'��鷴��',16,'�������') as ����״̬,Nvl(U.����״̬,1) as ����״ֵ̬,C.ID as �ύID,(Select Count(1) From �����ύ��¼ F Where F.����ID=U.����ID and F.��ҳID=U.��ҳID And F.��¼״̬=2) as �ύ����," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', 'δ����', '�ѱ�Ŀδ�Ǽ�') As ״̬ " & _
        " From ������ҳ U, ������Ϣ X,�����ύ��¼ C " & _
        " Where U.����ID = X.����ID And U.����ID = C.����ID And U.��ҳID = C.��ҳID And Not Exists" & _
        "      (Select 1 From �������ռ�¼ A Where A.����id = U.����id And A.��ҳid = U.��ҳid) And U.��ҳID <> 0 And U.��Ժ���� is not null And C.��¼״̬=1 " & _
               IIf(mblnShow, " And (U.��Ŀ���� is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��,Decode(Nvl(U.����״̬,1),1,'�ύ����',10,'���մ���',2,'�ܾ�����',3,'�������',4,'��鷴��',5,'���鵵',6,'�������',13,'���ڳ��',14,'��鷴��',16,'�������') as ����״̬,Nvl(U.����״̬,1) as ����״ֵ̬,C.ID as �ύID,0 as �ύ����," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', '�ѽ���', '�ѱ�Ŀ') As ״̬" & _
        " From ������ҳ U, ������Ϣ X,�������ռ�¼ A,�����ύ��¼ C" & _
        " Where U.����id = X.����id And U.��ҳID <> 0  And A.����ID = C.����ID And A.��ҳID = C.��ҳID And  C.��¼״̬<>2 And " & _
        "       A.����id = U.����id And A.��ҳID = U.��ҳID"
        
        strSQL = "" & _
        "   Select distinct A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����,��Ժ����id," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��,A.״̬,A.����״̬,A.����״ֵ̬,A.�ύID,A.�ύ����" & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & strFind & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
    Else
        strBillHead = "" & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, Null As ������, Null As ������, Null As ����ʱ��, Null As ��¼ʱ��, Decode(Nvl(U.����״̬,1),1,'�ύ����',10,'���մ���',2,'�ܾ�����',3,'�������',4,'��鷴��',5,'���鵵',6,'�������',13,'���ڳ��',14,'��鷴��',16,'�������') as ����״̬,Nvl(U.����״̬,1) as ����״ֵ̬,C.ID as �ύID,(Select Count(1) From �����ύ��¼ F Where F.����ID=U.����ID and F.��ҳID=U.��ҳID And F.��¼״̬=2) as �ύ����," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', 'δ����', '�ѱ�Ŀδ�Ǽ�') As ״̬ " & _
        " From ������ҳ U, ������Ϣ X,�����ύ��¼ C " & _
        " Where U.����ID = X.����ID And U.��Ժ����id =[11] And U.����ID = C.����ID And U.��ҳID = C.��ҳID And Not Exists" & _
        "      (Select 1 From �������ռ�¼ A Where A.����id = U.����id And A.��ҳid = U.��ҳid) And U.��ҳID <> 0 And U.��Ժ���� is not null And C.��¼״̬=1 " & _
               IIf(mblnShow, " And (U.��Ŀ���� is null )", "")
        strBillHead = strBillHead & " Union All "
        strBillHead = strBillHead & _
        " Select Distinct X.����id, U.��ҳid, U.סԺ��, X.����, X.�Ա�, X.����, U.��ҳid As ��סԺ����, X.��������, X.�����ص�," & _
        "      X.���֤��, X.ְҵ, X.����״��, X.��ͥ��ַ, X.��ͥ�绰, X.��ϵ������, X.��ϵ�˹�ϵ, X.��ϵ�˵�ַ," & _
        "      X.��ϵ�˵绰, X.������λ, X.��λ�绰, U.��Ժ����, U.��Ժ����, U.��Ժ����id, U.��Ժ����id," & _
        "      U.סԺ����, U.���ú�, Decode(U.�����־, 1, '��', 2, '��', 3, '��', '') As �Ƿ�����," & _
        "      U.��ĿԱ���� As ��ĿԱ, U.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��, Decode(Nvl(U.����״̬,1),1,'�ύ����',10,'���մ���',2,'�ܾ�����',3,'�������',4,'��鷴��',5,'���鵵',6,'�������',13,'���ڳ��',14,'��鷴��',16,'�������') as ����״̬,Nvl(U.����״̬,1) as ����״ֵ̬,C.ID as �ύID,0 as �ύ����," & _
        "      Decode(Nvl(to_char(U.��Ŀ����), '0'), '0', '�ѽ���', '�ѱ�Ŀ') As ״̬" & _
        " From ������ҳ U, ������Ϣ X,�������ռ�¼ A,�����ύ��¼ C" & _
        " Where U.����id = X.����id And U.��ҳID <> 0 And U.��Ժ����id =[11] And A.����ID = C.����ID And A.��ҳID = C.��ҳID And C.��¼״̬<>2 And  " & _
        "       A.����id = U.����id And A.��ҳID = U.��ҳID"
        strSQL = "" & _
        "   Select distinct A.����id, A.��ҳid, A.סԺ��, A.����, A.�Ա�, A.����, A.��סԺ����, A.��������, A.�����ص�," & _
        "    A.���֤��, A.ְҵ, A.����״��, A.��ͥ��ַ, A.��ͥ�绰, A.��ϵ������, A.��ϵ�˹�ϵ, A.��ϵ�˵�ַ," & _
        "    A.��ϵ�˵绰, A.������λ, A.��λ�绰, A.��Ժ����, A.��Ժ����, B1.���� As ��Ժ����, B2.���� As ��Ժ����,��Ժ����id," & _
        "    A.סԺ����, A.���ú�, A.�Ƿ�����,A.��ĿԱ, A.��Ŀ����, A.������, A.������, A.����ʱ��, A.��¼ʱ��,A.״̬,A.����״̬,A.����״ֵ̬,A.�ύID,A.�ύ����" & _
        "    From (" & strBillHead & ") A,���ű� B1, ���ű� B2" & _
        "    Where A.��Ժ����id=B1.id And A.��Ժ����id=B2.id " & strFind & zl_��ȡվ������(True, "B1") & zl_��ȡվ������(True, "B2") & _
        "    Order by A.��Ժ���� desc "
    End If
            
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(cllTemp(1)), CLng(cllTemp(2)), cllTemp(3), cllTemp(4), CDate(cllTemp(5)), _
                        CDate(cllTemp(6)), CDate(cllTemp(7)), CDate(cllTemp(8)), CDate(cllTemp(9)), CDate(cllTemp(10)), mlngApplyId)
    With vfgList
        Call initVfgList
        .Rows = IIf(rsTemp.EOF, 0, rsTemp.RecordCount) + 1
        If Not rsTemp.EOF Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("���")) = i
                .TextMatrix(i, .ColIndex("סԺ��")) = IIf(IsNull(rsTemp!סԺ��), 0, rsTemp!סԺ��)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("�Ա�")) = IIf(IsNull(rsTemp!�Ա�), "", rsTemp!�Ա�)
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("סԺ����")) = IIf(IsNull(rsTemp!��סԺ����), "", rsTemp!��סԺ����)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("��Ժ����")) = IIf(IsNull(rsTemp!��Ժ����), "", rsTemp!��Ժ����)
                .TextMatrix(i, .ColIndex("������")) = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                .TextMatrix(i, .ColIndex("������")) = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                .TextMatrix(i, .ColIndex("����ʱ��")) = IIf(IsNull(rsTemp!����ʱ��), "", Format(rsTemp!����ʱ��, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��Ŀʱ��")) = IIf(IsNull(rsTemp!��Ŀ����), "", Format(rsTemp!��Ŀ����, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��¼ʱ��")) = IIf(IsNull(rsTemp!��¼ʱ��), "", Format(rsTemp!��¼ʱ��, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("״̬")) = IIf(IsNull(rsTemp!״̬), "", rsTemp!״̬)
                .TextMatrix(i, .ColIndex("��Ժʱ��")) = IIf(IsNull(rsTemp!��Ժ����), "", Format(rsTemp!��Ժ����, "yyyy-MM-dd hh:mm:ss"))
                .TextMatrix(i, .ColIndex("����״̬")) = IIf(IsNull(rsTemp!����״̬), "", rsTemp!����״̬)
                .TextMatrix(i, .ColIndex("��������")) = IIf(IsNull(rsTemp!��������), "", Format(rsTemp!��������, "yyyy-MM-dd"))
                .TextMatrix(i, .ColIndex("��ͥ��ַ")) = IIf(IsNull(rsTemp!��ͥ��ַ), "", rsTemp!��ͥ��ַ)
                .TextMatrix(i, .ColIndex("����ID")) = IIf(IsNull(rsTemp!����ID), 0, rsTemp!����ID)
                .TextMatrix(i, .ColIndex("��ҳid")) = IIf(IsNull(rsTemp!��ҳID), 0, rsTemp!��ҳID)
                .TextMatrix(i, .ColIndex("��Ժ����id")) = IIf(IsNull(rsTemp!��Ժ����ID), 0, rsTemp!��Ժ����ID)
                .TextMatrix(i, .ColIndex("����״ֵ̬")) = IIf(IsNull(rsTemp!����״ֵ̬), 0, rsTemp!����״ֵ̬)
                .TextMatrix(i, .ColIndex("�ύID")) = IIf(IsNull(rsTemp!�ύId), 0, rsTemp!�ύId)
                .TextMatrix(i, .ColIndex("�ύ����")) = IIf(IsNull(rsTemp!�ύ����), 0, rsTemp!�ύ����)
                
                rsTemp.MoveNext
                strStatus = Trim(.TextMatrix(i, .ColIndex("״̬")))
                Select Case strStatus
                Case "δ����"
                    If Val(.TextMatrix(i, .ColIndex("�ύ����"))) = 0 Then
                        .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(0)
                    Else
                        .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(3)
                    End If
                Case "�ѽ���"
                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(1)
                Case "�ѱ�Ŀ", "�ѱ�Ŀδ�Ǽ�"
                    .Cell(flexcpPicture, i, .ColIndex("סԺ��")) = imgCelNo(2)
                End Select

                If InStr(mstrPrivs, "����;") <> 0 And strStatus = "δ����" Then
'                    lngApplyId = Val(.TextMatrix(i, .ColIndex("��Ժ����id")))
'                    If mlngTempId = 0 Then mlngTempId = lngApplyId
'                    If lngApplyId <> mlngTempId Then
                        .TextMatrix(i, .ColIndex("ѡ��")) = 0
'                    Else
'                        .TextMatrix(i, .ColIndex("ѡ��")) = -1
'                    End If
                    
                Else
                    .TextMatrix(i, .ColIndex("ѡ��")) = 0
                End If
            Next
        End If


    End With
    
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
'    stbThis.Panels(2).Text = "ֻ��״̬Ϊδ���ռ������Ժ����ʱ���ܽ���ѡ���������ǰ����" & rsTemp.RecordCount & "�����˲���������Ϣ��"
    Call GetStatusCount
    Call SetInitVfgListFormat(vfgList)
    Call RestoreHead(vfgList, 1)
    rsTemp.Close
    vfgList.Redraw = True
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetNoSHowData()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "" & _
        " Select min(U.��Ժ����) as ��Ժ����" & _
        " From ������ҳ U, �������ռ�¼ X " & _
        " Where U.����ID = X.����ID And U.��ҳid = X.��ҳid And X.����ʱ�� = " & _
        "      (Select min(A.����ʱ��) From �������ռ�¼ A)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstrNoShowDate = IIf(IsNull(rsTemp!��Ժ����), Format(zlDatabase.Currentdate, "yyyy-MM-dd"), Format(rsTemp!��Ժ����, "yyyy-MM-dd"))
    Else
        mstrNoShowDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End If
    rsTemp.Close
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
   
Private Sub SetRefresh()
    If mlngModule = 201 Then
        Call GetListData
    Else
        Call GetListDataElectron
    End If
End Sub

Private Sub SetVfgSelect()
    frm����ѡ����.ShowColSet Me, "������Ϣ������", vfgList
End Sub


Private Sub RightHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim strHeadInfo As String
    Dim vRect  As RECT
    vRect = GetControlRect(vsGrid.hWnd)
    lngLeft = vRect.Left + vsGrid.Left
    lngTop = vRect.Top + vsGrid.RowHeight(0) 'vsGrid.CellTop ' + vsGrid.CellHeight '
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, vsGrid.RowHeight(0))
    Call SaveHead(vsGrid, intListOrDetail)
End Sub

Private Sub SaveHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim strHeadInfo As String
     If intListOrDetail = 1 Then
        strHeadInfo = "����������ͷ��Ϣ"
    End If
    zl_VsGrid_SaveToPara vsGrid, Me.Caption, mlngModule, strHeadInfo, True, True
End Sub

Private Sub RestoreHead(ByVal vsGrid As VSFlexGrid, intListOrDetail As Integer)
    Dim strHeadInfo As String
    If intListOrDetail = 1 Then
        strHeadInfo = "����������ͷ��Ϣ"
    End If
    zl_VsGrid_FromParaRestore vsGrid, Me.Caption, mlngModule, strHeadInfo, True, True
End Sub

Private Sub txtApplyman_GotFocus()
    With txtApplyman
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtApplyman_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtApplyman.Text) = "" Then
            zlCommFun.PressKey (vbKeyTab)
            Exit Sub
        End If
        txtApplyman.Text = Replace(UCase(txtApplyman.Text), "'", "")
        vRect = GetControlRect(txtApplyman.hWnd)
        
        strSQL = "" & _
            "   Select ���,����,����,id " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) " & zl_��ȡվ������(True) & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
            
        strTemp = Trim(txtApplyman.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        lngHeigth = txtApplyman.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Աѡ��", False, txtApplyman.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
            If txtApplyman.Enabled Then
                zlCommFun.PressKey (vbKeyTab)
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtApplyman)) = "TEXTBOX" Then
                txtApplyman = IIf(IsNull(!����), "", !����)
                zlCommFun.PressKey (vbKeyTab)
            Else
                txtApplyman.SetFocus
                txtApplyman.SelStart = 0
                txtApplyman.SelLength = Len(txtApplyman.Text)
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtApplyman_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtApplyman_LostFocus()
    If zlCommFun.ActualLen(Trim(txtApplyman.Text)) > 20 Then
        MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
        txtApplyman.SetFocus
        txtApplyman.SelStart = 0
        txtApplyman.SelLength = Len(txtApplyman.Text)
        If txtApplyman.Enabled Then txtApplyman.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtAuditingMan_GotFocus()
    With txtAuditingMan
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAuditingMan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim vRect As RECT
    Dim LfPBF As String
    Dim RgPbf As String
    Dim lngHeigth As Long
    Dim blnCancel As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtAuditingMan.Text) = "" Then
            zlCommFun.PressKey (vbKeyTab)
            Exit Sub
        End If
        txtAuditingMan.Text = Replace(UCase(txtAuditingMan.Text), "'", "")
        vRect = GetControlRect(txtAuditingMan.hWnd)
        
        strSQL = "" & _
            "   Select ���,����,����,id " & _
            "   From ��Ա�� " & _
            "   Where (���� like [1] or ��� like [1] or ���� like [1] ) " & zl_��ȡվ������(True) & _
            "       and (����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
            "   order by ���"
            
        strTemp = Trim(txtAuditingMan.Text)
        
        If gstrMatchMethod = "0" Then
            LfPBF = "%"
            RgPbf = "%"
        Else
            LfPBF = ""
            RgPbf = "%"
        End If
        strTemp = LfPBF & strTemp & RgPbf
        lngHeigth = txtAuditingMan.Height

        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Աѡ��", False, txtAuditingMan.Tag, "", False, False, True, vRect.Left - 15, vRect.Top, lngHeigth, blnCancel, False, False, strTemp)
               
        If rsTemp Is Nothing Then
            If Not blnCancel Then MsgBox "û����������������,����[��Ա��Ϣ]!", vbInformation, gstrSysName
            If txtAuditingMan.Enabled Then
                txtAuditingMan.SetFocus
                txtAuditingMan.SelStart = 0
                txtAuditingMan.Text = ""
                Exit Sub
            End If
        End If
       
        With rsTemp
            If UCase(TypeName(txtAuditingMan)) = "TEXTBOX" Then
                txtAuditingMan = IIf(IsNull(!����), "", !����)
                zlCommFun.PressKey (vbKeyTab)
            Else
                txtAuditingMan.SetFocus
                txtAuditingMan.SelStart = 0
                txtAuditingMan.SelLength = Len(txtAuditingMan.Text)
                zlCommFun.PressKey vbKeyTab
            End If
            .Close
        End With
    End If
End Sub

Private Sub txtAuditingMan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtAuditingMan_LostFocus()
    If zlCommFun.ActualLen(Trim(txtAuditingMan.Text)) > 20 Then
        MsgBox "�����˳���,���������10�����ֻ�20���ַ�!", vbInformation + vbOKOnly, gstrSysName
        txtAuditingMan.SetFocus
        txtAuditingMan.SelStart = 0
        txtAuditingMan.SelLength = Len(txtAuditingMan.Text)
        If txtAuditingMan.Enabled Then txtAuditingMan.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtBeginNo_GotFocus()
    With txtBeginNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    zlCommFun.OpenIme False
End Sub

Private Sub txtBeginNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtBeginNo.Text = Replace(UCase(txtBeginNo.Text), "'", "")
'        If Len(txtBeginNo) < 20 And Len(txtBeginNo) > 0 Then
'            strNo = txtBeginNo.Text
'            Call MakeNO(117, lng�ⷿID, strNo)
'            txtBeginNo.Text = strNo
'        End If
        zlCommFun.PressKey (vbKeyTab)
'        txtEndNo.SetFocus
    End If
End Sub

Private Sub txtBeginNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> 8 Then
            If KeyAscii <> 46 Then
                KeyAscii = 0
            ElseIf InStr(txtBeginNo.Text, ".") = 0 Then
                KeyAscii = 0
'            Else
'                KeyAscii = 0
            End If
        End If
    End If
    
    If KeyAscii = 39 Then KeyAscii = 0

End Sub

Private Sub txtBeginNo_LostFocus()
    If zlCommFun.ActualLen(Trim(txtBeginNo.Text)) > 18 Then
        MsgBox "��ʼסԺ�ų���,���������9�����ֻ�18���ַ�!", vbInformation + vbOKOnly, gstrSysName
        txtBeginNo.SetFocus
        txtBeginNo.SelStart = 0
        txtBeginNo.SelLength = Len(txtBeginNo.Text)
        If txtBeginNo.Enabled Then txtBeginNo.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtEndNo_GotFocus()
    With txtEndNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    zlCommFun.OpenIme False
End Sub

Private Sub txtEndNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtEndNo.Text = Replace(UCase(txtEndNo.Text), "'", "")
'        If Len(txtEndNo) < 20 And Len(txtEndNo) > 0 Then
'            strNo = txtEndNo.Text
'            Call MakeNO(117, lng�ⷿID, strNo)
'            txtEndNo.Text = strNo
'        End If
        zlCommFun.PressKey (vbKeyTab)
'        txtEndNo.SetFocus
    End If
End Sub

Private Sub txtEndNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> 8 Then
            If KeyAscii <> 46 Then
                KeyAscii = 0
            ElseIf InStr(txtBeginNo.Text, ".") = 0 Then
                KeyAscii = 0
'            Else
'                KeyAscii = 0
            End If
        End If
    End If
    
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtEndNo_LostFocus()
    If zlCommFun.ActualLen(Trim(txtEndNo.Text)) > 18 Then
        MsgBox "����סԺ�ų���,���������9�����ֻ�18���ַ�!", vbInformation + vbOKOnly, gstrSysName
        txtEndNo.SetFocus
        txtEndNo.SelStart = 0
        txtEndNo.SelLength = Len(txtEndNo.Text)
        If txtEndNo.Enabled Then txtEndNo.SetFocus
        Exit Sub
    End If
End Sub

Private Sub vfgList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call SaveHead(vfgList, 1)
End Sub

'Private Sub vfgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'''    If OldRow > 0 Then
''        Call zl_VsGridRowChange(vfgList, OldRow, NewRow, OldCol, NewCol)
'''    End If
'End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strStatus As String
    Dim strTemp As String
    If mlngModule = 201 Then '���Ӳ���
        strTemp = ";����;"
    Else
        strTemp = "����;"
    End If
    
    If vfgList.Editable = flexEDNone Then
        Cancel = True
        Exit Sub
    End If
    
    strStatus = Trim(vfgList.TextMatrix(Row, vfgList.ColIndex("״̬")))
                
    Select Case Col
        Case vfgList.ColIndex("ѡ��")
            If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "δ����" Then  'And mlngApplyId <> 0
                Cancel = False
            Else
                Cancel = True
            End If
            
            Exit Sub
        Case Else
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vfgList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Select Case Col
        Case vfgList.ColIndex("���"), vfgList.ColIndex("ѡ��")
            Position = -1
            Exit Sub
    End Select
    If Position = 0 Or Position = 1 Then
        Position = Col
    End If
End Sub

Private Sub vfgList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call SaveHead(vfgList, 1)
End Sub

Private Sub vfgList_Click()
    Dim lngApplyId As String
    Dim strStatus As String
    Dim strStuffMan As String
    Dim lngRows As Long
    Dim lngFlag As Long
    Dim i As Long
    Dim strTemp As String
    
    If mlngModule = 201 Then '���Ӳ���
        strTemp = ";����;"
    Else
        strTemp = "����;"
    End If
    
    lngFlag = 0
    With vfgList
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                strStatus = Trim(.TextMatrix(i, .ColIndex("״̬")))
                If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "δ����" And .TextMatrix(i, .ColIndex("ѡ��")) = -1 Then
                    lngFlag = 1
                    i = lngRows
                End If
            Next
        End If
        
        If lngFlag = 0 Then
            mlngTempId = 0
        End If
        If .Row > 0 Then
            lngApplyId = Val(.TextMatrix(.Row, .ColIndex("��Ժ����id")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
            If InStr(mstrPrivs, strTemp) <> 0 And strStatus = "δ����" And .TextMatrix(.Row, .ColIndex("ѡ��")) = -1 Then
                If mlngTempId = 0 Then mlngTempId = lngApplyId
                If lngApplyId <> mlngTempId And cboOutDept.ListIndex <> 0 Then
                    MsgBox "�ò�����Ժ�������Ѿ�ѡ��ĳ�Ժ���Ҳ�һ�£�ϵͳ�Զ�ȡ��ѡ��!", vbInformation, gstrSysName
                    .TextMatrix(.Row, .ColIndex("ѡ��")) = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vfgList_DblClick()
    '˫���༭
    Call SetModify_DblClick
End Sub

Private Sub vfgList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    If KeyCode = vbKeyReturn Then
        Call zlVsMoveGridCell(vfgList, vfgList.ColIndex("���"), vfgList.ColIndex("��ͥ��ַ"), False, lngRow)
    End If
End Sub

Private Sub vfgList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim intGetHeight As Integer
    Dim intGetWidth As Integer
    
    intGetWidth = vfgList.ColWidth(0)
    intGetHeight = vfgList.RowHeight(0)
    If (Button = 2) Then
        If x < intGetWidth And y < intGetHeight Then
            Call RightHead(vfgList, 1)
        Else
            objExtendedBar.ShowPopup
        End If
    End If
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    Dim intCol As Long
    With vfgList
        '���ѡ���е���ɫ
''        For intCol = 0 To .Cols - 1
''            .Col = intCol
''            .CellBackColor = glngGetFocus_Font
'''            .CellForeColor = glngLostFocus_Font
''        Next
        .GridLines = flexGridInset
    End With
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = mstrListTitle
        
    Set objRow = New zlTabAppRow

    If cboOutDept.Visible Then
        objRow.Add "��Ժ����:" & cboOutDept.Text
    End If
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vfgList
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    With vfgList
        .GridLines = flexGridNone
    End With
End Sub

Private Sub SetAdd()
    Dim strNo As String
    Dim intEditState As Integer
    Dim strPatientSum As String
    Dim lngApplyId As Long
    Dim blnReturn As Boolean
    
    intEditState = 1
    strPatientSum = GetChoiceData()
'    If mlngTempId = 0 Then
'        lngApplyId = mlngApplyId
'    Else
'        lngApplyId = mlngTempId
'    End If
    lngApplyId = cboOutDept.ItemData(cboOutDept.ListIndex)
    
    frm�������ձ༭.ShowCard Me, intEditState, strPatientSum, lngApplyId, blnReturn, mlngModule
    If blnReturn Then
        mintDelete = 1
        mintEditState = intEditState
        Call SetRefresh
        mintDelete = 0
    End If
End Sub

Private Sub SetModify()
    Dim intEditState As Integer
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim blnReturn As Boolean
    Dim strPatientSum As String
    Dim lngApplyId As Long
    Dim lngRow As Long
    
    intEditState = 2
    
    With vfgList
        If .Row > 0 Then
            lngRow = .Row
            lngPatientlId = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            lngMtyId = Val(.TextMatrix(.Row, .ColIndex("��ҳID")))
            lngApplyId = Val(.TextMatrix(.Row, .ColIndex("��Ժ����id")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
            
            If strStatus = "�ѽ���" Then
                strPatientSum = lngPatientlId & "_" & lngMtyId
                frm�������ձ༭.ShowCard Me, intEditState, strPatientSum, lngApplyId, blnReturn
            End If
        End If
    End With
    If blnReturn Then
        mintDelete = 1
        mintEditState = intEditState
        Call SetRefresh
        mintDelete = 0
        If lngRow > 0 And lngRow < vfgList.Rows Then vfgList.Select lngRow, 2
    End If
End Sub

Private Sub SetDelete()
    Dim strSQL As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim strPname As String
    Dim lngRow As Long
    
    strSQL = ""
    
    With vfgList
        If .Row > 0 Then
            lngRow = .Row
            lngPatientlId = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            lngMtyId = Val(.TextMatrix(.Row, .ColIndex("��ҳID")))
            strPname = Trim(.TextMatrix(.Row, .ColIndex("����")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
            
            If strStatus = "�ѽ���" Then
                strSQL = "Zl_�������ռ�¼_Delete(" & lngPatientlId & "," & lngMtyId & ")"
            End If
        End If
    End With
    If Trim(strSQL) <> "" Then
        If MsgBox("��ȷ��ɾ������Ϊ��" & strPname & "���Ĳ������յǼ���ɾ�����ָܻ���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
            Err = 0: On Error GoTo errHand:
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            mintDelete = 1
            Call SetRefresh
            If lngRow + 1 > vfgList.Rows Then
                lngRow = lngRow - 1
            End If
            If lngRow > 0 And lngRow < vfgList.Rows Then vfgList.Select lngRow, 2
            mintDelete = 0
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDisplay()
    Dim intEditState As Integer
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim blnReturn As Boolean
    Dim strPatientSum As String
    Dim lngApplyId As Long
    
    intEditState = 3
    
    With vfgList
        If .Row > 0 Then
            lngPatientlId = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            lngMtyId = Val(.TextMatrix(.Row, .ColIndex("��ҳID")))
            lngApplyId = Val(.TextMatrix(.Row, .ColIndex("��Ժ����id")))
            strStatus = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
            
            If strStatus <> "δ����" Then
                strPatientSum = lngPatientlId & "_" & lngMtyId
                frm�������ձ༭.ShowCard Me, intEditState, strPatientSum, lngApplyId
            End If
        End If
    End With
    mintEditState = intEditState
End Sub

Private Sub SetModify_DblClick()
    If mintDblClick = 1 Then
        Call SetModify
'    Else
'        Call SetDisplay
    End If
End Sub

Private Sub SetVerify_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strStatus As String
   
    Control.Visible = True
    Control.Enabled = False
  
    If vfgList.Row > 0 Then
        strStatus = Trim(vfgList.TextMatrix(vfgList.Row, vfgList.ColIndex("״̬")))
        If strStatus = "�ѽ���" Then
            Control.Enabled = True
        End If
    End If
End Sub

Private Function GetChoiceData() As String
    Dim lngApplyId As String
    Dim lngPatientlId As Long
    Dim lngMtyId As Long
    Dim strStatus As String
    Dim lngRows As Long
    Dim strTemp As String
    Dim i As Long
    Dim j As Long
    Dim intCount As Integer
    Dim strTempEle As String
    If mlngModule = 201 Then '���Ӳ���
        strTempEle = ";����;"
    Else
        strTempEle = "����;"
    End If
    
    intCount = 0
    strTemp = ""
    GetChoiceData = ""
    With vfgList
        If .Rows > 1 Then
            lngRows = .Rows - 1
            For i = 1 To lngRows
                lngPatientlId = Val(.TextMatrix(i, .ColIndex("����ID")))
                lngMtyId = Val(.TextMatrix(i, .ColIndex("��ҳID")))
                strStatus = Trim(.TextMatrix(i, .ColIndex("״̬")))
                If InStr(mstrPrivs, strTempEle) <> 0 And strStatus = "δ����" And .TextMatrix(i, .ColIndex("ѡ��")) = -1 Then
                    intCount = intCount + 1
                    If intCount > 100 Then
                        GetChoiceData = strTemp
                        MsgBox "����ѡ�Ĳ�����̫���ˣ�ֻ����ǰ��ѡ�е�100�ݡ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If strTemp = "" Then
                        strTemp = lngPatientlId & "_" & lngMtyId
                    Else
                        strTemp = strTemp & "," & lngPatientlId & "_" & lngMtyId
                    End If
                End If
            Next
        End If
    End With
    GetChoiceData = strTemp
End Function

'==============================================================================
'=���ܣ� �鿴��ҳ
'==============================================================================
Private Sub RecordLook()
    
    On Error GoTo ErrH
    With vfgList
        If .Row < 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("����id"))) = 0 Then GoTo ErrH
        Call frmArchiveView.ShowArchive(Me, Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), False)
        
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetMedicalExits() As ADODB.Recordset
    '******************************************************************************************************************
    '����:����Ƿ�װ�˲���ϵͳ
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim strSQL As String
    strSQL = "Select ��� From zlsystems where ���=300"
    
    Set GetMedicalExits = zlDatabase.OpenSQLRecord(strSQL, "����ϵͳ")
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

