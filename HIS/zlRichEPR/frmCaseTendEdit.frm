VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendEdit 
   Caption         =   "�����¼"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCaseTendEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2850
      Index           =   2
      Left            =   240
      ScaleHeight     =   2850
      ScaleWidth      =   10920
      TabIndex        =   32
      Top             =   4215
      Width           =   10920
      Begin VB.Frame fraTime 
         Height          =   525
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   11235
         Begin VB.ComboBox cbo 
            Height          =   300
            Left            =   8235
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   150
            Width           =   1680
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Left            =   1065
            TabIndex        =   37
            Top             =   150
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   177471491
            UpDown          =   -1  'True
            CurrentDate     =   38702
         End
         Begin MSComctlLib.TabStrip tbs 
            Height          =   300
            Left            =   3765
            TabIndex        =   40
            Top             =   165
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   529
            MultiRow        =   -1  'True
            Style           =   2
            TabMinWidth     =   529
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   2
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "2"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��¼��(&G)"
            Height          =   180
            Index           =   0
            Left            =   2925
            TabIndex        =   41
            Top             =   210
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��(&T)"
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   38
            Top             =   195
            Width           =   990
         End
      End
      Begin zlRichEPR.VsfGrid vsf 
         Height          =   1575
         Left            =   150
         TabIndex        =   34
         Top             =   630
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2778
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   1
      Left            =   3210
      ScaleHeight     =   1695
      ScaleWidth      =   8445
      TabIndex        =   26
      Top             =   6105
      Width           =   8445
      Begin VB.Frame fra 
         Height          =   525
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   11235
         Begin MSComCtl2.UpDown udnDay 
            Height          =   270
            Left            =   585
            TabIndex        =   31
            Top             =   150
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   476
            _Version        =   393216
            BuddyControl    =   "txtDay"
            BuddyDispid     =   196618
            OrigLeft        =   810
            OrigTop         =   180
            OrigRight       =   1065
            OrigBottom      =   405
            Max             =   30
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtDay 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   270
            Left            =   270
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "1"
            Top             =   165
            Width           =   315
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ǰ       ����ʷ��¼���"
            Height          =   180
            Index           =   12
            Left            =   60
            TabIndex        =   29
            Top             =   195
            Width           =   2070
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfHistory 
         Height          =   1020
         Left            =   240
         TabIndex        =   27
         Top             =   660
         Width           =   1995
         _cx             =   3519
         _cy             =   1799
         Appearance      =   2
         BorderStyle     =   0
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picCustom 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1740
      ScaleHeight     =   300
      ScaleWidth      =   1995
      TabIndex        =   2
      Top             =   75
      Width           =   1995
      Begin VB.CommandButton cmd 
         Height          =   300
         Left            =   1665
         Picture         =   "frmCaseTendEdit.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   0
         Width           =   330
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1665
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3435
      Index           =   0
      Left            =   810
      ScaleHeight     =   3435
      ScaleWidth      =   9885
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   645
      Width           =   9885
      Begin VB.Frame fraInfo 
         Height          =   705
         Left            =   0
         TabIndex        =   4
         Top             =   -90
         Width           =   9780
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   8370
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   255
            Width           =   1350
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   9
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   435
            Width           =   1200
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   8
            Left            =   5910
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   435
            Width           =   2355
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   7
            Left            =   555
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   435
            Width           =   1185
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   6
            Left            =   7050
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   5
            Left            =   3990
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   435
            Width           =   1455
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   4
            Left            =   5910
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   180
            Width           =   600
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   3990
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   180
            Width           =   1455
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   2985
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   375
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   180
            Width           =   360
         End
         Begin VB.TextBox txtShow 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   0
            Left            =   555
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   180
            Width           =   1185
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��  Ժ:"
            Height          =   180
            Index           =   11
            Left            =   3375
            TabIndex        =   23
            Top             =   435
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����:"
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   13
            Top             =   435
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���:"
            Height          =   180
            Index           =   9
            Left            =   5475
            TabIndex        =   12
            Top             =   420
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����:"
            Height          =   180
            Index           =   8
            Left            =   5475
            TabIndex        =   11
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ҽʦ:"
            Height          =   180
            Index           =   7
            Left            =   6615
            TabIndex        =   10
            Top             =   180
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����:"
            Height          =   180
            Index           =   6
            Left            =   1755
            TabIndex        =   9
            Top             =   435
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�Ա�:"
            Height          =   180
            Index           =   5
            Left            =   1740
            TabIndex        =   8
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "סԺ��:"
            Height          =   180
            Index           =   4
            Left            =   3375
            TabIndex        =   7
            Top             =   165
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����:"
            Height          =   180
            Index           =   3
            Left            =   2550
            TabIndex        =   6
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����:"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   165
            Width           =   450
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   2025
         Left            =   270
         TabIndex        =   33
         Top             =   975
         Width           =   2700
         _Version        =   589884
         _ExtentX        =   4762
         _ExtentY        =   3572
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendEdit.frx":10C8
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14843
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmCaseTendEdit.frx":195A
            Text            =   "��Χ��"
            TextSave        =   "��Χ��"
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
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCaseTendEdit.frx":81BC
      Left            =   345
      Top             =   600
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCaseTendEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���ڼ���������
'######################################################################################################################

Private mblnStartUp As Boolean
Private mblnOk As Boolean
Private mstrSQL As String
Private mbytMode As Byte                    '1-������¼;2-�޸ļ�¼;3-����ǩ��;4-ȡ��ǩ��:5-��ʷ�汾
Private mlngKey As Long
Private mstrTime As String
Private mrsParam As ADODB.Recordset
Private mblnChanged As Boolean
Private mblnNoChanged As Boolean
Private mstrSvrDate As String
Private mint������ As Integer
Private mblnReading As Boolean
Private mstrSvr���� As String
Private mstr���￨��ĸǰ׺ As String
Private mint���￨���볤�� As Integer
Private mstrPrivs As String
Private mrsPatient As ADODB.Recordset
Private mlngRowNum As Long
Private mstrFindKey As String
Private mobjFindKey As CommandBarControl
Private mint����Ӧ�� As Integer
Private mblnDefault As Boolean
Private mclsVsfHistory As clsVsf
Private mintPreDays As Integer

Private Enum mCol
    ��¼�� = 1
    ������Ŀ
    ��Ŀ��λ
    ��Ŀ����
    ��Ŀ����
    ��ĿС��
    ��Ŀ��ʾ
    ��Ŀֵ��
    ��Ŀȱʡ
    ��Ŀ����
    ��Ŀid
    �Ƿ�䶯
    ��¼���
    ���
    ��λ
    δ��˵��
End Enum

Public Event AfterDataChanged()

'�Զ������/��������
'######################################################################################################################

Private Property Let DataChanged(vData As Boolean)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mblnChanged = vData
        
End Property

Private Property Get DataChanged() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    DataChanged = mblnChanged
    cmd.Enabled = Not mblnChanged And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5)
    tbs.Enabled = Not mblnChanged
    cboBaby.Enabled = Not mblnChanged
        
    For intLoop = 0 To tbcPage.ItemCount - 1
        If Not tbcPage.Item(intLoop) Is Nothing Then
            tbcPage.Item(intLoop).Enabled = Not mblnChanged
        End If
    Next

End Property

Public Function ShowEdit(ByVal frmParent As Form, ByVal strParam As String, Optional ByVal bytMode As Byte = 1, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������strParam->����id;��ҳid;����id;Ӥ��;��Դ;ʱ��,ID
    '���أ�
    '******************************************************************************************************************
    Dim varParam As Variant
        
    mblnStartUp = True
    
    mbytMode = bytMode
    mblnOk = False
    mstrPrivs = strPrivs
        
    '------------------------------------------------------------------------------------------------------------------
    Set mrsParam = New ADODB.Recordset
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "��ҳid", adBigInt)
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "����id", adBigInt)
    Call CreateParam(mrsParam, "Ӥ��", adTinyInt)
    Call CreateParam(mrsParam, "�汾", adTinyInt)
    Call CreateParam(mrsParam, "��Դ", adTinyInt)
    Call CreateParam(mrsParam, "��Ժ", adTinyInt)
    Call CreateParam(mrsParam, "ʱ��", adVarChar, 20)
    Call CreateParam(mrsParam, "����ȼ�", adTinyInt)
    Call CreateParam(mrsParam, "��Ժ��ʼ����", adVarChar, 30)
    Call CreateParam(mrsParam, "��Ժ��������", adVarChar, 30)
    Call CreateParam(mrsParam, "��Ժ����", adTinyInt)
    Call CreateParam(mrsParam, "��Ժ����", adTinyInt)
    Call CreateParam(mrsParam, "ת������", adTinyInt)
    Call CreateParam(mrsParam, "����Ʋ���", adTinyInt)
    Call CreateParam(mrsParam, "ת������", adTinyInt)
    Call CreateParam(mrsParam, "��¼id", adBigInt)
    
    '------------------------------------------------------------------------------------------------------------------
    If strParam <> "" Then varParam = Split(strParam, ";")
    mrsParam.Open
    mrsParam.AddNew
                    
    mrsParam("����id").Value = Val(varParam(0))
    mrsParam("��ҳid").Value = Val(varParam(1))
    mrsParam("����id").Value = Val(varParam(2))
    mrsParam("����id").Value = Val(varParam(2))
    mrsParam("����ȼ�").Value = 3
    mrsParam("Ӥ��").Value = 0
    mrsParam("�汾").Value = 0
    If UBound(varParam) >= 3 Then mrsParam("Ӥ��").Value = Val(varParam(3))
    If UBound(varParam) >= 4 Then mrsParam("��Դ").Value = Val(varParam(4))
    If UBound(varParam) >= 5 Then mrsParam("ʱ��").Value = CStr(varParam(5))
    If UBound(varParam) >= 5 Then mrsParam("ʱ��").Value = CStr(varParam(5))
    
    '��ʼ�ؼ�
    '------------------------------------------------------------------------------------------------------------------
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    
    '------------------------------------------------------------------------------------------------------------------
    Call ExecuteCommand("ˢ�»�����Ϣ")
    
    If mbytMode <> 1 Then
        Call ExecuteCommand("��ȡ��¼")
    End If
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("��ȡ����")
    
    If mbytMode <> 1 Then
        '��ȡָ����¼id��ָ����Ļ�������
        Call ExecuteCommand("��ȡ�������")
    Else
        '����,�����ȱʡֵ����������
        DataChanged = mblnDefault
    End If
    
    Vsf.Col = mCol.��¼���
    
    DataChanged = False
    mblnStartUp = False
    
    Me.Show , frmParent
    
    ShowEdit = mblnOk
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
EndHand:
    mblnStartUp = False
    Unload Me
End Function

Private Function ReadPatient() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim strParam As String
    On Error GoTo ErrHand
    
    '��Ժ�ͳ�Ժ����:��Ժ���˿������ж��סԺ
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("��Ժ����").Value) <> 0 Or Val(mrsParam("��Ժ����").Value) <> 0 Or Val(mrsParam("����Ʋ���").Value) <> 0 Then
        gstrSQL = _
            "Select Decode(B.��Ժ����,NULL,Decode(B.״̬,3,2,1),Decode(B.��Ժ��ʽ,'����',4,3)) as ����," & _
            " Decode(B.��Ժ����,NULL,Decode(B.״̬,3,'Ԥ��Ժ����','��Ժ����'),Decode(B.��Ժ��ʽ,'����','��������','��Ժ����')) as ����," & _
            " A.����ID,B.��ҳID,B.סԺ��,A.�����,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,C.���� as ����,B.סԺҽʦ," & _
            " B.��Ժ���� as ����,B.�ѱ�,B.��Ժ����,B.��Ժ����,B.״̬,B.����,A.���￨��" & _
            " From ������Ϣ A,������ҳ B,���ű� C" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And ([6]=1 Or Nvl(B.״̬,0)<>1) And B.��Ժ����ID=C.ID" & _
            " And B.��ǰ����ID=[1] And ([4]<>0 And B.��Ժ���� is NULL Or [5]<>0 And B.��Ժ���� Between [2] And [3]) "
    End If
    
    'ת������:��Ժ,ҽ���ʹ�����ʾ����ת��ǰ��
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("ת������").Value) <> 0 Then
        gstrSQL = gstrSQL & IIf(gstrSQL <> "", " Union All ", "") & _
            "Select Distinct 5 as ����,'ת������' as ����," & _
            " A.����ID,B.��ҳID,B.סԺ��,A.�����,NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,D.���� as ����,C.����ҽʦ as סԺҽʦ," & _
            " C.����,B.�ѱ�,B.��Ժ����,B.��Ժ����,B.״̬,B.����,A.���￨��" & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D" & _
            " Where A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And C.����ID=D.ID" & _
            " And Nvl(B.״̬,0)=0 And B.��Ժ���� is NULL And B.��ǰ����ID<>[1]" & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID And C.����ID=[1]" & _
            " And C.��ֹԭ��=3 And C.��ֹʱ�� Between Sysdate-[7] And Sysdate "
    End If
    gstrSQL = gstrSQL & " Order by ����,����,��ҳID Desc"
    gstrSQL = "Select RowNum As ID,1 As ĩ��,A.* From (" & gstrSQL & ") A"
    
    If mbytMode <> 5 Then
    
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("����id").Value), _
                                                                CDate(Format(mrsParam("��Ժ��ʼ����").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("��Ժ��������").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("��Ժ����").Value), _
                                                                0, _
                                                                Val(mrsParam("����Ʋ���").Value), _
                                                                Val(mrsParam("ת������").Value))
                                                            
    Else
        
        Set mrsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
                                                                Val(mrsParam("����id").Value), _
                                                                CDate(Format(mrsParam("��Ժ��ʼ����").Value, "yyyy-MM-dd 00:00:00")), _
                                                                CDate(Format(mrsParam("��Ժ��������").Value, "yyyy-MM-dd 23:59:59")), _
                                                                Val(mrsParam("��Ժ����").Value), _
                                                                Val(mrsParam("��Ժ����").Value), _
                                                                Val(mrsParam("����Ʋ���").Value), _
                                                                Val(mrsParam("ת������").Value))
    End If
                                                            
    ReadPatient = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim RS As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo ErrHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '�����
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�������"): cbrControl.ToolTipText = "�������(Alt+C)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(Alt+P)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(Alt+N)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "���"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�����Ŀ(Alt+A)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��"):  cbrControl.ToolTipText = "ɾ����Ŀ(Alt+D)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(Alt+S)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "��¼ǩ��"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "��¼ǩ��(Alt+R)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��ǩ��"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "ȡ��ǩ��(Alt+U)"
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")

        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): cbrControl.ToolTipText = "�˳�(Esc)"

    End With
    
    '��λ������
    '------------------------------------------------------------------------------------------------------------------
    
    Set objExtendedBar = cbsThis.Add("��λ", xtpBarTop)
    
    objExtendedBar.ContextMenuPresent = False
    objExtendedBar.ShowTextBelowIcons = False
    objExtendedBar.EnableDocking xtpFlagHideWrap
    
    With objExtendedBar.Controls

        mstrFindKey = Trim(zlDatabase.GetPara("���ҷ���", glngSys, 1255, "��  ��"))
        If mstrFindKey = "" Then mstrFindKey = "��  ��"
        
        Set mobjFindKey = .Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
        mobjFindKey.IconId = conMenu_View_Find
        mobjFindKey.BeginGroup = True
        mobjFindKey.ToolTipText = "��ݼ�:F4"
        Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.��  ��"): cbrControl.Parameter = "��  ��"
        Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.סԺ��"): cbrControl.Parameter = "סԺ��"
        Set cbrControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&3.���￨"): cbrControl.Parameter = "���￨"
        
        Set cbrCustom = .Add(xtpControlCustom, conMenu_View_Location, "")
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Handle = picCustom.hWnd
        txt.ToolTipText = "���Ҳ���(F3)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "ǰһ����"): cbrControl.flags = xtpFlagRightAlign: cbrControl.ToolTipText = "ǰһ����(Ctrl+Left)"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "��һ����"): cbrControl.flags = xtpFlagRightAlign: cbrControl.ToolTipText = "��һ����(Ctrl+Right)"
    End With
    
    Call SetDockRight(objExtendedBar, cbrToolBar)
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FALT, Asc("C"), conMenu_Edit_Audit
        .Add FALT, Asc("N"), conMenu_Edit_NewItem
        .Add FALT, Asc("A"), conMenu_Edit_Append
        .Add FALT, Asc("D"), conMenu_Edit_Delete
        .Add FALT, Asc("S"), conMenu_Edit_Transf_Save
        .Add FALT, Asc("R"), conMenu_Tool_Sign
        .Add FALT, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
        
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
    End With
    
    InitMenuBar = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Function InitData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    mint������ = 0
    Select Case mbytMode
    Case 1
        Me.Caption = Me.Caption & " - �Ǽ�"
        cbo.Visible = False
    Case 2
        Me.Caption = Me.Caption & " - �޸�"
        cbo.Visible = False
    Case 3
        Me.Caption = Me.Caption & " - ǩ��"
        cbo.Visible = False
    Case 4
        Me.Caption = Me.Caption & " - ȡ��ǩ��"
        cbo.Visible = False
    Case 5
        Me.Caption = Me.Caption & " - ��ʷ�汾"
        cbo.Visible = True
    End Select
    
    cmd.Enabled = (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5)
    txt.Enabled = cmd.Enabled
    dtp.Enabled = (mbytMode = 1 Or mbytMode = 2)
    
    '------------------------------------------------------------------------------------------------------------------
    With Vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "��¼��", 1500, 1
        .NewColumn "��¼��Ŀ", 1590, 1
        .NewColumn "��λ", 750, 1
        .NewColumn "��Ŀ����", 0, 1
        .NewColumn "��Ŀ����", 0, 1
        .NewColumn "��ĿС��", 0, 1
        .NewColumn "��Ŀ��ʾ", 0, 1
        .NewColumn "��Ŀֵ��", 0, 1
        .NewColumn "��Ŀȱʡ", 0, 1
        .NewColumn "��Ŀ����", 0, 1
        .NewColumn "��Ŀid", 0, 1
        .NewColumn "�Ƿ�䶯", 0, 1
        
        .NewColumn "��¼����", 3750, 1, , 1
        .NewColumn "���", 900, 1
        .NewColumn "��λ", 900, 1
        .NewColumn "δ��˵��", 900, 1, "...", 1
        
        .FixedCols = 4
                
        .Body.ColHidden(mCol.��Ŀ����) = True
        .Body.ColHidden(mCol.��Ŀ����) = True
        .Body.ColHidden(mCol.��ĿС��) = True
        .Body.ColHidden(mCol.��Ŀ��ʾ) = True
        .Body.ColHidden(mCol.��Ŀֵ��) = True
        .Body.ColHidden(mCol.��Ŀȱʡ) = True
        .Body.ColHidden(mCol.��Ŀ����) = True
        .Body.ColHidden(mCol.��Ŀid) = True
        .Body.MergeCells = flexMergeFree
        .Body.MergeCol(mCol.��¼��) = True
        .Body.WordWrap = True
        
        If mbytMode > 2 Then
            .Body.Editable = flexEDNone
'            cmdCalc.Enabled = False
        End If
    End With
    
    Set mclsVsfHistory = New clsVsf
    With mclsVsfHistory
        Call .Initialize(Me.Controls, vsfHistory, True, False)
        Call .ClearColumn
        Call .AppendColumn("��¼ʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
        Call .AppendColumn("��ʷ���", 15, flexAlignLeftCenter, flexDTString, "", , True)
        vsfHistory.FixedCols = 1
        vsfHistory.ExplorerBar = flexExNone
        vsfHistory.RowHidden(0) = True
        .AppendRows = False
    End With
        
    Dim objPane As Pane
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    
    dkpMain.SetCommandBars cbsThis
    
    Set objPane = dkpMain.CreatePane(1, 100, 200, DockTopOf, Nothing): objPane.Title = "�༭": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane): objPane.Title = "��ʷ": objPane.Options = PaneNoCaption
        
    Call InitTabControl
    
    If mbytMode <> 1 And mbytMode <> 2 Then
        dkpMain.Panes(2).Close
        picPane(1).Visible = False
    End If
    
    InitData = True
    
End Function

Private Function InitTabControl() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼTab�ؼ�
    '������
    '���أ�
    '******************************************************************************************************************
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
'            .COLOR = xtpTabColorDefault
            .ColorSet.ButtonSelected = &HFFFF&
            .DisableLunaColors = False
        End With

        Set .Icons = zlCommFun.GetPubIcons

        .InsertItem 0, "������1  ", picPane(2).hWnd, 0

        .Item(0).Selected = True
        
    End With
    
    InitTabControl = True
    
End Function

Private Function OpenPatientMap(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ������Ϣ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error GoTo ErrHand
    
    mblnDefault = False
    
    mrsParam("����id").Value = lng����ID
    mrsParam("��ҳid").Value = lng��ҳID
    mrsParam("Ӥ��").Value = intӤ��
        
    '������Ϣ
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ����,B.סԺ��,B.��Ժ����,B.ҽ�Ƹ��ʽ," & _
        " D.�������,B.����,B.��ǰ����,C.���� as ����ȼ�,B.��Ժ����," & _
        " B.��Ժ����,B.״̬,B.����ת��,B.��Ժ����ID,B.��ǰ����ID,A.סԺ����,B.סԺҽʦ " & _
        " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,������ϼ�¼ D" & _
        " Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2] And B.����ȼ�ID=C.ID(+)" & _
        " And D.����id(+)=B.����id And D.��ҳid(+)=B.��ҳid And D.�������(+)=1 "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value))
    
    If RS.BOF = False Then
        txt.Text = zlCommFun.NVL(RS("����").Value)
        txt.Tag = ""
        
        txtShow(0).Text = zlCommFun.NVL(RS("����").Value)
        txtShow(1).Text = zlCommFun.NVL(RS("�Ա�").Value)
        txtShow(2).Text = zlCommFun.NVL(RS("����").Value)
        txtShow(3).Text = zlCommFun.NVL(RS("סԺ��").Value)
        txtShow(4).Text = zlCommFun.NVL(RS("��Ժ����").Value)
        txtShow(5).Text = Format(zlCommFun.NVL(RS("��Ժ����").Value), "yyyy-MM-dd HH:mm")
        txtShow(6).Text = zlCommFun.NVL(RS("סԺҽʦ").Value)
        txtShow(7).Text = zlCommFun.NVL(RS("����ȼ�").Value)
        txtShow(8).Text = zlCommFun.NVL(RS("�������").Value)
        txtShow(9).Text = zlCommFun.NVL(RS("��ǰ����").Value)

    End If
    mstrSvr���� = txt.Text
    
    '
    '------------------------------------------------------------------------------------------------------------------
    cboBaby.Clear
    cboBaby.AddItem "���˱���"
    gstrSQL = "Select a.���,Decode(a.Ӥ������,Null,NVL(c.����,b.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������" & _
        " From ������Ϣ b,������ҳ c,������������¼ a Where b.����id=c.����id And a.����id=c.����ID And a.��ҳID=c.��ҳID And c.����id=[1] And c.��ҳid=[2]  Order By a.���"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value))
    If RS.BOF = False Then
        Do While Not RS.EOF
            cboBaby.AddItem RS("Ӥ������").Value
            RS.MoveNext
        Loop
    End If
    On Error Resume Next
    cboBaby.ListIndex = Val(mrsParam("Ӥ��").Value)
    On Error GoTo ErrHand
    If cboBaby.ListIndex = -1 Then cboBaby.ListIndex = 0
    cboBaby.Visible = (cboBaby.ListCount > 1)
    
    
    '��ȡ��Ժʱ��
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Select Min(��ʼʱ��) As ��ʼʱ�� From ���˱䶯��¼ Where ����id=[1] and ��ҳid=[2]"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If RS.BOF = False Then
        If IsNull(RS("��ʼʱ��").Value) = False Then
            On Error Resume Next
            dtp.MinDate = Format(DateAdd("n", 1, CDate(Format(RS("��ʼʱ��").Value, "yyyy-MM-dd HH:mm") & ":00")), dtp.CustomFormat)
            On Error GoTo ErrHand
        End If
    End If
    
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))
    
    mstrSQL = "Select ��Ժ����,��Ժ���� From ������ҳ Where ����id=[1] and ��ҳid=[2]"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If RS.BOF = False Then

        On Error Resume Next
        
        If IsNull(RS("��Ժ����").Value) Then
            
            If mintPreDays > 0 Then
                dtp.MaxDate = Format(Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd") & " 23:59:59", dtp.CustomFormat)
            Else
                dtp.MaxDate = Format(zlDatabase.Currentdate, dtp.CustomFormat)
            End If
            
        Else
            dtp.MaxDate = Format(zlCommFun.NVL(RS("��Ժ����").Value), dtp.CustomFormat)
        End If
        
        dtp.MaxDate = Format(zlCommFun.NVL(RS("��Ժ����").Value, Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd HH:mm:ss")), dtp.CustomFormat)
        On Error GoTo ErrHand

    End If
        
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If RS.BOF = False Then
        mrsParam("����ȼ�").Value = zlCommFun.NVL(RS("����ȼ�").Value)
    End If
    
'    If mrsParam("ʱ��").Value = "" Then
        
        Vsf.Rows = 2
        Vsf.RowData(1) = 0
        Vsf.Cell(flexcpData, 1, 0, 1, Vsf.Cols - 1) = ""
        
'        dtp.Enabled = True
        
        mstrSQL = " Select ��ĿID,��Ŀ���,������,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀֵ��,��Ŀ��λ From �����¼��Ŀ A " & _
                  " Where Nvl(��Ŀ����,1)=1 And Nvl(A.Ӧ�÷�ʽ,0)=1 And Nvl(a.���ò���,0) In (0,[3]) And A.����ȼ�>=[1] " & _
                  " And (A.���ÿ���=1 Or (A.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=A.��Ŀ��� And D.����id=[2]))) " & _
                  " Order By A.������,A.��Ŀ���"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("����ȼ�").Value), Val(mrsParam("����id").Value), IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2))
        If RS.BOF = False Then
            Do While Not RS.EOF
                
                If Val(Vsf.RowData(Vsf.Rows - 1)) <> 0 Then Vsf.Rows = Vsf.Rows + 1
                
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��¼��) = zlCommFun.NVL(RS("������").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.������Ŀ) = zlCommFun.NVL(RS("��Ŀ����").Value)

                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀ����) = zlCommFun.NVL(RS("��Ŀ����").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀ����) = zlCommFun.NVL(RS("��Ŀ����").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��ĿС��) = zlCommFun.NVL(RS("��ĿС��").Value)
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀ��ʾ) = zlCommFun.NVL(RS("��Ŀ��ʾ").Value)
                

                If zlCommFun.NVL(RS("��Ŀֵ��")) <> "" Then

                    varAry = Split(zlCommFun.NVL(RS("��Ŀֵ��")), ";")

                    For lngLoop = 0 To UBound(varAry)
                        If Left(varAry(lngLoop), 1) = "��" Then
                            mblnDefault = True
                            
                            strTmp = Mid(varAry(lngLoop), 2)

                            If Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀȱʡ) = "" Then
                                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀȱʡ) = strTmp
                            Else
                                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀȱʡ) = Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀȱʡ) & ";" & strTmp
                            End If
                        Else
                            strTmp = CStr(varAry(lngLoop))
                        End If

                        If Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀֵ��) = "" Then
                            Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀֵ��) = strTmp
                        Else
                            Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀֵ��) = Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀֵ��) & "|" & strTmp
                        End If
                    Next
                End If

                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀid) = zlCommFun.NVL(RS("��Ŀid").Value)
'                If mbytMode = 1 Then
'                    vsf.TextMatrix(vsf.Rows - 1, mCol.��¼���) = vsf.TextMatrix(vsf.Rows - 1, mCol.��Ŀȱʡ)
'                    vsf.TextMatrix(vsf.Rows - 1, mCol.�Ƿ�䶯) = "1"
'                Else
                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��¼���) = ""
'                End If

                Vsf.TextMatrix(Vsf.Rows - 1, mCol.��Ŀ��λ) = zlCommFun.NVL(RS("��Ŀ��λ").Value)
                Vsf.RowData(Vsf.Rows - 1) = zlCommFun.NVL(RS("��Ŀ���").Value, 0)
                
                RS.MoveNext
            Loop
        End If
        
'    Else
        '20090914:ÿ������������¼,���ڶ��ָ�Ϊȱʡֵ,��������
        'dtp.Value = Format(mrsParam("ʱ��").Value, dtp.CustomFormat)
'        dtp.Enabled = False
'        cboBaby.Enabled = False
'    End If
    
'    Call ReadData
    OpenPatientMap = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ReadDrink() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim RS As ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    Dim int������ As Integer
    Dim int������ As Integer
    Dim strValue As String
    
    
    On Error GoTo ErrHand
    
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm") & ":00"
    strEnd = Format(DateAdd("n", 1, CDate(strStart)), "yyyy-MM-dd HH:mm") & ":00"
    
    For intLoop = 1 To Vsf.Rows - 1
        If Val(Vsf.RowData(intLoop)) = 6 Then
            int������ = intLoop
        End If
        If Val(Vsf.RowData(intLoop)) = 7 Then
            int������ = intLoop
        End If
    Next
    
    If int������ = 0 And int������ = 0 Then Exit Function
    
    strSQL = "Select zl_PatitDrink([1],[2],[3],[4]) As ���� From Dual"
    
    Set RS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), CDate(strStart), CDate(strEnd))
    If RS.BOF = False Then
        
        strTmp = zlCommFun.NVL(RS("����"))
        
        If strTmp <> "" Then
            
            strValue = Trim(Split(strTmp, ";")(0))
            If int������ > 0 Then Vsf.TextMatrix(int������, mCol.��¼���) = strValue
            
            strValue = Trim(Split(strTmp, ";")(1))
            If UBound(Split(strTmp, ";")) > 1 Then strValue = strValue & "��"
            
            If int������ > 0 Then Vsf.TextMatrix(int������, mCol.��¼���) = strValue
            
            
        End If
    End If
        
    ReadDrink = True
                            
    Exit Function
    
ErrHand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '��ȡָ����¼id��ָ����Ļ������ݣ�ע����Ҫ���ôˣ�û������ʱ��Ҫ������Ŀ�б�
    '------------------------------------------------------------------------------------------------------------------
    Dim RS As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lngColor As Long
    Dim strStart As String
    Dim strEnd As String
    Dim blnAllow As Boolean
    
    On Error GoTo ErrHand
        
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    strEnd = Format(DateAdd("n", 1, CDate(strStart)), "yyyy-MM-dd HH:mm:ss")
    
    mint������ = 0
    
    Vsf.Rows = 2
    Vsf.RowData(1) = 0
    Vsf.Cell(flexcpText, 1, 1, 1, Vsf.Cols - 1) = ""
    
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Select X. *, " & _
                     "Y.��Ŀ���, " & _
                     "Y.��Ŀ����, " & _
                     "Y.��Ŀ��λ, " & _
                     "Y.������, " & _
                     "Y.��Ŀ��ʾ, " & _
                     "Y.��Ŀֵ��, " & _
                     "Y.��Ŀ����, " & _
                     "Y.��Ŀ����, " & _
                     "Y.��ĿС��, " & _
                     "Y.��Ŀid,Y.������Ŀ,Y.������,Y.��Ŀ���� " & _
                "From "
    
    If mint����Ӧ�� = 2 Then
        mstrSQL = mstrSQL & _
                    "(Select A.��¼���� As ��¼���, " & _
                                 "C.������ As ��¼��, " & _
                                 "C.����ʱ�� As ��¼ʱ��,Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,b.��¼���� As ���,b.��¼���," & _
                                 "A.��Ŀ���, " & _
                                 "C.����ʱ�� As �������,A.��¼id,a.δ��˵�� " & _
                             "From ���˻������� A, ���˻������� B,���˻����¼ C " & _
                            "Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                                  "AND A.��¼���� = 1 " & _
                                  "AND C.������Դ = 2 " & _
                                  "AND NVL(A.��¼���,0) <> 1 " & _
                                  "AND C.ID = [1] And A.��¼���=[5] "
    Else
    
        mstrSQL = mstrSQL & _
                    "(Select A.��¼���� As ��¼���, " & _
                                 "C.������ As ��¼��, " & _
                                 "C.����ʱ�� As ��¼ʱ��,Decode(a.��¼����,Null,'',A.���²�λ) As ��λ,Decode(a.��Ŀ���,2,'',-1,'',b.��¼����) As ���,Decode(a.��Ŀ���,2,'',-1,'',b.��¼���) As ��¼���," & _
                                 "A.��Ŀ���, " & _
                                 "C.����ʱ�� As �������,A.��¼id,a.δ��˵�� " & _
                             "From ���˻������� A, ���˻������� B,���˻����¼ C " & _
                            "Where C.ID = A.��¼id And b.��¼id(+)=a.��¼id And b.��¼���(+)=a.��¼��� And b.��¼���(+) =1 " & _
                                  "AND A.��¼���� = 1 " & _
                                  "AND C.������Դ = 2 " & _
                                  "AND ((NVL(A.��¼���,0) <> 1 And a.��Ŀ���>0) or a.��Ŀ���=-1) " & _
                                  "AND C.ID = [1] And A.��¼���=[5] "
                                  
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("�汾").Value) = 0 Then
    
        mstrSQL = mstrSQL & _
                    " And a.��ֹ�汾 Is Null And b.��ֹ�汾 Is Null "
                    
    Else
                
        mstrSQL = mstrSQL & _
                    " And Nvl(a.��ʼ�汾,1)<=[4] And Nvl(a.��ֹ�汾,10000)>[4] And Nvl(b.��ʼ�汾,1)<=[4] And Nvl(b.��ֹ�汾,10000)>[4] "
    End If

    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = mstrSQL & _
                        " and Decode(a.��Ŀ���,2,-1,a.��Ŀ���)=b.��Ŀ���(+)) X, " & _
                      "�����¼��Ŀ Y " & _
                "Where Y.��Ŀ��� = X.��Ŀ���(+) And Nvl(y.Ӧ�÷�ʽ,0)=1 And Nvl(y.���ò���,0) In (0,[6]) And (Y.���ÿ���=1 Or (Y.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=Y.��Ŀ��� And D.����id=[2])))  " & _
                        "AND Y.����ȼ� >=[3]  " & _
                "Order By Y.������,Y.��Ŀ���,X.��¼��� "
                
    'Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(tbs.Tag), Val(mrsParam("����id").Value), Val(mrsParam("����ȼ�").Value), Val(mrsParam("�汾").Value), Val(tbs.SelectedItem.Tag), IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2))
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(tbcPage.Selected.Tag), Val(mrsParam("����id").Value), Val(mrsParam("����ȼ�").Value), Val(mrsParam("�汾").Value), Val(tbs.SelectedItem.Tag), IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2))
    If RS.BOF = False Then
        
'        mrsParam("��¼id").Value = Val(tbs.Tag)
        mrsParam("��¼id").Value = Val(tbcPage.Selected.Tag)
        
        With Vsf
            Do While Not RS.EOF
                
                blnAllow = False
                If zlCommFun.NVL(RS("��Ŀ����"), 1) = 2 Then
                    If zlCommFun.NVL(RS("��¼���")) <> "" Then
                        blnAllow = True
                    End If
                Else
                    blnAllow = True
                End If
                
                If blnAllow Then
                    If Val(.RowData(.Rows - 1)) <> 0 Then .Rows = .Rows + 1
                    
                    Call WriteItemData(RS, .Rows - 1)

                End If
                
                RS.MoveNext
            Loop
        
            Call .Body.AutoSize(mCol.��¼���, mCol.��¼���)
        End With
        
        Call ExecuteCommand("��ʷ����", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(Vsf.Row)))
    End If

    ReadData = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function WriteItem(ByVal rsData As ADODB.Recordset, ByVal intRow As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngColor As Long
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    mblnDefault = False
    With Vsf
        .RowData(intRow) = zlCommFun.NVL(rsData("��Ŀ���"))
        
        .TextMatrix(intRow, mCol.��¼��) = zlCommFun.NVL(rsData("������").Value)
        .TextMatrix(intRow, mCol.������Ŀ) = zlCommFun.NVL(rsData("��Ŀ����"))
        .TextMatrix(intRow, mCol.��Ŀ����) = zlCommFun.NVL(rsData("��Ŀ����"))
        .TextMatrix(intRow, mCol.��Ŀ����) = zlCommFun.NVL(rsData("��Ŀ����"))
        .TextMatrix(intRow, mCol.��ĿС��) = zlCommFun.NVL(rsData("��ĿС��"), 0)
        .TextMatrix(intRow, mCol.��Ŀ��ʾ) = zlCommFun.NVL(rsData("��Ŀ��ʾ"))
        .TextMatrix(intRow, mCol.��Ŀ����) = zlCommFun.NVL(rsData("��Ŀ����"), 1)
        .TextMatrix(intRow, mCol.��Ŀid) = zlCommFun.NVL(rsData("��Ŀid"), 0)
                            
        If zlCommFun.NVL(rsData("��Ŀֵ��")) <> "" Then
        
            varAry = Split(zlCommFun.NVL(rsData("��Ŀֵ��")), ";")
                            
            For lngLoop = 0 To UBound(varAry)
                If Left(varAry(lngLoop), 1) = "��" Then
                    mblnDefault = True
                    strTmp = Mid(varAry(lngLoop), 2)
                    
                    If .TextMatrix(intRow, mCol.��Ŀȱʡ) = "" Then
                        .TextMatrix(intRow, mCol.��Ŀȱʡ) = strTmp
                    Else
                        .TextMatrix(intRow, mCol.��Ŀȱʡ) = .TextMatrix(intRow, mCol.��Ŀȱʡ) & ";" & strTmp
                    End If
                Else
                    strTmp = CStr(varAry(lngLoop))
                End If
                                    
                If .TextMatrix(intRow, mCol.��Ŀֵ��) = "" Then
                    .TextMatrix(intRow, mCol.��Ŀֵ��) = strTmp
                Else
                    .TextMatrix(intRow, mCol.��Ŀֵ��) = .TextMatrix(intRow, mCol.��Ŀֵ��) & "|" & strTmp
                End If
            Next
        End If
        
        If zlCommFun.NVL(rsData("��Ŀ���")) = 7 And zlCommFun.NVL(rsData("������Ŀ"), 0) = 1 Then mint������ = intRow
        .TextMatrix(intRow, mCol.��Ŀ��λ) = zlCommFun.NVL(rsData("��Ŀ��λ").Value)
        
    End With
    
    WriteItem = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function


Private Function WriteItemData(ByVal rsData As ADODB.Recordset, ByVal intRow As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngColor As Long
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim strTmp As String
    
    With Vsf
        .RowData(intRow) = zlCommFun.NVL(rsData("��Ŀ���"))
        
        Call WriteItem(rsData, intRow)
        
        If zlCommFun.NVL(rsData("��¼���")) <> "" Or zlCommFun.NVL(rsData("δ��˵��")) <> "" Then
            
            .TextMatrix(intRow, mCol.��λ) = zlCommFun.NVL(rsData("��λ"))
            
            Select Case zlCommFun.NVL(rsData("��Ŀ���"))
            Case 9
                If Right(zlCommFun.NVL(rsData("��¼���")), 2) = "/C" Then
                
                    .TextMatrix(intRow, mCol.��¼���) = Left(zlCommFun.NVL(rsData("��¼���")), Len(zlCommFun.NVL(rsData("��¼���"))) - 2)
                    .TextMatrix(intRow, mCol.���) = "/C"
                    
                ElseIf Right(zlCommFun.NVL(rsData("��¼���")), 1) = "C" Then
                    .TextMatrix(intRow, mCol.��¼���) = Left(zlCommFun.NVL(rsData("��¼���")), Len(zlCommFun.NVL(rsData("��¼���"))) - 1)
                    .TextMatrix(intRow, mCol.���) = "C"
                Else
                    .TextMatrix(intRow, mCol.��¼���) = zlCommFun.NVL(rsData("��¼���"))
                    .TextMatrix(intRow, mCol.���) = zlCommFun.NVL(rsData("���"))
                End If
            Case 10
                If Right(zlCommFun.NVL(rsData("��¼���")), 2) = "/E" Then
                    .TextMatrix(intRow, mCol.��¼���) = Left(zlCommFun.NVL(rsData("��¼���")), Len(zlCommFun.NVL(rsData("��¼���"))) - 2)
                    .TextMatrix(intRow, mCol.���) = "/E"
                ElseIf Right(zlCommFun.NVL(rsData("��¼���")), 1) = "E" Then
                    .TextMatrix(intRow, mCol.��¼���) = Left(zlCommFun.NVL(rsData("��¼���")), Len(zlCommFun.NVL(rsData("��¼���"))) - 1)
                    .TextMatrix(intRow, mCol.���) = "E"
                ElseIf Right(zlCommFun.NVL(rsData("��¼���")), 1) = "*" Then
                    .TextMatrix(intRow, mCol.���) = "*"
                Else
                    .TextMatrix(intRow, mCol.��¼���) = zlCommFun.NVL(rsData("��¼���").Value)
                    .TextMatrix(intRow, mCol.���) = zlCommFun.NVL(rsData("���").Value)
                End If
            Case Else
            
                lngColor = GridTextColor(zlCommFun.NVL(rsData("��Ŀ����")), zlCommFun.NVL(rsData("��¼���").Value))
                .Cell(flexcpForeColor, intRow, mCol.��¼���, intRow, mCol.��¼���) = lngColor
                
                .TextMatrix(intRow, mCol.��¼���) = zlCommFun.NVL(rsData("��¼���").Value)
                .TextMatrix(intRow, mCol.���) = zlCommFun.NVL(rsData("���").Value)
                .TextMatrix(intRow, mCol.δ��˵��) = zlCommFun.NVL(rsData("δ��˵��").Value)
            End Select
        End If
            
    End With
    
    WriteItemData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function SignName() As Boolean
    Dim RS As New ADODB.Recordset
    Dim oSign As cEPRSign
    Dim strSource As String
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim strDate As String
    Dim strStart As String
    Dim lngLoop As Long
    
    On Error GoTo ErrHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    ReDim Preserve strSQL(1 To 1)
    
    strDate = Format(dtp.Value, "yyyy-MM-dd HH:mm")
    strStart = strDate & ":00"
    strSource = ""
    
    '��鵱ǰ�Ƿ��Ѿ�ǩ����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From ���˻������� a,���˻����¼ b Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��¼����=5 And Nvl(a.��ʼ�汾,1)=Nvl(b.���汾,1)"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), CDate(strStart), Val(mrsParam("Ӥ��").Value), Val(mrsParam("�汾").Value))
    If RS.BOF = False Then
        ShowSimpleMsg "��ǰû����Ҫǩ������Ϣ��"
        Exit Function
    End If
        
    '��ȡҪǩ��������
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.�޸�ʱ��" & vbNewLine & _
             " From ���˻������� a,���˻����¼ b " & vbNewLine & _
             " Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��ֹ�汾 Is Null" & vbNewLine & _
             " Order by A.��Ŀ���"
    If mblnMoved_HL Then
        gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
    End If
    
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), CDate(strStart), Val(mrsParam("Ӥ��").Value))
    If RS.BOF = False Then
        Do While Not RS.EOF
            For lngLoop = 0 To RS.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(RS.Fields(lngLoop).Value, ""))
            Next
            RS.MoveNext
        Loop
    End If
    Debug.Print "ǩ����" & Now & vbCrLf & strSource
    If strSource = "" Then
        MsgBox "��ǰû����Ҫǩ������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    '76223:������,2014-08-05,����ǩ�����ʱ�����Ϣ
    '------------------------------------------------------------------------------------------------------------------
    Set oSign = frmCaseTendSign.ShowMe(Me, mstrPrivs, strSource, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("����id").Value))
    If Not oSign Is Nothing Then

        mstrSQL = "ZL_���ӻ����¼_SignName("
        mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
        mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
        mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
        mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        mstrSQL = mstrSQL & "'" & oSign.���� & "',"
        mstrSQL = mstrSQL & "'" & oSign.ǩ����Ϣ & "',"
        mstrSQL = mstrSQL & oSign.֤��ID & ","
        mstrSQL = mstrSQL & oSign.ǩ����ʽ & ",'" & oSign.ʱ��� & "','" & oSign.ʱ�����Ϣ & "')"

        strSQL(ReDimArray(strSQL)) = mstrSQL

        'ִ��
        '--------------------------------------------------------------------------------------------------------------
        blnTran = True
        gcnOracle.BeginTrans
        For lngLoop = 1 To UBound(strSQL)
            If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
        Next
        gcnOracle.CommitTrans
        blnTran = False
        
        SignName = True
    End If
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
ErrHand:

    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    

End Function

Private Function UnSignName() As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim strSource As String
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim strDate As String
    Dim strStart As String
    Dim lngLoop As Long
    Dim RS As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    ReDim Preserve strSQL(1 To 1)
    strDate = Format(dtp.Value, "yyyy-MM-dd HH:mm")
    strStart = strDate & ":00"
    
    '��鵱ǰ�Ƿ��Ѿ�ǩ����
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From ���˻������� a,���˻����¼ b Where b.����id=[1] And b.��ҳid=[2] And b.����ʱ��=[3] And Nvl(b.Ӥ��,0)=[4] And a.��¼id=b.ID And a.��¼����=5"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), CDate(strStart), Val(mrsParam("Ӥ��").Value), Val(mrsParam("�汾").Value))
    If RS.BOF Then
        ShowSimpleMsg "��ǰû����Ҫȡ����ǩ����"
        Exit Function
    End If
    
    
    '����ǵ���ǩ��,����Ҫ��֤
    '------------------------------------------------------------------------------------------------------------------
    If Val(Me.Tag) > 0 Then
        '����ǩ����֤
        Err.Clear
        If gobjTendESign Is Nothing Then
            On Error Resume Next
            Set gobjTendESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err.Clear
            On Error GoTo 0
            If Not gobjTendESign Is Nothing Then Call gobjTendESign.Initialize(gcnOracle, glngSys)
        End If
        If Not gobjTendESign Is Nothing Then
            If Not gobjTendESign.CheckCertificate(gstrDBUser) Then Exit Function
        Else
            MsgBox "����ǩ������δ����ȷ��װ�����˲������ܼ�����", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "Zl_���ӻ����¼_Unsignname("
    mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
    mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
    mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
    mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
    strSQL(ReDimArray(strSQL)) = mstrSQL

    'ִ��
    '------------------------------------------------------------------------------------------------------------------
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    UnSignName = True
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
ErrHand:

    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    

End Function

Private Function SaveDataAll() As Boolean
    Dim RS As New ADODB.Recordset
    Dim blnTrans As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim intCol As Integer
    Dim intRow As Integer
    
    On Error GoTo ErrHand
        
    ReDim Preserve strSQL(1 To 1)
    
    If SaveData(strSQL) = False Then GoTo EndHand
    
    
    intCol = tbs.SelectedItem.Index
    
'    For intRow = 1 To tbs.Tabs.Count
'        If intRow <> intCol Then
'            tbs.Tabs(intRow).Selected = True
'
'            If SaveData(strSQL) = False Then GoTo errHand

'        End If
'    Next

    blnTrans = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTrans = False
    
    If mbytMode = 1 Or tbcPage.Selected.Tag = "" Then
    
        gstrSQL = "Select a.ID,b.��¼��� From ���˻����¼ a,���˻������� b Where a.ID=b.��¼id And a.����id=[1] And a.��ҳid=[2] And a.����ʱ��=[3] And Nvl(a.Ӥ��,0)=[4] And b.��¼����<>5 Group By a.id,b.��¼��� Order By a.id,b.��¼���"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), Val(mrsParam("Ӥ��")))
        If RS.BOF = False Then
'            tbs.Tag = Val(rs("ID").Value)
            tbcPage.Selected.Tag = Val(RS("ID").Value)
        End If
        
    End If
    
    tbs.Tabs(intCol).Selected = True
    
    SaveDataAll = True
    
    Exit Function
    
ErrHand:
    '������
    
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    
EndHand:
    
End Function

Private Function SaveData(ByRef strSQL() As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strDate As String
    Dim strStart As String
    Dim strEnd As String
    Dim lng����ID As Long
    Dim RS As New ADODB.Recordset
    Dim strTmp As String
    Dim intAllow As Integer
    
    On Error GoTo ErrHand
            
    mstrSQL = " Select D.ID,D.����,��ʼ,��ֹ" & _
            " From ���ű� D," & _
            "   (Select ����id,To_Date(To_Char(Min(��ʼʱ��), 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi') as ��ʼ,Max(Nvl(��ֹʱ��,Sysdate+100)) as ��ֹ" & _
            "    From ���˱䶯��¼" & _
            "    Where ��ʼʱ�� is Not Null And ����ID=[1] And ��ҳID=[2]" & _
            "    Group by ����id) L" & _
            " Where L.����id=D.ID"
    mstrSQL = mstrSQL & " And To_Date('" & Format(dtp.Value, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') between ��ʼ and ��ֹ "
    
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If RS.BOF = False Then
        lng����ID = RS("ID").Value
    Else
        ShowSimpleMsg "����ʱ�䲻�ܴ��ڵ�ǰʱ�����С�ڿ�ʼʱ�䣡"
        Exit Function
    End If
    
    strDate = Format(dtp.Value, "yyyy-MM-dd HH:mm")
    strStart = strDate & ":00"
    strEnd = Format(DateAdd("n", 1, CDate(strDate)), "yyyy-MM-dd HH:mm") & ":00"
    intAllow = IIf(InStr(mstrPrivs, "���˻����¼") > 0, 1, 0)
    
    '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
    If Not CheckTime(Val(mrsParam("����id")), Val(mrsParam("��ҳid")), Mid(strStart, 1, 16), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")) Then Exit Function
    
    Dim int��¼��� As Long
    int��¼��� = Val(tbs.SelectedItem.Tag)
    
    If Val(mrsParam("��¼id").Value) > 0 Then
        
        mstrSQL = "Select ID From ���˻����¼ Where ����id = [1] And ��ҳid = [2] And Nvl(Ӥ��, 0) = Nvl([3], 0) And ������Դ = 2 And ����ʱ�� = [4] And ID<>[5]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), Val(mrsParam("Ӥ��")), CDate(strStart), Val(mrsParam("��¼id").Value))
        If RS.BOF = False Then
            If MsgBox("��ǰ����ʱ�仹���������ļ�¼���Ƿ񸲸ǣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
            
            'ɾ����ͬid,��������id��¼�ķ���ʱ���Ӥ�����
            mstrSQL = "Zl_���˻����¼_UpdateReplace(" & Val(mrsParam("��¼id").Value) & "," & Val(RS("ID").Value) & "," & Val(mrsParam("Ӥ��").Value) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
            strSQL(ReDimArray(strSQL)) = mstrSQL
        Else
            mstrSQL = "Zl_���˻����¼_UpdateReplace(" & Val(mrsParam("��¼id").Value) & ",0," & Val(mrsParam("Ӥ��").Value) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
            strSQL(ReDimArray(strSQL)) = mstrSQL
        End If
    Else
        mstrSQL = "Select ID From ���˻����¼ Where ����id = [1] And ��ҳid = [2] And Nvl(Ӥ��, 0) = Nvl([3], 0) And ������Դ = 2 And ����ʱ�� = [4]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, gstrSysName, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), Val(mrsParam("Ӥ��")), CDate(strStart))
        If RS.BOF = False Then
            If MsgBox("��ǰ����ʱ�仹���������ļ�¼���Ƿ񸲸ǣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Function
            End If
            
            'ɾ����ͬid,��������id��¼�ķ���ʱ���Ӥ�����
            mstrSQL = "Zl_���˻����¼_UpdateReplace(" & Val(mrsParam("��¼id").Value) & "," & Val(RS("ID").Value) & "," & Val(mrsParam("Ӥ��").Value) & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
            strSQL(ReDimArray(strSQL)) = mstrSQL
        End If
    End If
    
    For lngLoop = 1 To Vsf.Rows - 1
        
        If Val(Vsf.RowData(lngLoop)) <> 0 And Val(Vsf.TextMatrix(lngLoop, mCol.�Ƿ�䶯)) = 1 Then
'        If Val(vsf.RowData(lngLoop)) <> 0 Then

            mstrSQL = "Zl_���˻����¼_UpdateRecord("
            mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
            mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "1,"
            mstrSQL = mstrSQL & Val(Vsf.RowData(lngLoop)) & ","
            
            If Val(Vsf.RowData(lngLoop)) = -1 Then
                mstrSQL = mstrSQL & "1,"
            Else
                mstrSQL = mstrSQL & "0,"
            End If
            
            Select Case Val(Vsf.RowData(lngLoop))
            Case 9, 10
                strTmp = Trim(Vsf.TextMatrix(lngLoop, mCol.���))
            Case Else
                strTmp = ""
            End Select
            
            If Trim(Vsf.TextMatrix(lngLoop, mCol.��¼���)) <> "" Then strTmp = Vsf.TextMatrix(lngLoop, mCol.��¼���) & strTmp
            mstrSQL = mstrSQL & "'" & strTmp & "',"
            mstrSQL = mstrSQL & "'" & Trim(Vsf.TextMatrix(lngLoop, mCol.��λ)) & "'," & intAllow & "," & IIf(IsNumeric(strTmp), 0, 1) & "," & int��¼��� & ",'" & Vsf.TextMatrix(lngLoop, mCol.δ��˵��) & "')"
                
            strSQL(ReDimArray(strSQL)) = mstrSQL
            
            
            If (Val(Vsf.RowData(lngLoop)) = 1 Or (Val(Vsf.RowData(lngLoop)) = 2) And mint����Ӧ�� = 2) Then

                mstrSQL = "Zl_���˻����¼_UpdateRecord("
                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & IIf(Val(Vsf.RowData(lngLoop)) = 2, -1, Val(Vsf.RowData(lngLoop))) & ","
                mstrSQL = mstrSQL & "1,"
                                                
                If Trim(Vsf.TextMatrix(lngLoop, mCol.���)) <> "" And Trim(Vsf.TextMatrix(lngLoop, mCol.��¼���)) <> "" Then
                    Select Case Val(Vsf.TextMatrix(lngLoop, mCol.��Ŀ����))
                    Case 0          '��ֵ
                        strTmp = Val(Trim(Vsf.TextMatrix(lngLoop, mCol.���)))
                    Case 1          '�ı�
                        strTmp = Trim(Trim(Vsf.TextMatrix(lngLoop, mCol.���)))
                    End Select
                    
                    mstrSQL = mstrSQL & "'" & strTmp & "',"
                    mstrSQL = mstrSQL & "NULL," & intAllow & "," & IIf(IsNumeric(strTmp), 0, 1) & "," & int��¼��� & ",Null)"
                Else
                    mstrSQL = mstrSQL & "NULL,"
                    mstrSQL = mstrSQL & "NULL," & intAllow & ",0," & int��¼��� & ",Null)"
                End If
                
                strSQL(ReDimArray(strSQL)) = mstrSQL
            End If
            
        End If
    Next
                
    
    Vsf.Cell(flexcpText, 1, mCol.�Ƿ�䶯, Vsf.Rows - 1, mCol.�Ƿ�䶯) = ""
    
    SaveData = True
    
    Exit Function
    
ErrHand:
    '������
    
    If ErrCenter = 1 Then
        Resume
    End If

    
End Function

Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim RS As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    Dim lngRow As Long
    Dim strStart As String
    Dim strEnd As String
    Dim curDate As Date
    Dim intDay As Integer
    Dim strPar As String
    
    On Error GoTo ErrHand


    Select Case strCmd
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        If InitData = False Then Exit Function
        Call InitMenuBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
    
        If mrsParam("ʱ��").Value = "" Then mrsParam("ʱ��").Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        dtp.Value = Format(mrsParam("ʱ��").Value, dtp.CustomFormat)
        txtDay.Text = Val(zlDatabase.GetPara("��ʷ��������", glngSys, 1255, "1"))
                
        '��Ժ��ʼ����;��Ժ��������;��Ժ����;��Ժ����;ת������;ת������
        '------------------------------------------------------------------------------------------------------------------
        strPar = zlDatabase.GetPara("������ʾ��Χ", glngSys, 1262, "10000")
        mrsParam("��Ժ����").Value = Val(Mid(strPar, 1, 1))
        mrsParam("��Ժ����").Value = Val(Mid(strPar, 2, 1))
        mrsParam("ת������").Value = Val(Mid(strPar, 4, 1))
        On Error Resume Next
        mrsParam("����Ʋ���").Value = Val(Mid(strPar, 5, 1))
        On Error GoTo 0
        mrsParam("ת������").Value = Val(zlDatabase.GetPara("���ת������", glngSys, 1262, 7))
        
        curDate = zlDatabase.Currentdate
        intDay = Val(zlDatabase.GetPara("��Ժ���˽������", glngSys, 1262, 7))
        mrsParam("��Ժ��������").Value = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
        intDay = Val(zlDatabase.GetPara("��Ժ���˿�ʼ���", glngSys, 1262, 30))
        mrsParam("��Ժ��ʼ����").Value = Format(CDate(mrsParam("��Ժ��������").Value) - intDay, "yyyy-MM-dd 00:00:00")
        
        '------------------------------------------------------------------------------------------------------------------
        '88776:���￨���Ȼ�ȡ�в�������Ϊ���ݱ�
        gstrSQL = "Select NVL(���ų���,8) ���ų��� From ҽ�ƿ���� Where �ض���Ŀ = '���￨'"
        Call zlDatabase.OpenRecordset(RS, gstrSQL, Me.Caption)
        If RS.EOF = False Then
            mint���￨���볤�� = Val("" & RS!���ų���)
        Else
            mint���￨���볤�� = 8
        End If
        
        mstr���￨��ĸǰ׺ = UCase(zlDatabase.GetPara(27, glngSys))
        
        '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select ��Ժ����ID from ������ҳ Where ����id=[1] And ��ҳid=[2] "
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value))
        If RS.BOF = False Then
            mrsParam("����id").Value = Val(zlCommFun.NVL(RS("��Ժ����ID").Value))
        End If
        
        mint����Ӧ�� = 2
        gstrSQL = "Select a.Ӧ�÷�ʽ From �����¼��Ŀ a Where a.��Ŀ���=-1"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If RS.BOF = False Then
            mint����Ӧ�� = zlCommFun.NVL(RS("Ӧ�÷�ʽ").Value, 2)
        End If
        
        '��ȡ�����б������浽���ؼ�¼���У��Է��������ѡ��
        If ReadPatient = False Then Exit Function
        
        '��λ����ǰ����
        mrsPatient.Filter = ""
        mrsPatient.Filter = "����id=" & Val(mrsParam("����id").Value)
        If mrsPatient.RecordCount > 0 Then mlngRowNum = Val(mrsPatient("ID").Value)
        mrsPatient.Filter = ""
    
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
        
        tbs.Tabs.Clear
        tbs.Tabs.Add 1, , "1"
        tbs.Tabs(1).Tag = 1
        cbo.Clear
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ�»�����Ϣ"
        
        Call OpenPatientMap(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ��¼"
        '����ʱ���ȡID,���봰��ʱ
        
        tbcPage.Item(0).Tag = ""
        tbcPage.Item(0).Selected = True
        For intLoop = tbcPage.ItemCount - 1 To 1 Step -1
            tbcPage.RemoveItem intLoop
        Next
        
        gstrSQL = "Select b.ID From ���˻����¼ b Where b.����id=[1] And b.��ҳid=[2] And Nvl(b.Ӥ��,0)=[3] And b.����ʱ��=[4]"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value), CDate(mrsParam("ʱ��").Value & ":00"))
        If RS.BOF = False Then
            tbcPage.Item(0).Tag = RS("ID").Value
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
    
        '��ȡָ����¼id�Ļ������ݣ�����������
        
        cbo.Clear
'        gstrSQL = "Select Distinct a.��¼id,Nvl(a.��ʼ�汾,1) As ��ʼ�汾,c.��¼�� As ǩ���� From ���˻������� a,���˻����¼ b,���˻������� c Where a.��¼����<>5 And c.��¼����(+)=5 And c.��¼id(+)=a.��¼id And c.��ʼ�汾(+)=Nvl(a.��ʼ�汾,1) And a.��¼id=b.ID And b.����id=[1] And b.��ҳid=[2] And Nvl(b.Ӥ��,0)=[3] And b.����ʱ��=[4] Order By Nvl(a.��ʼ�汾,1) Desc"
'        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value), CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")))
'
        gstrSQL = "Select Distinct a.��¼id,Nvl(a.��ʼ�汾,1) As ��ʼ�汾,c.��¼�� As ǩ����,b.����ʱ�� From ���˻������� a,���˻����¼ b,���˻������� c Where a.��¼����<>5 And c.��¼����(+)=5 And c.��¼id(+)=a.��¼id And c.��ʼ�汾(+)=Nvl(a.��ʼ�汾,1) And a.��¼id=b.ID And b.ID=[1] Order By Nvl(a.��ʼ�汾,1) Desc"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tbcPage.Selected.Tag))
        
        If RS.BOF = False Then
            
            dtp.Value = Format(RS("����ʱ��").Value, "yyyy-MM-dd HH:mm")
            gstrSQL = "Select a.��Ŀid As ֤��id,Nvl(a.��ʼ�汾,1) As ��ʼ�汾 From ���˻������� a Where a.��¼����=5 And a.��¼id=[1] Order By a.��ʼ�汾 Desc "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(RS("��¼id").Value))
            If rsTmp.BOF = False Then
                If mbytMode = 4 Then Me.Caption = "ȡ������ " & rsTmp("��ʼ�汾").Value & " �桱��ǩ����"
                Me.Tag = zlCommFun.NVL(rsTmp("֤��id").Value, 0)
            End If
            
            Do While Not RS.EOF
                
                If zlCommFun.NVL(RS("ǩ����").Value, "") = "" Then
                    cbo.AddItem "�� " & RS("��ʼ�汾").Value & " ��"
                Else
                    cbo.AddItem "�� " & RS("��ʼ�汾").Value & " ��(��ǩ��)"
                End If
                
                cbo.ItemData(cbo.NewIndex) = RS("��ʼ�汾").Value
                RS.MoveNext
            Loop
    
        End If
        If cbo.ListCount = 0 And mbytMode = 4 Then
            ShowSimpleMsg "Ŀǰ��û���κ�ǩ���İ汾��"
            Exit Function
        End If
        If cbo.ListCount > 0 And cbo.ListIndex = -1 Then cbo.ListIndex = 0
        
        '��ȡ��¼��
        '------------------------------------------------------------------------------------------------------------------
        tbs.Tabs.Clear
        intLoop = 0
        gstrSQL = "Select a.ID,b.��¼��� From ���˻����¼ a,���˻������� b Where a.ID=b.��¼id And a.ID=[1] And b.��¼����<>5 Group By a.id,b.��¼��� Order By a.id,b.��¼���"
        Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(tbcPage.Selected.Tag))
        If RS.BOF = False Then
            Do While Not RS.EOF
                intLoop = intLoop + 1
                tbs.Tabs.Add intLoop, , CStr(intLoop)
                tbs.Tabs(intLoop).Tag = RS("��¼���").Value
                tbcPage.Selected.Tag = RS("ID").Value
    '            tbs.Tag = rs("ID").Value
                RS.MoveNext
            Loop
        Else
            tbs.Tabs.Add 1, , "1"
            tbs.Tabs(1).Tag = 1
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ�������"
        
        Call ReadData
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʷ����"
        
        lbl(12).Caption = "ǰ       ��ġ�" & Vsf.TextMatrix(Vsf.Row, mCol.������Ŀ) & "����ʷ��¼���"
        strStart = Format(DateAdd("d", 0 - Val(txtDay.Text), CDate(varParam(0))), "yyyy-MM-dd HH:mm:ss")
        strEnd = Format(DateAdd("n", -1, CDate(varParam(0))), "yyyy-MM-dd HH:mm:ss")
        
        mclsVsfHistory.ClearGrid
        
        '��ʾָ��ǰN���ָ��ָ������
        strSQL = _
            "Select a.����ʱ�� As ��¼ʱ��, b.��¼���� As ��ʷ���" & vbNewLine & _
            "From ���˻����¼ a, ���˻������� b" & vbNewLine & _
            "Where a.����ʱ�� Between [3] And [4] And a.����id=[1] And a.��ҳid=[2] And a.Id = b.��¼id And b.��Ŀ��� = [5] And" & vbNewLine & _
            "           b.��¼���� Is Not Null" & vbNewLine & _
            "Order By a.����ʱ�� Desc"
        Set RS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), CDate(strStart), CDate(strEnd), Val(varParam(1)))
        
        If RS.BOF = False Then
            Call mclsVsfHistory.LoadGrid(RS)
        End If
        
        vsfHistory.AutoSize 1, 1
        
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
    
    End Select
        
    ExecuteCommand = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click()
    Dim lng�汾 As Long
    
    If mblnStartUp Then Exit Sub
    If mblnReading Then Exit Sub

    lng�汾 = cbo.ItemData(cbo.ListIndex)
    If Val(mrsParam("�汾").Value) = lng�汾 Then Exit Sub
    mrsParam("�汾").Value = lng�汾
    
    Call ReadData
    
End Sub

Private Sub cboBaby_Click()
    If mblnStartUp Then Exit Sub
    If Val(mrsParam("Ӥ��").Value) = cboBaby.ListIndex Then Exit Sub
    mrsParam("Ӥ��").Value = cboBaby.ListIndex
    
    '�����������ģʽ���͸���ʱ���ȡ��Ӧ�ļ����¼id
    Call ExecuteCommand("��ȡ��¼")
    Call ExecuteCommand("�������")
    Call ExecuteCommand("��ȡ����")
    Call ExecuteCommand("��ȡ�������")
    
    DataChanged = False
    
'    Call ReadData
    
'    If mbytMode = 1 Or mbytMode = 2 Then DataChanged = True
    
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intCol As Integer
    Dim intRow As Integer
    Dim blnCancel As Boolean
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        
        Call ReadDrink
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewParent                 '���������¼
        
        tbcPage.InsertItem tbcPage.ItemCount, "   " & CStr(tbcPage.ItemCount + 1) & "   ", picPane(2).hWnd, 0
        
        Call ExecuteCommand("�������")
            
        tbcPage.Item(tbcPage.ItemCount - 1).Selected = True
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem                   '������¼���
        
        tbs.Tabs.Add tbs.Tabs.Count + 1, , CStr(tbs.Tabs.Count + 1)
        tbs.Tabs(tbs.Tabs.Count).Selected = True
        For intRow = 1 To tbs.Tabs.Count + 10
            If intRow <> tbs.Tabs(intRow).Tag Then
                tbs.Tabs(tbs.Tabs.Count).Tag = intRow
                Exit For
            End If
        Next
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append                '�����Ŀ
        
        Dim rsData As New ADODB.Recordset
        Dim rsTmp As New ADODB.Recordset
        Dim strNotItem As String
        Dim intLoop As Integer
        Dim strTmp As String
        
        strNotItem = ""
        For intLoop = 1 To Vsf.Rows - 1
            
            If Val(Vsf.TextMatrix(intLoop, mCol.��Ŀ����)) = 2 Then
                strNotItem = strNotItem & "," & Val(Vsf.RowData(intLoop))
            End If
            
        Next
        If strNotItem <> "" Then strNotItem = Mid(strNotItem, 2)

        Set rsData = GetGridItem(Val(mrsParam("����ȼ�").Value), Val(mrsParam("����id").Value), IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2), 2, strNotItem, False)
        
        If rsData.BOF = False Then
            If ShowTxtSelDialog(Me, Nothing, "����,1500,0,1;��λ,900,0,0", Me.Name & "\������Ŀѡ��", "�������ѡ��һ��������Ŀ��", rsData, rsTmp, 6000, 3000, , , 2, False) Then
                If rsTmp.BOF = False Then
                    
                    '���ں��ʵ�λ�ô����һ����
                    strTmp = ""
                    
                    For intLoop = 1 To Vsf.Rows - 1
                        If strTmp = "" Then
                            If Vsf.TextMatrix(intLoop, mCol.��¼��) = zlCommFun.NVL(rsTmp("������").Value) Then
                                strTmp = Vsf.TextMatrix(intLoop, mCol.��¼��)
                            End If
                        ElseIf strTmp <> Vsf.TextMatrix(intLoop, mCol.��¼��) Then
                            Exit For
                        End If
                    Next
                    '��д��Ŀ����
                    If intLoop = Vsf.Rows Then
                        Vsf.Rows = Vsf.Rows + 1
                        intLoop = Vsf.Rows - 1
                    Else
                        Vsf.Body.AddItem "", intLoop
                    End If
                    If WriteItem(rsTmp, intLoop) Then
                        Call LocationGrid(Vsf, intLoop, Vsf.Col)
                        If mblnDefault Then DataChanged = mblnDefault
                    End If

                End If
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                'ɾ����Ŀ
        
        With Vsf
            If Val(.TextMatrix(.Row, mCol.��Ŀ����)) = 2 Then
                
                    
                '����Ƿ������ݣ����������ʱ������ɾ��
                '���α���֮ǰ�������Լ���ǰ�����������ݣ����֮Ϊ������
                
                If Trim(.TextMatrix(.Row, mCol.��¼���)) <> "" Or .TextMatrix(.Row, mCol.�Ƿ�䶯) = "1" Then
                    ShowSimpleMsg "�Բ�����Ҫɾ������������ݻ�����ǰ�����ݣ�"
                    Exit Sub
                End If
                
                If MsgBox("ȷʵҪɾ����ǰ�ı����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                
                .RemoveItem .Row
                
                Call vsf_AfterRowColChange(0, 0, .Row, .Col)
                
            End If
        End With
        
'        '�ж��Ƿ�ֻ��һ������
'        For intRow = 1 To vsf.Rows - 1
'            If intRow <> vsf.Row Then
'                If Val(vsf.RowData(intRow)) = Val(vsf.RowData(vsf.Row)) Then
'                    vsf.Body.RemoveItem vsf.Row
'                    DataChanged = True
'                    Exit For
'                End If
'            End If
'        Next
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save
    
        Select Case mbytMode
        Case 1, 2
            
            blnCancel = False
            Call vsf_ValidateEdit(Vsf.Row, Vsf.Col, blnCancel)
            If blnCancel = False Then mblnOk = SaveDataAll
                        
            If mblnOk Then
                RaiseEvent AfterDataChanged
            End If
        End Select
        
        If mblnOk Then DataChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle
        
        Select Case mbytMode
        Case 1, 2
            
            Call OpenPatientMap(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
            Call tbcPage_SelectedChanged(tbcPage.Selected)
            
            DataChanged = False
        End Select
        
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign                  'ǩ��
        
        If mbytMode = 3 Then
            mblnOk = SignName
            If mblnOk Then

                RaiseEvent AfterDataChanged

            
                DataChanged = False
                Unload Me
            End If
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread                 'ȡ��ǩ��
        If mbytMode = 4 Then
            mblnOk = UnSignName
            If mblnOk Then
                
                RaiseEvent AfterDataChanged
                
                DataChanged = False
                Unload Me
            End If
        End If
    
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsThis.RecalcLayout
            
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        
        If mlngRowNum = 1 Then mlngRowNum = mrsPatient.RecordCount + 1
        mrsPatient.Filter = ""
        mrsPatient.Filter = "ID<" & mlngRowNum
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveLast
            mlngRowNum = Val(mrsPatient("ID").Value)
            txt.Text = zlCommFun.NVL(mrsPatient("����").Value)
            mrsParam("����id").Value = Val(mrsPatient("����id").Value)
            mrsParam("��ҳid").Value = Val(mrsPatient("��ҳid").Value)
            mrsParam("Ӥ��").Value = 0
            mrsParam("�汾").Value = 0
            Select Case CStr(mrsPatient("����").Value)
            Case "����", "��������", "��Ժ����"
                mrsParam("��Ժ").Value = 1
            Case Else
                mrsParam("��Ժ").Value = 0
            End Select

            Call OpenPatientMap(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
            Call ExecuteCommand("��ȡ��¼")
            Call ExecuteCommand("�������")
            Call ExecuteCommand("��ȡ����")
            Call ExecuteCommand("��ȡ�������")
            DataChanged = False
            
            txt.Tag = ""
        End If
        mrsPatient.Filter = ""
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        
        If mlngRowNum = mrsPatient.RecordCount Then mlngRowNum = 0
        mrsPatient.Filter = ""
        mrsPatient.Filter = "ID>" & mlngRowNum
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveFirst
            mlngRowNum = Val(mrsPatient("ID").Value)
            txt.Text = zlCommFun.NVL(mrsPatient("����").Value)
            mrsParam("����id").Value = Val(mrsPatient("����id").Value)
            mrsParam("��ҳid").Value = Val(mrsPatient("��ҳid").Value)
            mrsParam("Ӥ��").Value = 0
            mrsParam("�汾").Value = 0
            Select Case CStr(mrsPatient("����").Value)
            Case "����", "��������", "��Ժ����"
                mrsParam("��Ժ").Value = 1
            Case Else
                mrsParam("��Ժ").Value = 0
            End Select

            Call OpenPatientMap(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
            Call ExecuteCommand("��ȡ��¼")
            Call ExecuteCommand("�������")
            Call ExecuteCommand("��ȡ����")
            Call ExecuteCommand("��ȡ�������")
            DataChanged = False
            
            txt.Tag = ""
        End If
        mrsPatient.Filter = ""
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
        Call LocationObj(txt)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Help_Help
    
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Exit
    
        Unload Me
        Exit Sub
        
    End Select

End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
'    Dim lngLeft As Long
'    Dim lngTop  As Long
'    Dim lngRight  As Long
'    Dim lngBottom  As Long
'
'    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'
'    On Error Resume Next
'
'    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error Resume Next
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        
        Control.Visible = (mbytMode < 3)
        Control.Enabled = Control.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem, conMenu_Edit_NewParent
        Control.Visible = (mbytMode < 3)
        Control.Enabled = Control.Visible And mblnChanged = False
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Append
        Control.Visible = (mbytMode < 3)
        Control.Enabled = Control.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Save
        
        Control.Visible = (mbytMode < 3)
        Control.Enabled = mblnChanged And Control.Visible

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Transf_Cancle
    
        Control.Visible = (mbytMode < 3)
        Control.Enabled = mblnChanged And Control.Visible
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                    'ɾ������ӵĻ��Ŀ
        
        Control.Visible = (mbytMode < 3)
        Control.Enabled = (Val(Vsf.TextMatrix(Vsf.Row, mCol.��Ŀ����)) = 2) And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Tool_Sign                      'ǩ��
        
        Control.Visible = (mbytMode = 3)
        Control.Enabled = Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread                   'ȡ��ǩ��
    
        Control.Visible = (mbytMode = 4)
        Control.Enabled = Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        Control.Checked = (mstrFindKey = Control.Parameter)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Location
        
        Control.Enabled = (mblnChanged = False)
        cmd.Enabled = Control.Enabled And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward
        
        Control.Enabled = (mrsPatient.RecordCount > 1 And mblnChanged = False And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5) And mlngRowNum > 1)
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward
        
        Control.Enabled = (mrsPatient.RecordCount > 1 And mblnChanged = False And (mbytMode = 1 Or mbytMode = 2 Or mbytMode = 5) And mlngRowNum < mrsPatient.RecordCount)
        
    End Select
End Sub

Private Sub cmd_Click()
    Dim RS As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim mstrSort As String
    
    '------------------------------------------------------------------------------------------------------------------
    mrsPatient.Filter = ""
    If mrsPatient.RecordCount > 0 Then
        mrsPatient.MoveFirst
        If ShowTxtSelDialog(Me, txt, "����,1200,0,0;����,1200,0,1;�Ա�,600,0,0;����,1800,0,0;סԺ��,1080,0,0", Me.Name & "\�����嵥ѡ��", "�������ѡ��һ�����ˡ�", mrsPatient, RS, 5600, 4500, , CStr(mlngRowNum), 2, True) Then
            
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(RS("����").Value)

            mrsParam("����id").Value = Val(RS("����id").Value)
            mrsParam("��ҳid").Value = Val(RS("��ҳid").Value)
            mrsParam("Ӥ��").Value = 0
            Select Case CStr(RS("����").Value)
            Case "����", "��������", "��Ժ����"
                mrsParam("��Ժ").Value = 1
            Case Else
                mrsParam("��Ժ").Value = 0
            End Select
            
            Call OpenPatientMap(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
            
            Call ExecuteCommand("��ȡ��¼")
            Call ExecuteCommand("�������")
            Call ExecuteCommand("��ȡ����")
            Call ExecuteCommand("��ȡ�������")
            
            DataChanged = False
    
            txt.Tag = ""
        End If
    End If
    mrsPatient.Filter = ""
    
    Call LocationObj(txt)

    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
End Sub




Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub dtp_Change()
    dtp.Tag = "Changed"
    DataChanged = True
    
    Call ExecuteCommand("��ʷ����", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(Vsf.Row)))
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        
        Vsf.Row = 1
        Vsf.Col = mCol.��¼���
        Vsf.SetFocus
    End If
    
End Sub


Private Sub dtp_LostFocus()

'    If dtp.Tag = "Changed" Then
'        '��ȡ��ʱ����ֵ
'        dtp.Tag = ""
'        Call ReadData
'    End If
    
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpMain, 2, 15, 100, Me.ScaleWidth, 150)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ�������棿", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    
    On Error Resume Next
    zlCommFun.OpenIme False
    
    Call zlDatabase.SetPara("���ҷ���", mstrFindKey, glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    Call zlDatabase.SetPara("��ʷ��������", Val(txtDay.Text), glngSys, 1255, InStr(mstrPrivs, "����ѡ������") > 0)
    
    Call SaveWinState(Me, App.ProductName)
    
    Set mrsPatient = Nothing
    Set mobjFindKey = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0
    
        fraInfo.Move 0, -90, picPane(Index).Width
        tbcPage.Move 0, fraInfo.Top + fraInfo.Height, picPane(Index).Width, picPane(Index).Height - (fraInfo.Top + fraInfo.Height)
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        fra.Move 0, -90, picPane(Index).Width
        vsfHistory.Move 15, fra.Top + fra.Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (fra.Top + fra.Height) - 30
        vsfHistory.AutoSize 1, 1
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
        fraTime.Move 0, -90, picPane(Index).Width
'        lbl(0).Move 30, fraTime.Top + fraTime.Height + 45
'        tbs.Move tbs.Left, fraTime.Top + fraTime.Height, picPane(Index).Width - tbs.Left
        tbs.Width = fraTime.Width - tbs.Left - cboBaby.Width - 90
        cboBaby.Left = tbs.Left + tbs.Width + 30
        Vsf.Move 0, fraTime.Top + fraTime.Height, picPane(Index).Width, picPane(Index).Height - (fraTime.Top + fraTime.Height)
        cbo.Move fraTime.Width - cbo.Width - 90
        
    End Select
    
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnStartUp Then Exit Sub
    
    Call ExecuteCommand("�������")
    
    If Val(tbcPage.Selected.Tag) > 0 Then
        Call ExecuteCommand("��ȡ����")
    End If
    
    Call ExecuteCommand("��ȡ�������")
    
    DataChanged = False
End Sub

Private Sub tbs_Click()
    Call ReadData
End Sub

Private Sub txt_Change()
    txt.Tag = "Changed"
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim bytMode As Byte
    Dim lng����ID As Long
    Dim strInput As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If txt.Tag = "Changed" And txt.Text <> "" Then
            If InStr(txt.Text, "'") Then
                ShowSimpleMsg "������������зǷ��ַ� ' ��"
                Exit Sub
            End If
            
            Select Case mstrFindKey
'            Case "����id"
'                strInput = "����id=" & Val(txt.Text)
'                bytMode = 2
'            Case "�����"
'                strInput = "�����=" & Val(txt.Text)
'                bytMode = 4
            Case "��  ��"
                strInput = "����='" & Trim(txt.Text) & "'"
                bytMode = 5
            Case "סԺ��"
                strInput = "סԺ��=" & Val(txt.Text)
                bytMode = 3
            Case "���￨"
                strInput = "���￨��='" & Trim(txt.Text) & "'"
                bytMode = 1
            End Select
                        
        End If

    ElseIf mstrFindKey = "���￨" And txt.Tag = "Changed" And txt.Text <> "" Then
        If Len(txt.Text) = mint���￨���볤�� - 1 And KeyAscii <> 8 Or KeyAscii = 13 And txt.Text <> "" Then
            If KeyAscii <> 13 Then
                txt.Text = txt.Text & Chr(KeyAscii)
                txt.SelStart = Len(txt.Text)
                KeyAscii = 0
            End If

            strInput = "���￨��='" & Trim(txt.Text) & "'"
            bytMode = 1
        End If
    End If
    
    If strInput <> "" Then
        txt.Tag = ""
        mrsPatient.Filter = ""
        mrsPatient.Filter = strInput
        If mrsPatient.RecordCount > 0 Then
            mrsPatient.MoveFirst
            lng����ID = Val(mrsPatient("����id").Value)
            mlngRowNum = Val(mrsPatient("ID").Value)
            
            txt.Text = zlCommFun.NVL(mrsPatient("����").Value)
            txt.Tag = ""
            
            mrsParam("����id").Value = Val(mrsPatient("����id").Value)
            mrsParam("��ҳid").Value = Val(mrsPatient("��ҳid").Value)
            mrsParam("Ӥ��").Value = 0
            Select Case CStr(mrsPatient("����").Value)
            Case "����", "��������", "��Ժ����"
                mrsParam("��Ժ").Value = 1
            Case Else
                mrsParam("��Ժ").Value = 0
            End Select
            
            Call OpenPatientMap(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
        Else
            ShowSimpleMsg "û���ҵ����������Ĳ��ˣ�"
            txt.Text = mstrSvr����
        End If
        mrsPatient.Filter = ""

        Call LocationObj(txt)
        
    End If

    Exit Sub

ErrHand:
End Sub

Private Sub txtShow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtShow(Index).Locked Then
        glngTXTProc = GetWindowLong(txtShow(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtShow(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtShow(Index).Locked Then
        Call SetWindowLong(txtShow(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub udnDay_Change()
    Call ExecuteCommand("��ʷ����", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(Vsf.Row)))
End Sub

Private Sub vsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
    Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
    Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        
    DataChanged = True
    
    Select Case Col
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.��¼���
    
        Call Vsf.Body.AutoSize(mCol.��¼���, mCol.��¼���)
        
        Vsf.TextMatrix(Row, mCol.δ��˵��) = ""
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.���
        
        Vsf.TextMatrix(Row, mCol.δ��˵��) = ""
        
        If Val(Vsf.RowData(Row)) = 10 Then
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.δ��˵��
    
    End Select
    
    Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strValue As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    With Vsf
        .ComboList(mCol.��λ) = ""
        .EditMode(mCol.��λ) = 0
        .ComboList(mCol.���) = ""
        .EditMode(mCol.���) = 0
            
        Select Case Val(.RowData(NewRow))
        Case 1
            .ComboList(mCol.��λ) = "����|Ҹ��|����"
            .EditMode(mCol.��λ) = 1
            
            .ComboList(mCol.���) = ""
            .EditMode(mCol.���) = 1
        Case 2
            .ComboList(mCol.��λ) = " |����"
            .EditMode(mCol.��λ) = 1
            If mint����Ӧ�� = 2 Then
                .ComboList(mCol.���) = ""
                .EditMode(mCol.���) = 1
            End If
        Case 3
            .ComboList(mCol.��λ) = "��������|������"
            .EditMode(mCol.��λ) = 1
            
            .ComboList(mCol.���) = ""
            .EditMode(mCol.���) = 1
        Case 9
            .ComboList(mCol.���) = " |C|/C"
            .EditMode(mCol.���) = 1
        Case 10
            .ComboList(mCol.���) = " |*|E|/E"
            .EditMode(mCol.���) = 1
        Case Else
            If Val(.TextMatrix(NewRow, mCol.��Ŀ����)) = 2 Then
                gstrSQL = " Select ��λ From ���²�λ Where ��Ŀ���=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�û��Ŀ��Ӧ�Ĳ�λ", CLng(.RowData(NewRow)))
                If Err = 0 Then
                    Do While Not rsTemp.EOF
                        strValue = strValue & "|" & rsTemp!��λ
                        rsTemp.MoveNext
                    Loop
                    If strValue <> "" Then
                        .ComboList(mCol.��λ) = Mid(strValue, 2)
                        .EditMode(mCol.��λ) = 1
                    End If
                End If
            End If
        End Select
        
        Select Case Trim(.TextMatrix(NewRow, mCol.���))
        Case "*"
            .EditMode(mCol.��¼���) = 0
        Case Else
            .EditMode(mCol.��¼���) = 1
        End Select
        
        Select Case Val(.TextMatrix(NewRow, mCol.��Ŀ��ʾ))
        Case 0  '�ı�
            strValue = ""
            If Val(.TextMatrix(NewRow, mCol.��Ŀ����)) >= 200 Then strValue = "..."
            .ComboList(mCol.��¼���) = strValue
            .Body.ColComboList(mCol.��¼���) = ""
        Case 1  '����
            .ComboList(mCol.��¼���) = ""
            .Body.ColComboList(mCol.��¼���) = ""
        Case 2  '��ѡ
            .ComboList(mCol.��¼���) = ""
            .Body.ColComboList(mCol.��¼���) = " |" & .TextMatrix(NewRow, mCol.��Ŀֵ��)
        Case 3  '��ѡ
            .ComboList(mCol.��¼���) = "..."
            .Body.ColComboList(mCol.��¼���) = "..."
        End Select
        
        Dim varAry As Variant
        Dim strTmp As String
        
        If Val(.TextMatrix(NewRow, mCol.��Ŀ����)) = 0 Then
            Select Case Val(.TextMatrix(NewRow, mCol.��Ŀ��ʾ))
            Case 0, 1
                If .TextMatrix(NewRow, mCol.��Ŀֵ��) <> "" Then
                    varAry = Split(.TextMatrix(NewRow, mCol.��Ŀֵ��), "|")
                    
                    If UBound(varAry) >= 1 Then
                        strTmp = Val(varAry(0)) & "��" & Val(varAry(1))
                    End If
                End If
            End Select
        End If
        
        stbThis.Panels(3).Text = "��Χ��" & strTmp
                
        Select Case Val(.RowData(NewRow))
        Case 1
            strTmp = "��Ǳ�ʾ�����µ��¶ȣ���λΪ�����µĲ�λ��"
        Case 2
            If mint����Ӧ�� = 2 Then
                strTmp = "��Ǳ�ʾ���ʵ�ֵ����������ͬʱ�ż�¼����"
            Else
                strTmp = "��λ��ѡ���Ƿ�ʹ������"
            End If
        Case 3
            strTmp = "��λΪ������ʽ����Ϊ������������������������"
        Case 9
            strTmp = "����е� C ��ʾ��������"
        Case 10
            strTmp = "����е� * ��ʾʧ����ٸ�; E ��ʾ�೦; /E ��ʾ�೦�����й��"
        Case Else
            strTmp = ""
        End Select
        
        stbThis.Panels(2).Text = strTmp
        
        If Val(.TextMatrix(NewRow, mCol.��Ŀ����)) = 1 Then
            zlCommFun.OpenIme True
        Else
            zlCommFun.OpenIme False
        End If
        
        
        If NewRow <> OldRow Then
            
            Call ExecuteCommand("��ʷ����", Format(dtp.Value, "yyyy-MM-dd HH:mm:ss"), Val(Vsf.RowData(NewRow)))
            
        End If
    End With
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case mCol.���, mCol.��¼���, mCol.δ��˵��
        If Vsf.TextMatrix(Row, Col) <> "" Then
            Vsf.TextMatrix(Row, Col) = ""
            Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"
            DataChanged = True
        End If
    End Select
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub


Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RS As New ADODB.Recordset
    Dim varAry As Variant
    Dim lngLoop As Long
    Dim objPoint As POINTAPI
    Dim lngX As Long
    Dim lngY As Long
    Dim lngCX As Long
    Dim strTmp As String
    
    Select Case Col
    Case mCol.��¼���
        strTmp = frmWordsEditor.ShowMe(Me, mrsParam!����ID, mrsParam!��ҳID, Vsf.TextMatrix(Row, mCol.��¼���))
        If strTmp = "" Then Exit Sub
        Vsf.EditText = strTmp
        Vsf.TextMatrix(Row, mCol.��¼���) = strTmp
        Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"
        mblnChanged = True
        
    Case mCol.δ��˵��
    
        gstrSQL = "Select ����,����,RowNum As ID,1 As ĩ�� From ��������˵��"
        If ShowGrdSelectDialog(Me, Vsf, "����,3000,0,0", Me.Name & "\��������˵��", "�������ѡ��һ��δ��¼˵����", gstrSQL, RS, 4500, 4500, False, 2) Then
            Vsf.EditText = zlCommFun.NVL(RS("����").Value)
            Vsf.Cell(flexcpData, Row, Col) = zlCommFun.NVL(RS("����").Value)
            Vsf.TextMatrix(Row, Col) = zlCommFun.NVL(RS("����").Value)
            Vsf.TextMatrix(Row, mCol.��¼���) = ""
            Vsf.TextMatrix(Row, mCol.���) = ""
            Vsf.TextMatrix(Row, mCol.��λ) = ""
            
            Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"

            mblnChanged = True
        End If
        
    Case Else
        '��Ը�ѡ���
        Call CreateParam(RS, "ID", adBigInt)
        Call CreateParam(RS, "ĩ��", adTinyInt)
        Call CreateParam(RS, "����", adVarChar, 200)
        Call CreateParam(RS, "ѡ��", adTinyInt)
        RS.Open
        If Vsf.TextMatrix(Row, mCol.��Ŀֵ��) <> "" Then
    
            strTmp = ";" & Vsf.TextMatrix(Row, Col) & ";"
    
            varAry = Split(Vsf.TextMatrix(Row, mCol.��Ŀֵ��), "|")
            For lngLoop = 0 To UBound(varAry)
                RS.AddNew
                RS("ID").Value = lngLoop
                RS("ĩ��").Value = 1
                RS("����").Value = CStr(varAry(lngLoop))
    
                If InStr(strTmp, ";" & CStr(varAry(lngLoop)) & ";") > 0 Then
                    RS("ѡ��").Value = 1
                Else
                    RS("ѡ��").Value = 0
                End If
            Next
            If RS.RecordCount > 0 Then RS.MoveFirst
        End If
    
        Call ClientToScreen(Vsf.hWnd, objPoint)
    
        lngX = objPoint.X * Screen.TwipsPerPixelX + Vsf.CellLeft
        lngY = objPoint.Y * Screen.TwipsPerPixelY + Vsf.CellTop + Vsf.CellHeight
    
        strTmp = ""
        
        lngCX = Vsf.Width - Vsf.Body.ColWidth(0) - Vsf.Body.ColWidth(1) - 75
        If lngCX < 3300 Then lngCX = 3300
        
        If frmSelectDialog.ShowSelect(Me, 2, RS, "����,3600,0,1", "����Ҫѡ�����Ŀǰ���ϡ�", lngX, lngY, lngCX, 3900, Vsf.CellHeight, , Me.Name & "\������ѡ��", , False, True) Then
            RS.Filter = ""
            RS.Filter = "ѡ��=1"
            If RS.RecordCount > 0 Then RS.MoveFirst
            Do While Not RS.EOF
                strTmp = strTmp & ";" & RS("����").Value
                RS.MoveNext
            Loop
    
            If strTmp <> "" Then strTmp = Mid(strTmp, 2)
            Vsf.TextMatrix(Row, Col) = strTmp
            Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"
            mblnChanged = True
        End If
    End Select
    
End Sub

Private Sub vsf_ChangeEdit()
    
    With Vsf
        Select Case .Col
        Case mCol.��¼���
            Select Case Val(.TextMatrix(.Row, mCol.��Ŀ��ʾ))
            Case 0
                .TextMatrix(.Row, mCol.��¼���) = .EditText
                Call .Body.AutoSize(mCol.��¼���, mCol.��¼���)
            Case 2
                '����
                .TextMatrix(.Row, mCol.��¼���) = .EditText
            End Select
            
            Vsf.TextMatrix(.Row, mCol.δ��˵��) = ""
        Case mCol.δ��˵��
            .TextMatrix(.Row, mCol.δ��˵��) = .EditText
            If Trim(.TextMatrix(.Row, mCol.δ��˵��)) <> "" Then
                
                .TextMatrix(.Row, mCol.��¼���) = ""
                .TextMatrix(.Row, mCol.���) = ""
                .TextMatrix(.Row, mCol.��λ) = ""

            End If
            
        Case mCol.���
            Select Case Val(.RowData(.Row))
            Case 9
                .TextMatrix(.Row, mCol.���) = .EditText
                
                Select Case Trim(.TextMatrix(.Row, mCol.���))
                Case "C"
                    .EditMode(mCol.��¼���) = 0
                    .TextMatrix(.Row, mCol.��¼���) = ""
                Case Else
                    .EditMode(mCol.��¼���) = 1
                End Select
                
            Case 10
                .TextMatrix(.Row, mCol.���) = .EditText
                Select Case Trim(.TextMatrix(.Row, mCol.���))
                Case "*"
                    .EditMode(mCol.��¼���) = 0
                    .TextMatrix(.Row, mCol.��¼���) = ""
                Case Else
                    .EditMode(mCol.��¼���) = 1
                End Select
            
            End Select
            Vsf.TextMatrix(.Row, mCol.δ��˵��) = ""
        End Select
        .TextMatrix(.Row, mCol.�Ƿ�䶯) = "1"
    End With
    
    DataChanged = True
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        
        If Col = mCol.δ��˵�� Or Col = mCol.��¼��� Then
            
            Vsf.Cell(flexcpData, Row, Col) = Vsf.EditText
            Vsf.TextMatrix(Row, Col) = Vsf.EditText
            Vsf.TextMatrix(Row, mCol.�Ƿ�䶯) = "1"
            
        End If
        
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(Vsf.TextMatrix(Row, mCol.��Ŀ����)) = 0 Then
            If Col = mCol.��� Or Col = mCol.��¼��� Then
                If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
            Else
                If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
            End If
        End If
    End If
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    On Error Resume Next
    
    If KeyAscii <> vbKeyReturn Then
        If Val(Vsf.TextMatrix(Row, mCol.��Ŀ����)) = 0 Then
'            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
        End If
    End If
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngColor As Long
    Dim varAry As Variant
    
    Select Case Col
    Case mCol.��¼���
        GoTo CheckPoint
    Case mCol.���
        
        Select Case Val(Vsf.RowData(Row))
        Case 1
            GoTo CheckPoint
        Case 2
            GoTo CheckPoint
        End Select
        
    End Select
    
    Exit Sub
    
CheckPoint:

    If Val(Vsf.TextMatrix(Row, mCol.��Ŀ����)) = 0 And Trim(Vsf.EditText) <> "" And IsNumeric(Vsf.EditText) Then
        Select Case Val(Vsf.TextMatrix(Row, mCol.��Ŀ��ʾ))
        Case 0, 1
            If Vsf.TextMatrix(Row, mCol.��Ŀֵ��) <> "" Then
                varAry = Split(Vsf.TextMatrix(Row, mCol.��Ŀֵ��), "|")
                
                If UBound(varAry) >= 1 Then
                    If Val(Vsf.EditText) < Val(varAry(0)) Or Val(Vsf.EditText) > Val(varAry(1)) Then
                        Vsf.TextMatrix(Row, Col) = Vsf.EditText

                        ShowSimpleMsg "��" & Vsf.TextMatrix(Row, mCol.������Ŀ) & " ���ķ�ΧӦ�ڣ�" & varAry(0) & "��" & varAry(1) & "��֮�䣡"

                    End If
                End If
                
            End If
            
            If CheckNumber(Val(Vsf.EditText), Val(Vsf.TextMatrix(Row, mCol.��Ŀ����)), Val(Vsf.TextMatrix(Row, mCol.��ĿС��))) = False Then
                
                Vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
        
        End Select
    End If
    
    Select Case Col
    Case mCol.��¼���
        
        lngColor = GridTextColor(Vsf.TextMatrix(Row, 2), Vsf.TextMatrix(Row, Col))
        Vsf.Cell(flexcpForeColor, Row, mCol.��¼���, Row, mCol.��¼���) = lngColor
        
    Case mCol.���

        
    End Select
    

End Sub


Private Sub vsfHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsfHistory.AutoSize 1, 1
End Sub

Private Sub vsfHistory_DblClick()
    With vsfHistory
        If Trim(.TextMatrix(.Row, 1)) <> "" Then
            Vsf.TextMatrix(Vsf.Row, mCol.��¼���) = .TextMatrix(.Row, 1)
            Vsf.TextMatrix(Vsf.Row, mCol.�Ƿ�䶯) = "1"
            Call Vsf.Body.AutoSize(mCol.��¼���, mCol.��¼���)
            DataChanged = True
        End If
        
    End With
End Sub

Private Sub vsfHistory_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsfHistory_DblClick
    End If
End Sub

Private Function CheckTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal strTime As String, ByVal strCurTime As String) As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    '���ݷ���ʱ������ڵ�ǰ���ҵ���Чʱ�䷶Χ��
    
    gstrSQL = " Select ��ʼԭ��,����ID,to_char(��ʼʱ��,'yyyy-MM-dd hh24:mi') AS ��ʼʱ��,to_char(nvl(��ֹʱ��,sysdate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS ��ֹʱ�� " & _
              " From ���˱䶯��¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]" & _
              " Order by ��ʼʱ��,��ʼԭ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ������Чʱ�䷶Χ", lng����ID, lng��ҳID)
    With rsTemp
        .Filter = "����ID=" & Val(mrsParam("����id"))
        Do While Not .EOF
            If strTime >= !��ʼʱ�� And strTime <= NVL(!��ֹʱ��, strCurTime) Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '�ҵ��˾��˳�
        If blnExist Then
            If Not IsAllowInput(lng����ID, lng��ҳID, strTime, strCurTime) Then
                MsgBox "����ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]", vbInformation, gstrSysName
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        'û�ҵ�,������ԭ�����׼ȷ����ʾ
        .Filter = "��ʼԭ��=1"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 1 And strTime < !��ʼʱ�� Then
                MsgBox "����ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ�����Ժʱ��:" & !��ʼʱ�� & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=2"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 2 And strTime < !��ʼʱ�� Then
                MsgBox "����ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ������ʱ��:" & !��ʼʱ�� & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "��ʼԭ��=10"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 10 And strTime > !��ֹʱ�� Then
                MsgBox "����ʱ��" & strTime & "����[����ʱ�䲻�ܴ��ڳ�Ժʱ��:" & !��ֹʱ�� & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '�������˵��
        MsgBox "����ʱ��" & strTime & "����[���ڵ�ǰ��������Чʱ�䷶Χ��]", vbInformation, gstrSysName
        GoTo exitHand
    End With
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
End Function
