VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathDesignOut 
   AutoRedraw      =   -1  'True
   Caption         =   "�����ٴ�·�����"
   ClientHeight    =   7830
   ClientLeft      =   2310
   ClientTop       =   2040
   ClientWidth     =   11565
   Icon            =   "frmPathDesignOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   11565
   Begin VB.PictureBox picCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4695
      ScaleWidth      =   14055
      TabIndex        =   4
      Top             =   2040
      Width           =   14055
      Begin VB.Frame fraSplit 
         BorderStyle     =   0  'None
         ForeColor       =   &H80000011&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   11
         Top             =   1560
         Width           =   9735
      End
      Begin VB.PictureBox picBottom 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2415
         ScaleWidth      =   12975
         TabIndex        =   5
         Top             =   2040
         Width           =   12975
         Begin VB.CommandButton cmdCheck 
            Caption         =   "��˲���"
            Height          =   300
            Index           =   1
            Left            =   8640
            TabIndex        =   15
            Top             =   360
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.CommandButton cmdCheck 
            Caption         =   "���ͨ��"
            Height          =   300
            Index           =   0
            Left            =   7440
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   1100
         End
         Begin zlCISPath.UCAdviceList ucAdvice 
            Height          =   1335
            Index           =   0
            Left            =   480
            TabIndex        =   13
            Top             =   720
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   2355
         End
         Begin VB.Frame fraSplit2 
            BorderStyle     =   0  'None
            Height          =   2655
            Left            =   6000
            MousePointer    =   9  'Size W E
            TabIndex        =   7
            Top             =   600
            Width           =   60
         End
         Begin VB.ComboBox cboTimes 
            Height          =   300
            Left            =   8880
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   120
            Width           =   3495
         End
         Begin zlCISPath.UCAdviceList ucAdvice 
            Height          =   1335
            Index           =   1
            Left            =   7200
            TabIndex        =   12
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2355
         End
         Begin VB.Label lblCurr 
            Caption         =   "��ǰҽ������"
            Height          =   255
            Left            =   480
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblChange 
            Caption         =   "ҽ���䶯����"
            Height          =   255
            Left            =   7440
            TabIndex        =   8
            Top             =   120
            Width           =   1215
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPath 
         Height          =   825
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   4695
         _cx             =   8281
         _cy             =   1455
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   10218651
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   10218651
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   1500
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         Editable        =   2
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
         FrozenRows      =   1
         FrozenCols      =   1
         AllowUserFreezing=   0
         BackColorFrozen =   14811105
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   4440
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdgXML 
      Left            =   1770
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7470
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPathDesignOut.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17489
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
   Begin VSFlex8Ctl.VSFlexGrid vsPathExport 
      Height          =   1305
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
      _cx             =   5530
      _cy             =   2302
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   10218651
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   10218651
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   1500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   2
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
      FrozenRows      =   1
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image ImgBranch 
      Height          =   240
      Left            =   3120
      Picture         =   "frmPathDesignOut.frx":0E1C
      Top             =   240
      Width           =   240
   End
   Begin XtremeSuiteControls.ShortcutCaption stcInfo 
      Height          =   390
      Left            =   1095
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   795
      Width           =   2955
      _Version        =   589884
      _ExtentX        =   5212
      _ExtentY        =   688
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      GradientColorLight=   16710907
      GradientColorDark=   16180453
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   285
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Bindings        =   "frmPathDesignOut.frx":766E
      Left            =   915
      Top             =   225
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPathDesignOut.frx":7682
   End
End
Attribute VB_Name = "frmPathDesignOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event DataChanged(ByVal ·��ID As Long)

Private WithEvents mfrmVersion          As frmVersionOut
Attribute mfrmVersion.VB_VarHelpID = -1
Private WithEvents mfrmPathStep         As frmPathStepEditOut
Attribute mfrmPathStep.VB_VarHelpID = -1
Private WithEvents mfrmPathItem         As frmPathItemEditOut
Attribute mfrmPathItem.VB_VarHelpID = -1
Private WithEvents mfrmEvalEdit         As frmEvaluateEdit
Attribute mfrmEvalEdit.VB_VarHelpID = -1
Private WithEvents mfrmAdviceContrast   As frmAdviceContrast
Attribute mfrmAdviceContrast.VB_VarHelpID = -1

Private mlng·��ID          As Long             '·��ID
Private mbytMode            As CONST_MODE       '����ģʽ��Mode_Show Ƕ��ʽ��Mode_Design ����ģʽ
Private mcolVersion         As Collection       '·���汾����
Private mcolItemRowCol      As Collection       'Value:Row,Col Key��"_ "& ��ĿID  LoadPathTableʱ��¼����Ŀ���к���
Private mcolItemID          As Collection       '��¼����һ�汾������ͬ�׶Σ���ͬ���࣬��ͬ���Ƶ�ҽ������Ŀ��ҽ�����ڲ������ĿID
Private mstrPrivs           As String           'Ȩ��
Private mblnReturn          As Boolean
Private mlngNewRow          As Long
Private mlngNewCol          As Long
Private mstrDeptInfo        As String           '·�����������ʾ�����ÿ�����Ϣ
Private mstrDiagInfo        As String           '·�����������ʾ�����ò�����Ϣ

Private mrsAdvice           As ADODB.Recordset  '��Ӧҽ����̬��¼��
Private mvEvalImport        As TYPE_PATH_EVAL   '������������
Private mblnEditable        As Boolean          '�Ƿ�����༭
Private mblnChange          As Boolean          '�����Ƿ��Ѿ��ı�
Private mstrDelStepIDs      As String           '��ɾ����ʱ��׶�ID��
Private mstrDelItemIDs      As String           '��ɾ����·����ĿID��
Private mstrChangeItemIDs   As String           '·���䶯��Ŀ��ID��
Private mblnNewVersion      As Boolean          '�Ƿ����°汾
Private mblnAddNew          As Boolean          '�ж��Ƿ���������֧
Private mlngDays            As Long
Private mstr��������        As String
Private mblnDiff            As Boolean
Private mbytFunc            As Byte             '����������Ŀ�䶯�Ͳ鿴�䶯��¼;1-�鿴�䶯��¼,2-��ʾ��Ŀ�䶯
Private marrTime            As Variant

Private Type PathTable_Clipboard
    Empty       As Boolean
    ��Ŀ��()    As TYPE_PATH_ITEM               '�����հ���Ŀ
    vStep       As TYPE_PATH_STEP               '·��ʱ��׶�
    BeginRow    As Long
End Type
Private mvClipboard As PathTable_Clipboard      '�ڲ�������

Private Const ROW_HEIGHT_MIN = 270
Private Const COl_WIDTH_BASE = 2000

Private Enum CONST_MODE
    Mode_Show = 0
    Mode_Design = 1
End Enum

Private Enum CONST_COLOR
    Color_NewBack = &HE1E1FF
    Color_AuditBack = &HE1FFE1
    Color_StopBack = &HE1E1E1
    Color_DiffBack = &HFAEADA                   'ǳ�� ҽ������Ŀ��֮ǰ�汾���ڲ���
    Color_NewLine = &H9B9BEC
    Color_AuditLine = &H9BEC9B
    Color_StopLine = &H9B9B9B
    Color_NeedAuditFore = &H9B9BEC              '·����Ŀ���ڴ����ҽ������Ŀ������ɫΪ��ɫ��
End Enum

Private Enum CONST_AREA
    Area_Cross = 0
    Area_Category = 1
    Area_Step = 2
    Area_Item = 3
End Enum

Private Enum CONST_FUNCTION
    '�ļ�-------------------------
    cmd_File_Save = 101
    cmd_File_SaveExit = 102
    cmd_File_CopyFrom = 111
    cmd_File_ImportXML = 112
    cmd_File_ExportXML = 121
    cmd_File_ExportExcel = 122
    cmd_File_PrintSetup = 131
    cmd_File_Preview = 132
    cmd_File_Print = 133
    cmd_File_Exit = 191
    '�༭-------------------------
    cmd_Edit_Undo = 301
    cmd_Edit_Redo = 302
    cmd_Edit_Copy = 303
    cmd_Edit_Paste = 304
    cmd_Edit_Caption = 310          '��ǩ
    cmd_Edit_Edit = 311             '����
    cmd_Edit_Insert = 312           '����
    cmd_Edit_InsertBefore = 3121    '��ǰ�����
    cmd_Edit_InsertAfter = 3122     '�ں������
    cmd_Edit_InsertBranch = 3123    '���ӷ�֧
    cmd_Edit_Delete = 313           'ɾ��
    cmd_Edit_Modify = 314           '�޸�
    cmd_Edit_Version = 320          '�汾ѡ��
    cmd_Edit_VersionInfo = 321      '�汾��Ϣ
    cmd_Edit_VersionNew = 322       '�汾���
    cmd_Edit_VersionDel = 323       '�汾ɾ��
    cmd_Edit_EvalImport = 324       '��������
    cmd_Edit_EvalStep = 325         '�׶�����
    cmd_Edit_EvalStepCopy = 326     '���ƽ׶�����
    cmd_Edit_BranchNew = 327        '��֧���
    cmd_Edit_BranchDel = 328        '��֧ɾ��
    cmd_Edit_Branch = 329           '��֧ѡ��
    cmd_Edit_ItemShow = 330         '��ʾ��Ŀ�䶯\������Ŀ�䶯
    '�鿴-------------------------
    cmd_View_ToolBar = 701
    cmd_View_ToolBar_Button = 7011
    cmd_View_ToolBar_Text = 7012
    cmd_View_ToolBar_Size = 7013
    cmd_View_StatusBar = 702
    cmd_View_Refresh = 791
    cmd_View_Find = 721
    '����-------------------------
    cmd_Help_Help = 901
    cmd_Help_Web = 902
    cmd_Help_Web_Home = 9021
    cmd_Help_Web_Forum = 9023
    cmd_Help_Web_Mail = 9022
    cmd_Help_About = 991
End Enum

Private Function CheckPathItem() As Boolean
'����: �������·���Ƿ������Ŀ�䶯
    Dim lngRow As Long, lngCol As Long
    Dim vItem As TYPE_PATH_ITEM
    With vsPath
        For lngRow = .FixedRows To .Rows - 1
            For lngCol = .FixedCols To .Cols - 1
                If TypeName(.Cell(flexcpData, lngRow, lngCol)) = TypeName(vItem) Then
                    vItem = .Cell(flexcpData, lngRow, lngCol)
                    If vItem.�����ҽ��IDs <> "" Then
                        CheckPathItem = True
                        Exit Function
                    End If
                End If
            Next
        Next
        MsgBox "�������ٴ�·��������·����Ŀ�䶯��", vbOKOnly + vbInformation, gstrSysName
    End With
End Function

Public Sub ShowDesign(frmParent As Object, ByVal lng·��ID As Long, ByVal strPrivs As String, Optional ByVal str�������� As String)
    mbytMode = Mode_Design
    mlng·��ID = lng·��ID
    mstrPrivs = strPrivs
    mstr�������� = str��������
    mbytFunc = 0

    Me.Show 1, frmParent
End Sub

Public Sub zlRefresh(ByVal lng·��ID As Long, ByVal strPrivs As String, Optional ByVal strDeptInfo As String, Optional ByVal strDiagInfo As String)
'�������鿴ģʽʱ���룬strDeptInfo=���ÿ�����Ϣ��strDiagInfo=���ò�����Ϣ
    mlng·��ID = lng·��ID
    mstrPrivs = strPrivs
    mstrDeptInfo = strDeptInfo
    mstrDiagInfo = strDiagInfo

    Call LoadPathVersion
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
'���ܣ���ʾģʽ�£���������������������
    Dim vVersion As TYPE_PATH_VERSION
    Dim objCombo As CommandBarComboBox
    Dim blnEnabled As Boolean

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If Not objCombo Is Nothing Then
        If objCombo.ListIndex > 0 Then
            vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
        End If
    End If

    Select Case Control.ID
        Case conMenu_File_ExportToXML                   '����ΪXML�ļ�
            If InStr(mstrPrivs, "����XML") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mlng·��ID <> 0 And vVersion.�汾�� > 0
            End If
        Case conMenu_File_BatPrint                      'ȫ�������Excel
            Control.Enabled = mbytMode = Mode_Show
        Case conMenu_Edit_Compend                       '���
            If InStr(mstrPrivs, "·�������") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mlng·��ID <> 0
            End If
        Case conMenu_Edit_Audit                         '���
            If InStr(mstrPrivs, "���") = 0 Then
                Control.Visible = False
            Else
                blnEnabled = mlng·��ID <> 0 And vVersion.�汾�� > 0 And vVersion.���ʱ�� = Empty
                If blnEnabled Then blnEnabled = objCombo.ListIndex = 1
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Edit_Untread                       'ȡ�����
            If InStr(mstrPrivs, "���") = 0 Then
                Control.Visible = False
            Else
                blnEnabled = mlng·��ID <> 0 And vVersion.�汾�� > 0 And vVersion.���ʱ�� <> Empty And vVersion.ͣ��ʱ�� = Empty
                If blnEnabled Then blnEnabled = objCombo.ListIndex = 1
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Edit_Stop                          'ͣ��
            If InStr(mstrPrivs, "ͣ��") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mlng·��ID <> 0 And vVersion.�汾�� > 0 _
                    And vVersion.���ʱ�� <> Empty And vVersion.ͣ��ʱ�� = Empty
            End If
        Case conMenu_Edit_Reuse                         'ȡ��ͣ��
            If InStr(mstrPrivs, "ͣ��") = 0 Then
                Control.Visible = False
            Else
                blnEnabled = mlng·��ID <> 0 And vVersion.�汾�� > 0 And vVersion.ͣ��ʱ�� <> Empty
                Control.Enabled = blnEnabled
            End If
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl, Optional ByVal blnIsAll As Boolean)
'���ܣ���ʾģʽ�£�ִ�������������
'      blnIsAll=�Ƿ����������Excel
    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Print
            Call FuncPathTableOutput(1, blnIsAll)
        Case conMenu_File_Preview
            Call FuncPathTableOutput(2, blnIsAll)
        Case conMenu_File_Excel
            Call FuncPathTableOutput(3, blnIsAll)
        Case conMenu_File_ExportToXML               '����XML
            Call FuncExportToXML
        Case conMenu_Edit_Compend                   '���
            '��������ֱ��ִ����
        Case conMenu_Edit_Audit                     '���
            Call FuncVersionAudit(True)
        Case conMenu_Edit_Untread                   'ȡ�����
            Call FuncVersionAudit(False)
        Case conMenu_Edit_Stop                      'ͣ��
            Call FuncVersionStop(True)
        Case conMenu_Edit_Reuse                     'ȡ��ͣ��
            Call FuncVersionStop(False)
    End Select
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCombo As CommandBarComboBox
    Dim objCustom As CommandBarControlCustom
    Dim lngCount As Long

    '�˵�����
    '-----------------------------------------------------
    If mbytMode = Mode_Design Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
        objMenu.ID = conMenu_FilePopup
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_File_Save, "����(&S)")
            Set objControl = .Add(xtpControlButton, cmd_File_SaveExit, "���沢�˳�(&X)")

            Set objControl = .Add(xtpControlButton, cmd_File_CopyFrom, "������·������(&C)��"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_ImportXML, "��&XML�ļ����롭")

            Set objControl = .Add(xtpControlButton, cmd_File_PrintSetup, "��ӡ����(&U)��"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_Preview, "Ԥ��(&V)")
            Set objControl = .Add(xtpControlButton, cmd_File_Print, "��ӡ(&P)")
            Set objControl = .Add(xtpControlButton, cmd_File_ExportExcel, "�����&Excel��")
            Set objControl = .Add(xtpControlButton, cmd_File_ExportXML, "����XM&L�ļ���"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
        End With

        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
        objMenu.ID = conMenu_EditPopup
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_Edit_Undo, "����(&U)")
            Set objControl = .Add(xtpControlButton, cmd_Edit_Redo, "����(&R)")
            Set objControl = .Add(xtpControlButton, cmd_Edit_Copy, "����(&C)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Paste, "ճ��(&V)")

            Set objControl = .Add(xtpControlButton, cmd_Edit_Edit, "����XXXX(&E)"): objControl.BeginGroup = True
            objControl.ShortcutText = "Enter"    'ֻ����ʾ
            Set objPopup = .Add(xtpControlButtonPopup, cmd_Edit_Insert, "����XXXX(&I)")
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBefore, "��ǰ�����(&1)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertAfter, "�ں������(&2)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBranch, "�����֧(&3)"): objControl.BeginGroup = True
            End With
            Set objControl = .Add(xtpControlButton, cmd_Edit_Modify, "�޸ķ���(&X)")
            objControl.ShortcutText = "Modify"    'ֻ����ʾ
            Set objControl = .Add(xtpControlButton, cmd_Edit_Delete, "ɾ��XXXX(&D)")
            objControl.ShortcutText = "Delete"    'ֻ����ʾ

            Set objControl = .Add(xtpControlButton, cmd_Edit_EvalImport, "������������(&P)"): objControl.BeginGroup = True
            Set objPopup = .Add(xtpControlSplitButtonPopup, cmd_Edit_EvalStep, "�׶���������(&J)")
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_EvalStepCopy, "����ǰ��׶���������(&C)")
            End With

            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionInfo, "��׼����(&B)")
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionNew, "�����µİ汾(&N)"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionDel, "ɾ����ǰ�汾(&M)")
            objControl.IconId = cmd_Edit_Delete
        End With

        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
        objMenu.ID = conMenu_ViewPopup
        With objMenu.CommandBar.Controls
            Set objPopup = .Add(xtpControlButtonPopup, cmd_View_ToolBar, "������(&T)")
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, cmd_View_ToolBar_Button, "��׼��ť(&S)", -1, False
                .Add xtpControlButton, cmd_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
                .Add xtpControlButton, cmd_View_ToolBar_Size, "��ͼ��(&B)", -1, False
            End With
            Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "��׼·���ο�")
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_View_StatusBar, "״̬��(&S)")
            Set objControl = .Add(xtpControlButton, conMenu_View_Difference, "��ʾ����")
            objControl.ID = conMenu_View_Difference
            objControl.ToolTipText = "�Բ�ͬ����ɫ������ʾҽ����������һ�汾�в������Ŀ"
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Contrast, "�ԱȲ鿴")
            objControl.ToolTipText = "ѡ�б���Ϊ��ɫ��ҽ������Ŀ����ִ�жԱȲ鿴"

            Set objControl = .Add(xtpControlButton, cmd_Edit_ItemShow, "��ʾ��Ŀ�䶯")
            objControl.IconId = cmd_Edit_ItemShow
            objControl.BeginGroup = True
            objControl.Parameter = "��ʾ"

            Set objControl = .Add(xtpControlButton, conMenu_View_Show, "�鿴�䶯��¼")
            objControl.IconId = cmd_View_Find
            objControl.BeginGroup = True
            objControl.Parameter = "��ʾ"

            Set objControl = .Add(xtpControlButton, cmd_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
        End With

        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
        objMenu.ID = conMenu_HelpPopup
        With objMenu.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_Help_Help, "��������(&H)")
            Set objPopup = .Add(xtpControlButtonPopup, cmd_Help_Web, "&WEB�ϵ�" & gstrProductName)
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, cmd_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
                .Add xtpControlButton, cmd_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
                .Add xtpControlButton, cmd_Help_Web_Mail, "���ͷ���(&M)", -1, False
            End With
            Set objControl = .Add(xtpControlButton, cmd_Help_About, "����(&A)��")
            objControl.BeginGroup = True
        End With

        '����������:������������
        '-----------------------------------------------------
        Set objBar = cbsMain.Add("������", xtpBarTop)
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, cmd_File_Save, "����")
            Set objControl = .Add(xtpControlButton, cmd_File_SaveExit, "�����˳�")

            Set objControl = .Add(xtpControlButton, cmd_Edit_Undo, "����"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Redo, "����")
            Set objControl = .Add(xtpControlButton, cmd_Edit_Copy, "����"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Paste, "ճ��")

            Set objControl = .Add(xtpControlLabel, cmd_Edit_Caption, "���ࣺ"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_Edit, "����")
            objControl.ToolTipText = "Enter"
            Set objPopup = .Add(xtpControlPopup, cmd_Edit_Insert, "����")
            objPopup.ID = cmd_Edit_Insert
            objPopup.IconId = cmd_Edit_Insert
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBefore, "��ǰ�����(&1)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertAfter, "�ں������(&2)")
                Set objControl = .Add(xtpControlButton, cmd_Edit_InsertBranch, "�����֧(&3)"): objControl.BeginGroup = True
            End With
            Set objControl = .Add(xtpControlButton, cmd_Edit_Modify, "�޸�")
            objControl.ToolTipText = "Modify"                                   'ֻ����ʾ
            Set objControl = .Add(xtpControlButton, cmd_Edit_Delete, "ɾ��")
            objControl.ToolTipText = "Delete"                                   'ֻ����ʾ

            Set objControl = .Add(xtpControlButton, cmd_Edit_EvalImport, "��������"): objControl.BeginGroup = True
            Set objPopup = .Add(xtpControlSplitButtonPopup, cmd_Edit_EvalStep, "�׶�����")
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, cmd_Edit_EvalStepCopy, "����ǰ��׶���������(&C)")
            End With

            Set objControl = .Add(xtpControlButton, conMenu_View_Difference, "��ʾ����")
            objControl.ToolTipText = "�Բ�ͬ����ɫ������ʾҽ����������һ�汾�в������Ŀ"
            objControl.ID = conMenu_View_Difference
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Contrast, "�ԱȲ鿴")
            objControl.ToolTipText = "ѡ�в����ҽ������Ŀ��ԱȲ鿴"

            Set objControl = .Add(xtpControlButton, cmd_Edit_ItemShow, "��ʾ��Ŀ�䶯")
            objControl.IconId = cmd_Edit_ItemShow
            objControl.BeginGroup = True
            objControl.Parameter = "��ʾ"

            Set objControl = .Add(xtpControlButton, conMenu_View_Show, "�鿴�䶯��¼")
            objControl.IconId = cmd_View_Find
            objControl.BeginGroup = True
            objControl.Parameter = "��ʾ"

            Set objControl = .Add(xtpControlButton, cmd_Help_Help, "����"): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_File_Exit, "�˳�")
            '����
            Set objControl = .Add(xtpControlLabel, 0, "����")
            objControl.IconId = cmd_View_Find
            objControl.Flags = xtpFlagRightAlign
            Set objCustom = .Add(xtpControlCustom, cmd_View_Find, "")
            objCustom.Handle = txtFind.Hwnd
            objCustom.Flags = xtpFlagRightAlign
        End With
    End If

    Set objBar = cbsMain.Add("�汾��", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "��      ��")
        objControl.IconId = cmd_Edit_Version
        Set objCombo = .Add(xtpControlComboBox, cmd_Edit_Version, "")    '�޷���ʾͼ��
        objCombo.Flags = xtpFlagControlStretched
        objCombo.DropDownListStyle = False
        If mbytMode = Mode_Design Then
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionInfo, "��׼����")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionNew, "�����汾")
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, cmd_Edit_VersionDel, "ɾ����ǰ�汾")
            objControl.IconId = cmd_Edit_Delete
            objControl.Flags = xtpFlagRightAlign
            objControl.Style = xtpButtonIconAndCaption
        End If
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    If mbytMode = Mode_Design Then
        With cbsMain.KeyBindings
            .Add FCONTROL, vbKeyS, cmd_File_Save    '����
            .Add FCONTROL, vbKeyZ, cmd_Edit_Undo    '����
            .Add FCONTROL, vbKeyR, cmd_Edit_Redo    '����
            .Add FCONTROL, vbKeyC, cmd_Edit_Copy    '����
            .Add FCONTROL, vbKeyV, cmd_Edit_Paste    'ճ��
            .Add FCONTROL, vbKeyF, cmd_View_Find    '����

            .Add FCONTROL, vbKeyE, cmd_Edit_EvalStep    '��ǰʱ��׶�������׼
            .Add FCONTROL, vbKeyB, cmd_Edit_InsertBefore
            .Add FCONTROL, vbKeyI, cmd_Edit_InsertAfter

            .Add 0, vbKeyF4, conMenu_View_Contrast       '�ԱȲ鿴
            .Add 0, vbKeyF5, conMenu_View_Refresh    'ˢ��
            .Add 0, vbKeyF3, conMenu_View_FindNext    '������һ��
            .Add 0, vbKeyF1, conMenu_Help_Help    '����
        End With

        '�ָ����̶���һЩ�˵�����
        cbsMain.ActiveMenuBar.Title = "�˵�"
        cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    ElseIf mbytMode = Mode_Show Then
        cbsMain.ActiveMenuBar.Visible = False
    End If

    For lngCount = 2 To cbsMain.count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagStretched + xtpFlagHideWrap
        If lngCount = 2 Then
            For Each objControl In cbsMain(lngCount).Controls
                If objControl.Type <> xtpControlLabel Then
                    If Not Between(objControl.ID, cmd_Edit_Undo, cmd_Edit_Paste) Then
                        objControl.Style = xtpButtonIconAndCaption
                    End If
                End If
            Next
        End If
    Next
End Sub

Private Sub cboTimes_Click()
    Dim strTmp As String
    Dim blnDo As Boolean

    If InStr(cboTimes.Text, cboTimes.Tag) = 0 Or cboTimes.Tag = "" Then
        cboTimes.Tag = marrTime(cboTimes.ListIndex)
        Call FuncShowAdvice(1)
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vArea As CONST_AREA, i As Long
    Dim strTmp As String

    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    zlCommFun.ShowTipInfo 0, ""
    vArea = GetArea(vsPath.Row, vsPath.Col)

    Select Case Control.ID
        Case cmd_File_Save, cmd_File_SaveExit    '����
            If Not CheckPathTable Then Exit Sub
            If Not SavePathTable() Then Exit Sub
            RaiseEvent DataChanged(mlng·��ID)
            If Control.ID = cmd_File_SaveExit Then Unload Me
        Case cmd_File_CopyFrom                  '��������
            Call FuncVersionCopy
        Case cmd_File_PrintSetup                '��ӡ����
            Call zlPrintSet
        Case cmd_File_Print                     '��ӡ
            Call FuncPathTableOutput(1)
        Case cmd_File_Preview                   'Ԥ��
            Call FuncPathTableOutput(2)
        Case cmd_File_ExportExcel               '����Excel
            Call FuncPathTableOutput(3)
        Case cmd_File_ExportXML                 '����XML
            Call FuncExportToXML
        Case cmd_File_ImportXML                 '����XML
            Call FuncPathImportFromXML
            RaiseEvent DataChanged(mlng·��ID)
        Case cmd_Edit_Copy                      '����
            Call FuncClipboradCopy
        Case cmd_Edit_Paste                     'ճ��
            Call FuncClipboradPaste
        Case cmd_Edit_Edit                      '����
            If vArea = Area_Step Then
                Call FuncStepEdit
            ElseIf vArea = Area_Item Then
                Call FuncItemEdit(Control)
            End If
        Case cmd_Edit_InsertBefore              'ǰ�����
            If vArea = Area_Category Then
                Call FuncCategoryInsert(-1)
            ElseIf vArea = Area_Step Then
                Call FuncStepInsert(-1)
            ElseIf vArea = Area_Item Then
                Call FuncItemInsert(-1)
            End If
        Case cmd_Edit_InsertAfter               '�������
            If vArea = Area_Category Then
                Call FuncCategoryInsert(1)
            ElseIf vArea = Area_Step Then
                Call FuncStepInsert(1)
            ElseIf vArea = Area_Item Then
                Call FuncItemInsert(1)
            End If
        Case cmd_Edit_InsertBranch              '�����֧
            Call FuncStepBranchInsert
        Case cmd_Edit_Modify                    '�޸�
            vsPath.EditCell
        Case cmd_Edit_Delete                    'ɾ��
            If vArea = Area_Category Then
                Call FuncCategoryDelete
            ElseIf vArea = Area_Step Then
                Call FuncStepDelete
            ElseIf vArea = Area_Item Then
                Call FuncItemDelete
            End If
        Case cmd_Edit_EvalImport                '��������
            Call FuncEvaluateImport
        Case cmd_Edit_EvalStep                  '�׶�����
            Call FuncEvaluateStep
        Case cmd_Edit_EvalStepCopy              '���ƽ׶�����
            Call FuncEvaluateStep(True)
        Case cmd_Edit_Version, cmd_View_Refresh   '�汾,��֧,ˢ��
            If Control.ID = cmd_Edit_Version Then
                Set objCombo = Control
            Else
                Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
            End If

            If objCombo.ListIndex > 0 And mblnChange Then
                If MsgBox("·���������ѱ�������δ����" & IIf(mvClipboard.Empty, "", ",���ҽ���ռ�����") & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    If Control.ID = cmd_Edit_Version Then
                        objCombo.ListIndex = Val(objCombo.Category)
                    End If
                    Exit Sub
                Else
                    mvClipboard.Empty = True
                End If
            End If
            If objCombo.ListIndex = 0 Then
                mblnNewVersion = True
                mblnEditable = False
            Else
                vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                mblnNewVersion = vVersion.�汾�� = 0
                mblnEditable = vVersion.���ʱ�� = Empty
            End If
            objCombo.Category = objCombo.ListIndex
            Call LoadPathTable(objCombo)
            Set objControl = cbsMain.FindControl(, conMenu_View_Show, True)
            If Not objControl Is Nothing Then
                If objControl.Parameter = "����" Then
                    objControl.Parameter = "��ʾ"
                    Call cbsMain_Execute(objControl)
                End If
            End If
            mblnDiff = False
        Case cmd_Edit_VersionInfo       '�汾��Ϣ
            Call FuncVersionEdit
        Case cmd_Edit_VersionNew        '��Ӱ汾
            Call FuncVersionNew
        Case cmd_Edit_VersionDel        '�汾ɾ��
            Call FuncVersionDelete
        Case cmd_View_ToolBar_Button    '������
            Me.cbsMain(2).Visible = Not Me.cbsMain(2).Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text  '��ť����
            For Each objControl In Me.cbsMain(2).Controls
                If objControl.Type <> xtpControlLabel Then
                    If Not Between(objControl.ID, cmd_Edit_Undo, cmd_Edit_Paste) Then
                        objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                    Else
                        objControl.Style = xtpButtonIcon
                    End If
                End If
            Next
            Me.cbsMain.RecalcLayout
        Case cmd_View_ToolBar_Size      '��ͼ��
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StPath        '�鿴��׼·���ο�
            Call frmStPathList.ShowMe(Me, mstr��������, 1)
        Case cmd_View_StatusBar         '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case cmd_View_Find              '����
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus        '��ʱ��Ҫ��λһ��
                If txtFind.Text <> "" Then
                    Call FuncFindItem
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext      '������һ��
            If txtFind.Text = "" Then
                If txtFind.Visible And txtFind.Enabled Then txtFind.SetFocus
            Else
                Call FuncFindItem(True)
            End If
        Case conMenu_View_Difference    '��ʾ����/���ز���
            mblnDiff = Not mblnDiff
            Call ShowContrast(IIf(Control.Caption = "��ʾ����", 1, 2))
        Case conMenu_View_Contrast      '�ԱȲ鿴
            Call CompareAdviceItem
        Case cmd_Edit_ItemShow          '��ʾ��Ŀ�䶯/������Ŀ�䶯
            If Control.Parameter = "��ʾ" Then
                If CheckPathItem Then
                    Control.Parameter = "����"
                    Control.Caption = "������Ŀ�䶯"
                    mbytFunc = 2
                Else
                    Exit Sub
                End If
            Else
                Control.Parameter = "��ʾ"
                Control.Caption = "��ʾ��Ŀ�䶯"
                mbytFunc = 0
            End If
            Call FuncResizeCenter
            Call FuncShowItemAdvice
        Case conMenu_View_Show          '�鿴�䶯��¼
            If Control.Parameter = "��ʾ" Then
                Control.Parameter = "����"
                Control.Caption = "���ر䶯��¼"
                mbytFunc = 1
            Else
                Control.Parameter = "��ʾ"
                Control.Caption = "�鿴�䶯��¼"
                mbytFunc = 0
            End If
            Call FuncSetItemBackColor
            Call FuncResizeCenter
            Call FuncShowItemAdvice
        Case cmd_Help_Web_Home          'Web�ϵ�����
            Call zlHomePage(Me.Hwnd)
        Case cmd_Help_Web_Forum         '������̳
            Call zlWebForum(Me.Hwnd)
        Case cmd_Help_Web_Mail          '���ͷ���
            Call zlMailTo(Me.Hwnd)
        Case cmd_Help_About             '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case cmd_Help_Help              '����
            Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
        Case cmd_File_Exit              '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next

    Me.stcInfo.Left = lngLeft
    Me.stcInfo.Top = lngTop
    Me.stcInfo.Width = lngRight - lngLeft

    picCenter.Move lngLeft, lngTop + Me.stcInfo.Height, lngRight - lngLeft, lngBottom - lngTop - Me.stcInfo.Height
    Call FuncResizeCenter

    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vArea As CONST_AREA, strTemp As String
    Dim blnEnabled As Boolean, blnRefresh As Boolean
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim blnAdjust As Boolean, i As Long

    vArea = GetArea(vsPath.Row, vsPath.Col)
    strTemp = Decode(vArea, Area_Category, "����", Area_Step, "�׶�", Area_Item, "��Ŀ")

    Select Case Control.ID
        Case cmd_File_Save, cmd_File_SaveExit    '����
            Control.Enabled = mblnChange = True
        Case cmd_File_CopyFrom                  '��������
            Control.Enabled = mblnEditable
        Case cmd_File_ExportXML                 '����XML
            If InStr(mstrPrivs, "����XML") = 0 Then
                Control.Visible = False
            Else
                Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
                If Not objCombo Is Nothing Then
                    If objCombo.ListIndex > 0 Then
                        vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                    End If
                End If
                Control.Enabled = vVersion.�汾�� > 0
            End If
        Case cmd_File_ImportXML    '����XML
            If InStr(mstrPrivs, "����XML") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnEditable
            End If
        Case cmd_Edit_Undo    '����
            Control.Visible = False
            Control.Enabled = mblnEditable
        Case cmd_Edit_Redo    '����
            Control.Visible = False
            Control.Enabled = mblnEditable
        Case cmd_Edit_Copy    '����
            Control.Enabled = mblnEditable And vArea = Area_Item And vsPath.Col = vsPath.ColSel
        Case cmd_Edit_Paste    'ճ��
            Control.Enabled = mblnEditable And Not mvClipboard.Empty
        Case cmd_Edit_Caption    '���ܱ���
            If Control.Caption <> strTemp & "��" Then
                Control.Caption = strTemp & "��"
                cbsMain.RecalcLayout
            End If
        Case cmd_Edit_Edit    '����
            If vArea = Area_Category Then
                Control.Visible = False
            Else
                Control.Visible = True
    
                If Control.Parent.Title <> "������" Then
                    If Control.Parent.Controls(Control.Index + 1).BeginGroup <> (vArea = Area_Category) Then
                        Control.Parent.Controls(Control.Index + 1).BeginGroup = (vArea = Area_Category)
                        blnRefresh = True
                    End If
                End If
                If Control.Parent.Title <> "������" Then
                    If Control.Caption <> "����" & strTemp & "(&E)" Then
                        Control.Caption = "����" & strTemp & "(&E)"
                        blnRefresh = True
                    End If
                End If
                If blnRefresh Then cbsMain.RecalcLayout
    
                If vArea = Area_Step Then
                    Control.Enabled = mblnEditable And vsPath.ColSel = vsPath.Col
                ElseIf vArea = Area_Item Then
                    blnEnabled = (vsPath.ColSel = vsPath.Col) And (vsPath.RowSel = vsPath.Row)
    
                    '�ж�����΢����������δͣ�õ�����˰汾�������ҽ����������΢��
                    If blnEnabled And Not mblnEditable And mbytMode = Mode_Design Then
                        Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
                        If Not objCombo Is Nothing Then
                            If objCombo.ListIndex > 0 Then
                                vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                            End If
                        End If
                        If vVersion.�汾�� > 0 And vVersion.���ʱ�� <> Empty And vVersion.ͣ��ʱ�� = Empty Then
                            If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) <> "Empty" Then
                                vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
                                If vItem.ҽ��IDs <> "" Or vItem.����IDs <> "" Or vItem.�°没��IDs <> "" Then
                                    blnAdjust = True
                                End If
                            End If
                        End If
                    End If
    
                    Control.Enabled = blnEnabled And (mblnEditable Or blnAdjust) And mbytFunc = 0
    
                    If Control.Enabled And blnAdjust Then
                        Control.Parameter = "Adjust"
                    Else
                        Control.Parameter = ""
                    End If
                End If
            End If
        Case cmd_Edit_Insert    '����
            If Control.Parent.Title <> "������" Then
                If Control.Caption <> "����" & strTemp & "(&I)" Then
                    Control.Caption = "����" & strTemp & "(&I)"
                    cbsMain.RecalcLayout
                End If
            End If
            Control.Enabled = mblnEditable
        Case cmd_Edit_InsertBefore    '��ǰ�����
            If vArea = Area_Category Then
                Control.Enabled = mblnEditable And (vsPath.RowSel = vsPath.Row)
            ElseIf vArea = Area_Step Then
                Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col)
            ElseIf vArea = Area_Item Then
                Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col) And (vsPath.RowSel = vsPath.Row)
            End If
        Case cmd_Edit_InsertAfter    '�ں������
            If vArea = Area_Category Then
                Control.Enabled = mblnEditable And (vsPath.RowSel = vsPath.Row)
            ElseIf vArea = Area_Step Then
                Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col)
            ElseIf vArea = Area_Item Then
                Control.Enabled = mblnEditable And (vsPath.ColSel = vsPath.Col) And (vsPath.RowSel = vsPath.Row)
            End If
        Case cmd_Edit_InsertBranch    '�����֧
            If vArea = Area_Step Then
                Control.Visible = True
    
                blnEnabled = vsPath.ColSel = vsPath.Col
                If blnEnabled Then
                    '�����˵�ʱ��׶β��ܲ����֧
                    blnEnabled = TypeName(vsPath.ColData(vsPath.Col)) <> "Empty"
                End If
                Control.Enabled = mblnEditable And blnEnabled
            Else
                Control.Visible = False
            End If
        Case cmd_Edit_Modify            '�޸�
            If strTemp = "����" Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
            Control.Enabled = mblnEditable
        Case cmd_Edit_Delete    'ɾ��
            If Control.Parent.Title <> "������" Then
                If Control.Caption <> "ɾ��" & strTemp & "(&D)" Then
                    Control.Caption = "ɾ��" & strTemp & "(&D)"
                    cbsMain.RecalcLayout
                End If
            End If
            Control.Enabled = mblnEditable
        Case cmd_Edit_EvalImport    '��������
            If InStr(mstrPrivs, "���������") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnEditable
            End If
        Case cmd_Edit_EvalStep    '�׶�����
            If InStr(mstrPrivs, "���������") = 0 Then
                Control.Visible = False
            Else
                blnEnabled = mblnEditable And vsPath.Col >= vsPath.FixedCols + vsPath.FrozenCols And vsPath.Cols > 0
                If blnEnabled Then
                    With vsPath
                        If TypeName(.ColData(.Col)) = "Empty" Then
                            blnEnabled = False
                        End If
                    End With
                End If
                Control.Enabled = blnEnabled
            End If
        Case cmd_Edit_EvalStepCopy    '���ƽ׶�����
            If InStr(mstrPrivs, "���������") = 0 Then
                Control.Visible = False
            Else
                blnEnabled = mblnEditable And vsPath.Col >= vsPath.FixedCols + vsPath.FrozenCols And vsPath.Cols > 0
                If blnEnabled Then
                    With vsPath
                        If TypeName(.ColData(.Col)) <> "Empty" Then
                            vStep = .ColData(.Col)
                            If Not vStep.����.������ Is Nothing And Not vStep.����.ָ�꼯 Is Nothing Then
                                If vStep.����.������.count > 0 Or vStep.����.ָ�꼯.count > 0 Then
                                    blnEnabled = False
                                End If
                            End If
                        Else
                            blnEnabled = False
                        End If
                    End With
                End If
                Control.Enabled = blnEnabled
            End If
        Case cmd_Edit_VersionInfo    '�汾��Ϣ
            Control.Enabled = mblnEditable
        Case cmd_Edit_VersionNew    '��Ӱ汾
            'û��δ��˰汾ʱ(�ѱ��������δ�����)����������µİ汾
            Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
            blnEnabled = Not objCombo Is Nothing
            If blnEnabled Then
                For i = 1 To objCombo.ListCount
                    vVersion = mcolVersion("_" & objCombo.ItemData(i))
                    If vVersion.���ʱ�� = Empty Then Exit For
                Next
                If i <= objCombo.ListCount Then blnEnabled = False
            End If
            Control.Enabled = blnEnabled
        Case cmd_Edit_VersionDel    'ɾ���汾
            Control.Enabled = mblnEditable
        Case cmd_View_ToolBar_Button    '������
            If cbsMain.count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text    'ͼ������
            If cbsMain.count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case cmd_View_ToolBar_Size    '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case cmd_View_StatusBar    '״̬��
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Difference, conMenu_View_Contrast   '��ʾ����/���ز��� '�ԱȲ鿴
            Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
            If Not objCombo Is Nothing Then
                If objCombo.ListIndex > 0 Then
                    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                End If
            End If
            If vVersion.�汾�� > 1 And vVersion.���ʱ�� = Empty Then
                If Control.ID = conMenu_View_Difference Then
                    Control.Enabled = True
                    Control.Caption = IIf(mblnDiff, "���ز���", "��ʾ����")
                End If
                If Control.ID = conMenu_View_Contrast Then
                    Control.Enabled = IIf(cbsMain.FindControl(, conMenu_View_Difference, True, True).Caption = "���ز���", True, False)
                End If
            Else
                Control.Enabled = False
            End If
        Case conMenu_View_Show  '�鿴�䶯��¼
            Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
            If Not objCombo Is Nothing Then
                If objCombo.ListIndex > 0 Then
                    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                    If vVersion.���ʱ�� = Empty Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = True And (mbytFunc <> 2)
                    End If
                End If
            End If
        Case cmd_Edit_ItemShow   '��ʾ��Ŀ�䶯
            If InStr(mstrPrivs, "·��ҽ������") = 0 Or mblnEditable Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
            If Control.Visible Then
                Control.Enabled = (mbytFunc <> 1)
            End If
    End Select
End Sub

Private Sub cmdCheck_Click(Index As Integer)
    Dim vItem As TYPE_PATH_ITEM
    Dim rsTmp As ADODB.Recordset
    Dim arrtmp As Variant
    Dim strDate As String
    Dim strSql As String
    Dim strTmp As String
    Dim i As Long

    On Error GoTo errH
    If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) = "Empty" Then Exit Sub
    vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
    strDate = "To_Date('" & cboTimes.Tag & "','YYYY-MM-DD HH24:MI:SS')"

    If Index = 0 Then
        If MsgBox("��ȷ����Ŀ""" & vItem.��Ŀ���� & """��ҽ�����ݡ����ͨ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            strSql = "Zl_����·��ҽ���䶯_Audit(" & vItem.ID & "," & strDate & ")"
        Else
            Exit Sub
        End If
    Else
        If MsgBox("��ȷ����Ŀ""" & vItem.��Ŀ���� & """��ҽ�����ݡ���˲�ͨ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            strSql = "Zl_����·��ҽ���䶯_Audit(" & vItem.ID & "," & strDate & ",1)"
        Else
            Exit Sub
        End If
    End If

    '�����ύ
    If strSql <> "" Then
        Call zlDatabase.ExecuteProcedure(strSql, "����·��ҽ�����")
    End If

    If Index = 0 Then
        strSql = "Select a.Id, a.���id, a.���, a.��Ч, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.ҽ������, a.ִ��Ƶ��, a.Ƶ�ʴ���," & vbNewLine & _
                "       a.Ƶ�ʼ��, a.�����λ, a.ִ������, a.ִ�б��, a.ִ�п���id, a.ʱ�䷽��, a.�Ƿ�ȱʡ, a.�Ƿ�ѡ, a.�䷽id, a.�����Ŀid" & vbNewLine & _
                "From ����·��ҽ������ A, ����·��ҽ�� B" & vbNewLine & _
                "Where a.Id = b.ҽ������id And b.·����Ŀid = [1] Order By a.���"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vItem.ID)
        strTmp = ""
        Do While Not rsTmp.EOF
            strTmp = strTmp & "," & rsTmp!ID
            mrsAdvice.AddNew

            mrsAdvice!ID = rsTmp!ID
            mrsAdvice!���id = rsTmp!���id
            mrsAdvice!�Ƿ�ȱʡ = Val(rsTmp!�Ƿ�ȱʡ & "")
            mrsAdvice!�Ƿ�ѡ = Val(rsTmp!�Ƿ�ѡ & "")
            mrsAdvice!��� = rsTmp!���
            mrsAdvice!��Ч = rsTmp!��Ч
            mrsAdvice!������ĿID = rsTmp!������ĿID
            mrsAdvice!�շ�ϸĿID = rsTmp!�շ�ϸĿID
            mrsAdvice!ҽ������ = rsTmp!ҽ������
            mrsAdvice!�������� = rsTmp!��������
            mrsAdvice!�ܸ����� = rsTmp!�ܸ�����
            mrsAdvice!�걾��λ = rsTmp!�걾��λ
            mrsAdvice!��鷽�� = rsTmp!��鷽��
            mrsAdvice!ҽ������ = rsTmp!ҽ������
            mrsAdvice!ִ��Ƶ�� = rsTmp!ִ��Ƶ��
            mrsAdvice!Ƶ�ʴ��� = rsTmp!Ƶ�ʴ���
            mrsAdvice!Ƶ�ʼ�� = rsTmp!Ƶ�ʼ��
            mrsAdvice!�����λ = rsTmp!�����λ
            mrsAdvice!ִ������ = rsTmp!ִ������
            mrsAdvice!ִ�п���ID = rsTmp!ִ�п���ID
            mrsAdvice!ʱ�䷽�� = rsTmp!ʱ�䷽��
            mrsAdvice!�䷽ID = rsTmp!�䷽ID
            mrsAdvice!�����ĿID = rsTmp!�����ĿID
            mrsAdvice!ִ�б�� = rsTmp!ִ�б��

            mrsAdvice.Update
            rsTmp.MoveNext
        Loop

        '��ջ���
        arrtmp = Split(vItem.ҽ��IDs, ",")
        For i = LBound(arrtmp) To UBound(arrtmp)
            mrsAdvice.Filter = "ID =" & arrtmp(i)
            If mrsAdvice.RecordCount > 0 Then
                mrsAdvice.Delete
                mrsAdvice.Update
            End If
        Next
        mrsAdvice.Filter = ""
        vItem.ҽ��IDs = Mid(strTmp, 2)
    End If

    strSql = "Select a.��Ŀid, a.ҽ������id From ����·��ҽ���䶯 A" & vbNewLine & _
                "Where a.��Ŀid = [1] And a.���ʱ�� Is Null And" & vbNewLine & _
                "      a.����ʱ�� = (Select Max(����ʱ��) From ����·��ҽ���䶯 C Where c.��Ŀid = [1] And c.���ʱ�� Is Null)" & vbNewLine & _
                "Order By a.��Ŀid, a.ҽ������id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vItem.ID)
    strTmp = ""
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!ҽ������ID
        rsTmp.MoveNext
    Loop
    vItem.�����ҽ��IDs = Mid(strTmp, 2)
    vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col) = vItem
    If vItem.�����ҽ��IDs = "" Then
        vsPath.Cell(flexcpForeColor, vsPath.Row, vsPath.Col) = vbBlack
    End If
    '������Ŀˢ��
    Call vsPath_AfterRowColChange(vsPath.Row, vsPath.Col, vsPath.Row, vsPath.Col)

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim objPane As Pane

    If mbytMode = Mode_Show Then
        Call zlControl.FormSetCaption(Me, False, False)
        vsPath.Editable = flexEDNone
        vsPath.AllowSelection = False
        vsPath.HighLight = flexHighlightWithFocus
        vsPath.FocusRect = flexFocusLight
        Me.stbThis.Visible = False
    End If

    vsPath.Editable = flexEDNone
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True     '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.Icons = imgMain.Icons

    If mbytMode = Mode_Design Then
        Call RestoreWinState(Me, App.ProductName)
    End If
    Call MainDefCommandBar

    If mbytMode = Mode_Design Then
        Set mfrmVersion = New frmVersionOut
        Set mfrmPathStep = New frmPathStepEditOut
        Set mfrmPathItem = New frmPathItemEditOut
        Set mfrmEvalEdit = New frmEvaluateEdit
        Set mfrmAdviceContrast = New frmAdviceContrast
        Me.WindowState = vbMaximized                    '����Ĭ�����
    End If

    '��ȡ����
    If mbytMode = Mode_Design Then
        vsPath.ExplorerBar = flexExSort
        Call LoadPathVersion
    Else
        vsPath.ExplorerBar = flexExNone
        mblnEditable = False
    End If

    mblnChange = False
    mvClipboard.Empty = True
    Erase mvClipboard.��Ŀ��
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytMode = Mode_Design And mblnChange Then
        If MsgBox("·���������ѱ�������δ���棬ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        mblnChange = False
    End If

    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
        Set mrsAdvice = Nothing
    End If

    mvClipboard.Empty = True
    Erase mvClipboard.��Ŀ��

    If mbytMode = Mode_Design Then
        Unload mfrmVersion
        Set mfrmVersion = Nothing

        Unload mfrmPathStep
        Set mfrmPathStep = Nothing

        Unload mfrmPathItem
        Set mfrmPathItem = Nothing

        Unload mfrmEvalEdit
        Set mfrmEvalEdit = Nothing

        Unload mfrmAdviceContrast
        Set mfrmAdviceContrast = Nothing
    End If

    If mbytMode = Mode_Design Then
        Call SaveWinState(Me, App.ProductName)
    End If
End Sub

Private Function LoadPathVersion(Optional ByVal intVersion As Integer = -1) As Boolean
'���ܣ���ȡ��������ʾ�����ٴ�·���İ汾�б�
'������intVersion=ȱʡ��λ�汾
    Dim vVersion As TYPE_PATH_VERSION
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim objCombo As CommandBarComboBox

    On Error GoTo errH

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Function
    objCombo.Clear: vsPath.Rows = 0: vsPath.Cols = 0
    If mlng·��ID = 0 Then Exit Function

    Set mcolVersion = New Collection

    strSql = " Select A.����,A.����,B.�汾��,B.��׼����ʱ��,B.��׼����,B.�汾˵��,B.������,B.����ʱ��,B.�����,B.���ʱ��,B.ͣ����,B.ͣ��ʱ��" & _
             " From ����·��Ŀ¼ A,����·���汾 B Where A.ID=B.·��ID(+) And A.ID=[1] Order by B.�汾�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ����·���汾", mlng·��ID)

    If rsTmp.RecordCount <= 0 Then Exit Function

    Me.Tag = rsTmp!���� & "-" & rsTmp!����
    If mbytMode = Mode_Design Then
        Me.Caption = "����·����� - " & rsTmp!����
    End If

    Do While Not rsTmp.EOF
        If Not IsNull(rsTmp!�汾��) Then
            objCombo.AddItem "�� " & rsTmp!�汾�� & " �棬" & _
                "������" & rsTmp!������ & "/" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm") & _
                IIf(Not IsNull(rsTmp!���ʱ��), "����ˣ�" & rsTmp!����� & "/" & Format(rsTmp!���ʱ��, "yyyy-MM-DD HH:mm"), "") & _
                IIf(Not IsNull(rsTmp!ͣ��ʱ��), "��ͣ�ã�" & rsTmp!ͣ���� & "/" & Format(rsTmp!ͣ��ʱ��, "yyyy-MM-dd HH:mm"), "")
            objCombo.ItemData(objCombo.ListCount) = rsTmp!�汾��
            If rsTmp!�汾�� = intVersion Then
                objCombo.ListIndex = objCombo.ListCount
            End If

            vVersion.�汾�� = rsTmp!�汾��
            vVersion.��׼����ʱ�� = NVL(rsTmp!��׼����ʱ��)
            vVersion.��׼���� = NVL(rsTmp!��׼����)
            vVersion.�汾˵�� = NVL(rsTmp!�汾˵��)
            vVersion.������ = rsTmp!������
            vVersion.����ʱ�� = rsTmp!����ʱ��
            vVersion.����� = NVL(rsTmp!�����)
            vVersion.���ʱ�� = IIf(IsNull(rsTmp!���ʱ��), Empty, rsTmp!���ʱ��)
            vVersion.ͣ���� = NVL(rsTmp!ͣ����)
            vVersion.ͣ��ʱ�� = IIf(IsNull(rsTmp!ͣ��ʱ��), Empty, rsTmp!ͣ��ʱ��)
            mcolVersion.Add vVersion, "_" & vVersion.�汾��
        End If
        rsTmp.MoveNext
    Loop

    '������1��ʼ��ֱ�Ӹ�ֵ��������Execute�¼�
    If objCombo.ListCount = 0 Then
        If mbytMode = Mode_Show Then
            cbsMain.RecalcLayout: Exit Function
        End If
        objCombo.AddItem "��������С���"

        vVersion.�汾�� = 0
        vVersion.��׼����ʱ�� = ""
        vVersion.��׼���� = ""
        vVersion.�汾˵�� = ""
        vVersion.������ = ""
        vVersion.����ʱ�� = Empty
        vVersion.����� = ""
        vVersion.���ʱ�� = Empty
        vVersion.ͣ���� = ""
        vVersion.ͣ��ʱ�� = Empty
        mcolVersion.Add vVersion, "_0"
    End If
    If objCombo.ListIndex = 0 Then objCombo.ListIndex = 1
    cbsMain.RecalcLayout

    Call cbsMain_Execute(objCombo)
    LoadPathVersion = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPathTable(objCombo As CommandBarComboBox) As Boolean
'���ܣ�������ѡ���·���汾������·������Ӧ�����ݽ�����ʾ
    Dim vVersion As TYPE_PATH_VERSION
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond

    Dim colCols As New Collection
    Dim colRows As New Collection

    Dim rsTmp As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim rsPathAdvice As ADODB.Recordset
    Dim rsPathEPR As ADODB.Recordset
    Dim rsEvalMark As ADODB.Recordset
    Dim rsEvalCond As ADODB.Recordset
    Dim strSql As String, strItems As String
    Dim i As Long
    Dim lngRow As Long, lngCol As Long
    Dim blnBranch As Boolean

    On Error GoTo errH

    vsPath.Redraw = flexRDNone
    vsPath.Rows = 0
    vsPath.Cols = 0
    
    If objCombo.ListIndex = 0 Then
        vsPath.Redraw = flexRDDirect
        Exit Function
    End If

    '�汾��Ϣ��ʾ
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    stcInfo.Caption = "��׼����ʱ�䣺" & IIf(vVersion.��׼����ʱ�� <> "", vVersion.��׼����ʱ�� & "��", "<δ�趨>") & _
                      "����׼���ã�" & IIf(vVersion.��׼���� <> "", vVersion.��׼���� & "Ԫ", "<δ�趨>") & _
                      "��˵����" & IIf(vVersion.�汾˵�� <> "", vVersion.�汾˵��, "<��>")
    '·������ɫ����
    If vVersion.ͣ��ʱ�� <> Empty Then
        vsPath.GridColor = Color_StopLine
        vsPath.SheetBorder = Color_StopLine
        vsPath.BackColorFrozen = Color_StopBack
    ElseIf vVersion.���ʱ�� <> Empty Then
        vsPath.GridColor = Color_AuditLine
        vsPath.SheetBorder = Color_AuditLine
        vsPath.BackColorFrozen = Color_AuditBack
    Else
        vsPath.GridColor = Color_NewLine
        vsPath.SheetBorder = Color_NewLine
        vsPath.BackColorFrozen = Color_NewBack
    End If

    '��ʼ����ǰ�汾ҽ�����ݱ�
    Call InitAdviceRecordset
    Set mvEvalImport.ָ�꼯 = New Collection
    Set mvEvalImport.������ = New Collection

    If vVersion.�汾�� = 0 Then
        '�յ�·����ȱʡ��ʽ
        With vsPath
            .Rows = 2 + 1: .FixedRows = 1: .FrozenRows = 1
            .Cols = 1 + 1: .FixedCols = 0: .FrozenCols = 1
            .ColWidth(-1) = COl_WIDTH_BASE: .ColWidth(0) = 1000
        End With
    Else
        '�ѱ����·������ʽ
        With vsPath
            .Rows = 3: .FixedRows = 1: .FrozenRows = 2
            .Cols = 1: .FixedCols = 0: .FrozenCols = 1

            '�������ݶ�ȡ
            strSql = " Select A.��������,A.�׶�ID,B.ID,B.���,B.����ָ��,B.ָ������,B.ָ���� From ����·������ A,����·������ָ�� B " & _
                     " Where A.ID=B.����ID And A.·��ID=[1] And A.�汾��=[2] " & _
                     " Order by A.��������,A.�׶�ID,B.���"
            Set rsEvalMark = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))

            strSql = " Select A.��������,A.�׶�ID,B.ָ��ID,B.��ĿID,B.��ϵʽ,B.����ֵ,B.������� From ����·������ A,����·���������� B " & _
                     " Where A.ID=B.����ID And A.·��ID=[1] And A.�汾��=[2] " & _
                     " Order by A.��������,A.�׶�ID"
            Set rsEvalCond = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))

            '0)��������
            rsEvalMark.Filter = "��������=1"
            Do While Not rsEvalMark.EOF
                vEvalMark.ID = rsEvalMark!ID
                vEvalMark.��� = rsEvalMark!���
                vEvalMark.����ָ�� = rsEvalMark!����ָ��
                vEvalMark.ָ������ = rsEvalMark!ָ������
                vEvalMark.ָ���� = rsEvalMark!ָ����
                mvEvalImport.ָ�꼯.Add vEvalMark
                rsEvalMark.MoveNext
            Loop
            rsEvalCond.Filter = "��������=1"
            Do While Not rsEvalCond.EOF
                vEvalCond.ָ��ID = NVL(rsEvalCond!ָ��ID, 0)
                vEvalCond.��ĿID = NVL(rsEvalCond!��ĿID, 0)
                vEvalCond.��ϵʽ = rsEvalCond!��ϵʽ
                vEvalCond.����ֵ = rsEvalCond!����ֵ
                vEvalCond.������� = rsEvalCond!�������
                mvEvalImport.������.Add vEvalCond
                rsEvalCond.MoveNext
            Loop

            '1)ʱ��׶β���
            strSql = " Select Distinct A.ID,Nvl(A.��ID,0) as ��ID,A.���,b.��� as ��ID���,A.����,A.��ʼ����,A.��������,A.����,A.˵��" & _
                     " From ����·���׶� A,����·���׶� B " & _
                     " Where a.��ID=b.ID(+) And A.·��ID=[1] And A.�汾��=[2]" & _
                     " Order by NVL(B.���,A.���),NVL(b.���,0),NVL(a.���,0)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))

            blnBranch = False
            Set rsClone = rsTmp.Clone
            rsTmp.Filter = "��ID=0"
            
            Do While Not rsTmp.EOF
                .Cols = .Cols + 1

                vStep.ID = rsTmp!ID
                vStep.��ID = 0
                vStep.��� = rsTmp!���
                vStep.���� = rsTmp!����
                vStep.��ʼ���� = NVL(rsTmp!��ʼ����, 0)
                vStep.�������� = NVL(rsTmp!��������, 0)
                vStep.���� = NVL(rsTmp!����)
                vStep.˵�� = NVL(rsTmp!˵��)

                '�׶�����
                Set vStep.����.ָ�꼯 = New Collection
                rsEvalMark.Filter = "��������=2 And �׶�ID=" & vStep.ID
                Do While Not rsEvalMark.EOF
                    vEvalMark.ID = rsEvalMark!ID
                    vEvalMark.��� = rsEvalMark!���
                    vEvalMark.����ָ�� = rsEvalMark!����ָ��
                    vEvalMark.ָ������ = rsEvalMark!ָ������
                    vEvalMark.ָ���� = rsEvalMark!ָ����
                    vStep.����.ָ�꼯.Add vEvalMark
                    rsEvalMark.MoveNext
                Loop
                Set vStep.����.������ = New Collection
                rsEvalCond.Filter = "��������=2 And �׶�ID=" & vStep.ID
                Do While Not rsEvalCond.EOF
                    vEvalCond.ָ��ID = NVL(rsEvalCond!ָ��ID, 0)
                    vEvalCond.��ĿID = NVL(rsEvalCond!��ĿID, 0)
                    vEvalCond.��ϵʽ = rsEvalCond!��ϵʽ
                    vEvalCond.����ֵ = rsEvalCond!����ֵ
                    vEvalCond.������� = rsEvalCond!�������
                    vStep.����.������.Add vEvalCond
                    rsEvalCond.MoveNext
                Loop

                .ColData(.Cols - 1) = vStep
                '.Cell(flexcpText, .FixedRows, .Cols - 1, .FixedRows + .FrozenRows - 1, .Cols - 1) = vStep.����
                '���ֱ�ӷ�Χ��ֵ����Ϊ�����س����Զ�ʶ��Ϊ�ָ��������������ֱ��ж�
                .TextMatrix(.FixedRows, .Cols - 1) = vStep.����
                .TextMatrix(.FixedRows + .FrozenRows - 1, .Cols - 1) = vStep.����

                If mbytMode = Mode_Design Then
                    .TextMatrix(.FixedRows - 1, .Cols - 1) = "�׶�������"
                    .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = False
                    If vStep.����.ָ�꼯.count > 0 Or vStep.����.������.count > 0 Then
                        .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = True
                    End If
                End If

                '���ڿ��ٶ�λ�ý׶ε��к�
                colCols.Add .Cols - 1, "_" & vStep.ID

                '���뱸ѡ��֧
                rsClone.Filter = "��ID=" & rsTmp!ID
                If rsClone.EOF Then
                    If vStep.���ڷ�֧ Then
                        .Cell(flexcpPicture, .FixedRows, .Cols - 1) = ImgBranch.Picture
                        .Cell(flexcpPictureAlignment, .FixedRows, .Cols - 1) = 1
                    End If
                Else
                    Do While Not rsClone.EOF
                        .Cols = .Cols + 1

                        vStep.ID = rsClone!ID
                        vStep.��ID = rsClone!��ID
                        vStep.��� = rsClone!���
                        vStep.���� = NVL(rsClone!����)
                        vStep.˵�� = NVL(rsClone!˵��)
                        '����Ӧ���ʱ��׶���ͬ
                        vStep.���� = rsClone!����
                        vStep.��ʼ���� = NVL(rsClone!��ʼ����, 0)
                        vStep.�������� = NVL(rsClone!��������, 0)

                        '�׶�����
                        Set vStep.����.ָ�꼯 = New Collection
                        rsEvalMark.Filter = "��������=2 And �׶�ID=" & vStep.ID
                        Do While Not rsEvalMark.EOF
                            vEvalMark.ID = rsEvalMark!ID
                            vEvalMark.��� = rsEvalMark!���
                            vEvalMark.����ָ�� = rsEvalMark!����ָ��
                            vEvalMark.ָ������ = rsEvalMark!ָ������
                            vEvalMark.ָ���� = rsEvalMark!ָ����
                            vStep.����.ָ�꼯.Add vEvalMark
                            rsEvalMark.MoveNext
                        Loop
                        Set vStep.����.������ = New Collection
                        rsEvalCond.Filter = "��������=2 And �׶�ID=" & vStep.ID
                        Do While Not rsEvalCond.EOF
                            vEvalCond.ָ��ID = NVL(rsEvalCond!ָ��ID, 0)
                            vEvalCond.��ĿID = NVL(rsEvalCond!��ĿID, 0)
                            vEvalCond.��ϵʽ = rsEvalCond!��ϵʽ
                            vEvalCond.����ֵ = rsEvalCond!����ֵ
                            vEvalCond.������� = rsEvalCond!�������
                            vStep.����.������.Add vEvalCond
                            rsEvalCond.MoveNext
                        Loop

                        .ColData(.Cols - 1) = vStep
                        .TextMatrix(.FixedRows, .Cols - 1) = vStep.����
                        .TextMatrix(.FixedRows + .FrozenRows - 1, .Cols - 1) = IIf(vStep.˵�� = "", "���÷�֧" & vStep.���, vStep.˵��) & IIf(vStep.���� = "", "", ",") & vStep.����
                        If vStep.��� = 1 Then
                            .TextMatrix(.FixedRows + .FrozenRows - 1, .Cols - 2) = "ȱʡ��֧"
                        End If

                        If vStep.���ڷ�֧ Then
                            .Cell(flexcpPicture, .FixedRows + .FrozenRows - 1, .Cols - 2) = ImgBranch.Picture
                            .Cell(flexcpPictureAlignment, .FixedRows + .FrozenRows - 1, .Cols - 2) = 3
                        End If

                        If mbytMode = Mode_Design Then
                            .TextMatrix(.FixedRows - 1, .Cols - 1) = "�׶�������"
                            .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = False
                            If vStep.����.ָ�꼯.count > 0 And vStep.����.������.count > 0 Then
                                .Cell(flexcpFontBold, .FixedRows - 1, .Cols - 1) = True
                            End If
                        End If

                        '���ڿ��ٶ�λ�ý׶ε��к�
                        colCols.Add .Cols - 1, "_" & vStep.ID

                        blnBranch = True
                        rsClone.MoveNext
                    Loop
                End If

                rsTmp.MoveNext
            Loop
            If Not blnBranch Then
                .FrozenRows = 1
                .RemoveItem .FixedRows + .FrozenRows
            End If

            '2)���ಿ��
            strSql = " Select A.���,A.����,Max(����) As ���� From (" & _
                     "   Select A.���,A.���� as ����,Nvl(B.�׶�ID,0),Count(Nvl(B.��Ŀ���,0)) as ���� " & _
                     "   From ����·������ A,����·����Ŀ B " & _
                     "   Where A.·��ID=[1] And A.�汾��=[2] " & _
                     "   And A.����=B.����(+) And B.·��ID(+)=[1] And B.�汾��(+)=[2] " & _
                     "   Group by A.���,A.����,Nvl(B.�׶�ID,0)) A " & _
                     " Group By A.���,A.���� Order By A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))

            Do While Not rsTmp.EOF
                '���ֻ�������򣬱���ʱ��������
                .Rows = .Rows + rsTmp!����
                .Cell(flexcpText, .Rows - rsTmp!����, .FixedCols, .Rows - 1, .FixedCols) = rsTmp!����

                '���ڿ��ٶ�λ�÷������ʼ�к�
                colRows.Add .Rows - rsTmp!����, "_" & rsTmp!����

                rsTmp.MoveNext
            Loop

            '3)��Ŀ����
            '--ҽ���������ݼ�
            strSql = " Select Distinct A.ID,A.���ID,A.���,A.��Ч,A.������ĿID,A.�շ�ϸĿID," & _
                     " A.ҽ������,A.��������,A.�ܸ�����,A.�걾��λ,A.��鷽��,A.ҽ������," & _
                     " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ������,A.ִ�б��,A.ִ�п���ID,A.ʱ�䷽��,A.�Ƿ�ȱʡ,A.�Ƿ�ѡ,A.�䷽ID,A.�����ĿID" & _
                     " From ����·��ҽ������ A,����·��ҽ�� B,����·����Ŀ C" & _
                     " Where A.ID=B.ҽ������ID And B.·����ĿID=C.ID And C.·��ID=[1] And C.�汾��=[2]" & _
                     " Order by A.���,A.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))
            Do While Not rsTmp.EOF
                mrsAdvice.AddNew
                mrsAdvice!ID = rsTmp!ID
                mrsAdvice!���id = rsTmp!���id
                mrsAdvice!�Ƿ�ȱʡ = Val(rsTmp!�Ƿ�ȱʡ & "")
                mrsAdvice!�Ƿ�ѡ = Val(rsTmp!�Ƿ�ѡ & "")
                mrsAdvice!��� = rsTmp!���
                mrsAdvice!��Ч = rsTmp!��Ч
                mrsAdvice!������ĿID = rsTmp!������ĿID
                mrsAdvice!�շ�ϸĿID = rsTmp!�շ�ϸĿID
                mrsAdvice!ҽ������ = rsTmp!ҽ������
                mrsAdvice!�������� = rsTmp!��������
                mrsAdvice!�ܸ����� = rsTmp!�ܸ�����
                mrsAdvice!�걾��λ = rsTmp!�걾��λ
                mrsAdvice!��鷽�� = rsTmp!��鷽��
                mrsAdvice!ҽ������ = rsTmp!ҽ������
                mrsAdvice!ִ��Ƶ�� = rsTmp!ִ��Ƶ��
                mrsAdvice!Ƶ�ʴ��� = rsTmp!Ƶ�ʴ���
                mrsAdvice!Ƶ�ʼ�� = rsTmp!Ƶ�ʼ��
                mrsAdvice!�����λ = rsTmp!�����λ
                mrsAdvice!ִ������ = rsTmp!ִ������
                mrsAdvice!ִ�п���ID = rsTmp!ִ�п���ID
                mrsAdvice!ʱ�䷽�� = rsTmp!ʱ�䷽��
                mrsAdvice!�䷽ID = rsTmp!�䷽ID
                mrsAdvice!�����ĿID = rsTmp!�����ĿID
                mrsAdvice!ִ�б�� = rsTmp!ִ�б��

                mrsAdvice.Update
                rsTmp.MoveNext
            Loop
            '--ҽ����Ӧ��ϵ
            strSql = " Select Distinct A.·����ĿID,A.ҽ������ID" & _
                     " From ����·��ҽ�� A,����·����Ŀ B Where A.·����ĿID=B.ID And B.·��ID=[1] And B.�汾��=[2]" & _
                     " Order by ·����ĿID,ҽ������ID"
            Set rsPathAdvice = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))
            '--������Ӧ��ϵ
            strSql = " Select Distinct A.��ĿID,A.�ļ�ID,A.ԭ��ID " & _
                     " From ����·������ A,����·����Ŀ B Where A.��ĿID=B.ID And B.·��ID=[1] And B.�汾��=[2] " & _
                     " Order by ��ĿID,�ļ�ID,ԭ��ID"
            Set rsPathEPR = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))
            '--·����Ŀ
            Set mcolItemRowCol = New Collection
            strSql = " Select a.ID,a.�׶�ID,a.����,a.��Ŀ���,a.��Ŀ����,a.ִ�з�ʽ,a.��Ŀ���,a.ͼ��ID,a.����Ҫ��,A.����ο�,Nvl(A.������,1) ������" & _
                     " From ����·����Ŀ A,����·���׶� B,����·���׶� C Where a.�׶�ID=b.ID And b.��ID=c.ID(+) And a.·��ID=[1] And a.�汾��=[2] " & _
                     " Order by NVL(c.���,b.���),NVL(c.���,0),a.����,a.��Ŀ���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", mlng·��ID, objCombo.ItemData(objCombo.ListIndex))
            Do While Not rsTmp.EOF
                vItem.ID = rsTmp!ID
                vItem.��Ŀ��� = rsTmp!��Ŀ���
                vItem.��Ŀ���� = rsTmp!��Ŀ����
                vItem.ִ�з�ʽ = NVL(rsTmp!ִ�з�ʽ, 0)
                vItem.��Ŀ��� = NVL(rsTmp!��Ŀ���)
                vItem.ͼ��ID = NVL(rsTmp!ͼ��ID, 0)
                vItem.����Ҫ�� = Val("" & rsTmp!����Ҫ��)
                vItem.����ο� = NVL(rsTmp!����ο�)
                vItem.������ = NVL(rsTmp!������)
                '������ҽ��
                rsPathAdvice.Filter = "·����ĿID=" & rsTmp!ID
                vItem.ҽ��IDs = ""
                Do While Not rsPathAdvice.EOF
                    vItem.ҽ��IDs = vItem.ҽ��IDs & "," & rsPathAdvice!ҽ������ID
                    rsPathAdvice.MoveNext
                Loop
                vItem.ҽ��IDs = Mid(vItem.ҽ��IDs, 2)
                If vVersion.���ʱ�� <> Empty And vVersion.ͣ��ʱ�� = Empty Then
                    vItem.ԭҽ��IDs = vItem.ҽ��IDs
                    If InStr(mstrPrivs, "·��ҽ������") > 0 Then strItems = strItems & "," & vItem.ID
                End If

                '�����Ĳ���
                rsPathEPR.Filter = "��ĿID=" & rsTmp!ID
                vItem.����IDs = "": vItem.�°没��IDs = ""
                Do While Not rsPathEPR.EOF
                    If rsPathEPR!�ļ�ID & "" <> "" Then
                        vItem.����IDs = vItem.����IDs & "," & rsPathEPR!�ļ�ID
                    Else
                        vItem.�°没��IDs = vItem.�°没��IDs & "," & rsPathEPR!ԭ��ID
                    End If
                    rsPathEPR.MoveNext
                Loop
                vItem.����IDs = Mid(vItem.����IDs, 2)
                vItem.�°没��IDs = Mid(vItem.�°没��IDs, 2)
                '��λ����ʾ
                lngCol = colCols("_" & rsTmp!�׶�id)
                lngRow = colRows("_" & rsTmp!����)

                Do While .TextMatrix(lngRow, lngCol) <> ""
                    lngRow = lngRow + 1
                Loop
                If vItem.ͼ��ID <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, lngCol) = GetPathIcon(vItem.ͼ��ID)
                    .Cell(flexcpPictureAlignment, lngRow, lngCol) = 1
                End If
                .TextMatrix(lngRow, lngCol) = vItem.��Ŀ����
                If vItem.ҽ��IDs <> "" Or vItem.����IDs <> "" Or vItem.�°没��IDs <> "" Then
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol) & "��"
                End If
                .Cell(flexcpData, lngRow, lngCol) = vItem
                If vItem.������ <> 1 Then
                    .Cell(flexcpBackColor, lngRow, lngCol) = &HE1FFE1
                End If

                mcolItemRowCol.Add lngRow & "," & lngCol, "_" & vItem.ID '��ʾ����ʱ�����ٶ�λ�����У��������ò��쵥Ԫ��ı���ɫ
                rsTmp.MoveNext
            Loop

            '�����ҽ��
            If strItems <> "" Then
                strItems = Mid(strItems, 2)
                strSql = " Select /*+Cardinality(b,10)*/" & vbNewLine & _
                         " a.��Ŀid, a.ҽ������id, ���״̬, ���ʱ��" & vbNewLine & _
                         " From ����·��ҽ���䶯 A, Table(f_Num2list([1])) B" & vbNewLine & _
                         " Where a.��Ŀid = b.Column_Value And a.���ʱ�� is Null And" & vbNewLine & _
                         "      a.����ʱ�� = (Select Max(����ʱ��) From ����·��ҽ���䶯 C Where c.��Ŀid = a.��Ŀid And c.���ʱ�� is Null)" & vbNewLine & _
                         " Order By a.��Ŀid, a.ҽ������id"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "LoadPathTable", strItems)
                For lngRow = .FixedRows + .FrozenRows To .Rows - 1
                    For lngCol = .FixedCols + .FrozenCols To .Cols - 1
                        If TypeName(.Cell(flexcpData, lngRow, lngCol)) = TypeName(vItem) Then
                            vItem = .Cell(flexcpData, lngRow, lngCol)
                            If vItem.ҽ��IDs <> "" Then
                                rsTmp.Filter = "��ĿID=" & vItem.ID
                                Do While Not rsTmp.EOF
                                    vItem.�����ҽ��IDs = vItem.�����ҽ��IDs & "," & rsTmp!ҽ������ID
                                    rsTmp.MoveNext
                                Loop
                                vItem.�����ҽ��IDs = Mid(vItem.�����ҽ��IDs, 2)
                                .Cell(flexcpData, lngRow, lngCol) = vItem
                                If vItem.�����ҽ��IDs <> "" Then
                                    .Cell(flexcpForeColor, lngRow, lngCol) = Color_NeedAuditFore
                                End If
                            End If
                        End If
                    Next
                Next
            End If

            For i = .FixedCols + .FrozenCols To .Cols - 1
                .ColWidth(i) = COl_WIDTH_BASE
            Next
        End With
    End If

    vsPath.Redraw = flexRDDirect
    vsPath.AutoSize vsPath.FixedCols, vsPath.Cols - 1, , 45         '��ҪDraw֮�����Ч
    Call SetTableCommonStyle(True)
    vsPath.Row = vsPath.FixedRows + vsPath.FrozenRows
    vsPath.Col = vsPath.FixedCols + vsPath.FrozenCols
    If mbytMode = Mode_Design And Visible Then vsPath.SetFocus

    mstrDelStepIDs = ""
    mstrDelItemIDs = ""
    mblnChange = False

    LoadPathTable = True
    Exit Function
errH:
    vsPath.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitAdviceRecordset()
'���ܣ���ʼ����ǰ�汾ҽ�����ݱ�
    If Not mrsAdvice Is Nothing Then
        If mrsAdvice.State = 1 Then mrsAdvice.Close
    End If
    Set mrsAdvice = New ADODB.Recordset

    mrsAdvice.Fields.Append "ID", adBigInt
    mrsAdvice.Fields.Append "�Ƿ�ȱʡ", adSmallInt
    mrsAdvice.Fields.Append "�Ƿ�ѡ", adSmallInt
    mrsAdvice.Fields.Append "���ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "���", adBigInt
    mrsAdvice.Fields.Append "��Ч", adSmallInt
    mrsAdvice.Fields.Append "������ĿID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "�շ�ϸĿID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "��������", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "�ܸ�����", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "�걾��λ", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "��鷽��", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "ҽ������", adVarChar, 1000, adFldIsNullable
    mrsAdvice.Fields.Append "ִ��Ƶ��", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "Ƶ�ʴ���", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "Ƶ�ʼ��", adSmallInt, , adFldIsNullable
    mrsAdvice.Fields.Append "�����λ", adVarChar, 10, adFldIsNullable
    mrsAdvice.Fields.Append "ִ������", adSmallInt
    mrsAdvice.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "ʱ�䷽��", adVarChar, 100, adFldIsNullable
    mrsAdvice.Fields.Append "�䷽ID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "�����ĿID", adBigInt, , adFldIsNullable
    mrsAdvice.Fields.Append "ִ�б��", adSingle, , adFldIsNullable
    mrsAdvice.Fields.Append "�����", adSingle, 1, adFldIsNullable
    mrsAdvice.Fields.Append "��ĿID", adBigInt, , adFldIsNullable       '·��ҽ���䶯ʱ��

    mrsAdvice.CursorLocation = adUseClient
    mrsAdvice.LockType = adLockOptimistic
    mrsAdvice.CursorType = adOpenStatic
    mrsAdvice.Open
End Sub

Private Sub SetTableCommonStyle(Optional ByVal blnKeep As Boolean)
'���ܣ���·��������һЩͳһ����ʽ����
'���ܣ���һЩ�������Ա��ֲ���
    Dim vRedraw As RedrawSettings
    Dim i As Long

    With vsPath
        vRedraw = .Redraw
        If Not blnKeep Then
            .RowHeight(-1) = ROW_HEIGHT_MIN '�������и�
        Else
            For i = .FixedRows To .Rows - 1
                If .RowHeight(i) < ROW_HEIGHT_MIN Then
                    .RowHeight(i) = ROW_HEIGHT_MIN
                End If
            Next
        End If
        '�п�����
        If mbytMode = Mode_Design Then
            .RowHeight(0) = ROW_HEIGHT_MIN
        Else
            .RowHeight(0) = 150
        End If
        .RowHeight(1) = 650 'ʱ��׶���ʾ��

        .Cell(flexcpText, .FixedRows, .FixedCols, .FixedRows + .FrozenRows - 1, .FixedCols + .FrozenCols - 1) = " ʱ��׶� "
        .Cell(flexcpAlignment, 0, 0, .FixedRows + .FrozenRows - 1, .Cols - 1) = 4 '���ͷ
        .Cell(flexcpAlignment, .FixedRows + .FrozenRows, 0, .Rows - 1, .FixedCols + .FrozenCols - 1) = 4 '����ͷ
        .Cell(flexcpAlignment, .FixedRows + .FrozenRows, .FixedCols + .FrozenCols, .Rows - 1, .Cols - 1) = 1 '��Ŀ���ݲ���

        .MergeCol(-1) = True
        .MergeRow(.FixedRows) = True

        '�Զ���ʱ��׶α�ͷ���յĽ׶�������Ϊ�ϲ�Ч��
        If .FrozenRows > 1 Then
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) = "Empty" Then
                    .Cell(flexcpText, .FixedRows, i, .FixedRows + .FrozenRows - 1, i) = Space((i Mod 2) + 1)
                End If
            Next
        End If

        If Not blnKeep Then
            .Row = .FixedRows + .FrozenRows
            .Col = .FixedCols + .FrozenCols
        End If

        .Redraw = vRedraw
    End With
End Sub

Private Function GetArea(ByVal lngRow As Long, ByVal lngCol As Long) As CONST_AREA
'���ܣ���ȡָ����������һ������
    With vsPath
        If lngRow = -1 Or lngCol = -1 Then
            GetArea = -1
        ElseIf lngRow <= .FixedRows - 1 Or lngCol <= .FixedCols - 1 Then
            GetArea = -1
        ElseIf lngCol >= .FixedCols And lngCol <= .FixedCols + .FrozenCols - 1 And lngRow >= .FixedRows And lngRow <= .FixedRows + .FrozenRows - 1 Then
            GetArea = Area_Cross
        ElseIf lngCol >= .FixedCols And lngCol <= .FixedCols + .FrozenCols - 1 Then
            GetArea = Area_Category
        ElseIf lngRow >= .FixedRows And lngRow <= .FixedRows + .FrozenRows - 1 Then
            GetArea = Area_Step
        Else
            GetArea = Area_Item
        End If
    End With
End Function

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = 1 Then
        If fraSplit.Top + Y < 100 Or fraSplit.Top + Y > picCenter.Height - 100 Then Exit Sub
        fraSplit.Top = fraSplit.Top + Y
        vsPath.Height = vsPath.Height + Y

        picBottom.Top = picBottom.Top + Y
        picBottom.Height = picBottom.Height - Y

        UCAdvice(0).Height = UCAdvice(0).Height - Y
        UCAdvice(1).Height = UCAdvice(1).Height - Y
    End If
End Sub

Private Sub fraSplit2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = 1 Then
        If fraSplit2.Left + X < (picBottom.Width / 10) * 1 Or fraSplit2.Left + X > (picBottom.Width / 10) * 9 Then Exit Sub
        fraSplit2.Left = fraSplit2.Left + X
        UCAdvice(0).Width = UCAdvice(0).Width + X

        lblChange.Left = lblChange.Left + X
        cboTimes.Left = cboTimes.Left + X
        If cmdCheck(0).Visible Then cmdCheck(0).Left = cmdCheck(0).Left + X
        If cmdCheck(1).Visible Then cmdCheck(1).Left = cmdCheck(1).Left + X
        UCAdvice(1).Left = UCAdvice(1).Left + X
        UCAdvice(1).Width = UCAdvice(1).Width - X
    End If
End Sub

Private Sub mfrmAdviceContrast_MovePathItemFocus(ByVal lngItemID As Long)
'����:������ĿID,����Ŀ��ý���
'������lngItemID:��ĿID
    Dim strTmp As String
    Dim lngRow As Long, lngCol As Long

    strTmp = mcolItemRowCol("_" & lngItemID)
    lngRow = Split(strTmp, ",")(0)
    lngCol = Split(strTmp, ",")(1)
    With vsPath
        .Row = lngRow
        .Col = lngCol
        '�ԱȲ鿴ʱ��ʵʱ���µ�ǰ��Ŀ����
        Call mfrmAdviceContrast.SetNoteInfo(.TextMatrix(.Row, .Col))
    End With
End Sub

Private Sub mfrmVersion_CalcPathCost(CostMin As Currency, CostMax As Currency)
'���ܣ�����·������
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim rsTmp As ADODB.Recordset
    Dim curCostMin As Currency, curCostMax As Currency
    Dim strSql As String, intDay As Integer
    Dim intDayMin As Integer, intDayMax As Integer

    If mblnChange Then
        MsgBox "·����������δ���棬���ȱ�����ܽ��й��㡣", vbInformation, gstrSysName
        Exit Sub
    End If

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If Not objCombo Is Nothing Then
        If objCombo.ListIndex > 0 Then
            vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
        End If
    End If
    If vVersion.�汾�� = 0 Then Exit Sub

    If InStr(vVersion.��׼����ʱ��, "-") > 0 Then
        intDayMin = Val(Split(vVersion.��׼����ʱ��, "-")(0))
        intDayMax = Val(Split(vVersion.��׼����ʱ��, "-")(1))
    Else
        intDayMin = Val(vVersion.��׼����ʱ��)
        intDayMax = intDayMin
    End If

    Screen.MousePointer = 11
    On Error GoTo errH
    For intDay = 1 To intDayMax
        strSql = "Select zl_GetPathChargeOut(0,0,[1],[2],0,[3],Sysdate,[4]) as ��� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mfrmVersion_CalcPathCost", mlng·��ID, vVersion.�汾��, intDay)

        If intDay <= intDayMin Then
            curCostMin = curCostMin + NVL(rsTmp!���, 0)
        End If
        curCostMax = curCostMax + NVL(rsTmp!���, 0)
    Next
    On Error GoTo 0
    Screen.MousePointer = 0

    If curCostMin = 0 And curCostMax = 0 Then
        MsgBox "�����޷��á�", vbInformation, gstrSysName
    Else
        CostMin = curCostMin: CostMax = curCostMax
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mfrmVersion_CheckDataValid(Version As TYPE_PATH_VERSION, Cancel As Boolean)
    Dim vStep As TYPE_PATH_STEP
    Dim i As Long
    Dim objComboBranch As CommandBarComboBox
    Dim lngBegin As Long, lngEnd As Long
    Dim strSql As String, rsTmp As Recordset
    Dim lngDays As Long

    With vsPath
        If Version.�汾�� = 0 Then
            '��׼����ʱ�䲻ӦС�����н׶ε�������Χ
            For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    'ֻ�������һ���о���������Χ��ʱ��׶�
                    If vStep.�������� <> 0 Or vStep.��ʼ���� <> 0 Then
                        If InStr(Version.��׼����ʱ��, "-") > 0 Then
                            If vStep.�������� <> 0 Then
                                If Val(Split(Version.��׼����ʱ��, "-")(1)) < vStep.�������� Then
                                    MsgBox "��׼����ʱ���������� " & Val(Split(Version.��׼����ʱ��, "-")(1)) & " �첻ӦС��ʱ��׶���ָ�������� " & vStep.�������� & " �졣", vbInformation, gstrSysName
                                    Cancel = True
                                    Exit Sub
                                End If
                            ElseIf vStep.��ʼ���� <> 0 Then
                                If Val(Split(Version.��׼����ʱ��, "-")(1)) < vStep.��ʼ���� Then
                                    MsgBox "��׼����ʱ���������� " & Val(Split(Version.��׼����ʱ��, "-")(1)) & " �첻ӦС��ʱ��׶���ָ�������� " & vStep.��ʼ���� & " �졣", vbInformation, gstrSysName
                                    Cancel = True
                                    Exit Sub
                                End If
                            End If
                        Else
                            If vStep.�������� <> 0 Then
                                If Val(Version.��׼����ʱ��) < vStep.�������� Then
                                    MsgBox "��׼����ʱ�� " & Version.��׼����ʱ�� & " �첻ӦС��ʱ��׶���ָ�������� " & vStep.�������� & " �졣", vbInformation, gstrSysName
                                    Cancel = True
                                    Exit Sub
                                End If
                            ElseIf vStep.��ʼ���� <> 0 Then
                                If Val(Version.��׼����ʱ��) < vStep.��ʼ���� Then
                                    MsgBox "��׼����ʱ�� " & Version.��׼����ʱ�� & " �첻ӦС��ʱ��׶���ָ�������� " & vStep.��ʼ���� & " �졣", vbInformation, gstrSysName
                                    Cancel = True
                                    Exit Sub
                                End If
                            End If
                        End If
                        Exit For
                    End If
                End If
            Next
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call FuncFindItem
    End If
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo txtFind.Hwnd, "����(Ctrl+F)" & vbCrLf & "������һ��(F3)", True
End Sub

Private Sub vsPath_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim vArea As CONST_AREA

    If mbytMode = Mode_Design Then
        vArea = GetArea(NewRow, NewCol)
        If vArea = Area_Category Then
            vsPath.FocusRect = flexFocusSolid
        Else
            vsPath.FocusRect = flexFocusHeavy
        End If
        If picBottom.Visible Then
            If vArea = Area_Item Then
                Call FuncShowItemAdvice
            Else
                Call FuncShowAdvice(2)
            End If
            If mbytFunc = 2 Then
                cmdCheck(0).Enabled = cboTimes.ListCount > 0
                cmdCheck(1).Enabled = cboTimes.ListCount > 0
            End If
        End If
    End If
End Sub

Private Sub vsPath_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsPath
        If Row = -1 And Col <> -1 Then
            .AutoSize .FixedCols, .Cols - 1, , 45
            Call SetTableCommonStyle(True)
        End If
    End With
End Sub

Private Sub vsPath_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
'���Ͻǽ��������������
    If GetArea(NewRow, NewCol) = Area_Cross Then
        If vsPath.Redraw <> flexRDNone Then Cancel = True
    Else
        mlngNewRow = NewRow: mlngNewCol = NewCol
    End If
End Sub

Private Sub vsPath_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    With vsPath
        If GetArea(NewRowSel, NewColSel) = Area_Cross Then
            If .Redraw <> flexRDNone Then Cancel = True '���Ͻǽ��������������
        ElseIf GetArea(NewRowSel, NewColSel) <> GetArea(.Row, .Col) And Not (mlngNewRow = NewRowSel And mlngNewCol = NewColSel) Then
            If .Redraw <> flexRDNone Then Cancel = True '������ͬ���򽻲�ѡ��
        End If
    End With
End Sub

Private Sub vsPath_BeforeSort(ByVal Col As Long, Order As Integer)
    Dim objControl As CommandBarControl

    Order = 0
    If Col >= vsPath.FixedCols + vsPath.FrozenCols And Col <= vsPath.Cols - 1 Then
        vsPath.Col = Col

        Set objControl = cbsMain.FindControl(, cmd_Edit_EvalStep, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then Call FuncEvaluateStep
        End If
    End If
End Sub

Private Sub FuncShowItemAdvice()
    With vsPath
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If mbytFunc = 0 Then Exit Sub

        If TypeName(.Cell(flexcpData, .Row, .Col)) <> "Empty" Then
            If (.Cell(flexcpBackColor, .Row, .Col) = Color_DiffBack Or .Cell(flexcpForeColor, .Row, .Col) = Color_NeedAuditFore) Then
                Call FuncLoadChangeTimes
            Else
                Call FuncShowAdvice(2)
            End If
            Call FuncShowAdvice(0)
        Else
            Call FuncShowAdvice(2)
        End If

        If mbytFunc = 2 Then
            cmdCheck(0).Enabled = cboTimes.ListCount > 0
            cmdCheck(1).Enabled = cboTimes.ListCount > 0
        End If
    End With
End Sub

Private Sub vsPath_DblClick()
    Dim vArea As CONST_AREA
    Dim lngRow As Long, lngCol As Long

    With vsPath
        lngRow = .MouseRow
        lngCol = .MouseCol

        vArea = GetArea(lngRow, lngCol)
        If vArea <> Area_Cross And vArea <> -1 Then
            If mbytMode = Mode_Design And mblnEditable And vArea = Area_Category Then
                '�ɱ༭��δ��ˣ�ʱ��˫�������У��Զ�����༭״̬
                .EditCell
                Exit Sub
            End If
            Call vsPath_KeyPress(13)
        End If
    End With
End Sub

Private Sub vsPath_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl

    If KeyCode = vbKeyDelete And Shift = 0 Then
        Set objControl = cbsMain.FindControl(, cmd_Edit_Delete, True, True)
        If Not objControl Is Nothing Then
            If objControl.Enabled Then objControl.Execute
        End If
    End If
End Sub

Private Sub vsPath_KeyPress(KeyAscii As Integer)
    Dim vArea As CONST_AREA
    Dim objControl As CommandBarControl

    vArea = GetArea(vsPath.Row, vsPath.Col)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If vArea = Area_Category Then
            Call CategoryEnterNextCell(vsPath.Row, vsPath.Col)
        Else
            Set objControl = cbsMain.FindControl(, cmd_Edit_Edit, True, True)
            If Not objControl Is Nothing Then
                If objControl.Enabled Then objControl.Execute
            End If
        End If
    End If
End Sub

Private Sub vsPath_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim vArea As CONST_AREA

    If Button = vbLeftButton Then
        With vsPath
            lngRow = .MouseRow
            lngCol = .MouseCol
            vArea = GetArea(lngRow, lngCol)
            If vArea = -1 Then
                Exit Sub
            ElseIf vArea = Area_Category And .TextMatrix(lngRow, lngCol) = "" Then
                .EditCell   '���Ϊ�գ�ǿ�Ʊ༭
            End If
        End With
    End If
End Sub

Private Sub vsPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ���ʾ������ʾ
    Dim lngRow As Long, lngCol As Long
    Dim vArea As CONST_AREA
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim strTip As String, strTmp As String, i As Long
    Dim rsTmp As ADODB.Recordset

    With vsPath
        If .Rows = 0 Or .Cols = 0 Then
            zlCommFun.ShowTipInfo 0, ""
            Exit Sub
        End If
        lngRow = .MouseRow
        lngCol = .MouseCol

        vArea = GetArea(lngRow, lngCol)
        If vArea = Area_Step Then
            If TypeName(.ColData(lngCol)) <> "Empty" Then
                vStep = .ColData(lngCol)
                If vStep.�������� <> 0 Then
                    strTip = strTip & "ʱ��׶Σ������" & vStep.��ʼ���� & "-" & vStep.�������� & "��"
                Else
                    strTip = strTip & "ʱ��׶Σ������" & vStep.��ʼ���� & "��"
                End If
                If lngRow = .FixedRows + .FrozenRows - 1 Then
                    If .TextMatrix(lngRow, lngCol) <> .TextMatrix(lngRow - 1, lngCol) Then
                        strTip = strTip & "��" & .TextMatrix(lngRow, lngCol)
                    End If
                End If
                If vStep.���� <> "" Then
                    strTip = strTip & vbCrLf & "����ࣺ" & vStep.����
                End If
                strTip = strTip & vbCrLf & "��˵����" & vStep.˵��
            End If
        ElseIf vArea = Area_Item Then
            If TypeName(.Cell(flexcpData, lngRow, lngCol)) <> "Empty" Then
                vItem = .Cell(flexcpData, lngRow, lngCol)
                strTip = "·����Ŀ��" & vItem.��Ŀ����
                If vItem.������ = 1 Then
                    If vItem.ҽ��IDs <> "" Then
                        If Not vItem.Tip Like vItem.ҽ��IDs & ":*" Then
                            vItem.Tip = vItem.ҽ��IDs & ":" & GetAdviceDefineText(vItem.ҽ��IDs, mrsAdvice)
                            .Cell(flexcpData, lngRow, lngCol) = vItem
                        End If
                        strTip = strTip & vbCrLf & "��ҽ��ժҪ��" & vbCrLf & Mid(vItem.Tip, InStr(vItem.Tip, ":") + 1)
                    End If
                    If vItem.����IDs <> "" Or vItem.�°没��IDs <> "" Then
                        If Not vItem.Tip Like vItem.����IDs & "|" & vItem.�°没��IDs & ":*" Then
                            If vItem.Edit = 0 Then
                                If vItem.����IDs <> "" And vItem.�°没��IDs <> "" Then
                                    strTmp = GetEPRDefineTextOut(, vItem.ID)
                                ElseIf vItem.����IDs <> "" Then
                                    strTmp = GetEPRDefineTextOut(vItem.����IDs)
                                Else
                                    strTmp = GetEPRDefineTextOut(vItem.�°没��IDs, vItem.ID)
                                End If
                            Else
                                Set rsTmp = FuncGetEMRInfo(vItem.��������)
                                strTmp = ""
                                Do While Not rsTmp.EOF
                                    strTmp = strTmp & "��" & rsTmp!����
                                    rsTmp.MoveNext
                                Loop
                                strTmp = Mid(strTmp, 2)
                            End If
                            vItem.Tip = vItem.����IDs & "|" & vItem.�°没��IDs & ":" & strTmp
                            .Cell(flexcpData, lngRow, lngCol) = vItem
                        End If
                        strTip = strTip & vbCrLf & "���Ӧ������" & Mid(vItem.Tip, InStr(vItem.Tip, ":") + 1)
                    End If
                    strTip = strTip & vbCrLf & "��ִ�з�ʽ��" & Decode(vItem.ִ�з�ʽ, 0, "����ִ��", 1, "����ִ��", 3, "��Ҫʱִ��")
                    If vItem.ִ�з�ʽ <> 0 Then
                        If vItem.��Ŀ��� <> "" Then
                            strTmp = ""
                            For i = 0 To UBound(Split(Split(vItem.��Ŀ���, vbTab)(0), ","))
                                strTmp = strTmp & "��" & Split(Split(Split(vItem.��Ŀ���, vbTab)(0), ",")(i), "|")(0)
                            Next
                            strTip = strTip & vbCrLf & "��ִ�н����" & Mid(strTmp, 2)
                            strTip = strTip & vbCrLf & "��ȱʡ�����" & Split(vItem.��Ŀ���, vbTab)(1)
                        End If
                    End If
                Else
                    strTip = vItem.����ο�
                End If
            End If
        ElseIf vArea = Area_Cross Or lngRow = .FixedRows - 1 And lngCol <= .FixedCols + .FrozenCols - 1 And lngCol >= 0 Then
            If Not mvEvalImport.ָ�꼯 Is Nothing Then
                If mvEvalImport.ָ�꼯.count > 0 Then
                    strTip = strTip & vbCrLf & "������ָ�꣺"
                    For i = 1 To mvEvalImport.ָ�꼯.count
                        vEvalMark = mvEvalImport.ָ�꼯(i)
                        strTip = strTip & vbCrLf & "����" & vEvalMark.����ָ�� & "�������" & Split(vEvalMark.ָ����, vbTab)(0)
                    Next
                End If
            End If
            If Not mvEvalImport.������ Is Nothing Then
                If mvEvalImport.������.count > 0 Then
                    strTip = strTip & vbCrLf & "�����������"
                    For i = 1 To mvEvalImport.������.count
                        vEvalCond = mvEvalImport.������(i)
                        strTip = strTip & vbCrLf & "����[" & GetMarkName(vEvalCond.ָ��ID, lngCol) & "] " & vEvalCond.��ϵʽ & " [" & vEvalCond.����ֵ & "]"
                    Next
                End If
            End If
            If strTip <> "" Then
                strTip = "����������Ϣ��" & strTip
            Else
                strTip = "û�����õ���������Ϣ��"
            End If
        ElseIf lngRow = .FixedRows - 1 And lngCol >= .FixedCols + .FrozenCols Then
            If TypeName(.ColData(lngCol)) <> "Empty" Then
                vStep = .ColData(lngCol)
                If Not vStep.����.ָ�꼯 Is Nothing Then
                    If vStep.����.ָ�꼯.count > 0 Then
                        strTip = strTip & vbCrLf & "������ָ�꣺"
                        For i = 1 To vStep.����.ָ�꼯.count
                            vEvalMark = vStep.����.ָ�꼯(i)
                            strTip = strTip & vbCrLf & "����" & vEvalMark.����ָ�� & "�������" & Split(vEvalMark.ָ����, vbTab)(0)
                        Next
                    End If
                End If
                If Not vStep.����.������ Is Nothing Then
                    If vStep.����.������.count > 0 Then
                        strTip = strTip & vbCrLf & "�����������"
                        For i = 1 To vStep.����.������.count
                            vEvalCond = vStep.����.������(i)
                            If vEvalCond.ָ��ID <> 0 Then
                                strTip = strTip & vbCrLf & "����[" & GetMarkName(vEvalCond.ָ��ID, lngCol) & "] " & vEvalCond.��ϵʽ & " [" & vEvalCond.����ֵ & "]"
                            ElseIf vEvalCond.��ĿID <> 0 Then
                                strTip = strTip & vbCrLf & "����[" & GetItemName(vEvalCond.��ĿID, lngCol) & "] " & vEvalCond.��ϵʽ & " [" & vEvalCond.����ֵ & "]"
                            End If
                        Next
                    End If
                End If
                If strTip <> "" Then
                    strTip = "�׶�������Ϣ��" & strTip
                Else
                    If mbytMode = Mode_Design And mblnEditable Then
                        strTip = "��δ���ø�ʱ��׶ε�������Ϣ���������á�"
                    Else
                        strTip = "û�����ø�ʱ��׶ε�������Ϣ��"
                    End If
                End If
            Else
                If mbytMode = Mode_Design And mblnEditable Then
                    strTip = "��δ���ø�ʱ��׶ε�������Ϣ���������á�"
                Else
                    strTip = "û�����ø�ʱ��׶ε�������Ϣ��"
                End If
            End If
        End If

        If strTip <> "" Then
            zlCommFun.ShowTipInfo .Hwnd, strTip, True
        Else
            zlCommFun.ShowTipInfo 0, ""
        End If
    End With
End Sub

Private Function GetItemName(ByVal lngItemID As Long, ByVal lngCol As Long) As String
'���ܣ���ȡָ���׶���ָ����ĿID����Ŀ����
    Dim vItem As TYPE_PATH_ITEM
    Dim i As Long

    With vsPath
        For i = .FixedRows + .FrozenRows To .Rows - 1
            If TypeName(.Cell(flexcpData, i, lngCol)) <> "Empty" Then
                vItem = .Cell(flexcpData, i, lngCol)
                If vItem.ID = lngItemID Then
                    GetItemName = vItem.��Ŀ����
                    Exit Function
                End If
            End If
        Next
    End With
End Function

Private Function GetMarkName(ByVal lngMarkID As Long, Optional ByVal lngCol As Long)
'���ܣ���ȡָ��ָ��ID��ָ������
'������lngCol=ָ��ʱΪ����Ľ׶��У������ʾ��������ָ��
    Dim vStep As TYPE_PATH_STEP
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim i As Long

    If lngCol = 0 Then
        If Not mvEvalImport.ָ�꼯 Is Nothing Then
            For i = 1 To mvEvalImport.ָ�꼯.count
                vEvalMark = mvEvalImport.ָ�꼯(i)
                If vEvalMark.ID = lngMarkID Then
                    GetMarkName = vEvalMark.����ָ��
                    Exit Function
                End If
            Next
        End If
    Else
        If TypeName(vsPath.ColData(lngCol)) <> "Empty" Then
            vStep = vsPath.ColData(lngCol)
            If Not vStep.����.ָ�꼯 Is Nothing Then
                For i = 1 To vStep.����.ָ�꼯.count
                    vEvalMark = vStep.����.ָ�꼯(i)
                    If vEvalMark.ID = lngMarkID Then
                        GetMarkName = vEvalMark.����ָ��
                        Exit Function
                    End If
                Next
            End If
        End If
    End If
End Function

Private Sub vsPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim vArea As CONST_AREA
    Dim objPopup As CommandBarPopup

    If Button = 2 Then
        lngRow = vsPath.MouseRow
        lngCol = vsPath.MouseCol
        vArea = GetArea(lngRow, lngCol)
        If vArea <> Area_Cross And vArea <> -1 Then
            '�Ⱥ�˳����BeforeRowColChange�¼���������
            If vsPath.Col = vsPath.FixedCols Then
                vsPath.Col = lngCol: vsPath.Row = lngRow
            Else
                vsPath.Row = lngRow: vsPath.Col = lngCol
            End If
            Set objPopup = cbsMain.FindControl(, conMenu_EditPopup, True)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub vsPath_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsPath.EditSelStart = 0
    vsPath.EditSelLength = zlCommFun.ActualLen(vsPath.EditText)
End Sub

Private Sub vsPath_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPath
        If GetArea(Row, Col) <> Area_Category Then
            Cancel = True                               '���������ݿ���ֱ������
        ElseIf .RowSel <> Row Or .ColSel <> Col Then
            Cancel = True                               'ѡ��Χʱ����������
        End If
    End With
End Sub

Private Sub vsPath_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngR1 As Long, lngR2 As Long
    Dim vArea As CONST_AREA, i As Long, j As Long
    Dim vItem As TYPE_PATH_ITEM

    vArea = GetArea(Row, Col)

    With vsPath
        If vArea = Area_Category Then
            .EditText = Trim(.EditText)
            If LenB(StrConv(.EditText, vbFromUnicode)) > 50 Then
                MsgBox "������ķ������Ƶ���������25�������������롣", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
            'û�иĶ�ʱ����������
            If .TextMatrix(Row, Col) = .EditText Then
                Exit Sub
            End If
            If Trim(.EditText) = "" Then
                '�൱��ɾ��,�ж�Ӧ��Ŀʱ���������
                .GetMergedRange Row, Col, lngR1, 0, lngR2, 0
                For i = lngR1 To lngR2
                    If Replace(.Cell(flexcpText, i, .FixedCols + .FrozenCols, i, .Cols - 1), vbTab, "") <> "" Then
                        MsgBox "�÷������Ѿ����ڶ�Ӧ����Ŀ��������������ơ�", vbInformation, gstrSysName
                        Cancel = True: Exit Sub
                    End If
                Next
            Else
                '���಻���ظ�:�ϲ���Ԫ��Χ�ؼ��Զ�����
                i = .FixedRows + .FrozenCols
                Do While i <= .Rows - 1
                    If i <> Row And .TextMatrix(i, Col) = .EditText Then
                        MsgBox "����ķ��������Ѿ����ڣ����������롣", vbInformation, gstrSysName
                        Cancel = True: Exit Sub
                    End If

                    .GetMergedRange i, Col, 0, 0, lngR2, 0
                    i = lngR2 + 1    '�����ϲ�����
                Loop
                .GetMergedRange Row, Col, lngR1, 0, lngR2, 0

                '�÷����������Ŀ�����Ϊ�޸�״̬
                For i = .FixedCols + .FrozenCols To .Cols - 1
                    For j = lngR1 To lngR2
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)
                            vItem.Edit = 2
                            .Cell(flexcpData, j, i) = vItem
                        End If
                    Next
                Next
            End If

            '���ݱ仯�󣬸��������Զ����и�
            For i = lngR1 To lngR2
                .TextMatrix(i, Col) = .EditText    '��Ȼ������Ч
            Next i
            .AutoSize .FixedCols, .Cols - 1, , 45
            Call SetTableCommonStyle(True)

            mblnChange = True

            '���������һ����
            If mblnReturn And Trim(.EditText) <> "" Then
                Call CategoryEnterNextCell(Row, Col)
            End If
        End If
    End With
End Sub

Private Sub CategoryEnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim lngR2 As Long

    With vsPath
        .GetMergedRange lngRow, lngCol, 0, 0, lngR2, 0
        If lngR2 + 1 <= .Rows - 1 Then
            .Row = lngR2 + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Sub FuncCategoryDelete()
'���ܣ�ɾ����ǰѡ��ķ�����
    Dim lngR1 As Long, lngR2 As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngRow As Long, i As Long

    With vsPath
        lngRow = .Row

        '��ǰѡ��Χ
        .GetSelection lngR1, 0, lngR2, 0
        lngBegin = lngR1

        '���Ǻϲ���Ԫ���ϲ���Ԫѡ��ʱRowSel,ColSel����
        .GetMergedRange lngR2, .Col, lngR1, 0, lngR2, 0
        lngEnd = lngR2

        For i = lngBegin To lngEnd
            If Replace(.Cell(flexcpText, i, .FixedCols + .FrozenCols, i, .Cols - 1), vbTab, "") <> "" Then
                MsgBox "��ѡ��������Ѿ����ڶ�Ӧ��·����Ŀ������ɾ����", vbInformation, gstrSysName
                Exit Sub
            End If
        Next

        If MsgBox("ȷʵҪɾ����ѡ��ķ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

        'ɾ������
        .Redraw = flexRDNone
        For i = lngEnd To lngBegin Step -1
            .RemoveItem i
        Next
        If .Rows = .FixedRows + .FrozenRows Then
            'ɾ�������ٱ���һ�з���
            .AddItem "": .RowHeight(.Rows - 1) = ROW_HEIGHT_MIN
            .Row = .FixedRows + .FrozenRows
        ElseIf lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
        .SetFocus
    End With

    mblnChange = True
End Sub

Private Sub FuncCategoryInsert(ByVal intPos As Integer)
'���ܣ������µķ�����
'������inPos=1���ڵ�ǰ�к��棬-1���ڵ�ǰ��ǰ��
    Dim lngRow As Long
    Dim lngR1 As Long, lngR2 As Long

    With vsPath
        If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "��ǰ�з�����δ���룬�������뵱ǰ�з��ࡣ", vbInformation, gstrSysName
            Exit Sub
        End If

        .GetMergedRange .Row, .Col, lngR1, 0, lngR2, 0    '��Ҫ���Ǻϲ��Χ
        lngRow = IIf(intPos = -1, lngR1, lngR2 + 1)
        .AddItem "", lngRow
        .RowHeight(lngRow) = ROW_HEIGHT_MIN
        .Row = lngRow
        .EditCell
        .ShowCell .Row, .Col
    End With
    mblnChange = True
End Sub

Private Sub mfrmPathStep_CheckDataValid(TimeStep As TYPE_PATH_STEP, Cancel As Boolean)
'���ܣ����������ʱ��׶����ݵ���ȷ��
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vStep As TYPE_PATH_STEP
    Dim strMsg As String, i As Long

    With vsPath
        '���׼����ʱ��֮��Ĺ�ϵ���
        Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
        If Not objCombo Is Nothing Then
            If Not objCombo.ListIndex = 0 Then
                vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
                If vVersion.��׼����ʱ�� <> "" Then
                    If InStr(vVersion.��׼����ʱ��, "-") > 0 Then
                        If TimeStep.�������� <> 0 And TimeStep.�������� > Val(Split(vVersion.��׼����ʱ��, "-")(1)) Then
                            MsgBox "��ǰʱ��׶εĽ������� " & TimeStep.�������� & " ������˱�׼����ʱ��ָ����������� " & Val(Split(vVersion.��׼����ʱ��, "-")(1)) & " �졣", vbInformation, gstrSysName
                            Cancel = True: Exit Sub
                        ElseIf TimeStep.��ʼ���� <> 0 And TimeStep.��ʼ���� > Val(Split(vVersion.��׼����ʱ��, "-")(1)) Then
                            MsgBox "��ǰʱ��׶ε����� " & TimeStep.��ʼ���� & " ������˱�׼����ʱ��ָ����������� " & Val(Split(vVersion.��׼����ʱ��, "-")(1)) & " �졣", vbInformation, gstrSysName
                            Cancel = True: Exit Sub
                        End If
                    Else
                        If TimeStep.�������� <> 0 And TimeStep.�������� > Val(vVersion.��׼����ʱ��) Then
                            MsgBox "��ǰʱ��׶εĽ������� " & TimeStep.�������� & " ������˱�׼����ʱ��ָ����������� " & Val(vVersion.��׼����ʱ��) & " �졣", vbInformation, gstrSysName
                            Cancel = True: Exit Sub
                        ElseIf TimeStep.��ʼ���� <> 0 And TimeStep.��ʼ���� > Val(vVersion.��׼����ʱ��) Then
                            MsgBox "��ǰʱ��׶ε����� " & TimeStep.��ʼ���� & " ������˱�׼����ʱ��ָ����������� " & Val(vVersion.��׼����ʱ��) & " �졣", vbInformation, gstrSysName
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
            End If
        End If

        '����������ʱ��׶�֮��Ĺ�ϵ���
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.��ID = 0 And vStep.ID <> TimeStep.ID Then
                    '������ΧӦ����ǰ��֮��,����֮ǰ��������Χ���Բ��ֽ���,Ҳ���ܰ���
                    If i < .Col Then
                        If TimeStep.��ʼ���� <= vStep.��ʼ���� And TimeStep.��ʼ���� <> 0 And vStep.��ʼ���� <> 0 Then
                            If TimeStep.��ʼ���� <> vStep.��ʼ���� Then
                                strMsg = "��ǰ�׶εĿ�ʼ����Ӧ�ô���ǰ��׶εĿ�ʼ������": Exit For
                            End If
                        End If
                        If IIf(TimeStep.�������� = 0, TimeStep.��ʼ����, TimeStep.��������) < IIf(vStep.�������� = 0, vStep.��ʼ����, vStep.��������) Then
                            strMsg = "��ǰ�׶εĽ�������Ӧ�ô���ǰ��׶εĽ���������": Exit For
                        End If
                        If IIf(vStep.�������� = 0, vStep.��ʼ����, vStep.��������) < TimeStep.��ʼ���� - 1 And i = .Col - 1 Then
                            strMsg = "��ǰ�׶εĿ�ʼ���������ǰһ���׶�Ϊ������,��ʼ��������С�ڻ����" & TimeStep.��ʼ���� - 1 & "��": Exit For
                        End If
                    ElseIf i > .Col Then
                        If TimeStep.��ʼ���� >= vStep.��ʼ���� And TimeStep.��ʼ���� <> 0 And vStep.��ʼ���� <> 0 Then
                            If TimeStep.��ʼ���� <> vStep.��ʼ���� Then
                                strMsg = "��ǰ�׶εĿ�ʼ����Ӧ��С�ں���׶εĿ�ʼ������": Exit For
                            End If
                        End If
                        If IIf(TimeStep.�������� = 0, TimeStep.��ʼ����, TimeStep.��������) > IIf(vStep.�������� = 0, vStep.��ʼ����, vStep.��������) Then
                            strMsg = "��ǰ�׶εĽ�������Ӧ��С�ں���׶εĽ���������": Exit For
                        End If
                        If IIf(TimeStep.�������� = 0, TimeStep.��ʼ����, TimeStep.��������) < vStep.��ʼ���� - 1 And i = .Col + 1 Then
                            strMsg = "��ǰ�׶εĿ�ʼ���������ǰһ���׶�Ϊ������,��������������ڻ���ڣ�" & vStep.��ʼ���� - 1 & "�졣": Exit For
                        End If
                    End If
                End If
            End If
        Next

        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetNearStep(ByVal lngCol As Long, ByVal intPos As Integer, _
    Optional ByVal blnSub As Boolean, Optional ByVal blnSkip As Boolean = True) As TYPE_PATH_STEP
'���ܣ���ȡ��ǰʱ��׶����ڵ�ʱ��׶���Ϣ
'������lngCol=��ǰ��
'      intPos=-1:ǰ�棬1:����
'      blnSub=�Ƿ����������ڵķ�֧
'      blnSkip=�Ƿ����������յ�ʱ��׶�
    Dim vStep As TYPE_PATH_STEP
    Dim i As Long

    With vsPath
        If intPos = -1 Then
            For i = lngCol - 1 To .FixedCols + .FrozenCols Step -1
                If TypeName(.ColData(i)) <> "Empty" Then
                    If blnSub Or .ColData(i).���� <> vStep.���� Then
                        vStep = .ColData(i): Exit For
                    End If
                Else
                    If Not blnSkip Then Exit For
                End If
            Next
        Else
            For i = lngCol + 1 To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    If blnSub Or .ColData(i).���� <> vStep.���� Then
                        vStep = .ColData(i): Exit For
                    End If
                Else
                    If Not blnSkip Then Exit For
                End If
            Next
        End If
    End With

    GetNearStep = vStep
End Function

Private Sub FuncStepEdit()
'���ܣ����õ�ǰʱ��׶�����
    Dim vStep As TYPE_PATH_STEP
    Dim vPreStep As TYPE_PATH_STEP
    Dim vNextStep As TYPE_PATH_STEP
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim str����s As String, i As Long, j As Long

    With vsPath
        If TypeName(.ColData(.Col)) <> "Empty" Then
            vStep = .ColData(.Col)
        End If

        '��ȡ����ʱ��׶ε�����
        vPreStep = GetNearStep(.Col, -1)
        vNextStep = GetNearStep(.Col, 1)

        '��ȡǰ���÷�֧�ķ�����
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                If .ColData(i).���� <> "" Then
                    If InStr(str����s & "|", "|" & .ColData(i).���� & "|") = 0 Then
                        str����s = str����s & "|" & .ColData(i).����
                    End If
                End If
            End If
        Next

        If mfrmPathStep.ShowEdit(Me, vStep, vPreStep, vNextStep, Mid(str����s, 2)) Then
            If vStep.ID = 0 Then
                '��֤�����ݵĽ׶�ID��Ϊ�գ���Ԥȡһ��ID
                vStep.ID = zlDatabase.GetNextId("����·���׶�")
                vStep.Edit = 1 '0-ԭʼ,1-����,2-�޸�
            Else
                If vStep.Edit = 0 Then vStep.Edit = 2
            End If

            If vStep.��ID <> 0 Then
                '��ѡ��֧����������˵���ͷ���
                .ColData(.Col) = vStep
                .TextMatrix(.Row, .Col) = IIf(vStep.˵�� = "", "���÷�֧" & vStep.���, vStep.˵��) & IIf(vStep.���� = "", "", ",") & vStep.����
            ElseIf vStep.��ID = 0 Then
                .ColData(.Col) = vStep

                '��ѡ��֧�����Ϣͬ���仯
                For i = .Col + 1 To .Cols - 1
                    If TypeName(.ColData(i)) <> "Empty" Then
                        vNextStep = .ColData(i)
                        If vNextStep.��ID = vStep.ID Then
                            vNextStep.���� = vStep.����
                            vNextStep.��ʼ���� = vStep.��ʼ����
                            vNextStep.�������� = vStep.��������
                            If vNextStep.Edit = 0 Then vNextStep.Edit = 2

                            .ColData(i) = vNextStep
                        Else
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next

                '��һ������ʱ����ʹ������Ҳ���ÿո�ϲ��˵�
                .GetMergedRange .Row, .Col, lngR1, lngC1, lngR2, lngC2
                If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
                    'ѡ��ȱʡ��֧�����,�����ʱ��׶ε�������ʾ�仯
                    .GetMergedRange .FixedRows, .Col, lngR1, lngC1, lngR2, lngC2
                End If

                '���������������������ϲ��Ķ����֧������ϲ���һ��ʱ��׶Σ�����û�ϲ���һ��ʱ��׶�
                '.Cell(flexcpText, lngR1, lngC1, lngR2, lngC2) = vStep.����
                '���ֱ�ӷ�Χ��ֵ����Ϊ�����س����Զ�ʶ��Ϊ�ָ��������������ֱ��ж�
                '������Ǳ༭��֧�У��򲻺ϲ���Ԫ�񣬼�һ���ո�
                If .Row = 1 Then
                    If vPreStep.���� = vStep.���� Then
                        vStep.���� = vStep.���� & " "
                        .ColData(.Col) = vStep
                    End If
                    If vNextStep.���� = vStep.���� Then
                        vStep.���� = vStep.���� & " "
                        .ColData(.Col) = vStep
                    End If
                End If
                For i = lngC1 To lngC2
                    For j = lngR1 To lngR2
                        .TextMatrix(j, i) = vStep.����
                    Next
                Next
            End If

            .TextMatrix(.FixedRows - 1, .Col) = "�׶�������"
            .Cell(flexcpFontBold, .FixedRows - 1, .Col) = False
            If Not vStep.����.ָ�꼯 Is Nothing And Not Not vStep.����.������ Is Nothing Then
                If vStep.����.ָ�꼯.count > 0 Or vStep.����.������.count > 0 Then
                    .Cell(flexcpFontBold, .FixedRows - 1, .Col) = True
                End If
            End If

            mblnChange = True
        End If
    End With
End Sub

Private Sub FuncStepInsert(ByVal intPos As Integer)
'���ܣ������µ�ʱ��׶�
'������inPos=1���ڵ�ǰʱ��׶κ��棬-1���ڵ�ǰʱ��׶�ǰ��
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long

    With vsPath
        If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "��ǰ�׶���δ���ã��������õ�ǰ�׶���Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If

        '��ȡ�����λ��
        .GetMergedRange .Row, .Col, lngR1, lngC1, lngR2, lngC2
        If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
            'ѡ�з�֧�����,GetMergedRange�����ںϲ���Χ���κε�Ԫ
            .GetMergedRange .FixedRows, .Col, 0, lngC1, 0, lngC2
        End If

        '�����µ�ʱ��׶���
        .Redraw = flexRDNone

        .Cols = .Cols + 1
        .ColWidth(.Cols - 1) = COl_WIDTH_BASE

        If intPos = -1 Then
            .ColPosition(.Cols - 1) = lngC1
            .Col = lngC1
        Else
            .ColPosition(.Cols - 1) = lngC2 + 1
            .Col = lngC2 + 1
        End If

        Call SetTableCommonStyle(True)
         .Row = .FixedRows
         .ShowCell .Row, .Col

        .Redraw = flexRDDirect

        '����֮�������������
        Call FuncStepEdit
    End With

    mblnChange = True
End Sub

Private Sub FuncStepBranchInsert()
'���ܣ��ڵ�ǰʱ��׶������µķ�֧
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim vStep As TYPE_PATH_STEP

    With vsPath
        '��ȡ�����λ��
        .GetMergedRange .Row, .Col, lngR1, lngC1, lngR2, lngC2
        If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
            'ѡ�з�֧�����,GetMergedRange�����ںϲ���Χ���κε�Ԫ
            .GetMergedRange .FixedRows, .Col, 0, lngC1, 0, lngC2
        End If

        .Redraw = flexRDNone

        '�����µ�ʱ��׶���
        .Cols = .Cols + 1
        .ColWidth(.Cols - 1) = COl_WIDTH_BASE
        .ColPosition(.Cols - 1) = lngC2 + 1
        .Col = lngC2 + 1

        '����ȱʡ��������
        vStep = .ColData(.Col - 1)
        vStep.��� = IIf(vStep.��ID <> 0, vStep.��� + 1, 1) '��֧��ű�֤1-N����
        vStep.��ID = IIf(vStep.��ID <> 0, vStep.��ID, vStep.ID)
        vStep.ID = zlDatabase.GetNextId("����·���׶�") '��֤���µ�ΨһID
        vStep.���� = ""
        vStep.˵�� = ""
        vStep.Edit = 1 '0-ԭʼ,1-����,2-�޸�

        Set vStep.����.������ = Nothing
        Set vStep.����.ָ�꼯 = Nothing

        .ColData(.Col) = vStep

        '���ý���ϲ���ʾЧ��
        If .FrozenRows = 1 Then
            .AddItem .Cell(flexcpText, .FixedRows, .FixedCols, .FixedRows, .Cols - 1), .FixedRows + 1
            .FrozenRows = 2
        End If
        .TextMatrix(.FixedRows, .Col) = vStep.����
        .TextMatrix(.FixedRows + .FrozenRows - 1, .Col) = IIf(vStep.˵�� = "", "���÷�֧" & vStep.���, vStep.˵��) & IIf(vStep.���� = "", "", ",") & vStep.����
        If vStep.��� = 1 Then
            .TextMatrix(.FixedRows + .FrozenRows - 1, .Col - 1) = "ȱʡ��֧"
        End If

        .Redraw = flexRDDirect

        .AutoSize .FixedCols, .Cols - 1, , 45 'Redraw����Ч
        Call SetTableCommonStyle(True)
         .Row = .FixedRows + .FrozenRows - 1
         .ShowCell .Row, .Col

        '����֮�������������
        Call FuncStepEdit
    End With

    mblnChange = True
End Sub

Private Sub FuncStepDelete()
'���ܣ��ڵ�ǰʱ��׶�ɾ����֧
    Dim lngR1 As Long, lngR2 As Long
    Dim lngC1 As Long, lngC2 As Long
    Dim vStep As TYPE_PATH_STEP
    Dim vSubStep As TYPE_PATH_STEP
    Dim lng��ID As Long, blnSub As Boolean
    Dim i As Long, j As Long
    Dim blnIsDelete As Boolean

    With vsPath
        '��ȡѡ��Χ
        .GetSelection lngR1, lngC1, lngR2, lngC2
        If lngC1 = lngC2 And lngR1 = lngR2 And lngR1 = .FixedRows + .FrozenRows - 1 And lngR1 - 1 = .FixedRows Then
            blnSub = True 'ѡ�з�֧�����
        End If
        If Not blnSub Then
            .GetMergedRange .FixedRows, lngC2, 0, 0, 0, lngC2
        End If

        For i = lngC1 To lngC2
            If Replace(.Cell(flexcpText, .FixedRows + .FrozenRows, i, .Rows - 1, i), vbCr, "") <> "" Then
                If MsgBox("��ѡ���ʱ��׶�(���֧)�д����Ѿ������·����Ŀ,ɾ���׶ν�ͬʱɾ����Щ��Ŀ���Ƿ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                'ɾ��·����Ŀ
                For j = .FixedRows + .FrozenRows To .Rows - 1
                    If .TextMatrix(j, i) <> "" Then
                        .Row = j: .Col = i
                        Call FuncItemDelete(False)
                    End If
                Next
                blnIsDelete = True
                Exit For
            End If
        Next
        If Not blnIsDelete Then
            If MsgBox("ȷʵҪɾ����ѡ���" & IIf(blnSub, "��֧", "ʱ��׶�") & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If

        .Redraw = flexRDNone

        'ɾ������(����)
        For i = lngC2 To lngC1 Step -1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.��ID <> 0 Then
                    '��������ķ�֧���
                    For j = i + 1 To .Cols - 1
                        If TypeName(.ColData(j)) <> "Empty" Then
                            vSubStep = .ColData(j)
                            If vSubStep.��ID = vStep.��ID Then
                                vSubStep.��� = vSubStep.��� - 1
                                .TextMatrix(.FixedRows + .FrozenRows - 1, j) = IIf(vSubStep.˵�� = "", "���÷�֧" & vSubStep.���, vSubStep.˵��) & IIf(vSubStep.���� = "", "", ",") & vSubStep.����
                                .ColData(j) = vSubStep
                            Else
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf vStep.��ID = 0 Then
                    'ȱʡ��֧ɾ��֮��ѡ��֧���Գ�Ϊȱʡ��֧
                    lng��ID = 0
                    For j = i + 1 To .Cols - 1
                        If TypeName(.ColData(j)) <> "Empty" Then
                            vSubStep = .ColData(j)
                            If vSubStep.��ID = vStep.ID Then
                                If j = i + 1 Then
                                    lng��ID = vSubStep.ID

                                    vSubStep.��ID = 0
                                    vSubStep.��� = 0 '����ʱ��������
                                    .TextMatrix(.FixedRows + .FrozenRows - 1, j) = "ȱʡ��֧"
                                Else
                                    vSubStep.��ID = lng��ID
                                    vSubStep.��� = vSubStep.��� - 1
                                    .TextMatrix(.FixedRows + .FrozenRows - 1, j) = IIf(vSubStep.˵�� = "", "���÷�֧" & vSubStep.���, vSubStep.˵��) & IIf(vSubStep.���� = "", "", ",") & vSubStep.����
                                End If
                                .ColData(j) = vSubStep
                            Else
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If

                '��¼ɾ������:0-ԭʼ,1-����,2-�޸�
                If vStep.Edit <> 1 Then
                    mstrDelStepIDs = mstrDelStepIDs & "," & vStep.ID
                End If
            End If
            .ColPosition(i) = .Cols - 1
            .Cols = .Cols - 1
        Next

        '����֧�����
        blnSub = False
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.��ID <> 0 Then
                    blnSub = True: Exit For
                End If
            End If
        Next
        If Not blnSub Then
            If .FrozenRows > 1 Then
                .FrozenRows = 1
                .RemoveItem .FixedRows + .FrozenRows
            End If
        Else
            '����޷�֧,������ʾ�˷�֧��ͷ������
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vStep.��ID = 0 Then
                        If GetNearStep(i, 1, True, False).��ID <> vStep.ID Then
                            If .TextMatrix(.FixedRows + .FrozenRows - 1, i) <> .TextMatrix(.FixedRows, i) Then
                                .TextMatrix(.FixedRows + .FrozenRows - 1, i) = .TextMatrix(.FixedRows, i)
                            End If
                        End If
                    End If
                End If
            Next
        End If

        '���ж�λ
        If lngC1 <= .Cols - 1 Then
            .Col = lngC1
        ElseIf .Cols > .FixedCols + .FrozenCols Then
            .Col = .Cols - 1
        ElseIf .Cols = .FixedCols + .FrozenCols Then
            .Cols = .Cols + 1: .Col = .Cols - 1
            .ColWidth(.Cols - 1) = COl_WIDTH_BASE
        End If
        .Row = .FixedRows

        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
    End With

    mblnChange = True
End Sub

Private Sub FuncItemEdit(Optional ByVal objControl As CommandBarControl)
'���ܣ����õ�ǰ·����Ŀ����
    Dim vStep    As TYPE_PATH_STEP
    Dim vPreStep As TYPE_PATH_STEP
    Dim vItem    As TYPE_PATH_ITEM
    Dim vBakItem As TYPE_PATH_ITEM
    Dim vPreItem As TYPE_PATH_ITEM
    Dim vTmpItem As TYPE_PATH_ITEM
    Dim i As Long, j As Long
    Dim lng�׶�ID As Long
    Dim blnAdjust As Boolean

    With vsPath
        If TypeName(.ColData(.Col)) = "Empty" Then
            MsgBox "�������õ�ǰ��Ŀλ������Ӧ��ʱ��׶Ρ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If Trim(.TextMatrix(.Row, .FixedCols)) = "" Then
            MsgBox "�������õ�ǰ��Ŀλ������Ӧ�ķ��ࡣ", vbInformation, gstrSysName
            Exit Sub
        End If
        vStep = .ColData(.Col)

        If TypeName(.Cell(flexcpData, .Row, .Col)) <> "Empty" Then
            vItem = .Cell(flexcpData, .Row, .Col)
            vBakItem = vItem
        End If

        '��ȡǰһ��ʱ��׶���ͬ��Ŀ������(������֧)
        For i = .Col - 1 To .FixedCols + .FrozenCols Step -1
            If TypeName(.ColData(i)) <> "Empty" Then
                vPreStep = .ColData(i)
                If IIf(vPreStep.��ID <> 0, vPreStep.��ID, vPreStep.ID) <> IIf(vStep.��ID <> 0, vStep.��ID, vStep.ID) Then '���ǵ�ǰ�׶ε�
                    If lng�׶�ID = 0 Then lng�׶�ID = IIf(vPreStep.��ID <> 0, vPreStep.��ID, vPreStep.ID)
                    If IIf(vPreStep.��ID <> 0, vPreStep.��ID, vPreStep.ID) = lng�׶�ID Then 'ǰһ���׶����з�֧,ѭ��ȡ��ǰ���֧����
                        For j = .FixedRows + .FrozenRows To .Rows - 1
                            If Trim(.TextMatrix(j, .FixedCols)) = Trim(.TextMatrix(.Row, .FixedCols)) Then '�뵱ǰͬ�����
                                If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                                    vTmpItem = .Cell(flexcpData, j, i)
                                    If vTmpItem.��Ŀ���� = vItem.��Ŀ���� Or vItem.ID = 0 And j = .Row Then
                                        vPreItem = vTmpItem: Exit For
                                    End If
                                End If
                            End If
                        Next
                    Else
                        Exit For
                    End If
                End If
            Else
                Exit For 'ֻȡǰ�������ʱ��׶�,û�����˳�
            End If
        Next

        If Not objControl Is Nothing Then
            If objControl.Parameter = "Adjust" Then blnAdjust = True
        End If
        If mfrmPathItem.ShowEdit(Me, mrsAdvice, vItem, vPreItem, blnAdjust, mlng·��ID, mstrPrivs) Then
            If vItem.ID = 0 Then
                '��֤�����ݵ���ĿID��Ϊ�գ���Ԥȡһ��ID
                vItem.ID = zlDatabase.GetNextId("����·����Ŀ")
                vItem.Edit = 1 '0-ԭʼ,1-����,2-�޸�
                '��Ŀ��ű���ǰ�Զ�����
            Else
                If vItem.Edit = 0 Then vItem.Edit = 2
            End If

            If vItem.������ = 1 Then
                .Cell(flexcpBackColor, .Row, .Col) = &H80000005
            End If
            '�������������Ԫ����Ŀ������ͬ��Ϊ�˷�ֹ�Զ��ϲ�����һ���ո�
            If .Row > 1 Then
                If TypeName(.Cell(flexcpData, .Row - 1, .Col)) <> "Empty" Then
                    vTmpItem = .Cell(flexcpData, .Row - 1, .Col)
                    If vTmpItem.��Ŀ���� = vItem.��Ŀ���� Then
                        vItem.��Ŀ���� = vItem.��Ŀ���� & " "
                        .Cell(flexcpData, .Row, .Col) = vItem
                    End If
                End If
            End If
            If .Row < .Rows - 1 Then
                If TypeName(.Cell(flexcpData, .Row + 1, .Col)) <> "Empty" Then
                    vTmpItem = .Cell(flexcpData, .Row + 1, .Col)
                    If vTmpItem.��Ŀ���� = vItem.��Ŀ���� Then
                        vItem.��Ŀ���� = vItem.��Ŀ���� & " "
                        .Cell(flexcpData, .Row, .Col) = vItem
                    End If
                End If
            End If

            '��ǰ��Ԫ��ʾ����
            If vItem.ͼ��ID <> 0 Then
                Set .Cell(flexcpPicture, .Row, .Col) = GetPathIcon(vItem.ͼ��ID)
                .Cell(flexcpPictureAlignment, .Row, .Col) = 1
            Else
                Set .Cell(flexcpPicture, .Row, .Col) = Nothing
            End If
            .TextMatrix(.Row, .Col) = vItem.��Ŀ����
            If vItem.ҽ��IDs <> "" Or vItem.����IDs <> "" Or vItem.�°没��IDs <> "" Then
                .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col) & "��"
            End If
            .Cell(flexcpData, .Row, .Col) = vItem

            '��������
            .AutoSize .FixedCols, .Cols - 1, , 45
            Call SetTableCommonStyle(True)

            mblnChange = True
        End If
    End With
End Sub

Private Sub FuncItemInsert(ByVal intPos As Integer)
'���ܣ������µ���Ŀ
'������inPos=1���ڵ�ǰ��Ŀ���棬-1���ڵ�ǰ��Ŀǰ��
    Dim lngRow As Long, strCategory As String

    With vsPath
        If .TextMatrix(.Row, .Col) = "" Then
            MsgBox "��ǰ��Ŀ��δ���ã��������õ�ǰ��Ŀ���ݡ�", vbInformation, gstrSysName
            Exit Sub
        End If

        strCategory = .TextMatrix(.Row, .FixedCols)
        lngRow = IIf(intPos = -1, .Row, .Row + 1)
        .AddItem "", lngRow
        .TextMatrix(lngRow, .FixedCols) = strCategory
        .RowHeight(lngRow) = ROW_HEIGHT_MIN

        .Row = lngRow
        .ShowCell .Row, .Col

        '��������
        Call FuncItemEdit
    End With

    mblnChange = True
End Sub

Private Sub FuncItemDelete(Optional ByVal blnIsMsg As Boolean = True)
'���ܣ�ɾ����ǰѡ�����Ŀ
'������blnIsMsg-�Ƿ񵯳�ȷ����Ϣ
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim lngRow As Long, lngCol As Long
    Dim i As Long, j As Long, k As Long
    Dim vItem As TYPE_PATH_ITEM
    Dim vStep As TYPE_PATH_STEP
    Dim vEvalCond As TYPE_PATH_EvalCond

    With vsPath
        If blnIsMsg Then
            If MsgBox("ȷʵҪɾ����ѡ���·����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        .Redraw = flexRDNone
        lngRow = .Row: lngCol = .Col

        '��¼ɾ������:0-ԭʼ,1-����,2-�޸�
        .GetSelection lngR1, lngC1, lngR2, lngC2
        For i = lngC1 To lngC2
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                For j = lngR1 To lngR2
                    If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                        vItem = .Cell(flexcpData, j, i)
                        If vItem.Edit <> 1 Then
                            mstrDelItemIDs = mstrDelItemIDs & "," & vItem.ID
                            'ɾ���׶�������ʹ�õ���Ŀָ��
                            If Not vStep.����.������ Is Nothing Then
                                For k = vStep.����.������.count To 1 Step -1
                                    vEvalCond = vStep.����.������(k)
                                    If vEvalCond.��ĿID = vItem.ID Then
                                        vStep.����.������.Remove k
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
                .ColData(i) = vStep
            End If
        Next

        'ɾ��ѡ������
        .GetSelection lngR1, lngC1, lngR2, lngC2
        .Cell(flexcpData, lngR1, lngC1, lngR2, lngC2) = Empty
        .Cell(flexcpText, lngR1, lngC1, lngR2, lngC2) = ""
        Set .Cell(flexcpPicture, lngR1, lngC1, lngR2, lngC2) = Nothing

        '���û��������Ŀ�Ķ���ķ�����
        Call ClearCategoryRow

        '��������
        .Redraw = flexRDDirect
        .AutoSize .FixedCols, .Cols - 1, , 45 'Redraw����Ч
        Call SetTableCommonStyle(True)

        '��λ����
        .Row = IIf(lngRow <= .Rows - 1, lngRow, .Rows - 1): .RowSel = .Row
        .Col = IIf(lngCol <= .Cols - 1, lngCol, .Cols - 1): .ColSel = .Col
        Call .ShowCell(.Row, .Col)
    End With

    mblnChange = True
End Sub

Private Sub FuncVersionDelete()
'���ܣ�ɾ����ǰ�汾
    Dim objCombo As CommandBarComboBox
    Dim strSql As String

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    If MsgBox("ȷʵҪɾ����ǰ�汾�������ٴ�·����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    strSql = "Zl_����·���汾_Delete(" & mlng·��ID & "," & objCombo.ItemData(objCombo.ListIndex) & ")"

    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0

    Call LoadPathVersion

    '���ݱ仯
    RaiseEvent DataChanged(mlng·��ID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncVersionCopy()
'���ܣ�������·�����Ƹ��ǵ�ǰ�汾
    Dim rsTmp As ADODB.Recordset
    Dim objCombo As CommandBarComboBox
    Dim intVersion As Integer, i As Long
    Dim strSql As String, blnCancel As Boolean
    Dim lngԴ·��ID As Long

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    If mblnChange Then
        MsgBox "·���������ѱ�������δ���棬�����ȱ�����ٵ��롣", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error GoTo errH

    'ѡ�����������ٴ�·��
    strSql = "Select ID,����,����,����,���°汾,Decode(Nvl(�����Ա�,0),0,'',1,'��',2,'Ů') as �����Ա�,��������,˵��" & _
            " From ����·��Ŀ¼ A Where Nvl(���°汾,0)>0 And ID<>[1] "
    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        strSql = strSql & " And ͨ��=2 And Not Exists(" & _
                " Select ����ID From ����·������ Where ·��ID=A.ID" & _
                " Minus Select ����ID From ������Ա Where ��ԱID=[2])"
    End If
    strSql = strSql & " Order by ����,����"

    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "�����ٴ�·��", False, "", "", _
        False, False, False, 0, 0, 0, blnCancel, False, False, mlng·��ID, UserInfo.ID, objCombo.ItemData(objCombo.ListIndex))
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û���ҵ��������õ�·����", vbInformation, gstrSysName
        End If
        Exit Sub
    End If

    lngԴ·��ID = rsTmp!ID

    intVersion = objCombo.ItemData(objCombo.ListIndex)

    mstrDelStepIDs = "": mstrDelItemIDs = "": mblnChange = False

    '����ָ��·�����°汾���ǵ�ǰ�汾����
    strSql = "Zl_����·���汾_Copy(" & lngԴ·��ID & ",0," & mlng·��ID & "," & intVersion & ")"

    '�ύ����
    zlDatabase.ExecuteProcedure strSql, Me.Caption

    'ˢ�½���
    Call LoadPathVersion

    '���ݱ仯
    RaiseEvent DataChanged(mlng·��ID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncVersionNew()
'���ܣ������汾
    Dim vVersion As TYPE_PATH_VERSION
    Dim objCombo As CommandBarComboBox
    Dim intVersion As Integer, strSql As String
    Dim i As Long

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    If mblnChange Then
        If MsgBox("·���������ѱ�������δ���棬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    intVersion = objCombo.ItemData(objCombo.ListIndex)
    If MsgBox("Ҫ���Ƶ�ǰѡ��汾�����ݲ����°汾��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then intVersion = 0

    mstrDelStepIDs = "": mstrDelItemIDs = "": mblnChange = False

    If intVersion > 0 Then
        '���Ƶ�ǰѡ��汾���ݲ����°汾����
        strSql = "Zl_����·���汾_Copy(" & mlng·��ID & "," & intVersion & "," & mlng·��ID & ",0)"

        '�ύ����
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        On Error GoTo 0

        'ˢ�½���
        Call LoadPathVersion

        '���ݱ仯
        RaiseEvent DataChanged(mlng·��ID)
    Else
        '���ӿյ�������
        objCombo.AddItem "��������С���", 1
        objCombo.ListIndex = 1
        mcolVersion.Add vVersion, "_0"
        Call cbsMain_Execute(objCombo)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncVersionAudit(ByVal blnAudit As Boolean)
'���ܣ����/ȡ����˵�ǰ�汾
'������blnAudit=���/ȡ�����
    Dim objCombo As CommandBarComboBox
    Dim strSql As String

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    If Not blnAudit Then
        If MsgBox("ȷʵҪȡ����˵�ǰ�汾�������ٴ�·����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    strSql = "Zl_����·���汾_Audit(" & mlng·��ID & "," & objCombo.ItemData(objCombo.ListIndex) & "," & IIf(blnAudit, 1, -1) & ")"

    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0

    Call LoadPathVersion(objCombo.ItemData(objCombo.ListIndex))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncVersionEdit()
'���ܣ����õ�ǰ�汾�����Ϣ
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vStep As TYPE_PATH_STEP
    Dim i As Long, j As Long
    Dim str���� As String
    Dim strԭʼ As String

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    If vVersion.��׼����ʱ�� = "" Then
        With vsPath
            For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vStep.��ʼ���� <> 0 Then
                        If vStep.�������� <> 0 Then
                            vVersion.��׼����ʱ�� = vStep.��������
                        Else
                            vVersion.��׼����ʱ�� = vStep.��ʼ����
                        End If
                        Exit For
                    End If
                End If
            Next
        End With
    End If

    If mfrmVersion.ShowMe(Me, vVersion, mlng·��ID) Then
        mcolVersion.Remove "_" & objCombo.ItemData(objCombo.ListIndex)
        mcolVersion.Add vVersion, "_" & objCombo.ItemData(objCombo.ListIndex)

        stcInfo.Caption = "��׼����ʱ�䣺" & IIf(vVersion.��׼����ʱ�� <> "", vVersion.��׼����ʱ�� & "��", "<δ�趨>") & _
                          "����׼���ã�" & IIf(vVersion.��׼���� <> "", vVersion.��׼���� & "Ԫ", "<δ�趨>") & _
                          "��˵����" & IIf(vVersion.�汾˵�� <> "", vVersion.�汾˵��, "<��>")
        mblnChange = True
    End If
End Sub

Private Sub FuncVersionStop(ByVal blnStop As Boolean)
'���ܣ�ͣ��/ȡ��ͣ�õ�ǰ�汾
'������blnAudit=���/ȡ�����
    Dim objCombo As CommandBarComboBox
    Dim strSql As String

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    If blnStop Then
        If MsgBox("ȷʵҪͣ�õ�ǰ�汾�������ٴ�·����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        If MsgBox("ȷʵҪȡ��ͣ�õ�ǰ�汾�������ٴ�·����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    strSql = "Zl_����·���汾_Stop(" & mlng·��ID & "," & objCombo.ItemData(objCombo.ListIndex) & "," & IIf(blnStop, 1, -1) & ")"

    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0

    Call LoadPathVersion(objCombo.ItemData(objCombo.ListIndex))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncEvaluateImport()
'���ܣ����õ�������
    If mfrmEvalEdit.ShowEdit(Me, 1, mvEvalImport) Then
        mblnChange = True
    End If
End Sub

Private Sub FuncEvaluateStep(Optional ByVal blnCopy As Boolean)
'���ܣ����ý׶�����
    Dim vStep As TYPE_PATH_STEP
    Dim vEval As TYPE_PATH_EVAL
    Dim vEvalPre As TYPE_PATH_EVAL
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim colMarkID As New Collection
    Dim vItem As TYPE_PATH_ITEM
    Dim colItems As New Collection
    Dim lngC1 As Long, lngC2 As Long
    Dim lngNewId As Long, i As Long, j As Long

    With vsPath
        If .Col >= .FixedCols + .FrozenCols Then
            If .Row = .FixedRows Then
                .GetMergedRange .Row, .Col, 0, lngC1, 0, lngC2
                If lngC1 <> lngC2 Then
                    .Row = .FixedRows + .FrozenRows - 1
                End If
            End If
        End If

        If TypeName(.ColData(.Col)) = "Empty" Then
            MsgBox "�������õ�ǰʱ��׶ε���Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If

        If blnCopy Then
            For i = .Col - 1 To .FixedCols + .FrozenCols Step -1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If Not vStep.����.ָ�꼯 Is Nothing Then
                        If vStep.����.ָ�꼯.count > 0 Then
                            vEvalPre = vStep.����
                            Set vEval.ָ�꼯 = New Collection
                            Set vEval.������ = New Collection

                            '�ռ�ָ��
                            For j = 1 To vEvalPre.ָ�꼯.count
                                vEvalMark = vEvalPre.ָ�꼯(j)
                                
                                lngNewId = zlDatabase.GetNextId("����·������ָ��")
                                colMarkID.Add lngNewId, "_" & vEvalMark.ID
                                
                                vEvalMark.ID = lngNewId
                                vEval.ָ�꼯.Add vEvalMark
                            Next

                            '�ռ���������
                            If Not vEvalPre.������ Is Nothing Then
                                For j = 1 To vEvalPre.������.count
                                    vEvalCond = vEvalPre.������(j)
                                    If vEvalCond.ָ��ID <> 0 Then
                                        vEvalCond.ָ��ID = colMarkID("_" & vEvalCond.ָ��ID)
                                        vEval.������.Add vEvalCond
                                    End If
                                Next
                            End If

                            Exit For
                        End If
                    End If
                End If
            Next
            If vEval.ָ�꼯 Is Nothing And vEval.������ Is Nothing Then
                MsgBox "ǰ���ʱ��׶���û�п��Ը��Ƶ��������á�", vbInformation, gstrSysName
                Exit Sub
            End If
            vStep = .ColData(.Col)
        Else
            vStep = .ColData(.Col)
            vEval = vStep.����
        End If

        '���׶ε���Ŀ(������Ϊ����ָ��)
        For i = .FixedRows + .FrozenRows To .Rows - 1
            If TypeName(.Cell(flexcpData, i, .Col)) <> "Empty" Then
                vItem = .Cell(flexcpData, i, .Col)
                colItems.Add vItem
            End If
        Next
    End With

    If mfrmEvalEdit.ShowEdit(Me, 2, vEval, colItems) Then
        With vsPath
            vStep.���� = vEval
            '0-ԭʼ,1-����,2-�޸�
            If vStep.Edit = 0 Then vStep.Edit = 2
            .ColData(.Col) = vStep

            .TextMatrix(.FixedRows - 1, .Col) = "�׶�������"
            .Cell(flexcpFontBold, .FixedRows - 1, .Col) = False
            If vStep.����.ָ�꼯.count > 0 Or vStep.����.������.count > 0 Then
                .Cell(flexcpFontBold, .FixedRows - 1, .Col) = True
            End If
        End With
        mblnChange = True
    End If
End Sub

Private Sub FuncClipboradCopy()
'���ܣ����Ƶ�ǰѡ�����Ŀ��Ϣ���ڲ�������
'˵����ֻ�ܶ�ͬһ�׶��е�һ��������Ŀ���и���
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vNullItem As TYPE_PATH_ITEM
    Dim lngR1 As Long, lngR2 As Long
    Dim lngC1 As Long, lngC2 As Long
    Dim i As Long

    With vsPath
        .GetSelection lngR1, lngC1, lngR2, lngC2
        If lngC1 <> lngC2 Then
            MsgBox "û�����ݱ����ơ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If TypeName(.ColData(lngC1)) = "Empty" Then
            MsgBox "û�����ݱ����ơ�", vbInformation, gstrSysName
            Exit Sub
        End If
        vStep = .ColData(lngC1)

        ReDim mvClipboard.��Ŀ��(lngR2 - lngR1)
        For i = lngR1 To lngR2
            If TypeName(.Cell(flexcpData, i, lngC1)) <> "Empty" Then
                vItem = .Cell(flexcpData, i, lngC1)
            Else
                vItem = vNullItem
            End If
            mvClipboard.��Ŀ��(i - lngR1) = vItem
        Next
        mvClipboard.Empty = False
        mvClipboard.vStep = vStep
        mvClipboard.BeginRow = lngR1
    End With
End Sub

Private Sub FuncClipboradPaste()
'���ܣ����ڲ�������ճ�����ݵ���ǰѡ������
'˵����ֻ�ܶ�ͬһ�׶��е�һ��������Ŀ���и���
    Dim vItem1 As TYPE_PATH_ITEM
    Dim vItem2 As TYPE_PATH_ITEM
    Dim vNullItem As TYPE_PATH_ITEM
    Dim lngR1 As Long, lngC1 As Long
    Dim lngR2 As Long, lngC2 As Long
    Dim i As Long
    Dim vStep As TYPE_PATH_STEP

    If mvClipboard.Empty Then
        MsgBox "�������ǿյġ�", vbInformation, gstrSysName
        Exit Sub
    End If
    With vsPath
        .GetSelection lngR1, lngC1, lngR2, lngC2

        If lngC2 <> lngC1 Then
            MsgBox "Ҫճ�����ݵ�Ŀ��ѡ�����򲻷���Ҫ��ֻ�ܶ�һ��ʱ��׶��е���Ŀ���и���ճ����", vbInformation, gstrSysName
            Exit Sub
        End If
        If TypeName(.ColData(lngC1)) = "Empty" Then
            MsgBox "�������õ�ǰ��Ŀλ������Ӧ��ʱ��׶Ρ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If .Rows - lngR1 < UBound(mvClipboard.��Ŀ��) + 1 Then
            MsgBox "Ŀ������̫С��������ճ�������Ƶ�Դ��Ŀ���ݡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("ȷʵҪճ�������Ƶ���Ŀ���ݸ��ǵ�ǰĿ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

        vStep = .ColData(lngC1)
        'ճ������
        .Redraw = flexRDNone
        For i = 0 To UBound(mvClipboard.��Ŀ��)
            vItem1 = mvClipboard.��Ŀ��(i)
            If TypeName(.Cell(flexcpData, lngR1, lngC1)) <> "Empty" Then
                vItem2 = .Cell(flexcpData, lngR1, lngC1)
            Else
                vItem2 = vNullItem
            End If

            'Edit��0-ԭʼ,1-����,2-�޸�
            If vItem1.ID <> 0 Then
                If vItem2.ID <> 0 Then
                    vItem1.ID = vItem2.ID
                    vItem1.Edit = vItem2.Edit
                    If vItem1.Edit = 0 Then vItem1.Edit = 2
                Else
                    vItem1.ID = zlDatabase.GetNextId("����·����Ŀ")
                    vItem1.Edit = 1
                End If

                '����ж�Ӧҽ��������Ϊ��������ҽ��
                If vItem1.ҽ��IDs <> "" Then
                    vItem1.ҽ��IDs = AdviceCopyNew(vItem1.ҽ��IDs)
                End If

                .Cell(flexcpData, lngR1, lngC1) = vItem1

                .TextMatrix(lngR1, lngC1) = vItem1.��Ŀ����
                If vItem1.ҽ��IDs <> "" Or vItem1.����IDs <> "" Or vItem1.�°没��IDs <> "" Then
                    .TextMatrix(lngR1, lngC1) = .TextMatrix(lngR1, lngC1) & "��"
                End If

                If vItem1.ͼ��ID <> 0 Then
                    Set .Cell(flexcpPicture, lngR1, lngC1) = GetPathIcon(vItem1.ͼ��ID)
                    .Cell(flexcpPictureAlignment, lngR1, lngC1) = 1
                End If
            Else
                .Cell(flexcpData, lngR1, lngC1) = Empty
                .TextMatrix(lngR1, lngC1) = ""
                Set .Cell(flexcpPicture, lngR1, lngC1) = Nothing

                '��¼ɾ������
                If vItem2.ID <> 0 And vItem2.Edit <> 1 Then
                    mstrDelItemIDs = mstrDelItemIDs & "," & vItem2.ID
                End If
            End If

            lngR1 = lngR1 + 1
        Next
        .GetSelection lngR1, lngC1, lngR2, lngC2
        .Select lngR1, lngC1, lngR1 + UBound(mvClipboard.��Ŀ��), lngC2
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect

        '��������
        .AutoSize .FixedCols, .Cols - 1, , 45 'Redraw����Ч
        Call SetTableCommonStyle(True)

        mblnChange = True
    End With
End Sub

Private Function AdviceCopyNew(ByVal strҽ��ID As String) As String
'���ܣ�����ҽ��ID���Ʋ����µ�ҽ��
    Dim rsCopy As ADODB.Recordset
    Dim strFilter As String, i As Long, arrAdvice As Variant
    Dim colAdviceID As New Collection
    Dim lngAdviceID As Long, strAdviceID As String
    Dim strSql As String
    Dim objCombo As CommandBarComboBox

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Function
    If objCombo.ListIndex = 0 Then Exit Function
    Set rsCopy = mrsAdvice.Clone

    arrAdvice = Split(strҽ��ID, ",")
    For i = 0 To UBound(arrAdvice)
        strFilter = strFilter & " Or ID=" & arrAdvice(i)
    Next
    rsCopy.Filter = Mid(strFilter, 5)

    If rsCopy.RecordCount = 0 Then
        '�������ʱû�м�¼�������������
        strSql = " Select /*+ Rule*/ Distinct A.ID,A.���ID,A.���,A.��Ч,A.������ĿID,A.�շ�ϸĿID," & _
                " A.ҽ������,A.��������,A.�ܸ�����,A.�걾��λ,A.��鷽��,A.ҽ������,A.ִ�б��, " & _
                " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ������,A.ִ�п���ID,A.ʱ�䷽��,A.�Ƿ�ȱʡ,A.�Ƿ�ѡ,A.�䷽ID,A.�����ĿID" & _
                " From ����·��ҽ������ A,����·��ҽ�� B,����·����Ŀ C" & _
                " Where A.ID=B.ҽ������ID And B.·����ĿID=C.ID And C.·��ID=[1] And C.�汾��=[2] And a.ID In (Select * From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))) " & _
                " Order by A.���,A.ID"
        On Error GoTo errH
        Set rsCopy = zlDatabase.OpenSQLRecord(strSql, "������Ŀ", mlng·��ID, objCombo.ItemData(objCombo.ListIndex), strҽ��ID)
    End If
    If rsCopy.RecordCount = 0 Then Exit Function

    '�Ȳ����µ�ҽ��ID
    Do While Not rsCopy.EOF
        lngAdviceID = zlDatabase.GetNextId("����·��ҽ������")
        colAdviceID.Add lngAdviceID, "_" & rsCopy!ID
        strAdviceID = strAdviceID & "," & lngAdviceID
        rsCopy.MoveNext
    Loop

    rsCopy.MoveFirst: i = 1
    Do While Not rsCopy.EOF
        lngAdviceID = colAdviceID("_" & rsCopy!ID)
        mrsAdvice.AddNew
        mrsAdvice!ID = lngAdviceID
        If Not IsNull(rsCopy!���id) Then
            mrsAdvice!���id = colAdviceID("_" & rsCopy!���id)
        End If
        mrsAdvice!��� = i
        mrsAdvice!��Ч = rsCopy!��Ч
        mrsAdvice!������ĿID = rsCopy!������ĿID
        mrsAdvice!�շ�ϸĿID = rsCopy!�շ�ϸĿID
        If IsNull(rsCopy!������ĿID) Then
            mrsAdvice!ҽ������ = rsCopy!ҽ������ '����¼��ҽ���ű���
        End If
        mrsAdvice!�������� = rsCopy!��������
        mrsAdvice!�ܸ����� = rsCopy!�ܸ�����
        mrsAdvice!ҽ������ = rsCopy!ҽ������
        mrsAdvice!ִ��Ƶ�� = rsCopy!ִ��Ƶ��
        mrsAdvice!Ƶ�ʴ��� = rsCopy!Ƶ�ʴ���
        mrsAdvice!Ƶ�ʼ�� = rsCopy!Ƶ�ʼ��
        mrsAdvice!�����λ = rsCopy!�����λ
        mrsAdvice!ʱ�䷽�� = rsCopy!ʱ�䷽��
        mrsAdvice!ִ�п���ID = rsCopy!ִ�п���ID
        mrsAdvice!ִ������ = rsCopy!ִ������
        mrsAdvice!�걾��λ = rsCopy!�걾��λ
        mrsAdvice!��鷽�� = rsCopy!��鷽��
        mrsAdvice!�䷽ID = rsCopy!�䷽ID
        mrsAdvice!�����ĿID = rsCopy!�����ĿID
        mrsAdvice!ִ�б�� = rsCopy!ִ�б��

        mrsAdvice.Update

        i = i + 1
        rsCopy.MoveNext
    Loop

    AdviceCopyNew = Mid(strAdviceID, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearCategoryRow()
'���ܣ����û��������Ŀ�Ķ���ķ�����
    Dim lngRow As Long
    Dim lngR1 As Long, lngR2 As Long
    Dim i As Long, j As Long
    Dim vRedraw As RedrawSettings

    With vsPath
        vRedraw = .Redraw: .Redraw = flexRDNone
        lngRow = .Row
        i = .Rows - 1
        Do While i >= .FixedRows + .FrozenRows
            .GetMergedRange i, .FixedCols, lngR1, 0, lngR2, 0
            If Replace(Replace(.Cell(flexcpText, lngR1, .FixedCols, lngR2, .FixedCols), vbTab, ""), vbCr, "") <> "" Then
                For j = lngR2 To lngR1 Step -1
                    If Replace(.Cell(flexcpText, j, .FixedCols, j, .FixedCols), vbTab, "") = "" Then
                        .RemoveItem j
                    End If
                Next
            End If

            i = lngR1 - 1
        Loop
        .Row = IIf(lngRow <= .Rows - 1, lngRow, .Rows - 1)
        .ShowCell .Row, .Col
        .Redraw = vRedraw
    End With
End Sub

Private Function CheckPathTable() As Boolean
'���ܣ����·������������ĺϷ���
    Dim lngR1 As Long, lngR2 As Long
    Dim i As Long, j As Long
    Dim strMsg As String
    Dim strPathItems As String
    Dim strAdviceIDs As String
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim lng�׶���� As Long
    Dim lng��֧��� As Long
    Dim lng��Ŀ��� As Long
    Dim strSql As String, rsTmp As Recordset

    With vsPath
        'û�����õĽ׶�
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) = "Empty" Then
                .Row = .FixedRows: .Col = i
                Call .ShowCell(.Row, .Col)
                MsgBox "�ý׶ε���Ϣ��δ�������á�", vbInformation, gstrSysName
                Exit Function
            End If
        Next

        '�������õķ���
        For i = .FixedRows + .FrozenRows To .Rows - 1
            If Trim(.TextMatrix(i, .FixedCols)) = "" Then
                .Row = i: .Col = .FixedCols
                Call .ShowCell(.Row, .Col)
                MsgBox "�÷����������δ���롣", vbInformation, gstrSysName
                Exit Function
            End If
        Next

        'û��������Ŀ�Ľ׶λ��߷���(����)
        strMsg = ""
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                If Replace(.Cell(flexcpText, .FixedRows + .FrozenRows, i, .Rows - 1, i), vbCr, "") = "" Then
                    strMsg = "���ִ�����δ����·����Ŀ�Ľ׶λ��߷��࣬Ҫ������"
                    Exit For
                End If
            End If
        Next
        i = .FixedRows + .FrozenRows
        Do While i <= .Rows - 1
            .GetMergedRange i, .FixedCols, lngR1, 0, lngR2, 0
            If Replace(Replace(.Cell(flexcpText, lngR1, .FixedCols + .FrozenCols, lngR2, .Cols - 1), vbTab, ""), vbCr, "") = "" Then
                strMsg = "���ִ�����δ����·����Ŀ�Ľ׶λ��߷��࣬Ҫ������"
                Exit Do
            End If
            i = lngR2 + 1
        Loop
        If strMsg <> "" Then
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If

        '��׼����ʱ��Ӧ�����н׶ε�����ƥ��
        Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
        If objCombo Is Nothing Then
            MsgBox "�����ٴ�·����ĵ�ǰ�汾��Ϣ��ȡʧ�ܡ�", vbInformation, gstrSysName: Exit Function
        End If
        If objCombo.ListIndex = 0 Then
            MsgBox "�����ٴ�·����ĵ�ǰ�汾��Ϣ��ȡʧ�ܡ�", vbInformation, gstrSysName: Exit Function
        End If
        vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))

        If vVersion.��׼����ʱ�� = "" Then
            MsgBox "��û�����õ�ǰ�汾�ı�׼����ʱ����Ϣ��", vbInformation, gstrSysName: Exit Function
        End If

        For i = .Cols - 1 To .FixedCols + .FrozenCols Step -1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.�������� <> 0 Or vStep.��ʼ���� <> 0 Then
                    If InStr(vVersion.��׼����ʱ��, "-") > 0 Then
                        If vStep.�������� <> 0 Then
                            If Val(Split(vVersion.��׼����ʱ��, "-")(1)) <> vStep.�������� Then
                                MsgBox "��׼����ʱ���������� " & Val(Split(vVersion.��׼����ʱ��, "-")(1)) & " ����ʱ��׶���ָ����������� " & vStep.�������� & " �첻����", vbInformation, gstrSysName
                                Exit Function
                            End If
                        ElseIf vStep.��ʼ���� <> 0 Then
                            If Val(Split(vVersion.��׼����ʱ��, "-")(1)) <> vStep.��ʼ���� Then
                                MsgBox "��׼����ʱ���������� " & Val(Split(vVersion.��׼����ʱ��, "-")(1)) & " ����ʱ��׶���ָ����������� " & vStep.��ʼ���� & " �첻����", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    Else
                        If vStep.�������� <> 0 Then
                            If Val(vVersion.��׼����ʱ��) <> vStep.�������� Then
                                MsgBox "��׼����ʱ�� " & vVersion.��׼����ʱ�� & " ����ʱ��׶���ָ����������� " & vStep.�������� & " �첻����", vbInformation, gstrSysName
                                Exit Function
                            End If
                        ElseIf vStep.��ʼ���� <> 0 Then
                            If Val(vVersion.��׼����ʱ��) <> vStep.��ʼ���� Then
                                MsgBox "��׼����ʱ�� " & vVersion.��׼����ʱ�� & " ����ʱ��׶���ָ����������� " & vStep.��ʼ���� & " �첻����", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                    Exit For
                End If
            End If
        Next

        '���׶��е���Ŀ�����ظ�
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                strMsg = "": strPathItems = ""
                For j = .FixedRows + .FrozenRows To .Rows - 1
                    If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                        vItem = .Cell(flexcpData, j, i)
                        If InStr(strPathItems & vbTab, vbTab & Trim(vItem.��Ŀ����) & vbTab) = 0 Then
                            strPathItems = strPathItems & vbTab & Trim(vItem.��Ŀ����)
                            .Cell(flexcpFontBold, j, i) = False
                        Else
                            .Cell(flexcpFontBold, j, i) = True
                            strMsg = Trim(vItem.��Ŀ����)
                        End If
                    End If
                Next
                If strMsg <> "" Then
                    '�ҵ���һ��
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            If .Cell(flexcpData, j, i).��Ŀ���� = strMsg Then
                                .Col = i: .Row = j: .ShowCell .Row, .Col
                                .Cell(flexcpFontBold, j, i) = True
                                Exit For
                            End If
                        End If
                    Next
                    If .FrozenRows > 1 And .TextMatrix(.FixedRows, i) <> .TextMatrix(.FixedRows + .FrozenRows - 1, i) Then
                        strMsg = "�׶�""" & Replace(vStep.����, vbLf, "") & "(" & .TextMatrix(.FixedRows + .FrozenRows - 1, i) & ")""�е�·����Ŀ""" & strMsg & """�ظ������顣"
                    Else
                        strMsg = "�׶�""" & Replace(vStep.����, vbLf, "") & """�е�·����Ŀ""" & strMsg & """�ظ������顣"
                    End If

                    MsgBox strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next

        '���û��������Ŀ�Ķ���ķ�����
        Call ClearCategoryRow

        '���ý׶κ���Ŀ�����
        lng�׶���� = 1
        For i = .FixedCols + .FrozenCols To .Cols - 1
            '�׶����
            If TypeName(.ColData(i)) <> "Empty" Then
                vStep = .ColData(i)
                If vStep.��ID = 0 Then
                    If vStep.��� <> lng�׶���� Then
                        vStep.��� = lng�׶����
                        If vStep.Edit = 0 Then vStep.Edit = 2   '0-ԭʼ,1-����,2-�޸�
                    End If
                    lng�׶���� = lng�׶���� + 1
                    lng��֧��� = 1
                Else
                    If vStep.��� <> lng��֧��� Then
                        vStep.��� = lng��֧���
                        If vStep.Edit = 0 Then vStep.Edit = 2   '0-ԭʼ,1-����,2-�޸�
                    End If
                    lng��֧��� = lng��֧��� + 1
                End If
                .ColData(i) = vStep
            End If

            '��Ŀ���
            If TypeName(.ColData(i)) <> "Empty" Then
                lngR1 = .FixedRows + .FrozenRows
                Do While lngR1 <= .Rows - 1
                    .GetMergedRange lngR1, .FixedCols, lngR1, 0, lngR2, 0

                    lng��Ŀ��� = 1
                    For j = lngR1 To lngR2
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)

                            If vItem.��Ŀ��� <> lng��Ŀ��� Then
                                vItem.��Ŀ��� = lng��Ŀ���
                                If vItem.Edit = 0 Then vItem.Edit = 2 '0-ԭʼ,1-����,2-�޸�
                            End If

                            .Cell(flexcpData, j, i) = vItem
                            lng��Ŀ��� = lng��Ŀ��� + 1
                        End If
                    Next

                    lngR1 = lngR2 + 1
                Loop
            End If
        Next

        '�����û��ʹ�õ�ҽ������ID
        strAdviceIDs = "": mstrChangeItemIDs = ""
        For i = .FixedCols + .FrozenCols To .Cols - 1
            For j = .FixedRows + .FrozenRows To .Rows - 1
                If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                    vItem = .Cell(flexcpData, j, i)
                    If vItem.ҽ��IDs <> "" Then
                        strAdviceIDs = strAdviceIDs & "," & vItem.ҽ��IDs
                    End If
                    If vItem.�����ҽ��IDs <> "" Then
                        strAdviceIDs = strAdviceIDs & "," & vItem.�����ҽ��IDs
                    End If
                    If (vItem.ԭҽ��IDs <> vItem.ҽ��IDs And vItem.�����ҽ��IDs = "") And vVersion.���ʱ�� <> Empty And vVersion.ͣ��ʱ�� = Empty Then
                        mstrChangeItemIDs = mstrChangeItemIDs & "," & vItem.ID         '��¼�±䶯��ĿID
                    End If
                End If
            Next
        Next
        strAdviceIDs = strAdviceIDs & ","
        mstrChangeItemIDs = Mid(mstrChangeItemIDs, 2)

        mrsAdvice.Filter = ""
        If Not mrsAdvice.EOF Then
            mrsAdvice.MoveFirst
            Do While Not mrsAdvice.EOF
                If InStr(strAdviceIDs, "," & mrsAdvice!ID & ",") = 0 Then
                    mrsAdvice.Delete
                    mrsAdvice.Update
                End If
                mrsAdvice.MoveNext
            Loop
        End If
    End With

    CheckPathTable = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePathTable() As Boolean
'���ܣ�����·��������
    Dim vVersion As TYPE_PATH_VERSION
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim vEvalMark As TYPE_PATH_EvalMark
    Dim vEvalCond As TYPE_PATH_EvalCond
    Dim objCombo As CommandBarComboBox
    Dim arrSQL As Variant, intVersion As Integer
    Dim i As Long, j As Long, k As Long
    Dim blnTrans As Boolean
    Dim strAddDate As String

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Function
    If objCombo.ListIndex = 0 Then Exit Function

    arrSQL = Array()
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))

    With vsPath
        If mblnNewVersion Then
            '�����µ������ٴ�·���汾
            k = 0
            For i = 1 To objCombo.ListCount
                If objCombo.ItemData(i) > k Then k = objCombo.ItemData(i)
            Next
            intVersion = k + 1

            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����·���汾_Update(" & _
                mlng·��ID & "," & intVersion & ",'" & vVersion.��׼����ʱ�� & "','" & vVersion.��׼���� & "','" & vVersion.�汾˵�� & "')"

            '��������
            If Not mvEvalImport.ָ�꼯 Is Nothing Then
                For i = 1 To mvEvalImport.ָ�꼯.count
                    vEvalMark = mvEvalImport.ָ�꼯(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·������ָ��_Insert(" & _
                        mlng·��ID & "," & intVersion & ",NULL,1," & _
                        vEvalMark.ID & "," & vEvalMark.��� & "," & _
                        "'" & vEvalMark.����ָ�� & "'," & vEvalMark.ָ������ & "," & _
                        "'" & vEvalMark.ָ���� & "')"
                Next
            End If
            If Not mvEvalImport.������ Is Nothing Then
                For i = 1 To mvEvalImport.������.count
                    vEvalCond = mvEvalImport.������(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·����������_Insert(" & _
                        mlng·��ID & "," & intVersion & ",NULL,1," & _
                        ZVal(vEvalCond.ָ��ID) & ",NULL," & _
                        "'" & vEvalCond.��ϵʽ & "','" & vEvalCond.����ֵ & "'," & _
                        vEvalCond.������� & ")"
                Next
            End If

            '�׶���Ϣ
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·���׶�_Insert(" & _
                        vStep.ID & "," & mlng·��ID & "," & intVersion & "," & _
                        ZVal(vStep.��ID) & "," & vStep.��� & ",'" & vStep.���� & "'," & _
                        ZVal(vStep.��ʼ����) & "," & ZVal(vStep.��������) & "," & _
                        "'" & vStep.˵�� & "','" & vStep.���� & "')"
                End If
            Next

            '������Ϣ
            k = 1: i = .FixedRows + .FrozenRows
            Do While i <= .Rows - 1
                .GetMergedRange i, .FixedCols, i, 0, j, 0
                If .TextMatrix(i, .FixedCols) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·������_Insert(" & _
                        mlng·��ID & "," & intVersion & "," & k & ",'" & .TextMatrix(i, .FixedCols) & "',Null)"
                    k = k + 1
                End If
                i = j + 1
            Loop

            '��Ŀ��Ӧ��ҽ������
            With mrsAdvice
               .Filter = "" '�Զ�MoveFirst,����Filter��û��
                Do While Not .EOF
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·��ҽ������_Insert(" & _
                        !ID & "," & ZVal(NVL(!���id, 0)) & "," & !��� & "," & !��Ч & "," & _
                        ZVal(NVL(!������ĿID, 0)) & ",'" & NVL(!ҽ������) & "'," & ZVal(NVL(!��������, 0)) & "," & _
                        ZVal(NVL(!�ܸ�����, 0)) & "," & ZVal(NVL(!�շ�ϸĿID, 0)) & ",'" & NVL(!�걾��λ) & "'," & _
                        "'" & NVL(!��鷽��) & "','" & NVL(!ִ��Ƶ��) & "'," & ZVal(NVL(!Ƶ�ʴ���, 0)) & "," & _
                        ZVal(NVL(!Ƶ�ʼ��, 0)) & ",'" & NVL(!�����λ) & "','" & NVL(!ҽ������) & "'," & _
                        NVL(!ִ������, 0) & "," & ZVal(NVL(!ִ�п���ID, 0)) & ",'" & NVL(!ʱ�䷽��) & "',Null,Null," & _
                        !�Ƿ�ȱʡ & "," & !�Ƿ�ѡ & _
                       "," & ZVal(Val(!�䷽ID & "")) & "," & ZVal(Val(!�����ĿID & "")) & "," & NVL(!ִ�б��, 0) & ")"

                   .MoveNext
                Loop
            End With

            '��Ŀ��Ϣ
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)

                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_����·����Ŀ_Insert(" & _
                                vItem.ID & "," & mlng·��ID & "," & intVersion & "," & _
                                vStep.ID & ",'" & .TextMatrix(j, .FixedCols) & "'," & _
                                vItem.��Ŀ��� & ",'" & vItem.��Ŀ���� & "'," & _
                                vItem.ִ�з�ʽ & "," & _
                                "'" & vItem.��Ŀ��� & "'," & ZVal(vItem.ͼ��ID) & "," & _
                                "'" & vItem.ҽ��IDs & "','" & vItem.�������� & "'," & vItem.����Ҫ�� & "," & _
                                "'" & vItem.����ο� & "'," & _
                                IIf(vItem.������ = 1 And Trim(vItem.����ο�) = "", "Null", vItem.������) & ")"
                        End If
                    Next
                End If
            Next

            '�׶��������ͽ׶κ���Ŀ��أ���������
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If Not vStep.����.ָ�꼯 Is Nothing Then
                        For j = 1 To vStep.����.ָ�꼯.count
                            vEvalMark = vStep.����.ָ�꼯(j)
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_����·������ָ��_Insert(" & _
                                mlng·��ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                vEvalMark.ID & "," & vEvalMark.��� & "," & _
                                "'" & vEvalMark.����ָ�� & "'," & vEvalMark.ָ������ & "," & _
                                "'" & vEvalMark.ָ���� & "')"
                        Next
                    End If
                    If Not vStep.����.������ Is Nothing Then
                        For j = 1 To vStep.����.������.count
                            vEvalCond = vStep.����.������(j)
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_����·����������_Insert(" & _
                                mlng·��ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                ZVal(vEvalCond.ָ��ID) & "," & ZVal(vEvalCond.��ĿID) & "," & _
                                "'" & vEvalCond.��ϵʽ & "','" & vEvalCond.����ֵ & "'," & _
                                vEvalCond.������� & ")"
                        Next
                    End If
                End If
            Next
        Else
            '��ԭ·���汾�����ϸ���
            intVersion = objCombo.ItemData(objCombo.ListIndex)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����·���汾_Update(" & _
                mlng·��ID & "," & intVersion & ",'" & vVersion.��׼����ʱ�� & "','" & vVersion.��׼���� & "','" & vVersion.�汾˵�� & "')"

            '��������
            If Not mvEvalImport.ָ�꼯 Is Nothing Then
                For i = 1 To mvEvalImport.ָ�꼯.count
                    vEvalMark = mvEvalImport.ָ�꼯(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·������ָ��_Insert(" & _
                        mlng·��ID & "," & intVersion & ",NULL,1," & _
                        vEvalMark.ID & "," & vEvalMark.��� & "," & _
                        "'" & vEvalMark.����ָ�� & "'," & vEvalMark.ָ������ & "," & _
                        "'" & vEvalMark.ָ���� & "')"
                Next
            End If
            If Not mvEvalImport.������ Is Nothing Then
                For i = 1 To mvEvalImport.������.count
                    vEvalCond = mvEvalImport.������(i)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·����������_Insert(" & _
                        mlng·��ID & "," & intVersion & ",NULL,1," & _
                        ZVal(vEvalCond.ָ��ID) & ",NULL," & _
                        "'" & vEvalCond.��ϵʽ & "','" & vEvalCond.����ֵ & "'," & _
                        vEvalCond.������� & ")"
                Next
            End If

            '�׶���Ϣ
            If mstrDelStepIDs <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����·���׶�_Delete('" & Mid(mstrDelStepIDs, 2) & "')"
            End If

            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vStep.Edit = 1 Then '����
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_����·���׶�_Insert(" & _
                            vStep.ID & "," & mlng·��ID & "," & intVersion & "," & _
                            ZVal(vStep.��ID) & "," & vStep.��� & ",'" & vStep.���� & "'," & _
                            ZVal(vStep.��ʼ����) & "," & ZVal(vStep.��������) & "," & _
                            "'" & vStep.˵�� & "','" & vStep.���� & "')"
                    ElseIf vStep.Edit = 2 Then '�޸�
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_����·���׶�_Update(" & _
                            vStep.ID & "," & mlng·��ID & "," & intVersion & "," & _
                            vStep.��� & ",'" & vStep.���� & "'," & _
                            ZVal(vStep.��ʼ����) & "," & ZVal(vStep.��������) & "," & _
                            "'" & vStep.˵�� & "','" & vStep.���� & "')"
                    End If
                End If
            Next

            '������Ϣ
            k = 1: i = .FixedRows + .FrozenRows
            Do While i <= .Rows - 1
                .GetMergedRange i, .FixedCols, i, 0, j, 0
                If .TextMatrix(i, .FixedCols) <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·������_Insert(" & _
                        mlng·��ID & "," & intVersion & "," & k & ",'" & .TextMatrix(i, .FixedCols) & "'," & _
                        IIf(k = 1, 1, 0) & ")"
                    k = k + 1
                End If
                i = j + 1
            Loop

            strAddDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")

            '���δͣ�õ�·����Ҫ����·��ҽ���䶯��¼(��SQLҪ���� Zl_����·��ҽ������_Insert ִ��)
            If vVersion.���ʱ�� <> Empty And vVersion.ͣ��ʱ�� = Empty Then
                If mstrChangeItemIDs <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����·��ҽ���䶯_Insert('" & mstrChangeItemIDs & "'," & "To_Date('" & strAddDate & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & UserInfo.���� & "')"
                End If
            End If

            '��Ŀ��Ӧ��ҽ������
            With mrsAdvice
                k = 1: .Filter = "" '�Զ�MoveFirst,����Filter��û��
                Do While Not .EOF
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If Val(!����� & "") = 0 Then
                        arrSQL(UBound(arrSQL)) = "Zl_����·��ҽ������_Insert(" & _
                            !ID & "," & ZVal(NVL(!���id, 0)) & "," & !��� & "," & !��Ч & "," & _
                            ZVal(NVL(!������ĿID, 0)) & ",'" & NVL(!ҽ������) & "'," & ZVal(NVL(!��������, 0)) & "," & _
                            ZVal(NVL(!�ܸ�����, 0)) & "," & ZVal(NVL(!�շ�ϸĿID, 0)) & ",'" & NVL(!�걾��λ) & "'," & _
                            "'" & NVL(!��鷽��) & "','" & NVL(!ִ��Ƶ��) & "'," & ZVal(NVL(!Ƶ�ʴ���, 0)) & "," & _
                            ZVal(NVL(!Ƶ�ʼ��, 0)) & ",'" & NVL(!�����λ) & "','" & NVL(!ҽ������) & "'," & _
                            NVL(!ִ������, 0) & "," & ZVal(NVL(!ִ�п���ID, 0)) & ",'" & NVL(!ʱ�䷽��) & "'," & _
                            IIf(k = 1, mlng·��ID, "NULL") & "," & IIf(k = 1, intVersion, "NULL") & "," & _
                            !�Ƿ�ȱʡ & "," & !�Ƿ�ѡ & "," & ZVal(Val(!�䷽ID & "")) & _
                            "," & ZVal(Val(!�����ĿID & "")) & "," & NVL(!ִ�б��, 0) & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_����·��ҽ���䶯_Insert(Null,To_Date('" & strAddDate & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & UserInfo.���� & "'," & _
                            !��ĿID & "," & !ID & "," & ZVal(NVL(!���id, 0)) & "," & !��� & "," & !��Ч & "," & _
                            ZVal(NVL(!������ĿID, 0)) & "," & ZVal(NVL(!�շ�ϸĿID, 0)) & ",'" & NVL(!ҽ������) & "'," & ZVal(NVL(!��������, 0)) & "," & _
                            ZVal(NVL(!�ܸ�����, 0)) & ",'" & NVL(!�걾��λ) & "'," & _
                            "'" & NVL(!��鷽��) & "','" & NVL(!ִ��Ƶ��) & "'," & ZVal(NVL(!Ƶ�ʴ���, 0)) & "," & _
                            ZVal(NVL(!Ƶ�ʼ��, 0)) & ",'" & NVL(!�����λ) & "','" & NVL(!ҽ������) & "'," & _
                            NVL(!ִ������, 0) & "," & NVL(!ִ�б��, 0) & "," & ZVal(NVL(!ִ�п���ID, 0)) & ",'" & NVL(!ʱ�䷽��) & "'," & _
                            ZVal(Val(!�Ƿ�ȱʡ & "")) & "," & ZVal(Val(!�Ƿ�ѡ & "")) & "," & ZVal(Val(!�䷽ID & "")) & "," & ZVal(Val(!�����ĿID & "")) & ")"
                    End If
                    k = k + 1: .MoveNext
                Loop
            End With

            '��Ŀ��Ϣ
            If mstrDelItemIDs <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����·����Ŀ_Delete('" & Mid(mstrDelItemIDs, 2) & "')"
            End If
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                            vItem = .Cell(flexcpData, j, i)

                            If vItem.Edit = 1 Then '����
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_����·����Ŀ_Insert(" & _
                                    vItem.ID & "," & mlng·��ID & "," & intVersion & "," & _
                                    vStep.ID & ",'" & .TextMatrix(j, .FixedCols) & "'," & _
                                    vItem.��Ŀ��� & ",'" & vItem.��Ŀ���� & "'," & _
                                    vItem.ִ�з�ʽ & "," & _
                                    "'" & vItem.��Ŀ��� & "'," & ZVal(vItem.ͼ��ID) & "," & _
                                    "'" & vItem.ҽ��IDs & "','" & vItem.�������� & "'," & vItem.����Ҫ�� & "," & _
                                    "Null,Null )"
                                    
                            ElseIf vItem.Edit = 2 Or (vItem.Edit = 0 And vItem.ҽ��IDs <> "") Then '�޸ģ�����ǿ�����±���ҽ����ϵ
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_����·����Ŀ_Update(" & _
                                    vItem.ID & "," & mlng·��ID & "," & intVersion & "," & _
                                    vItem.��Ŀ��� & ",'" & vItem.��Ŀ���� & "'," & _
                                    vItem.ִ�з�ʽ & "," & _
                                    "'" & vItem.��Ŀ��� & "'," & ZVal(vItem.ͼ��ID) & "," & _
                                    "'" & vItem.ҽ��IDs & "','" & vItem.�������� & "'," & vItem.����Ҫ�� & ",'" & .TextMatrix(j, .FixedCols) & "')"
                            End If
                        End If
                    Next
                End If
            Next

            '�׶��������ͽ׶κ���Ŀ��أ���������
            For i = .FixedCols + .FrozenCols To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    If vStep.Edit = 1 Or vStep.Edit = 2 Then '�������޸�
                        If Not vStep.����.ָ�꼯 Is Nothing Then
                            For j = 1 To vStep.����.ָ�꼯.count
                                vEvalMark = vStep.����.ָ�꼯(j)
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_����·������ָ��_Insert(" & _
                                    mlng·��ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                    vEvalMark.ID & "," & vEvalMark.��� & "," & _
                                    "'" & vEvalMark.����ָ�� & "'," & vEvalMark.ָ������ & "," & _
                                    "'" & vEvalMark.ָ���� & "')"
                            Next
                        End If
                        If Not vStep.����.������ Is Nothing Then
                            For j = 1 To vStep.����.������.count
                                vEvalCond = vStep.����.������(j)
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = "Zl_����·����������_Insert(" & _
                                    mlng·��ID & "," & intVersion & "," & vStep.ID & ",2," & _
                                    ZVal(vEvalCond.ָ��ID) & "," & ZVal(vEvalCond.��ĿID) & "," & _
                                    "'" & vEvalCond.��ϵʽ & "','" & vEvalCond.����ֵ & "'," & _
                                    vEvalCond.������� & ")"
                            Next
                        End If
                    End If
                End If
            Next
        End If
    End With

    'ִ���ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0

    '---
    mstrDelStepIDs = ""
    mstrDelItemIDs = ""
    mstrChangeItemIDs = ""
    mblnChange = False
    mblnNewVersion = False

    'List��ֻ������ֻ�����¼���
    i = vsPath.Row: j = vsPath.Col

    Call LoadPathVersion(intVersion)
    
    If i <= vsPath.Rows - 1 Then vsPath.Row = i
    If j <= vsPath.Cols - 1 Then vsPath.Col = j

    Call vsPath.ShowCell(vsPath.Row, vsPath.Col)

    SavePathTable = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub FuncPathTableOutput(bytStyle As Byte, Optional ByVal blnIsAll As Boolean, Optional ByVal blnIsMe As Boolean)
'���ܣ���������ٴ�·����
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
'      blnIsAll=�Ƿ��������
'      blnIsMe=ȫ�����ʱ��������
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION

    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim lngRow As Long, lngCol As Long
    Dim bytR As Byte, strTemp As String
    Dim vItem As TYPE_PATH_ITEM
    Dim lngStart As Long

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    If blnIsAll Then
        'ֻ���������ʱ�ŷֱ��ӡ��֧
        Call LoadPathTable(objCombo)
    End If

    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))

    '��ͷ
    objOut.Title.Text = Me.Tag
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��׼����ʱ�䣺" & vVersion.��׼����ʱ�� & "��"
    objRow.Add "��׼���ã�" & vVersion.��׼���� & "Ԫ"
    objOut.UnderAppRows.Add objRow

    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.���� & vbCrLf & "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow


    '����ǰ����Ƶ��������ϣ����Ҽ��ض�Ӧ��ҽ����Ϣ�����
    With vsPathExport
        .Redraw = flexRDNone
        .Clear
        .Rows = 1: .Rows = vsPath.Rows
        .Cols = (vsPath.Cols - 1) * 2 + 1   '���̶����⣬����ÿ�к�����һ��������ʾ��Ӧ��ҽ����
        .FixedRows = 0: .FixedCols = 0
        .Width = vsPath.Width
        .Height = vsPath.Height

        .Redraw = flexRDDirect

        .Redraw = flexRDNone
        '��һ����ʾ�汾��Ϣ
        .TextMatrix(0, 1) = "��ǰ�汾����" & vVersion.�汾�� & "��" & IIf(vVersion.�汾˵�� <> "", "��" & vVersion.�汾˵��, "") & _
                            vbCrLf & IIf(vVersion.����� = "", "����ʱ�䣺" & Format(vVersion.����ʱ��, "yyyy��MM��dd��") & vbCrLf & "�����ˣ�" & vVersion.������ & "(δ���)", _
                            "���ʱ�䣺" & Format(vVersion.���ʱ��, "yyyy��MM��dd��") & vbCrLf & "����ˣ�" & vVersion.�����)
        .TextMatrix(0, 2) = "���ÿ��ң�" & mstrDeptInfo & vbCrLf & "���ò��֣�" & mstrDiagInfo

        '�ڶ��У������У�·���׶���Ϣ
        '�ӵ�1�п�ʼ����0��Ϊ�̶��У�������
        If Trim(vsPath.TextMatrix(2, 0)) = "ʱ��׶�" Then
            lngStart = 2
        Else
            lngStart = 1
        End If
        For lngRow = 1 To lngStart
            For lngCol = 1 To vsPath.Cols - 1
                If vsPath.TextMatrix(lngRow, lngCol) <> vsPath.TextMatrix(lngRow - 1, lngCol) Then
                    .TextMatrix(lngRow, lngCol * 2 - 1) = Replace(Replace(vsPath.TextMatrix(lngRow, lngCol), vbLf, ""), vbCr, "")
                End If
            Next
        Next

        '�����У�·����Ŀ
        For lngCol = 0 To vsPath.Cols - 1
            'ҽ����
            .ColAlignment(lngCol * 2) = vsPath.ColAlignment(lngCol)
            .ColWidth(lngCol * 2) = vsPath.ColWidth(lngCol) * 1.6

            If lngCol = 0 Then
                '��Ŀ���
                For lngRow = lngStart + 1 To vsPath.Rows - 1
                    If vsPath.TextMatrix(lngRow, 0) <> vsPath.TextMatrix(lngRow - 1, 0) Then .TextMatrix(lngRow, 0) = vsPath.TextMatrix(lngRow, 0)
                Next
            Else
                .ColAlignment(lngCol * 2 - 1) = vsPath.ColAlignment(lngCol)
                .ColWidth(lngCol * 2 - 1) = vsPath.ColWidth(lngCol)


                '��ǰ�е�����·����Ŀ��
                .Cell(flexcpText, lngStart + 1, lngCol * 2 - 1, .Rows - 1, lngCol * 2 - 1) = vsPath.Cell(flexcpText, lngStart + 1, lngCol, .Rows - 1, lngCol)

                For lngRow = lngStart + 1 To vsPath.Rows - 1

                    If TypeName(vsPath.Cell(flexcpData, lngRow, lngCol)) <> "Empty" Then
                        vItem = vsPath.Cell(flexcpData, lngRow, lngCol)
                        strTemp = vItem.Tip 'ҽ�����ݻ�������ժҪ
                        If InStr(strTemp, ":") > 0 Then
                            strTemp = Trim(Mid(strTemp, InStr(strTemp, ":") + 1))
                        Else
                            If vItem.ҽ��IDs <> "" Then
                                strTemp = GetAdviceDefineText(vItem.ҽ��IDs, mrsAdvice)
                            ElseIf vItem.����IDs <> "" Or vItem.�°没��IDs <> "" Then
                                If vItem.����IDs <> "" And vItem.�°没��IDs <> "" Then
                                    strTemp = GetEPRDefineTextOut(, vItem.ID)
                                ElseIf vItem.����IDs <> "" Then
                                    strTemp = GetEPRDefineTextOut(vItem.����IDs)
                                Else
                                    strTemp = GetEPRDefineTextOut(vItem.�°没��IDs, vItem.ID)
                                End If
                            End If
                        End If

                        strTemp = Replace(strTemp, "��", "")
                        .TextMatrix(lngRow, lngCol * 2) = strTemp
                    End If
                Next
            End If
        Next
        .Redraw = flexRDDirect
    End With

    '����
    Set objOut.Body = vsPathExport
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If

    '����������ͷ��ڴ�
    vsPathExport.Clear
    vsPathExport.Rows = 1: vsPathExport.Cols = 1

    '�����ȫ���������ѭ�����ã�ֱ�����һ��
    If blnIsAll Or blnIsMe Then
        Call LoadPathTable(objCombo)
        Call FuncPathTableOutput(bytStyle, False, True)
    End If
End Sub

Private Sub FuncExportToXML()
'���ܣ�������XML�ļ�
    Dim objCombo As CommandBarComboBox
    Dim vVersion As TYPE_PATH_VERSION

    If mbytMode = Mode_Design And mblnChange Then
        MsgBox "·�������ݱ������δ���棬���ȱ��档", vbInformation, gstrSysName
        Exit Sub
    End If

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))

    '����Ŀ¼ѡ��
    cdgXML.DialogTitle = "���������ٴ�·��"
    cdgXML.Filter = "XML�ļ�|*.xml"
    cdgXML.Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
    cdgXML.InitDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼")
    cdgXML.FileName = Replace(Me.Tag, vbCrLf, "_") & ".xml"
    cdgXML.CancelError = True
    On Error Resume Next
    cdgXML.ShowSave
    If Err.Number <> 0 Then
        '����ȡ��ʱ
        If Err.Number <> 32755 Then MsgBox "�������̷�������:" & Err.Description, vbInformation, gstrSysName
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼", gobjFile.GetParentFolderName(cdgXML.FileName)

    '����
    Screen.MousePointer = 11
    Call ExportOutPathToXML(mlng·��ID, vVersion.�汾��, cdgXML.FileName)
    Screen.MousePointer = 0
End Sub

Private Sub FuncPathImportFromXML()
    Dim objCombo As CommandBarComboBox
    Dim intVersion As Integer, k As Long, i As Long

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    If objCombo Is Nothing Then Exit Sub
    If objCombo.ListIndex = 0 Then Exit Sub

    cdgXML.DialogTitle = "���������ٴ�·��"
    cdgXML.Filter = "XML�ļ�|*.xml"
    cdgXML.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgXML.InitDir = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼")
    cdgXML.CancelError = True
    On Error Resume Next
    cdgXML.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "�����ٴ�·��XMLĿ¼", gobjFile.GetParentFolderName(cdgXML.FileName)

    'ȷ������汾��
    If mblnNewVersion Then
        k = 0
        For i = 1 To objCombo.ListCount
            If objCombo.ItemData(i) > k Then k = objCombo.ItemData(i)
        Next
        intVersion = k + 1
    Else
        intVersion = objCombo.ItemData(objCombo.ListIndex)
    End If

    '����·��
    Screen.MousePointer = 11
    If ImportOutPathFromXML(cdgXML.FileName, mlng·��ID, intVersion) Then
        mstrDelStepIDs = ""
        mstrDelItemIDs = ""
        mblnChange = False
        mblnNewVersion = False
        Call LoadPathVersion(intVersion)
    End If
    Screen.MousePointer = 0
End Sub

Private Function GetParentStep(vStep As TYPE_PATH_STEP) As TYPE_PATH_STEP
'���ܣ���ȡ��֧�׶εĸ��׶�
    Dim i As Long

    With vsPath
        For i = .FixedCols + .FrozenCols To .Cols - 1
            If TypeName(.ColData(i)) <> "Empty" Then
                If .ColData(i).ID = vStep.��ID Then
                    GetParentStep = .ColData(i)
                    Exit For
                End If
            End If
        Next
    End With
End Function

Private Sub FuncFindItem(Optional ByVal blnNext As Boolean)
'������blnNext=�Ƿ������һ��
    Dim blnHave As Boolean, i As Long, j As Long
    Dim vStep As TYPE_PATH_STEP
    Dim vItem As TYPE_PATH_ITEM
    Dim lngRow As Long, lngCol As Long
    Dim blnOver As Boolean

    If Trim(txtFind.Text) = "" Then Exit Sub
    Call zlControl.TxtSelAll(txtFind)

    '��ʼ������
    With vsPath
        If .Row < .FixedRows + .FrozenRows Or .Col < .FixedCols + .FrozenCols Then .Row = .FixedRows + .FrozenRows: .Col = .FixedCols + .FrozenCols

        If blnNext Then
            If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                blnOver = True
            Else
                lngRow = .Row: lngCol = .Col
                If .Row = .Rows - 1 Then
                    lngRow = .FixedRows + .FrozenRows
                    lngCol = .Col + 1
                Else
                    lngRow = .Row + 1
                End If
            End If
        Else
            lngCol = .FixedCols + .FrozenCols: lngRow = .FixedRows + .FrozenRows
        End If
        '�ӵ�ǰ�п�ʼ������(�����ң��������£�
        If Not blnOver Then
            For i = lngCol To .Cols - 1
                If TypeName(.ColData(i)) <> "Empty" Then
                    vStep = .ColData(i)
                    For j = .FixedRows + .FrozenRows To .Rows - 1
                        If i <> lngCol Or j >= lngRow Then
                            If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                                vItem = .Cell(flexcpData, j, i)
                                If vItem.��Ŀ���� Like IIf(gstrLike <> "", "*", "") & txtFind.Text & "*" Then
                                    blnHave = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If blnHave Then Exit For
                End If
            Next
        End If

        If blnHave And Not blnOver Then
            .Row = j: .Col = i
            .ShowCell .Row, .Col
            If .Visible Then .SetFocus
        Else
            MsgBox IIf(blnNext, "������", "") & "�Ҳ�������������·����Ŀ��", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub ShowContrast(ByVal bytMode As Byte)
'����:1.�Բ�ͬ����ɫ������ʾҽ����������һ�汾�в������Ŀ
'     ����ɫ����ɫ��ʾ��&H00FFEADA&������,
'����:bytMode  1-��ʾ���죬2-���ز���

    Dim rsNew As ADODB.Recordset, rsOld As ADODB.Recordset
    Dim rsAdviceNew As ADODB.Recordset, rsAdviceOld As ADODB.Recordset
    Dim strSql As String
    Dim objCombo As CommandBarComboBox
    Dim vItem As TYPE_PATH_ITEM
    Dim strTmp As String
    Dim lngRow As Long, lngCol As Long
    Dim lngVersion As Long
    Dim i As Long, j As Long, lngCount As Long
    Dim intOldItemId As Long
    Dim blnDo As Boolean

    On Error GoTo errH
    blnDo = False
    If bytMode = 1 Then
        Set mcolItemID = New Collection
        Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
        lngVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex)).�汾��    '��ǰ�汾��
        If lngVersion < 2 Then
            Exit Sub
        End If

        '����ҽ�������Ŀ���� ��������ٴ�·������ Ϊ�˰������������󣬱��ڰ��մ��ϵ��£������ҵ�˳�����mcolItemID
        strSql = "Select a.Id As �׶�id, a.���, Nvl(b.���, 0) As ��id���, a.����, a.��ʼ����, Nvl(a.��������, 0) As ��������, c.����, c.Id As ��Ŀid, c.��Ŀ����" & vbNewLine & _
                 "From ����·���׶� A, ����·���׶� B, ����·����Ŀ C, ����·������ D" & vbNewLine & _
                 "Where a.·��id = [1] And a.�汾�� = [2] " & _
                 "  And a.��id = b.Id(+) And a.Id = c.�׶�id And d.·��id = c.·��id And" & vbNewLine & _
                 "      d.�汾�� = c.�汾�� And d.���� = c.���� And Exists" & vbNewLine & _
                 " (Select 1 From ����·��ҽ�� D Where c.Id = d.·����Ŀid)" & vbNewLine & _
                 "Order By Nvl(b.���, a.���), Nvl(b.���, 0), a.���, d.���, c.��Ŀ���"
        Set rsNew = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, lngVersion)      '�°汾
        '�ɰ�������Ŀ����
        strSql = "Select a.Id As �׶�id, a.���, Nvl(b.���, 0) As ��id���, a.����, a.��ʼ����, Nvl(a.��������, 0) As ��������, c.����, c.Id As ��Ŀid, c.��Ŀ����" & vbNewLine & _
                 "From ����·���׶� A, ����·���׶� B, ����·����Ŀ C,����·������ D" & vbNewLine & _
                 "Where a.·��id = [1] And a.�汾�� = [2] " & _
                 " And a.��id = b.Id(+) And a.Id = c.�׶�id  And d.·��id = c.·��id And" & vbNewLine & _
                 "     d.�汾�� = c.�汾�� And d.���� = c.���� " & vbNewLine & _
                 "Order By Nvl(b.���, a.���), Nvl(b.���, 0), Nvl(a.���, 0), d.���, c.��Ŀ���"
        Set rsOld = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, lngVersion - 1)       '�ɰ汾

        Do While Not rsNew.EOF
            rsOld.Filter = "��� =" & Val(NVL(rsNew!���)) & " And ��id��� = " & Val(NVL(rsNew!��id���)) & " And ��ʼ���� =" & Val(NVL(rsNew!��ʼ����)) & _
                           " And ��������= " & Val(NVL(rsNew!��������)) & " And ���� ='" & NVL(rsNew!����) & "' And ��Ŀ���� = '" & NVL(rsNew!��Ŀ����) & "'"

            If rsOld.RecordCount > 0 Then
                'ͬ�׶Σ�ͬ���࣬ͬ��Ŀ
                'strSql����н�������������ֵת����Ϊ�˱�����Filter������
                strSql = "Select b.���, b.��Ч, Nvl(b.������Ŀid,0) as ������ĿID, Nvl(b.�շ�ϸĿid, 0) as �շ�ϸĿid, Nvl(b.ҽ������, 0) As ҽ������, Nvl(b.��������, 0) As ��������," & vbNewLine & _
                         "       Nvl(b.�ܸ�����, 0) As �ܸ�����, Nvl(b.ִ��Ƶ��,0) as ִ��Ƶ��, b.ִ������, Nvl(b.��鷽��, 0) As ��鷽��, Nvl(b.�걾��λ, 0) As �걾��λ," & vbNewLine & _
                         "       Nvl(b.ִ�п���id, 0) As ִ�п���id, Nvl(b.ʱ�䷽��, 0)  as ʱ�䷽��" & vbNewLine & _
                         "From ����·��ҽ�� A, ����·��ҽ������ B" & vbNewLine & _
                         "Where a.·����Ŀid = [1] And a.ҽ������id = b.Id" & vbNewLine & _
                         "Order By b.���"

                Set rsAdviceNew = zlDatabase.OpenSQLRecord(strSql, Me.Caption, rsNew!��ĿID)
                Set rsAdviceOld = zlDatabase.OpenSQLRecord(strSql, Me.Caption, rsOld!��ĿID)

                If rsAdviceNew.RecordCount > 0 And rsAdviceOld.RecordCount = 0 Then
                    '��һ��,�°���ҽ����Ŀ���ɰ治��ҽ����Ŀ
                    intOldItemId = rsOld!��ĿID
                    blnDo = True
                ElseIf rsAdviceNew.RecordCount > 0 And rsAdviceOld.RecordCount > 0 Then
                    '�ڶ��֣��°�ɰ涼��ҽ����Ŀ
                    For i = 1 To rsAdviceNew.RecordCount
                        rsAdviceOld.Filter = "��Ч = " & Val(NVL(rsAdviceNew!��Ч)) & " And ������ĿID = " & Val(NVL(rsAdviceNew!������ĿID)) & " and �շ�ϸĿID=" & Val(NVL(rsAdviceNew!�շ�ϸĿID)) & _
                                             " And ҽ������ ='" & NVL(rsAdviceNew!ҽ������) & "' And �������� =" & Val(NVL(rsAdviceNew!��������)) & " And �ܸ����� = " & Val(NVL(rsAdviceNew!�ܸ�����)) & _
                                             " And ִ��Ƶ�� = '" & NVL(rsAdviceNew!ִ��Ƶ��) & "' And ִ������ ='" & NVL(rsAdviceNew!ִ������) & "' And ��鷽�� = '" & NVL(rsAdviceNew!��鷽��) & "'" & _
                                             " And �걾��λ = '" & NVL(rsAdviceNew!�걾��λ) & "' And ִ�п���ID =" & Val(NVL(rsAdviceNew!ִ�п���ID)) & " And ʱ�䷽�� = '" & NVL(rsAdviceNew!ʱ�䷽��) & "'"
                        'һ����һ��ҽ������ͬ,���˳�ѭ��
                        If rsAdviceOld.RecordCount = 0 Then
                            intOldItemId = rsOld!��ĿID
                            blnDo = True
                            Exit For
                        End If
                        rsAdviceNew.MoveNext
                    Next
                End If
            ElseIf rsOld.RecordCount < 1 Then    '��������ͬ�Ľ׶λ���ͬ�������ͬ��Ŀʱ
                intOldItemId = 0
                blnDo = True
            End If
            If blnDo Then

                lngCount = lngCount + 1
                '��¼�´��ڲ������ĿID,����ԱȲ鿴ʱ,��һ������һ����ȡ��ĿID
                'item�� �°���ĿID:�ϰ���ĿID:�±�λ��
                mcolItemID.Add Val(rsNew!��ĿID) & ":" & intOldItemId & ":" & lngCount, "_" & Val(rsNew!��ĿID)
                strTmp = mcolItemRowCol("_" & rsNew!��ĿID)
                lngRow = Split(strTmp, ",")(0)
                lngCol = Split(strTmp, ",")(1)
                With vsPath
                    vItem = .Cell(flexcpData, lngRow, lngCol)
                    vItem.ǰһ�汾��ĿID = intOldItemId
                    .Cell(flexcpData, lngRow, lngCol) = vItem
                End With
                blnDo = False
            End If

            rsNew.MoveNext
        Loop

        If mcolItemID.count = 0 Then
            mblnDiff = False
            MsgBox "�ð汾ҽ������Ŀͬ��һ�汾��ͬ��", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
    End If

    '���ò�����ɫ/���ز���
    For i = 1 To mcolItemID.count
        strTmp = mcolItemRowCol("_" & Split(mcolItemID(i), ":")(0))
        lngRow = Split(strTmp, ",")(0)
        lngCol = Split(strTmp, ",")(1)
        vsPath.Cell(flexcpBackColor, lngRow, lngCol) = IIf(bytMode = 1, Color_DiffBack, Empty)
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CompareAdviceItem()
    Dim vItem As TYPE_PATH_ITEM
    Dim i As Long

    '�ԱȲ鿴
    With vsPath
        If .TextMatrix(.Row, .Col) <> "" Then
            vItem = .Cell(flexcpData, .Row, .Col)
            If .Cell(flexcpBackColor, .Row, .Col) = Color_DiffBack Then
                mfrmAdviceContrast.ShowMe Me, vItem.ID, vItem.ǰһ�汾��ĿID, mcolItemID, 1
            Else
                MsgBox "��ѡ��һ����ɫ�����ĵ�Ԫ����ִ�жԱȲ鿴��", vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            MsgBox "�㵱ǰѡ��ĵ�Ԫ��û�ж���·����Ŀ����ѡ��һ����ɫ�����ĵ�Ԫ��", vbOKOnly + vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub FuncResizeCenter()
    Dim objControl As CommandBarControl

    On Error Resume Next

    If mbytFunc = 0 Then
        Me.picBottom.Visible = False
        Me.fraSplit.Visible = False
        vsPath.Move 0, 0, picCenter.Width, picCenter.Height
    ElseIf mbytFunc = 1 Or mbytFunc = 2 Then
        Me.picBottom.Visible = True
        Me.fraSplit.Visible = True
        UCAdvice(0).Setѡ���еĿɼ��� (True)
        UCAdvice(1).Setѡ���еĿɼ��� (True)
        vsPath.Move 0, 0, picCenter.Width, picCenter.Height / 10 * 7
        fraSplit.Move 0, picCenter.Height / 10 * 7, picCenter.Width, 45
        picBottom.Move 0, fraSplit.Top + 45, picCenter.Width, picCenter.Height - fraSplit.Top - 50

        If mbytFunc = 2 Then
            cmdCheck(0).Visible = True
            cmdCheck(1).Visible = True
        Else
            cmdCheck(0).Visible = False
            cmdCheck(1).Visible = False
        End If
        Call FuncResizeBottom
    End If
End Sub

Private Sub FuncResizeBottom()
'����:���µ����䶯��¼λ��
    On Error Resume Next

    lblCurr.Move 120, 50, 1095, 300
    UCAdvice(0).Move 120, 360, picBottom.Width / 2 - 120, picBottom.Height - 300
    fraSplit2.Move picBottom.ScaleWidth / 2, 400, 60, picBottom.Height
    lblChange.Move fraSplit2.Left + 120, 50, 1095, 300
    With cboTimes
        .Left = fraSplit2.Left + 60 + lblChange.Width + 120: .Top = 15
        .Width = IIf(mbytFunc = 1, 8000, 5000)
        .Height = 300
    End With
    If cmdCheck(0).Visible Then cmdCheck(0).Move cboTimes.Left + cboTimes.Width + 500, cboTimes.Top, 1100, 360
    If cmdCheck(1).Visible Then cmdCheck(1).Move cmdCheck(0).Left + cmdCheck(0).Width + 120, cmdCheck(0).Top, 1100, 360
    UCAdvice(1).Move fraSplit2.Left + 60, 360, picBottom.Width - fraSplit2.Left - 60 - 120, picBottom.Height - 300
End Sub

Private Sub FuncShowAdvice(Optional ByVal bytModel As Byte = 0)
'����:��ʾ�䶯��¼
'����:bytModel = 0 ��ʾ��ǰҽ������
'              = 1 ��ʾָ����·��ҽ���䶯��¼
'              = 2 ���ҽ����¼
    Dim lng��ĿID As Long
    Dim vItem As TYPE_PATH_ITEM
    Dim strSQLOne As String
    Dim strSQLTwo As String

    On Error GoTo errH
    If vsPath.Row < 0 Or vsPath.Col < 0 Then Exit Sub

    If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) <> "Empty" And InStr(",0,1,", "," & bytModel & ",") > 0 Then
        vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
        If bytModel = 0 Then
            strSQLOne = "Select a.Id, a.���id, a.���, a.��Ч, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.ҽ������, a.ִ��Ƶ��, a.Ƶ�ʴ���," & vbNewLine & _
                 "       a.Ƶ�ʼ��, a.�����λ, a.ִ������, a.ִ�п���id, a.ʱ�䷽��, a.�Ƿ�ȱʡ, a.�Ƿ�ѡ, a.�䷽id, a.�����Ŀid,a.ִ�б�� " & vbNewLine & _
                 "From ����·��ҽ������ A, ����·��ҽ�� B" & vbNewLine & _
                 "Where a.Id = b.ҽ������id And b.·����Ŀid =[3] "
            UCAdvice(0).ShowAdvice 0, strSQLOne, , , True, vItem.ID
        ElseIf bytModel = 1 Then
            strSQLTwo = "Select a.ҽ������ID as Id, a.���id, a.���, a.��Ч, a.������Ŀid, a.�շ�ϸĿid, a.ҽ������, a.��������, a.�ܸ�����, a.�걾��λ, a.��鷽��, a.ҽ������, a.ִ��Ƶ��, a.Ƶ�ʴ���," & vbNewLine & _
                "       a.Ƶ�ʼ��, a.�����λ, a.ִ������, a.ִ�п���id, a.ʱ�䷽��, a.�Ƿ�ȱʡ, a.�Ƿ�ѡ, a.�䷽id, a.�����Ŀid,a.ִ�б�� " & vbNewLine & _
                "From ����·��ҽ���䶯 A " & vbNewLine & _
                "Where a.��ĿId = [3] and a.����ʱ��= To_Date('" & Format(cboTimes.Tag, "yyyy-mm-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
            UCAdvice(1).ShowAdvice 0, strSQLTwo, , , True, vItem.ID
        End If
    Else
        UCAdvice(0).ShowAdvice 0, "", , , True
        UCAdvice(1).ShowAdvice 0, "", , , True
        cboTimes.Clear
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncSetItemBackColor()
'����:���Ҵ���·��ҽ���䶯��·����Ŀ,����������ҽ���䶯��·����Ŀ��������Ϊ��ɫ
    Dim i As Long
    Dim j As Long
    Dim vVersion As TYPE_PATH_VERSION
    Dim vItem As TYPE_PATH_ITEM
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strIDs As String
    Dim objCombo As CommandBarComboBox

    Set objCombo = cbsMain(cbsMain.count).FindControl(, cmd_Edit_Version, True)
    vVersion = mcolVersion("_" & objCombo.ItemData(objCombo.ListIndex))
    On Error GoTo errH
    If mbytFunc = 1 Then
        strSql = " Select Distinct b.��Ŀid From ����·����Ŀ A, ����·��ҽ���䶯 B Where a.·��id = [1] And a.�汾�� = [2] And a.Id = b.��Ŀid And B.���״̬=1 "

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, vVersion.�汾��)
        If rsTmp.RecordCount < 1 Then
            MsgBox "�������ٴ�·��������ҽ���䶯��¼��", vbOKOnly + vbInformation, gstrSysName
        End If
        For i = 1 To rsTmp.RecordCount
            strIDs = strIDs & "," & rsTmp!��ĿID
            rsTmp.MoveNext
        Next
        strIDs = strIDs & ","
    End If
    With vsPath
        For i = .FixedCols To .Cols - 1
            For j = 1 To .Rows - 1
                If TypeName(.Cell(flexcpData, j, i)) <> "Empty" Then
                    If mbytFunc = 1 Then
                        vItem = .Cell(flexcpData, j, i)
                        If InStr(strIDs, "," & vItem.ID & ",") > 0 Then
                            .Cell(flexcpBackColor, j, i) = Color_DiffBack
                        End If
                    Else
                        If .Cell(flexcpBackColor, j, i) = Color_DiffBack Then
                            .Cell(flexcpBackColor, j, i) = 0
                        End If
                    End If
                End If
            Next
        Next
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub FuncLoadChangeTimes()
'����:ҽ���䶯����䶯��������
'����:mbytFunc=1 ҽ���䶯��ʷ��¼(�����\���ʱ�� ��Ϊ��
'     mbytFunc=2 ·����Ŀ�䶯����˵ļ�¼ ҽ���䶯��¼(�����=NULL)�ļ�¼
    Dim strSql As String, strWhere As String
    Dim rsTmp As ADODB.Recordset
    Dim vItem As TYPE_PATH_ITEM
    Dim i As Long

    On Error GoTo errH

    cboTimes.Clear: cboTimes.Tag = ""

    If TypeName(vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)) = "Empty" Then Exit Sub
    marrTime = Array()
    If mbytFunc = 1 Then
        strSql = "Select Rownum As ���, a.*" & vbNewLine & _
                    "From (Select Distinct a.����ʱ��, a.����Ա,a.���״̬, a.�����, a.���ʱ��" & vbNewLine & _
                    "       From ����·��ҽ���䶯 A" & vbNewLine & _
                    "       Where a.��Ŀid = [1] And a.����� is Not Null " & vbNewLine & _
                    "       Order By a.����ʱ��) A" & vbNewLine & _
                    "Order By Rownum Desc"
    ElseIf mbytFunc = 2 Then
        strSql = " Select Distinct a.����ʱ��, a.����Ա, a.�����, a.���ʱ��" & vbNewLine & _
                 " From ����·��ҽ���䶯 A" & vbNewLine & _
                 " Where a.��Ŀid = [1] And a.����� is Null " & vbNewLine & _
                 " Order By a.����ʱ�� Desc"
    End If
    vItem = vsPath.Cell(flexcpData, vsPath.Row, vsPath.Col)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vItem.ID)
    If mbytFunc = 1 Then
        For i = 1 To rsTmp.RecordCount
            cboTimes.AddItem "��" & rsTmp!��� & "��,�Ǽ�:" & rsTmp!����Ա & "/" & Format(rsTmp!����ʱ��, "yyyy-mm-dd hh:mm:ss") & " ���:" & rsTmp!����� & "/" & rsTmp!���ʱ�� & Space(1) & IIf(Val(rsTmp!���״̬ & "") = 0, "���δͨ��", "���ͨ��")
            ReDim Preserve marrTime(UBound(marrTime) + 1)
            marrTime(UBound(marrTime)) = Format(rsTmp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            rsTmp.MoveNext
        Next
    ElseIf mbytFunc = 2 Then
        For i = 1 To rsTmp.RecordCount
            cboTimes.AddItem "�Ǽ�:" & rsTmp!����Ա & "/" & Format(rsTmp!����ʱ��, "yyyy-mm-dd hh:mm:ss") & Space(1) & "�����"
            ReDim Preserve marrTime(UBound(marrTime) + 1)
            marrTime(UBound(marrTime)) = Format(rsTmp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            rsTmp.MoveNext
        Next
    End If

    If cboTimes.ListCount > 0 Then
        cboTimes.ListIndex = 0   'ȱʡ��λ�����±䶯��¼\���´���˼�¼
    Else
        UCAdvice(1).ShowAdvice 0, "", 0, 0, True '�������
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
