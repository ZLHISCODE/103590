VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppUpgrade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ϣ����ƽ̨�ͻ�������"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8025
   Icon            =   "frmAppUpgrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6765
      TabIndex        =   6
      Top             =   4710
      Width           =   1100
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   45
      ScaleHeight     =   840
      ScaleWidth      =   7995
      TabIndex        =   4
      Top             =   0
      Width           =   7995
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϣ����ƽ̨�ͻ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   225
         Width           =   2835
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   7170
         Picture         =   "frmAppUpgrade.frx":6852
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   345
      Left            =   270
      TabIndex        =   3
      Top             =   4725
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ��(&P)"
      Height          =   345
      Left            =   5610
      TabIndex        =   2
      Top             =   4725
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   870
      Width           =   8100
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   8100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1560
      Top             =   4605
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppUpgrade.frx":9CD4
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppUpgrade.frx":A26E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppUpgrade.frx":10AD0
            Key             =   "�����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppUpgrade.frx":1106A
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppUpgrade.frx":11604
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAppUpgrade.frx":11B9E
            Key             =   "ȫ��"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   3
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   21
      Top             =   840
      Width           =   7950
      Begin VB.Frame Frame4 
         Height          =   2430
         Left            =   780
         TabIndex        =   22
         Top             =   1080
         Width           =   6840
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   9
            Left            =   1320
            TabIndex        =   25
            Top             =   1170
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   26
            Top             =   1575
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   7
            Left            =   1320
            TabIndex        =   24
            Top             =   765
            Width           =   5070
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   6
            Left            =   1320
            TabIndex        =   23
            Top             =   360
            Width           =   5070
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�û�"
            Height          =   180
            Index           =   16
            Left            =   750
            TabIndex        =   30
            Top             =   1215
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   15
            Left            =   750
            TabIndex        =   29
            Top             =   1635
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������˿�"
            Height          =   180
            Index           =   14
            Left            =   210
            TabIndex        =   28
            Top             =   810
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������ַ"
            Height          =   180
            Index           =   13
            Left            =   210
            TabIndex        =   27
            Top             =   405
            Width           =   900
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������������ַ���˿ڡ��û������룬�Դ�����Ϣ����ƽ̨���������ӡ�"
         Height          =   180
         Index           =   17
         Left            =   810
         TabIndex        =   32
         Top             =   765
         Width           =   5940
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   4
         Left            =   135
         Picture         =   "frmAppUpgrade.frx":12138
         Top             =   555
         Width           =   480
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   4
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   15
      Top             =   915
      Width           =   7950
      Begin VB.CommandButton cmdUpgrade 
         Caption         =   "����(&U)"
         Height          =   345
         Left            =   960
         TabIndex        =   16
         Top             =   3195
         Width           =   1100
      End
      Begin MSComctlLib.ProgressBar pgb 
         Height          =   225
         Left            =   2130
         TabIndex        =   17
         Top             =   3360
         Visible         =   0   'False
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStep 
         Height          =   2490
         Left            =   975
         TabIndex        =   18
         Top             =   600
         Width           =   6840
         _cx             =   2088840993
         _cy             =   2088833320
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���ڰ�װ.."
         Height          =   180
         Index           =   12
         Left            =   2130
         TabIndex        =   31
         Top             =   3165
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   3
         Left            =   195
         Picture         =   "frmAppUpgrade.frx":13ABA
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������������ʼ������Ϣ�ͻ���"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   20
         Top             =   210
         Width           =   2880
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   180
         Index           =   6
         Left            =   7395
         TabIndex        =   19
         Top             =   3390
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   1
      Left            =   45
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   10
      Top             =   915
      Width           =   7950
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1155
         TabIndex        =   12
         Top             =   795
         Width           =   6330
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   0
         Left            =   7500
         Picture         =   "frmAppUpgrade.frx":1543C
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   795
         Width           =   315
      End
      Begin MSComDlg.CommonDialog cdl 
         Left            =   7410
         Top             =   165
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѡ����Ϣ����ƽ̨�ͻ������������ļ�"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   14
         Top             =   270
         Width           =   3060
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ŀ��汾"
         Height          =   180
         Index           =   2
         Left            =   1185
         TabIndex        =   13
         Top             =   1305
         Width           =   720
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   2
         Left            =   225
         Picture         =   "frmAppUpgrade.frx":1BC8E
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3645
      Index           =   2
      Left            =   30
      ScaleHeight     =   3645
      ScaleWidth      =   7950
      TabIndex        =   7
      Top             =   900
      Width           =   7950
      Begin VSFlex8Ctl.VSFlexGrid vsfUser 
         Height          =   2820
         Left            =   975
         TabIndex        =   8
         Top             =   600
         Width           =   6840
         _cx             =   2088840993
         _cy             =   2088833902
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����������ݿ��û��ĵ�¼����"
         Height          =   180
         Index           =   3
         Left            =   960
         TabIndex        =   9
         Top             =   165
         Width           =   2520
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   0
         Left            =   180
         Picture         =   "frmAppUpgrade.frx":1D610
         Top             =   150
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmAppUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mobjFso As New FileSystemObject
Private mclsOracle As clsDataOracle
Private mblnStep(1 To 2) As Boolean
Private mstrManageVersion As String
Private mstrVersion As String
Private mstrDBVersion As String
Private mintPage As Integer
Private mclsVsf As zlVSFlexGrid.clsVsf
Private mclsVsfStep As zlVSFlexGrid.clsVsf
Private mclsVsfUser As zlVSFlexGrid.clsVsf
Private mblnEMR As Boolean
Private mstrBeforeEditStr As String
Private WithEvents mclsMipClientManage As clsMipClientManage
Attribute mclsMipClientManage.VB_VarHelpID = -1
Private mblnImportDB As Boolean
Private mfrmErrorInfo As frmErrorInfo

Private WithEvents mobjScript As clsOracleScript
Attribute mobjScript.VB_VarHelpID = -1

Public Function ShowDialog() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    mblnOK = False
    
    Set mclsOracle = New clsDataOracle
    
    
    For intLoop = 1 To picPage.UBound
        picPage(intLoop).Left = 0
        picPage(intLoop).Top = 915
        picPage(intLoop).Width = 7950
        picPage(intLoop).Height = 3645
    Next
    
    Call InitGrid
    
    
    mblnEMR = False
    mintPage = 1
    Call ShowPage(mintPage)
    
    Me.Show 1
    ShowDialog = mblnOK
End Function

Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
        
    '------------------------------------------------------------------------------------------------------------------
    
    Set mclsVsfUser = New zlVSFlexGrid.clsVsf
    With mclsVsfUser
        Call .Initialize(Me.Controls, vsfUser, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("������", 0, flexAlignLeftCenter, flexDTString, "*", "trigger_type", True)
        Call .AppendColumn("����", 0, flexAlignLeftCenter, flexDTString, , "id", True)
        Call .AppendColumn("ϵͳ", 1800, flexAlignLeftCenter, flexDTString, , "item_title", True)
        Call .AppendColumn("�û�", 1800, flexAlignLeftCenter, flexDTString, , "item_title", True)
        Call .AppendColumn("����", 1800, flexAlignLeftCenter, flexDTString, "*", "trigger_type", True)
        
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditText)
        .AppendRows = True
        
        
    End With
        
'    With vsfUser
'        .Rows = 4
'        .TextMatrix(1, 2) = "ZLTOOLS"
'        .TextMatrix(2, 2) = "ZLHIS"
'        .TextMatrix(3, 2) = "ZLEMR"
'
'        .TextMatrix(1, 3) = "**************"
'        .TextMatrix(2, 3) = "**************"
'        .TextMatrix(3, 3) = "**************"
'
'        mclsVsfUser.AppendRows = True
'    End With
    
    '------------------------------------------------------------------------------------------------------------------
            
    Set mclsVsfStep = New zlVSFlexGrid.clsVsf
    With mclsVsfStep
        Call .Initialize(Me.Controls, vsfStep, True, False)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        Call .AppendColumn("code", 0, flexAlignLeftCenter, flexDTString, , "data_code", True)
        Call .AppendColumn("step", 1500, flexAlignLeftCenter, flexDTString, , "item_note", True)
        Call .AppendColumn("pname", 0, flexAlignLeftCenter, flexDTString, , "item_note", True)
        vsfStep.RowHidden(0) = True
    End With
    
'    With vsfStep
'        .Rows = 8
'        .TextMatrix(1, 2) = "�������ݽṹ"
'        .TextMatrix(2, 2) = "װ�ر�׼����Ϣ����"
'        .TextMatrix(3, 2) = "װ���°没����Ϣ����"
'        .TextMatrix(4, 2) = "װ�ؼ�����Ϣ����"
'        .TextMatrix(5, 2) = "װ�������Ϣ����"
'        .TextMatrix(6, 2) = "װ��������Ϣ����"
'        .TextMatrix(7, 2) = "װ��Ѫ����Ϣ����"
'
'    End With
    
    InitGrid = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Sub ShowPage(ByVal intPage As Integer)
    Dim intLoop As Integer
    
    For intLoop = 1 To picPage.UBound
        picPage(intLoop).Visible = False
    Next
    
    picPage(intPage).Visible = True
        
    cmdNext.Enabled = (intPage < picPage.UBound)
    cmdPrevious.Enabled = (intPage > 1)
    
End Sub

Private Function OpenDialog(ByRef objDlg As Object, ByVal strDialogTitle As String, ByVal strFilter As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim strTmp As String
            
    With objDlg
        .DialogTitle = strDialogTitle
        .Filter = strFilter
    
        On Error Resume Next
    
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowOpen
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            OpenDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    MsgBox "���ܴ��ļ�(" & strTmp & "),���ļ���������ʹ�û��ļ�������!"
End Function

Private Function SaveDialog(ByRef objDlg As Object, ByVal strDialogTitle As String, ByVal strFilter As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    Dim strTmp As String
            
    With objDlg
        .DialogTitle = strDialogTitle
        .Filter = strFilter
    
        On Error Resume Next
    
        .flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        .FileName = ""
        .MaxFileSize = 32767
        .CancelError = True
        .ShowSave
    
        If Err.Number = 0 And .FileName <> "" Then
    
            strTmp = .FileName
    
            On Error GoTo errHand
                                                    
            SaveDialog = strTmp
            
        Else
            Err.Clear
        End If
    End With
    
    Exit Function

errHand:
    MsgBox "���ܱ���Ϊ�ļ�(" & strTmp & "),���ļ���������ʹ�û��ļ��Ѿ�����!"
End Function

Private Sub cmd_Click(Index As Integer)
    Dim strFile As String
    
    Select Case Index
    Case 0
        strFile = OpenDialog(cdl, "��ѡ�������ļ�", "�����ļ�(*.ini)|*.ini")
        
        If strFile <> "" Then
            txt(0).Text = strFile
            mblnStep(1) = CheckSetupFile(strFile)
        End If
    Case 1
        strFile = SaveDialog(cdl, "��ѡ����־�ļ�", "��־�ļ�(*.log)|*.log")
        
        If strFile <> "" Then
            txt(1).Text = strFile
        End If
    End Select
    
End Sub

Private Function CheckPassword(ByVal strUser As String, ByVal strPassword As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    CheckPassword = mclsOracle.OraDataOpen(gstrServerName, strUser, strPassword, True)
End Function

Private Function CheckSetupFile(ByVal strFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ������Ͱ�װ�����ļ�����ȷ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strIniPath As String
    Dim strTemp As String
    Dim objText As TextStream
    Dim strManageVersion As String
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim aryTemp As Variant
    Dim aryItem As Variant
    Dim aryFlag As Variant
    Dim strSys As String
    Dim rsData As ADODB.Recordset
    Dim dblDataVer As String
    Dim dblFileVer As String
    Dim aryFile() As Variant
    Dim strPath As String
    
    strIniPath = Mid(strFile, 1, Len(strFile) - 11)
    
    '����ļ�ƥ���Լ��
    '------------------------------------------------------------------------------------------------------------------
    strTemp = ""
'    If Dir(strIniPath & "zlMipClientStruct.SQL") = "" Then strTemp = strTemp & vbCr & "�ṹ�ļ�" & strIniPath & "zlMipClientStruct.SQL"
'    If Dir(strIniPath & "zlMipClientData.SQL") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & strIniPath & "zlMipClientData.SQL"
    
'    If strTemp <> "" Then
'        MsgBox "���°�װ������ļ���ʧ�����ܼ�����������" & strTemp, vbExclamation, gstrSysName
'        Exit Function
'    End If
    
    '��װ�����ļ�����
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error Resume Next
    Set objText = mobjFso.OpenTextFile(strFile)
    
    mstrVersion = ""
    mstrManageVersion = ""
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 6) = "[Ŀ��汾]" Then
        mstrVersion = Trim(Mid(strTemp, 7))
    Else
        Err.Raise 10
    End If
    
    mstrDBVersion = gclsBusiness.Get_Ver
    '�Ƚϰ汾���Ƿ���Ҫ����
    dblDataVer = GetVerDouble(mstrDBVersion)
    dblFileVer = GetVerDouble(mstrVersion)
    
    If dblDataVer >= dblFileVer Then
        MsgBox "��ǰ�ͻ����������°汾��", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '��������ϵͳ�б�
    vsfUser.Rows = 1
    With vsfStep
        .Rows = 0
        strTemp = Trim(objText.ReadLine)
        .Rows = .Rows + 1
        .TextMatrix(0, .ColIndex("step")) = "��������ҵ������"
        .TextMatrix(0, .ColIndex("code")) = ""
        .Cell(flexcpPicture, 0, vsfStep.ColIndex("ͼ��")) = imgList.ListImages("��ִ��").Picture
        If Left(strTemp, 5) = "[������]" Then
            strTemp = Trim((Mid(strTemp, 7)))
            aryTemp = Split(strTemp, "|")
            For intLoop = 0 To UBound(aryTemp)
                aryItem = Split(aryTemp(intLoop), "=")
                aryFlag = Split(aryItem(1), ",")
                If gclsBusiness.CheckSetuped(aryFlag(0)) = 1 Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("step")) = "����" & aryItem(0) & "��Ϣ����"
                    .TextMatrix(.Rows - 1, .ColIndex("code")) = aryFlag(0)
                    .Cell(flexcpPicture, vsfStep.Rows - 1, vsfStep.ColIndex("ͼ��")) = imgList.ListImages("��ִ��").Picture
                    
                    '�ж��Ƿ��°���Ӳ���
                    If aryFlag(0) = "EMR" And aryFlag(2) = 2 Then
                        vsfUser.Rows = vsfUser.Rows + 1
                        Set rsData = GetSetupInfo(aryFlag(0))
                        If Not (rsData Is Nothing) Then
                            vsfUser.TextMatrix(vsfUser.Rows - 1, vsfUser.ColIndex("�û�")) = rsData("data_owner").Value
                            vsfUser.TextMatrix(vsfUser.Rows - 1, vsfUser.ColIndex("ϵͳ")) = aryItem(0)
                            vsfUser.TextMatrix(vsfUser.Rows - 1, vsfUser.ColIndex("����")) = GetXMLNode("ip", rsData("data_source").Value) & "/" & GetXMLNode("db_instance", rsData("data_source").Value) & "/" & GetXMLNode("db_owner", rsData("data_source").Value) & "/" & GetXMLNode("port", rsData("data_source").Value)
                            vsfUser.Cell(flexcpPicture, vsfUser.Rows - 1, vsfUser.ColIndex("ͼ��")) = imgList.ListImages("����").Picture
                            mclsVsfUser.AppendRows = True
                            mblnEMR = True
                        End If
                    End If
                End If
            Next
            
            '��������DB�ļ�
            
            strPath = Left(txt(0).Text, Len(txt(0).Text) - Len("zlUpgrade.ini"))
            
            '��������DB
            aryFile = GetFileList(strPath, "", GetVerDouble(mstrVersion), "zlMipServerData", "db")
            If UBound(aryFile) > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("step")) = "���빫��������Ϣ����"
                .TextMatrix(.Rows - 1, .ColIndex("code")) = "m_PUBLIC"
                .TextMatrix(.Rows - 1, .ColIndex("pname")) = GetVerString(aryFile(UBound(aryFile)))
                .Cell(flexcpPicture, vsfStep.Rows - 1, vsfStep.ColIndex("ͼ��")) = imgList.ListImages("��ִ��").Picture
            End If
            
            '����ҵ�񲿷�DB
            For intLoop = 0 To UBound(aryTemp)
                aryItem = Split(aryTemp(intLoop), "=")
                aryFlag = Split(aryItem(1), ",")
                If gclsBusiness.CheckSetuped(aryFlag(0)) = 1 Then
                    '����Ƿ����������DB�ļ�
                    aryFile = GetFileList(strPath, aryFlag(0), GetVerDouble(mstrVersion), "zlMipServerData", "db")
                    If UBound(aryFile) > 0 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, .ColIndex("step")) = "����" & aryItem(0) & "��Ϣ����"
                        .TextMatrix(.Rows - 1, .ColIndex("code")) = "m_" & aryFlag(0)
                        .TextMatrix(.Rows - 1, .ColIndex("pname")) = GetVerString(aryFile(UBound(aryFile)))
                        .Cell(flexcpPicture, vsfStep.Rows - 1, vsfStep.ColIndex("ͼ��")) = imgList.ListImages("��ִ��").Picture
                    End If
                End If
            Next
        End If
    End With
    
    lbl(2).Caption = "Ŀ��汾��" & mstrVersion
        
    objText.Close
    
    
    CheckSetupFile = True
End Function

Private Function GetSetupInfo(ByVal strCode As String) As ADODB.Recordset
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select data_owner,data_source From zlmip_data_setup Where data_code='" & strCode & "'"
    Set rsData = zlDataBase.OpenSQLRecord(strSQL, gstrSysName)
    If rsData.BOF = False Then
        Set GetSetupInfo = rsData
    Else
        Set GetSetupInfo = Nothing
    End If
    
End Function

Private Function GetXMLNode(ByVal strNode As String, strContent) As String
    '��ȡ�ڵ�����
    Dim lngLoop As Long
    Dim lngLength As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim blnFlag As Boolean
    
    lngLength = Len(strNode) + 2
    For lngLoop = 1 To Len(strContent)
        If Mid(strContent, lngLoop, lngLength) = "<" & strNode & ">" Then
            lngStart = lngLoop + lngLength
            blnFlag = True
        End If
        
        If blnFlag Then
            If Mid(strContent, lngLoop, 2) = "</" Then
                lngEnd = lngLoop
                Exit For
            End If
        End If
    Next
    
    If lngStart <> 0 And lngEnd <> 0 Then
        GetXMLNode = Mid(strContent, lngStart, lngEnd - lngStart)
    End If
    
End Function

Private Function GetSysOwner(ByVal strOwners As String) As ADODB.Recordset
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select Distinct ������ From zlSystems Where Upper(������) In(" & strOwners & ")"
    Set rsData = zlDataBase.OpenSQLRecord(strSQL, gstrSysName)
    If rsData.BOF = False Then
        Set GetSysOwner = rsData
    Else
        Set GetSysOwner = Nothing
    End If
    
End Function

Private Function GetVerDouble(ByVal varVer As Variant) As Double
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    '���ܣ����ݰ汾�ַ������������ֻ��İ汾
    '������varVer   �汾�ַ�������9.5.0
    Dim varArray As Variant
    
    varVer = IIf(IsNull(varVer), "", varVer)
    varArray = Split(varVer, ".")
    
    If UBound(varArray) < 2 Then Exit Function
    
    GetVerDouble = Val(varArray(0)) * 10 ^ 8 + Val(varArray(1)) * 10 ^ 4 + Val(varArray(2))
End Function

Public Function GetVerString(ByVal dblVer As Double) As String
'���ܣ��������ֻ��İ汾�������汾�ַ���
'������dblVer   �汾�ַ�������900050000
    
    GetVerString = dblVer \ 10 ^ 8 & "." & (dblVer Mod 10 ^ 8) \ 10 ^ 4 & "." & dblVer Mod 10 ^ 4
End Function

Private Function GetFileList(ByVal strFolderPath As String, ByVal strCode As String, ByVal strEndVer As Double, ByVal strFirstName, ByVal strLastName As String) As Variant
    Dim aryFile As Variant
    Dim aryVer As Variant
    Dim strPath As String
    Dim objFolder As Scripting.Folder
    Dim objFile As Scripting.File
    Dim strFileName As String
    Dim strLine As String
    Dim intPosition As Integer
    Dim i As Integer
    Dim j As Integer
    Dim varVer As String
    
    On Error Resume Next
    
    If strFolderPath = "" Then Exit Function
    If strCode <> "" Then
        strFolderPath = strFolderPath & strCode & "\"
    End If
    Set objFolder = gobjFile.GetFolder(strFolderPath)
    If Err.Number <> 0 Then
        MsgBox "�������ű���Žű�Ŀ¼����"
        Exit Function
    End If
    ReDim aryVer(0 To 0) '���ȳ�ʼ��
    For Each objFile In objFolder.Files
        Select Case strCode
        Case "" 'ϵͳ���������ļ�
            If (UCase(objFile.Name) Like UCase(strFirstName & "_*." & strLastName)) Then
                '�����ļ�
                intPosition = Len(strFirstName & "_") + 1
                strLine = Mid(objFile.Name, intPosition)
                If UCase(strLastName) = "DB" Then
                    strLine = Mid(strLine, 1, Len(strLine) - 3)
                Else
                    strLine = Mid(strLine, 1, Len(strLine) - 4)
                End If
                
                If GetVerDouble(strLine) <= strEndVer Then
                    i = UBound(aryVer) + 1
                    ReDim Preserve aryVer(0 To i)
                    aryVer(i) = GetVerDouble(strLine)
                End If
            End If
        Case Else
            If UCase(objFile.Name) Like UCase(strFirstName & "_*." & strLastName) Then
                '�����ļ�
                intPosition = Len(strFirstName & "_") + 1
                strLine = Mid(objFile.Name, intPosition)
                If UCase(strLastName) = "DB" Then
                    strLine = Mid(strLine, 1, Len(strLine) - 3)
                Else
                    strLine = Mid(strLine, 1, Len(strLine) - 4)
                End If
                
                If GetVerDouble(strLine) <= strEndVer Then
                    i = UBound(aryVer) + 1
                    ReDim Preserve aryVer(0 To i)
                    aryVer(i) = GetVerDouble(strLine)
                End If
            End If
        End Select
    Next

    '���汾��С���������
    For i = 1 To UBound(aryVer) - 1
        For j = i + 1 To UBound(aryVer)
            If aryVer(i) > aryVer(j) Then
                varVer = aryVer(i)
                aryVer(i) = aryVer(j)
                aryVer(j) = varVer
            End If
        Next
    Next
    
    GetFileList = aryVer
    
End Function

Public Function UpdateMipClient(ByVal strInstallFile As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strPath As String
    Dim intPercent As Integer
    Dim intRow As Integer
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim intFlag As Integer
    Dim i As Integer
    Dim strSQL As String
    Dim aryFile As Variant
    Dim rsErr As ADODB.Recordset
    
    On Error GoTo errHand
    
    strPath = Left(strInstallFile, Len(strInstallFile) - Len("zlUpgrade.ini"))
    
    Set mobjScript = New clsOracleScript
    lbl(12).Visible = True
    lbl(6).Visible = True
    pgb.Visible = True
    
    
    
    '��������
    '������������������
    With vsfStep
        .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("ִ����").Picture
        
        '�����ṹ����
        aryFile = GetFileList(strPath, "", GetVerDouble(mstrVersion), "zlMipClient", "SQL")
        For i = 1 To UBound(aryFile)
            If aryFile(i) > GetVerDouble(mstrDBVersion) And aryFile(i) <= GetVerDouble(mstrVersion) Then
                If mobjScript.OpenScriptFile(strPath & "zlMipClient_" & GetVerString(aryFile(i)) & ".SQL") Then
                    lbl(12).Caption = "���������������ֽṹ����..."
                    For intLoop = 1 To mobjScript.SQLCount
                        Call mobjScript.ExecuteSQL(gclsMsgOracle.DatabaseConnection, mobjScript.SQL(intLoop))
                        intPercent = 100 * intLoop / mobjScript.SQLCount
                        If pgb.Value <> intPercent Then pgb.Value = intPercent
                        lbl(6).Caption = intPercent & "%"
                    Next
                End If
            End If
        Next
        
        '�������ݲ���
        aryFile = GetFileList(strPath, "", GetVerDouble(mstrVersion), "zlMipClientData", "SQL")
        For i = 1 To UBound(aryFile)
            If aryFile(i) > GetVerDouble(mstrDBVersion) And aryFile(i) <= GetVerDouble(mstrVersion) Then
                If mobjScript.OpenScriptFile(strPath & "zlMipClientData_" & GetVerString(aryFile(i)) & ".SQL") Then
                    lbl(12).Caption = "���������������ֽṹ����..."
                    For intLoop = 1 To mobjScript.SQLCount
                        Call mobjScript.ExecuteSQL(gclsMsgOracle.DatabaseConnection, mobjScript.SQL(intLoop))
                        intPercent = 100 * intLoop / mobjScript.SQLCount
                        If pgb.Value <> intPercent Then pgb.Value = intPercent
                        lbl(6).Caption = intPercent & "%"
                    Next
                End If
            End If
        Next
        .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("�����").Picture
        intCount = intCount + 1
        
        '����ҵ�񲿷�
        For intRow = 1 To .Rows - 1
            If Mid(.TextMatrix(intRow, .ColIndex("code")), 1, 2) <> "m_" Then
                aryFile = GetFileList(strPath, .TextMatrix(intRow, .ColIndex("code")), GetVerDouble(mstrVersion), "zlMipClientData", "SQL")
                .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("ִ����").Picture
                For i = 1 To UBound(aryFile)
                    If aryFile(i) > GetVerDouble(mstrDBVersion) And aryFile(i) <= GetVerDouble(mstrVersion) Then
                        If mobjScript.OpenScriptFile(strPath & .TextMatrix(intRow, .ColIndex("code")) & "\zlMipClientData" & "_" & GetVerString(aryFile(i)) & ".SQL") Then
                            lbl(12).Caption = "��������" & .TextMatrix(intRow, .ColIndex("code")) & "�������ݽű�..."
                            vsfStep.Cell(flexcpPicture, intCount, vsfStep.ColIndex("ͼ��")) = imgList.ListImages("ִ����").Picture
                            For intLoop = 1 To mobjScript.SQLCount
                                Call mobjScript.ExecuteSQL(gclsMsgOracle.DatabaseConnection, mobjScript.SQL(intLoop))
                                intPercent = 100 * intLoop / mobjScript.SQLCount
                                If pgb.Value <> intPercent Then pgb.Value = intPercent
                                lbl(6).Caption = intPercent & "%"
                            Next
                        End If
                    End If
                Next
                .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("�����").Picture
                intCount = intCount + 1
            End If
        Next
    End With
    
    intFlag = intCount
0:
    With vsfStep
        intCount = intFlag
        '��ʼ�������¼��
        Set rsErr = Nothing
        If rsErr Is Nothing Then
            Set rsErr = New ADODB.Recordset
            rsErr.Fields.Append "���", adBSTR
            rsErr.Fields.Append "����", adBSTR
            rsErr.Open
        End If
        
        '���빫����������
        For intRow = 0 To .Rows - 1
            If Mid(.TextMatrix(intRow, .ColIndex("code")), 1, 2) = "m_" Then
                If Mid(.TextMatrix(intRow, .ColIndex("code")), 3) = "PUBLIC" Then
                    If Dir(strPath & "zlMipServerData_" & .TextMatrix(intRow, .ColIndex("pname")) & ".DB") <> "" Then
                        lbl(12).Caption = "��������������빫��������Ϣ..."
                        .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("ִ����").Picture
                        Call mclsMipClientManage.CommunicateProxyInstall(strPath & Mid(.TextMatrix(intRow, .ColIndex("code")), 3) & "\zlMipServerData_" & .TextMatrix(intRow, .ColIndex("pname")) & ".DB", rsErr)
                        .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("�����").Picture
                        intCount = intCount + 1
                    End If
                Else
                    If Dir(strPath & Mid(.TextMatrix(intRow, .ColIndex("code")), 3) & "\zlMipServerData_" & .TextMatrix(intRow, .ColIndex("pname")) & ".DB") <> "" Then
                        lbl(12).Caption = "���������������" & Mid(.TextMatrix(intRow, .ColIndex("code")), 3) & "��Ϣ..."
                        .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("ִ����").Picture
                        Call mclsMipClientManage.CommunicateProxyInstall(strPath & Mid(.TextMatrix(intRow, .ColIndex("code")), 3) & "\zlMipServerData_" & .TextMatrix(intRow, .ColIndex("pname")) & ".DB", rsErr)
                        .Cell(flexcpPicture, intCount, .ColIndex("ͼ��")) = imgList.ListImages("�����").Picture
                        intCount = intCount + 1
                    End If
                End If
            End If
        Next
        
    End With
    
    If Not (rsErr Is Nothing) Then
        If rsErr.RecordCount > 0 Then
            If mfrmErrorInfo Is Nothing Then
                Set mfrmErrorInfo = New frmErrorInfo
            End If
            
            If mfrmErrorInfo.ShowError(Me, rsErr) = False Then
                GoTo 0
            End If
        End If
    End If
    If mblnImportDB Then
        Call mclsMipClientManage.CommunicateProxyLogout
    End If
    Set mclsMipClientManage = Nothing
    
    '�޸İ�װ�汾
    strSQL = "Update zlRegInfo Set ����='" & mstrVersion & "' Where ��Ŀ='��Ϣ����ƽ̨�ͻ���' And �к�=1"
    gclsMsgOracle.ExecuteSQL strSQL, gstrSysName
    lbl(12).Caption = "�����������!"
    UpdateMipClient = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If MsgBox("�������д����Ƿ������" & vbCrLf & "    " & Err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
    
    'ж���Ѿ���װ������
    '------------------------------------------------------------------------------------------------------------------
'    lbl(4).Caption = "����ж���Ѿ���װ������..."
    
End Function

Private Sub Command2_Click()

End Sub

Private Sub cmdNext_Click()
    Dim intLoop As Integer
    
    Select Case mintPage
    '------------------------------------------------------------------------------------------------------------------
    Case 1
        
        If txt(0).Text = "" Then
            ShowSimpleMsg "����ѡ����Ϣ����ƽ̨�ͻ��˰�װ�����ļ���"
            Exit Sub
        End If
                
        If Dir(txt(0).Text) = "" Then
            ShowSimpleMsg "ѡ����Ϣ����ƽ̨�ͻ��˰�װ�����ļ������ڻ�����Ч��"
            Exit Sub
        End If
        
        If mblnEMR = False Then
            mintPage = mintPage + 1
        Else
            '��֤����
        End If
        
        mintPage = mintPage + 1
        Call ShowPage(mintPage)
        
        '��ʼ��һҳ
        
    '------------------------------------------------------------------------------------------------------------------
    Case 2
        
'        '���ZLTOOLS������Ч��
'        If CheckPassword("ZLTOOLS", txt(2).Text) = False Then
'            Exit Sub
'        End If
        Dim clsOracle As New clsDataOracle
        Dim aryData() As String
        With vsfUser
            For intLoop = 1 To .Rows - 1
                aryData = Split(.TextMatrix(intLoop, .ColIndex("����")), "/")
                If OraDataOpen(aryData(0), aryData(1), aryData(2), .TextMatrix(intLoop, .ColIndex("������")), aryData(3)) = False Then
                    MsgBox .TextMatrix(intLoop, .ColIndex("ϵͳ")) & "�޷����ӵ�������", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
            Next
        End With
        mintPage = mintPage + 1
        Call ShowPage(mintPage)
    Case 3
        '��֤���ӣ��Ƿ���������Ϣ������
        Set mclsMipClientManage = Nothing
        If mclsMipClientManage Is Nothing Then
            Set mclsMipClientManage = New clsMipClientManage
        End If
        
        If mclsMipClientManage.CommunicateProxyLogin(txt(6).Text, txt(7).Text, txt(9).Text, txt(8).Text) = False Then
            mblnImportDB = False
            Exit Sub
        Else
            mblnImportDB = True
        End If
        
        mintPage = mintPage + 1
        Call ShowPage(mintPage)
    End Select
    
    
End Sub

Private Function OraDataOpen(ByVal strServerIP As String, ByVal strSID As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal strPort As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    Dim strServer As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strServerIP & ")(PORT = " & strPort & ")))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If zlComLib.ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function

Private Sub cmdPrevious_Click()

    Select Case mintPage
    '------------------------------------------------------------------------------------------------------------------
    Case 2, 4
        
        mintPage = mintPage - 1
        Call ShowPage(mintPage)
    Case 3
        '�ж��Ƿ�ѡ���˲���
        If mblnEMR = True Then
            mintPage = mintPage - 1
        Else
            mintPage = mintPage - 2
        End If
        Call ShowPage(mintPage)
    End Select
    
End Sub

Private Sub cmdUpgrade_Click()
    '��ȡ����ϵͳ��Ӧ�����������ű�
    If MsgBox("ȷ����Ҫ������Ϣ����ƽ̨�ͻ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    '�����ű�
    If UpdateMipClient(txt(0).Text) Then
        MsgBox "��Ϣ���������ɹ�!", vbInformation + vbOKOnly, "��ʾ��Ϣ"
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strValue As String
    Dim strArr() As String
    
    '��ע���
    strValue = GetSetting("ZLSOFT", "����ȫ��\��Ϣƽ̨�ͻ���", "��Ϣ����������", "")
    If strValue <> "" Then
        strArr = Split(strValue, ";")
        txt(6).Text = strArr(0)
        txt(7).Text = strArr(1)
        txt(9).Text = strArr(2)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If Not (mclsOracle Is Nothing) Then
'        Set mclsOracle = Nothing
'    End If
'
'    Dim frmThis As Form
'
'    On Error Resume Next
'
'    '�رձ���������
'    For Each frmThis In Forms
'        If frmThis.Caption <> Me.Caption Then
'            Unload frmThis
'        End If
'    Next
    
End Sub

Private Sub mclsMipClientManage_AfterCommunicateChange(ByVal strTitle As String, ByVal strNumber As String)
    lbl(12).Caption = strTitle
    lbl(6).Caption = strNumber & "%"
    pgb.Value = strNumber
End Sub

Private Sub mobjScript_AfterAnalyseLine(ByVal Line As Long, ByVal Lines As Long)
    Dim intPercent As Integer
    
'    If pgb.Visible = False Then pgb.Visible = True
'    If lbl(4).Visible = False Then
'        lbl(4).Visible = True
'        lbl(4).Caption = "���ڷ����ű��ļ�...."
'    End If
'
'    intPercent = 100 * Line / Lines
'    If pgb.Value <> intPercent Then pgb.Value = intPercent
'
End Sub

Private Sub mobjScript_BeforeAnalyseLine(ByVal Line As Long, ByVal Lines As Long)
'    If pgb.Visible = False Then pgb.Visible = True
'    If lbl(4).Visible = False Then
'        lbl(4).Visible = True
'        lbl(4).Caption = "���ڷ����ű��ļ�...."
'    End If
End Sub

Private Sub vsfUser_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsfUser.AfterEdit(Row, Col)
End Sub

Private Sub vsfUser_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfUser.BeforeEdit(Row, Col, Cancel)
    If Cancel = False Then
        If Col = vsfUser.ColIndex("����") Then
            mstrBeforeEditStr = vsfUser.TextMatrix(Row, vsfUser.ColIndex("������"))
        End If
    End If
End Sub

Private Sub vsfUser_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim intTruePwd As Integer
    If Col <> vsfUser.ColIndex("����") Then Exit Sub
    
    intTruePwd = vsfUser.ColIndex("������")
    
    If KeyAscii = vbKeyEscape Then
        vsfUser.TextMatrix(Row, intTruePwd) = mstrBeforeEditStr
        Exit Sub
    End If
    
    If KeyAscii = 8 And Len(vsfUser.TextMatrix(Row, intTruePwd)) > 0 Then
        vsfUser.TextMatrix(Row, intTruePwd) = Mid(vsfUser.TextMatrix(Row, intTruePwd), 1, Len(vsfUser.TextMatrix(Row, intTruePwd)) - 1)
    Else
        If KeyAscii <> vbKeyReturn Then
            vsfUser.TextMatrix(Row, intTruePwd) = vsfUser.TextMatrix(Row, intTruePwd) & Chr(KeyAscii)
            KeyAscii = 42
        End If
    End If
End Sub

