VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParaSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   Icon            =   "frmParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSet 
      Caption         =   "�豸(&V)"
      Height          =   350
      Left            =   5970
      TabIndex        =   65
      ToolTipText     =   "�������������豸"
      Top             =   6855
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8310
      TabIndex        =   64
      Top             =   6855
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7140
      TabIndex        =   63
      Top             =   6855
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6585
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   11615
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   2822
      TabCaption(0)   =   "���շ���"
      TabPicture(0)   =   "frmParaSet.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmReceiveSet"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkStorage"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "WorkList"
      TabPicture(1)   =   "frmParaSet.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmWorkList"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkDWL"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Query/Retrieve"
      TabPicture(2)   =   "frmParaSet.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmQueryRetrieve"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkQuery"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "�������ݿ�"
      TabPicture(3)   =   "frmParaSet.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "�����ʱ��"
         Height          =   2535
         Left            =   120
         TabIndex        =   55
         Top             =   700
         Width           =   9255
         Begin VB.CommandButton cmdClear 
            Caption         =   "�������"
            Height          =   350
            Left            =   240
            TabIndex        =   62
            Top             =   2040
            Width           =   1100
         End
         Begin VB.TextBox txtClearInterval 
            Height          =   300
            Left            =   960
            MaxLength       =   3
            TabIndex        =   61
            Top             =   1485
            Width           =   975
         End
         Begin VB.CheckBox chkAutoClear 
            Caption         =   "���                           �죬�Զ����"
            Height          =   375
            Left            =   240
            TabIndex        =   60
            Top             =   1440
            Width           =   4575
         End
         Begin VB.Frame Frame5 
            Caption         =   "���ݿ��"
            Height          =   855
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   8775
            Begin VB.CheckBox chkClearTempTB 
               Caption         =   "Ӱ���������"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   57
               Top             =   360
               Width           =   1815
            End
            Begin VB.CheckBox chkClearTempTB 
               Caption         =   "������־"
               Height          =   255
               Index           =   1
               Left            =   2640
               TabIndex        =   58
               Top             =   360
               Width           =   1815
            End
            Begin VB.CheckBox chkClearTempTB 
               Caption         =   "DICOMͨѶ��־"
               Height          =   255
               Index           =   2
               Left            =   5280
               TabIndex        =   59
               Top             =   360
               Width           =   2775
            End
         End
      End
      Begin VB.CheckBox chkStorage 
         Caption         =   "����ͼ����շ���"
         Height          =   195
         Left            =   -74850
         TabIndex        =   0
         Top             =   500
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Frame frmReceiveSet 
         Height          =   5715
         Left            =   -74850
         TabIndex        =   69
         Top             =   700
         Width           =   9255
         Begin VB.Frame Frame3 
            Caption         =   "�Զ�ƥ������"
            Height          =   2055
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   8895
            Begin VB.CheckBox chkImageType 
               Caption         =   "����ͼ�����Ͳ������"
               Height          =   350
               Left            =   4440
               TabIndex        =   75
               Top             =   240
               Width           =   3015
            End
            Begin VB.CheckBox chkMatchStudyUID 
               Caption         =   "���� ""���UID"" ƥ��"
               Height          =   350
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   3015
            End
            Begin VB.Frame Frame2 
               Caption         =   "���ݿ���Ŀ"
               Height          =   1335
               Left            =   4440
               TabIndex        =   14
               Top             =   600
               Width           =   4365
               Begin VB.OptionButton optMatch 
                  Caption         =   "�� ""����ʶ��"" ƥ��"
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   17
                  ToolTipText     =   "������ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
                  Top             =   960
                  Width           =   2775
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "�� ""���˱�ʶ�ţ�����/סԺ�ţ�"" ƥ��"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   16
                  ToolTipText     =   "�����˱�ʶ�Ž����˺ͽ��յ�Ӱ�����ƥ��"
                  Top             =   600
                  Width           =   3975
               End
               Begin VB.OptionButton optMatch 
                  Caption         =   "�� ""����"" ƥ��"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   15
                  ToolTipText     =   "�����Ž����˺ͽ��յ�Ӱ�����ƥ��"
                  Top             =   240
                  Width           =   2265
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "ͼ����Ŀ"
               Height          =   1335
               Left            =   120
               TabIndex        =   10
               Top             =   600
               Width           =   4250
               Begin VB.OptionButton optImgMatch 
                  Caption         =   "Patient Name"
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   13
                  Top             =   960
                  Width           =   2055
               End
               Begin VB.OptionButton optImgMatch 
                  Caption         =   "Accession Number"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   12
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.OptionButton optImgMatch 
                  Caption         =   "Patient ID"
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   11
                  Top             =   240
                  Width           =   2055
               End
            End
         End
         Begin VB.Frame frmAutoRoutSet 
            Caption         =   "�Զ�·������"
            Height          =   2385
            Left            =   120
            TabIndex        =   18
            Top             =   3210
            Width           =   9015
            Begin VB.CommandButton cmdInsert 
               Caption         =   "���(&A)"
               Height          =   350
               Left            =   1800
               TabIndex        =   26
               Top             =   1890
               Width           =   1100
            End
            Begin VB.CommandButton cmdModify 
               Caption         =   "�޸�(&M)"
               Height          =   350
               Left            =   3660
               TabIndex        =   27
               Top             =   1890
               Width           =   1100
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "ɾ��(&D)"
               Height          =   350
               Left            =   5520
               TabIndex        =   28
               Top             =   1890
               Width           =   1100
            End
            Begin VB.OptionButton optType 
               Caption         =   "Ӱ�����(&S)"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   20
               Top             =   1455
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optType 
               Caption         =   "����豸(&R)"
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   22
               Top             =   1455
               Width           =   1335
            End
            Begin VB.ComboBox cobCondition 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   4530
               TabIndex        =   23
               Top             =   1425
               Width           =   1365
            End
            Begin VB.ComboBox cobCondition 
               Height          =   315
               Index           =   1
               Left            =   1530
               TabIndex        =   21
               Top             =   1440
               Width           =   1365
            End
            Begin VB.ComboBox cobDestination 
               Height          =   315
               Left            =   7290
               TabIndex        =   25
               Top             =   1425
               Width           =   1605
            End
            Begin MSFlexGridLib.MSFlexGrid MSFAutoRout 
               Height          =   1125
               Left            =   150
               TabIndex        =   19
               Top             =   240
               Width           =   8775
               _ExtentX        =   15478
               _ExtentY        =   1984
               _Version        =   393216
               FixedCols       =   0
               SelectionMode   =   1
               AllowUserResizing=   1
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Ŀ���豸(&B)"
               Height          =   180
               Left            =   6150
               TabIndex        =   24
               Top             =   1485
               Width           =   990
            End
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   255
            Left            =   8760
            TabIndex        =   71
            TabStop         =   0   'False
            ToolTipText     =   "ѡ����ʱĿ¼"
            Top             =   750
            Width           =   285
         End
         Begin VB.ComboBox cboEncode 
            Height          =   300
            ItemData        =   "frmParaSet.frx":007C
            Left            =   6240
            List            =   "frmParaSet.frx":0089
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   2835
         End
         Begin VB.TextBox txtItem 
            BackColor       =   &H80000009&
            DataField       =   "315"
            Height          =   300
            Index           =   0
            Left            =   1320
            MaxLength       =   5
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   280
            Width           =   855
         End
         Begin VB.ComboBox cboDevice 
            Height          =   300
            ItemData        =   "frmParaSet.frx":00AC
            Left            =   3315
            List            =   "frmParaSet.frx":00B9
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   300
            Width           =   1575
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   465
            Left            =   150
            TabIndex        =   70
            Top             =   1050
            Width           =   8925
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   1
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   200
            TabIndex        =   8
            Top             =   720
            Width           =   7740
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "��ʱĿ¼(&T)"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   780
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ѹ����ʽ(&Y)"
            Height          =   180
            Index           =   0
            Left            =   5160
            TabIndex        =   5
            Top             =   345
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "�����˿�(&P)"
            Height          =   180
            Left            =   240
            TabIndex        =   1
            Top             =   345
            Width           =   990
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            Caption         =   "�洢�豸(&F)"
            Height          =   180
            Index           =   8
            Left            =   2280
            TabIndex        =   3
            Top             =   345
            Width           =   990
         End
      End
      Begin VB.CheckBox chkDWL 
         Caption         =   "���� DICOM WorkList ����"
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   2775
      End
      Begin VB.Frame frmWorkList 
         Enabled         =   0   'False
         Height          =   5790
         Left            =   -74880
         TabIndex        =   68
         Top             =   705
         Width           =   9255
         Begin VB.CommandButton cmdResetWLResult 
            Caption         =   "�ָ�Ĭ�Ͻ��"
            Height          =   350
            Left            =   6240
            TabIndex        =   38
            Top             =   600
            Width           =   1335
         End
         Begin VB.Frame Frame8 
            Caption         =   "���������"
            Height          =   4695
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   8895
            Begin VB.CheckBox chkUseResult 
               Caption         =   "ѡ��ʹ�øý����"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   3000
               Width           =   1935
            End
            Begin VB.Frame frmSetResult 
               Height          =   1575
               Left            =   120
               TabIndex        =   72
               Top             =   3000
               Width           =   8655
               Begin VB.CommandButton cmdBuildResult 
                  Appearance      =   0  'Flat
                  Caption         =   "��"
                  Height          =   235
                  Index           =   0
                  Left            =   7950
                  MaskColor       =   &H80000000&
                  Style           =   1  'Graphical
                  TabIndex        =   73
                  Top             =   765
                  Width           =   315
               End
               Begin VB.TextBox txtResult 
                  Height          =   300
                  Index           =   1
                  Left            =   1200
                  TabIndex        =   47
                  Top             =   1080
                  Width           =   7095
               End
               Begin VB.CheckBox chkResult 
                  Caption         =   "�Ƿ����"
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   43
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.TextBox txtResult 
                  Height          =   300
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   45
                  Top             =   720
                  Width           =   7095
               End
               Begin VB.Label Label12 
                  Caption         =   "ǿ�ƽ��ֵ"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   46
                  Top             =   1110
                  Width           =   975
               End
               Begin VB.Label Label11 
                  Caption         =   "����ֵ"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   44
                  Top             =   743
                  Width           =   735
               End
               Begin VB.Label lblResult 
                  Caption         =   "�������"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   42
                  Top             =   360
                  Width           =   7215
               End
            End
            Begin MSFlexGridLib.MSFlexGrid MSFResult 
               Height          =   2535
               Left            =   120
               TabIndex        =   40
               Top             =   360
               Width           =   8655
               _ExtentX        =   15266
               _ExtentY        =   4471
               _Version        =   393216
               AllowBigSelection=   0   'False
               SelectionMode   =   1
               AllowUserResizing=   1
            End
         End
         Begin VB.CheckBox chkForceResult 
            Caption         =   "ʹ��ǿ�ƽ��"
            Height          =   255
            Left            =   3600
            TabIndex        =   35
            Top             =   660
            Width           =   1515
         End
         Begin VB.CheckBox chkModel 
            Caption         =   "������豸����"
            Height          =   225
            Left            =   3600
            TabIndex        =   34
            Top             =   278
            Width           =   1755
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   6
            Left            =   7080
            MaxLength       =   4
            TabIndex        =   37
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtDWLLocalAE 
            Height          =   300
            Left            =   1080
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   637
            Width           =   1695
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   4
            Left            =   1080
            MaxLength       =   5
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "�������                �������"
            Height          =   195
            Left            =   6210
            TabIndex        =   36
            Top             =   300
            Width           =   2355
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "����AE"
            Height          =   180
            Left            =   195
            TabIndex        =   32
            Top             =   690
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "�����˿�"
            Height          =   180
            Left            =   195
            TabIndex        =   30
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.CheckBox chkQuery 
         Caption         =   "���� Query/Retrieve ��ѯ����"
         Height          =   255
         Left            =   -74850
         TabIndex        =   48
         Top             =   500
         Width           =   2955
      End
      Begin VB.Frame frmQueryRetrieve 
         Enabled         =   0   'False
         Height          =   975
         Left            =   -74880
         TabIndex        =   49
         Top             =   700
         Width           =   9255
         Begin VB.TextBox txtQueryAE 
            Height          =   300
            Left            =   4320
            MaxLength       =   20
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtItem 
            Height          =   300
            Index           =   5
            Left            =   1245
            MaxLength       =   5
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox chkAcceptCGET 
            Caption         =   "֧��C-GET"
            Height          =   255
            Left            =   7200
            TabIndex        =   54
            Top             =   380
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "����AE"
            Height          =   180
            Left            =   3600
            TabIndex        =   52
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "�����˿�"
            Height          =   180
            Left            =   360
            TabIndex        =   50
            Top             =   420
            Width           =   720
         End
      End
   End
End
Attribute VB_Name = "frmParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ifOK As Boolean
Private mblnchkResultFocus As Boolean
Private mblnchkUseResultFocus As Boolean

Private aDevices() As Variant
Private mintMaxDevs As Integer
Private mblnModifyMWLResult As Boolean          '��¼�Ƿ��޸���Worklist����ֵ������

Public Function ShowMe(objParent As Object, Optional iMaxDevs As Integer = 2) As Boolean
    mintMaxDevs = iMaxDevs
    Me.Show vbModal, objParent
    ShowMe = ifOK
End Function

Private Sub cboDevice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboEncode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkAcceptCGET_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkAutoClear_Click()
    If chkAutoClear.value = 1 Then
        txtClearInterval.Enabled = True
    Else
        txtClearInterval.Enabled = False
    End If
End Sub

Private Sub chkClearTempTB_Click(Index As Integer)
    Dim i As Integer
    gstrClearTable = "" '��ձ��б�
    For i = 0 To 2
        If chkClearTempTB(i).value = 1 Then
            gstrClearTable = gstrClearTable & IIf(Trim(gstrClearTable) = "", "", ";") & chkClearTempTB(i).Caption
        End If
    Next i
End Sub

Private Sub chkClearTempTB_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkDWL_Click()
    If Me.chkDWL.value = 0 Then
        Me.frmWorkList.Enabled = False
    Else
        Me.frmWorkList.Enabled = True
    End If
End Sub

Private Sub chkDWL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub




Private Sub chkQuery_Click()
    If Me.chkQuery.value = 0 Then
        Me.frmQueryRetrieve.Enabled = False
    Else
        Me.frmQueryRetrieve.Enabled = True
    End If
End Sub

Private Sub chkQuery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub chkResult_Click()
    If mblnchkResultFocus Then
        mblnModifyMWLResult = True
        subChangeMSFResult
    End If
End Sub

Private Sub chkResult_GotFocus()
    mblnchkResultFocus = True
End Sub

Private Sub chkResult_LostFocus()
    mblnchkResultFocus = False
End Sub


Private Sub chkStorage_Click()
    If Me.chkStorage.value = 0 Then
        Me.frmReceiveSet.Enabled = False
    Else
        Me.frmReceiveSet.Enabled = True
    End If
    
End Sub

Private Sub chkUseResult_Click()
    If mblnchkUseResultFocus Then
        mblnModifyMWLResult = True
        subChangeMSFResult
    End If
    If chkUseResult.value = 1 Then
        frmSetResult.Enabled = True
    Else
        frmSetResult.Enabled = False
    End If
End Sub

Private Sub chkUseResult_GotFocus()
    mblnchkUseResultFocus = True
End Sub

Private Sub chkUseResult_LostFocus()
    mblnchkUseResultFocus = False
End Sub

Private Sub cmdBuildResult_Click(Index As Integer)
    frmBuildResult.strReturnString = ""
    frmBuildResult.txtBuildResult.Text = Me.txtResult(Index).Text
    frmBuildResult.Show 1, Me
    If frmBuildResult.strReturnString <> "" Then
        Me.txtResult(Index).Text = frmBuildResult.strReturnString
        mblnModifyMWLResult = True
        subChangeMSFResult
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdClear_Click()
    subClearTempTable True
End Sub

Private Sub cmdDelete_Click()
    'ɾ���Զ�·�������е�ֵ
    Dim iRow As Integer
    Dim i As Integer
    iRow = MSFAutoRout.RowSel
    '�ƶ���������
    For i = iRow + 1 To UBound(aAutoRoutSetting)
        aAutoRoutSetting(i - 1).Type = aAutoRoutSetting(i).Type
        aAutoRoutSetting(i - 1).strCondition = aAutoRoutSetting(i).strCondition
        aAutoRoutSetting(i - 1).strFTPDeviceNo = aAutoRoutSetting(i).strFTPDeviceNo
    Next
    '�޸������С
    If UBound(aAutoRoutSetting) = 0 Then Exit Sub
    ReDim Preserve aAutoRoutSetting(0 To UBound(aAutoRoutSetting) - 1)
    'ˢ���Զ�·�ɹ�����ʾ�б�
    subFillMsfAutoRout
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdInsert_Click()
    '��������Ƿ�Ϸ�
    Dim iType As Integer
    iType = IIf(optType(1).value = True, 1, 2)
    If cobDestination.Text = "" Then MsgBox "�������Զ�·�ɵ�Ŀ���豸��": Exit Sub
    If cobCondition(iType).Text = "" Then MsgBox IIf(iType = 1, "������Ӱ�����", "���������豸"): Exit Sub
    '���Զ�·�ɹ�����������¹���
    Dim iCount As Integer
    iCount = UBound(aAutoRoutSetting) + 1
    
    ReDim a(2)
    a(1).Type = 4
    a(1).strCondition = "DFDFD"
    ReDim Preserve a(5)
    ReDim Preserve a(3)
    
    ReDim Preserve aAutoRoutSetting(0 To iCount)
    aAutoRoutSetting(iCount).Type = iType
    aAutoRoutSetting(iCount).strCondition = cobCondition(iType).Text
    aAutoRoutSetting(iCount).strFTPDeviceNo = GetDeviceNameNum(aDevices, cobDestination.Text, 1)
    '������б�����¹���
    With MSFAutoRout
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = IIf(aAutoRoutSetting(iCount).Type = 1, "Ӱ�����", "����豸")
        .TextMatrix(.Rows - 1, 1) = aAutoRoutSetting(iCount).strCondition
        .TextMatrix(.Rows - 1, 2) = GetDeviceNameNum(aDevices, aAutoRoutSetting(iCount).strFTPDeviceNo, 0)
    End With
End Sub

Private Sub cmdModify_Click()
    '�޸��Զ�·������
    Dim iType  As Integer
    Dim iRow As Integer
    iRow = MSFAutoRout.RowSel
    iType = IIf(optType(1).value = True, 1, 2)
    aAutoRoutSetting(iRow).Type = iType
    aAutoRoutSetting(iRow).strCondition = cobCondition(iType).Text
    aAutoRoutSetting(iRow).strFTPDeviceNo = GetDeviceNameNum(aDevices, cobDestination.Text, 1)
    '�޸��Զ�·�ɹ����б�
    MSFAutoRout.TextMatrix(iRow, 0) = IIf(aAutoRoutSetting(iRow).Type = 1, "Ӱ�����", "����豸")
    MSFAutoRout.TextMatrix(iRow, 1) = aAutoRoutSetting(iRow).strCondition
    MSFAutoRout.TextMatrix(iRow, 2) = GetDeviceNameNum(aDevices, aAutoRoutSetting(iRow).strFTPDeviceNo, 0)
End Sub

Private Sub CmdOK_Click()
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo DBError
    '����ͼ���������
    If Me.chkStorage.value = 1 Then
        If Len(Trim(txtItem(0))) = 0 Then
            MsgBox "������˿ںţ�", vbInformation, gstrSysName
            txtItem(0).SetFocus: Exit Sub
        End If
        If Len(Trim(txtItem(1))) = 0 Then
            MsgBox "��������ʱĿ¼��", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        If LenB(StrConv(Trim(txtItem(1).Text), vbFromUnicode)) > txtItem(1).MaxLength Then
            MsgBox "��ʱĿ¼���������" & txtItem(1).MaxLength & "���ַ���" & CInt(txtItem(1).MaxLength / 2) & "�����֣���", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�˿�", txtItem(0)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�豸��", aDevices(0, cboDevice.ListIndex)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "��ʱĿ¼", txtItem(1)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����ѹ��", cboEncode.ListIndex
        For i = 0 To optMatch.count - 1
            If optMatch(i).value Then Exit For
        Next
        If i > optMatch.count - 1 Then i = 0
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "���ݿ�ƥ�䷽ʽ", i
        
        For i = 0 To optImgMatch.count - 1
            If optImgMatch(i).value Then Exit For
        Next
        If i > optImgMatch.count - 1 Then i = 0
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "ͼ��ƥ�䷽ʽ", i
        
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "���ü��UIDƥ��", IIf(chkMatchStudyUID.value, 1, 0)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����ͼ�����Ͳ������", IIf(chkImageType.value, 1, 0)
        
        '�����Զ�·������
        Dim strAutoRoutSet As String
        If UBound(aAutoRoutSetting) >= 1 Then
            strAutoRoutSet = aAutoRoutSetting(1).Type
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(1).strCondition
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(1).strFTPDeviceNo
        End If
        For i = 2 To UBound(aAutoRoutSetting)
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(i).Type
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(i).strCondition
            strAutoRoutSet = strAutoRoutSet & "," & aAutoRoutSetting(i).strFTPDeviceNo
        Next
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�Զ�·��", strAutoRoutSet
    End If
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "��������", Me.chkStorage.value
    
    '����WorkList������
    If Me.chkDWL.value = 1 Then
        '������ȷ�Լ��
        If Len(Trim(txtItem(4))) = 0 Then
            MsgBox "������WorkList�˿ںţ�", vbInformation, gstrSysName
            txtItem(4).SetFocus: Exit Sub
        End If
        If Len(Trim(txtDWLLocalAE)) = 0 Then
            MsgBox "������WorkList�ı���AE���ƣ�", vbInformation, gstrSysName
            txtDWLLocalAE.SetFocus: Exit Sub
        End If
        '�����������
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList�˿�", txtItem(4)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList����AE", txtDWLLocalAE
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList��������", Val(txtItem(6))
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList���豸����", IIf(chkModel.value, 1, 0)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkListʹ��ǿ�ƽ��", IIf(chkForceResult.value, 1, 0)
        
        '����Worklist ����ֵ���޸�
        If mblnModifyMWLResult = True And gcnAccess.State <> adStateClosed Then
            With Me.MSFResult
                For i = 1 To .Rows - 1
                    strSQL = "update ǿ�ƽ�� set ����ֵ = '" & .TextMatrix(i, 4) & "' , ǿ�ƽ��ֵ='" _
                             & .TextMatrix(i, 5) & "' , �Ƿ���� = " & .TextMatrix(i, 6) _
                             & ",��ѡ�� = " & .TextMatrix(i, 3) & " where ��� = '" _
                             & Mid(.TextMatrix(i, 0), 2, InStr(.TextMatrix(i, 0), ",") - 2) & "' and Ԫ�غ� = '" _
                             & Mid(.TextMatrix(i, 0), InStr(.TextMatrix(i, 0), ",") + 1, Len(.TextMatrix(i, 0)) - InStr(.TextMatrix(i, 0), ",") - 1) & "'"
                    gcnAccess.Execute strSQL
                Next i
            End With
        End If
    End If
    
    '����Query/Retrieve������
    If Me.chkQuery.value = 1 Then
        '������ȷ�Լ��
        If Len(Trim(txtItem(5))) = 0 Then
            MsgBox "������Query/Retrieve�Ķ˿ںţ�", vbInformation, gstrSysName
            txtItem(5).SetFocus: Exit Sub
        End If
        If Len(Trim(txtQueryAE)) = 0 Then
            MsgBox "������Query/Retrieve�ı���AE���ƣ�", vbInformation, gstrSysName
            txtQueryAE.SetFocus: Exit Sub
        End If
        '�����������
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "Query/Retrieve�˿�", txtItem(5)
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "Query/Retrieve����AE", txtQueryAE
        SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "֧��C-GET", Me.chkAcceptCGET.value
    End If
    
    '���桰�������ݿ⡱������
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�Զ������ʱ��", gstrClearTable
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�Զ������ʱ��������", IIf(txtClearInterval.Enabled = True, Val(Me.txtClearInterval.Text), 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�Զ������ʱ������", Date
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����WorkList", Me.chkDWL.value
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����Query/Retrieve", Me.chkQuery.value
    ifOK = True
    Unload Me
    Exit Sub
DBError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdResetWLResult_Click()
    Dim strSQL As String
    
    If gcnAccess.State = adStateClosed Then Exit Sub
    strSQL = "update ǿ�ƽ�� set ����ֵ = Ĭ��ֵ,ǿ�ƽ��ֵ=Ĭ��ǿ�ƽ��,�Ƿ���� = False,��ѡ�� = Ĭ��ѡ��"
    gcnAccess.Execute strSQL
    subFillMsfResult
End Sub

Private Sub cmdSel_Click()
    Dim strTmp As String
    '�õ�·��
    strTmp = BrowPath(Me.hwnd, "��ѡ��Ӱ�񱣴����ʱĿ¼��")
    '�����µ�·��ʱ�ű���
    If strTmp <> "" Then
        If Mid(strTmp, Len(strTmp), 1) <> "\" Then strTmp = strTmp + "\"
        txtItem(1) = strTmp
    End If
End Sub

Private Sub cmdSet_Click()
    frmIPConfig.ShowEdit Me, mintMaxDevs
End Sub



Private Sub cobCondition_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cobDestination_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call CmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strExeRoom As String
    Dim strDeviceNO As String
    Dim iMatchStyle As Integer
    Dim iImgMatchStyle As Integer
    Dim strTemp As String
    Dim i As Integer
    
    ifOK = False
    
    On Error GoTo DBError
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����= [1]"
    Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption, 1)
    If rsTmp.EOF Then
        MsgBox "δ����Ӱ��洢�豸���뵽Ӱ���豸Ŀ¼�����ã�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    aDevices = rsTmp.GetRows: rsTmp.MoveFirst: strDeviceNO = rsTmp(0)
    Me.cboDevice.Clear
    Do While Not rsTmp.EOF
        cboDevice.AddItem Nvl(rsTmp(1))
        '����Զ�·�������е�Ŀ���豸�����б��ƽ�
        cobDestination.AddItem Nvl(rsTmp(1))
        rsTmp.MoveNext
    Loop
    
    txtItem(0) = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�˿�", 104)
    strDeviceNO = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�豸��", strDeviceNO)
    cboDevice.ListIndex = GetComboxIndex(aDevices, strDeviceNO)
    txtItem(1) = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "��ʱĿ¼", "C:\TmpImage\")
    cboEncode.ListIndex = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����ѹ��", 0))
    iMatchStyle = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "���ݿ�ƥ�䷽ʽ", 0))
    optMatch(iMatchStyle).value = True
    iImgMatchStyle = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "ͼ��ƥ�䷽ʽ", 0))
    optImgMatch(iImgMatchStyle).value = True
    chkMatchStudyUID.value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "���ü��UIDƥ��", 1))
    chkImageType.value = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����ͼ�����Ͳ������", 0))
    
    
    '����Զ�·������
    subFillMsfAutoRout
    '����Զ�·�������У�Ӱ����𣬺ͼ���豸�б�
    gstrSQL = "Select ���� From Ӱ�������"
    OpenRecordset rsTmp, Me.Caption
    Do While Not rsTmp.EOF
        cobCondition(1).AddItem rsTmp(0)
        rsTmp.MoveNext
    Loop
    
    gstrSQL = "Select distinct ����豸 From Ӱ�����¼"
    OpenRecordset rsTmp, Me.Caption
    Do While Not rsTmp.EOF
        cobCondition(2).AddItem Nvl(rsTmp(0))
        rsTmp.MoveNext
    Loop
    chkStorage.value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "��������", 1)
    
    '���WorkList�Ĳ���
    txtItem(4) = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList�˿�", 1024)
    txtDWLLocalAE = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList����AE", "ZLPACSWL")
    chkDWL.value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����WorkList", 0)
    txtItem(6) = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList��������", 3)
    chkModel = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkList���豸����", 0))
    chkForceResult = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "WorkListʹ��ǿ�ƽ��", 0))
    subFillMsfResult    '���Worklist����ֵ���ñ�
    
    '���Query/Retrieve�Ĳ���
    txtItem(5) = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "Query/Retrieve�˿�", 1024)
    txtQueryAE.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "Query/Retrieve����AE", "ZLPACSQR")
    chkQuery.value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "����Query/Retrieve", 0)
    chkAcceptCGET.value = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "֧��C-GET", 0)
    
    '��䡰�������ݿ⡱�Ĳ���
    strTemp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�Զ������ʱ��", "")
    Dim strTempArray() As String
    strTempArray = Split(strTemp, ";")
    For i = 0 To 2
        chkClearTempTB(i).value = 0
    Next i
    For i = 0 To UBound(strTempArray)
        If strTempArray(i) = "Ӱ���������" Then
            chkClearTempTB(0).value = 1
        ElseIf strTempArray(i) = "������־" Then
            chkClearTempTB(1).value = 1
        ElseIf strTempArray(i) = "DICOMͨѶ��־" Then
            chkClearTempTB(2).value = 1
        End If
    Next i
    txtClearInterval = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�Զ������ʱ��������", "0")
    If txtClearInterval <= 0 Then
        chkAutoClear.value = 0
        txtClearInterval.Enabled = False
    Else
        chkAutoClear.value = 1
        txtClearInterval.Enabled = True
    End If
    
    SetPrivs gstrPrivs
    
    SSTab1.Tab = 0
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub subFillMsfAutoRout()
    Dim lngRowPos As Long
    Dim i As Integer
    With MSFAutoRout
        .Clear
        .Rows = 1
        .Cols = 3
        .ColWidth(1) = 3000
        .TextMatrix(0, 0) = "��������"
        .TextMatrix(0, 1) = "��������"
        .TextMatrix(0, 2) = "Ŀ���豸"
        lngRowPos = 1
        For i = 1 To UBound(aAutoRoutSetting)
            .Rows = .Rows + 1
            .TextMatrix(lngRowPos, 0) = IIf(aAutoRoutSetting(i).Type = 1, "Ӱ�����", "����豸")
            .TextMatrix(lngRowPos, 1) = aAutoRoutSetting(i).strCondition
            .TextMatrix(lngRowPos, 2) = GetDeviceNameNum(aDevices, aAutoRoutSetting(i).strFTPDeviceNo, 0)
            lngRowPos = .Rows
        Next
    End With
End Sub
Private Function GetDeviceNameNum(aSource() As Variant, ByVal SeekString As String, iType As Integer) As String
    '��ȡ�豸�����ƻ��豸��
    'iType=0---����SeekStringΪ�豸�ţ������豸����
    'iType=1---����SeekStringΪ�豸���������豸�š�
    Dim i As Long
    For i = 0 To UBound(aSource, 2)
        If aSource(iType, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then GetDeviceNameNum = "": Exit Function
    GetDeviceNameNum = IIf(iType = 1, aSource(0, i), aSource(1, i))
End Function
Private Function GetComboxIndex(aSource() As Variant, ByVal SeekString As String) As Long
    Dim i As Long
    
    For i = 0 To UBound(aSource, 2)
        If aSource(0, i) = SeekString Then Exit For
    Next
    If i > UBound(aSource, 2) Then i = 0
    GetComboxIndex = i
End Function


Private Sub MSFAutoRout_Click()
    Dim iSelected As Integer
    With MSFAutoRout
        iSelected = .RowSel
        Me.optType(IIf(.TextMatrix(iSelected, 0) = "Ӱ�����", 1, 2)).value = True
        Me.cobCondition(IIf(.TextMatrix(iSelected, 0) = "Ӱ�����", 1, 2)).Text = .TextMatrix(iSelected, 1)
        Me.cobDestination = .TextMatrix(iSelected, 2)
    End With
End Sub

Private Sub MSFResult_Click()
    Dim iSelected As Integer
    With MSFResult
        iSelected = .RowSel
        Me.chkUseResult.value = IIf(.TextMatrix(iSelected, 3) = "True", 1, 0)
        Me.lblResult.Caption = .TextMatrix(iSelected, 0) & " " & .TextMatrix(iSelected, 1) & " : " & .TextMatrix(iSelected, 2)
        Me.txtResult(0).Text = .TextMatrix(iSelected, 4)
        Me.txtResult(1).Text = .TextMatrix(iSelected, 5)
        Me.chkResult.value = IIf(.TextMatrix(iSelected, 6) = True, 1, 0)
    End With
End Sub

Private Sub optImgMatch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optMatch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optType_Click(Index As Integer)
    Me.cobCondition(Index).Enabled = True
    Me.cobCondition(IIf(Index = 1, 2, 1)).Enabled = False
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case SSTab1.Tab
        Case 0
            chkStorage.SetFocus
        Case 1
            chkDWL.SetFocus
        Case 2
            chkQuery.SetFocus
        Case 3
            chkClearTempTB(0).SetFocus
    End Select
End Sub

Private Sub txtClearInterval_GotFocus()
    With txtClearInterval
        .SelStart = 0: .SelLength = .MaxLength
    End With
End Sub

Private Sub txtClearInterval_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtClearInterval_KeyPress(KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
End Sub

Private Sub txtDWLLocalAE_GotFocus()
    txtDWLLocalAE.SelStart = 0
    txtDWLLocalAE.SelLength = Len(txtDWLLocalAE.Text)
End Sub

Private Sub txtDWLLocalAE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    With Me.txtItem(Index)
        .SelStart = 0: .SelLength = .MaxLength
    End With
End Sub

Private Sub txtItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    
    If LenB(StrConv(Trim(txtItem(Index).Text), vbFromUnicode)) >= txtItem(Index).MaxLength Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 0, 2, 3, 6
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
        Case -1
            Call zlCommFun.OpenIme(False)
    End Select
End Sub

'�ж��Ƿ�Ϊ�༭��
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Sub txtQueryAE_GotFocus()
    txtQueryAE.SelStart = 0
    txtQueryAE.SelLength = Len(txtQueryAE.Text)
End Sub

Private Sub txtQueryAE_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub SetPrivs(strPrivs As String)
    '---------------------------------------------------------------
    '���ܣ�                                  ������Աʹ��Ȩ��
    '������
    '���أ�                                  ��
    '�ϼ���������̣�                        frmParaSet.Form_load
    '�¼���������̣�                        ��
    '���õ��ⲿ������                        mstrPrivs
    '�����ˣ�                                ���� 2005-8-25
    '---------------------------------------------------------------
    If InStr(strPrivs, "�洢�Զ�·��") = 0 Then
        cmdInsert.Enabled = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
    If InStr(strPrivs, "DICOM�����б����") = 0 Then
        chkDWL.Enabled = False
        frmWorkList.Enabled = False
    End If
    If InStr(strPrivs, "DICOM��������") = 0 Then
        chkQuery.Enabled = False
        frmQueryRetrieve.Enabled = False
    End If
End Sub

Private Sub subFillMsfResult()
    Dim lngRowPos As Long
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    
    With MSFResult
        .Clear
        .Rows = 1
        .Cols = 7
        .ColWidth(0) = 800
        .ColWidth(1) = 1800
        .ColWidth(2) = 1800
        .ColWidth(3) = 600
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 600
        .FixedCols = 3
        .TextMatrix(0, 0) = "���"
        .TextMatrix(0, 1) = "���ı���"
        .TextMatrix(0, 2) = "Ӣ�ı���"
        .TextMatrix(0, 3) = "��ѡ��"
        .TextMatrix(0, 4) = "����ֵ"
        .TextMatrix(0, 5) = "ǿ�ƽ��ֵ"
        .TextMatrix(0, 6) = "����"
        lngRowPos = 1
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        Set rsTmp = gcnAccess.Execute("select * from ǿ�ƽ��")
        While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(lngRowPos, 0) = "(" & rsTmp!��� & "," & rsTmp!Ԫ�غ� & ")"
            .TextMatrix(lngRowPos, 1) = Nvl(rsTmp!���ı���)
            .TextMatrix(lngRowPos, 2) = Nvl(rsTmp!Ӣ�ı���)
            .TextMatrix(lngRowPos, 3) = rsTmp!��ѡ��
            .TextMatrix(lngRowPos, 4) = Nvl(rsTmp!����ֵ)
            .TextMatrix(lngRowPos, 5) = Nvl(rsTmp!ǿ�ƽ��ֵ)
            .TextMatrix(lngRowPos, 6) = rsTmp!�Ƿ����
            rsTmp.MoveNext
            lngRowPos = .Rows
        Wend
    End With
End Sub

Private Sub subChangeMSFResult()
    Dim iSelect As Integer
    With Me.MSFResult
        iSelect = .RowSel
        If frmBuildResult.funVerifyResult(Me.txtResult(0).Text) <> 0 _
        Or frmBuildResult.funVerifyResult(Me.txtResult(1).Text) <> 0 Then
            Exit Sub
        End If
       .TextMatrix(iSelect, 3) = IIf(Me.chkUseResult.value = 0, "False", "True")
        .TextMatrix(iSelect, 4) = Me.txtResult(0).Text
        .TextMatrix(iSelect, 5) = Me.txtResult(1).Text
        .TextMatrix(iSelect, 6) = IIf(Me.chkResult.value = 0, "False", "True")
    End With
End Sub

Private Sub txtResult_Change(Index As Integer)
    mblnModifyMWLResult = True
    If Index = 1 Then
        Me.MSFResult.TextMatrix(Me.MSFResult.RowSel, 5) = Me.txtResult(1).Text
    End If
End Sub

Private Sub txtResult_Click(Index As Integer)
    If Index = 0 Then
        cmdBuildResult_Click (0)
    End If
End Sub

Private Sub txtResult_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        KeyAscii = 0
        cmdBuildResult_Click (0)
    End If
End Sub
