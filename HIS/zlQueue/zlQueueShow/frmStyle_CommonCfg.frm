VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmStyle_CommonCfg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ʽ����"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   Icon            =   "frmStyle_CommonCfg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkCallTarget 
      Caption         =   "��ʾ����Ŀ�ĵ�"
      Height          =   255
      Left            =   2160
      TabIndex        =   41
      Top             =   5950
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd��ʾ�豸���� 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   4040
      TabIndex        =   38
      Top             =   5880
      Width           =   1200
   End
   Begin VB.CheckBox chkScrollDisplay 
      Caption         =   "������ʾ�ѹ�������"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5950
      Width           =   1935
   End
   Begin TabDlg.SSTab sstFormSetup 
      Height          =   5640
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9948
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Һ����λ��"
      TabPicture(0)   =   "frmStyle_CommonCfg.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��ʾ��������"
      TabPicture(1)   =   "frmStyle_CommonCfg.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfQueueSetup"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "����Ƥ������"
      TabPicture(2)   =   "frmStyle_CommonCfg.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(1)=   "cboStyleType"
      Tab(2).Control(2)=   "fraRemarkInfo"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "ҽ��\������Ϣ"
      TabPicture(3)   =   "frmStyle_CommonCfg.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "fraDeptInfo"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame3 
         Height          =   700
         Left            =   -74880
         TabIndex        =   42
         Top             =   360
         Width           =   7140
         Begin VB.ComboBox cboCurRoom 
            Height          =   300
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   240
            Width           =   2205
         End
         Begin VB.ComboBox cboCurDept 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��ǰ����/ִ�м�"
            Height          =   180
            Left            =   3360
            TabIndex        =   46
            Top             =   285
            Width           =   720
         End
         Begin VB.Label lblCurDept 
            AutoSize        =   -1  'True
            Caption         =   "��ǰ����"
            Height          =   180
            Left            =   120
            TabIndex        =   45
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   -74880
         TabIndex        =   31
         Top             =   3600
         Width           =   7140
         Begin VB.CheckBox chkConvertQueueName 
            Caption         =   "ת�����ϰ��������"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "���������ƵĴ洢��ʽת��Ϊ�ϰ汾�ĸ�ʽ"
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkShowDeptName 
            Caption         =   "��ʾ������"
            Height          =   255
            Left            =   3950
            TabIndex        =   37
            ToolTipText     =   "����ʽ��������ʾ���ұ���ʱ���Ƿ���Ҫ��ʾ��Ӧ�Ŀ�����"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtQueueRows 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   6400
            TabIndex        =   33
            Top             =   220
            Width           =   375
         End
         Begin VB.CheckBox chkFontAutoSizeToList 
            Caption         =   "�б���������Ӧ"
            Height          =   255
            Left            =   2240
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblQueue 
            AutoSize        =   -1  'True
            Caption         =   "�Ŷ��б���ʾ     ��"
            Height          =   180
            Left            =   5300
            TabIndex        =   34
            Top             =   270
            Width           =   1710
         End
      End
      Begin VB.Frame fraRemarkInfo 
         Caption         =   "�׶��ı�"
         Height          =   800
         Left            =   -74880
         TabIndex        =   29
         Top             =   4720
         Width           =   7140
         Begin VB.TextBox txtRemarkInfo 
            Appearance      =   0  'Flat
            Height          =   460
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   240
            Width           =   6945
         End
      End
      Begin VB.ComboBox cboStyleType 
         Height          =   300
         Left            =   -73990
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   400
         Width           =   2535
      End
      Begin VB.Frame Frame9 
         Caption         =   "Ƥ������"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   27
         Top             =   440
         Width           =   7140
         Begin VB.PictureBox picStyleView 
            BackColor       =   &H80000008&
            Height          =   3735
            Left            =   120
            ScaleHeight     =   3675
            ScaleWidth      =   6840
            TabIndex        =   40
            Top             =   360
            Width           =   6900
            Begin VB.Image imgStyleView 
               Height          =   3780
               Left            =   0
               Picture         =   "frmStyle_CommonCfg.frx":007C
               Stretch         =   -1  'True
               Top             =   0
               Width           =   6900
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "��ʾ���ݹ�����������"
         Height          =   1140
         Left            =   -74880
         TabIndex        =   24
         Top             =   4320
         Width           =   7140
         Begin VB.TextBox txtFilter 
            Appearance      =   0  'Flat
            Height          =   705
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   300
            Width           =   6945
         End
      End
      Begin VB.Frame fraDeptInfo 
         Caption         =   "���Ҽ��"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   22
         Top             =   3975
         Width           =   7140
         Begin VB.TextBox txtDeptInfo 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   300
            Width           =   6825
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "ҽ�������Ϣ"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   18
         Top             =   1200
         Width           =   7140
         Begin VB.TextBox txtIntroduction 
            Appearance      =   0  'Flat
            Height          =   2025
            Index           =   0
            Left            =   4080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   500
            Visible         =   0   'False
            Width           =   2865
         End
         Begin VB.CommandButton cmdSetDocPhoto 
            Caption         =   "��Ƭ����(&S)"
            Height          =   350
            Left            =   2360
            TabIndex        =   20
            Top             =   2160
            Width           =   1605
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   2360
            ScaleHeight     =   1785
            ScaleWidth      =   1575
            TabIndex        =   19
            Top             =   285
            Width           =   1605
            Begin VB.Image imgDoctorPhoto 
               Height          =   1815
               Index           =   0
               Left            =   0
               Picture         =   "frmStyle_CommonCfg.frx":30D77
               Stretch         =   -1  'True
               Top             =   0
               Visible         =   0   'False
               Width           =   1605
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfDoctorInfo 
            Height          =   2235
            Left            =   120
            TabIndex        =   21
            Top             =   285
            Width           =   2115
            _cx             =   3731
            _cy             =   3942
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483638
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   2
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   272
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
         Begin VB.Label lblDoctorIntro 
            AutoSize        =   -1  'True
            Caption         =   "ҽ����飺"
            Height          =   180
            Left            =   4080
            TabIndex        =   36
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5100
         Left            =   120
         TabIndex        =   3
         Top             =   400
         Width           =   7140
         Begin VB.OptionButton optFullScreen 
            Caption         =   "ȫ��"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   760
            Width           =   855
         End
         Begin VB.OptionButton optCustom 
            Caption         =   "�Զ���"
            Height          =   255
            Left            =   2160
            TabIndex        =   14
            Top             =   760
            Width           =   900
         End
         Begin VB.Frame frmCustom 
            Caption         =   "�Զ���λ��(�ֱ���Ϊ��λ)"
            Height          =   1245
            Left            =   120
            TabIndex        =   5
            Top             =   1200
            Width           =   6900
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   1
               Left            =   840
               TabIndex        =   9
               Top             =   360
               Width           =   2535
            End
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   2
               Left            =   840
               TabIndex        =   8
               Top             =   830
               Width           =   2535
            End
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   3
               Left            =   4440
               TabIndex        =   7
               Top             =   360
               Width           =   2295
            End
            Begin VB.TextBox txtRect 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   4
               Left            =   4440
               TabIndex        =   6
               Top             =   840
               Width           =   2295
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   1
               Left            =   240
               TabIndex        =   13
               Top             =   405
               Width           =   180
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   2
               Left            =   240
               TabIndex        =   12
               Top             =   888
               Width           =   180
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "���"
               Height          =   180
               Index           =   3
               Left            =   3840
               TabIndex        =   11
               Top             =   405
               Width           =   360
            End
            Begin VB.Label lblRect 
               AutoSize        =   -1  'True
               Caption         =   "�߶�"
               Height          =   180
               Index           =   4
               Left            =   3840
               TabIndex        =   10
               Top             =   885
               Width           =   360
            End
         End
         Begin VB.ComboBox cboLCDNum 
            Height          =   300
            ItemData        =   "frmStyle_CommonCfg.frx":32128
            Left            =   1200
            List            =   "frmStyle_CommonCfg.frx":3212F
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   250
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��ʾ�����"
            Height          =   180
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   900
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfQueueSetup 
         Height          =   3105
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   7140
         _cx             =   12594
         _cy             =   5477
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483638
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6540
      TabIndex        =   1
      Top             =   5895
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5400
      TabIndex        =   0
      Top             =   5895
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgDoctorPhoto 
      Left            =   4800
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmStyle_CommonCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��������ô����������ô���frmMain����
'�ڸ�ͨ����ʽ���ô����£����õ����ݰ������£�
'1.������ʾλ�ã�������ʾ����ţ���ʾ����(left,top,right,bottom),��Ϊȫ����ʾʱ������Ҫ����ʾ�����������
'2.��ǰ�Ŷӽк���ʾ���ڿ��Ҽ���������
'3.��ʾ�������ã����ҵ������Ϊ0��ʾ���������Ŷӽкŵ���ʾҵ��ҵ�����Ͷ���ɲο�����glngBusinessType��
'                ��ͬ��ҵ�����ͻ�ȡ�������Ƶķ�ʽҲ��������
'                   pacsҵ��������ƹ���Ϊ���������� + "-" + ִ�м� �硰�����-DR1�ҡ�
'4.�Զ����Ŷӽк���ʾ���ݹ�����������
'5.�����ʾ����

Private mlngWindowNo As Long
Private mlngStyleType As Long
Private mobjParent As Object
Private mrsRecord As ADODB.Recordset
Private mstrDoctorPhoto() As String         'ҽ����Ƭ��ʮ�����ƴ�
Private mstrStyleTylePath As String
Private mintSelectedQueueNum As Integer     '��ѡ��Ķ�����
Private mlngQueueListMaxRows As Long        '�б������ʾ���ݵ�������
Private mblnShowListHeader As Boolean       '��ʾ�б����
Private mrsDept As ADODB.Recordset
Private mstrCurDiagnoseRoom As String       '����ִ�м�
Private mstrPreSelectDept As String         '��һ��ѡ��Ŀ�����

Public Function OpenShowConfig(ByVal lngWindowNo As Long, ByVal lngStyleType As TShowStyle, objOwner As Object) As Boolean
'��ʾ���ô���
    OpenShowConfig = False

    mlngWindowNo = lngWindowNo
    mlngStyleType = lngStyleType
    mintSelectedQueueNum = 0
    Set mobjParent = objOwner
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical, TBusinessType.btPacs, TBusinessType.btPeis
            Call InitFace
            
            Call InitQueueSetup
            
            Call InitDoctorInfo
            
            Call InitLocalPars
            
            Call Me.Show(vbModal, objOwner)
    End Select
    
    OpenShowConfig = True
End Function

Private Sub InitFace()
'������ʾ��ʽ��ʼ������������
    Dim i As Integer
    
    If mlngStyleType = TShowStyle.ssSingleMan Then         '������
        cmd��ʾ�豸����.Visible = False
        chkScrollDisplay.Visible = False
        lblQueue.Visible = False
        txtQueueRows.Visible = False
        fraRemarkInfo.Enabled = False
        
        lblDoctorIntro.Enabled = False
        For i = 0 To txtIntroduction.Count - 1
            txtIntroduction(i).Enabled = False
        Next
        fraDeptInfo.Enabled = False
    
    ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then   '������
        cmd��ʾ�豸����.Visible = False
        chkCallTarget.Visible = True
    
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then   '�����
        lblQueue.Visible = False
        txtQueueRows.Visible = False
        chkShowDeptName.Visible = False
        cmd��ʾ�豸����.Visible = False
        sstFormSetup.TabVisible(3) = False
        
    ElseIf mlngStyleType = TShowStyle.ssOld Then   '�ϰ汾
        chkFontAutoSizeToList.Visible = False
        lblQueue.Visible = False
        txtQueueRows.Visible = False
        chkShowDeptName.Visible = False
        chkScrollDisplay.Visible = False
        
        sstFormSetup.TabVisible(0) = False
        sstFormSetup.TabVisible(2) = False
        sstFormSetup.TabVisible(3) = False
    End If
End Sub

Private Sub InitDoctorInfo()
'��ʼ��ҽ�������Ϣ����
    Dim i As Integer
    Dim strDoctorInfo As String
    Dim strDoctorPhoto As String
    Dim strIntroduction As String
    Dim strWorkingTime As String
    
On Error GoTo ErrorHand

    If mlngStyleType <> TShowStyle.ssSingleMan And mlngStyleType <> TShowStyle.ssSingleQueue Then Exit Sub
    
    With vsfDoctorInfo
        .Cols = 3
        .Rows = 20
        ReDim mstrDoctorPhoto(.Rows - 2) As String
        
        For i = 1 To .Rows - 1
            Load imgDoctorPhoto(i)
            Load txtIntroduction(i)
            imgDoctorPhoto(i).Visible = True
            txtIntroduction(i).Visible = True
        Next
        
        .TextMatrix(0, 0) = "ҽ������"
        .TextMatrix(0, 1) = "ֵ��ʱ��"
        
        .Editable = flexEDKbdMouse
        
        .ColHidden(2) = True
        .ColComboList(1) = " |������|����һ|���ڶ�|������|������|������|������"
        
        strDoctorInfo = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ϣ")    '������ְλ
        strDoctorPhoto = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ƭ")    '
        strWorkingTime = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ֵ��ʱ��")   '
        strIntroduction = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ�����")   '
        
        For i = 1 To .Rows - 1
            If strDoctorInfo <> "" Then
                vsfDoctorInfo.TextMatrix(i, 0) = Split(Split(Mid(strDoctorInfo, 2), "|")(i - 1), "-")(1)
                vsfDoctorInfo.TextMatrix(i, 2) = Split(Split(Mid(strDoctorInfo, 2), "|")(i - 1), "-")(0)
            End If

            If strDoctorPhoto <> "" Then
                Call LoadPictureInfo(imgDoctorPhoto(i), Split(Mid(strDoctorPhoto, 2), "|")(i - 1))
                mstrDoctorPhoto(i - 1) = Split(Mid(strDoctorPhoto, 2), "|")(i - 1)
            End If
            If strWorkingTime <> "" Then vsfDoctorInfo.TextMatrix(i, 1) = Split(Mid(strWorkingTime, 2), "|")(i - 1)
            If strIntroduction <> "" Then txtIntroduction(i).Text = Split(Mid(strIntroduction, 2), "|")(i - 1)
            
            vsfDoctorInfo.AutoSize 0, vsfDoctorInfo.Cols - 1
        Next
        
        If .Rows > 1 Then .RowSel = 1
        .AutoSize 0, .Cols - 1
        
        .Editable = flexEDNone
    End With

Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub LoadDoctorInfo(ByVal lngCurDeptID As Long)
'������ʱ��������ʵ�������ã����ض�Ӧ���ҵ�ҽ�����ֵ���Ϣ
    Dim i As Integer
    Dim strSql As String
    Dim strDoctorNames As String
    Dim str�������� As String
    
    vsfDoctorInfo.ColComboList(0) = ""
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            str�������� = "�ٴ�"
        Case TBusinessType.btPacs
            str�������� = "���"
        Case TBusinessType.btPeis
            str�������� = "���"
        'case
        '
        '
    End Select
    
    strSql = "select A.ID as ����ID,A.����,C.����,C.ID as ��ԱID from ���ű� A,������Ա B,��Ա�� C,��������˵�� D " & _
             "Where A.ID=[1] And A.ID = B.����ID And B.��ԱID = C.ID And D.����ID = A.ID " & _
             "And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
             "And D.�������� IN('" & str�������� & "') Order by A.����"
    
    Set mrsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngCurDeptID)
    
    If mrsRecord.RecordCount <= 0 Then Exit Sub
    
    Do While Not mrsRecord.EOF
        strDoctorNames = strDoctorNames & "|" & Nvl(mrsRecord!����)
        mrsRecord.MoveNext
    Loop
    
   vsfDoctorInfo.Editable = flexEDKbdMouse
   vsfDoctorInfo.ColComboList(0) = " " & strDoctorNames
End Sub

Private Sub InitLocalPars()
'��ʼ����ʽ���ò���
    Dim i As Integer, j As Integer
    Dim lngLCDNum As Long
    Dim lngCurLCDNo As Long
    Dim strCallingQueues As String
    Dim strLCDLocation As String
    Dim objForder As Folder
    Dim objFile As File

On Error GoTo ErrorHand
    '��������
    txtFilter.Text = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��������", "")
    chkConvertQueueName.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ת����������", 0))
    
    '�Ŷ��б���ʾ�Ķ�����
    strCallingQueues = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����")
    
    For i = 1 To vsfQueueSetup.Rows - 1
        For j = 0 To vsfQueueSetup.Cols - 1
            If InStr(strCallingQueues, vsfQueueSetup.TextMatrix(0, j) & "|" & vsfQueueSetup.TextMatrix(vsfQueueSetup.Rows - 1, j) & ":" & vsfQueueSetup.TextMatrix(i, j)) > 0 Then
                If vsfQueueSetup.TextMatrix(i, j) <> "" Then
                    vsfQueueSetup.Cell(flexcpChecked, i, j) = 1
                    
                    mintSelectedQueueNum = mintSelectedQueueNum + 1
                End If
            End If
        Next
    Next
    
    If mlngStyleType = TShowStyle.ssOld Then Exit Sub
    
    '�������п���
    mstrCurDiagnoseRoom = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "����ִ�м�", "")
    Call InitRoom
    
    If mlngStyleType = TShowStyle.ssMultiQueue Then
        vsfQueueSetup.ToolTipText = "����ѡ����" & mintSelectedQueueNum & "�����У�" & vbCrLf & "�������ѡ��" & mlngQueueListMaxRows & "�����У�"
    End If

    '��ʾģʽ,0-ȫ����1-�Զ���
    If Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾģʽ", 0)) = 0 Then
        optFullScreen.value = True
    Else
        optCustom.value = True
    End If
    
    '�Զ�����ʾλ��
    If optCustom.value Then
        strLCDLocation = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Զ���λ��")
        
        For i = 0 To UBound(Split(strLCDLocation, "|"))
            txtRect(i + 1).Text = Mid(Split(strLCDLocation, "|")(i), 3)
        Next
    End If
    
    '������ʾ�����
    Call InitMonitor
    
    lngLCDNum = UBound(gmonitors) - 1
    lngCurLCDNo = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ�����")) - 1
    
    cboLCDNum.Clear
    
    For i = 1 To lngLCDNum
        cboLCDNum.AddItem i
        If i - 1 = lngCurLCDNo Then cboLCDNum.ListIndex = i - 1
    Next
    
    If cboLCDNum.ListIndex < 0 Then cboLCDNum.ListIndex = 0
    
    '���ض�����ʽ
    cboStyleType.Clear
    
    If mlngStyleType = TShowStyle.ssSingleMan Then
        If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
            Set objForder = gobjFile.GetFolder(App.Path & "\Skin\��������ʽ")
        Else
            Set objForder = gobjFile.GetFolder(App.Path & "\zlQueueShow\Skin\��������ʽ")
        End If
    ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
            Set objForder = gobjFile.GetFolder(App.Path & "\Skin\��������ʽ")
        Else
            Set objForder = gobjFile.GetFolder(App.Path & "\zlQueueShow\Skin\��������ʽ")
        End If
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\�������ʽ") Then
            Set objForder = gobjFile.GetFolder(App.Path & "\Skin\�������ʽ")
        Else
            Set objForder = gobjFile.GetFolder(App.Path & "\zlQueueShow\Skin\�������ʽ")
        End If
    End If
    
    For Each objFile In objForder.Files
        If Mid(objFile.Name, Len(objFile.Name) - 2) = "jpg" Then
            cboStyleType.AddItem Mid(objFile.Name, 1, Len(objFile.Name) - 4)
            
            If objFile.Path = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ") & ".jpg" Then
                cboStyleType.Text = Mid(objFile.Name, 1, Len(objFile.Name) - 4)
            End If
        End If
    Next
    
    If cboStyleType.ListCount > 0 And cboStyleType.ListIndex < 0 Then cboStyleType.ListIndex = 0
    If cboStyleType.ListIndex >= 0 Then
        If mlngStyleType = TShowStyle.ssSingleMan Then
            '''''''''''''
            
        ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then
            If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
                Call SetIniFile(App.Path & "\Skin\��������ʽ\" & cboStyleType.Text & ".ini")
            Else
                Call SetIniFile(App.Path & "\zlQueueShow\Skin\��������ʽ\" & cboStyleType.Text & ".ini")
            End If
            
            mblnShowListHeader = CBool(ReadValue("�Ŷ��б�����", "�Ƿ���ʾ�б����"))
            
            If mblnShowListHeader Then
                mlngQueueListMaxRows = Val(ReadValue("�Ŷ��б�����", "������")) - 1
            Else
                mlngQueueListMaxRows = Val(ReadValue("�Ŷ��б�����", "������"))
            End If
        Else
            If gobjFile.FolderExists(App.Path & "\Skin\�������ʽ") Then
                Call SetIniFile(App.Path & "\Skin\�������ʽ\" & cboStyleType.Text & ".ini")
            Else
                Call SetIniFile(App.Path & "\zlQueueShow\Skin\�������ʽ\" & cboStyleType.Text & ".ini")
            End If
            
            mlngQueueListMaxRows = Val(ReadValue("׼�������б�����", "������")) - 1
        End If
    End If
    
    chkFontAutoSizeToList.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�б���������Ӧ", 1))
    chkShowDeptName.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "���ұ����Ƿ���ʾ������", 0))
    
    If mlngStyleType = TShowStyle.ssSingleMan Then Exit Sub
    
    '�Ƿ������ʾ������Ϣ
    chkScrollDisplay.value = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "������ʾ", 1)
    
    If mlngStyleType = TShowStyle.ssSingleQueue Then
        chkCallTarget.value = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����Ŀ�ĵ�", 0))
        txtQueueRows.Text = Val(GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Ŷ��б���ʾ��", mlngQueueListMaxRows))
        txtQueueRows.ToolTipText = "�Ŷ��б�������ʾ�����������ʾ" & mlngQueueListMaxRows & "��"
        txtDeptInfo.Text = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "���Ҽ��")
    End If
    
    '�׶��ı�
    txtRemarkInfo.Text = GetSetting("ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�׶��ı�", "��δ�е��ŵĻ������ĵȴ�!")
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub InitRoom()
'���ܣ���������ִ�п���
    Dim i As Integer
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
    
    cboCurDept.Clear
    
    If mrsDept.RecordCount <= 0 Then Exit Sub
    mrsDept.MoveFirst
    
    For i = 0 To mrsDept.RecordCount - 1
        cboCurDept.AddItem Nvl(mrsDept!����)
        cboCurDept.ItemData(i) = Nvl(mrsDept!ID)
        
        If mstrCurDiagnoseRoom <> "" Then
            If Nvl(mrsDept!����) = Split(mstrCurDiagnoseRoom, "-")(0) Then
                mstrPreSelectDept = Nvl(mrsDept!����)
                cboCurDept.ListIndex = i
            End If
        End If
        
        mrsDept.MoveNext
    Next
    
    If cboCurDept.ListCount > 0 And cboCurDept.ListIndex < 0 Then
        strSql = "select d.���� from �ϻ���Ա�� A,��Ա�� B,������Ա C,���ű� D " & _
                 "where A.��ԱID=B.ID And b.id=c.��Աid and c.����id=d.id and c.ȱʡ=1 and A.�û���=[1]"
        
        Set rsRecord = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", UCase(gstrUserName))
        
        If rsRecord.RecordCount > 0 Then
            For i = 0 To cboCurDept.ListCount - 1
                If cboCurDept.List(i) = Nvl(rsRecord!����) Then
                    mstrPreSelectDept = Nvl(rsRecord!����)
                    cboCurDept.ListIndex = i
                End If
            Next
        End If
    End If
    
    If cboCurDept.ListCount > 0 And cboCurDept.ListIndex < 0 Then cboCurDept.ListIndex = 0
End Sub

Private Sub InitQueueSetup()
'���ض�����Ϣ
    Dim i As Integer, j As Integer
    Dim intQueueNum As Integer  'ÿ�����Ҷ�Ӧ�Ķ�����
    Dim str��Դ As String, str�������� As String
    Dim strSql As String
    Dim rsRoom As ADODB.Recordset
    Dim lngRoomMaxNum As Long

On Error GoTo ErrorHand
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            str�������� = "�ٴ�"
            str��Դ = ",1,3,"
        Case TBusinessType.btPacs
            str�������� = "���"
            str��Դ = ",1,2,3,"
        Case TBusinessType.btPeis
            str�������� = "���"
            str��Դ = ",1,2,3,"
        'case
        '
        '
    End Select
    
    strSql = "Select Distinct A.ID,A.����,A.���� From ���ű� A,��������˵�� B " & _
             "Where B.����ID = A.ID  And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
             "And B.�������� IN('" & str�������� & "') And instr('" & str��Դ & "',','||B.�������||',')> 0 Order by A.����"
    
    Set mrsDept = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ʾ��������")
    
    If mrsDept.RecordCount <= 0 Then Exit Sub
    
    '���ݲ�ͬҵ����Ҷ�Ӧ��������
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "select ����id,����," & _
                           "case " & _
                                "when instr(����, 'һ') > 0 then replace(����, 'һ', '1') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '2') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '3') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '4') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '5') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '6') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '7') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '8') " & _
                                "else replace(����, '��', '9') " & _
                           "end As ord " & _
                    "from ( Select Distinct P.����ID, 'ִ�м�-'||R.���� as ���� From �������� R, �ҺŰ������� S, �ҺŰ��� P " & _
                    "Where R.���� = S.�������� And S.�ű�id = P.ID ) a order by ����id,ord "
            
        Case TBusinessType.btPacs
            If gstrCompareVersion < "010.034.000" Then
                strSql = "select ����id,����," & _
                              "case " & _
                                  "when instr(����, 'һ') > 0 then replace(����, 'һ', '1') " & _
                                  "when instr(����, '��') > 0 then replace(����, '��', '2') " & _
                                  "when instr(����, '��') > 0 then replace(����, '��', '3') " & _
                                  "when instr(����, '��') > 0 then replace(����, '��', '4') " & _
                                  "when instr(����, '��') > 0 then replace(����, '��', '5') " & _
                                  "when instr(����, '��') > 0 then replace(����, '��', '6') " & _
                                  "when instr(����, '��') > 0 then replace(����, '��', '7') " & _
                                  "when instr(����, '��') > 0 then replace(����, '��', '8') " & _
                                  "else replace(����, '��', '9') end As ord " & _
                              "from (select A.����id, 'ִ�м�-'||ִ�м� as ���� from ҽ��ִ�з��� A,���ű� B,��������˵�� C " & _
                                  "Where A.����ID=B.id And B.ID=C.����ID And C.��������='" & str�������� & "' And ����id not in " & _
                                  "(select ����id from Ӱ�����̲��� where ������='�Ŷӽкŷ�ʽ' and ����ֵ=1 ) " & _
                                  "union select ����id,'ִ�м�-���Ҷ���' as ���� from Ӱ�����̲��� where ������='�Ŷӽкŷ�ʽ' and ����ֵ=1) " & _
                              "order by ����id,ord"
            Else
                strSql = "select ����id,����,ord from " & _
                         "(select ����id,����, " & _
                            "case " & _
                                "when instr(����, 'һ') > 0 then replace(����, 'һ', '1') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '2') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '3') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '4') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '5') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '6') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '7') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '8') " & _
                                "else replace(����, '��', '9') " & _
                            "end As ord " & _
                        "from (select ����id, 'ִ�м�-'||ִ�м� as ���� from ҽ��ִ�з��� A,���ű� B,��������˵�� C " & _
                              "Where A.����ID=B.id And B.ID=C.����ID And C.��������='" & str�������� & "') " & _
                        "union select b.id,'ִ�з���-'||a.���� ����,'��' as ord from Ӱ��ִ�з��� a,���ű� b where a.����id=b.id) " & _
                        "order by ����id,ord"

            End If
            
        Case TBusinessType.btPeis
            strSql = "select ����id,����," & _
                           "case " & _
                                "when instr(����, 'һ') > 0 then replace(����, 'һ', '1') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '2') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '3') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '4') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '5') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '6') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '7') " & _
                                "when instr(����, '��') > 0 then replace(����, '��', '8') " & _
                                "else replace(����, '��', '9') " & _
                          "end As ord " & _
                    "from (select ����id, 'ִ�м�-'||ִ�м� as ���� from ҽ��ִ�з��� A,���ű� B,��������˵�� C  " & _
                     "Where A.����ID=B.id And B.ID=C.����ID And C.��������='" & str�������� & "')" & _
                    "order by ����id,ord"""

        'case
        '.
        '.
    End Select
    
    Set rsRoom = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ʾ��������")
    
    With vsfQueueSetup
        .Editable = flexEDKbdMouse
        .Cols = mrsDept.RecordCount
        
        If rsRoom.RecordCount > 0 Then
            .Rows = rsRoom.RecordCount + 3
        Else
            .Rows = 2
        End If
        
        If glngBusinessType = TBusinessType.btClinical And (mlngStyleType = TShowStyle.ssSingleQueue Or mlngStyleType = TShowStyle.ssMultiQueue) Then
            .Rows = .Rows + 1
        End If
        
        mrsDept.MoveFirst

        For i = 0 To .Cols - 1
            '������ͷ������������
            .TextMatrix(0, i) = Nvl(mrsDept!����)
            
            intQueueNum = 0
            
            If rsRoom.RecordCount > 0 Then
                '���ض�Ӧ����ִ�м�
                rsRoom.MoveFirst
                For j = 0 To rsRoom.RecordCount - 1
                    If Nvl(mrsDept!ID) = Nvl(rsRoom!����id) Then
                        intQueueNum = intQueueNum + 1
                        
                        .TextMatrix(intQueueNum, i) = Split(Nvl(rsRoom!����), "-")(1)
                        
                        '�Է��������ɫ����
                        Select Case Split(Nvl(rsRoom!����), "-")(0)
                            Case "ִ�м�"
                                '''''
                            Case "ִ�з���"
                                .Cell(flexcpForeColor, intQueueNum, i) = vbRed
                        End Select
                        
                        .Cell(flexcpChecked, intQueueNum, i) = 2
                    End If
                    
                    rsRoom.MoveNext
                Next
            End If

            If glngBusinessType = TBusinessType.btPacs And gstrCompareVersion >= "010.034.000" Then 'PACSҵ������"δ�������"
                intQueueNum = intQueueNum + 1
                .TextMatrix(intQueueNum, i) = "δ�������"
                .Cell(flexcpChecked, intQueueNum, i) = 2
                .Cell(flexcpForeColor, intQueueNum, i) = vbRed
            End If
            
            '�ٴ�ҵ�񣬵��ӺͶ����ʱ���ӿ��Ҷ��У�������ʾ�������ҵ��Ŷ���Ϣ
            If glngBusinessType = TBusinessType.btClinical And (mlngStyleType = TShowStyle.ssSingleQueue Or mlngStyleType = TShowStyle.ssMultiQueue) Then
                intQueueNum = intQueueNum + 1
                .TextMatrix(intQueueNum, i) = "���Ҷ���"
                .Cell(flexcpChecked, intQueueNum, i) = 2
                .Cell(flexcpForeColor, intQueueNum, i) = vbRed
            End If
            
            If intQueueNum = 0 Then .ColHidden(i) = True     'û��ִ�м�ʱ����ʽΪ������ʱ������ʾ����
            If intQueueNum > lngRoomMaxNum Then lngRoomMaxNum = intQueueNum
            
            '�洢��Ӧ�п��ҵ�ID�ͱ��룬���в���ʾ
            .TextMatrix(.Rows - 1, i) = Nvl(mrsDept!ID) & "_" & Nvl(mrsDept!����)
        
            mrsDept.MoveNext
        Next
        
        For i = lngRoomMaxNum + 1 To .Rows - 1
            .RowHidden(i) = True
        Next
        
        '�Զ��п�
        .AutoSize 0, .Cols - 1
        
        '���һ���Զ�������б�
        .ExtendLastCol = True
    End With
Exit Sub
ErrorHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Private Sub cboCurDept_Click()
'���ص�ǰ�����µ�ִ�з���/����
    Dim strSql As String
    Dim rsRoom As ADODB.Recordset
On Error GoTo ErrorHand
    If mstrPreSelectDept <> "" Then
        If mstrPreSelectDept <> cboCurDept.Text Then
            If MsgBox("�˲������������ҽ���������Ϣ���ã��Ƿ������", vbYesNo + vbDefaultButton2) = vbNo Then
                cboCurDept.Text = mstrPreSelectDept
                Exit Sub
            Else
                '���ҽ�������Ϣ����
                Call ClearDoctorInfo
            End If
        End If
    End If
    
    mstrPreSelectDept = cboCurDept.Text
    
    '����ҽ�������Ϣ
    Call LoadDoctorInfo(Val(cboCurDept.ItemData(cboCurDept.ListIndex)))
    
    '��������
    cboCurRoom.Clear
    
    Select Case glngBusinessType
        Case TBusinessType.btClinical
            strSql = "Select Distinct P.����ID, 'ִ�м�-'||R.���� as ���� From �������� R, �ҺŰ������� S, �ҺŰ��� P " & _
                     "Where R.���� = S.�������� And S.�ű�id = P.ID and p.����id=[1]  And R.ȱʡ��־ <> 1 "
                     
        Case TBusinessType.btPacs
            strSql = "select ����id, 'ִ�м�-'||ִ�м� as ���� from ҽ��ִ�з��� A,���ű� B,��������˵�� C " & _
                     "Where A.����ID=B.id And B.ID=C.����ID and a.����id=[1] And C.��������='���'"
                     
        Case TBusinessType.btPeis
            strSql = "select ����id, 'ִ�м�-'||ִ�м� as ���� from ҽ��ִ�з��� A,���ű� B,��������˵�� C  " & _
                   "Where A.����ID=B.id And B.ID=C.����ID and a.����id=[1]  And C.��������='���'"
                   
    End Select
    
    Set rsRoom = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ѯ����ִ�м�", Val(cboCurDept.ItemData(cboCurDept.ListIndex)))
    
    If rsRoom.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsRoom.EOF
        cboCurRoom.AddItem Split(Nvl(rsRoom!����), "-")(1)
        
        If mstrCurDiagnoseRoom <> "" Then
            If Split(Nvl(rsRoom!����), "-")(1) = Split(mstrCurDiagnoseRoom, "-")(1) Then cboCurRoom.Text = Split(Nvl(rsRoom!����), "-")(1)
        End If
        
        rsRoom.MoveNext
    Loop
    
    If cboCurRoom.ListCount > 0 And cboCurRoom.ListIndex < 0 Then cboCurRoom.ListIndex = 0
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cboStyleType_Click()
    If mlngStyleType = TShowStyle.ssSingleMan Then
        If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
            mstrStyleTylePath = App.Path & "\Skin\��������ʽ\" & cboStyleType.Text & ".jpg"
        Else
            mstrStyleTylePath = App.Path & "\zlQueueShow\Skin\��������ʽ\" & cboStyleType.Text & ".jpg"
        End If
    ElseIf mlngStyleType = TShowStyle.ssSingleQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
            mstrStyleTylePath = App.Path & "\Skin\��������ʽ\" & cboStyleType.Text & ".jpg"
        Else
            mstrStyleTylePath = App.Path & "\zlQueueShow\Skin\��������ʽ\" & cboStyleType.Text & ".jpg"
        End If
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\�������ʽ") Then
            mstrStyleTylePath = App.Path & "\Skin\�������ʽ\" & cboStyleType.Text & ".jpg"
        Else
            mstrStyleTylePath = App.Path & "\zlQueueShow\Skin\�������ʽ\" & cboStyleType.Text & ".jpg"
        End If
    End If
    
    If gobjFile.FileExists(mstrStyleTylePath) Then
        imgStyleView.Picture = LoadPicture(mstrStyleTylePath)
    
        Call ResizeImg(imgStyleView, 0, 0, picStyleView.Width, picStyleView.Height)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
'��������
On Error GoTo ErrorHand
    Call SaveLocalPars
    
    Unload Me
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub SaveLocalPars()
'��������
    Dim i As Integer, j As Integer
    Dim strCallingQueues As String
    Dim strDoctorInfo As String     '�����ʽ��"ҽ��1��������ְλ|ҽ��2��������ְλ|��������"
    Dim strDoctorPhoto As String    '�����ʽ��"ҽ��1����Ƭ|ҽ��2����Ƭ|��������"
    Dim strIntroduction As String   '�����ʽ��"ҽ��1�ļ��|ҽ��2�ļ��|��������"
    Dim strWorkingTime As String    '�����ʽ��"ҽ��1��ֵ��ʱ��|ҽ��2��ֵ��ʱ��|��������"
    
    For j = 0 To vsfQueueSetup.Cols - 1
        For i = 1 To vsfQueueSetup.Rows - 1
            If vsfQueueSetup.Cell(flexcpChecked, i, j) = 1 Then
                strCallingQueues = strCallingQueues & "," & vsfQueueSetup.TextMatrix(0, j) & "|" & vsfQueueSetup.TextMatrix(vsfQueueSetup.Rows - 1, j) & ":" & vsfQueueSetup.TextMatrix(i, j)
            End If
        Next
    Next
    '�������õĶ��У���ʽ��"��������|����ID:ִ�м�"���磺"�����|64:CTһ��"
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����", Mid(strCallingQueues, 2)
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��������", txtFilter.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ת����������", chkConvertQueueName.value
    
    If mlngStyleType = TShowStyle.ssOld Then Exit Sub
    
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾģʽ", IIf(optFullScreen.value = True, 0, 1)
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Զ���λ��", "��:" & Val(txtRect(1).Text) & "|��:" & Val(txtRect(2).Text) & "|��:" & Val(txtRect(3).Text) & "|��:" & Val(txtRect(4).Text)
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ�����", cboLCDNum.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�б���������Ӧ", chkFontAutoSizeToList.value
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "���ұ����Ƿ���ʾ������", chkShowDeptName.value
    
    If cboCurDept.Text <> "" Then
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "����ִ�м�", cboCurDept.Text & "-" & cboCurRoom.Text
    Else
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "����ִ�м�", ""
    End If

    If mlngStyleType = TShowStyle.ssSingleMan Or mlngStyleType = TShowStyle.ssSingleQueue Then
        For i = 1 To vsfDoctorInfo.Rows - 1
            strDoctorInfo = strDoctorInfo & "|" & Nvl(vsfDoctorInfo.TextMatrix(i, 2)) & "-" & Nvl(vsfDoctorInfo.TextMatrix(i, 0))
            strDoctorPhoto = strDoctorPhoto & "|" & mstrDoctorPhoto(i - 1)
            strWorkingTime = strWorkingTime & "|" & Nvl(vsfDoctorInfo.TextMatrix(i, 1))
            strIntroduction = strIntroduction & "|" & txtIntroduction(i)
        Next
        
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ϣ", strDoctorInfo     '������ְλ
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ����Ƭ", strDoctorPhoto    '
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ֵ��ʱ��", strWorkingTime    '
        SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "ҽ�����", strIntroduction   '
        
        If mlngStyleType = TShowStyle.ssSingleMan Then
            If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\Skin\��������ʽ\" & cboStyleType.Text
            Else
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\zlQueueShow\Skin\��������ʽ\" & cboStyleType.Text
            End If
            
            Exit Sub
        Else
            SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "��ʾ����Ŀ�ĵ�", chkCallTarget.value
            
            If gobjFile.FolderExists(App.Path & "\Skin\��������ʽ") Then
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\Skin\��������ʽ\" & cboStyleType.Text
            Else
                SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\zlQueueShow\Skin\��������ʽ\" & cboStyleType.Text
            End If
        End If
    ElseIf mlngStyleType = TShowStyle.ssMultiQueue Then
        If gobjFile.FolderExists(App.Path & "\Skin\�������ʽ") Then
            SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\Skin\�������ʽ\" & cboStyleType.Text
        Else
            SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "Ƥ����ʽ", App.Path & "\zlQueueShow\Skin\�������ʽ\" & cboStyleType.Text
        End If
    End If
    
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�׶��ı�", txtRemarkInfo.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "������ʾ", chkScrollDisplay.value
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "�Ŷ��б���ʾ��", txtQueueRows.Text
    SaveSetting "ZLSOFT", G_STR_REGPATH & "\" & mlngWindowNo, "���Ҽ��", txtDeptInfo.Text
End Sub

Private Sub cmdSetDocPhoto_Click()
    Dim strFileName As String
    Dim arrByte() As Byte
    Dim arrPic() As String
    Dim lngCount As Long, lngFileSize As Long
On Error GoTo ErrorHand
    dlgDoctorPhoto.Filter = "(*.jpg)|*.jpg|(*.gif)|*.gif|(*.bmp)|*.bmp|(*.*)|*.*"
    dlgDoctorPhoto.ShowOpen

    strFileName = dlgDoctorPhoto.FileName

    If strFileName = "" Then Exit Sub

    '��ȡ�ļ�����
    lngFileSize = FileLen(strFileName)

    ReDim arrByte(0 To lngFileSize - 1) '������ֵ����
    ReDim arrPic(0 To lngFileSize - 1) '������ֵ����

    Open strFileName For Binary As #1
    Get #1, , arrByte
    Close #1

    '���ֽ�ת��Ϊ16����
    For lngCount = LBound(arrByte) To UBound(arrByte)
        arrPic(lngCount) = Hex(arrByte(lngCount))
        If Len(arrPic(lngCount)) = 1 Then arrPic(lngCount) = "0" & arrPic(lngCount)
    Next

    imgDoctorPhoto(vsfDoctorInfo.RowSel).Picture = LoadPicture(strFileName)
    
    mstrDoctorPhoto(vsfDoctorInfo.RowSel - 1) = Join(arrPic, "")
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub cmd��ʾ�豸����_Click()
    Call InitOldLCDShow
    
    Call gobjQueueShow.zlSetup(Me)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If KeyAscii = vbKeyEscape Then Unload Me
Exit Sub
ErrorHand:
    Debug.Print Err.Description
    Err.Clear
End Sub

Private Sub optCustom_Click()
    Dim i As Integer

On Error GoTo ErrorHand
    frmCustom.Enabled = True
    
    For i = 1 To txtRect.Count
        lblRect(i).Enabled = True
        txtRect(i).Enabled = True
    Next
    
    txtRect(1).Text = 0
    txtRect(2).Text = 0
    txtRect(3).Text = Screen.Width / Screen.TwipsPerPixelX
    txtRect(4).Text = Screen.Height / Screen.TwipsPerPixelY
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub optFullScreen_Click()
    Dim i As Integer
On Error GoTo ErrorHand
    frmCustom.Enabled = False
    
    For i = 1 To txtRect.Count
        lblRect(i).Enabled = False
        txtRect(i).Enabled = False
        txtRect(i).Text = ""
    Next
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtQueueRows_Change()
On Error GoTo ErrorHand

    If txtQueueRows.Text <= 0 Then txtQueueRows.Text = 1
    If txtQueueRows.Text > mlngQueueListMaxRows Then txtQueueRows.Text = mlngQueueListMaxRows
    txtQueueRows.Text = Val(txtQueueRows.Text)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtQueueRows_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHand
    If InStr("01234567890." & Chr(8), Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub txtRect_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrorHand
    If InStr("01234567890." & Chr(8), Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHand
    With vsfDoctorInfo
        If .ColSel <> 0 Then Exit Sub
        If vsfDoctorInfo.TextMatrix(.RowSel, .ColSel) = "" Then Exit Sub

        mrsRecord.Filter = ""
        mrsRecord.Filter = "����='" & Trim(vsfDoctorInfo.TextMatrix(.RowSel, .ColSel)) & "'"

        If mrsRecord.RecordCount > 0 Then vsfDoctorInfo.TextMatrix(.RowSel, 2) = Nvl(mrsRecord!��Աid)
        
        .AutoSize 0, vsfDoctorInfo.Cols - 1
    End With
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrorHand
    FinishEdit = True
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_KeyDown(KeyCode As Integer, Shift As Integer)
'�������б��а��¡�delete����ʱ����ʾ�Ƿ�ɾ��ҽ�������Ϣ����
    
On Error GoTo ErrorHand
    If KeyCode = vbKeyDelete Then  '���ҽ�������Ϣ����
        If MsgBox("�˲������������ҽ���������Ϣ���ã��Ƿ������", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        
        Call ClearDoctorInfo
    End If
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub ClearDoctorInfo()
'���ҽ�������Ϣ����
    Dim i As Integer, j As Integer
    
    For i = 1 To vsfDoctorInfo.Rows - 1
        For j = 0 To vsfDoctorInfo.Cols - 1
            vsfDoctorInfo.TextMatrix(i, j) = ""
        Next
        
        imgDoctorPhoto(i).Picture = imgDoctorPhoto(0).Picture
    Next
    
    For i = 0 To UBound(mstrDoctorPhoto)
        mstrDoctorPhoto(i) = ""
    Next
    
    For i = 0 To txtIntroduction.Count - 1
        txtIntroduction(i).Text = ""
    Next
End Sub

Private Sub vsfQueueSetup_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrorHand
    Cancel = True
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfQueueSetup_Click()
    Dim i As Integer, j As Integer
    Dim lngRowSel As Long, lngColSel As Long

On Error GoTo ErrorHand

    lngRowSel = vsfQueueSetup.RowSel
    lngColSel = vsfQueueSetup.ColSel
    
    If lngRowSel <= 0 Then Exit Sub
    If vsfQueueSetup.TextMatrix(lngRowSel, lngColSel) = "" Then Exit Sub

    Select Case mlngStyleType
        Case TShowStyle.ssSingleMan
            For i = 1 To vsfQueueSetup.Rows - 1
                For j = 0 To vsfQueueSetup.Cols - 1
                    '��PACS�Ŷ�ҵ���£�������ʱ��ѡ��ͬһ�������µ�һ����������
                    If glngBusinessType = TBusinessType.btPacs Then
                        If vsfQueueSetup.TextMatrix(i, j) <> "" And j <> lngColSel Then
                            vsfQueueSetup.Cell(flexcpChecked, i, j) = 2
                        End If
                    Else    '����ҵ�������ֻ��ѡ��һ������
                        If vsfQueueSetup.TextMatrix(i, j) <> "" And (i <> lngRowSel Or j <> lngColSel) Then
                            vsfQueueSetup.Cell(flexcpChecked, i, j) = 2
                        End If
                    End If
                Next
            Next
            
        Case TShowStyle.ssSingleQueue  '������ʱֻ��ѡ��һ��ִ�м�
            For i = 1 To vsfQueueSetup.Rows - 1
                For j = 0 To vsfQueueSetup.Cols - 1
                    If vsfQueueSetup.TextMatrix(i, j) <> "" And (i <> lngRowSel Or j <> lngColSel) Then
                        vsfQueueSetup.Cell(flexcpChecked, i, j) = 2
                    End If
                Next
            Next

        Case TShowStyle.ssMultiQueue  '�����ʱ������ѡ��
        
        Case TShowStyle.ssOld '�ϰ汾ʱ������ѡ��
        ''''''''''''''
    End Select
    
    If vsfQueueSetup.Cell(flexcpChecked, lngRowSel, lngColSel) = 1 Then
        If mlngStyleType = TShowStyle.ssMultiQueue Then mintSelectedQueueNum = mintSelectedQueueNum - 1
        
        vsfQueueSetup.Cell(flexcpChecked, lngRowSel, lngColSel) = 2
        vsfDoctorInfo.Editable = flexEDNone         'û��ѡ��ִ�м�ʱ����������ҽ��ֵ����Ϣ
        
        If mlngStyleType = TShowStyle.ssSingleMan Or mlngStyleType = TShowStyle.ssSingleQueue Then vsfDoctorInfo.ColComboList(0) = ""
    Else
        If mlngStyleType = TShowStyle.ssMultiQueue Then
            If mintSelectedQueueNum >= mlngQueueListMaxRows Then
                MsgBox "����ѡ����" & mintSelectedQueueNum & "�����У�" & vbCrLf & "�������ѡ��" & mlngQueueListMaxRows & "�����У�", vbExclamation, gstrSysName
                Exit Sub
            End If
        
            mintSelectedQueueNum = mintSelectedQueueNum + 1
        End If
        
        vsfQueueSetup.Cell(flexcpChecked, lngRowSel, lngColSel) = 1
    End If
    
    If mlngStyleType = TShowStyle.ssMultiQueue Then
        vsfQueueSetup.ToolTipText = "����ѡ����" & mintSelectedQueueNum & "�����У�" & vbCrLf & "�������ѡ��" & mlngQueueListMaxRows & "�����У�"
    End If
    Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub vsfDoctorInfo_SelChange()
On Error GoTo ErrorHand
    Call ShowImageAndIntroduction(vsfDoctorInfo.RowSel)
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub

Private Sub ShowImageAndIntroduction(ByVal intIndex As Integer)
    Dim i As Integer
On Error GoTo ErrorHand
    For i = 0 To imgDoctorPhoto.Count - 1
        imgDoctorPhoto(i).Visible = False
        txtIntroduction(i).Visible = False
    Next
    
    imgDoctorPhoto(intIndex).Visible = True
    txtIntroduction(intIndex).Visible = True
Exit Sub
ErrorHand:
    MsgBox Err.Description, vbExclamation, gstrSysName
    Err.Clear
End Sub


