VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdviceReport 
   AutoRedraw      =   -1  'True
   Caption         =   "��ӡִ�е�"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16590
   Icon            =   "frmAdviceReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   16590
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraDetail 
      Height          =   8160
      Left            =   120
      TabIndex        =   11
      Top             =   -15
      Width           =   16425
      Begin VB.PictureBox PicView 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7875
         Left            =   7200
         ScaleHeight     =   7875
         ScaleWidth      =   9135
         TabIndex        =   27
         Top             =   180
         Width           =   9135
         Begin VB.CommandButton cmdRePrint 
            Caption         =   "�ش�ֹͣ���ϴδ�ӡǰ��ҽ��"
            Height          =   345
            Left            =   6645
            TabIndex        =   34
            Top             =   18
            Width           =   2495
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "��ѯ(&Q)"
            Height          =   350
            Left            =   5595
            TabIndex        =   33
            Top             =   15
            Width           =   975
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
            Height          =   7380
            Left            =   0
            TabIndex        =   28
            Top             =   420
            Width           =   9135
            _cx             =   16113
            _cy             =   13017
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
            BackColorSel    =   16771802
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
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
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmAdviceReport.frx":014A
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
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
            OwnerDraw       =   1
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
         Begin MSComCtl2.DTPicker dtpViewBegin 
            Height          =   300
            Left            =   1200
            TabIndex        =   29
            Top             =   45
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   123994115
            CurrentDate     =   37953
         End
         Begin MSComCtl2.DTPicker dtpViewEnd 
            Height          =   300
            Left            =   3480
            TabIndex        =   30
            Top             =   45
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   123994115
            CurrentDate     =   37953
         End
         Begin VB.Label lblTo 
            Caption         =   "~"
            Height          =   135
            Left            =   3345
            TabIndex        =   32
            Top             =   165
            Width           =   135
         End
         Begin VB.Label lblView 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ѵ�ִ��ʱ��"
            Height          =   180
            Left            =   80
            TabIndex        =   31
            Top             =   105
            Width           =   1080
         End
      End
      Begin VB.Frame fraPati 
         BorderStyle     =   0  'None
         Height          =   3960
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   7005
         Begin VB.CheckBox chkOut 
            Caption         =   "������Ժ����"
            Height          =   375
            Left            =   0
            TabIndex        =   39
            Top             =   705
            Width           =   915
         End
         Begin VB.Frame fraBaby 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3600
            TabIndex        =   23
            Top             =   60
            Visible         =   0   'False
            Width           =   3195
            Begin VB.OptionButton optBaby 
               Caption         =   "����ҽ��"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   26
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "����ҽ��"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "Ӥ��ҽ��"
               Height          =   180
               Index           =   2
               Left            =   2175
               TabIndex        =   24
               Top             =   0
               Width           =   1020
            End
         End
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   0
            Width           =   2535
         End
         Begin VB.CommandButton cmdNoPati 
            Caption         =   "ȫ��"
            Height          =   330
            Left            =   90
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   3525
            Width           =   870
         End
         Begin VB.CommandButton cmdAllPati 
            Caption         =   "ȫѡ"
            Height          =   330
            Left            =   90
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   3150
            Width           =   870
         End
         Begin MSComctlLib.ImageList img16 
            Left            =   240
            Top             =   1680
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":0273
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList img32 
            Left            =   255
            Top             =   1215
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":03CD
                  Key             =   "Left"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":0CA7
                  Key             =   "Right"
               EndProperty
            EndProperty
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPati 
            Bindings        =   "frmAdviceReport.frx":1581
            Height          =   3515
            Left            =   1035
            TabIndex        =   37
            Top             =   360
            Width           =   5775
            _cx             =   10186
            _cy             =   6200
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
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmAdviceReport.frx":1595
            ScrollTrack     =   -1  'True
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
         Begin MSComctlLib.ImageList imgPati 
            Left            =   240
            Top             =   2280
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":16E2
                  Key             =   "Child"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����(&I)"
            Height          =   180
            Left            =   0
            TabIndex        =   4
            Top             =   435
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����(&U)"
            Height          =   180
            Left            =   0
            TabIndex        =   2
            Top             =   60
            Width           =   990
         End
      End
      Begin VB.Frame fraline 
         Height          =   30
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   22
         Top             =   3960
         Width           =   6855
      End
      Begin VB.Frame fraReport 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6975
         Begin VB.Frame fraCondition 
            BorderStyle     =   0  'None
            Height          =   3600
            Left            =   4680
            TabIndex        =   15
            Top             =   15
            Width           =   2325
            Begin VB.CheckBox ChkWaitPrint 
               Caption         =   "ֻ��ʾ����ӡ�Ĳ���"
               Height          =   195
               Left            =   0
               TabIndex        =   36
               Top             =   3360
               Width           =   1980
            End
            Begin VB.CommandButton cmdView 
               Caption         =   "�����Ѵ�ӡ��Ϣ"
               Height          =   350
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   0
               Width           =   2055
            End
            Begin VB.CheckBox chk��Ч 
               Caption         =   "����(&L)"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   18
               Top             =   1665
               Value           =   1  'Checked
               Width           =   930
            End
            Begin VB.CheckBox chk��Ч 
               Caption         =   "��ʱ(&T)"
               Height          =   195
               Index           =   1
               Left            =   1080
               TabIndex        =   17
               Top             =   1665
               Value           =   1  'Checked
               Width           =   930
            End
            Begin VB.CheckBox chk�ظ���ӡ 
               Caption         =   "�����Ѵ�ӡ����(&A)"
               Height          =   195
               Left            =   0
               TabIndex        =   16
               Top             =   1965
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Left            =   0
               TabIndex        =   19
               Top             =   825
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   529
               _Version        =   393216
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
               Format          =   123994115
               CurrentDate     =   37953
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Left            =   0
               TabIndex        =   20
               Top             =   1185
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   529
               _Version        =   393216
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
               Format          =   123994115
               CurrentDate     =   37953
            End
            Begin VB.Label lblִ��ʱ�� 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ִ��ʱ��(&E)"
               Height          =   180
               Left            =   0
               TabIndex        =   21
               Top             =   480
               Width           =   1350
            End
         End
         Begin VB.CommandButton cmdSetup 
            Caption         =   "��ӡ����"
            Height          =   330
            Left            =   0
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + S"
            Top             =   405
            Width           =   990
         End
         Begin MSComctlLib.ListView lvwReport 
            Height          =   3600
            Left            =   1035
            TabIndex        =   1
            Top             =   0
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   6350
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "img16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   6615
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ�е�(&R)"
            Height          =   180
            Left            =   180
            TabIndex        =   0
            Top             =   60
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   16590
      TabIndex        =   13
      Top             =   8190
      Width           =   16590
      Begin VB.CommandButton cmdGoOn 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   1560
         TabIndex        =   38
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ(&P)"
         Height          =   350
         Left            =   3870
         TabIndex        =   8
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "�˳�(&X)"
         Height          =   350
         Left            =   5235
         TabIndex        =   9
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Ԥ��(&V)"
         Height          =   350
         Left            =   2760
         TabIndex        =   7
         Top             =   0
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmAdviceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1

Private mMainPrivs As String 'IN:���������������е�Ȩ��,ע����ڲ�ģ��Ȩ��
Private mlng����ID As Long 'IN
Private mlng����ID As Long 'IN
Private mblnOnePati As Boolean 'IN��������ģʽ
Private mstr��ǰ���� As String '�����ӡʱ��ǰ��ӡ��:����ID,��ҳID
Private mstrPrintedID As String '��ӡ����ҽ��ID��
Private mlngҽ������Χ As Long    '0-����ҽ����1-����ҽ����2-Ӥ��ҽ��(������Ϊ-1��0��Ӥ�����)
Private mlngLastRow As Long     '��ӡ�����һ��
Private mintType As Long    '��ӡ���ͣ�0-��ӡ��1-�ش�2-����
Private mbln������˻�ҳ��ӡ As Boolean
Private mbln���� As Boolean
Private mblnӤ������ As Boolean

Private Enum AdviceCol
    col���� = 0
    COL��λ = 1
    COLӤ�� = 2
    col��Ч = 3
    colҽ������ = 4
    COLƵ�� = 5
    COL�ϴδ�ӡʱ�� = 6
    COLֹͣʱ�� = 7
    col��� = 8
    col���ID = 9
End Enum

Private Enum PatiCol
    COL_ѡ�� = 0
    COL_���� = 1
    COL_סԺ�� = 2
    COL_��λ = 3
    COL_סԺҽʦ = 4
    COL_�ѱ� = 5
    COL_����ȼ� = 6
    COL_���� = 7
    COL_��Ժʱ�� = 8
    COL_�������� = 9
    COL_��ҳID = 10
    COL_Ӥ�� = 11
End Enum

Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, _
    ByVal lng����ID As Long, ByVal lng����ID As Long, _
     ByVal blnOnePati As Boolean, Optional ByVal lngҽ������ID As Long, Optional ByVal lngӤ������ID As Long) As Boolean
'������
    mMainPrivs = MainPrivs
    
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    If lngӤ������ID <> 0 Then
        If lngӤ������ID = lngҽ������ID Then
            mlng����ID = lngӤ������ID
        End If
    End If
   
    mblnOnePati = blnOnePati
        
    Me.Show 1, frmParent
End Function

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long, lngUnitID As Long
    Dim lngColor As Long, lngҽ������Χ As Long
    Dim blnIsWowen As Boolean
        
    On Error GoTo errH
    
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    blnIsWowen = DeptIsWoman(0, Get����IDs(lngUnitID))
    mbln���� = blnIsWowen
    If blnIsWowen Then
        'ҽ������Χ
        optBaby(mlngҽ������Χ).value = True
        fraBaby.Visible = True
        '���빴ѡ������˴�ӡ����ʾӤ��
        If mbln������˻�ҳ��ӡ Then
            lngҽ������Χ = IIF(optBaby(0).value, -1, IIF(optBaby(1).value, 0, 1))
            mblnӤ������ = True
        Else
            lngҽ������Χ = 0
            mblnӤ������ = False
        End If
    Else
        fraBaby.Visible = False
        optBaby(0).value = True
        lngҽ������Χ = 0
    End If
    Call SetBabyVisible
    vsPati.Rows = 1
    
    str����IDs = IIF(mblnOnePati, "", zldatabase.GetPara("���Ͳ���", glngSys, pסԺҽ������))
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
    With vsPati
        Set rsTmp = GetPatiRsByUnit(lngUnitID, 0, False, False, chkOut.value = 1, True, lngҽ������Χ)
        vsPati.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            If Not (blnIsWowen And mbln������˻�ҳ��ӡ And rsTmp!Ӥ������ID <> "" And (rsTmp!Ӥ������ID & "") = mlng����ID And rsTmp!Ӥ������ & "" = "") And Not (mblnOnePati And rsTmp!����ID <> mlng����ID) Then
                .RowData(i) = Val(rsTmp!����ID & "")
                .TextMatrix(i, COL_����) = IIF(rsTmp!Ӥ������ & "" = "", rsTmp!���� & "", rsTmp!Ӥ������ & "")
                If rsTmp!Ӥ������ & "" <> "" Then .Cell(flexcpPicture, i, COL_����) = imgPati.ListImages("Child").Picture
                .TextMatrix(i, COL_סԺ��) = IIF(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
                .TextMatrix(i, COL_��λ) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
                .TextMatrix(i, COL_סԺҽʦ) = IIF(IsNull(rsTmp!סԺҽʦ), "", rsTmp!סԺҽʦ)
                .TextMatrix(i, COL_�ѱ�) = IIF(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
                .TextMatrix(i, COL_����ȼ�) = IIF(IsNull(rsTmp!����ȼ�), "", rsTmp!����ȼ�)
                .TextMatrix(i, COL_����) = IIF(IsNull(rsTmp!����), "", rsTmp!����)
                .TextMatrix(i, COL_��Ժʱ��) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COL_��������) = NVL(rsTmp!��������)
                .TextMatrix(i, COL_��ҳID) = rsTmp!��ҳID & ""
                .TextMatrix(i, COL_Ӥ��) = rsTmp!Ӥ����� & ""
                
                '������ɫ
                lngColor = zldatabase.GetPatiColor(NVL(rsTmp!��������))
                .Cell(flexcpForeColor, i, COL_סԺ��, i, COL_סԺ��) = lngColor
                .Cell(flexcpForeColor, i, COL_��������, i, COL_��������) = lngColor
                
                '�ϴ��Ƿ�ѡ��
                If lngUnitID = lng����ID And str����IDs <> "" Then
                    If InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 Then
                        .TextMatrix(i, COL_ѡ��) = -1
                        If k = 0 Then 'Ϊ�˿�����ѡ���
                            .ShowCell i, COL_����
                            k = 1
                        End If
                    End If
                ElseIf rsTmp!����ID = mlng����ID Then
                    .TextMatrix(i, COL_ѡ��) = -1
                    .ShowCell i, COL_����
                End If
            End If
            rsTmp.MoveNext
        Next
        For i = .Rows - 1 To 1 Step -1
            If Val(.RowData(i) & "") = 0 Then .RemoveItem i
        Next
    End With
    mlng����ID = lngUnitID
    '��ѡ�˲Ŵ���
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ChkWaitPrint_Click()
    Call LoadWaitPrint(ChkWaitPrint.value = 1)
    Call SetBabyVisible
End Sub

Private Sub LoadWaitPrint(ByVal blnIsLoadWart As Boolean)
'���ܣ����˴���ӡ�Ĳ���
'������LoadWaitPrint-true���˴���ӡ�ģ�������ʾȫ��
    Dim i As Long
    Dim strFiter As String, str��Ч As String
    Dim strSql As String, rsTmp As Recordset
    
    If blnIsLoadWart Then
        'ɾ������Ҫ��ӡ��
        'Ӥ��
        If optBaby(1).value = True Then
            strFiter = " And Nvl(a.Ӥ��, 0) = 0 "
        ElseIf optBaby(2).value = True Then
            strFiter = " And Nvl(a.Ӥ��, 0) <> 0 "
        End If
        '����
        Select Case UCase(Mid(lvwReport.SelectedItem.Key, 2))
        Case "ZL1_INSIDE_1254_4" '��ҩ��
            strFiter = strFiter & " And B.��������='2' And B.ִ�з���=4"
        Case "ZL1_INSIDE_1254_5" 'ע�䵥
            strFiter = strFiter & " And B.��������='2' And B.ִ�з���=2"
        Case "ZL1_INSIDE_1254_6" '��Һ��
            strFiter = strFiter & " And B.��������='2' And B.ִ�з���=1"
        End Select
        '��Ч
        If chk��Ч(0).value = 1 And chk��Ч(1).value = 1 Then
            strFiter = strFiter & " And a.ҽ����Ч In(0,1) "
        ElseIf chk��Ч(0).value = 1 And chk��Ч(1).value = 0 Then
            strFiter = strFiter & " And a.ҽ����Ч =0 "
        ElseIf chk��Ч(0).value = 0 And chk��Ч(1).value = 1 Then
            strFiter = strFiter & " And a.ҽ����Ч =1 "
        End If
        strSql = "Select distinct a.����id, a.��ҳid," & IIF(optBaby(1).value = False And fraBaby.Visible And mbln������˻�ҳ��ӡ, "NVL(A.Ӥ��,0)", "0") & " as Ӥ��" & vbNewLine & _
                "From ����ҽ����¼ A, ������ĿĿ¼ B ,������Ϣ C,������ҳ D,��Ժ���� R" & vbNewLine & _
                "Where a.������Ŀid = b.Id And a.���id Is Null And a.У��ʱ�� Is Not Null AND C.����ID=D.����ID and C.��ҳID=D.��ҳID And a.ҽ��״̬ <> 4 And ([1]" & vbNewLine & _
                "       <= a.ִ����ֹʱ�� Or a.ִ����ֹʱ�� Is Null) And [2]" & vbNewLine & _
                "      >= a.��ʼִ��ʱ��  And" & vbNewLine & _
                "      a.����id=c.����id and a.��ҳid=c.��ҳid and C.����ID=R.����ID And C.��ǰ����ID=R.����ID and (R.����id = [4] or D.Ӥ������ID = [4]) And" & vbNewLine & _
                "      Zl_Adviceexecount(a.Id, [1], [2], [5], [3]) > 0" & strFiter
        On Error GoTo errH
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, CDate(dtpBegin.value), CDate(dtpEnd.value), lvwReport.SelectedItem.Tag, mlng����ID, chk�ظ���ӡ.value)
        
        For i = 1 To vsPati.Rows - 1
            rsTmp.Filter = "����ID=" & vsPati.RowData(i) & " And ��ҳID=" & Val(vsPati.TextMatrix(i, COL_��ҳID)) & " And Ӥ��=" & Val(vsPati.TextMatrix(i, COL_Ӥ��))
            If rsTmp.RecordCount = 0 Then
                '����
                vsPati.RowHidden(i) = True
            Else
                '��ʾ
                vsPati.RowHidden(i) = False
            End If
        Next
    Else
        For i = 1 To vsPati.Rows - 1
            '��ʾ
            vsPati.RowHidden(i) = False
        Next
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chk��Ч_Click(Index As Integer)
    If chk��Ч((Index + 1) Mod 2).value = 0 And chk��Ч(Index).value = 0 Then chk��Ч(Index).value = 1
    '��ѡ�˲Ŵ���
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub chk�ظ���ӡ_Click()
    '��ѡ�˲Ŵ���
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(vsPati, True)
    vsPati.SetFocus
End Sub

Private Sub SelectLVW(objVsg As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objVsg.Rows - 1
        objVsg.TextMatrix(i, COL_ѡ��) = IIF(blnCheck, -1, 0)
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGoOn_Click()
    Call PrintOrPreview(2, 2) '����
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(vsPati, False)
    vsPati.SetFocus
End Sub

Private Sub cmdPreview_Click()
    Call PrintOrPreview(1) 'Ԥ��
End Sub

Private Sub cmdPrint_Click()
    Call PrintOrPreview(2) '��ӡ
End Sub

Private Sub PrintOrPreview(ByVal bytMode As Byte, Optional ByVal bytType As Byte)
'������bytMode=1-Ԥ��,2-��ӡ
'      bytType=0-��ӡ����ҽ����1-ֻ��ӡֹͣ���ϴδ�ӡʱ��֮ǰ��ҽ��,2-���򵥸����˵�ִ�е�
    Dim curDate As Date, strTmp As String, i As Long
    Dim arrPati As Variant, str����IDs As String
    Dim str��Ч As String, str���� As String
    Dim str�ظ���ӡ As String
    Dim datBegin As Date, datEnd As Date
    Dim strReports As String, j As Long, k As Long, z As Long
    Dim lng��ʼ�к� As Long, str��ʼ�к� As String
    Dim strRPTNO As String '1-û�й�ѡ����ȡ��궨λ�У�2-�й�ѡ����궨λ��δ��ѡ����ȡ��һ����ѡ�У�3-�й�ѡ����궨λ�ڹ�ѡ���ϣ���ȡ��궨λ��
    
    mintType = bytType
    If bytType = 1 Then
        datBegin = dtpViewBegin.value: datEnd = dtpViewEnd.value
    Else
        datBegin = dtpBegin.value: datEnd = dtpEnd.value
    End If
    
    If datBegin >= datEnd Then
        MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
        IIF(bytType = 1, dtpViewBegin, dtpBegin).SetFocus: Exit Sub
    End If
    
    mstrPrintedID = ""
    
    '���汨��������
    str����IDs = ""
    arrPati = Array()
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, COL_ѡ��)) = -1 And vsPati.RowHidden(i) = False Then
            str����IDs = str����IDs & "," & vsPati.RowData(i)
            ReDim Preserve arrPati(UBound(arrPati) + 1)
            arrPati(UBound(arrPati)) = vsPati.RowData(i) & "," & vsPati.TextMatrix(i, COL_��ҳID) & "," & Val(vsPati.TextMatrix(i, COL_Ӥ��))
        End If
    Next
    str����IDs = Mid(str����IDs, 2)
    For i = 1 To lvwReport.ListItems.Count
        If lvwReport.ListItems(i).Checked Then
            strReports = strReports & "," & lvwReport.ListItems(i).Tag
            strRPTNO = strRPTNO & "," & Mid(lvwReport.ListItems(i).Key, 2)
        End If
    Next
    strReports = Mid(strReports, 2)
    
    If strRPTNO <> "" Then
        If InStr(strRPTNO & ",", "," & Mid(lvwReport.SelectedItem.Key, 2) & ",") > 0 Then
            strRPTNO = Mid(lvwReport.SelectedItem.Key, 2)
        Else
            strRPTNO = Split(strRPTNO, ",")(1)
        End If
    Else
        strRPTNO = Mid(lvwReport.SelectedItem.Key, 2)
    End If

    
    '���ѡ����δ��ѡ������ʾ�û�ֻ��ӡ��ѡ�˵�
    If bytMode = 2 And bytType = 0 Then
        If strReports <> "" And lvwReport.SelectedItem.Checked = False Then
            If MsgBox("��ǰ��ѡ��" & UBound(Split(strReports, ",")) + 1 & "�ű������δ�ӡֻ��ӡ��ѡ�˵ı���,�Ƿ������", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    '�����飬������������˴�ӡ���ҵ������ˣ�����ִ�е�
    If bytType = 2 Then
        If mbln������˻�ҳ��ӡ = False Then
            MsgBox "��������""������˻�ҳ��ӡ""����������", vbInformation, gstrSysName
            Exit Sub
        ElseIf UBound(Split(str����IDs, ",")) <> 0 Then
            MsgBox "ֻ��ѡ��һ�����˽�������", vbInformation, gstrSysName
            vsPati.SetFocus: Exit Sub
        End If
    End If
    If str����IDs = "" Then
        MsgBox "������ѡ��һ��סԺ���ˡ�", vbInformation, gstrSysName
        vsPati.SetFocus: Exit Sub
    End If

    '�����ӡӤ��ҽ��ʱ���������
    If bytType = 0 Then
        If UBound(Split(str����IDs, ",")) = 0 And Val(str����IDs) = mlng����ID Then
            Call zldatabase.SetPara("���Ͳ���", "", glngSys, pסԺҽ������)
        Else
            Call zldatabase.SetPara("���Ͳ���", cboUnit.ItemData(cboUnit.ListIndex) & ":" & str����IDs, glngSys, pסԺҽ������)
        End If
        Call zldatabase.SetPara("ִ�е���ӡ����", strReports, glngSys, pסԺҽ������)
    End If
    '��������
    curDate = zldatabase.Currentdate
    If bytType = 0 Then
        Call zldatabase.SetPara("���ñ�����Ч", chk��Ч(0).value & chk��Ч(1).value, glngSys, pסԺҽ������)
        Call zldatabase.SetPara("���ñ���ʼʱ��", Format(dtpBegin.value, "HH:mm:ss"), glngSys, pסԺҽ������)
        Call zldatabase.SetPara("���ñ���ʼ���", Int(CDate(Format(dtpBegin.value, "yyyy-MM-dd")) - CDate(Format(curDate, "yyyy-MM-dd"))), glngSys, pסԺҽ������)
        Call zldatabase.SetPara("���ñ������ʱ��", Format(dtpEnd.value, "HH:mm:ss"), glngSys, pסԺҽ������)
        Call zldatabase.SetPara("���ñ���������", Int(CDate(Format(dtpEnd.value, "yyyy-MM-dd")) - CDate(Format(curDate, "yyyy-MM-dd"))), glngSys, pסԺҽ������)
    End If
    
    'ֻ��ʾ����ӡ�Ĳ���
    '��ʹ�ò������棬1�����ܲ����ã�2���ұ������ִ��ʱ����Ч���й��ˣ��´ν���ʱʱ���ַ����˱仯��
    
    '��Ч����
    If chk��Ч(0).value = 1 And chk��Ч(1).value = 1 Then
        str��Ч = "0,1"
    ElseIf chk��Ч(0).value = 1 Then
        str��Ч = "0"
    Else
        str��Ч = "1"
    End If
    
    str�ظ���ӡ = IIF(chk�ظ���ӡ.Visible, chk�ظ���ӡ.value, 0)
    If bytType = 1 Then str�ظ���ӡ = "2"
    
    '��������
    If mbln������˻�ҳ��ӡ = False Then
        If UBound(arrPati) = 0 Then
            '��������
            str���� = "(" & Mid(arrPati(0), 1, Len(arrPati(0)) - 2) & ")"
        Else
            '�������
            strTmp = ""
            For i = 0 To UBound(arrPati)
                strTmp = strTmp & "," & Replace(Mid(arrPati(i), 1, Len(arrPati(i)) - 2), ",", ":")
            Next
            strTmp = Mid(strTmp, 2)
            str���� = " Select  C1 As ����ID,C2 As ��ҳID From Table(f_Num2list2('" & strTmp & "')) "
        End If
        
        'ִ��
        If strReports = "" Or bytMode = 1 Then ' ֻԤ����ǰ��
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strRPTNO, Me, _
                "��ʼʱ��=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                "����ʱ��=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                "��Ч=" & str��Ч, "����=" & str����, "��ʼ�к�=1", "������ID=0", "�ظ���ӡ=" & str�ظ���ӡ, "����ID=" & lvwReport.SelectedItem.Tag, "ҽ������Χ=" & IIF(optBaby(0).value, -1, IIF(optBaby(1).value, 0, -2)), bytMode)
        Else
            '������ӡ
            For i = 1 To lvwReport.ListItems.Count
                If lvwReport.ListItems(i).Checked Then
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, Mid(lvwReport.ListItems(i).Key, 2), Me, _
                        "��ʼʱ��=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                        "����ʱ��=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                        "��Ч=" & str��Ч, "����=" & str����, "������ID=0", "��ʼ�к�=1", "�ظ���ӡ=" & str�ظ���ӡ, "����ID=" & lvwReport.ListItems(i).Tag, "ҽ������Χ=" & IIF(optBaby(0).value, -1, IIF(optBaby(1).value, 0, -2)), bytMode)
                End If
            Next
        End If
    Else
        '������˽���
        Screen.MousePointer = 11
        If vsPati.Visible Then vsPati.SetFocus: Me.Refresh
        '������ӡ
        '�����Ԥ������һ������δ��ѡ����Ĭ��ѡ���б���
        For z = 1 To IIF(strReports = "" Or bytMode = 1 Or bytType = 2, 1, lvwReport.ListItems.Count)
            If lvwReport.ListItems(z).Checked Or (strReports = "" Or bytMode = 1 Or bytType = 2) Then
                For i = 0 To UBound(arrPati)
                    str���� = "_" & Split(arrPati(i), ",")(0)
                    j = vsPati.FindRow(Val(Split(arrPati(i), ",")(0)))
                    If (optBaby(2).value Or optBaby(0).value) And mbln������˻�ҳ��ӡ And j <> -1 Then
                        'Ӥ���б�
                        For k = j To vsPati.Rows - 1
                            If Val(Split(arrPati(i), ",")(0)) = vsPati.RowData(k) And Val(vsPati.TextMatrix(k, COL_Ӥ��)) = Val(Split(arrPati(i), ",")(2)) Then
                                j = k: Exit For
                            End If
                        Next
                    End If
                    If j <> -1 Then
                        If Val(Split(arrPati(i), ",")(2)) <> Val(vsPati.TextMatrix(j, COL_Ӥ��)) Then
                            
                            j = vsPati.FindRow(Val(Split(arrPati(i), ",")(0)), j + 1)
                        End If
                        vsPati.TextMatrix(j, COL_ѡ��) = -1
                        Call vsPati.ShowCell(j, COL_����)
                        vsPati.Refresh: Me.Refresh
                        
                        str���� = "(" & Mid(arrPati(i), 1, Len(arrPati(i)) - 2) & ")"
                        mstr��ǰ���� = Mid(arrPati(i), 1, Len(arrPati(i)) - 2)
                        '����ʱ��ʾ
                        lng��ʼ�к� = 1
                        If bytType = 2 Then
                            lng��ʼ�к� = Get��ʼ�к�(Val(Split(arrPati(i), ",")(0)), Val(Split(arrPati(i), ",")(1)), Val(Split(arrPati(i), ",")(2)), Val(lvwReport.SelectedItem.Tag))
                            str��ʼ�к� = lng��ʼ�к�
                            'ȷ����ʼ�к�
                            If zlCommFun.ShowMsgBox("������ʼ�к�", "��ǰ���ӵ�" & str��ʼ�к� & "�п�ʼ������ȷ�ϡ�" & vbCrLf & "��������������ȷ����ʼ�кš�", "!ȷ��(&O),?ȡ��(&C)", Me, vbInformation, , , , , , "��ʼ�к�", 10, str��ʼ�к�) <> "ȷ��" Then
                                Exit Sub
                            End If
                            lng��ʼ�к� = Val(str��ʼ�к�)
                            If lng��ʼ�к� = 0 Then
                                MsgBox "�������������ȷ��", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                        'Ԥ����δ��ѡ����������Ե�ǰѡ��ı��������ǹ�ѡ�ı���
                        If strReports = "" Or bytMode = 1 Or bytType = 2 Then ' ֻԤ����ǰ��
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, strRPTNO, Me, _
                                "��ʼʱ��=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                                "����ʱ��=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                                "��Ч=" & str��Ч, "����=" & str����, "�ظ���ӡ=" & str�ظ���ӡ, _
                                "����ID=" & lvwReport.SelectedItem.Tag, "PressWorkFirst=" & IIF(lng��ʼ�к� = 1, 0, 1), "��ʼ�к�=" & lng��ʼ�к�, _
                                "������ID=" & Val(Split(arrPati(i), ",")(0)), "ҽ������Χ=" & IIF(optBaby(1).value, 0, Val(Split(arrPati(i), ",")(2))), bytMode)
                        Else
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, Mid(lvwReport.ListItems(z).Key, 2), Me, _
                                "��ʼʱ��=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                                "����ʱ��=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                                "��Ч=" & str��Ч, "����=" & str����, "�ظ���ӡ=" & str�ظ���ӡ, _
                                "����ID=" & lvwReport.ListItems(z).Tag, "PressWorkFirst=" & IIF(lng��ʼ�к� = 1, 0, 1), "��ʼ�к�=" & lng��ʼ�к�, _
                                "������ID=" & Val(Split(arrPati(i), ",")(0)), "ҽ������Χ=" & IIF(optBaby(1).value, 0, Val(Split(arrPati(i), ",")(2))), bytMode)
                        End If
                        'ֻԤ����һ�����˵�����
                        If bytMode = 1 And i = 0 Then Exit For
                    End If
                Next
            End If
        Next
        Screen.MousePointer = 0
    End If
    If bytMode = 2 Then
        '��ѡ�˲Ŵ���
        If ChkWaitPrint.value = 1 Then
            Call LoadWaitPrint(ChkWaitPrint.value)
        End If
    End If
End Sub

Private Function Get��ʼ�к�(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, ByVal lng����ID As Long) As Long
'���ܣ�ȡ��Ӧ�����Ӧ���ˣ���Ӥ�������ϴδ�ӡ��ĩ���к�
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngRow As Long
    
    On Error GoTo errH
    
    strSql = "select ĩҳĩ�к� from ����ִ�е���ӡ where ����ID=[1] And ��ҳID=[2]  And ����ID=[4] And Ӥ��=[3]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳID, intӤ��, lng����ID)
    If rsTmp.RecordCount > 0 Then lngRow = Val(rsTmp!ĩҳĩ�к� & "")
    Get��ʼ�к� = lngRow + 1
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
                                
Private Sub cmdRePrint_Click()
     Call PrintOrPreview(2, 1) '��ӡֹͣ���ϴδ�ӡʱ��֮ǰ��ҽ��
End Sub

Private Sub cmdSelect_Click()
    Call LoadAdvice
End Sub

Private Sub cmdSetup_Click()
    Call mobjReport.ReportPrintSet(gcnOracle, glngSys, Mid(lvwReport.SelectedItem.Key, 2), Me)
End Sub

Private Sub cmdView_Click()
    If cmdView.Caption = "�����Ѵ�ӡ��Ϣ" Then
        cmdView.Caption = "��ʾ�Ѵ�ӡ��Ϣ"
        PicView.Visible = False
        If Me.WindowState = 0 Then Me.Width = fraCondition.Left + fraCondition.Width + 650
    Else
        cmdView.Caption = "�����Ѵ�ӡ��Ϣ"
        PicView.Visible = True
        If Me.WindowState = 0 Then Me.Width = PicView.Left + PicView.Width + 580
    End If
    Call Form_Resize
End Sub

Private Sub dtpBegin_Change()
    '��ѡ�˲Ŵ���
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub dtpEnd_Change()
    '��ѡ�˲Ŵ���
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdAllPati.Visible Then Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdNoPati.Visible Then Call cmdNoPati_Click
    ElseIf KeyCode = vbKeyS And Shift = vbCtrlMask Then
        Call cmdSetup_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strTmp As String, lngTmp As Long
        
    Call InitReports '��ȡ����
    If lvwReport.ListItems.Count = 0 Then
        MsgBox "��û��Ȩ�޴�ӡ�κ�һ�ű�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    'ȱʡ����ʱ��
    curDate = zldatabase.Currentdate
    
    'ȱʡҽ����Ч
    strTmp = zldatabase.GetPara("���ñ�����Ч", glngSys, pסԺҽ������, "11", Array(chk��Ч(0), chk��Ч(1)))
    chk��Ч(0).value = Val(Left(strTmp, 1))
    chk��Ч(1).value = Val(Right(strTmp, 1))

    
    strTmp = zldatabase.GetPara("���ñ���ʼʱ��", glngSys, pסԺҽ������, "00:00:00", Array(dtpBegin))
    lngTmp = Val(zldatabase.GetPara("���ñ���ʼ���", glngSys, pסԺҽ������, "0", Array(dtpBegin)))
    dtpBegin.value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    dtpViewBegin.value = dtpBegin.value
    
    strTmp = zldatabase.GetPara("���ñ������ʱ��", glngSys, pסԺҽ������, "23:59:59", Array(dtpEnd))
    lngTmp = Val(zldatabase.GetPara("���ñ���������", glngSys, pסԺҽ������, "0", Array(dtpEnd)))
    dtpEnd.value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    dtpViewEnd.value = dtpEnd.value
    
    If mblnOnePati = False Then mbln������˻�ҳ��ӡ = Val(zldatabase.GetPara("���ñ��������ӡ", glngSys, pסԺҽ������, "0")) = 1
    mlngҽ������Χ = Val(zldatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0"))
    If mblnOnePati Then
        cboUnit.Enabled = False
        mbln������˻�ҳ��ӡ = True
        ChkWaitPrint.Visible = False
    End If
    Call InitUnits '��ȡ����/����
    
    Call zlControl.LvwFlatColumnHeader(lvwReport)
    '֧������
    vsPati.ExplorerBar = flexExSort
    vsPati.Editable = flexEDKbdMouse

    If Val(zldatabase.GetPara("��ʾ�Ѵ�ӡ��Ϣ", glngSys, pסԺҽ������, "0")) = 0 Then
        Call cmdView_Click
    End If

    
    Call RestoreWinState(Me, App.ProductName, IIF(mblnOnePati, "OnePati", ""))
    
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSql As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mMainPrivs, "ȫԺ����") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSql = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = strSql & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSql & ") Group by ID,����,���� Order by ����"
    End If
    
    cboUnit.Clear
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 14205 And PicView.Visible Then Me.Width = 14205
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdPrint.Left = cmdCancel.Left - cmdPrint.Width - 240
    cmdPreview.Left = cmdPrint.Left - cmdPreview.Width - 30
    cmdGoOn.Left = cmdPreview.Left - cmdGoOn.Width - 120
    
    fraDetail.Width = Me.ScaleWidth - 240
    fraDetail.Height = Me.ScaleHeight - picBottom.Height - 120
    
    fraReport.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - 240
    lvwReport.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - lvwReport.Left - fraCondition.Width - 300
    lvwReport.ColumnHeaders(1).Width = lvwReport.Width - 140
        
    lvwReport.Height = fraReport.Height - lvwReport.Top - 60
    fraCondition.Left = lvwReport.Left + lvwReport.Width + 120
    
    fraline.Top = fraReport.Top + fraReport.Height
    fraline.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - fraline.Left - 120
    
    PicView.Left = fraCondition.Width + fraCondition.Left + 100
    PicView.Height = Me.Height - 1400
    If vsAdvice.Visible Then vsAdvice.Height = Me.Height - 1300
    
    fraPati.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - fraPati.Left - 120
    vsPati.Width = fraPati.Width - vsPati.Left
    
    fraPati.Top = fraline.Top + fraline.Height + 60
    fraPati.Height = fraDetail.Height - fraline.Top - 120
    vsPati.Height = fraPati.Height - vsPati.Top - 60
    
    cmdNoPati.Top = vsPati.Top + vsPati.Height - 30 - cmdNoPati.Height
    cmdAllPati.Top = cmdNoPati.Top - cmdAllPati.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
    mlng����ID = 0
    mlng����ID = 0
    Call zldatabase.SetPara("��ʾ�Ѵ�ӡ��Ϣ", IIF(PicView.Visible, 1, 0), glngSys, pסԺҽ������)
    'Set mobjReport = Nothing '�Զ������Ա㱨�����еĻ������ظ�ʹ��
    
    Call SaveWinState(Me, App.ProductName, IIF(mblnOnePati, "OnePati", ""))
End Sub

Private Function InitReports() As Boolean
'���ܣ���ȡ���ñ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim objItem As ListItem
    Dim strReports As String
    
    On Error GoTo errH
    
    strReports = zldatabase.GetPara("ִ�е���ӡ����", glngSys, pסԺҽ������)
    strSql = "Select ID,���,����,���� From zlReports Where ϵͳ=[1] And ��� IN('ZL1_INSIDE_1254_4','ZL1_INSIDE_1254_5','ZL1_INSIDE_1254_6'" & _
         ",'ZL1_INSIDE_1254_7','ZL1_INSIDE_1254_8','ZL1_INSIDE_1254_9','ZL1_INSIDE_1254_10','ZL1_INSIDE_1254_11'" & _
         ",'ZL1_INSIDE_1254_12','ZL1_INSIDE_1254_13','ZL1_INSIDE_1254_14','ZL1_INSIDE_1254_15','ZL1_INSIDE_1254_16') Order by ID"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, glngSys)
    Do While Not rsTmp.EOF
        If InStr(GetInsidePrivs(pסԺҽ������), ";" & rsTmp!���� & ";") > 0 Then
            Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!���, rsTmp!����, , 1)
            objItem.Tag = Val(rsTmp!ID)
            If InStr("," & strReports & ",", "," & Val(rsTmp!ID) & ",") > 0 And strReports <> "" Then
                objItem.Checked = True
            End If
        End If
        rsTmp.MoveNext
    Loop
    InitReports = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub fraline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next
        If fraReport.Height + Y < 1000 Or fraReport.Height - Y < 500 Then Exit Sub
        If fraReport.Height + Y > (fraDetail.Height - cmdAllPati.Height * 7) Then Exit Sub
        
        fraline.Top = fraline.Top + Y
        fraReport.Height = fraReport.Height + Y
        fraPati.Top = fraPati.Top + Y
        fraPati.Height = fraPati.Height - Y
        
        Call Form_Resize
    End If
End Sub

Private Sub lvwReport_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub lvwReport_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '��ѡ�˲Ŵ���
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
'���ܣ��������ݴ�ӡ�¼�����¼ҽ����ӡ������
'˵�����������������Ҫ��ӡʱ���ǲ��ἤ����¼���
    If ID <> 0 Then
        If InStr(mstrPrintedID & ",", "," & ID & ",") = 0 Then
            mstrPrintedID = mstrPrintedID & "," & ID
        End If
        mlngLastRow = Row
    End If
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'���ܣ���ӡ֮�󣬸���ҽ�����ϴδ�ӡʱ��
    Dim rsTmp As ADODB.Recordset
    Dim arrPati As Variant, arrSQL As Variant
    Dim strSql As String, i As Long
    Dim strSQLPati As String, strPatis As String, strTemp As String
    Dim strThis As String, p As Long, n As Long, lngParStar As Long
    Dim varPar(0 To 10) As String, blnTrans As Boolean, lngReportID As Long
    
    On Error GoTo errH
    
    If mstrPrintedID <> "" Then
        mstrPrintedID = Mid(mstrPrintedID, 2)
        n = 0
        Do While True
            If Len(mstrPrintedID) < 4000 Then
                p = Len(mstrPrintedID) + 1
            Else
                p = InStrRev(Mid(mstrPrintedID, 1, 4000), ",")
            End If
            strThis = Mid(mstrPrintedID, 1, p - 1)
            
            If n > 10 Then
                '̫�����ٴ���ʹ֮����
                varPar(10) = varPar(10) & "," & strThis
            Else
                varPar(n) = strThis
            End If
            
            n = n + 1
            mstrPrintedID = Mid(mstrPrintedID, p + 1)
            If mstrPrintedID = "" Then Exit Do
        Loop
        For i = 1 To lvwReport.ListItems.Count
            If Mid(lvwReport.ListItems(i).Key, 2) = ReportNum Then lngReportID = Val(lvwReport.ListItems(i).Tag): Exit For
        Next
        arrSQL = Array()
        For i = 0 To UBound(varPar)
            If varPar(i) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_ҽ��ִ�е�_��ӡ('" & varPar(i) & "'," & lngReportID & "," & _
                    "To_Date('" & Format(IIF(mintType = 1, dtpViewBegin.value, dtpBegin.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(IIF(mintType = 1, dtpViewEnd.value, dtpEnd.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIF(mbln������˻�ҳ��ӡ, mlngLastRow, 0) & ")"
            End If
        Next
    
    Else
        'ҽ����Ч
        If chk��Ч(0).value = 1 And chk��Ч(1).value = 1 Then
            strSql = strSql & " And A.ҽ����Ч IN(0,1)"
        ElseIf chk��Ч(0).value = 1 Then
            strSql = strSql & " And A.ҽ����Ч=0"
        Else
            strSql = strSql & " And A.ҽ����Ч=1"
        End If
        
        '�������
        Select Case UCase(Mid(lvwReport.SelectedItem.Key, 2))
        Case "ZL1_INSIDE_1254_4" '��ҩ��
            strSql = " And B.ִ�з���=4"
        Case "ZL1_INSIDE_1254_5" 'ע�䵥
            strSql = " And B.ִ�з���=2"
        Case "ZL1_INSIDE_1254_6" '��Һ��
            strSql = " And B.ִ�з���=1"
        Case Else
            '�����ı������Զ���ģ�����¼�ϴδ�ӡʱ�䣬��֧�ֱ����ظ���ӡ
            Exit Sub
        End Select
        
        '����ѡ��
        arrPati = Array()
        For i = 1 To vsPati.Rows - 1
            If Val(vsPati.TextMatrix(i, COL_ѡ��)) = -1 And vsPati.RowHidden(i) = False Then
                ReDim Preserve arrPati(UBound(arrPati) + 1)
                arrPati(UBound(arrPati)) = vsPati.RowData(i) & "," & vsPati.TextMatrix(i, COL_��ҳID)
            End If
        Next
        If mbln������˻�ҳ��ӡ Then
            strSql = strSql & " And A.����ID = [4] And A.��ҳID = [5]"
            varPar(0) = Split(mstr��ǰ����, ",")(0)
            varPar(1) = Split(mstr��ǰ����, ",")(1)
        Else
            If UBound(arrPati) = 0 Then '��������
                strSql = strSql & " And (A.����ID,A.��ҳID) IN((" & arrPati(0) & "))"
            Else
                For i = 0 To UBound(arrPati)
                    strPatis = strPatis & "," & Replace(arrPati(i), ",", ":")
                Next
                        
                strPatis = Mid(strPatis, 2)   'ȥ��ǰ��,��
                strTemp = "Select a.C1 As ����ID,a.C2 As ��ҳID From Table(f_Num2list2([1])) a"
                n = 0
                lngParStar = 3
                Do While True
                    If Len(strPatis) < 4000 Then
                        p = Len(strPatis) + 1
                    Else
                        p = InStrRev(Mid(strPatis, 1, 4000), ",")
                    End If
                    strThis = Mid(strPatis, 1, p - 1)
                    
                    If n > 10 Then
                        strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
                    Else
                        varPar(n) = strThis
                        strSQLPati = IIF(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (lngParStar + n + 1) & "]")
                    End If
                    
                    n = n + 1
                    strPatis = Mid(strPatis, p + 1)
                    If strPatis = "" Then Exit Do
                Loop
                
                strSql = strSql & " And (A.����ID,A.��ҳID) IN(" & strSQLPati & ")"
            End If
        End If
            
        '��ȡ���δ�ӡ��ҽ��
        strSql = _
            " Select /*+ Rule*/A.ID,zl_AdviceExeCount(A.Id,[1],[2]) As ����" & _
            " From ����ҽ����¼ A,������ĿĿ¼ B" & _
            " Where A.������ĿID=B.ID And A.�������='E' And B.��������='2'" & _
            " And A.У��ʱ�� Is Not Null And A.ҽ��״̬<>4" & strSql & _
            " And ([1]<=ִ����ֹʱ�� Or ִ����ֹʱ�� Is Null) And [2]>=��ʼִ��ʱ��"
        strSql = "Select * From (" & strSql & ") Where ����>0"
        
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, dtpBegin.value, dtpEnd.value, mlng����ID, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
        
        arrSQL = Array()
        Do While Not rsTmp.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����ҽ��ִ��_��ӡ(" & rsTmp!ID & "," & _
                "To_Date('" & Format(IIF(mintType = 1, dtpViewBegin.value, dtpBegin.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "To_Date('" & Format(IIF(mintType = 1, dtpViewEnd.value, dtpEnd.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            rsTmp.MoveNext
        Loop
    End If
    mlngLastRow = 0
    'ִ���ύ����
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zldatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optBaby_Click(Index As Integer)
    If fraBaby.Visible Then
        mlngҽ������Χ = Index
        Call cboUnit_Click
    End If
End Sub

Private Sub chkOut_Click()
'��ʾ��Ժ����
    Call cboUnit_Click
End Sub

Private Sub LoadAdvice()
'����: ����ҽ��
'����: �Ƿ������ִ��ҽ�� , Ϊ��Ϊ���ش����ҽ��
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long, j As Long
    Dim intBedLen As Integer
    Dim str����IDs As String
    Dim blnDo As String
    Dim bln��ҩ;�� As Boolean
    
    intBedLen = GetMaxBedLen(mlng����ID, False)
    str����IDs = ""
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, COL_ѡ��)) = -1 And vsPati.RowHidden(i) = False Then
            If InStr(str����IDs & ",", "," & vsPati.RowData(i) & ",") = 0 Then
                str����IDs = str����IDs & "," & vsPati.RowData(i)
            End If
        End If
    Next
    str����IDs = Mid(str����IDs, 2)
    If str����IDs = "" Then Exit Sub
    strSql = "Select /*+ Rule*/ b.Id, b.���id, e.Ӥ������,b.�������, b.����, Decode(b.ҽ����Ч, 0, '����', '����') As ��Ч, LPAD(c.��Ժ����," & intBedLen & ",' ') as ����," & vbNewLine & _
            "       Decode(b.���id, Null, b.ҽ������ || ' ' || b.ִ��Ƶ��, b.ҽ������) As ҽ������, b.ִ��Ƶ��, a.�ϴδ�ӡʱ��, b.ִ����ֹʱ��" & vbNewLine & _
            "From ҽ��ִ�д�ӡ A, ����ҽ����¼ B, ������ҳ C,������Ϣ D,������������¼ E,��Ժ���� R" & vbNewLine & _
            "Where a.ҽ��id = NVL(B.���ID,b.Id)  And e.����ID(+)=b.����ID and e.��ҳID(+)=b.��ҳID And e.���(+)=b.Ӥ�� And b.����id = c.����id And b.��ҳid = c.��ҳid And c.����id=d.����id and c.��ҳid=d.��ҳid" & _
            " And (R.����ID=[1] OR C.Ӥ������ID=[1]) and D.����ID=R.����ID And D.��ǰ����ID=R.����ID And zl_AdviceExeCount(b.Id,[2],[3],1)>0 " & _
            " And R.����ID In(Select Column_Value From Table(Cast(f_Str2List([4]) As zlTools.t_StrList)))  And A.����ID=[5] "
            

    strSql = strSql & " Order By ����,Nvl(b.���id, b.Id),b.���"
    

    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, CDate(dtpViewBegin.value), CDate(dtpViewEnd.value), str����IDs, Val(lvwReport.SelectedItem.Tag))

    With vsAdvice
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                .RowData(i) = rsTmp!ID & ""
                .TextMatrix(i, col����) = rsTmp!���� & ""
                .TextMatrix(i, col��Ч) = rsTmp!��Ч & ""
                .TextMatrix(i, COL��λ) = rsTmp!���� & ""
                .TextMatrix(i, COLӤ��) = rsTmp!Ӥ������ & ""
                .TextMatrix(i, COLƵ��) = rsTmp!ִ��Ƶ�� & ""
                .TextMatrix(i, COL�ϴδ�ӡʱ��) = Format(rsTmp!�ϴδ�ӡʱ�� & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLֹͣʱ��) = Format(rsTmp!ִ����ֹʱ�� & "", "yyyy-MM-dd HH:mm")
                If Format(rsTmp!ִ����ֹʱ�� & "", "yyyy-MM-dd HH:mm:ss") < Format(rsTmp!�ϴδ�ӡʱ�� & "", "yyyy-MM-dd HH:mm:ss") And rsTmp!ִ����ֹʱ�� & "" <> "" Then
                    'ֹͣ���ϴδ�ӡʱ��֮ǰ��ҽ������ɫ��ע
                    .Cell(flexcpBackColor, i, col����, i, COLֹͣʱ��) = &HE1FFE1
                End If
                .TextMatrix(i, col���) = rsTmp!������� & ""
                .TextMatrix(i, colҽ������) = rsTmp!ҽ������ & ""
                .TextMatrix(i, col���ID) = rsTmp!���ID & ""
                bln��ҩ;�� = False
                If IsNull(rsTmp!���ID) And rsTmp!������� & "" = "E" Then
                    If Val(.TextMatrix(i - 1, col���ID)) = .RowData(i) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, col���)) > 0 Then
                            bln��ҩ;�� = True
                        End If
                    End If
                End If
                '����һЩ������
                If (InStr(",F,G,D,7,E,C,", rsTmp!�������) > 0 And Not IsNull(rsTmp!���ID)) Or bln��ҩ;�� Then
                    .RemoveItem i
                    i = i - 1
                End If
                rsTmp.MoveNext
                i = i + 1
            Loop
        Else
            .AddItem ""
        End If
        vsAdvice.ColHidden(COLӤ��) = Not fraBaby.Visible
        '�Զ������и�
        .AutoSize colҽ������
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
        '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = col��Ч: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COLƵ��: lngRight = COLƵ��
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd, vsAdvice) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If .Cell(flexcpBackColor, Row, col����) = &HE1FFE1 Then
                SetBkColor hDC, OS.SysColor2RGB(&HE1FFE1)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If .TextMatrix(lngRow, col���) = "" Then Exit Function
        If .TextMatrix(lngRow, col���) = "���" Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 Or Val(.TextMatrix(lngRow - 1, col���ID)) = Val(.RowData(lngRow)) Or Val(.RowData(lngRow - 1)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.RowData(lngRow - 1)) <> 0 Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow + 1, col���ID)) <> 0 Or Val(.RowData(lngRow + 1)) = Val(.TextMatrix(lngRow, col���ID)) Or Val(.TextMatrix(lngRow + 1, col���ID)) = Val(.RowData(lngRow)) Then
                blnTmp = True
            End If
        End If
        lngBegin = lngRow
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 And Val(.RowData(i)) <> Val(.RowData(lngRow)) Or Val(.TextMatrix(i, col���ID)) = Val(.RowData(lngRow)) Or Val(.RowData(i)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.RowData(i)) <> 0 Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        lngEnd = lngRow
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, col���ID)) = Val(.TextMatrix(lngRow, col���ID)) And Val(.TextMatrix(lngRow, col���ID)) <> 0 And Val(.RowData(i)) <> Val(.RowData(lngRow)) Or Val(.RowData(i)) = Val(.TextMatrix(lngRow, col���ID)) Or Val(.TextMatrix(i, col���ID)) = Val(.RowData(lngRow)) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub vsPati_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> COL_ѡ�� Then Cancel = True
End Sub

Private Sub vsPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsPati.Row > 0 And KeyCode = vbKeySpace And vsPati.Col <> COL_ѡ�� Then
        vsPati.TextMatrix(vsPati.Row, COL_ѡ��) = IIF(Val(vsPati.TextMatrix(vsPati.Row, COL_ѡ��)) = -1, 0, -1)
    End If
End Sub

Private Sub SetBabyVisible()
'���ܣ�����Ӥ�����������Ŀɼ���
    Dim blnTmp As Boolean
    If mbln���� Then
        If mblnӤ������ Then
            blnTmp = True
        Else
            If ChkWaitPrint.value = 1 Then
                blnTmp = True
            End If
        End If
    End If
    fraBaby.Visible = blnTmp
End Sub
