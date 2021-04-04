VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBalanceEdit 
   Caption         =   "团体体检结算"
   ClientHeight    =   7305
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11880
   Icon            =   "frmBalanceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11880
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   6945
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBalanceEdit.frx":076A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   30
      ScaleHeight     =   555
      ScaleWidth      =   10650
      TabIndex        =   23
      Top             =   5565
      Width           =   10650
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   90
         TabIndex        =   18
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9315
         TabIndex        =   17
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8100
         TabIndex        =   16
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   645
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   10635
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   1
         Left            =   7095
         ScaleHeight     =   360
         ScaleWidth      =   3435
         TabIndex        =   26
         Top             =   180
         Width           =   3435
         Begin VB.TextBox txt 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Index           =   1
            Left            =   630
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   30
            Width           =   1005
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00E0E0E0&
            Height          =   300
            Index           =   2
            Left            =   2265
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   30
            Width           =   1140
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单据号"
            Height          =   180
            Index           =   3
            Left            =   1680
            TabIndex        =   30
            Top             =   90
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票据号"
            Height          =   180
            Index           =   4
            Left            =   30
            TabIndex        =   29
            Top             =   90
            Width           =   540
         End
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "团体体检结算"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   105
         TabIndex        =   21
         Top             =   255
         Width           =   1530
      End
   End
   Begin VB.Frame fra1 
      Height          =   585
      Left            =   0
      TabIndex        =   20
      Top             =   510
      Width           =   10110
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1350
         TabIndex        =   2
         Top             =   180
         Width           =   3150
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   0
         Left            =   4545
         Picture         =   "frmBalanceEdit.frx":0FFE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结帐团体(&N)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame fra2 
      Height          =   4410
      Left            =   -180
      TabIndex        =   22
      Top             =   1050
      Width           =   6690
      Begin MSComctlLib.TabStrip tbs 
         Height          =   300
         Left            =   45
         TabIndex        =   4
         Top             =   165
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   529
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.结 帐 表"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   3825
         Index           =   0
         Left            =   75
         TabIndex        =   5
         Top             =   495
         Width           =   6525
         _cx             =   11509
         _cy             =   6747
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         GridColorFixed  =   12698049
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
         Begin VB.Line lnX0 
            Index           =   0
            Visible         =   0   'False
            X1              =   -555
            X2              =   1230
            Y1              =   555
            Y2              =   555
         End
         Begin VB.Line lnY0 
            Index           =   0
            Visible         =   0   'False
            X1              =   270
            X2              =   270
            Y1              =   420
            Y2              =   1635
         End
      End
   End
   Begin VB.Frame fra3 
      Height          =   4410
      Left            =   6735
      TabIndex        =   24
      Top             =   1065
      Width           =   3780
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1065
         MaxLength       =   12
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   165
         Width           =   1170
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   1305
         Index           =   0
         Left            =   90
         ScaleHeight     =   1305
         ScaleWidth      =   3240
         TabIndex        =   25
         Top             =   3000
         Width           =   3240
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   90
            Width           =   1170
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   1005
            MaxLength       =   12
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   480
            Width           =   1170
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1005
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   885
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&3.应补金额"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   180
            Left            =   45
            TabIndex        =   10
            Top             =   150
            Width           =   900
         End
         Begin VB.Label lbl缴款 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&4.缴款金额"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   180
            Left            =   45
            TabIndex        =   12
            Top             =   555
            Width           =   900
         End
         Begin VB.Label lbl找补 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&5.找补金额"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   180
            Left            =   45
            TabIndex        =   14
            Top             =   960
            Width           =   900
         End
      End
      Begin zl9Medical.VsfGrid vsfPay 
         Height          =   1215
         Left            =   90
         TabIndex        =   9
         Top             =   900
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   2143
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&1.结帐金额"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   105
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&2.结算方式"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Top             =   645
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   10230
      Top             =   4035
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
            Picture         =   "frmBalanceEdit.frx":1588
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceEdit.frx":65F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceEdit.frx":68EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceEdit.frx":6E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceEdit.frx":7420
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBalanceEdit.frx":757A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBalanceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnModify As Boolean
Private mlng领用ID As Long
Private mstrItem As String '要结的收据费目
Private mstrALLItem As String '病人所有未结收据费目
Private mbytKind As Byte
Private mcurTotal As Currency
Private mblnZero As Boolean

Private Enum mCol
    结算方式 = 1
    金额
    结算号码
    缺省
    性质
    
    单据号 = 0
    费目 = 1
    项目
    未结金额
    结帐金额
    费用时间
    科室
    记录性质
    记录状态
    执行状态
    序号
End Enum

Private Type Items
    团体名称 As String
    ID As Long
    结帐金额 As String
End Type

Private mblnPrint As Boolean
Private usrSaveGroup As Items

'（２）自定义过程或函数************************************************************************************************
Private Function CheckBillRange(ByVal strFact As String, ByVal strItems As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：检查剩余的发票是否足够打印
    '参数：strFact=开始票据号
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim intPages As Integer
    Dim intRows As Integer
    
    On Error GoTo errHand
    
    '1.获取参数,门诊收据能够打印的项目总数
    intRows = Val(GetSysParameter(4))
            
    '2.检查票据是否够用
    If gblnBill结帐 And intRows > 0 Then
    
        '2.1.计算要打印的票据张数
'        strItems = IIf(mstrItem = "", mstrALLItem, mstrItem)
        intPages = IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / intRows)
        
        '2.2.检查每张票据是否可用
        For intLoop = 1 To intPages
            mlng领用ID = CheckUsedBill(mbytKind, IIf(mlng领用ID > 0, mlng领用ID, glng结帐ID), strFact)
            If mlng领用ID <= 0 Then
                
                Select Case mlng领用ID
                    Case 0 '操作失败
                    Case -1
                        ShowSimpleMsg "本次结帐要使用 " & intPages & " 张票据,你已没有足够的自用和共用的收费票据！" & vbCrLf & _
                            "请先领用一批票据或设置本地共用票据！"
                    Case -2
                        ShowSimpleMsg "本次结帐要使用 " & intPages & " 张票据,你已没有足够的本地共用票据！" & vbCrLf & _
                            "请先领用一批票据或重新设置本地共用票据！"
                    Case -3
                        ShowSimpleMsg "本次结帐要使用 " & intPages & " 张票据,但当前领用的票据剩余号码不足！"
                End Select
                
                Exit Function
                
            End If
            strFact = IncStr(strFact)
        Next
    End If
    
    CheckBillRange = True
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InheritAppendSpaceRows(ByVal intIndex As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '功能：补齐表格空行
    '------------------------------------------------------------------------------------------------------------------
    Select Case intIndex
    Case 0
        Call AppendRows(vsf(intIndex), lnX0, lnY0)
    End Select
End Sub

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
'
'    mnuFileSave.Enabled = vData
'
'    tbrThis.Buttons("结算").Enabled = mnuFileSave.Enabled
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next

    For lngLoop = 0 To txt.UBound
        txt(lngLoop).Text = ""
        txt(lngLoop).Tag = ""
    Next

    On Error GoTo 0
    
    Call ResetVsf(vsf(0))
    Call ResetVsf(vsfPay)
    
    Call AppendRows(vsf(0), lnX0, lnY0)
    
    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, Optional ByVal blnModify As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnPrint = True
    mblnOK = False
    mblnModify = blnModify
    mlngKey = lngKey
    
    Set mfrmMain = frmMain

    If InitData = False Then Exit Function
    If ReadData(mlngKey) = False Then Exit Function

    stbThis.Panels(2).Text = "只结算“记帐”的体检项目所产生的费用。"
    
    EditChanged = False

    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand

    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    mblnZero = (Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "对零费用进行结帐", 1)) = 1)
    mlng领用ID = 0
    mbytKind = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据类型", 3))
    
    strVsf = "单据号,900,1,1,1,;费目,0,1,1,0,;项目,2400,1,1,1,;未结金额,1080,7,1,1,;结帐金额,1080,7,1,1,;费用时间,0,1,1,0,;科室,1200,1,1,1,;记录性质,0,1,1,0,;记录状态,0,1,1,0,;执行状态,0,1,1,0,;序号,0,1,1,0,"
    Call CreateVsf(vsf(0), strVsf)
    vsf(0).Cols = vsf(0).Cols + 1
    vsf(0).ColWidth(vsf(0).Cols - 1) = 15
    Call AppendRows(vsf(0), lnX0, lnY0)
    
    With vsfPay
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "结算方式", 900, 1
        .NewColumn "金额", 1080, 7, , 1
        .NewColumn "结算号码", 1080, 1, , 1
        .NewColumn "缺省", 0, 1
        .NewColumn "性质", 0, 1
        .NewColumn "", 15, 1
        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .AppendRow = True
    End With
    
    strTmp = Trim(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "缺省结算方式", ""))
    
    gstrSQL = "SELECT A.结算方式,NULL AS 金额,NULL AS 结算号码,Decode(A.结算方式,'" & strTmp & "',1,0) AS 缺省,B.性质 " & _
                    "from 结算方式应用 A,结算方式 B where A.结算方式=B.名称 AND A.应用场合=[1] AND 性质 in (1,2)"
    gstrSQL = "Select * From (" & gstrSQL & ") Order By 缺省 Desc"
    
    If mbytKind = 1 Then
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "收费")
        
    Else
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "结帐")
    
    End If

    If rs.BOF = False Then
        Call FillGrid(vsfPay, rs)
        
        For lngLoop = 1 To vsfPay.Rows - 1
            If Val(vsfPay.TextMatrix(lngLoop, 4)) = 1 Then
                vsfPay.Cell(flexcpFontBold, lngLoop, 1, lngLoop, 1) = True
                Exit For
            End If
        Next
        
    End If
    
    gbytBalanceRows = 40
    gbytBalanceRows = Val(zlDatabase.GetPara(4, glngSys, , "40"))
    
    '共用票据ID
    glng结帐ID = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0)
    glngShareUseID = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0)
    
    '票据号
    txt(1).Text = RefreshFact(mbytKind)
    If txt(1).Text = "" And gblnStrictCtrl Then Exit Function
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function RefreshFact(ByVal bytKind As Byte) As String

    '功能：刷新收费票据号
    
'    If gint结帐打印 = 0 And gblnNotPrint Then Exit Sub
    If gblnStrictCtrl Then
        mlng领用ID = CheckUsedBill(bytKind, IIf(mlng领用ID > 0, mlng领用ID, glngShareUseID))
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    ShowSimpleMsg "你没有自用和共用的结算票据,请先领用一批票据或设置本地共用票据！"
                Case -2
                    ShowSimpleMsg "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！"
            End Select
            RefreshFact = ""
        Else
            '严格：取下一个号码
            RefreshFact = GetNextBill(mlng领用ID)
        End If
    Else
        '松散：取下一个号码
        RefreshFact = IncStr(UCase(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "当前结帐票据号", "")))
    End If
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim curMoney As Currency
    Dim strItems As String
    
    '1.检查是否还有余额未付
'    If Val(txt(3).Text) <= 0 Then
'        ShowSimpleMsg "没有要结算的团体体检费用！"
'        Exit Function
'    End If
    
    If Val(txt(6).Text) <> 0 Then
        ShowSimpleMsg "必须全部结算，所有结算方式的结算金额和不等于" & Val(txt(3).Text) & "！"
        Exit Function
    End If
    
    '2.票据号码有效性检查
    If mblnPrint Then
        If gblnStrictCtrl Then   '严格票据管理
            If Trim(txt(1).Text) = "" Then
                ShowSimpleMsg "必须输入一个有效的票据号码！"
                Call LocationObj(txt(1))
                Exit Function
            End If
            mlng领用ID = GetInvoiceGroupID(mbytKind, 1, mlng领用ID, glngShareUseID, txt(1).Text)
            If mlng领用ID <= 0 Then
                Select Case mlng领用ID
                    Case 0 '操作失败
                        
                    Case -1
                        ShowSimpleMsg "你没有自用和共用的结帐票据，请先领用一批票据或设置本地共用票据！"
                    Case -2
                        ShowSimpleMsg "本地的共用票据已经用完，请先领用一批票据或重新设置本地共用票据！"
                    Case -3
                        ShowSimpleMsg "当前票据号码不在可用领用批次的有效票据号范围内，请重新输入！"
                        Call LocationObj(txt(1))
                End Select
                Exit Function
            End If
        Else
            If Len(txt(1).Text) <> ParamInfo.结帐票据号长度 And txt(1).Text <> "" Then
                ShowSimpleMsg "票据号码长度应该为 " & ParamInfo.结帐票据号长度 & " 位！"
                Call LocationObj(txt(1))
                Exit Function
            End If
        End If
    End If
        
'    If gblnBill结帐 Then
'
'        '2.1.检查是否有票据号码
'        If Trim(txt(1).Text) = "" Then
'            ShowSimpleMsg "必须输入一个有效的票据号码！"
'            LocationObj txt(1)
'            Exit Function
'        End If
'
'        '2.2.检查是否有领用或共用票据
'        mlng领用ID = CheckUsedBill(mbytKind, IIf(mlng领用ID > 0, mlng领用ID, glng结帐ID), txt(1).Text)
'        If mlng领用ID <= 0 Then
'            Select Case mlng领用ID
'                Case 0 '操作失败
'                Case -1
'                    ShowSimpleMsg "你没有自用和共用的结帐票据,请先领用一批票据或设置本地共用票据！"
'                Case -2
'                    ShowSimpleMsg "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！"
'                Case -3
'                    ShowSimpleMsg "票据号码不在当前有效领用范围内,请重新输入！"
'                    LocationObj txt(1)
'            End Select
'            Exit Function
'        End If
'
'        '2.3.检查号码是否够用
'        '2.3.1.搜索要打印的费目种类
'        For lngLoop = 1 To vsf(0).Rows - 1
'            If InStr(strItems & ",", "," & vsf(0).TextMatrix(lngLoop, mCol.费目) & ",") = 0 Then
'                strItems = strItems & "," & vsf(0).TextMatrix(lngLoop, mCol.费目)
'            End If
'        Next
'        If strItems <> "" Then strItems = Mid(strItems, 2)
'
'        '2.3.2.检查票据是否够用
'        If strItems <> "" Then
'            If Not CheckBillRange(txt(1).Text, strItems) Then Exit Function
'        End If
'    Else
'        '2.4.不严格控制票据情况下，输入了票据号的处理
'        If Len(txt(1).Text) <> ParamInfo.结帐票据号长度 And txt(1).Text <> "" Then
'            ShowSimpleMsg "票据号码长度应该为 " & ParamInfo.结帐票据号长度 & " 位！"
'            Call LocationObj(txt(1))
'            Exit Function
'        End If
'    End If
    
    ValidEdit = True

End Function

Private Function SaveEdit(ByRef lngSaveID As Long, ByRef curDate As Date) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim lng结帐ID As Long
    Dim strNo As String
    Dim strNow As String
    
    On Error GoTo errHand

    ReDim Preserve strSQL(1 To 1)
    
    strNo = GetNextNo(15)
    txt(2).Text = strNo
    
    lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
    curDate = zlDatabase.Currentdate
    strNow = Format(curDate, "yyyy-MM-dd HH:mm:ss")
    
    lngSaveID = lng结帐ID
    
    strSQL(ReDimArray(strSQL)) = "zl_病人结帐记录_Insert(" & lng结帐ID & ",'" & _
                                                    strNo & "'," & _
                                                    "NULL," & _
                                                    "NULL," & _
                                                    "NULL," & _
                                                    "To_Date('" & strNow & "','YYYY-MM-DD HH24:MI:SS')," & _
                                                    "NULL," & _
                                                    "NULL)"
   
    '病人预交记录-结帐补：结算方式,金额,结算号码
    For lngLoop = 1 To vsfPay.Rows - 1
        If Val(vsfPay.TextMatrix(lngLoop, mCol.金额)) <> 0 Then
            strSQL(ReDimArray(strSQL)) = "zl_结帐缴款记录_Insert('" & strNo & "'," & _
                                                                "NULL," & _
                                                                "0," & _
                                                                mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex) & ",'" & _
                                                                vsfPay.TextMatrix(lngLoop, mCol.结算方式) & "','" & _
                                                                vsfPay.TextMatrix(lngLoop, mCol.结算号码) & "'," & _
                                                                CCur(vsfPay.TextMatrix(lngLoop, mCol.金额)) & "," & _
                                                                lng结帐ID & ",'" & _
                                                                UserInfo.编号 & "','" & _
                                                                UserInfo.姓名 & "'," & _
                                                                "To_Date('" & strNow & "','YYYY-MM-DD HH24:MI:SS')," & _
                                                                "NULL," & _
                                                                "NULL," & _
                                                                "NULL)"
        End If
    Next

    For lngLoop = 1 To vsf(0).Rows - 1

        If Val(vsf(0).TextMatrix(lngLoop, mCol.结帐金额)) <> 0 Or Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额)) = 0 Then
'            strSQL(ReDimArray(strSQL)) = "zl_结帐费用记录_Insert(" & Val(vsf(0).RowData(lngLoop)) & ",'" & _
'                                                                vsf(0).TextMatrix(lngLoop, mCol.单据号) & "'," & _
'                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.记录性质)) & "," & _
'                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.记录状态)) & "," & _
'                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.执行状态)) & "," & _
'                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.序号)) & "," & _
'                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.结帐金额)) & "," & _
'                                                                lng结帐ID & ",To_Date('" & strNow & "','yyyy-mm-dd'))"
            
            strSQL(ReDimArray(strSQL)) = "zl_结帐费用记录_Insert(" & Val(vsf(0).RowData(lngLoop)) & ",'" & _
                                                                vsf(0).TextMatrix(lngLoop, mCol.单据号) & "'," & _
                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.记录性质)) & "," & _
                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.记录状态)) & "," & _
                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.执行状态)) & "," & _
                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.序号)) & "," & _
                                                                Val(vsf(0).TextMatrix(lngLoop, mCol.结帐金额)) & "," & _
                                                                lng结帐ID & ")"
        End If
    Next
    
    Dim lngTmp As Long
    
    lngTmp = zlDatabase.GetNextId("体检结算记录")
    strSQL(ReDimArray(strSQL)) = "ZL_体检结算记录_INSERT(" & lngTmp & "," & mlngKey & "," & lng结帐ID & "," & Val(txt(3).Text) & "," & mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex) & ")"
    
'    '产生结帐票据
'    If Trim(txt(1).Text) <> "" Then
'        strSQL(ReDimArray(strSQL)) = "zl_病人结帐票据_Insert('" & strNO & "','" & _
'                                                            Trim(txt(1).Text) & "'," & _
'                                                            IIf(mlng领用ID = 0, "NULL", mlng领用ID) & ",'" & _
'                                                            UserInfo.姓名 & "'," & _
'                                                            "To_Date('" & strNow & "','YYYY-MM-DD HH24:MI:SS'))"
'    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True

    Exit Function

errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans

End Function

Private Function RefreshData(ByVal strMenu As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim dblMoney As Double
    
    Dim strKey As String
    
    On Error GoTo errHand
    
    Select Case strMenu
    Case "未结体检"
    
        
    Case "明细"
        
        gstrSQL = GetPublicSQL(SQL.团体未结明细)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            Call LoadGrid(vsf(0), rs, Array("", "", "", "0.00##", "0.00##"))
            Call InheritAppendSpaceRows(0)
        End If
        
        dblMoney = 0
        For lngLoop = 1 To vsf(0).Rows - 1
            dblMoney = dblMoney + Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额))
        Next
        
        txt(3).Text = Format(dblMoney, "0.00")
        mcurTotal = Val(txt(3).Text)
        usrSaveGroup.结帐金额 = txt(3).Text
        txt(3).Tag = ""
        
        Call AssignMoney(Val(txt(3).Text))
        
    End Select
    
    RefreshData = True
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    gstrSQL = GetPublicSQL(SQL.体检团体选择)
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt(0), "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1", Me.Name & "\体检团体选择", "请在下表中选择一个团体/单位。", rsData, rs, 8790, 5100) Then
        
        txt(0).Text = zlCommFun.NVL(rs("名称").Value)
        mlngKey = zlCommFun.NVL(rs("ID").Value, 0)
        
        usrSaveGroup.团体名称 = txt(0).Text
        
        '清除
        Call ResetVsf(vsf(0))
        vsfPay.Cell(flexcpText, 1, 2, vsfPay.Rows - 1, 2) = ""
        Call AppendRows(vsf(0), lnX0, lnY0)
        
        DoEvents
        
        Call RefreshData("明细")
        
    End If
    txt(0).SetFocus
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub cmdOK_Click()
    Dim lngKey As Long
    Dim curDate As Date
    
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit(lngKey, curDate) Then
        mblnOK = True
        
        '票据打印
        If mblnPrint Then
            Call frmPrint.ReportPrint(1, txt(2).Text, lngKey, mlng领用ID, txt(1).Text, curDate, txt(4).Text, txt(5).Text, mbytKind)
        End If
        
        stbThis.Panels(2).Text = "上一张单据号:" & txt(2).Text

        Call ClearData
        
        txt(1).Text = RefreshFact(mbytKind)
        
        EditChanged = False
        Unload Me
    End If
End Sub


Private Sub Form_Load()
    glngFormW = 12000
    glngFormH = 7710
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With fraTitle
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With pic(1)
        .Left = fraTitle.Width - .Width - 45
    End With
    
    With fra1
        .Left = fraTitle.Left
        .Top = fraTitle.Top + fraTitle.Height - 90
        .Width = fraTitle.Width
    End With
    
    With fra2
        .Left = fra1.Left
        .Top = fra1.Top + fra1.Height - 90
        .Width = fra1.Width - fra3.Width - 15
        .Height = Me.ScaleHeight - .Top - picButton.Height - stbThis.Height
    End With
                               
    With tbs
        .Left = 30
        .Top = 45 + 90
    End With
    
    With vsf(0)
        .Left = 45
        .Top = tbs.Top + tbs.Height + 45
        .Width = fra2.Width - .Left - 45
        .Height = fra2.Height - .Top - 45
    End With
    
    With fra3
        .Left = fra2.Left + fra2.Width + 15
        .Top = fra2.Top
        .Height = fra2.Height
    End With
                                
    txt(3).Width = fra3.Width - txt(3).Left - 75
        
    With vsfPay
        .Top = lbl(2).Top + lbl(2).Height + 30
        .Width = fra3.Width - .Left - 60
        .Height = fra3.Height - .Top - pic(0).Height - 30
    End With
    
    With pic(0)
        .Top = vsfPay.Top + vsfPay.Height
        .Width = vsfPay.Width
    End With
    
    txt(4).Width = pic(0).Width - txt(4).Left - 45
    txt(5).Width = pic(0).Width - txt(5).Left - 45
    txt(6).Width = pic(0).Width - txt(6).Left - 45
    
    With picButton
        .Left = fra1.Left
        .Top = fra2.Top + fra2.Height
        .Width = fra1.Width
    End With
    
    cmdCancel.Left = picButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    
    Call AppendRows(vsf(0), lnX0, lnY0)
    vsfPay.AppendRow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub AssignMoney(ByVal dbl金额 As Double)
    Dim lngLoop As Long
    Dim dbl合计 As Double
    Dim lngDefault As Long
    
    On Error GoTo errHand
        
    For lngLoop = 1 To vsfPay.Rows - 1
    
        If Val(vsfPay.TextMatrix(lngLoop, 4)) = 1 Then lngDefault = lngLoop
        dbl合计 = dbl合计 + Val(vsfPay.TextMatrix(lngLoop, 2))
        
    Next
    
    dbl金额 = dbl金额 - dbl合计
    
    If dbl金额 > 0 Then
        
        If lngDefault = 0 Then lngDefault = 1
        
        vsfPay.TextMatrix(lngDefault, 2) = Format(Val(vsfPay.TextMatrix(lngDefault, 2)) + dbl金额, "0.00")
        
    Else
        If lngDefault > 0 Then
        
            If Val(vsfPay.TextMatrix(lngDefault, 2)) + dbl金额 >= 0 Then
                
                vsfPay.TextMatrix(lngDefault, 2) = Format(Val(vsfPay.TextMatrix(lngDefault, 2)) + dbl金额, "0.00")
                dbl金额 = 0
                
            Else
                                
                dbl金额 = Val(vsfPay.TextMatrix(lngDefault, 2)) + dbl金额
                vsfPay.TextMatrix(lngDefault, 2) = "0.00"
                
            End If
            
        End If
        
        If dbl金额 <> 0 Then
            For lngLoop = 1 To vsfPay.Rows - 1
                If Val(vsfPay.TextMatrix(lngLoop, 2)) + dbl金额 >= 0 Then
                    vsfPay.TextMatrix(lngLoop, 2) = Format(Val(vsfPay.TextMatrix(lngLoop, 2)) + dbl金额, "0.00")
                    dbl金额 = 0
                Else
                    dbl金额 = Val(vsfPay.TextMatrix(lngLoop, 2)) + dbl金额
                    vsfPay.TextMatrix(lngLoop, 2) = "0.00"
                    
                End If
                
                If dbl金额 = 0 Then Exit For
            Next
        End If
    End If
    
    dbl合计 = 0
    For lngLoop = 1 To vsfPay.Rows - 1
        dbl合计 = dbl合计 + Val(vsfPay.TextMatrix(lngLoop, 2))
    Next
    
    txt(6).Text = Format(Val(txt(3).Text) - dbl合计, "0.00")
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
    '
End Sub

Private Sub tbs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    EditChanged = True
    
    Select Case Index
    Case 0
        txt(Index).Tag = "Changed"
        mlngKey = 0
    Case 3
        txt(Index).Tag = "Changed"
    End Select
    
'    If Index = 0 Then
'        txt(Index).Tag = "Changed"
'        mlngKey = 0
'    End If
'
    If Index = 4 Or Index = 3 Then
        txt(5).Text = Format(Val(txt(4).Text) - Val(txt(3).Text), "0.00")
    End If
    
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 0
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngLoop As Long
'    Dim dbl结帐金额 As Double
    Dim curMoney As Currency
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
            If txt(Index).Tag = "Changed" Then
                gstrSQL = GetPublicSQL(SQL.团体过滤选择)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
                If ShowTxtFilter(Me, txt(Index), "名称,1800,0,0;编码,900,0,0;简码,900,0,0;联系人,900,0,0;电话,1200,0,0", Me.Name & "\团体过滤选择", "请从下面选择一个团体单位", rsData, rs) Then
                    
                    txt(0).Text = zlCommFun.NVL(rs("名称"))
                    mlngKey = zlCommFun.NVL(rs("ID"))
                    txt(0).Tag = ""
                    
                    usrSaveGroup.团体名称 = txt(0).Text
                    
                    '清除
                    Call ResetVsf(vsf(0))
                    vsfPay.Cell(flexcpText, 1, 2, vsfPay.Rows - 1, 2) = ""
                    Call AppendRows(vsf(0), lnX0, lnY0)
        
                    Call RefreshData("未结体检")
                    Call RefreshData("明细")
                    
                Else
                    txt(0).Text = usrSaveGroup.团体名称
                    Exit Sub
                End If
            End If
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
        '--------------------------------------------------------------------------------------------------------------
        Case 3
            If Not IsNumeric(txt(3).Text) Then
                stbThis.Panels(2) = "输入错误！": Beep
                txt(3).Text = Format(mcurTotal, "0.00")
                LocationObj txt(3)
            ElseIf Val(txt(3).Text) <> 0 And Val(txt(3).Text) > mcurTotal Then
                stbThis.Panels(2) = "输入金额不能大于本次结帐的金额:" & Format(mcurTotal, "0.00"): Beep
                txt(3).Text = Format(mcurTotal, "0.00")
                LocationObj txt(3)
            Else
                '自动处理合计分配
                stbThis.Panels(2) = ""
                curMoney = Format(txt(3).Text, "0.00")
                For lngLoop = vsf(0).Rows - 1 To 1 Step -1
                    If curMoney = 0 Then
                        vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = "0.00"
                    Else
                        If Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额)) >= curMoney Then
                            vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = Format(curMoney, "0.00")
                        ElseIf Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额)) < curMoney Then
                            vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = Format(vsf(0).TextMatrix(lngLoop, mCol.未结金额), "0.00")
                        End If
                        curMoney = curMoney - Val(vsf(0).TextMatrix(lngLoop, mCol.结帐金额))
                    End If
                Next
                If curMoney <> 0 Then
                    vsf(0).TextMatrix(1, mCol.结帐金额) = Format(Val(vsf(0).TextMatrix(1, mCol.结帐金额)) + curMoney, "0.00")
                End If
                
                Call AssignMoney(Val(txt(3).Text))

            End If
            txt(3).Text = Format(txt(3).Text, "0.00")
            txt(3).Tag = ""
            usrSaveGroup.结帐金额 = txt(3).Text
            
            vsfPay.Col = 2
            vsfPay.SetFocus
            
        End Select
        
'        If Val(txt(3).Text) >= 0 And Index = 3 Then
'            If Val(txt(3).Text) < 0 Or Val(txt(3).Text) > Val(vsf(0).Tag) Then
'                txt(3).Text = txt(3).Tag
'            Else
'                If Val(txt(3).Tag) <> Val(txt(3).Text) Then
'
'                    txt(3).Tag = txt(3).Text
'                    Call AssignMoney(Val(txt(3).Text))
'
'                    '调整结帐金额
''                    dbl结帐金额 = Val(vsf(0).Tag) - Val(txt(3).Text)
'                    dbl结帐金额 = Val(txt(3).Text)
'
'                    For lngLoop = 1 To vsf(0).Rows - 1
'
'                        If dbl结帐金额 <> 0 Then
'                            If Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额)) >= dbl结帐金额 Then
'                                vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = Format(dbl结帐金额, "0.00")
'                                dbl结帐金额 = 0
'                            Else
'                                vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = vsf(0).TextMatrix(lngLoop, mCol.未结金额)
'                                dbl结帐金额 = dbl结帐金额 - Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额))
'                            End If
'                        Else
'                            vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = "0.00"
'                        End If
'
''                        If Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额)) <= dbl结帐金额 Then
''                            vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = "0.00"
''                            dbl结帐金额 = dbl结帐金额 - Val(vsf(0).TextMatrix(lngLoop, mCol.未结金额))
''                        Else
''                            vsf(0).TextMatrix(lngLoop, mCol.结帐金额) = Format(Val(vsf(0).TextMatrix(lngLoop, mCol.结帐金额)) - dbl结帐金额, "0.00")
''                            dbl结帐金额 = 0
''                        End If
'
''                        If dbl结帐金额 = 0 Then Exit For
'                    Next
'
'                End If
'            End If
            
'        End If
        
'        If Index = 3 Then
'            vsfPay.Col = 2
'            vsfPay.SetFocus
'        Else
'            zlCommFun.PressKey vbKeyTab
'        End If
'
'        If Index = 0 Then zlCommFun.PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        zlCommFun.OpenIme False
    Case 3
        txt(3).Text = Format(txt(3).Text, "0.00")
        
'        If Not IsNumeric(txt(3).Text) Then
'            txt(3).SetFocus
'        ElseIf mcurTotal <> CCur(txt(3).Text) Then
'            txt(3).Text = Format(mcurTotal, "0.00")
'        End If
    
    End Select
    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    Select Case Index
    Case 0
        If txt(Index).Tag = "Changed" Then
            txt(Index).Text = usrSaveGroup.团体名称
            txt(Index).Tag = ""
        End If
    Case 3
        If txt(Index).Tag = "Changed" Then
            txt(Index).Text = usrSaveGroup.结帐金额
            txt(Index).Tag = ""
        End If
    End Select
End Sub


Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsfPay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    Dim dbl合计 As Double
    
    If Col = 2 Then
        vsfPay.TextMatrix(Row, Col) = Format(vsfPay.TextMatrix(Row, Col), "0.00")
    End If
    
    For lngLoop = 1 To vsfPay.Rows - 1
        dbl合计 = dbl合计 + Val(vsfPay.TextMatrix(lngLoop, 2))
    Next
    
    txt(6).Text = Format(Val(txt(3).Text) - dbl合计, "0.00")
End Sub

Private Sub vsfPay_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsfPay_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsfPay_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn Then
        If Row = vsfPay.Rows - 1 And Col = 3 Then LocationObj txt(4)
    End If
End Sub
