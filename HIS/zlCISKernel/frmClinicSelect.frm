VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmClinicSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "诊疗项目选择器"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   Icon            =   "frmClinicSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9120
   Begin VB.CheckBox chkShowCause 
      Caption         =   "显示未匹配成功的项目"
      Height          =   195
      Left            =   4080
      TabIndex        =   23
      Top             =   5658
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame fraInfo 
      Height          =   435
      Left            =   30
      TabIndex        =   14
      Top             =   -75
      Width           =   9075
      Begin VB.CheckBox chkSub 
         Caption         =   "包含下级项目(&T)"
         Height          =   195
         Left            =   7380
         TabIndex        =   10
         Top             =   165
         Width           =   1650
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   180
         Left            =   225
         TabIndex        =   15
         Top             =   165
         Width           =   270
      End
   End
   Begin VB.Frame fraStat 
      Height          =   5160
      Left            =   15
      TabIndex        =   17
      Top             =   285
      Visible         =   0   'False
      Width           =   2715
      Begin VB.CommandButton cmdSelClear 
         Caption         =   "全清(&R)"
         Height          =   350
         Left            =   1425
         TabIndex        =   7
         ToolTipText     =   "Ctrl+R"
         Top             =   3705
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "全选(&A)"
         Height          =   350
         Left            =   1425
         TabIndex        =   6
         ToolTipText     =   "Ctrl+A"
         Top             =   3345
         Width           =   1100
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "全部显示"
         Height          =   195
         Left            =   1050
         TabIndex        =   4
         Top             =   1965
         Width           =   1020
      End
      Begin VB.CommandButton cmdStat 
         Caption         =   "统计(&S)"
         Height          =   350
         Left            =   1425
         TabIndex        =   5
         Top             =   2835
         Width           =   1100
      End
      Begin VB.TextBox txtCount 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "100"
         Top             =   1620
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   106233859
         CurrentDate     =   38434
      End
      Begin VB.ComboBox cboDate 
         Height          =   300
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmClinicSelect.frx":058A
         Left            =   1050
         List            =   "frmClinicSelect.frx":05A6
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   825
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "如果统计的时间范围较长，速度可能会较慢，请耐心等待。"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   165
         TabIndex        =   22
         Top             =   2340
         Width           =   2400
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "条"
         Height          =   180
         Left            =   2100
         TabIndex        =   21
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "显示最前"
         Height          =   180
         Left            =   285
         TabIndex        =   20
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "统计时间"
         Height          =   180
         Left            =   285
         TabIndex        =   19
         Top             =   885
         Width           =   720
      End
      Begin VB.Label lblStatTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "自动统计""XXXXXX""最近常用的诊疗项目："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   18
         Top             =   270
         Width           =   2400
      End
   End
   Begin MSComctlLib.ImageList imgOften 
      Left            =   1110
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":060A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":0D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":1AF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOften 
      Height          =   450
      Left            =   495
      TabIndex        =   16
      Top             =   5505
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   794
      ButtonWidth     =   1561
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgOften"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "常用"
            Key             =   "Often"
            Description     =   "常用"
            Object.ToolTipText     =   "显示常用项目(F2)"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "统计"
            Key             =   "Stat"
            Description     =   "统计"
            Object.ToolTipText     =   "统计常用项目"
            Object.Tag             =   "统计"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "加入"
            Key             =   "New"
            Description     =   "加入"
            Object.ToolTipText     =   "加入常用项目(F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "移除"
            Key             =   "Del"
            Description     =   "移除"
            Object.ToolTipText     =   "移除常用项目(Del)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   4665
      Left            =   2790
      TabIndex        =   8
      Top             =   375
      Width           =   6300
      _cx             =   11112
      _cy             =   8229
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicSelect.frx":21F2
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
      ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList imgSort 
         Left            =   930
         Top             =   900
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":227F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":2759
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   13
      Top             =   555
      Width           =   45
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7335
      TabIndex        =   12
      Top             =   5580
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6240
      TabIndex        =   11
      Top             =   5580
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   600
      Left            =   2835
      TabIndex        =   9
      Top             =   4815
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   1058
      TabWidthStyle   =   2
      TabFixedWidth   =   1623
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      Placement       =   1
      ImageList       =   "img16"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "全部(0)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "西成药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中成药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中草药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "治疗"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2235
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
            Picture         =   "frmClinicSelect.frx":2C33
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":31CD
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":3767
            Key             =   "成药"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":3D01
            Key             =   "诊疗"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":429B
            Key             =   "草药"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":4835
            Key             =   "方案"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   4995
      Left            =   15
      TabIndex        =   0
      Top             =   390
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   8811
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "可用药房库存"
      Height          =   180
      Left            =   2760
      TabIndex        =   24
      Top             =   5595
      Width           =   1080
   End
   Begin VB.Shape Shp 
      Height          =   405
      Left            =   4800
      Top             =   5550
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   10000
      Y1              =   5445
      Y2              =   5445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "frmClinicSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint场合 As Integer
Private mint范围 As Integer
Private mlng病人性质 As Long
Private mlng病人科室id As Long
Private mblnOK As Boolean
Private mrsItem As ADODB.Recordset
Private mint期效 As Integer
Private mstr性别 As String
Private mstr输入 As String
Private mobjTXT As Object
Private mlng分类ID As Long
Private mlng定位分类ID As Long
Private mint险类 As Integer
Private mintType As Integer  '当mlng药名ID>0时 该变量的值有效: 0-读取指定品种药品的所有规格;1-读取指定品种卫材的所有规格

Private mstr诊疗分类 As String
Private mstr操作类型 As String
Private mstr执行分类 As String

Private mstrSaveTag As String
Private mstrPreNode As String
Private mblnClick As Boolean

Private mbln价格 As Boolean
Private mbln简码 As Boolean
Private mint简码 As Integer
Private mstrLike As String

Private mstr可用西药房 As String
Private mstr可用成药房 As String
Private mstr可用中药房 As String
Private mstr发料部门 As String

Private mlng西药房 As Long
Private mlng成药房 As Long
Private mlng中药房 As Long
Private mlng发料部门 As Long
Private mlng药名ID As Long '读取指定品种药品的所有规格;读取指定品种卫材的所有规格

Private mstr科室ID As String '诊疗适用的科室ID
Private mstrPrivs As String
Private mbyt匹配 As Byte '不匹配的原因，1表示期效不匹配
Private mbytSize As Byte
Private mbln显示库存 As Boolean
Private mstr药品价格等级 As String '病人的药品价格等级
Private mstr卫材价格等级 As String '病人的卫材价格等级
Private mstr普通项目价格等级 As String '病人的普通项目价格等级
Private mlng医技科室ID As Long

Public Function ShowSelect(frmParent As Object, ByVal int场合 As Integer, ByVal lng病区ID As Long, ByVal lng科室id As Long, _
    ByVal int期效 As Integer, ByVal str性别 As String, Optional ByVal str输入 As String, _
    Optional objTxt As Object, Optional ByVal int范围 As Integer = 2, _
    Optional ByVal lng分类ID As Long, Optional ByVal int险类 As Integer, Optional ByVal lng病人性质 As Integer, Optional ByVal lng药名ID As Long, _
    Optional ByVal str使用科室 As String, Optional ByRef byt匹配 As Byte, Optional str诊疗分类 As String, _
    Optional ByVal str操作类型 As String, Optional ByVal str执行分类 As String, Optional ByVal lng定位分类ID As Long, _
    Optional ByVal str药品价格等级 As String, Optional ByVal str卫材价格等级 As String, Optional ByVal str普通项目价格等级 As String, _
    Optional ByVal intType As Integer, Optional ByVal lng医技科室ID As Long) As ADODB.Recordset
'功能：显示诊疗项目选择器
'参数：int场合=(-1)-成套方案编辑,0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      lng病区ID/lng科室ID=病人的病区/科室ID
'      int期效=医嘱期效
'      str性别=病人性别
'      str输入=输入匹配的内容,如果没有则为选择器方式,否则为列表方式
'      objTXT=用于列表定位的输入框
'      blnCancel(O):是否取消
'      int范围=1-门诊,2-住院,3-门诊和住院
'      lng分类ID=选择器时(str输入="")，从这个分类开始显示
'      lng药名ID=读取指定品种药品\卫材的所有规格，传入药品\卫材的诊疗项目ID
'      byt匹配 出参，不匹配原因 =1期效 =2简码
'      str诊疗分类 =（-1）成套方案编辑且路径项目批量调整时传人
'      str操作类型=（-1）成套方案编辑且路径项目批量调整时传人
'      str执行分类=（-1）成套方案编辑且路径项目批量调整时传人
'      lng定位分类ID=定位到该分类下面
'      intType-该参数用于区分【lng药名ID】是药品的诊疗项目ID还是卫材的诊疗项目ID。=0：【lng药名ID】为药品的诊疗项目ID;=1 【lng药名ID】为卫材的诊疗项目ID(临床路径生成时)
'      lng医技科室ID 医技工作站调用时当前界面科室ID
'返回：如果没有数据,或取消,则返回Nothing；否则为一条包含诊疗项目数据的记录
    mint场合 = int场合
    mint范围 = int范围
    mint期效 = int期效
    mstr性别 = str性别
    mstr输入 = str输入
    mlng药名ID = lng药名ID
    mlng病人科室id = lng科室id
    mlng医技科室ID = lng医技科室ID
    If mlng药名ID <> 0 Then mstr输入 = ""
    
    Set mobjTXT = objTxt
    mlng分类ID = lng分类ID
    mlng定位分类ID = lng定位分类ID
    mint险类 = int险类
    mlng病人性质 = lng病人性质
    mstr药品价格等级 = str药品价格等级
    mstr卫材价格等级 = str卫材价格等级
    mstr普通项目价格等级 = str普通项目价格等级
    
    mstrSaveTag = mint范围 & IIF(mstr输入 <> "", 1, 0) & IIF(gbln药品按规格下医嘱 Or mint期效 = 1, 1, 0)
    
    '诊疗适用科室
    If mint场合 = -1 Then
        '成套方案编辑：接口不传入，取操作员所属所有科室
        If str使用科室 = "" Then
            mstr科室ID = GetUser科室IDs
            If mstr科室ID <> "" Then mstr科室ID = "," & mstr科室ID & ","
        Else
            mstr科室ID = "," & str使用科室 & ","
        End If
    Else
        '医技工作站：根据病人科室，限制诊疗项目的使用科室
        '医生工作站：
        '    住院：根据病人科室，限制诊疗项目的使用科室
        '    门诊：根据病人科室，限制诊疗项目的使用科室
        mstr科室ID = "," & lng科室id & ","
    End If
    mbyt匹配 = 0
    mstr诊疗分类 = str诊疗分类
    mstr操作类型 = str操作类型
    mstr执行分类 = str执行分类
    mintType = intType
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        byt匹配 = mbyt匹配
        Set ShowSelect = mrsItem
    Else
        byt匹配 = 0
        Set ShowSelect = Nothing
    End If
End Function

Private Sub cboDate_Click()
    Dim curDate As Date
    
    If cboDate.ListIndex = cboDate.ListCount - 1 Then
        dtpDate.Enabled = True
        dtpDate.SetFocus
    Else
        dtpDate.Enabled = False
        curDate = zlDatabase.Currentdate
        Select Case cboDate.ListIndex
            Case 0 '一周
                dtpDate.value = DateAdd("ww", -1, curDate)
            Case 1 '半月(15天)
                dtpDate.value = DateAdd("d", -15, curDate)
            Case 2 '一月
                dtpDate.value = DateAdd("m", -1, curDate)
            Case 3 '二月
                dtpDate.value = DateAdd("m", -2, curDate)
            Case 4 '三月
                dtpDate.value = DateAdd("m", -3, curDate)
            Case 5 '半年
                dtpDate.value = DateAdd("m", -6, curDate)
            Case 6 '一年
                dtpDate.value = DateAdd("yyyy", -1, curDate)
        End Select
    End If
End Sub

Private Sub chkAll_Click()
    txtCount.Enabled = chkAll.value = 0
End Sub

Private Sub chkShowCause_Click()
    If chkShowCause.value = 1 Then tbrOften.Buttons(1).value = tbrUnpressed
    Call FillList
    Call SetFormSize
    If chkShowCause.value = 1 Then
        cmdOK.Enabled = False
        If mrsItem.RecordCount = 1 Then
            If InStr(mrsItem!未匹配原因, "项目简码不匹配") > 0 Or InStr(mrsItem!未匹配原因, "诊疗项目的执行频率为") > 0 And mint范围 <> 1 Then
                cmdOK.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkSub_Click()
    If Not Visible Then Exit Sub
    vsItem.SetFocus
    Call FillList(True)
End Sub

Private Sub cmdCancel_Click()
    Set mrsItem = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If chkShowCause.value = 1 Then
        If mrsItem.RecordCount = 1 Then
            If InStr(mrsItem!未匹配原因, "项目简码不匹配") > 0 Then
                mbyt匹配 = 2
            ElseIf InStr(mrsItem!未匹配原因, "诊疗项目的执行频率为") > 0 Then
                mbyt匹配 = 1
            End If
        End If
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdStat_Click()
    If chkAll.value = 0 Then
        If Val(txtCount.Text) <= 0 Then
            MsgBox "请输入正确的显示条数。", vbInformation, gstrSysName
            txtCount.SetFocus: Exit Sub
        End If
    End If
    
    Call FillStat(True)
    vsItem.SetFocus
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        If Val(vsItem.TextMatrix(i, 1)) <> 0 Then
            vsItem.TextMatrix(i, 2) = 1
        End If
    Next
End Sub

Private Sub cmdSelClear_Click()
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        vsItem.TextMatrix(i, 2) = 0
    Next
End Sub

Private Sub Form_Activate()
    If Not tvw_s.Visible And vsItem.Visible Then vsItem.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    ElseIf Shift = vbAltMask Then
        If Between(KeyCode, vbKey0, vbKey9) Then
            lngIdx = KeyCode - vbKey0 + 1
        End If
        If tabClass.SelectedItem.Index <> lngIdx And Between(lngIdx, 1, tabClass.Tabs.Count) Then
            tabClass.Tabs(lngIdx).Selected = True
        End If
    ElseIf Shift = vbCtrlMask Then
        If KeyCode = vbKeyA Then
            If fraStat.Visible Then cmdSelALL_Click
        ElseIf KeyCode = vbKeyR Then
            If fraStat.Visible Then cmdSelClear_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If tbrOften.Buttons("Often").Visible And tbrOften.Buttons("Often").Enabled Then
            If tbrOften.Buttons("Often").value = tbrPressed Then
                tbrOften.Buttons("Often").value = tbrUnpressed
            Else
                tbrOften.Buttons("Often").value = tbrPressed
            End If
            Call tbrOften_ButtonClick(tbrOften.Buttons("Often"))
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If tbrOften.Buttons("New").Visible And tbrOften.Buttons("New").Enabled Then
            Call tbrOften_ButtonClick(tbrOften.Buttons("New"))
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If tbrOften.Buttons("Del").Visible And tbrOften.Buttons("Del").Enabled Then
            Call tbrOften_ButtonClick(tbrOften.Buttons("Del"))
        End If
    End If
End Sub

Private Function ExistOftenItem(Optional ByVal lng分类ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng分类ID <> 0 Then
        strSQL = "Select ID From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Start With ID=[2] Connect by Prior ID=上级ID"
        strSQL = "Select 1 From 诊疗个人项目 A,诊疗项目目录 B" & _
            " Where A.诊疗项目ID=B.ID And B.分类ID IN(" & strSQL & ") And A.人员ID=[1]" & _
            " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And rownum<2"
    Else
        strSQL = "Select 1 From 诊疗个人项目 Where 人员ID=[1] And rownum<2"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, UserInfo.ID, lng分类ID)
    ExistOftenItem = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim blnDo As Boolean
    Dim str诊疗项目IDs As String, str收费细目IDs As String
    
    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    Call SetFontSize(mbytSize)
    mblnOK = False
    mblnClick = True
    mstrPreNode = ""
    Set mrsItem = Nothing
    mstrPrivs = GetInsidePrivs(IIF(mint范围 = 1, p门诊医嘱下达, p住院医嘱下达))
    mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔
    mlng西药房 = Val(zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "缺省西药房", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id))
    mlng成药房 = Val(zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "缺省成药房", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id))
    mlng中药房 = Val(zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "缺省中药房", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id))
    mlng发料部门 = Val(zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "缺省发料部门", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id))
    
    mstr可用西药房 = zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "可用西药房", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id)
    mstr可用成药房 = zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "可用成药房", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id)
    mstr可用中药房 = zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "可用中药房", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id)
    mstr发料部门 = zlDatabase.GetPara(Decode(mint范围, 1, "门诊", 2, "住院", "") & "可用发料部门", glngSys, Decode(mint范围, 1, p门诊医嘱下达, 2, p住院医嘱下达, 0), , , , , mlng病人科室id)
    If mint场合 = 0 Then
        '字体设置
        mbytSize = zlDatabase.GetPara("字体", glngSys, p门诊医生站, "0")
    ElseIf mint场合 = 2 Then
        mbytSize = zlDatabase.GetPara("字体", glngSys, p医技工作站, "0")
    End If
    '选择器中的设置
    mbln简码 = True '是否显示简码：暂未加参数，固定显示
    If mint场合 <> -1 Then
        mbln价格 = True '是否显示价格：暂未加参数，固定显示
        mbln显示库存 = Val(zlDatabase.GetPara("显示药品库存", glngSys, IIF(mint范围 = 1, p门诊医嘱下达, p住院医嘱下达))) = 1 '是否显示药品库存
    Else
        mbln价格 = False
    End If
    
    lblStatTitle.Caption = Replace(lblStatTitle.Caption, "XXXXXX", UserInfo.姓名)
    cboDate.ListIndex = 0
    If mlng药名ID = 0 Then
        Call SetOftenToolBar(mstr输入 = "")
    Else
        tbrOften.Visible = False
    End If
    
    If mstr输入 = "" Then
        tvw_s.Visible = mlng药名ID = 0
        chkSub.Visible = mlng药名ID = 0
        
        '读取类别失败,已提示,非取消退出
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '无类别,提示,非取消退出
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "没有设置相关诊疗类别,请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
        '如果有个人项目,缺省转到个人项目
        If ExistOftenItem(mlng分类ID) Then
            tbrOften.Buttons("Often").value = tbrPressed
            Call tbrOften_ButtonClick(tbrOften.Buttons("Often"))
        End If
    Else
        fraInfo.Visible = False
        tvw_s.Visible = False
        fraLR.Visible = False
        chkSub.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        Line1(0).Visible = False
        Line1(1).Visible = False
        Shp.Visible = True
        chkShowCause.Visible = True

        '如果有匹配的个人项目,优先显示个人项目
        If ExistOftenItem Then
            tbrOften.Buttons("Often").value = tbrPressed
            Call SwitchToOften(False, False)
            Call FillList(True, str诊疗项目IDs, str收费细目IDs)
            
            '如果没有则切换回来
            If Not cmdOK.Enabled Then
                tbrOften.Buttons("Often").value = tbrUnpressed
                Call SwitchToOften(False, False)
                Call FillList(True, str诊疗项目IDs, str收费细目IDs)
            End If
        Else
            '填充匹配数据
            Call FillList(True, str诊疗项目IDs, str收费细目IDs)
        End If
        
        If cmdOK.Enabled And vsItem.Rows = vsItem.FixedRows + 1 Then
            '只有一个项目时,直接返回
            If tbrOften.Buttons("Often").value = tbrUnpressed Then
                mblnOK = True: Unload Me: Exit Sub
            Else
                blnDo = True '常用项目匹配时始终显示
            End If
        End If
        If (cmdOK.Enabled And vsItem.Rows > vsItem.FixedRows + 1) Or blnDo Then
            '多行是同一个项目时,直接返回:可能无收费细目ID
            If mstr输入 <> "" Then
                If UBound(Split(str诊疗项目IDs, ",")) = 1 _
                    And UBound(Split(str收费细目IDs, ",")) <= 1 Then
                    '常用项目匹配时始终显示
                    If tbrOften.Buttons("Often").value = tbrUnpressed Then
                        mblnOK = True: Unload Me: Exit Sub
                    End If
                End If
            End If
        
            vsItem.Appearance = ccFlat
            vsItem.BorderStyle = ccFixedSingle
            
            Call SetFormSize
            Call Form_Resize
        Else
            '无数据,提示,非取消退出,提示是否查看未匹配的项目。
            If MsgBox("未找到可以使用的诊疗项目、药品或卫材，可能是库存不足等问题，是否查看详细的原因？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set mrsItem = Nothing
                mblnOK = True: Unload Me: Exit Sub
            Else
                chkShowCause.value = 1
            End If
        End If
    End If
End Sub

Private Sub SetFormSize()
    Dim vRect As RECT, i As Long
    Dim lngUpH As Long, lngDnH As Long
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long

    Call zlControl.FormSetCaption(Me, False, False)
    Call GetWindowRect(mobjTXT.hwnd, vRect) '输入框位置
    
    '设置窗体尺寸和位置
    '计算宽度
    Me.Left = vRect.Left * Screen.TwipsPerPixelX
    lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 60 '+3D边框
    For i = 0 To vsItem.Cols - 1
        lngColW = lngColW + IIF(vsItem.ColHidden(i), 0, vsItem.ColWidth(i))
    Next
    If Me.Left + lngColW + lngScrW > Screen.Width - lngScrW Then
        Me.Width = Screen.Width - lngScrW - Me.Left
    Else
        Me.Width = lngColW + lngScrW
    End If
    
    '计算高度
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '屏幕可用高度
    lngUpH = vRect.Top * Screen.TwipsPerPixelY '上面可用高度
    lngDnH = lngScrH - vRect.Bottom * Screen.TwipsPerPixelY '下面可用高度
    Me.Height = vsItem.Rows * vsItem.RowHeight(0) + tbrOften.Height + 45 '395 '+类别卡片高度
    If Me.Height < 2000 Then Me.Height = IIF(mbytSize = 0, 2000, 2500) '窗体最小高度
    If Me.Height > lngUpH And Me.Height > lngDnH Then
        Me.Height = IIF(lngUpH < lngDnH, lngDnH, lngUpH)
    End If
    If Me.Height > lngScrH / 2 Then Me.Height = lngScrH / 2 '窗体最大高度
    If Me.Height <= lngDnH Then
        Me.Top = vRect.Bottom * Screen.TwipsPerPixelY
    ElseIf Me.Height <= lngUpH Then
        Me.Top = vRect.Top * Screen.TwipsPerPixelY - Me.Height
    End If
End Sub
    
Private Sub SetOftenToolBar(ByVal blnCaption As Boolean)
'功能：设置工具条是否显示文本
    Dim lngW As Long, i As Long
    
    For i = 1 To tbrOften.Buttons.Count
        tbrOften.Buttons(i).Caption = IIF(blnCaption, tbrOften.Buttons(i).Description, "")
    Next
    If blnCaption Then
        tbrOften.TextAlignment = tbrTextAlignRight
    Else
        tbrOften.TextAlignment = tbrTextAlignBottom
    End If
    
    For i = 1 To tbrOften.Buttons.Count
        If tbrOften.Buttons(i).Visible Then
            lngW = lngW + tbrOften.Buttons(i).Width
        End If
    Next
    tbrOften.Width = lngW
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long
    
    On Error Resume Next
    
    If mstr输入 = "" Then
        fraInfo.Left = 0
        fraInfo.Width = Me.ScaleWidth
        chkSub.Left = fraInfo.Width - IIF(chkSub.Visible, chkSub.Width, 0) - 45
        lblInfo.Width = IIF(chkSub.Visible, chkSub.Left, fraInfo.Width) - lblInfo.Left - 45
        
        lngLeft = IIF(tvw_s.Visible, tvw_s.Width, 0) + IIF(fraStat.Visible, fraStat.Width, 0) + IIF(fraLR.Visible, fraLR.Width, 0)
        
        If tvw_s.Visible Then
            tvw_s.Left = 0
            tvw_s.Top = fraInfo.Top + fraInfo.Height + 15
            tvw_s.Height = Me.ScaleHeight - tvw_s.Top - 615
        End If
        If fraStat.Visible Then
            fraStat.Left = 0
            fraStat.Top = fraInfo.Top + fraInfo.Height - 90
            fraStat.Height = Me.ScaleHeight - fraStat.Top - 615
        End If
        If fraLR.Visible Then
            fraLR.Top = tvw_s.Top
            fraLR.Left = tvw_s.Left + tvw_s.Width
            fraLR.Height = tvw_s.Height
        End If
        
        vsItem.Top = fraInfo.Top + fraInfo.Height + 15
        vsItem.Left = lngLeft
        vsItem.Width = Me.ScaleWidth - lngLeft
        vsItem.Height = Me.ScaleHeight - vsItem.Top - IIF(mbytSize = 1, 750, 615) - IIF(tabClass.Visible, 350, 0)
        
        If tabClass.Visible Then
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
            tabClass.Left = vsItem.Left + 30
            tabClass.Width = vsItem.Width - 60
        End If
        
        Line1(0).X1 = 0: Line1(0).X2 = Me.ScaleWidth
        Line1(0).Y1 = tvw_s.Top + vsItem.Height + IIF(tabClass.Visible, 350, 0) + 75: Line1(0).Y2 = Line1(0).Y1
        
        Line1(1).X1 = Line1(0).X1: Line1(1).X2 = Line1(0).X2
        Line1(1).Y1 = Line1(0).Y1 - 15: Line1(1).Y2 = Line1(1).Y1
        
        cmdOK.Top = Line1(1).Y1 + 120
        cmdCancel.Top = cmdOK.Top
        
        If Me.ScaleWidth - cmdCancel.Width * 1.8 < 4000 Then
            cmdCancel.Left = 4000
        Else
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.2
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 15
        
        tbrOften.Top = cmdOK.Top + (cmdOK.Height - tbrOften.Height) / 2
        
        lblStore.Top = cmdOK.Top + 70
        lblStore.Left = tbrOften.Left + tbrOften.Width + 80
        
    Else
        Shp.Left = 0
        Shp.Top = 0
        Shp.Width = Me.ScaleWidth
        Shp.Height = Me.ScaleHeight
        
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        'vsItem.Height = Me.ScaleHeight - IIf(tabClass.Tabs.Count > 1, 380, 0)
        vsItem.Height = Me.ScaleHeight - tbrOften.Height + 15 - chkShowCause.Height
        
        tbrOften.Left = Me.ScaleWidth - tbrOften.Width - 15
        tbrOften.Top = vsItem.Top + vsItem.Height + chkShowCause.Height - 30
        
        If chkShowCause.Visible Then
            chkShowCause.Left = tbrOften.Left - chkShowCause.Width - 60
            chkShowCause.Top = tbrOften.Top + tbrOften.Height - chkShowCause.Height - 60
        End If
        
        If tabClass.Tabs.Count > 1 Then
            tabClass.Left = vsItem.Left + 60
            tabClass.Width = vsItem.Width - tbrOften.Width - 120
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
        End If
        
        lblStore.Top = tbrOften.Top + tbrOften.Height - chkShowCause.Height - 20
        lblStore.Left = 80
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mlng药名ID = 0 Then
        If tbrOften.Buttons("Often").value = tbrPressed Then
            Call SaveColPosition("Often")
            Call SaveColWidth("Often")
        Else
            Call SaveColPosition
            Call SaveColWidth
        End If
        Call SaveWinState(Me, App.ProductName, mstrSaveTag)
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
        If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsItem.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvw_s.Width = tvw_s.Width + X
        vsItem.Left = vsItem.Left + X
        vsItem.Width = vsItem.Width - X
        tabClass.Left = tabClass.Left + X
        tabClass.Width = tabClass.Width - X
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node
    
    On Error GoTo errH
    
    If mlng分类ID <> 0 Then
        strSQL = _
            " Select 1 as 级,类型,ID,上级ID,编码,名称 From 诊疗分类目录 Where ID=[1]" & _
            " Union ALL " & _
            " Select Level+1 as 级,类型,ID,上级ID,编码,名称 From 诊疗分类目录" & _
            " Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With 上级ID=[1] Connect by Prior ID=上级ID" & _
            " Order by 级,编码"
    Else
        strSQL = _
            " Select 0 as 级,类型,-类型 as ID,-Null as 上级ID,类型||'' as 编码," & _
            " 类型||'.'||Decode(类型,1,'西成药',2,'中成药',3,'中草药',4,'中药配方',5,'诊疗项目',6,'成套诊疗','7','卫生材料') as 名称" & _
            " From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Group by 类型"
        strSQL = strSQL & " Union ALL " & _
            " Select Level as 级,类型,ID,Nvl(上级ID,-类型) as 上级ID,编码,名称 From 诊疗分类目录" & _
            " Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
            " Order by 级,编码"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng分类ID)
        
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!上级ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!名称, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Close")
        End If
        objNode.Tag = rsTmp!类型 '存放分类类型
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        If mlng定位分类ID = 0 Then
            tvw_s.Nodes(1).Expanded = True
            If tvw_s.Nodes(1).Children > 0 Then
                tvw_s.Nodes(1).Child.Selected = True
            Else
                tvw_s.Nodes(1).Selected = True
            End If
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        Else
            tvw_s.Nodes("_" & mlng定位分类ID).Selected = True
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        End If
    End If
    
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tabClass_Click()
    If Not mblnClick Then Exit Sub
    
    If fraStat.Visible Then
        Call FillStat
    Else
        Call FillList
    End If
    vsItem.SetFocus
End Sub

Private Sub SwitchToOften(Optional ByVal blnFill As Boolean = True, Optional ByVal blnSaveColPos As Boolean = True)
'功能：在常用项目和选择项目界面之间切换
'参数：blnFill=切换之后是否立即刷新清单
    Dim blnNoStat As Boolean
    
    '配方和成套无法统计常用
    If mlng分类ID <> 0 And Not tvw_s.SelectedItem Is Nothing Then
        If InStr(",4,6,", Val(tvw_s.SelectedItem.Tag)) > 0 Then
            blnNoStat = True
        End If
    End If
    tbrOften.Buttons("Stat").Visible = tbrOften.Buttons("Often").value = tbrPressed And mstr输入 = "" And Not blnNoStat
    If mint范围 = 1 Or mint范围 = 2 Then
        '门诊和住院如果没有常用医嘱统计的权限，则隐藏统计按钮
        If InStr(mstrPrivs, ";常用医嘱统计;") = 0 Then
            tbrOften.Buttons("Stat").Visible = False
        End If
    End If
    tbrOften.Buttons("New").Visible = tbrOften.Buttons("Often").value = tbrUnpressed
    tbrOften.Buttons("Del").Visible = tbrOften.Buttons("Often").value = tbrPressed
    If mstr输入 = "" Then
        chkSub.Visible = tbrOften.Buttons("Often").value = tbrUnpressed
        tvw_s.Visible = tbrOften.Buttons("Often").value = tbrUnpressed
        fraLR.Visible = tbrOften.Buttons("Often").value = tbrUnpressed
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").value = tbrPressed Then
                Call SaveColPosition(tvw_s.SelectedItem.Tag)
                Call SaveColWidth(tvw_s.SelectedItem.Tag)
            Else
                Call SaveColPosition("Often")
                Call SaveColWidth("Often")
            End If
        End If
    Else
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").value = tbrPressed Then
                Call SaveColPosition
                Call SaveColWidth
            Else
                Call SaveColPosition("Often")
                Call SaveColWidth("Often")
            End If
        End If
    End If
    Call SetOftenToolBar(mstr输入 = "")
    Call Form_Resize
    
    If blnFill Then Call FillList(True)
End Sub

Private Sub SwitchToState(Optional ByVal blnFill As Boolean = True)
'功能：在常用项目界面，再在统计和常用界面之间切换
'参数：blnFill=切换之后是否立即刷新清单
    If tbrOften.Buttons("Stat").value = tbrUnpressed Then
        fraStat.Visible = False
        tbrOften.Buttons("Del").Visible = True
        tbrOften.Buttons("New").Visible = False
        Call Form_Resize
        If blnFill Then Call FillList(True)
        If Visible Then vsItem.SetFocus
    Else
        fraStat.Visible = True
        lblInfo.Caption = "当前选择：""" & UserInfo.姓名 & """的个人常用项目"
        tbrOften.Buttons("Del").Visible = False
        tbrOften.Buttons("New").Visible = True
        Call Form_Resize
        vsItem.FixedRows = 0: vsItem.Rows = 0
        vsItem.FixedCols = 0: vsItem.Cols = 0
        If Visible Then cboDate.SetFocus
    End If
End Sub

Private Sub tbrOften_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Often" Then
        If Visible And mstr输入 <> "" Then
            zlControl.FormLock Me.hwnd
            Call SwitchToOften
            Call SetFormSize
            Call Form_Resize
            zlControl.FormLock 0
        Else
            '切换回选择界面时先关闭统计界面
            If tbrOften.Buttons("Stat").value = tbrPressed Then
                tbrOften.Buttons("Stat").value = tbrUnpressed
                Call SwitchToState(False)
                Call SwitchToOften(, False)
            Else
                Call SwitchToOften
            End If
        End If
    ElseIf Button.Key = "Stat" Then
        Call SwitchToState
    ElseIf Button.Key = "New" Then
        Call NewOftenNew
    ElseIf Button.Key = "Del" Then
        Call OftenItemDel
    End If
End Sub

Private Sub NewOftenNew()
'功能：将当前诊疗项目加入个人常用项目
    Dim arrSQL As Variant, i As Long, blnTran As Boolean
    Dim lngCol项目 As Long, lngCol次数 As Long
    Dim lngCol收费细目ID As Long, lngCol类别 As Long
    
    arrSQL = Array()
    If Not fraStat.Visible Then
        If mrsItem.EOF Then Exit Sub
        
        ReDim arrSQL(0)
        arrSQL(0) = "ZL_诊疗个人项目_Insert(" & UserInfo.ID & "," & mrsItem!诊疗项目ID & ",Null,'" & _
                 mrsItem!类别ID & "'," & ZVal(Val("" & mrsItem!收费细目ID)) & ")"
    Else
        lngCol项目 = GetCol("诊疗项目ID")
        lngCol收费细目ID = GetCol("收费细目ID")
        lngCol类别 = GetCol("类别ID")
        lngCol次数 = GetCol("次数")
        With vsItem
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 2)) <> 0 And Val(.TextMatrix(i, lngCol项目)) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_诊疗个人项目_Insert(" & UserInfo.ID & "," & Val(.TextMatrix(i, lngCol项目)) & "," & _
                        Val(.TextMatrix(i, lngCol次数)) & ",'" & .TextMatrix(i, lngCol类别) & "'," & ZVal(Val(.TextMatrix(i, lngCol收费细目ID))) & ")"
                End If
            Next
        End With
        If UBound(arrSQL) < 0 Then
            MsgBox "请至少选择一个要加入的常用项目。", vbInformation, gstrSysName
            vsItem.SetFocus: Exit Sub
        Else
            If MsgBox("你当前选择了 " & UBound(arrSQL) + 1 & " 个项目，要把这些项目设置为你的个人常用项目吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
        
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    Screen.MousePointer = 0
    
    If Not fraStat.Visible Then
        MsgBox "项目""" & mrsItem!名称 & """已经加入你的个人常用项目。", vbInformation, gstrSysName
    Else
        MsgBox "所选择的项目已经加入你的个人常用项目。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCol(ByVal strName As String) As Long
    Dim i As Long
    For i = 1 To vsItem.Cols - 1
        If UCase(vsItem.TextMatrix(0, i)) = UCase(strName) Then
            GetCol = i: Exit Function
        End If
    Next
End Function

Private Sub OftenItemDel()
'功能：将当前个人诊疗项目移除
    Dim strSQL As String, lngRow As Long
    
    If mrsItem.EOF Then Exit Sub
    If MsgBox("确实要把""" & mrsItem!名称 & """从你的个人项目中移除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    lngRow = vsItem.Row
    
    strSQL = "ZL_诊疗个人项目_Delete(" & UserInfo.ID & "," & mrsItem!诊疗项目ID & "," & ZVal(Val("" & mrsItem!收费细目ID)) & ")"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    With vsItem
        If lngRow = .FixedRows And .Rows = .FixedRows + 1 Then
            vsItem.Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
        Else
            .RemoveItem lngRow
            If lngRow <= .Rows - 1 Then
                .Row = lngRow
            Else
                .Row = .Rows - 1
            End If
        End If
        Call .ShowCell(.Row, .Col)
        Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mstrPreNode Then Exit Sub
    '结点改变时,保存当前顺序(分类型)
    If Visible Then
        Call SaveColPosition(tvw_s.Nodes(mstrPreNode).Tag)
        Call SaveColWidth(tvw_s.Nodes(mstrPreNode).Tag)
    End If
    mstrPreNode = Node.Key
    
    Call FillList(True)
End Sub

Private Function GetTreePath(ByVal objNode As Node) As String
'功能：获取结点的路径串
    Dim tmpNode As Node, strTmp As String
    Set tmpNode = objNode
    Do While Not tmpNode Is Nothing
        strTmp = IIF(InStr(tmpNode.Text, "[") > 0, zlCommFun.GetNeedName(tmpNode.Text), Mid(tmpNode.Text, 3)) & "\" & strTmp
        Set tmpNode = tmpNode.Parent
    Loop
    GetTreePath = strTmp
End Function

Private Sub txtCount_GotFocus()
    Call zlControl.TxtSelAll(txtCount)
End Sub

Private Sub txtCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then
        If chkShowCause.value = 1 Then
            cmdOK.Enabled = False
            If NewRow >= vsItem.FixedRows Then
                mrsItem.Filter = "KeyID=" & Val(vsItem.TextMatrix(NewRow, GetCol("KeyID")))
                If mrsItem.RecordCount = 1 Then
                    If InStr(mrsItem!未匹配原因, "项目简码不匹配") > 0 Or InStr(mrsItem!未匹配原因, "诊疗项目的执行频率为") > 0 And mint范围 <> 1 Then
                        cmdOK.Enabled = True
                    End If
                End If
            End If
        Else
            If NewRow >= vsItem.FixedRows Then
                mrsItem.Filter = "KeyID=" & Val(vsItem.TextMatrix(NewRow, GetCol("KeyID")))
                '统计出的项目不能直接选择,因为没有管权限,临长嘱等
                cmdOK.Enabled = mrsItem.RecordCount = 1 And Not fraStat.Visible
            Else
                cmdOK.Enabled = False
            End If
        End If
        cmdOK.Visible = Not fraStat.Visible And mstr输入 = ""
        Call ShowDrugStore(NewRow)
    End If
End Sub

Private Sub vsItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strType As String, i As Long
    
    If Order = 0 Then Exit Sub
    
    With vsItem
        .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(1).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(2).Picture
        End If
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Redraw = flexRDDirect
            Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
        End If
            
        '因为可能列顺序改变,所以保存原始列号
        If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
        If tbrOften.Buttons("Often").value = tbrPressed Then strType = "Often" '固定
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", .ColData(Col) & "," & Order
    End With
End Sub

Private Sub vsItem_BeforeSort(ByVal Col As Long, Order As Integer)
    If vsItem.ColDataType(Col) = flexDTBoolean Then
        Order = 0
    Else
        '强制编码列按字符串排序
        If vsItem.TextMatrix(0, Col) = "编码" Then
            If Order = 1 Then Order = 7
            If Order = 2 Then Order = 8
        End If
    End If
End Sub

Private Sub vsItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) = flexDTBoolean Then Cancel = True
End Sub

Private Sub ShowDrugStore(ByVal lngRow As Long)
'功能：显示可用药房中的库存
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str药房 As String
    Dim lng药品ID As Long
    Dim str范围 As String
    Dim i As Long
    Dim strTmp As String
    Dim blnDo As Boolean
    Dim lngCol库存 As Long
    Dim lng类别 As Long
    
    On Error GoTo errH
    
    lblStore.Caption = ""
    lblStore.ToolTipText = ""
    
    With vsItem
        If .Row >= .FixedRows Then
            lng类别 = Val(.TextMatrix(.Row, GetCol("类别ID")))
            If InStr(",5,6,7,", lng类别) > 0 Then
                lngCol库存 = GetCol("库存")
                If .ColHidden(lngCol库存) = False Then
                    If vsItem.TextMatrix(0, lngCol库存) = "库存" Then
                        lng药品ID = Val(.TextMatrix(.Row, GetCol("收费细目ID")))
                        If lng药品ID <> 0 And (mint范围 = 1 Or mint范围 = 2) Then
                            blnDo = True
                        End If
                    End If
                End If
            End If
            If .Cell(flexcpData, .Row, lngCol库存) <> "" Then
                lblStore.Caption = .Cell(flexcpData, .Row, lngCol库存)
                lblStore.ToolTipText = .Cell(flexcpData, .Row, lngCol库存)
                blnDo = False
            End If
        End If
        If blnDo Then
            Select Case lng类别
            Case 5
                str药房 = mstr可用西药房
            Case 6
                str药房 = mstr可用成药房
            Case 7
                str药房 = mstr可用中药房
            End Select
            
            str范围 = Decode(mint范围, 1, "C.门诊", 2, "C.住院")
            
            strSQL = "Select a.名称,Decode(x.库存,Null,Null,Round(x.库存/" & str范围 & "包装,5)||" & str范围 & "单位) As 库存" & _
                " From (Select a.药品id, a.库房id, Nvl(Sum(a.可用数量), 0) As 库存  From 药品库存 A" & _
                " Where a.性质 = 1 And a.库房id In (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X)" & _
                " And (Nvl(a.批次, 0) = 0 Or a.效期 Is Null Or a.效期 > Trunc(Sysdate)) And a.药品id=[2]" & _
                " Group By a.药品id,a.库房id" & _
                " Having Nvl(Sum(a.可用数量),0) <> 0) X, 药品规格 C, 部门表 A" & _
                " Where x.药品id = c.药品id And a.Id = x.库房id"
                
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str药房, lng药品ID)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    strTmp = strTmp & "," & rsTmp!名称 & ":" & rsTmp!库存
                    rsTmp.MoveNext
                Next
            End If
            If strTmp = "" Then
                strTmp = "可用药房中无库存."
            Else
                strTmp = Mid(strTmp, 2) & "."
            End If
            .Cell(flexcpData, .Row, lngCol库存) = strTmp
            lblStore.Caption = strTmp
            lblStore.ToolTipText = strTmp
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsItem_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) <> flexDTBoolean Then
        Cancel = True
    ElseIf Val(vsItem.TextMatrix(Row, 1)) = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsItem_DblClick()
    If vsItem.MouseRow >= vsItem.FixedRows Then
        If cmdOK.Enabled Then
            Call vsItem_KeyPress(13)
        ElseIf fraStat.Visible Then
            Call vsItem_KeyPress(32)
        End If
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdOK.Enabled Then
            cmdOK_Click
        ElseIf fraStat.Visible Then
            If vsItem.Row + 1 <= vsItem.Rows - 1 Then
                vsItem.Row = vsItem.Row + 1
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If fraStat.Visible Then
            If Val(vsItem.TextMatrix(vsItem.Row, 1)) <> 0 Then
                If Val(vsItem.TextMatrix(vsItem.Row, 2)) = 0 Then
                    vsItem.TextMatrix(vsItem.Row, 2) = 1
                Else
                    vsItem.TextMatrix(vsItem.Row, 2) = 0
                End If
            End If
        End If
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If vsItem.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
                vsItem.Row = Val(strIdx)
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    End If
End Sub

Private Sub SaveColPosition(Optional ByVal strType As String)
'功能：保存列顺序:列号,顺序|...
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
    Dim strPos As String, i As Long
        
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    If tbrOften.Buttons("Stat").value = tbrPressed Or fraStat.Visible Then Exit Sub
    
    With vsItem
        For i = 0 To .Cols - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr输入 = "" And strType = "" And tvw_s.Visible And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub RestoreColPosition()
'功能：恢复列顺序
'说明：应放在排序处理之前
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    With vsItem
        If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
        If tbrOften.Buttons("Often").value = tbrPressed Then strType = "Often" '固定
        strPos = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", "")
        If strPos <> "" Then
            rsPos.Fields.Append "Col", adBigInt
            rsPos.Fields.Append "Position", adBigInt
            rsPos.CursorLocation = adUseClient
            rsPos.LockType = adLockOptimistic
            rsPos.CursorType = adOpenStatic
            rsPos.Open
            
            For i = 0 To UBound(Split(strPos, "|"))
                rsPos.AddNew
                rsPos!Col = Split(Split(strPos, "|")(i), ",")(0)
                rsPos!Position = Split(Split(strPos, "|")(i), ",")(1)
                rsPos.Update
            Next
            rsPos.Sort = "Position"
            
            'ColPosition:>=0,ReadOnly,改变后相关列号也改变
            For i = 1 To rsPos.RecordCount
                For j = i - 1 To .Cols - 1
                    If .ColData(j) = rsPos!Col Then Exit For
                Next
                If j <= .Cols - 1 Then
                    .ColPosition(j) = rsPos!Position
                End If
                rsPos.MoveNext
            Next
        End If
    End With
End Sub

Private Sub SaveColWidth(Optional ByVal strType As String)
'功能：保存列宽度
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
    Dim strPos As String, i As Long
        
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'功能：恢复列宽度
'说明：应放在恢复列序之后
    Dim strType As String
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColSort()
'功能：排序处理
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If Val(zlDatabase.GetPara("使用个性化风格")) <> 0 Then
            If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
            If tbrOften.Buttons("Often").value = tbrPressed Then strType = "Often" '固定
            strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", "")
            If strSort <> "" Then
                '因为可能调整列顺序,所以查找真实的排序列
                For i = 0 To .Cols - 1
                    If .ColData(i) = Val(Split(strSort, ",")(0)) Then Exit For
                Next
                If i <= .Cols - 1 Then
                    .Col = i
                    .Sort = Val(Split(strSort, ",")(1))
                    
                    If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(1).Picture
                    Else
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(2).Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function FillList(Optional ByVal blnClass As Boolean, _
    Optional str诊疗项目IDs As String, Optional str收费细目IDs As String) As Boolean
'功能：根据当前界面条件装入诊疗项目目录
'参数：blnClass=是否重建分类卡(应在树形项目改变时才重建)
    Dim objNode As Node, objItem As ListItem
    Dim strSQL As String, strInside As String
    Dim arrClass As Variant, strClass As String
    Dim strSub As String, str操作类型 As String
    Dim str性别 As String, strStock As String
    Dim strInput As String, lng药房ID As Long
    Dim blnLoad As Boolean, objTab As MSComctlLib.Tab
    Dim str范围 As String, str药品 As String
    Dim blnOften As Boolean, blnStock As Boolean
    Dim str库存限制 As String, strPriv As String
    Dim i As Long, j As Long
    Dim strCommIF As String, strScope As String
    Dim blnIsHaveKSS As Boolean
    Dim str诊疗类别 As String
    Dim bln药品 As Boolean, bln卫材 As Boolean, bln其他 As Boolean
    Dim lng分类ID As Long, int类型 As Integer, str类别 As String
    Dim bln显示库存 As Boolean
    Dim str提取字段 As String
    Dim blnBarcode As Boolean          '过滤时卫材是否根据药品库存的商品条码/内部条码来匹配
    Dim str医技成套 As String
    Dim strTsPriv As String

    str诊疗项目IDs = "": str收费细目IDs = ""
    Set objNode = tvw_s.SelectedItem '可能为Nothing
    blnOften = tbrOften.Buttons("Often").value = tbrPressed And mlng药名ID = 0 '是否显示常用项目
    
    '
    bln药品 = True: bln卫材 = True: bln其他 = True
    If mstr诊疗分类 <> "" And mstr输入 <> "" Then
       bln药品 = InStr(",5,6,7,", mstr诊疗分类) > 0
       bln卫材 = mstr诊疗分类 = "4"
       bln其他 = Not (InStr(",4,5,6,7,", mstr诊疗分类) > 0)
    ElseIf mlng药名ID <> 0 Then
        If mintType = 0 Then
            bln卫材 = False: bln药品 = True: bln其他 = False    '显示指定药品所有规格
        ElseIf mintType = 1 Then
            bln卫材 = True: bln药品 = False: bln其他 = False    '显示指定卫材所有规格
        End If
    End If
    
    '是否显示库存选项
    If mint场合 <> -1 Then
        blnStock = mstr输入 <> "" And tabClass.SelectedItem.Index = 1 _
            And ((gbln药品按规格下医嘱 Or mint期效 = 1) And Not (mlng西药房 = 0 And mlng成药房 = 0 And mlng中药房 = 0) _
            Or mlng发料部门 <> 0)
    Else
        blnStock = False
    End If
    
    '清除项目清单及分类卡片
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
    Me.Refresh
    
    '共公条件及字段设置(mstr诊疗分类不为空时,允许显示不能单独应用的项目)
    '------------------------------------------------------------------------
    If mint场合 = 2 Then
        str医技成套 = " and a.类别 <> '9' Or a.类别 = '9' And Exists (Select 1 From 诊疗适用科室 Where 项目id = a.Id And Instr([23], ',' || 科室id || ',') > 0) "
    End If
    strCommIF = " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.服务对象 IN([8],3) Or [8]=3 And Nvl(A.服务对象,0)<>0)"
    strScope = " And ((A.类别<>'9' Or A.类别='9' And (A.人员ID=[11] Or A.人员ID is Null))" & _
            " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And Instr([17],','||科室ID||',')>0)" & str医技成套 & _
            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID)))" & _
             IIF(mstr诊疗分类 <> "", "", "And Nvl(A.单独应用,0)=1") & " And Instr([10],','||Nvl(a.适用性别,0)||',')>0 And Nvl(A.执行频率,0) IN(0,[9])"
            
    If mstr性别 Like "*男*" Then
        str性别 = "0,1"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "0,2"
    ElseIf mstr性别 = "" Then
        str性别 = "0,1,2"
    Else
        str性别 = "0"
    End If
    
    If chkShowCause.value = 1 Then
        strCommIF = "": strScope = " And (A.类别<>'9' Or A.类别='9' And (A.人员ID=[11] Or A.人员ID is Null))"
    End If
    
    '诊疗项目的操作类型
    str操作类型 = "Decode(A.类别," & _
        "'H',Decode(A.操作类型,'1','护理等级','护理常规')," & _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法','4','中药用法','5','特殊治疗','6','采集方法','7','配血方法','8','输血途径',Null)," & _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','7','会诊','8','抢救','9','病重','10','病危','11','死亡','12','记录入出量','14','术前',NULL)," & _
        "A.操作类型)"
    
    If mstr输入 = "" Then
        If mlng药名ID <> 0 Then
            strSub = " And A.ID = [19]"
        Else
            int类型 = Val(objNode.Tag): lng分类ID = Val(Mid(objNode.Key, 2))
            If Not blnOften Then
                '树形中的分类ID
                If chkSub.value = 1 Then
                    '显示下级的项目
                    If Val(Mid(objNode.Key, 2)) < 0 Then
                        strSub = " And A.分类ID IN(" & _
                            " Select ID From 诊疗分类目录 Where 类型=[1] And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                            " )"
                    Else
                        strSub = " And A.分类ID IN(" & _
                            " Select ID From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
                            " Start With ID=[3] Connect by Prior ID=上级ID)"
                    End If
                Else
                    strSub = " And A.分类ID=[3]"
                End If
            ElseIf mlng分类ID <> 0 Then '通过快捷面板确定的分类下面的所有常用项目
                strSub = " And A.分类ID IN(" & _
                    " Select ID From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
                    " Start With ID=[4] Connect by Prior ID=上级ID" & _
                    " )"
            Else
                '显示所有分类,类别中的个人常用项目
            End If
            
            If Not blnOften Or mlng分类ID <> 0 Then
                '树形中的类型确定类别
                If Val(objNode.Tag) = 5 Then
                    strSub = strSub & " And A.类别 Not IN('4','5','6','7','8','9')"
                Else
                    If Val(objNode.Tag) < 8 And Val(objNode.Tag) > 0 Then str类别 = Choose(Val(objNode.Tag), "5", "6", "7", "8", "", "9", "4")
                    If str类别 <> "" Then strSub = strSub & " And A.类别=[2]"
                End If
            End If
        End If
    Else
        '输入匹配:无法确定分类及类别,在所有项目中匹配
        If Len(mstr输入) < 2 Then mstrLike = "" '优化
        strInput = " And (A.编码 Like [5] And B.码类=[7]" & _
            " Or B.名称 Like [6] And B.码类=[7] Or B.简码 Like [6] And B.码类 IN([7],3))"
    End If
    
    '类别卡片确定类别
    If tabClass.SelectedItem.Key <> "" Then
        str类别 = Mid(tabClass.SelectedItem.Key, 2)
        strSub = strSub & " And A.类别=[2]"
    End If
        
    '模块权限
    If mint范围 = 1 Then
        strPriv = GetInsidePrivs(p门诊医嘱下达)
    ElseIf mint范围 = 2 Then
        strPriv = GetInsidePrivs(p住院医嘱下达)
    End If
    
    If mint范围 = 1 Then
        strTsPriv = GetTsPrivs(p门诊医嘱下达)
    ElseIf mint范围 = 2 Then
        strTsPriv = GetTsPrivs(p住院医嘱下达)
    End If
    
    '特殊药品权限
    str药品 = ""
    If strTsPriv <> "" And chkShowCause.value <> 1 Then
        If InStr(strTsPriv, "下达麻醉药嘱") = 0 Then str药品 = str药品 & " And D.毒理分类<>'麻醉药'"
        If InStr(strTsPriv, "下达毒性药嘱") = 0 Then str药品 = str药品 & " And D.毒理分类<>'毒性药'"
        If InStr(strTsPriv, "下达精神药嘱") = 0 Then str药品 = str药品 & " And D.毒理分类 Not IN('精神I类')"
        If InStr(strTsPriv, "下达贵重药嘱") = 0 Then str药品 = str药品 & " And D.价值分类 Not IN('贵重','昂贵')"
    End If
    
    '路径批量调整时,只显示指定诊疗分类的诊疗项目
    If mstr诊疗分类 <> "" Then
        str诊疗类别 = "A.类别 ='" & mstr诊疗分类 & "'"
        If InStr(",C,D,F,G,E,H,Z,", mstr诊疗分类) > 0 Then
            If mstr操作类型 <> "" Then str诊疗类别 = str诊疗类别 & " And A.操作类型='" & mstr操作类型 & "'"
        End If
        If mstr诊疗分类 = "E" Or (mstr诊疗分类 = "D" And mstr操作类型 = "18") Then
            If Val(mstr执行分类) <> 0 Then str诊疗类别 = str诊疗类别 & " And A.执行分类=" & mstr执行分类
        End If
        
    End If
    
    '读取数据
    
    '1.药品列表
    If mstr输入 <> "" And bln药品 Then
        strInput = " And (A.编码 Like [5] And B.码类=[7]" & _
            " Or B.名称 Like [6] And B.码类=[7] Or B.简码 Like [6] And B.码类 IN([7],3))"
        If IsNumeric(mstr输入) Then
            '1X.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.编码 Like [5] And B.码类=[7] Or B.简码 Like [6] And B.码类=3)"
        ElseIf zlCommFun.IsCharAlpha(mstr输入) Then
            'X1.输入全是字母时只匹配简码
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [6] And B.码类=[7]"
        ElseIf zlCommFun.IsCharChinese(mstr输入) Then
            '包含汉字,则只匹配名称
            strInput = " And B.名称 Like [6] And B.码类=[7]"
        End If
    End If
    '按品种下达的长嘱
    If Not (gbln药品按规格下医嘱 Or mint期效 = 1) And bln药品 Then
        
        '药品诊疗项目部分:当分类是药品类型时才读取
        '--------------------------------------------------------------------------------------
        blnLoad = False
        If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
        End If
        If blnLoad Then
            If mstr输入 <> "" Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & IIF(Not mbln简码, "Distinct", "") & _
                        " A.类别 As 类别ID,A.ID as 诊疗项目ID,-Null as 收费细目ID," & _
                        " F.名称 As 类别,Null as 基本,A.编码,B.名称,Null as 商品名," & IIF(mbln简码, "B.简码,", "") & _
                        " A.计算单位,Null as 规格,Null as 产地,D.药品剂型," & str操作类型 & " As 项目特性," & _
                        " Null as 费用类型,Null as 医保大类,Null as 说明,D.处方职务 as 处方职务ID,Null as 价格,Null as 库存,Decode(d.抗生素,0,'',1,'非限制使用',2,'限制使用',3,'特殊使用') as 抗菌等级,D.临床自管药 as 临床自管药ID,NULL as 批次" & _
                        IIF(chkShowCause.value = 1, ",Null as 收费撤挡时间,a.撤档时间,A.站点,Nvl(D.抗生素,0) as 抗生素,null as 费用服务对象,D.毒理分类 as 毒理分类,D.价值分类 as 价值分类,a.服务对象,Nvl(a.执行频率, 0) as 执行频率,Null as 西药可用数量," & vbNewLine & _
                        "  Null as 成药可用数量,Null as 中药可用数量,Null as 使用科室ID,Null as 单独应用,Null as 核算材料,Nvl(A.适用性别,0) as 适用性别,b.码类,NULL AS 未匹配原因", "") & _
                    " From 药品特性 D,诊疗项目类别 F,诊疗项目别名 B,诊疗项目目录 A" & _
                    " Where A.ID=B.诊疗项目ID And A.ID=D.药名ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",5,6,7,", "," & mstr诊疗分类 & ",") > 0, str诊疗类别, "A.类别 = Null"), "A.类别 IN ('5','6','7')") & strCommIF & _
                        IIF(chkShowCause.value = 1, "", " And Instr([10],','||Nvl(A.适用性别,0)||',')>0 And Nvl(A.执行频率,0) IN(0,[9])") & strInput & strSub & str药品
            Else
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select A.类别 As 类别ID,A.ID as 诊疗项目ID,-Null as 收费细目ID," & _
                        " F.名称 As 类别,Null as 基本,A.编码,A.名称,Null as 商品名," & IIF(mbln简码, "Null as 简码,", "") & _
                        "A.计算单位,Null as 规格,Null as 产地, D.药品剂型," & str操作类型 & " As 项目特性," & _
                        "Null as 费用类型,Null as 医保大类,Null as 说明,D.处方职务 as 处方职务ID, Null as 价格,Null as 库存 ,Decode(d.抗生素,0,'',1,'非限制使用',2,'限制使用',3,'特殊使用') as 抗菌等级,D.临床自管药 as 临床自管药ID,NULL as 批次" & _
                    " From 药品特性 D,诊疗项目类别 F,诊疗项目目录 A" & _
                    " Where A.ID=D.药名ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",5,6,7,", "," & mstr诊疗分类 & ",") > 0, str诊疗类别, "A.类别 = Null"), "A.类别 IN ('5','6','7')") & strCommIF & _
                        " And Instr([10],','||Nvl(A.适用性别,0)||',')>0 And Nvl(A.执行频率,0) IN(0,[9])" & strSub & str药品
            End If
        End If
    Else
        
        '药品规格部分:当分类是药品类型时才读取
        '--------------------------------------------------------------------------------------
        blnLoad = False
        If bln药品 Then
            If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
                blnLoad = True
            Else
                blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
            End If
        End If
        
        
        If blnLoad Then
            '药品库存,某一类药房未指定时,读不出库存记录
            strStock = ""
            If mint场合 <> -1 Then
                If mstr输入 = "" Then '根据分类可以确定药品类别
                    If Val(objNode.Tag) = 1 Then
                        lng药房ID = mlng西药房
                    ElseIf Val(objNode.Tag) = 2 Then
                        lng药房ID = mlng成药房
                    ElseIf Val(objNode.Tag) = 3 Then
                        lng药房ID = mlng中药房
                    End If
                Else
                    '没有分类,无法确定药品类别
                    If Mid(tabClass.SelectedItem.Key, 2) = "5" Then
                        lng药房ID = mlng西药房
                    ElseIf Mid(tabClass.SelectedItem.Key, 2) = "6" Then
                        lng药房ID = mlng成药房
                    ElseIf Mid(tabClass.SelectedItem.Key, 2) = "7" Then
                        lng药房ID = mlng中药房
                    End If
                End If
                If chkShowCause.value <> 1 And mbln显示库存 Then
                    If lng药房ID <> 0 Then
                        strStock = _
                            "Select A.药品ID,Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
                            " Where A.性质 = 1 And A.库房ID=[12]" & _
                            " And (Nvl(A.批次, 0) = 0 Or A.效期 Is Null Or A.效期 > Trunc(Sysdate))" & _
                            " Group by A.药品ID Having Nvl(Sum(A.可用数量),0)<>0"
                        bln显示库存 = True
                    ElseIf blnStock And Not (mlng西药房 = 0 And mlng成药房 = 0 And mlng中药房 = 0) Then
                        strStock = _
                            "Select C.药品ID,Nvl(Sum(C.可用数量),0) as 库存" & _
                            " From 药品库存 C,收费项目目录 A" & IIF(strInput <> "", ",收费项目别名 B", "") & _
                            " Where C.性质 = 1 And (Nvl(C.批次,0)=0 Or C.效期 Is Null Or C.效期>Trunc(Sysdate))" & _
                                " And C.库房ID=Decode(A.类别,'5',[13],'6',[14],'7',[15],Null)" & _
                                " And C.药品ID=A.ID And A.类别 IN('5','6','7')" & _
                                 IIF(strInput <> "", " And A.ID=B.收费细目id " & strInput, "") & _
                            " Group by C.药品ID Having Nvl(Sum(C.可用数量),0)<>0"
                        bln显示库存 = True
                        'strStock = "" '优化
                    End If
                End If
            End If
            
            str范围 = Decode(mint范围, 1, "C.门诊", 2, "C.住院", 3, "A.零售")
                
            '是否必须有库存:指定药房时根据系统参数是否要限制库存
            If strStock = "" Then
                str库存限制 = ""
            Else
                str库存限制 = " And A.ID=X.药品ID(+)"
            End If
            If Not (mstr可用西药房 = "" And mstr可用成药房 = "" And mstr可用中药房 = "") And chkShowCause.value <> 1 Then
                '不使用绑定变量，因为这三个参数的相对静态的
                If gblnStock Then
                    str库存限制 = str库存限制 & " And (D.临床自管药=1 Or (" & _
                        " A.类别='5'" & IIF(mstr可用西药房 = "", "", " And Exists(Select 1 From 药品库存" & _
                        " Where 药品ID = c.药品ID And 性质 = 1 And (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And 可用数量>0 And 库房ID In(" & mstr可用西药房 & "))") & _
                        " Or A.类别='6'" & IIF(mstr可用成药房 = "", "", " And Exists(Select 1 From 药品库存" & _
                        " Where 药品ID = c.药品ID And 性质 = 1 And (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And 可用数量>0 And 库房ID In(" & mstr可用成药房 & "))") & _
                        " Or A.类别='7'" & IIF(mstr可用中药房 = "", "", " And Exists(Select 1 From 药品库存 " & _
                        " Where 药品ID = c.药品ID And 性质 = 1 And (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And 可用数量>0 And 库房ID In(" & mstr可用中药房 & "))") & _
                        "))"
                Else
                '只显示指定药房的药品（不用管适用的病人科室）
                    str库存限制 = str库存限制 & " And (D.临床自管药=1 Or Exists(Select 1 From 收费执行科室 X Where x.收费细目id = c.药品ID And " & _
                           "(A.类别='5'" & IIF(mstr可用西药房 = "", "", " And x.执行科室id In (" & mstr可用西药房 & ")") & _
                        " Or A.类别='6'" & IIF(mstr可用成药房 = "", "", " And x.执行科室id In (" & mstr可用成药房 & ")") & _
                        " Or A.类别='7'" & IIF(mstr可用中药房 = "", "", " And x.执行科室id In (" & mstr可用中药房 & ")") & _
                        ")" & IIF(mint范围 <> 3, " And (x.病人来源 is NULL Or x.病人来源=[8])", "") & "))"
                End If
            End If
            If mstr输入 <> "" Then
                '名称根据输入的匹配显示
                strInside = "Select " & IIF(Not mbln简码, "Distinct", "") & _
                    " A.ID,A.类别,A.编码," & IIF(gbyt输入药品显示 = 1, "C2.名称 ,C1.名称 as 商品名,", "B.名称,Null as 商品名,") & IIF(mbln简码, "B.简码,", "") & _
                    " A.计算单位 as 零售单位,1 as 零售包装,A.规格,A.产地,A.费用类型," & _
                        IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')') As 医保大类,", "Null as 医保大类,") & "A.说明,A.是否变价" & _
                        IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,b.码类", "") & _
                    " From 收费项目别名 B,收费项目目录 A" & IIF(gbyt输入药品显示 = 1, ",收费项目别名 C2,收费项目别名 C1", "") & _
                      IIF(mint险类 <> 0, ",保险支付项目 M,保险支付大类 N", "") & _
                    " Where A.ID=B.收费细目ID And A.类别 IN ('5','6','7')" & _
                    IIF(gbyt输入药品显示 = 1, " And A.ID=C1.收费细目ID(+) And C1.码类(+)=1 And C1.性质(+)=3 And A.ID=C2.收费细目ID(+) And C2.码类(+)=1 And C2.性质(+)=1", "") & _
                    strCommIF & strInput & _
                    IIF(mint险类 <> 0, " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[18]", "")
                If mbln价格 Then
                    strInside = "Select A.ID,A.类别,A.编码,A.名称,A.商品名," & IIF(mbln简码, "A.简码,", "") & _
                        " A.零售单位,A.零售包装,A.规格,A.产地,A.费用类型,A.医保大类,A.说明,Sum(Decode(A.是否变价,1,NULL,B.现价)) as 价格,Sum(b.现价) as 现价" & _
                        IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,A.码类", "") & _
                        " From 收费价目 B,(" & strInside & ") A" & _
                        " Where A.ID=B.收费细目ID And Sysdate Between B.执行日期+0 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "20", "21", "22") & _
                        " Group by A.ID,A.类别,A.编码,A.名称,A.商品名," & IIF(mbln简码, "A.简码,", "") & _
                        " A.零售单位,A.零售包装,A.规格,A.产地,A.费用类型,A.医保大类,A.说明" & IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,A.码类", "")
                ElseIf mstrLike = "" And strStock <> "" Then
                    '当可以利用简码索引时(单向匹配),如果有(+)连接(药品库存),则需要Group By一下(奇怪)
                    '当Group by 和Distinct 同时存在时(Not mbln简码)，Oracle会只选择进行Group by
                    strInside = Replace(strInside, "A.是否变价", "Null as 价格")
                    strInside = strInside & " Group By A.ID,A.类别,A.编码," & IIF(gbyt输入药品显示 = 1, "C2.名称 ,C1.名称,", "B.名称,") & IIF(mbln简码, "B.简码,", "") & _
                        " A.计算单位,A.规格,A.产地,A.费用类型," & IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')'),", "") & "A.说明,A.是否变价" & IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,b.码类", "")
                End If
                
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & " Select " & _
                        " A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                        " F.名称 AS 类别,C.基本药物 as 基本,A.编码,A.名称,A.商品名," & IIF(mbln简码, "A.简码,", "") & _
                        " E.计算单位,A.规格,A.产地,D.药品剂型,Null as 项目特性,A.费用类型,A.医保大类,A.说明,D.处方职务 as 处方职务ID" & _
                        IIF(mbln价格, IIF(chkShowCause.value = 1, ",Sum(Decode(A.价格, Null, decode(C.上次售价,null,A.现价,C.上次售价), A.现价)) * ", ",Decode(A.价格, Null, decode(C.上次售价,null,A.现价,C.上次售价), A.现价) * ") & str范围 & "包装 || '/' || " & str范围 & "单位 As 价格", ",Null as 价格") & _
                        IIF(strStock <> "", _
                            IIF(InStr(strPriv, "显示药品库存") = 0, _
                                ",Decode(Sign(Nvl(X.库存,0)),1,'有','') as 库存", _
                                ",Decode(X.库存,NULL,NULL,Round(X.库存/" & str范围 & "包装,5)||" & str范围 & "单位) as 库存"), _
                            ",Null as 库存") & ",Decode(d.抗生素,0,'',1,'非限制使用',2,'限制使用',3,'特殊使用') as 抗菌等级,D.临床自管药 as 临床自管药ID,NULL as 批次" & _
                            IIF(chkShowCause.value = 1, ",a.撤档时间 as 收费撤挡时间,e.撤档时间,A.站点,Nvl(D.抗生素,0)  as 抗生素,a.服务对象 as 费用服务对象,D.毒理分类,D.价值分类," & vbNewLine & _
                    "              e.服务对象,Nvl(e.执行频率, 0) as 执行频率,decode(D.临床自管药,1,1,max(decode(a.类别,'5', decode(instr('," & mstr可用西药房 & ",',',' || n." & IIF(gblnStock, "库房id", "执行科室id") & " || ','),0,0,NVL(N." & IIF(gblnStock, "可用数量", "执行科室id") & ",0)),0))) as 西药可用数量," & vbNewLine & _
                    "              decode(D.临床自管药,1,1,max(decode(a.类别,'6', decode(instr('," & mstr可用成药房 & ",',',' || n." & IIF(gblnStock, "库房id", "执行科室id") & " || ','),0,0,NVL(N." & IIF(gblnStock, "可用数量", "执行科室id") & ",0)),0))) as 成药可用数量," & vbNewLine & _
                    "              decode(D.临床自管药,1,1,max(decode(a.类别,'7',decode(instr('," & mstr可用中药房 & ",',',' || n." & IIF(gblnStock, "库房id", "执行科室id") & " || ','),0,0,NVL(N." & IIF(gblnStock, "可用数量", "执行科室id") & ",0)),0))) as 中药可用数量,null as 使用科室ID, Null as 单独应用,null as 核算材料, nvl(e.适用性别,0) as 适用性别,A.码类,NULL AS 未匹配原因", "") & _
                    " From 药品规格 C,药品特性 D,诊疗项目目录 E,收费项目类别 F,(" & strInside & ") A" & IIF(chkShowCause.value = 1, IIF(gblnStock, ",药品库存 N", ",收费执行科室 N"), "") & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                    " Where A.ID=C.药品ID And C.药名ID=D.药名ID And D.药名ID=E.ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",5,6,7,", "," & mstr诊疗分类 & ",") > 0, Replace(str诊疗类别, "A.", "E."), "E.类别 = Null"), " E.类别 IN ('5','6','7')") & _
                        IIF(chkShowCause.value = 1, "", " And Instr([10],','||Nvl(e.适用性别,0)||',')>0") & _
                        IIF(chkShowCause.value = 1, IIF(gblnStock, " And N.药品ID(+) = c.药品ID AND n.性质(+) = 1 And (Nvl(n.批次, 0) = 0 Or n.效期 Is Null Or n.效期 > Trunc(Sysdate))", " And N.收费细目id(+) = c.药品ID " & IIF(mint范围 <> 3, " And (N.病人来源 is NULL Or N.病人来源=[8])", "")), "") & _
                        IIF(chkShowCause.value <> 1, " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                        " And (E.服务对象 IN([8],3) Or [8]=3 And Nvl(E.服务对象,0)<>0) And Nvl(E.执行频率,0) IN(0,[9])", "") & _
                        str库存限制 & str药品 & Replace(strSub, "A.", "E.") & _
                        IIF(chkShowCause.value = 1, " Group by a.类别, e.Id, a.Id, f.名称,c.基本药物,a.编码, a.名称," & vbNewLine & _
                        "              a.商品名" & IIF(mbln简码, ",A.简码", "") & ", e.计算单位, a.规格, a.产地, d.药品剂型, a.费用类型, a.医保大类, a.说明, d.处方职务," & IIF(mbln价格, "a.价格," & str范围 & "包装, " & str范围 & "单位,", "") & vbNewLine & _
                        "              d.抗生素,a.撤档时间,e.撤档时间,A.站点,a.服务对象,e.服务对象,e.执行频率,D.临床自管药,D.毒理分类,D.价值分类,A.码类,e.适用性别 ", "")

            Else
                '西药名称根据参数设置显示
                If mbln价格 Then
                    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                        "Select A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                            " F.名称 AS 类别,C.基本药物 as 基本,A.编码,Nvl(G1.名称,A.名称) as 名称," & IIF(gbyt药品名称显示 = 2, "G2.名称", "Null") & " as 商品名," & IIF(mbln简码, "Null as 简码,", "") & "E.计算单位,A.规格,A.产地," & _
                            " D.药品剂型,Null as 项目特性,A.费用类型," & IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')') As 医保大类,", "Null as 医保大类,") & "A.说明,D.处方职务 as 处方职务ID," & _
                            " Decode(A.是否变价,1,decode(Sum(c.上次售价),null,Sum(B.现价),sum(c.上次售价))*" & str范围 & "包装 || '/' || " & str范围 & "单位,Sum(B.现价)* " & str范围 & "包装 || '/' || " & str范围 & "单位) As 价格" & _
                            IIF(strStock <> "", _
                                IIF(InStr(strPriv, "显示药品库存") = 0, _
                                    ",Decode(Sign(Nvl(X.库存,0)),1,'有','') as 库存", _
                                    ",Decode(X.库存,NULL,NULL,Round(X.库存/" & str范围 & "包装,5)||" & str范围 & "单位) as 库存"), _
                                ",Null as 库存") & ",Decode(d.抗生素,0,'',1,'非限制使用',2,'限制使用',3,'特殊使用') as 抗菌等级,D.临床自管药 as 临床自管药ID,NULL as 批次" & _
                        " From 收费价目 B,收费项目目录 A,药品规格 C,药品特性 D,诊疗项目目录 E,收费项目类别 F,收费项目别名 G1" & IIF(gbyt药品名称显示 = 2, ",收费项目别名 G2", "") & _
                          IIF(strStock <> "", ",(" & strStock & ") X", "") & IIF(mint险类 <> 0, ",保险支付项目 M,保险支付大类 N", "") & _
                        " Where A.ID=C.药品ID And C.药名ID=D.药名ID And D.药名ID=E.ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",5,6,7,", "," & mstr诊疗分类 & ",") > 0, Replace(str诊疗类别, "A.", "E."), "E.类别 = Null"), " E.类别 IN ('5','6','7')") & _
                            " And Instr([10],','||Nvl(e.适用性别,0)||',')>0 And A.ID=G1.收费细目ID(+) And G1.码类(+)=1 And G1.性质(+)=" & IIF(gbyt药品名称显示 = 1, 3, 1) & _
                            IIF(gbyt药品名称显示 = 2, " And A.ID=G2.收费细目ID(+) And G2.码类(+)=1 And G2.性质(+)=3", "") & _
                            " And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",5,6,7,", "," & mstr诊疗分类 & ",") > 0, Replace(str诊疗类别, "A.", "E."), "E.类别 = Null"), " E.类别 IN ('5','6','7')") & strCommIF & _
                            " And (E.服务对象 IN([8],3) Or [8]=3 And Nvl(E.服务对象,0)<>0) And Nvl(E.执行频率,0) IN(0,[9])" & _
                            " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                            str库存限制 & str药品 & Replace(strSub, "A.", "E.") & _
                            IIF(mint险类 <> 0, " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[18]", "") & _
                            " And A.ID=B.收费细目ID And Sysdate Between B.执行日期+0 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "20", "21", "22") & _
                        " Group by A.类别,E.ID,A.ID,F.名称,C.基本药物,A.编码,Nvl(G1.名称,A.名称)," & IIF(gbyt药品名称显示 = 2, "G2.名称,", "") & "E.计算单位,A.规格,A.产地,D.药品剂型,A.费用类型," & _
                        IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')'),", "") & "A.说明,D.处方职务,A.是否变价," & str范围 & "包装," & str范围 & "单位" & IIF(strStock <> "", ",X.库存", "") & ",d.抗生素,D.临床自管药"
                Else
                    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                        " Select A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                            " F.名称 AS 类别,C.基本药物 as 基本,A.编码,Nvl(G1.名称,A.名称) as 名称," & IIF(gbyt药品名称显示 = 2, "G2.名称", "Null") & " as 商品名," & IIF(mbln简码, "Null as 简码,", "") & _
                            " E.计算单位,A.规格,A.产地,D.药品剂型,Null as 项目特性,A.费用类型," & _
                            IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')') As 医保大类,", "Null as 医保大类,") & "A.说明,D.处方职务 as 处方职务ID,Null as 价格" & _
                            IIF(strStock <> "", _
                                IIF(InStr(strPriv, "显示药品库存") = 0, _
                                    ",Decode(Sign(Nvl(X.库存,0)),1,'有','') as 库存", _
                                    ",Decode(X.库存,NULL,NULL,Round(X.库存/" & str范围 & "包装,5)||" & str范围 & "单位) as 库存"), _
                                ",Null as 库存") & ",Decode(d.抗生素,0,'',1,'非限制使用',2,'限制使用',3,'特殊使用') as 抗菌等级,D.临床自管药 as 临床自管药ID,NULL as 批次" & _
                        " From 收费项目目录 A,药品规格 C,药品特性 D,诊疗项目目录 E,收费项目类别 F,收费项目别名 G1" & IIF(gbyt药品名称显示 = 2, ",收费项目别名 G2", "") & _
                            IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                            IIF(mint险类 <> 0, ",保险支付项目 M,保险支付大类 N", "") & _
                        " Where A.ID=C.药品ID And C.药名ID=D.药名ID And D.药名ID=E.ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",5,6,7,", "," & mstr诊疗分类 & ",") > 0, Replace(str诊疗类别, "A.", "E."), "E.类别 = Null"), " E.类别 IN ('5','6','7')") & _
                            " And Instr([10],','||Nvl(e.适用性别,0)||',')>0 And A.ID=G1.收费细目ID(+) And G1.码类(+)=1 And G1.性质(+)=" & IIF(gbyt药品名称显示 = 1, 3, 1) & _
                            IIF(gbyt药品名称显示 = 2, " And A.ID=G2.收费细目ID(+) And G2.码类(+)=1 And G2.性质(+)=3", "") & _
                            " And A.类别 " & IIF(mstr诊疗分类 <> "" And InStr(",5,6,7,", "," & mstr诊疗分类 & ",") > 0, "='" & mstr诊疗分类 & "'", " IN ('5','6','7')") & strCommIF & _
                            " And (E.服务对象 IN([8],3) Or [8]=3 And Nvl(E.服务对象,0)<>0) And Nvl(E.执行频率,0) IN(0,[9])" & _
                            " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                            str库存限制 & str药品 & Replace(strSub, "A.", "E.") & _
                            IIF(mint险类 <> 0, " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[18]", "")
                End If
            End If
        End If
    End If
        
        
    '2.非药品卫材的诊疗项目部份:分类不是药品类型时不必读取
    '--------------------------------------------------------------------------------------
    blnLoad = False
    If bln其他 Then
        If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,7,", Val(objNode.Tag)) = 0
        End If
    End If
    
    If blnLoad Then
        If mstr输入 <> "" Then
            strInput = " And (A.编码 Like [5] Or B.名称 Like [6] Or B.简码 Like [6]) And B.码类=[7]"
            If IsNumeric(mstr输入) Then
                '1X.输入全是数字时只匹配编码
                If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.编码 Like [5] And B.码类=[7]"
            ElseIf zlCommFun.IsCharAlpha(mstr输入) Then
                'X1.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [6] And B.码类=[7]"
            ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                '包含汉字,则只匹配名称
                strInput = " And B.名称 Like [6] And B.码类=[7]"
            End If
            
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & IIF(Not mbln简码, "Distinct", "") & _
                    " A.类别 As 类别ID,A.ID as 诊疗项目ID,-Null as 收费细目ID," & _
                    " D.名称 As 类别,Null as 基本,A.编码,B.名称,Null as 商品名," & IIF(mbln简码, "B.简码,", "") & _
                    " A.计算单位,A.标本部位 as 规格,Null as 产地,Null as 药品剂型," & str操作类型 & " As 项目特性," & _
                    " Null as 费用类型,Null as 医保大类,Null as 说明,Null as 处方职务ID,Null as 价格,Null as 库存" & ",Null As 抗菌等级,NULL AS 临床自管药ID,NULL as 批次" & _
                    IIF(chkShowCause.value = 1, ",Null as 收费撤挡时间,a.撤档时间,A.站点,null as 抗生素,null as 费用服务对象,Null as 毒理分类,Null as 价值分类,a.服务对象,Nvl(a.执行频率, 0) as 执行频率,Null as 西药可用数量," & vbNewLine & _
                "              Null as 成药可用数量,Null as 中药可用数量,Max(Decode(NVL(e.科室id,0),0,0,Decode(instr([17],',' || e.科室id || ','),0,-1,e.科室id))) as 使用科室ID,a.单独应用,Null as 核算材料,Nvl(A.适用性别,0) as 适用性别,b.码类,NULL AS 未匹配原因", "") & _
                " From 诊疗项目类别 D,诊疗项目别名 B,诊疗项目目录 A" & IIF(chkShowCause.value = 1, ",诊疗适用科室 E", "") & _
                " Where A.ID=B.诊疗项目ID And A.类别=D.编码 And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",4,5,6,7,", "," & mstr诊疗分类 & ",") = 0, str诊疗类别, "A.类别 = Null"), " A.类别 Not IN ('4','5','6','7')") & strScope & strSub & strInput & _
                IIF(mlng病人性质 = 1 And strCommIF <> "", Mid(strCommIF, 1, IIF(Len(strCommIF) = 0, 1, Len(strCommIF)) - 1) & " Or A.类别 = 'Z')", strCommIF) & _
                IIF(chkShowCause.value = 1, " And E.项目ID(+)=A.ID Group by a.类别 , a.Id ,  d.名称, a.编码, b.名称,  " & IIF(mbln简码, "B.简码,", "") & _
                " a.计算单位,a.标本部位 ,a.操作类型,a.撤档时间,A.站点,a.服务对象,a.执行频率,a.单独应用,A.适用性别,b.码类", "")
        Else
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & _
                    " A.类别 As 类别ID,A.ID as 诊疗项目ID,-Null as 收费细目ID,D.名称 As 类别,Null as 基本," & _
                    " A.编码,A.名称,Null as 商品名," & IIF(mbln简码, "Null as 简码,", "") & "A.计算单位,A.标本部位 as 规格,Null as 产地," & _
                    " Null as 药品剂型," & str操作类型 & " As 项目特性,Null as 费用类型,Null as 医保大类,Null as 说明,Null as 处方职务ID," & _
                    " Null as 价格,Null as 库存" & ",Null As 抗菌等级,NULL AS 临床自管药ID,NULL as 批次" & _
                " From 诊疗项目类别 D,诊疗项目目录 A" & _
                " Where A.类别=D.编码 And " & IIF(mstr诊疗分类 <> "", IIF(InStr(",4,5,6,7,", "," & mstr诊疗分类 & ",") = 0, str诊疗类别, "A.类别 = Null"), " A.类别 Not IN ('4','5','6','7')") & strScope & strSub & _
                IIF(mlng病人性质 = 1 And strCommIF <> "", Mid(strCommIF, 1, IIF(Len(strCommIF) = 0, 1, Len(strCommIF)) - 1) & " Or A.类别 = 'Z')", strCommIF)
        End If
    End If
    
    '3.卫材部份:当分类是卫材类型时才读取，都按规格信息读取
    '--------------------------------------------------------------------------------------
    strStock = "" '卫材库存,发料部门未指定时,读不出库存记录
    If mint场合 <> -1 Then
        blnLoad = False
        If mstr输入 = "" Then
            If Val(objNode.Tag) = 7 Then blnLoad = True
        Else
            If Mid(tabClass.SelectedItem.Key, 2) = "4" Then blnLoad = True
        End If
        If (blnLoad Or blnStock And mbln显示库存) And mlng发料部门 <> 0 And chkShowCause.value <> 1 Then
            strStock = _
                "Select A.药品ID,Nvl(Sum(A.可用数量),0) as 库存 From 药品库存 A" & _
                " Where A.性质 = 1 And A.库房ID=[16]" & _
                " And (Nvl(A.批次, 0) = 0 Or A.效期 Is Null Or A.效期 > Trunc(Sysdate))" & _
                " Group by A.药品ID Having Nvl(Sum(A.可用数量),0)<>0"
        End If
    End If
    
    '是否必须有库存
    If strStock = "" Then
        str库存限制 = ""
    Else
        str库存限制 = " And A.ID=X.药品ID(+)"
    End If
    
    If mstr发料部门 <> "" And chkShowCause.value <> 1 Then
        '不使用绑定变量，参数的相对静态的
        If gblnStock Then
            str库存限制 = str库存限制 & " And A.类别='4' And Exists(Select 1 From 药品库存 Where 药品ID = c.材料ID And 性质 = 1 And (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate)) And 可用数量>0 And 库房ID In (" & mstr发料部门 & "))"
        Else
            '只显示指定药房的药品（不用管适用的病人科室）
            str库存限制 = str库存限制 & " And Exists(Select 1 From 收费执行科室 X Where x.收费细目id = c.材料ID And A.类别='4' And x.执行科室id In (" & mstr发料部门 & ")" & IIF(mint范围 <> 3, " And (x.病人来源 is NULL Or x.病人来源=[8])", "") & ")"
        End If
    End If
    
    blnLoad = False
    If bln卫材 Then
        If mlng药名ID <> 0 Then
            blnLoad = True
        Else
            If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
                blnLoad = True
            Else
                blnLoad = Val(objNode.Tag) = 7
            End If
        End If
    End If
    
    If blnLoad Then
        If mstr输入 <> "" Then
            '使用条码匹配的规则：1、全数字或者数字+字母；长度10位数以上；
            If (Not zlCommFun.IsCharChinese(mstr输入)) And Len(mstr输入) >= 10 Then
                strInput = " And (A.编码 Like [5] Or B.名称 Like [6] Or B.简码 Like [6] Or c.商品条码 Like [5] Or c.内部条码 Like [5] ) And B.码类=[7] "
                blnBarcode = True
            Else
                strInput = " And (A.编码 Like [5] Or B.名称 Like [6] Or B.简码 Like [6] ) And B.码类=[7] "
                blnBarcode = False
            End If
            If IsNumeric(mstr输入) Then
                '1X.输入全是数字时只匹配编码
                If Mid(gstrMatchMode, 1, 1) = "1" Then
                    If Len(mstr输入) >= 10 Then
                        strInput = " And (A.编码 Like [5] Or c.商品条码 Like [5] Or c.内部条码 Like [5] ) And B.码类=[7] "
                        blnBarcode = True
                    Else
                        strInput = " And (A.编码 Like [5] ) And B.码类=[7] "
                        blnBarcode = False
                    End If
                End If
            ElseIf zlCommFun.IsCharAlpha(mstr输入) Then
                'X1.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then
                    If Len(mstr输入) >= 10 Then
                        strInput = " And (B.简码 Like [6] Or c.商品条码 Like [5] Or c.内部条码 Like [5] ) And B.码类=[7] "
                        blnBarcode = True
                    Else
                        strInput = " And (B.简码 Like [6] ) And B.码类=[7] "
                        blnBarcode = False
                    End If
                End If
            ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                '包含汉字,则只匹配名称
                strInput = " And B.名称 Like [6] And B.码类=[7]"
                blnBarcode = False
            End If
            '名称根据输入的匹配显示
            strInside = "Select " & IIF(Not mbln简码 Or blnBarcode, "Distinct", "") & _
                " A.ID,A.类别,A.编码,B.名称," & IIF(mbln简码, "B.简码,", "") & "A.计算单位,A.规格,A.产地,A.费用类型," & _
                IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')') As 医保大类,", "Null as 医保大类,") & "A.说明,A.是否变价," & IIF(blnBarcode, "C.批次 ", "NULL as 批次") & _
                IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,b.码类", "") & _
                " From 收费项目别名 B,收费项目目录 A " & IIF(blnBarcode, " , 药品库存 C ", "") & _
                    IIF(mint险类 <> 0, ",保险支付项目 M,保险支付大类 N", "") & _
                " Where A.ID=B.收费细目ID " & IIF(blnBarcode, " And a.Id = c.药品id ", "") & "  And A.类别='4'" & strCommIF & strInput & _
                IIF(mint险类 <> 0, " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[18]", "")
            If mbln价格 Then
                strInside = "Select A.ID,A.类别,A.编码,A.名称," & IIF(mbln简码, "A.简码,", "") & _
                    " A.计算单位,A.规格,A.产地,A.费用类型,A.医保大类,A.说明,Sum(Decode(A.是否变价,1,NULL,B.现价)) as 价格,Sum(b.现价) as 现价," & IIF(blnBarcode, "a.批次", "null as 批次") & _
                    IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,A.码类", "") & _
                    " From 收费价目 B,(" & strInside & ") A" & _
                    " Where A.ID=B.收费细目ID And Sysdate Between B.执行日期+0 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "20", "21", "22") & _
                    " Group by A.ID,A.类别,A.编码,A.名称," & IIF(mbln简码, "A.简码,", "") & _
                    " A.计算单位,A.规格,A.产地,A.费用类型,A.医保大类,A.说明" & IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,A.码类", "") & IIF(blnBarcode, ",a.批次 ", "")
            ElseIf mstrLike = "" And strStock <> "" Then
                '当可以利用简码索引时(单向匹配),如果有(+)连接(药品库存),则需要Group By一下(奇怪)
                '当Group by 和Distinct 同时存在时(Not mbln简码)，Oracle会只选择进行Group by
                strInside = strInside & " Group By A.ID,A.类别,A.编码,B.名称," & IIF(mbln简码, "B.简码,", "") & "A.计算单位,A.规格,A.产地," & _
                " A.费用类型," & IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')'),", "") & "A.说明,A.是否变价" & IIF(chkShowCause.value = 1, ",a.撤档时间,A.站点,a.服务对象,b.码类", "") & IIF(blnBarcode, ",C.批次 ", "")
            End If
            
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                " Select " & _
                    " A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                    " F.名称 AS 类别,Null as 基本,A.编码,A.名称,Null as 商品名," & IIF(mbln简码, "A.简码,", "") & _
                    " A.计算单位,A.规格,A.产地,Null as 药品剂型,Null as 项目特性,A.费用类型,A.医保大类,A.说明,Null as 处方职务ID" & _
                    IIF(mbln价格, ",Decode(A.价格, Null, decode(C.上次售价,null,A.现价,C.上次售价), A.现价)||'/'||A.计算单位 as 价格", ",Null as 价格") & _
                    IIF(strStock <> "", _
                        IIF(InStr(strPriv, "显示药品库存") = 0, _
                            ",Decode(Sign(Nvl(X.库存,0)),1,'有','') as 库存", _
                            ",Decode(X.库存,NULL,NULL,X.库存||A.计算单位) as 库存"), _
                        ",Null as 库存") & ",Null As 抗菌等级,NULL AS 临床自管药ID,A.批次" & _
                IIF(chkShowCause.value = 1, ",a.撤档时间 as 收费撤挡时间,e.撤档时间,A.站点,NULL AS  抗生素,a.服务对象 as 费用服务对象,NULL AS  毒理分类,NULL AS 价值分类," & vbNewLine & _
                "              e.服务对象,Nvl(e.执行频率, 0) as 执行频率,Null as 西药可用数量,Null as 成药可用数量," & vbNewLine & _
                "              Null as 中药可用数量,null as 使用科室ID, Null as 单独应用,c.核算材料 ,Nvl(e.适用性别,0) as 适用性别,A.码类,NULL AS 未匹配原因", "") & _
                " From 材料特性 C,诊疗项目目录 E,收费项目类别 F,(" & strInside & ") A" & IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                " Where A.ID=C.材料ID And C.诊疗ID=E.ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(mstr诊疗分类 = "4", Replace(str诊疗类别, "A.", "E."), " E.类别 = Null"), " E.类别 ='4'") & _
                    IIF(chkShowCause.value <> 1, " And Instr([10],','||Nvl(e.适用性别,0)||',')>0 And C.核算材料=0 And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                    " And (E.服务对象 IN([8],3) Or [8]=3 And Nvl(E.服务对象,0)<>0) And Nvl(E.执行频率,0) IN(0,[9])", "") & _
                    str库存限制 & Replace(strSub, "A.", "E.")
        Else
            If mbln价格 Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    "Select A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                        " F.名称 AS 类别,Null as 基本,A.编码,A.名称,Null as 商品名," & IIF(mbln简码, "Null as 简码,", "") & "A.计算单位,A.规格,A.产地,Null as 药品剂型," & _
                        " Null as 项目特性,A.费用类型," & IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')') As 医保大类,", "Null as 医保大类,") & "A.说明,Null as 处方职务ID," & _
                        " Decode(A.是否变价,1,decode(Sum(c.上次售价),null,Sum(B.现价),sum(c.上次售价)) || '/' || A.计算单位,Sum(B.现价)|| '/' || A.计算单位) As 价格" & _
                        IIF(strStock <> "", _
                            IIF(InStr(strPriv, "显示药品库存") = 0, _
                                ",Decode(Sign(Nvl(X.库存,0)),1,'有','') as 库存", _
                                ",Decode(X.库存,NULL,NULL,X.库存||A.计算单位) as 库存"), _
                            ",Null as 库存") & ",Null As 抗菌等级,NULL AS 临床自管药ID,NULL as 批次" & _
                    " From 收费价目 B,收费项目目录 A,材料特性 C,诊疗项目目录 E,收费项目类别 F" & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        IIF(mint险类 <> 0, ",保险支付项目 M,保险支付大类 N", "") & _
                    " Where A.ID=C.材料ID And C.诊疗ID=E.ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(mstr诊疗分类 = "4", Replace(str诊疗类别, "A.", "E."), " E.类别 = Null"), " E.类别 ='4'") & " And C.核算材料=0" & _
                        " And A.类别='4' And Instr([10],','||Nvl(e.适用性别,0)||',')>0" & strCommIF & _
                        " And (E.服务对象 IN([8],3) Or [8]=3 And Nvl(E.服务对象,0)<>0) And Nvl(E.执行频率,0) IN(0,[9])" & _
                        " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                        IIF(strStock <> "", IIF(gblnStock, " And A.ID=X.药品ID", " And A.ID=X.药品ID(+)"), "") & Replace(strSub, "A.", "E.") & _
                        IIF(mint险类 <> 0, " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[18]", "") & _
                        " And A.ID=B.收费细目ID And Sysdate Between B.执行日期+0 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "A", "B", "20", "21", "22") & _
                    " Group by A.类别,E.ID,A.ID,F.名称,A.编码,A.名称,A.计算单位,A.规格,A.产地,A.费用类型," & IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')'),", "") & _
                    " A.说明,A.是否变价" & IIF(strStock <> "", ",X.库存", "")
            Else
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                        " F.名称 AS 类别,Null as 基本,A.编码,A.名称 as 名称,Null as 商品名," & IIF(mbln简码, "Null as 简码,", "") & "A.计算单位,A.规格,A.产地," & _
                        " Null as 药品剂型,Null as 项目特性,A.费用类型," & IIF(mint险类 <> 0, "n.名称 || Decode(m.保险费用等级, Null, Null, '(' || m.保险费用等级 || ')') As 医保大类,", "Null as 医保大类,") & _
                        " A.说明,Null as 处方职务ID,Null as 价格" & _
                        IIF(strStock <> "", _
                            IIF(InStr(strPriv, "显示药品库存") = 0, _
                                ",Decode(Sign(Nvl(X.库存,0)),1,'有','') as 库存", _
                                ",Decode(X.库存,NULL,NULL,X.库存||A.计算单位) as 库存"), _
                            ",Null as 库存") & ",Null As 抗菌等级,NULL AS 临床自管药ID,NULL as 批次" & _
                    " From 收费项目目录 A,材料特性 C,诊疗项目目录 E,收费项目类别 F" & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        IIF(mint险类 <> 0, ",保险支付项目 M,保险支付大类 N", "") & _
                    " Where A.ID=C.材料ID And C.诊疗ID=E.ID And A.类别=F.编码 And " & IIF(mstr诊疗分类 <> "", IIF(mstr诊疗分类 = "4", Replace(str诊疗类别, "A.", "E."), " E.类别 = Null"), " E.类别 ='4'") & " And C.核算材料=0" & _
                        " And A.类别='4' And Instr([10],','||Nvl(e.适用性别,0)||',')>0" & strCommIF & _
                        " And (E.服务对象 IN([8],3) Or [8]=3 And Nvl(E.服务对象,0)<>0) And Nvl(E.执行频率,0) IN(0,[9])" & _
                        " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                        str库存限制 & Replace(strSub, "A.", "E.") & _
                        IIF(mint险类 <> 0, " And A.ID=M.收费细目ID(+) And M.大类ID=N.ID(+) And M.险类(+)=[18]", "")
            End If
        End If
    End If
    
    '统一的SQL读取结果字段，可见列中[]项目会根据情况进行隐藏
    '类别ID,诊疗项目ID,收费细目ID,类别,基本,编码,名称,[商品名],[简码],计算单位,[规格],[产地],[药品剂型],[项目特性],[费用类型],[医保大类],[说明],处方职务ID,[价格],[库存],[抗菌等级],临床自管药id
    '-------------------------------------------------------------------------------------------------------------------------------------------------
    str提取字段 = "a.类别ID,a.诊疗项目ID,a.收费细目ID,a.类别,a.基本,a.编码,a.名称,a.商品名,a.简码,a.计算单位,a.规格,a.库存,a.产地,a.药品剂型," & _
        "a.项目特性,a.费用类型,a.医保大类,a.说明,a.处方职务ID,a.价格,a.抗菌等级,a.临床自管药id,A.批次"
        
    If chkShowCause.value = 1 Then
        str提取字段 = str提取字段 & ",a.收费撤挡时间,a.撤档时间,a.站点,a.抗生素,a.费用服务对象,a.毒理分类,a.价值分类,a.服务对象,a.执行频率," & _
            "a.西药可用数量,a.成药可用数量,a.中药可用数量,a.使用科室ID,a.单独应用,a.核算材料,a.适用性别,a.码类,a.未匹配原因"
    End If
    
    If blnOften Then
        '包含按品种下达的长嘱药品(非药品类，avg(R.频度)=R.频度,卫材多规格时简化和药品一样)
        If Not (gbln药品按规格下医嘱 Or mint期效 = 1) Then
            strSQL = "Select /*+ rule*/Rownum as KeyID," & str提取字段 & ",r.频度ID " & vbNewLine & _
                    "From (" & strSQL & ") A," & vbNewLine & _
                    "(Select 诊疗项目id, Avg(频度) 频度ID From 诊疗个人项目 Where 人员id = [11] Group By 诊疗项目id) R Where r.诊疗项目id = a.诊疗项目id" & vbNewLine & _
                    "Order by 频度ID Desc,Decode(类别ID,'4','Z',类别ID),类别,编码"
        Else
            strSQL = "Select " & str提取字段 & ",R.频度 as 频度ID From (" & strSQL & ") A,诊疗个人项目 R" & _
                    " Where R.诊疗项目ID=A.诊疗项目ID And (A.收费细目ID is Null Or A.收费细目ID = R.收费细目ID) And R.人员ID=[11]"
                    
            strSQL = "Select /*+ rule*/Rownum as KeyID,A.* From (" & strSQL & ") A Order by 频度ID Desc,Decode(类别ID,'4','Z',类别ID),类别,编码"
        End If
    ElseIf mint范围 = 1 And (mint场合 = 0 Or mint场合 = 2) And mstr输入 <> "" Then
        strSQL = "Select " & str提取字段 & ",R.频度ID From (" & strSQL & ") A," & _
                " (select 诊疗项目id,Avg(使用次数) as 频度ID from 医生常用医嘱 Where 人员id =[11] Group By 诊疗项目id) R" & _
                " Where A.诊疗项目ID=R.诊疗项目ID(+)"
        strSQL = "Select /*+ rule*/Rownum as KeyID,A.* From (" & strSQL & ") A Order by nvl(频度ID,0) Desc,Decode(类别ID,'4','Z',类别ID),类别,编码"
    Else
        strSQL = "Select /*+ rule*/Rownum as KeyID," & str提取字段 & " From (" & strSQL & ") A Order by Decode(类别ID,'4','Z',类别ID),类别,编码"
    End If
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    If chkShowCause.value = 1 Then
        '替换简码
        strSQL = Replace(strSQL, "B.码类=[7]", "B.码类 In(1,2)")
        strSQL = Replace(strSQL, "B.码类 IN([7],3)", "B.码类 IN(1,2,3)")
    End If
    
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int类型, str类别, lng分类ID, mlng分类ID, _
        UCase(mstr输入) & "%", mstrLike & UCase(mstr输入) & "%", mint简码 + 1, mint范围, IIF(mint期效 = 0, 2, 1), _
        "," & str性别 & ",", UserInfo.ID, lng药房ID, mlng西药房, mlng成药房, mlng中药房, mlng发料部门, mstr科室ID, mint险类, mlng药名ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "," & mlng医技科室ID & ",")
    
    '未匹配原因分析
    If chkShowCause.value = 1 Then
        If mrsItem.RecordCount > 0 Then
            Set mrsItem = zlDatabase.CopyNewRec(mrsItem)
            mrsItem.MoveFirst
            Do While Not mrsItem.EOF
                If mrsItem!撤档时间 & "" <> "" And Format(mrsItem!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    mrsItem!未匹配原因 = Decode(mrsItem!类别ID & "", 5, "药品", 6, "药品", 7, "药品", 4, "卫材", "诊疗项目") & "已经停用。"
                ElseIf mrsItem!收费撤挡时间 & "" <> "" And Format(mrsItem!收费撤挡时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    mrsItem!未匹配原因 = Decode(mrsItem!类别ID & "", 5, "药品规格", 6, "药品规格", 7, "药品规格", 4, "卫材规格", "收费项目") & "已经停用。"
                ElseIf mrsItem!站点 & "" <> "" And mrsItem!站点 & "" <> gstrNodeNo Then
                    mrsItem!未匹配原因 = "不是本站点下的项目。"
                ElseIf (Val(mrsItem!类别ID & "") = 4 Or InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And (gbln药品按规格下医嘱 Or mint期效 = 1)) And Val(mrsItem!费用服务对象 & "") <> 3 And Val(mrsItem!费用服务对象 & "") <> mint范围 And mint范围 <> 3 Then
                    mrsItem!未匹配原因 = "规格服务对象不匹配。"
                ElseIf Val(mrsItem!服务对象 & "") <> 3 And Val(mrsItem!服务对象 & "") <> mint范围 And mint范围 <> 3 Or (mint范围 = 3 And Val(mrsItem!服务对象 & "") = 0) Then
                    mrsItem!未匹配原因 = "诊疗项目服务对象不匹配。"
                ElseIf InStr(",4,5,6,7,", "," & mrsItem!类别ID & ",") = 0 And Val(mrsItem!使用科室ID & "") = -1 Then
                    mrsItem!未匹配原因 = "诊疗项目适用科室不匹配，当前科室不可用。"
                ElseIf InStr(",4,5,6,7,", "," & mrsItem!类别ID & ",") = 0 And Val(mrsItem!单独应用 & "") <> 1 Then
                    mrsItem!未匹配原因 = "诊疗项目不可单独使用。"
                ElseIf InStr("," & str性别 & ",", "," & mrsItem!适用性别 & ",") = 0 Then
                    mrsItem!未匹配原因 = "诊疗项目性别不匹配当前病人。"
                ElseIf mrsItem!执行频率 & "" <> "0" And mrsItem!执行频率 & "" <> IIF(mint期效 = 0, 2, 1) & "" Then
                    mrsItem!未匹配原因 = "诊疗项目的执行频率为" & IIF(mrsItem!执行频率 & "" = "1", "一次性", "持续性") & ",不能作为" & IIF(mrsItem!执行频率 & "" = "1", "长嘱。", "临嘱。")
                ElseIf InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And InStr(strTsPriv, "下达麻醉药嘱") = 0 And mrsItem!毒理分类 & "" = "麻醉药" Then
                    mrsItem!未匹配原因 = "药品为麻醉类药品，当前用户没有下达麻醉类药品的权限。"
                ElseIf InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And InStr(strTsPriv, "下达毒性药嘱") = 0 And mrsItem!毒理分类 & "" = "毒性药" Then
                    mrsItem!未匹配原因 = "药品为毒性类药品，当前用户没有下达毒性类药品的权限。"
                ElseIf InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And InStr(strTsPriv, "下达精神药嘱") = 0 And (mrsItem!毒理分类 & "" = "精神I类") Then
                    mrsItem!未匹配原因 = "药品为精神类药品，当前用户没有下达精神类药品的权限。"
                ElseIf InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And InStr(strTsPriv, "下达贵重药嘱") = 0 And (mrsItem!价值分类 & "" = "贵重" Or mrsItem!价值分类 & "" = "昂贵") Then
                    mrsItem!未匹配原因 = "药品为贵重类药品，当前用户没有下达贵重类药品的权限。"
                ElseIf gblnKSSStrict And mint范围 = 1 And InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And mrsItem!抗生素 & "" = "3" Then
                    mrsItem!未匹配原因 = "药品为特殊类抗菌药物，门诊不允许下达。"
                ElseIf InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And Not (mstr可用西药房 = "" And mstr可用成药房 = "" And mstr可用中药房 = "") And gblnStock Then
                    If mrsItem!类别ID = "5" And Val(mrsItem!西药可用数量 & "") = 0 Or mrsItem!类别ID = "6" And Val(mrsItem!成药可用数量 & "") = 0 Or mrsItem!类别ID = "7" And Val(mrsItem!中药可用数量 & "") = 0 Then
                        mrsItem!未匹配原因 = "药品库存不足，系统限制了不允许下达库存不足的药品。"
                    End If
                End If
                If mrsItem!未匹配原因 & "" = "" And InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And Not (mstr可用西药房 = "" And mstr可用成药房 = "" And mstr可用中药房 = "") And Not gblnStock Then
                    If mrsItem!类别ID = "5" And Val(mrsItem!西药可用数量 & "") = 0 Or mrsItem!类别ID = "6" And Val(mrsItem!成药可用数量 & "") = 0 Or mrsItem!类别ID = "7" And Val(mrsItem!中药可用数量 & "") = 0 Then
                        mrsItem!未匹配原因 = "药品只能在适用的药房的执行。"
                    End If
                End If
                If mrsItem!未匹配原因 & "" = "" And mrsItem!类别ID = "4" And mrsItem!核算材料 <> 0 Then
                    mrsItem!未匹配原因 = "卫材是核算材料，不允许单独下达。"
                ElseIf mrsItem!未匹配原因 & "" = "" And mrsItem!码类 <> mint简码 + 1 And zlCommFun.IsCharAlpha(mstr输入) And zlCommFun.IsCharChinese(mrsItem!名称 & "") Then
                    mrsItem!未匹配原因 = "项目简码不匹配，当前是使用的" & IIF(mint简码 + 1 = 1, "拼音", "五笔") & "。"
                End If
                mrsItem.MoveNext
            Loop
            mrsItem.Filter = "未匹配原因 <> ''"
            If mrsItem.RecordCount > 0 Then mrsItem.MoveFirst
        End If
    End If
    '绑定数据
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '可能统计常用项目时设置为了0行0列
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '列属性调整
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        If InStr(",库存,价格,", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.TextMatrix(0, i) = "医保大类" Or vsItem.TextMatrix(0, i) = "医保大类(费用等级)" Then
            vsItem.ColWidth(i) = 2000
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '记录原始列号,用于处理列顺序
    Next
    
    '恢复列顺序:应放在排序处理之前
    Call RestoreColPosition
    Call RestoreColWidth
    '排序处理:先排序,以便后面处理行号
    Call RestoreColSort
    
    '卡片相关数据计算
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '收集类别卡片信息
        If InStr(strClass & ",", "," & mrsItem!类别ID & mrsItem!类别 & ",") = 0 Then
            strClass = strClass & "," & mrsItem!类别ID & mrsItem!类别
        End If
        If InStr(",5,6,7,", "," & mrsItem!类别ID & ",") > 0 And bln显示库存 = True And chkShowCause.value <> 1 Then
            If Val(mrsItem!临床自管药ID & "") = 0 And mrsItem!库存 & "" = "" And (mlng西药房 <> 0 And mrsItem!类别ID = "5" Or mlng成药房 <> 0 And mrsItem!类别ID = "6" Or mlng中药房 <> 0 And mrsItem!类别ID = "7") Then
                '显示了库存但库存不足的药品，用灰色背景显示，但排除自管药
                vsItem.Cell(flexcpBackColor, i, vsItem.FixedCols, i, vsItem.Cols - 1) = &H8000000F
            End If
        End If
        
        '收集项目ID:只收集最多2个
        If mstr输入 <> "" Then
            If UBound(Split(str诊疗项目IDs, ",")) < 2 Then
                If InStr(str诊疗项目IDs & ",", "," & mrsItem!诊疗项目ID & ",") = 0 Then
                    str诊疗项目IDs = str诊疗项目IDs & "," & mrsItem!诊疗项目ID
                End If
            End If
            If UBound(Split(str收费细目IDs, ",")) < 2 Then
                If Not IsNull(mrsItem!收费细目ID) Then
                    If InStr(str收费细目IDs & ",", "," & mrsItem!收费细目ID & ",") = 0 Then
                        str收费细目IDs = str收费细目IDs & "," & mrsItem!收费细目ID
                    End If
                End If
            End If
        End If
        If NVL(mrsItem!抗菌等级) <> "" Then blnIsHaveKSS = True
        mrsItem.MoveNext
    Next
    
    '建立分类卡片:有多类时且项目数较多时
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '用Alt快捷键焦点无法处理
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '根据结果数据类别等情况隐藏一些不必要的列
    For i = 1 To vsItem.Cols - 1
        If vsItem.TextMatrix(0, i) = "商品名" Then
            If (mstr输入 <> "" And gbyt输入药品显示 = 0) Or (mstr输入 = "" And gbyt药品名称显示 <> 2) Then
                vsItem.ColHidden(i) = True '输入时才显示，选择器直接根据参数来
            ElseIf Not ((InStr(strClass, ",5") > 0 Or InStr(strClass, ",6") > 0) And (gbln药品按规格下医嘱 Or mint期效 = 1)) Then
                vsItem.ColHidden(i) = True '成药按规格下达才需要
            End If
        ElseIf vsItem.TextMatrix(0, i) = "简码" Then
            '选择器方式时为空字段
            If mstr输入 = "" Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "规格" Then
            '按品种下达的中西成药没有(诊疗项目为标本部位内容)
            If (InStr(strClass, ",5") > 0 Or InStr(strClass, ",6") > 0) _
                And Not (gbln药品按规格下医嘱 Or mint期效 = 1) Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "基本" Then
            '按品种下达的中西成药不显示，没有药品项时不显示
            If Not (gbln药品按规格下医嘱 Or mint期效 = 1) Then
                vsItem.ColHidden(i) = True
            ElseIf InStr(strClass, ",5") = 0 And InStr(strClass, ",6") = 0 And InStr(strClass, ",7") = 0 Then
                vsItem.ColHidden(i) = True
            End If
        
        ElseIf InStr(",产地,费用类型,医保大类,医保大类(费用等级),说明,价格,库存,", vsItem.TextMatrix(0, i)) > 0 Then
            '固定有收费细目的卫材、中药，或者按规格下达的中西成药
            If Not (InStr(strClass, ",4") > 0 Or InStr(strClass, ",7") > 0 _
                Or ((InStr(strClass, ",5") > 0 Or InStr(strClass, ",6") > 0) _
                    And (gbln药品按规格下医嘱 Or mint期效 = 1))) Then
                vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "产地" And Not gblnShowOrigin Then
                vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "价格" Then
                If Not mbln价格 Then vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "医保大类" Then
                vsItem.TextMatrix(0, i) = "医保大类(费用等级)"
                If mint险类 = 0 Then vsItem.ColHidden(i) = True
            ElseIf vsItem.TextMatrix(0, i) = "医保大类(费用等级)" Then
                If mint险类 = 0 Then vsItem.ColHidden(i) = True
            End If
            If vsItem.TextMatrix(0, i) = "价格" Or vsItem.TextMatrix(0, i) = "库存" Then
                For j = vsItem.FixedRows To vsItem.Rows - 1
                    If Mid(vsItem.TextMatrix(j, i), 1, 1) = "." Then vsItem.TextMatrix(j, i) = "0" & vsItem.TextMatrix(j, i)
                Next
            End If
            If chkShowCause.value = 1 Then
                If vsItem.TextMatrix(0, i) = "库存" Then
                    vsItem.ColHidden(i) = True
                End If
            End If
        ElseIf vsItem.TextMatrix(0, i) = "药品剂型" Then
            '只有药品才有
            If InStr(strClass, ",5") = 0 And InStr(strClass, ",6") = 0 _
                And InStr(strClass, ",7") = 0 And strClass <> "" Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "项目特性" Then
            '药品、卫材不需要
            strSub = Replace(strClass, "4卫材", "")
            strSub = Replace(strSub, "5西成药", "")
            strSub = Replace(strSub, "6中成药", "")
            strSub = Replace(strSub, "7中草药", "")
            strSub = Replace(strSub, ",", "")
            If strSub = "" Then vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "抗菌等级" Then
            If Not blnIsHaveKSS Then vsItem.ColHidden(i) = True
        ElseIf InStr(",收费撤挡时间,撤档时间,站点,抗生素,费用服务对象,毒理分类,价值分类,服务对象,执行频率,西药可用数量,成药可用数量,中药可用数量,单独应用,核算材料,适用性别,码类,", vsItem.TextMatrix(0, i)) > 0 Then
            '隐藏未匹配原因的计算列
            vsItem.ColHidden(i) = True
        ElseIf vsItem.TextMatrix(0, i) = "未匹配原因" Then
            vsItem.ColWidth(i) = 4500
                ElseIf vsItem.TextMatrix(0, i) = "批次" Then
            vsItem.ColHidden(i) = True
        End If
    Next
    
    '行号列宽度
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    If blnOften Then
        lblInfo.Caption = "当前选择：""" & UserInfo.姓名 & """的个人常用项目，共有 " & mrsItem.RecordCount & " 个项目"
    Else
        lblInfo.Caption = "当前选择：" & GetTreePath(tvw_s.SelectedItem) & tabClass.SelectedItem.Tag & "，共有 " & mrsItem.RecordCount & " 个项目"
    End If
    
    vsItem.FrozenCols = 0
    vsItem.Editable = flexEDNone
    vsItem.SheetBorder = vsItem.BackColor
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    zlControl.FormLock 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function

Private Function FillStat(Optional ByVal blnClass As Boolean) As Boolean
'参数：blnClass=是否重建分类卡(应在树形项目改变时才重建)
    Dim objNode As Node, objItem As ListItem
    Dim strSQL As String, i As Long, j As Long
    Dim arrClass As Variant, strClass As String
    Dim str分类 As String, str类别 As String
    Dim str操作类型 As String, str选择类别 As String
    Dim objTab As MSComctlLib.Tab

    Set objNode = tvw_s.SelectedItem '可能为Nothing
    
    '清除项目清单及分类卡片
    '------------------------------------------------------------------------
    vsItem.Editable = flexEDKbdMouse
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
    Me.Refresh
    
    '共公条件及字段设置
    '------------------------------------------------------------------------
    '诊疗项目的操作类型
    str操作类型 = "Decode(A.类别," & _
        "'H',Decode(A.操作类型,'1','护理等级','护理常规')," & _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法','4','中药用法','5','特殊治疗','6','采集方法','7','配血方法',Null)," & _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','7','会诊','8','抢救','9','病重','10','病危','11','死亡','12','记录入出量','14','术前',NULL)," & _
        "A.操作类型)"
    
    '只统计该分类下面的常用项目
    If mlng分类ID <> 0 Then
        str分类 = " And A.分类ID IN(" & _
            " Select ID From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With ID=[1] Connect by Prior ID=上级ID" & _
            " )"
        
        '树形中的类型确定类别
        If Val(objNode.Tag) = 5 Then
            str类别 = str类别 & " And A.诊疗类别 Not IN('4','5','6','7','8','9')"
        Else
            If Val(objNode.Tag) = 1 Then
                str选择类别 = "5"
            ElseIf Val(objNode.Tag) = 2 Then
                str选择类别 = "6"
            ElseIf Val(objNode.Tag) = 3 Then
                str选择类别 = "7"
            ElseIf Val(objNode.Tag) = 4 Then
                str选择类别 = "8"
            ElseIf Val(objNode.Tag) = 6 Then
                str选择类别 = "9"
            ElseIf Val(objNode.Tag) = 7 Then
                str选择类别 = "4"
            End If
            If str选择类别 <> "" Then
                str类别 = str类别 & " And A.诊疗类别=[2]"
            End If
        End If
    End If

    '类别卡片确定类别
    If tabClass.SelectedItem.Key <> "" Then
        str选择类别 = Mid(tabClass.SelectedItem.Key, 2)
        str类别 = str类别 & " And A.诊疗类别=[2]"
    End If
    
    '读取数据:没有限制长嘱/临嘱应用,服务对象,性别范围,药品权限;中药缺省不能单独应用
    '------------------------------------------------------------------------------
    If InStr(",5,6,7,", str选择类别) > 0 And str选择类别 <> "" Then
        strSQL = " Select A.收费细目ID,Count(A.收费细目ID) As 次数" & _
            " From 病人医嘱记录 A" & _
            " Where A.开嘱时间>=[3] And A.开嘱医生=[4] " & str类别 & _
            " Group By A.收费细目ID"
    Else
        strSQL = " Select A.诊疗项目ID,Count(A.诊疗项目ID) As 次数" & _
            " From 病人医嘱记录 A" & _
            " Where A.开嘱时间>=[3] And A.开嘱医生=[4] " & str类别 & _
            " Group By A.诊疗项目ID"
    End If
    If zlDatabase.DateMoved(dtpDate.value) Then
        strSQL = strSQL & " Union ALL " & Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    
    If InStr(",5,6,7,", str选择类别) > 0 And str选择类别 <> "" Then
        '药品诊疗项目部分
        strSQL = _
            " Select A.类别 as 类别ID,A.ID as 诊疗项目ID,B.收费细目ID," & _
            " D.名称 as 类别,A.编码,A.名称,F.规格,A.计算单位,C.药品剂型,B.次数" & _
            " From 诊疗项目类别 D,药品特性 C,诊疗项目目录 A,药品规格 E,(" & strSQL & ") B,收费项目目录 F" & _
            " Where A.类别=D.编码 And A.ID=C.药名ID And A.ID=E.药名ID And E.药品ID=B.收费细目ID And E.药品ID=F.ID" & _
            "   And Not (A.类别='E' And Nvl(A.操作类型,'0')<>'0')" & str分类 & _
            "   And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            "   And Nvl(A.服务对象,0)<>0 And (Nvl(A.单独应用,0)=1 Or A.类别 IN('4','7'))"
        If chkAll.value = 0 Then
            strSQL = _
                "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,规格,计算单位,药品剂型,-1*次数 as 次数 From (" & strSQL & ")" & _
                " Group by -1*次数,类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,规格,计算单位,药品剂型"
            strSQL = "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,规格,计算单位,药品剂型,Abs(次数) as 次数 From (" & strSQL & ") Where Rownum<=[5]"
        End If
    Else
        '非药品部份或所有部份
        strSQL = _
            " Select A.类别 as 类别ID,A.ID as 诊疗项目ID,Nvl(h.id,f.id) as 收费细目ID," & _
            " D.名称 as 类别,Nvl(Nvl(h.编码,f.编码),A.编码) as 编码,A.名称,Nvl(h.规格,f.规格) as 规格,A.计算单位,A.标本部位," & str操作类型 & " As 项目特性,B.次数" & _
            " From 诊疗项目类别 D,诊疗项目目录 A,(" & strSQL & ") B,材料特性 G,收费项目目录 H,药品规格 E,收费项目目录 F" & _
            " Where A.类别=D.编码 And A.ID=B.诊疗项目ID" & str分类 & _
            "   And Not (A.类别='E' And Nvl(A.操作类型,'0')<>'0') And a.id=g.诊疗id(+) And g.材料id=h.id(+) and a.id=e.药名id(+) and e.药品id=f.id(+)" & _
            "   And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            "   And Nvl(A.服务对象,0)<>0 And (Nvl(A.单独应用,0)=1 Or A.类别 IN('4','7'))"
        If chkAll.value = 0 Then
            strSQL = _
                "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,规格,计算单位,标本部位,项目特性,-1*次数 as 次数 From (" & strSQL & ")" & _
                " Group by -1*次数,类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,规格,计算单位,标本部位,项目特性"
            strSQL = "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,规格,计算单位,标本部位,项目特性,Abs(次数) as 次数 From (" & strSQL & ") Where Rownum<=[5]"
        End If
    End If
    strSQL = "Select Rownum as KeyID,Null as 选择,A.* From (" & strSQL & ") A Order by 次数 Desc,编码"
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng分类ID, str选择类别, CDate(Format(dtpDate.value, "yyyy-MM-dd")), UserInfo.姓名, Val(txtCount.Text))
    
    '绑定数据
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '可能统计常用项目时设置为了0行0列
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '列属性调整
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        vsItem.ColAlignment(i) = 1
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '记录原始列号,用于处理列顺序
    Next
    
    '卡片相关数据计算
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '收集类别卡片信息
        If InStr(strClass & ",", "," & mrsItem!类别ID & mrsItem!类别 & ",") = 0 Then
            strClass = strClass & "," & mrsItem!类别ID & mrsItem!类别
        End If
        mrsItem.MoveNext
    Next
    
    '建立分类卡片:有多类时且项目数较多时
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '用Alt快捷键焦点无法处理
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '行号列宽度
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    lblInfo.Caption = "当前选择：""" & UserInfo.姓名 & """的个人常用项目，共有 " & mrsItem.RecordCount & " 个项目"
    
    vsItem.TextMatrix(0, 2) = ""
    vsItem.ColWidth(2) = 300
    vsItem.ColDataType(2) = flexDTBoolean
    vsItem.FrozenCols = 2
    vsItem.Editable = flexEDKbdMouse
    vsItem.SheetBorder = vbBlack
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillStat = True
    Exit Function
errH:
    zlControl.FormLock 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function

Private Sub SetFontSize(ByVal bytSize As Byte)
'功能：进行界面字体的统一设置
'参数：bytSize  0-9号字体，1-12号字体
    Call zlControl.SetPubFontSize(Me, bytSize)
    If bytSize = 1 Then vsItem.RowHeightMin = 300: tvw_s.Font.Size = IIF(bytSize = 0, 9, 12)
End Sub
