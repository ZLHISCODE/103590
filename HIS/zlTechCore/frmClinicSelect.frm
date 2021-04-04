VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin VB.CheckBox chkStock 
      Caption         =   "显示药品库存(&S)"
      Height          =   195
      Left            =   2760
      TabIndex        =   11
      Top             =   5658
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Frame fraInfo 
      Height          =   435
      Left            =   30
      TabIndex        =   15
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
         TabIndex        =   16
         Top             =   165
         Width           =   270
      End
   End
   Begin VB.Frame fraStat 
      Height          =   5160
      Left            =   15
      TabIndex        =   18
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
         Format          =   67174403
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
         TabIndex        =   23
         Top             =   2340
         Width           =   2400
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "条"
         Height          =   180
         Left            =   2100
         TabIndex        =   22
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "显示最前"
         Height          =   180
         Left            =   285
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
      TabIndex        =   17
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
      TabIndex        =   14
      Top             =   555
      Width           =   45
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7335
      TabIndex        =   13
      Top             =   5580
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6240
      TabIndex        =   12
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
Private mblnOK As Boolean
Private mstrPrivs As String
Private mrsItem As ADODB.Recordset
Private mint期效 As Integer
Private mstr性别 As String
Private mstr输入 As String
Private mobjTXT As Object
Private mint范围 As Integer '1-门诊,2-住院
Private mlng分类ID As Long

Private mstrSaveTag As String
Private mstrPreNode As String
Private mblnClick As Boolean

Private mbln价格 As Boolean
Private mbln简码 As Boolean
Private mint简码 As Integer
Private mstrLike As String
Private mlng中药房 As Long
Private mlng西药房 As Long
Private mlng成药房 As Long

Public Function ShowSelect(frmParent As Object, ByVal strPrivs As String, ByVal int期效 As Integer, ByVal str性别 As String, _
    Optional ByVal str输入 As String, Optional objTXT As Object, Optional ByVal int范围 As Integer = 2, _
    Optional ByVal lng分类ID As Long) As ADODB.Recordset
'功能：显示诊疗项目选择器
'参数：int期效=医嘱期效
'      str性别=病人性别
'      str输入=输入匹配的内容,如果没有则为选择器方式,否则为列表方式
'      objTXT=用于列表定位的输入框
'      blnCancel(O):是否取消
'      int范围=1-门诊,2-住院
'      lng分类ID=选择器时(str输入="")，从这个分类开始显示
'返回：如果没有数据,或取消,则返回Nothing；否则为一条包含诊疗项目数据的记录
    mstrPrivs = strPrivs
    mint期效 = int期效
    mstr性别 = str性别
    mstr输入 = str输入
    Set mobjTXT = objTXT
    mint范围 = int范围
    mlng分类ID = lng分类ID
    
    mstrSaveTag = mint范围 & IIF(mstr输入 <> "", 1, 0) & IIF(gbln药品按规格下医嘱 Or mint期效 = 1, 1, 0)
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        Set ShowSelect = mrsItem
    Else
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
                dtpDate.Value = DateAdd("ww", -1, curDate)
            Case 1 '半月(15天)
                dtpDate.Value = DateAdd("d", -15, curDate)
            Case 2 '一月
                dtpDate.Value = DateAdd("m", -1, curDate)
            Case 3 '二月
                dtpDate.Value = DateAdd("m", -2, curDate)
            Case 4 '三月
                dtpDate.Value = DateAdd("m", -3, curDate)
            Case 5 '半年
                dtpDate.Value = DateAdd("m", -6, curDate)
            Case 6 '一年
                dtpDate.Value = DateAdd("yyyy", -1, curDate)
        End Select
    End If
End Sub

Private Sub chkAll_Click()
    txtCount.Enabled = chkAll.Value = 0
End Sub

Private Sub chkStock_Click()
    If Visible Then
        Call FillList
        vsItem.SetFocus
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
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdStat_Click()
    If chkAll.Value = 0 Then
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
    If Not tvw_s.Visible Then vsItem.SetFocus
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
            If tbrOften.Buttons("Often").Value = tbrPressed Then
                tbrOften.Buttons("Often").Value = tbrUnpressed
            Else
                tbrOften.Buttons("Often").Value = tbrPressed
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
        strSQL = "Select ID From 诊疗分类目录 Start With ID=[2] Connect by Prior ID=上级ID"
        strSQL = "Select Count(A.诊疗项目ID) as Num From 诊疗个人项目 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And B.分类ID IN(" & strSQL & ") And A.人员ID=[1]"
    Else
        strSQL = "Select Count(*) as Num From 诊疗个人项目 Where 人员ID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, UserInfo.ID, lng分类ID)
    ExistOftenItem = Nvl(rsTmp!Num, 0) > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim blnDo As Boolean
    Dim str诊疗项目IDs As String, str收费细目IDs As String
    
    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)

    mblnOK = False
    mblnClick = True
    mstrPreNode = ""
    Set mrsItem = Nothing
    
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "") '输入匹配方式
    mint简码 = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0)) '简码匹配方式：0-拼音,1-五笔
    mlng中药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint范围 = 1, "门诊", "住院") & "缺省中药房", 0))
    mlng西药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint范围 = 1, "门诊", "住院") & "缺省西药房", 0))
    mlng成药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint范围 = 1, "门诊", "住院") & "缺省成药房", 0))
    
    '选择器中的设置
    mbln简码 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示项目简码", 1)) <> 0 '是否显示简码
    mbln价格 = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示药品价格", 1)) <> 0 '是否显示药品价格
    chkStock.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示药品库存", 0)) '是否显示药品库存
    
    lblStatTitle.Caption = Replace(lblStatTitle.Caption, "XXXXXX", UserInfo.姓名)
    cboDate.ListIndex = 0
    Call SetOftenToolBar(mstr输入 = "")
    
    If mstr输入 = "" Then
        tvw_s.Visible = True
        
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
            tbrOften.Buttons("Often").Value = tbrPressed
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

        '如果有匹配的个人项目,优先显示个人项目
        If ExistOftenItem Then
            tbrOften.Buttons("Often").Value = tbrPressed
            Call SwitchToOften(False, False)
            Call FillList(True, str诊疗项目IDs, str收费细目IDs)
            
            '如果没有则切换回来
            If Not cmdOK.Enabled Then
                tbrOften.Buttons("Often").Value = tbrUnpressed
                Call SwitchToOften(False, False)
                Call FillList(True, str诊疗项目IDs, str收费细目IDs)
            End If
        Else
            '填充匹配数据
            Call FillList(True, str诊疗项目IDs, str收费细目IDs)
        End If
        
        If cmdOK.Enabled And vsItem.Rows = vsItem.FixedRows + 1 Then
            '只有一个项目时,直接返回
            If tbrOften.Buttons("Often").Value = tbrUnpressed Then
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
                    If tbrOften.Buttons("Often").Value = tbrUnpressed Then
                        mblnOK = True: Unload Me: Exit Sub
                    End If
                End If
            End If
        
            vsItem.Appearance = ccFlat
            vsItem.BorderStyle = ccFixedSingle
                        
            Call SetFormSize
            Call Form_Resize
        Else
            '无数据,提示,非取消退出
            MsgBox "没有找到匹配的诊疗项目。", vbInformation, gstrSysName
            Set mrsItem = Nothing
            mblnOK = True: Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub SetFormSize()
    Dim vRect As RECT, i As Long
    Dim lngUpH As Long, lngDnH As Long
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long

    Call FormSetCaption(Me, False, False)
    Call GetWindowRect(mobjTXT.Hwnd, vRect) '输入框位置
    
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
    If Me.Height < 1500 Then Me.Height = 2000 '窗体最小高度
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
        vsItem.Height = Me.ScaleHeight - vsItem.Top - 615 - IIF(tabClass.Visible, 350, 0)
        
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
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.8
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 15
        
        tbrOften.Top = cmdOK.Top + (cmdOK.Height - tbrOften.Height) / 2
    Else
        Shp.Left = 0
        Shp.Top = 0
        Shp.Width = Me.ScaleWidth
        Shp.Height = Me.ScaleHeight
        
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        'vsItem.Height = Me.ScaleHeight - IIf(tabClass.Tabs.Count > 1, 380, 0)
        vsItem.Height = Me.ScaleHeight - tbrOften.Height - 15
        
        tbrOften.Left = Me.ScaleWidth - tbrOften.Width - 15
        tbrOften.Top = vsItem.Top + vsItem.Height
        
        If chkStock.Visible Then
            chkStock.Left = tbrOften.Left - chkStock.Width - 30
            chkStock.Top = tbrOften.Top + (tbrOften.Height - chkStock.Height) / 2
        End If
        
        If tabClass.Tabs.Count > 1 Then
            tabClass.Left = vsItem.Left + 60
            tabClass.Width = vsItem.Width - tbrOften.Width - IIF(chkStock.Visible, chkStock.Width, 0) - 120
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
        End If
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '选择器中的设置
    If chkStock.Visible Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "显示药品库存", chkStock.Value '是否显示药品库存
    End If

    If tbrOften.Buttons("Often").Value = tbrPressed Then
        Call SaveColPosition("Often")
        Call SaveColWidth("Often")
    Else
        Call SaveColPosition
        Call SaveColWidth
    End If
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or vsItem.Width - x < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + x
        tvw_s.Width = tvw_s.Width + x
        vsItem.Left = vsItem.Left + x
        vsItem.Width = vsItem.Width - x
        tabClass.Left = tabClass.Left + x
        tabClass.Width = tabClass.Width - x
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
            " Select Level+1 as 级,类型,ID,上级ID,编码,名称" & _
            " From 诊疗分类目录 Where 类型<>7 Start With 上级ID=[1] Connect by Prior ID=上级ID" & _
            " Order by 级,编码"
    Else
        strSQL = _
            " Select 0 as 级,类型,-类型 as ID,-NULL as 上级ID,NULL as 编码," & _
            " 类型||'.'||Decode(类型,1,'西成药',2,'中成药',3,'中草药',4,'中药配方',5,'诊疗项目',6,'成套诊疗','7','卫生材料') as 名称" & _
            " From 诊疗分类目录 Where 类型<>7 Group by 类型"
        strSQL = strSQL & " Union ALL " & _
            " Select Level as 级,类型,ID,Nvl(上级ID,-类型) as 上级ID,编码,名称" & _
            " From 诊疗分类目录 Where 类型<>7 Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
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
        tvw_s.Nodes(1).Expanded = True
        If tvw_s.Nodes(1).Children > 0 Then
            tvw_s.Nodes(1).Child.Selected = True
        Else
            tvw_s.Nodes(1).Selected = True
        End If
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
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
    tbrOften.Buttons("Stat").Visible = tbrOften.Buttons("Often").Value = tbrPressed And mstr输入 = "" And Not blnNoStat
    
    tbrOften.Buttons("New").Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
    tbrOften.Buttons("Del").Visible = tbrOften.Buttons("Often").Value = tbrPressed
    If mstr输入 = "" Then
        chkSub.Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
        tvw_s.Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
        fraLR.Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").Value = tbrPressed Then
                Call SaveColPosition(tvw_s.SelectedItem.Tag)
                Call SaveColWidth(tvw_s.SelectedItem.Tag)
            Else
                Call SaveColPosition("Often")
                Call SaveColWidth("Often")
            End If
        End If
    Else
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").Value = tbrPressed Then
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
    If tbrOften.Buttons("Stat").Value = tbrUnpressed Then
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
            LockWindowUpdate Me.Hwnd
            Call SwitchToOften
            Call SetFormSize
            Call Form_Resize
            LockWindowUpdate 0
        Else
            '切换回选择界面时先关闭统计界面
            If tbrOften.Buttons("Stat").Value = tbrPressed Then
                tbrOften.Buttons("Stat").Value = tbrUnpressed
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
    Dim arrSQL As Variant, i As Long
    Dim lngCol项目 As Long, lngCol次数 As Long
    
    arrSQL = Array()
    If Not fraStat.Visible Then
        If mrsItem.EOF Then Exit Sub
        
        ReDim arrSQL(0)
        arrSQL(0) = "ZL_诊疗个人项目_Insert(" & UserInfo.ID & "," & mrsItem!诊疗项目ID & ")"
    Else
        lngCol项目 = GetCol("诊疗项目ID")
        lngCol次数 = GetCol("次数")
        With vsItem
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 2)) <> 0 And Val(.TextMatrix(i, lngCol项目)) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_诊疗个人项目_Insert(" & UserInfo.ID & "," & Val(.TextMatrix(i, lngCol项目)) & "," & Val(.TextMatrix(i, lngCol次数)) & ")"
                End If
            Next
        End With
        If UBound(arrSQL) < 0 Then
            MsgBox "请至少选择一个要加入的常用项目。", vbInformation, gstrSysName
            vsItem.SetFocus: Exit Sub
        Else
            If MsgBox("你当前选择了 " & UBound(arrSQL) + 1 & " 个项目，要把这些项目设置为你的个人常用项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
        
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    Screen.MousePointer = 0
    
    If Not fraStat.Visible Then
        MsgBox "项目""" & mrsItem!名称 & """已经加入你的个人常用项目。", vbInformation, gstrSysName
    Else
        MsgBox "所选择的项目已经加入你的个人常用项目。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
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
        If vsItem.TextMatrix(0, i) = strName Then
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
    
    strSQL = "ZL_诊疗个人项目_Delete(" & UserInfo.ID & "," & mrsItem!诊疗项目ID & ")"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans
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
    gcnOracle.RollbackTrans
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
        strTmp = IIF(InStr(tmpNode.Text, "[") > 0, NeedName(tmpNode.Text), Mid(tmpNode.Text, 3)) & "\" & strTmp
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
        If NewRow >= vsItem.FixedRows Then
            mrsItem.Filter = "KeyID=" & Val(vsItem.TextMatrix(NewRow, 1))
            '统计出的项目不能直接选择,因为没有管权限,临长嘱等
            cmdOK.Enabled = mrsItem.RecordCount = 1 And Not fraStat.Visible
        Else
            cmdOK.Enabled = False
        End If
        cmdOK.Visible = Not fraStat.Visible And mstr输入 = ""
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
        If tbrOften.Buttons("Often").Value = tbrPressed Then strType = "Often" '固定
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
        
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    If tbrOften.Buttons("Stat").Value = tbrPressed Or fraStat.Visible Then Exit Sub
    
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
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    
    With vsItem
        If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
        If tbrOften.Buttons("Often").Value = tbrPressed Then strType = "Often" '固定
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
        
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'功能：恢复列宽度
'说明：应放在恢复列序之后
    Dim strType As String
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    
    If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColSort()
'功能：排序处理
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) <> 0 Then
            If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
            If tbrOften.Buttons("Often").Value = tbrPressed Then strType = "Often" '固定
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
    Dim strSQL As String, i As Long, j As Long
    Dim arrClass As Variant, strClass As String
    Dim strSub As String, str操作类型 As String
    Dim str性别 As String, strStock As String
    Dim strInput As String, lng药房ID As Long
    Dim blnLoad As Boolean, objTab As MSComctlLib.Tab
    Dim str范围 As String, str药品 As String
    Dim blnOften As Boolean, blnStock As Boolean
    Dim str库存限制 As String
    
    Dim lng分类ID As Long, int类型 As Integer, str类别 As String

    str诊疗项目IDs = "": str收费细目IDs = ""
    Set objNode = tvw_s.SelectedItem '可能为Nothing
    blnOften = tbrOften.Buttons("Often").Value = tbrPressed '是否显示常用项目
    
    '是否显示库存选项
    blnStock = mstr输入 <> "" And tabClass.SelectedItem.Index = 1 _
        And Not blnOften And (gbln药品按规格下医嘱 Or mint期效 = 1) _
        And Not (mlng西药房 = 0 And mlng成药房 = 0 And mlng中药房 = 0)
    chkStock.Visible = blnStock
    
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
    
    '共公条件及字段设置
    '------------------------------------------------------------------------
    '诊疗项目的适用性别
    If mstr性别 Like "*男*" Then
        str性别 = "0,1"
    ElseIf mstr性别 Like "*女*" Then
        str性别 = "0,2"
    Else
        str性别 = "0"
    End If
    
    '诊疗项目的操作类型
    str操作类型 = "Decode(A.类别," & _
        "'H',Decode(A.操作类型,'1','护理等级','护理常规')," & _
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法','4','中药用法',Null)," & _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院','7','会诊','8','抢救','9','病重','10','病危','11','死亡',NULL)," & _
        "A.操作类型)"
    
    If mstr输入 = "" Then
        int类型 = Val(objNode.Tag): lng分类ID = Val(Mid(objNode.Key, 2))
        If Not blnOften Then
            '树形中的分类ID
            If chkSub.Value = 1 Then
                '显示下级的项目
                If Val(Mid(objNode.Key, 2)) < 0 Then
                    strSub = " And A.分类ID IN(Select ID From 诊疗分类目录 Where 类型=[1])"
                Else
                    strSub = " And A.分类ID IN(Select ID From 诊疗分类目录 Start With ID=[3] Connect by Prior ID=上级ID)"
                End If
            Else
                strSub = " And A.分类ID=[3]"
            End If
            
            '树形中的类型确定类别
            If Val(objNode.Tag) = 5 Then
                strSub = strSub & " And A.类别 Not IN('4','5','6','7','8','9')"
            Else
                If Val(objNode.Tag) = 1 Then
                    str类别 = "5"
                ElseIf Val(objNode.Tag) = 2 Then
                    str类别 = "6"
                ElseIf Val(objNode.Tag) = 3 Then
                    str类别 = "7"
                ElseIf Val(objNode.Tag) = 4 Then
                    str类别 = "8"
                ElseIf Val(objNode.Tag) = 6 Then
                    str类别 = "9"
                ElseIf Val(objNode.Tag) = 7 Then
                    str类别 = "4"
                End If
                If str类别 <> "" Then
                    strSub = strSub & " And A.类别=[2]"
                End If
            End If
        ElseIf mlng分类ID <> 0 Then '通过快捷面板确定的分类下面的所有常用项目
            strSub = " And A.分类ID IN(Select ID From 诊疗分类目录 Start With ID=[4] Connect by Prior ID=上级ID)"
        
            '树形中的类型确定类别
            If Val(objNode.Tag) = 5 Then
                strSub = strSub & " And A.类别 Not IN('4','5','6','7','8','9')"
            Else
                If Val(objNode.Tag) = 1 Then
                    str类别 = "5"
                ElseIf Val(objNode.Tag) = 2 Then
                    str类别 = "6"
                ElseIf Val(objNode.Tag) = 3 Then
                    str类别 = "7"
                ElseIf Val(objNode.Tag) = 4 Then
                    str类别 = "8"
                ElseIf Val(objNode.Tag) = 6 Then
                    str类别 = "9"
                ElseIf Val(objNode.Tag) = 7 Then
                    str类别 = "4"
                End If
                If str类别 <> "" Then
                    strSub = strSub & " And A.类别=[2]"
                End If
            End If
        Else
            '显示所有分类,类别中的个人常用项目
        End If
    Else
        '输入匹配:无法确定分类及类别,在所有项目中匹配
        If Len(mstr输入) < 2 Then mstrLike = "" '优化
        strInput = " And A.类别<>'4' And (A.编码 Like [5] And B.码类=[7]" & _
            " Or B.名称 Like [6] And B.码类=[7] Or B.简码 Like [6] And B.码类 IN([7],3))"
    End If
    '类别卡片确定类别
    If tabClass.SelectedItem.Key <> "" Then
        str类别 = Mid(tabClass.SelectedItem.Key, 2)
        strSub = strSub & " And A.类别=[2]"
    End If
    
    '读取数据
    '------------------------------------------------------------------------
    If Not (gbln药品按规格下医嘱 Or mint期效 = 1) Then
        '特殊药品权限
        str药品 = ""
        If InStr(mstrPrivs, "下达麻醉药嘱") = 0 Then
            str药品 = str药品 & " And C.毒理分类<>'麻醉药'"
        End If
        If InStr(mstrPrivs, "下达毒性药嘱") = 0 Then
            str药品 = str药品 & " And C.毒理分类<>'毒性药'"
        End If
        If InStr(mstrPrivs, "下达贵重药嘱") = 0 Then
            str药品 = str药品 & " And C.价值分类 Not IN('贵重','昂贵')"
        End If
        
        '药品诊疗项目部分:当分类是药品类型时才读取
        blnLoad = False
        If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
        End If
        If blnLoad Then
            If mstr输入 <> "" Then
                strInput = " And A.类别<>'4' And (A.编码 Like [5] And B.码类=[7]" & _
                    " Or B.名称 Like [6] And B.码类=[7] Or B.简码 Like [6] And B.码类 IN([7],3))"
                If IsNumeric(mstr输入) Then
                    '1X.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.类别<>'4' And (A.编码 Like [5] And B.码类=[7] Or B.简码 Like [6] And B.码类=3)"
                ElseIf zlCommFun.IsCharAlpha(mstr输入) Then
                    'X1.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.类别<>'4' And B.简码 Like [6] And B.码类=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                    '包含汉字,则只匹配名称
                    strInput = " And A.类别<>'4' And B.名称 Like [6] And B.码类=[7]"
                End If
                
                strSQL = _
                    " Select " & IIF(Not mbln简码, "Distinct", "") & _
                        " A.类别 As 类别ID,A.ID as 诊疗项目ID,NULL as 收费细目ID," & _
                        " D.名称 As 类别,A.编码,B.名称," & IIF(mbln简码, "B.简码,", "") & _
                        " A.计算单位,A.标本部位,C.药品剂型," & str操作类型 & " As 项目特性," & _
                        " C.处方职务 as 处方职务ID" & IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                    " From 药品特性 C,诊疗项目类别 D,诊疗项目别名 B,诊疗项目目录 A" & IIF(blnOften, ",诊疗个人项目 R", "") & _
                    " Where A.ID=B.诊疗项目ID And A.ID=C.药名ID And A.类别=D.编码 And A.类别 IN ('5','6','7')" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                        " And A.服务对象 IN([8],3) And Nvl(A.单独应用,0)=1 And Instr([10],','||Nvl(A.适用性别,0)||',')>0" & _
                        " And Nvl(A.执行频率,0) IN(0,[9])" & strInput & strSub & str药品 & _
                        IIF(blnOften, " And R.诊疗项目ID=A.ID And R.人员ID=[11]", "")
            Else
                strSQL = _
                    " Select " & _
                        " A.类别 As 类别ID,A.ID as 诊疗项目ID,NULL as 收费细目ID," & _
                        " D.名称 As 类别,A.编码,A.名称,A.计算单位,A.标本部位," & _
                        " C.药品剂型," & str操作类型 & " As 项目特性,C.处方职务 as 处方职务ID" & _
                        IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                    " From 药品特性 C,诊疗项目类别 D,诊疗项目目录 A" & IIF(blnOften, ",诊疗个人项目 R", "") & _
                    " Where A.ID=C.药名ID And A.类别=D.编码 And A.类别 IN ('5','6','7')" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                        " And A.服务对象 IN([8],3) And Nvl(A.单独应用,0)=1 And Instr([10],','||Nvl(A.适用性别,0)||',')>0" & _
                        " And Nvl(A.执行频率,0) IN(0,[9])" & strSub & str药品 & _
                        IIF(blnOften, " And R.诊疗项目ID=A.ID And R.人员ID=[11]", "")
            End If
        End If
        
        '非药品诊疗项目部份:分类不是药品类型时不必读取
        blnLoad = False
        If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) = 0
        End If
        If blnLoad Then
            If mstr输入 <> "" Then
                strInput = " And A.类别<>'4' And (A.编码 Like [5] Or B.名称 Like [6] Or B.简码 Like [6]) And B.码类=[7]"
                If IsNumeric(mstr输入) Then
                    '1X.输入全是数字时只匹配编码
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.类别<>'4' And A.编码 Like [5] And B.码类=[7]"
                ElseIf zlCommFun.IsCharAlpha(mstr输入) Then
                    'X1.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.类别<>'4' And B.简码 Like [6] And B.码类=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                    '包含汉字,则只匹配名称
                    strInput = " And A.类别<>'4' And B.名称 Like [6] And B.码类=[7]"
                End If
            
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & IIF(Not mbln简码, "Distinct", "") & _
                        " A.类别 As 类别ID,A.ID as 诊疗项目ID,NULL as 收费细目ID," & _
                        " D.名称 As 类别,A.编码,B.名称," & IIF(mbln简码, "B.简码,", "") & _
                        " A.计算单位,A.标本部位,NULL as 药品剂型," & str操作类型 & " As 项目特性," & _
                        " Null as 处方职务ID" & IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                    " From 诊疗项目类别 D,诊疗项目别名 B,诊疗项目目录 A" & IIF(blnOften, ",诊疗个人项目 R", "") & _
                    " Where A.ID=B.诊疗项目ID And A.类别=D.编码 And A.类别 Not IN('5','6','7')" & _
                        " And (A.类别<>'9' Or A.人员ID=[11] Or A.人员ID is Null)" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                        " And A.服务对象 IN([8],3) And Nvl(A.单独应用,0)=1 And Instr([10],','||Nvl(A.适用性别,0)||',')>0" & _
                        " And Nvl(A.执行频率,0) IN(0,[9])" & strInput & strSub & _
                        IIF(blnOften, " And R.诊疗项目ID=A.ID And R.人员ID=[11]", "")
            Else
                '从树形分类中选择的非药品诊疗项目,不可能与药品同时显示,因此可以减少一些字段
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & _
                        " A.类别 As 类别ID,A.ID as 诊疗项目ID,NULL as 收费细目ID,D.名称 As 类别," & _
                        " A.编码,A.名称,A.计算单位,A.标本部位" & IIF(blnOften, ",Null as 药品剂型", "") & "," & _
                        str操作类型 & " As 项目特性,Null as 处方职务ID" & IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                    " From 诊疗项目类别 D,诊疗项目目录 A" & IIF(blnOften, ",诊疗个人项目 R", "") & _
                    " Where A.类别=D.编码 And A.类别 Not IN('5','6','7')" & _
                        " And (A.类别<>'9' Or A.人员ID=[11] Or A.人员ID is Null)" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                        " And A.服务对象 IN([8],3) And Nvl(A.单独应用,0)=1 And Instr([10],','||Nvl(A.适用性别,0)||',')>0" & _
                        " And Nvl(A.执行频率,0) IN(0,[9])" & strSub & _
                        IIF(blnOften, " And R.诊疗项目ID=A.ID And R.人员ID=[11]", "")
            End If
        End If
    Else
        '药品库存,某一类药房未指定时,读不出库存记录
        If blnOften Then
            strStock = "" '没有分类,无法确定药品类别,否则按下面的方式较慢
        Else
            If mstr输入 = "" Then
                '根据分类可以确定药品类别
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
            If lng药房ID <> 0 Then
                strStock = _
                    "Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存 From 药品库存 A" & _
                    " Where A.性质 = 1 And A.库房ID=[12]" & _
                    " And (Nvl(A.批次, 0) = 0 Or A.效期 Is Null Or A.效期 > Trunc(Sysdate))" & _
                    " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
            ElseIf blnStock And chkStock.Value = 1 Then
                strStock = _
                    "Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存" & _
                    " From 药品库存 A,收费项目目录 B" & _
                    " Where A.性质 = 1 And (Nvl(A.批次,0)=0 Or A.效期 Is Null Or A.效期>Trunc(Sysdate))" & _
                        " And A.库房ID=Decode(B.类别,'5',[13],'6',[14],'7',[15],Null)" & _
                        " And A.药品ID=B.ID And B.类别 IN('5','6','7')" & _
                    " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
                'strStock = "" '优化
            End If
        End If
        
        '特殊药品权限
        str药品 = ""
        If InStr(mstrPrivs, "下达麻醉药嘱") = 0 Then
            str药品 = str药品 & " And D.毒理分类<>'麻醉药'"
        End If
        If InStr(mstrPrivs, "下达毒性药嘱") = 0 Then
            str药品 = str药品 & " And D.毒理分类<>'毒性药'"
        End If
        If InStr(mstrPrivs, "下达贵重药嘱") = 0 Then
            str药品 = str药品 & " And D.价值分类 Not IN('贵重','昂贵')"
        End If
        
        '药品规格部分:当分类是药品类型时才读取
        blnLoad = False
        If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
        End If
        If blnLoad Then
            str范围 = IIF(mint范围 = 1, "门诊", "住院")
                
            '是否必须有库存:指定药房时根据系统参数是否要限制库存
            str库存限制 = " And A.ID=X.药品ID(+)"
            If gblnStock Then
                str库存限制 = str库存限制 & " And (" & _
                    " A.类别='5' And ([13]=0 Or X.药品ID Is Not Null)" & _
                    " Or A.类别='6' And ([14]=0 Or X.药品ID Is Not Null)" & _
                    " Or A.类别='7' And ([15]=0 Or X.药品ID Is Not Null)" & _
                    ")"
            End If
                
            If mstr输入 <> "" Then
                strInput = " And A.类别<>'4' And (A.编码 Like [5] And B.码类=[7]" & _
                    " Or B.名称 Like [6] And B.码类=[7] Or B.简码 Like [6] And B.码类 IN([7],3))"
                If IsNumeric(mstr输入) Then
                    '1X.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.类别<>'4' And (A.编码 Like [5] And B.码类=[7] Or B.简码 Like [6] And B.码类=3)"
                ElseIf zlCommFun.IsCharAlpha(mstr输入) Then
                    'X1.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.类别<>'4' And B.简码 Like [6] And B.码类=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                    '包含汉字,则只匹配名称
                    strInput = " And A.类别<>'4' And B.名称 Like [6] And B.码类=[7]"
                End If
            
                '名称根据输入的匹配显示
                strSQL = "Select " & IIF(Not mbln简码, "Distinct", "") & _
                    " A.ID,A.类别,A.编码,B.名称," & IIF(mbln简码, "B.简码,", "") & _
                    " A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                    " From 收费项目别名 B,收费项目目录 A" & IIF(blnOften, ",药品规格 Q,诊疗个人项目 R", "") & _
                    " Where A.ID=B.收费细目ID And A.类别 IN ('5','6','7')" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And A.服务对象 IN([8],3)" & strInput & _
                    IIF(blnOften, " And Q.药品ID=A.ID And R.诊疗项目ID=Q.药名ID And R.人员ID=[11]", "")
                If mbln价格 Then
                    strSQL = "Select A.ID,A.类别,A.编码,A.名称," & IIF(mbln简码, "A.简码,", "") & _
                        " A.规格,A.产地,A.费用类型,A.说明,Sum(Decode(A.是否变价,1,NULL,B.现价)) as 价格" & IIF(blnOften, ",A.频度ID", "") & _
                        " From 收费价目 B,(" & strSQL & ") A" & _
                        " Where A.ID=B.收费细目ID And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Group by A.ID,A.类别,A.编码,A.名称," & IIF(mbln简码, "A.简码,", "") & _
                        "A.规格,A.产地,A.费用类型,A.说明" & IIF(blnOften, ",A.频度ID", "")
                ElseIf mstrLike = "" And strStock <> "" Then
                    '当可以利用简码索引时(单向匹配),如果有(+)连接(药品库存),则需要Group By一下(奇怪)
                    '当Group by 和Distinct 同时存在时(Not mbln简码)，Oracle会只选择进行Group by
                    strSQL = strSQL & " Group By A.ID,A.类别,A.编码,B.名称," & IIF(mbln简码, "B.简码,", "") & _
                        " A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "")
                End If
                
                strSQL = _
                    " Select " & _
                        " A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                        " F.名称 AS 类别,A.编码,A.名称," & IIF(mbln简码, "A.简码,", "") & _
                        " E.计算单位,A.规格,A.产地,D.药品剂型,NULL AS 项目特性,A.费用类型,A.说明,D.处方职务 as 处方职务ID" & _
                        IIF(mbln价格, ",Decode(A.价格,NULL,NULL,LTrim(To_Char(A.价格*C." & str范围 & "包装,'999990.0000'))||'/'||C." & str范围 & "单位) as 价格", "") & _
                        IIF(strStock <> "", ",Decode(X.库存,NULL,NULL,LTrim(To_Char(X.库存/C." & str范围 & "包装,'999990.0000'))||C." & str范围 & "单位) as 库存", "") & _
                        IIF(blnOften, ",A.频度ID", "") & _
                    " From 药品规格 C,药品特性 D,诊疗项目目录 E,收费项目类别 F,(" & strSQL & ") A" & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                    " Where A.ID=C.药品ID And C.药名ID=D.药名ID And D.药名ID=E.ID" & _
                        " And A.类别=F.编码 And E.类别 IN('5','6','7')" & _
                        " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                        " And E.服务对象 IN([8],3) And Nvl(E.执行频率,0) IN(0,[9])" & _
                        IIF(strStock <> "", str库存限制, "") & str药品 & Replace(strSub, "A.", "E.")
            Else
                '西药名称根据参数设置显示
                If mbln价格 Then
                    strSQL = _
                        " Select A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                            " F.名称 AS 类别,A.编码,Nvl(G.名称,A.名称) as 名称,E.计算单位,A.规格,A.产地,D.药品剂型,NULL AS 项目特性,A.费用类型,A.说明,D.处方职务 as 处方职务ID," & _
                            " Decode(A.是否变价,1,NULL,LTrim(To_Char(Sum(B.现价)*C." & str范围 & "包装,'999990.0000'))||'/'||C." & str范围 & "单位) as 价格" & _
                            IIF(strStock <> "", ",Decode(X.库存,NULL,NULL,LTrim(To_Char(X.库存/C." & str范围 & "包装,'999990.0000'))||C." & str范围 & "单位) as 库存", "") & _
                            IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                        " From 收费价目 B,收费项目目录 A,药品规格 C,药品特性 D,诊疗项目目录 E,收费项目类别 F,收费项目别名 G" & _
                            IIF(blnOften, ",诊疗个人项目 R", "") & IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        " Where A.ID=C.药品ID And C.药名ID=D.药名ID And D.药名ID=E.ID" & _
                            " And A.类别=F.编码 And E.类别 IN('5','6','7')" & _
                            " And A.ID=G.收费细目ID(+) And G.码类(+)=1 And G.性质(+)=" & IIF(gbln商品名, 3, 1) & _
                            " And E.服务对象 IN([8],3) And Nvl(E.执行频率,0) IN(0,[9])" & _
                            " And A.类别 IN ('5','6','7') And A.服务对象 IN([8],3)" & _
                            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                            " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                            IIF(blnOften, " And R.诊疗项目ID=E.ID And R.人员ID=[11]", "") & _
                            IIF(strStock <> "", str库存限制, "") & str药品 & Replace(strSub, "A.", "E.") & _
                            " And A.ID=B.收费细目ID And Sysdate Between B.执行日期 And Nvl(B.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Group by A.类别,E.ID,A.ID,F.名称,A.编码,Nvl(G.名称,A.名称),E.计算单位,A.规格,A.产地,D.药品剂型,A.费用类型,A.说明,D.处方职务,A.是否变价," & _
                            "C." & str范围 & "包装,C." & str范围 & "单位" & IIF(strStock <> "", ",X.库存", "") & IIF(blnOften, ",Nvl(R.频度,1)", "")
                Else
                    strSQL = _
                        " Select A.类别 AS 类别ID,E.ID as 诊疗项目ID,A.ID as 收费细目ID," & _
                            " F.名称 AS 类别,A.编码,Nvl(G.名称,A.名称) as 名称,E.计算单位,A.规格,A.产地,D.药品剂型,NULL AS 项目特性,A.费用类型,A.说明,D.处方职务 as 处方职务ID" & _
                            IIF(strStock <> "", ",Decode(X.库存,NULL,NULL,LTrim(To_Char(X.库存/C." & str范围 & "包装,'999990.0000'))||C." & str范围 & "单位) as 库存", "") & _
                            IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                        " From 收费项目目录 A,药品规格 C,药品特性 D,诊疗项目目录 E,收费项目类别 F,收费项目别名 G" & _
                            IIF(blnOften, ",诊疗个人项目 R", "") & IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        " Where A.ID=C.药品ID And C.药名ID=D.药名ID And D.药名ID=E.ID" & _
                            " And A.类别=F.编码 And E.类别 IN('5','6','7')" & _
                            " And A.ID=G.收费细目ID(+) And G.码类(+)=1 And G.性质(+)=" & IIF(gbln商品名, 3, 1) & _
                            " And E.服务对象 IN([8],3) And Nvl(E.执行频率,0) IN(0,[9])" & _
                            " And A.类别 IN ('5','6','7') And A.服务对象 IN([8],3)" & _
                            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                            " And (E.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or E.撤档时间 IS NULL)" & _
                            IIF(blnOften, " And R.诊疗项目ID=E.ID And R.人员ID=[11]", "") & _
                            IIF(strStock <> "", str库存限制, "") & str药品 & Replace(strSub, "A.", "E.")
                End If
            End If
        End If
        
        '非药品诊疗项目部分:分类不是药品类型时不必读取
        blnLoad = False
        If mstr输入 <> "" Or (blnOften And mlng分类ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) = 0
        End If
        If blnLoad Then
            If mstr输入 <> "" Then
                strInput = " And A.类别<>'4' And (A.编码 Like [5] Or B.名称 Like [6] Or B.简码 Like [6]) And B.码类=[7]"
                If IsNumeric(mstr输入) Then
                    '1X.输入全是数字时只匹配编码
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.类别<>'4' And A.编码 Like [5] And B.码类=[7]"
                ElseIf zlCommFun.IsCharAlpha(mstr输入) Then
                    'X1.输入全是字母时只匹配简码
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.类别<>'4' And B.简码 Like [6] And B.码类=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                    '包含汉字,则只匹配名称
                    strInput = " And A.类别<>'4' And B.名称 Like [6] And B.码类=[7]"
                End If
            
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & IIF(Not mbln简码, "Distinct", "") & _
                        " A.类别 As 类别ID,A.ID as 诊疗项目ID,-NULL as 收费细目ID," & _
                        " D.名称 As 类别,A.编码,B.名称," & IIF(mbln简码, "B.简码,", "") & _
                        " A.计算单位,A.标本部位 as 规格,NULL AS 产地,NULL as 药品剂型," & _
                        str操作类型 & " As 项目特性,NULL as 费用类型,Null as 说明,Null as 处方职务ID" & _
                        IIF(mbln价格, ",NULL as 价格", "") & IIF(strStock <> "", ",NULL As 库存", "") & _
                        IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                    " From 诊疗项目类别 D,诊疗项目别名 B,诊疗项目目录 A" & IIF(blnOften, ",诊疗个人项目 R", "") & _
                    " Where A.ID=B.诊疗项目ID And A.类别=D.编码 And A.类别 Not IN('5','6','7')" & _
                        " And (A.类别<>'9' Or A.人员ID=[11] Or A.人员ID is Null)" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                        " And A.服务对象 IN([8],3) And Nvl(A.单独应用,0)=1 And Instr([10],','||Nvl(A.适用性别,0)||',')>0" & _
                        " And Nvl(A.执行频率,0) IN(0,[9])" & strInput & strSub & _
                        IIF(blnOften, " And R.诊疗项目ID=A.ID And R.人员ID=[11]", "")
            Else
                '从树形分类中选择的非药品诊疗项目,不可能与药品同时显示,因此可以减少一些字段
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & _
                        " A.类别 as 类别ID,A.ID as 诊疗项目ID,-NULL as 收费细目ID,D.名称 as 类别," & _
                        " A.编码,A.名称,A.计算单位,A.标本部位 as 规格" & IIF(blnOften, ",Null as 产地,Null as 药品剂型", "") & "," & _
                        str操作类型 & " as 项目特性" & IIF(blnOften, ",Null as 费用类型,Null as 说明,Null as 处方职务ID" & _
                        IIF(mbln价格, ",NULL as 价格", "") & IIF(strStock <> "", ",Null as 库存", ""), "") & _
                        IIF(blnOften, ",Nvl(R.频度,1) as 频度ID", "") & _
                    " From 诊疗项目类别 D,诊疗项目目录 A" & IIF(blnOften, ",诊疗个人项目 R", "") & _
                    " Where A.类别=D.编码 And A.类别 Not IN('5','6','7')" & _
                        " And (A.类别<>'9' Or A.人员ID=[11] Or A.人员ID is Null)" & _
                        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                        " And A.服务对象 IN([8],3) And Nvl(A.单独应用,0)=1 And Instr([10],','||Nvl(A.适用性别,0)||',')>0" & _
                        " And Nvl(A.执行频率,0) IN(0,[9])" & strSub & _
                        IIF(blnOften, " And R.诊疗项目ID=A.ID And R.人员ID=[11]", "")
            End If
        End If
    End If
    If blnOften Then
        strSQL = "Select Rownum as KeyID,A.* From (" & strSQL & ") A Order by 频度ID Desc,Decode(类别ID,'4','I',类别ID),类别,编码" '无序号,排在护理之后
    Else
        strSQL = "Select Rownum as KeyID,A.* From (" & strSQL & ") A Order by Decode(类别ID,'4','I',类别ID),类别,编码" '无序号,排在护理之后
    End If
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int类型, str类别, lng分类ID, mlng分类ID, _
        UCase(mstr输入) & "%", mstrLike & UCase(mstr输入) & "%", mint简码 + 1, mint范围, _
        IIF(mint期效 = 0, 2, 1), "," & str性别 & ",", UserInfo.ID, lng药房ID, mlng西药房, mlng成药房, mlng中药房)
    
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
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
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
    LockWindowUpdate 0
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
        "'E',Decode(A.操作类型,'1','过敏试验','2','给药途径','3','中药煎法','4','中药用法',Null)," & _
        "'Z',Decode(A.操作类型,'1','留观','2','住院','3','转科','4','术后','5','出院','6','转院',NULL)," & _
        "A.操作类型)"
    
    '只统计该分类下面的常用项目
    If mlng分类ID <> 0 Then
        str分类 = " And A.分类ID IN(Select ID From 诊疗分类目录 Start With ID=[1] Connect by Prior ID=上级ID)"
        
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
    strSQL = _
        " Select A.诊疗项目ID,Count(A.诊疗项目ID) As 次数" & _
        " From 病人医嘱记录 A,病人医嘱状态 B" & _
        " Where A.ID=B.医嘱ID And B.操作类型=1" & str类别 & _
        "   And B.操作时间>=[3] And B.操作人员=[4]" & _
        " Group By A.诊疗项目ID"
    If MovedByDate(dtpDate.Value) Then
        strSQL = strSQL & " Union ALL " & _
            " Select A.诊疗项目ID,Count(A.诊疗项目ID) As 次数" & _
            " From H病人医嘱记录 A,H病人医嘱状态 B" & _
            " Where A.ID=B.医嘱ID And B.操作类型=1" & str类别 & _
            "   And B.操作时间>=[3] And B.操作人员=[4]" & _
            " Group By A.诊疗项目ID"
    End If
    
    If InStr(str类别, "='5'") > 0 Or InStr(str类别, "='6'") > 0 Or InStr(str类别, "='7'") > 0 Then
        '药品诊疗项目部分
        strSQL = _
            " Select A.类别 as 类别ID,A.ID as 诊疗项目ID,NULL as 收费细目ID," & _
            " D.名称 as 类别,A.编码,A.名称,A.计算单位,C.药品剂型,B.次数" & _
            " From 诊疗项目类别 D,药品特性 C,诊疗项目目录 A,(" & strSQL & ") B" & _
            " Where A.类别=D.编码 And A.ID=C.药名ID And A.ID=B.诊疗项目ID" & _
            "   And Not (A.类别='E' And Nvl(A.操作类型,'0')<>'0')" & str分类 & _
            "   And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            "   And Nvl(A.服务对象,0)<>0 And (Nvl(A.单独应用,0)=1 Or A.类别='7')"
        If chkAll.Value = 0 Then
            strSQL = _
                "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,计算单位,药品剂型,-1*次数 as 次数 From (" & strSQL & ")" & _
                " Group by -1*次数,类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,计算单位,药品剂型"
            strSQL = "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,计算单位,药品剂型,Abs(次数) as 次数 From (" & strSQL & ") Where Rownum<=[5]"
        End If
    Else
        '非药品部份或所有部份
        strSQL = _
            " Select A.类别 as 类别ID,A.ID as 诊疗项目ID,NULL as 收费细目ID," & _
            " D.名称 as 类别,A.编码,A.名称,A.计算单位,A.标本部位," & str操作类型 & " As 项目特性,B.次数" & _
            " From 诊疗项目类别 D,诊疗项目目录 A,(" & strSQL & ") B" & _
            " Where A.类别=D.编码 And A.ID=B.诊疗项目ID" & str分类 & _
            "   And Not (A.类别='E' And Nvl(A.操作类型,'0')<>'0')" & _
            "   And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            "   And Nvl(A.服务对象,0)<>0 And (Nvl(A.单独应用,0)=1 Or A.类别='7')"
        If chkAll.Value = 0 Then
            strSQL = _
                "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,计算单位,标本部位,项目特性,-1*次数 as 次数 From (" & strSQL & ")" & _
                " Group by -1*次数,类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,计算单位,标本部位,项目特性"
            strSQL = "Select 类别ID,诊疗项目ID,收费细目ID,类别,编码,名称,计算单位,标本部位,项目特性,Abs(次数) as 次数 From (" & strSQL & ") Where Rownum<=[5]"
        End If
    End If
    strSQL = "Select Rownum as KeyID,Null as 选择,A.* From (" & strSQL & ") A Order by 次数 Desc,编码"
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng分类ID, str选择类别, CDate(Format(dtpDate.Value, "yyyy-MM-dd")), UserInfo.姓名, Val(txtCount.Text))
    
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
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
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
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function
