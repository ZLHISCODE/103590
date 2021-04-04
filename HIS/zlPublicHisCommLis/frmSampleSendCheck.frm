VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSampleSendCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "送检核对"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10470
   Icon            =   "frmSampleSendCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame FraAdvice 
      Caption         =   "医嘱信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5445
      Left            =   60
      TabIndex        =   12
      Top             =   1680
      Width           =   10365
      Begin VB.CheckBox chkAll 
         Caption         =   "全选"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8610
         TabIndex        =   15
         Top             =   420
         Width           =   870
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1575
         TabIndex        =   4
         Top             =   375
         Width           =   3090
      End
      Begin VB.ComboBox cboCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   375
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4470
         Left            =   60
         TabIndex        =   13
         Top             =   840
         Width           =   10260
         _cx             =   18098
         _cy             =   7885
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         Editable        =   2
         ShowComboButton =   0
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
      Begin VB.Label lblRefresh 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "清除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F56C58&
         Height          =   240
         Left            =   9615
         MouseIcon       =   "frmSampleSendCheck.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   405
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   5265
         TabIndex        =   14
         Top             =   405
         Width           =   2955
      End
   End
   Begin VB.Frame fraPerson 
      Caption         =   "员工信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   10365
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3255
         TabIndex        =   2
         Top             =   945
         Width           =   720
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   855
         TabIndex        =   1
         Top             =   945
         Width           =   1605
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   855
         MaxLength       =   20
         TabIndex        =   0
         Top             =   405
         Width           =   3120
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9135
         TabIndex        =   5
         Top             =   405
         Width           =   1035
      End
      Begin VB.CommandButton cmdCnacel 
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9135
         TabIndex        =   6
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   330
         TabIndex        =   11
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2715
         TabIndex        =   10
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   465
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7125
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSampleSendCheck.frx":0BD4
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13864
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   9750
      Top             =   -270
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
            Picture         =   "frmSampleSendCheck.frx":1468
            Key             =   "紧急"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSampleSendCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCbo As Boolean
Private mintSendCount As Integer    '当前人员送检标本数
Private mintCurrentCount As Integer    '当前标本数
Private mintCheckCount As Integer    '勾选标本数
Private mintDays As Integer  '允许查询天数
Private mstrSDate As String  '开始时间
Private mstrEDate As String  '结束时间
Private mstrAdvice As String    '已核对医嘱

Private Sub cboCode_Click()
    If mblnCbo Then txtCode.SetFocus
End Sub

Private Sub chkAll_Click()
    Dim i As Integer
    With vsfList
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                .Cell(flexcpChecked, i, .ColIndex("选择"), i, .ColIndex("选择")) = chkAll.value
            Next
        End If
    End With
End Sub

Private Sub cmdCnacel_Click()
    Unload Me
End Sub

Public Function ShowMe(ByVal frmParent As Form, intType As Integer, intDays As Integer) As Boolean
    mintDays = intDays
    Me.Show vbModal, frmParent
End Function

Private Sub cmdOK_Click()
          Dim lngRow As Long

1         On Error GoTo cmdOK_Click_Error

2         If txtInfo(0).Tag = "" Then
3             MsgBox "请先确定送检员工信息！", vbInformation, "中联信息"
4             txtInfo(0).SetFocus
5             Exit Sub
6         End If

7         If mintCurrentCount = 0 Then
8             MsgBox "请先扫描要本次要送检核对的标本！", vbInformation, "中联信息"
9             txtCode.SetFocus
10            Exit Sub
11        End If

12        If SaveSampleNum Then
13            MsgBox "送检核对成功！", vbInformation, "中联信息"
14            mintSendCount = 0
15            mintCheckCount = 0

16            With vsfList
17                For lngRow = .Rows - 1 To 1 Step -1
18                    If .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 1 Then
19                        .RemoveItem lngRow
20                    End If
21                Next
22            End With

23            txtCode.Text = ""
24            txtCode.SetFocus
25        End If

26        Call ShowInfo


27        Exit Sub
cmdOK_Click_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "执行(cmdOK_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
29        Err.Clear
End Sub

Private Function SaveSampleNum() As Boolean
      '记录检验核对数量
          Dim strSQL As String
          Dim rsTemp As Recordset
          Dim strAdvice As String
          Dim strYes As String
          Dim strNO As String
          Dim lngRow As Long
          Dim strMsg As String
          Dim intCount As Integer
          Dim strBatchNO As String
          Dim rsSampleCode As ADODB.Recordset
          Dim strSendAdivce As String
          Dim blnTre As Boolean
          Dim strErr As String

1         On Error GoTo SaveSampleNum_Error

2         lblInfo.Caption = ""

          '获取医嘱内容
3         With vsfList
4             If .Rows > 1 Then
5                 For lngRow = 1 To .Rows - 1
6                     If .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 1 Then
7                         strAdvice = strAdvice & ";" & .TextMatrix(lngRow, .ColIndex("医嘱id")) & "^" & .TextMatrix(lngRow, .ColIndex("试管编码"))
8                         strSendAdivce = strSendAdivce & "," & .TextMatrix(lngRow, .ColIndex("id")) & "," & .TextMatrix(lngRow, .ColIndex("医嘱id"))
9                     End If
10                Next
11            End If
12        End With

13        If strAdvice <> "" Then
14            strAdvice = Mid(strAdvice, 2)
15            strSendAdivce = Mid(strSendAdivce, 2)
16        Else
17            MsgBox "请选择要核对的记录！", vbInformation, "中联信息"
18            Exit Function
19        End If


          '采集工作站
20        strSQL = "Select 人员id, 登记数量, 登记项目 From 标本送检记录 Where 核对时间 Is Null And 人员id = [1] And 登记时间 Between [2] And [3]"
21        Set rsTemp = ComOpenSQL(Sel_Lis_DB, strSQL, "标本送检记录", Val(txtInfo(0).Tag), CDate(mstrSDate), CDate(mstrEDate))

22        If rsTemp.EOF Then
23            strSQL = "Zl_标本送检记录_Edit(1," & Val(txtInfo(0).Tag) & ",'" & Trim(txtInfo(1).Text) & "'," & mintCheckCount & ",'" & strAdvice & "',To_Date('" & mstrSDate & "','yyyy-mm-dd hh24:mi:ss'),To_Date('" & mstrEDate & "','yyyy-mm-dd hh24:mi:ss'))"
24            Call ComExecuteProc(Sel_Lis_DB, strSQL, "标本送检记录")
25            SaveSampleNum = True
26            SaveDBLog 18, 6, 0, "送检核对", "送检标本登记,登记人:" & Trim(txtInfo(1).Text), 1018, "标本签收"

27            gcnHisOracle.BeginTrans
              '核对之后送检标本
              '生成发送批号
              '提交老版LIS数据
28            strSQL = "select 病人医嘱发送_标本发送批号.NEXTVAL from dual"
29            Set rsSampleCode = ComOpenSQL(Sel_His_DB, strSQL, "标本发送批号", "")
30            strBatchNO = rsSampleCode(0) & ""
31            strSQL = "Zl_Lis预置条码_标本送出('" & strSendAdivce & "',0,'" & txtInfo(1).Text & "','" & strBatchNO & "')"
32            Call ComExecuteProc(Sel_His_DB, strSQL, "标本送检")

              '提交新版LIS数据
33            If funSampleSendInfo(strSendAdivce, 0, txtInfo(1).Text, strErr) = False Then
34                gcnHisOracle.RollbackTrans
35                If strErr <> "" Then
36                    MsgBox strErr, vbInformation, "送检核对"
37                End If
38                Exit Function
39            End If
40            gcnHisOracle.CommitTrans
41            blnTre = False

42        Else
43            MsgBox "人员【" & Trim(txtInfo(1).Text) & "】存在未核对的送检记录，请先核对！", vbInformation, "中联信息"
44            SaveSampleNum = False
45        End If


46        Exit Function
SaveSampleNum_Error:
47        If blnTre Then gcnHisOracle.RollbackTrans
48        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "执行(SaveSampleNum)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
49        Err.Clear
End Function

Private Function CheckAdvice(ByVal strOldAdvice As String, ByVal strNewAdvice As String, strMsg As String, Optional blnModify As Boolean = False) As Boolean
      '核对医嘱内容
          Dim arrOld As Variant
          Dim arrNew As Variant
          Dim strOld As String
          Dim strNew As String
          Dim strTemp As String
          Dim i As Integer

1         On Error GoTo CheckAdvice_Error

2         If strNewAdvice = "" Or strOldAdvice = "" Then Exit Function

3         arrNew = Split(strNewAdvice, ";")
4         arrOld = Split(strOldAdvice, ";")

5         If blnModify Then
              '修正核对失败的送检记录
6             For i = 0 To UBound(arrOld)
7                 If InStr(";" & strNewAdvice & ";", ";" & arrOld(i) & ";") > 0 Then
                      '能修正的医嘱
8                     strMsg = strMsg & ";" & arrOld(i)
9                 Else
                      '不能修正的医嘱
10                    strTemp = strTemp & ";" & arrOld(i)
11                End If
12            Next
13            If strMsg = "" Then Exit Function
14            strMsg = Mid(strMsg, 2) & "|" & Mid(strTemp, 2)
15        Else
              '核对送检记录
16            If UBound(arrOld) >= UBound(arrNew) Then
                  '缺少项目
17                For i = 0 To UBound(arrOld)
18                    If InStr(";" & strNewAdvice & ";", ";" & arrOld(i) & ";") > 0 Then
19                    Else
20                        strOld = strOld & ";" & arrOld(i)
21                    End If
22                Next
23                If strOld <> "" Then
24                    strOld = Mid(strOld, 2)
25                    strMsg = strOld & "|"
26                    Call GetMsg(strOld)
27                    strMsg = strMsg & "核对数量" & UBound(arrNew) + 1 & "小于等于送检数量" & UBound(arrOld) + 1 & vbCrLf & vbCrLf & "缺少项目：" & strOld
28                    Exit Function
29                End If
30            Else
                  '多余项目
31                For i = 0 To UBound(arrNew)
32                    If InStr(";" & strOldAdvice & ";", ";" & arrNew(i) & ";") > 0 Then
33                    Else
34                        strNew = strNew & ";" & arrNew(i)
35                    End If
36                Next
37                If strNew <> "" Then
38                    strNew = Mid(strNew, 2)
39                    strMsg = strNew & "|"
40                    Call GetMsg(strNew)
41                    strMsg = strMsg & "核对数量" & UBound(arrNew) + 1 & "大于送检数量" & UBound(arrOld) + 1 & vbCrLf & vbCrLf & "多余项目：" & strNew
42                    Exit Function
43                End If
44            End If
45        End If

46        CheckAdvice = True


47        Exit Function
CheckAdvice_Error:
48        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "执行(CheckAdvice)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
49        Err.Clear
End Function

Private Sub GetMsg(strMsg As String)
      '获取核对有误的医嘱信息
          Dim strSQL As String
          Dim rsTemp As Recordset
          Dim str医嘱内容 As String
          Dim str紧急 As String
          Dim str病人id As String
          Dim str标本类型 As String
          Dim str样本条码 As String
          Dim str开嘱科室id As String
          Dim str执行科室id As String
          Dim str试管编码 As String
          Dim arrAdvice As Variant
          Dim i As Integer

1         On Error GoTo GetMsg_Error

2         If strMsg = "" Then Exit Sub

3         arrAdvice = Split(strMsg, ";")
4         strMsg = ""

5         For i = 0 To UBound(arrAdvice)
6             strSQL = "Select Distinct a.病人id, a.相关id 医嘱id, Decode(a.紧急标志, 1, '紧急', '') 紧急, Decode(a.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检') 病人来源," & vbNewLine & _
                     "                a.姓名, a.性别, d.医嘱内容, a.标本部位 标本类型, b.样本条码, b.采样人, b.采样时间, b.送检人, b.标本送出时间 送检时间, a.开嘱科室id, a.执行科室id, c.试管编码" & vbNewLine & _
                       "From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C, 病人医嘱记录 D, (Select Column_Value id From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) E" & vbNewLine & _
                       "Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And a.相关id = d.Id And c.类别 = 'C' And a.相关id Is Not Null And b.接收时间 Is Null And" & vbNewLine & _
                     "      b.执行状态 = 0 And a.相关id = e.id And b.发送时间 Between [2] And [3] And c.试管编码 = [4]"
7             Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "医嘱信息", Split(arrAdvice(i), "^")(0), CDate(mstrSDate), CDate(mstrEDate), Split(arrAdvice(i), "^")(1))

8             Do While Not rsTemp.EOF
9                 If str医嘱内容 <> "" & rsTemp!医嘱内容 And str紧急 = "" & rsTemp!紧急 And str病人id = "" & rsTemp!病人ID And str标本类型 = "" & rsTemp!标本类型 And _
                     str样本条码 = "" & rsTemp!样本条码 And str开嘱科室id = "" & rsTemp!开嘱科室ID And str执行科室id = "" & rsTemp!执行科室id And str试管编码 = "" & rsTemp!试管编码 Then
                      '医嘱合并
10                    strMsg = strMsg & "," & rsTemp!医嘱内容
11                Else
12                    str医嘱内容 = "" & rsTemp!医嘱内容
13                    str紧急 = "" & rsTemp!紧急
14                    str病人id = "" & rsTemp!病人ID
15                    str标本类型 = "" & rsTemp!标本类型
16                    str样本条码 = "" & rsTemp!样本条码
17                    str开嘱科室id = "" & rsTemp!开嘱科室ID
18                    str执行科室id = "" & rsTemp!执行科室id
19                    str试管编码 = "" & rsTemp!试管编码

20                    strMsg = strMsg & vbCrLf & rsTemp!姓名 & "  " & rsTemp!医嘱内容
21                End If

22                rsTemp.MoveNext
23            Loop
24        Next

25        Exit Sub
GetMsg_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "执行(GetMsg)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear
End Sub

Private Sub Form_Load()
    Dim intOnlyBarcode As Integer

    Me.Caption = "标本送检记录"

    intOnlyBarcode = Val(ComGetPara(Sel_Lis_DB, "仅支持条码扫描录入数据", 2500, 1018, 0))
    If intOnlyBarcode = 1 Then cboCode.Enabled = False

    mstrEDate = Format(Currentdate, "yyyy-mm-dd") & " 23:59:59"
    mstrSDate = Format(CDate(mstrEDate) - mintDays, "yyyy-mm-dd") & " 00:00:00"

    Call CreateCbo
    Call GetAdvice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnCbo = False
    mintSendCount = 0
    mintCurrentCount = 0
    mintCheckCount = 0
    mstrAdvice = ""
End Sub

Private Sub lblRefresh_Click()
    vsfList.Rows = 1
    vsfList.Rows = 2
    mintCurrentCount = 0
    mintCheckCount = 0
    lblInfo.Caption = ""
    Call ShowInfo
End Sub

Private Sub txtCode_GotFocus()
    txtCode.SelStart = 0
    txtCode.SelLength = LenB(StrConv(Trim(txtCode.Text), vbFromUnicode))
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtCode.Text) <> "" Then
            Call GetAdvice(Trim(txtCode.Text))
            Call txtCode_GotFocus
        End If
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case cboCode.Text
        Case "条码扫描", "门 诊 号", "住 院 号"
            If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = LenB(StrConv(Trim(txtInfo(Index).Text), vbFromUnicode))
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Index = 0 And KeyCode = 13 Then
        For i = 1 To 2
            txtInfo(i).Text = ""
        Next
        txtInfo(0).Tag = ""
        If Trim(txtInfo(0).Text) <> "" Then
            Call GetPerson(Trim(txtInfo(0).Text))
            Call ShowInfo
        End If
    End If
End Sub

Private Sub ShowInfo()
'提示标本数量
    stbThis.Panels(2).Text = ""
    If mintSendCount <> 0 Then
        stbThis.Panels(2).Text = "人员【" & Trim(txtInfo(1).Text) & "】送检标本数：" & mintSendCount & " "
    End If

    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "当前标本数：" & mintCurrentCount & " "
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "核对标本数：" & mintCheckCount
End Sub

Private Sub GetPerson(ByVal strCode As String)
      '获取病人信息
          Dim rsTemp As Recordset
          Dim strSQL As String

1         On Error GoTo GetPerson_Error

2         mintSendCount = 0

3         strSQL = "Select a.id, a.姓名, a.性别 From 人员表 A Where  a.编号 = [1]"

4         Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "人员信息", strCode)

5         If rsTemp.EOF Then
6             MsgBox "未找到人员！", vbInformation, "中联信息"
7             Call txtInfo_GotFocus(0)
8         Else
9             txtInfo(0).Tag = rsTemp!ID
10            txtInfo(1).Text = "" & rsTemp!姓名
11            txtInfo(2).Text = "" & rsTemp!性别

              '人员送检记录情况
12            strSQL = "Select 人员id, 登记数量, 登记项目 From 标本送检记录 Where 核对时间 Is Null And 人员id = [1] And 登记时间 Between [2] And [3]"
13            Set rsTemp = ComOpenSQL(Sel_Lis_DB, strSQL, "标本送检记录", Val(txtInfo(0).Tag), CDate(mstrSDate), CDate(mstrEDate))

14            If rsTemp.EOF Then

15            Else
16                MsgBox "人员【" & Trim(txtInfo(1).Text) & "】存在未核对的送检记录，请先核对！", vbInformation, "中联信息"
17            End If


18            txtCode.SetFocus
19        End If


20        Exit Sub
GetPerson_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "执行(GetPerson)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
22        Err.Clear
End Sub

Private Sub CreateCbo()
    cboCode.AddItem "条码扫描"
    cboCode.AddItem "门 诊 号"
    cboCode.AddItem "住 院 号"
    cboCode.AddItem "挂 号 单"
    cboCode.ListIndex = 0
    mblnCbo = True
End Sub

Private Sub GetAdvice(Optional ByVal strCode As String)
      '获取医嘱信息
          Dim strSQL As String
          Dim rsTemp As Recordset
          Dim strTitle As String

1         On Error GoTo GetAdvice_Error

2         If cboCode.Text = "条码扫描" Then
3             strSQL = "Select Distinct 1 选择, a.病人id,a.id, a.相关id 医嘱id, Decode(a.紧急标志, 1, '紧急', '') 紧急, Decode(a.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检') 病人来源," & vbNewLine & _
                     "                a.姓名, a.性别, d.医嘱内容, a.标本部位 标本类型, b.样本条码, b.采样人, b.采样时间, b.送检人, b.标本送出时间 送检时间, a.开嘱科室id, a.执行科室id, c.试管编码" & vbNewLine & _
                       "From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C, 病人医嘱记录 D" & vbNewLine & _
                       "Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And a.相关id = d.Id And c.类别 = 'C' And a.相关id Is Not Null And b.标本送出时间 is null And b.接收时间 Is Null And b.采样人 is not null And" & vbNewLine & _
                     "      b.执行状态 = 0 And b.样本条码 = [1] And b.发送时间 Between [2] And [3]"
4         Else
5             strSQL = "Select Distinct 1 选择, a.病人id,a.id, a.相关id 医嘱id, Decode(a.紧急标志, 1, '紧急', '') 紧急, Decode(a.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检') 病人来源," & vbNewLine & _
                     "                a.姓名, a.性别, d.医嘱内容, a.标本部位 标本类型, b.样本条码, b.采样人, b.采样时间, b.送检人, b.标本送出时间 送检时间, a.开嘱科室id, a.执行科室id, c.试管编码" & vbNewLine & _
                       "From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C, 病人医嘱记录 D, 病人信息 E" & vbNewLine & _
                       "Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And a.相关id = d.Id And a.病人id = e.病人id And c.类别 = 'C' And a.相关id Is Not Null And" & vbNewLine & _
                     "       b.标本送出时间  is null And b.接收时间 Is Null And b.采样人 is not null And b.执行状态 = 0 And e.病人id = [1] And b.发送时间 Between [2] And [3]"

6             If cboCode.Text = "门 诊 号" Then
7                 strSQL = Replace(strSQL, "e.病人id = [1]", "e.门诊号 = [1]")
8             ElseIf cboCode.Text = "住 院 号" Then
9                 strSQL = Replace(strSQL, "e.病人id = [1]", "e.住院号 = [1]")
10            ElseIf cboCode.Text = "挂 号 单" Then
11                strSQL = Replace(strSQL, "e.病人id = [1]", "a.挂号单 = [1]")
12            End If
13        End If

14        If vsfList.TextMatrix(0, 1) = "" Then
              '首次进入不取数据，仅列表初始化

15            strSQL = strSQL & " And 1=0"
16            Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "医嘱信息", strCode, CDate(mstrSDate), CDate(mstrEDate))

17            Call vfgLoadFromRecord(vsfList, rsTemp, "", imgList)
18            With vsfList
19                .ExplorerBar = flexExSortShow
20                .ColDataType(.ColIndex("选择")) = flexDTBoolean
21                .ColWidth(.ColIndex("选择")) = 250: .TextMatrix(0, .ColIndex("选择")) = ""
22                .ColWidth(.ColIndex("紧急")) = 250: .ColHidden(.ColIndex("紧急")) = False
23                .ColWidth(.ColIndex("病人来源")) = 1200: .ColHidden(.ColIndex("病人来源")) = False
24                .ColWidth(.ColIndex("姓名")) = 1200: .ColHidden(.ColIndex("姓名")) = False
25                .ColWidth(.ColIndex("性别")) = 600: .ColHidden(.ColIndex("性别")) = False
26                .ColWidth(.ColIndex("医嘱内容")) = 3000: .ColHidden(.ColIndex("医嘱内容")) = False
27                .ColWidth(.ColIndex("标本类型")) = 1200: .ColHidden(.ColIndex("标本类型")) = False
28                .ColWidth(.ColIndex("样本条码")) = 1600: .ColHidden(.ColIndex("样本条码")) = False
29                .ColWidth(.ColIndex("采样人")) = 1200: .ColHidden(.ColIndex("采样人")) = False
30                .ColWidth(.ColIndex("采样时间")) = 1400: .ColHidden(.ColIndex("采样时间")) = False
31                .ColWidth(.ColIndex("送检人")) = 1200: .ColHidden(.ColIndex("送检人")) = False
32                .ColWidth(.ColIndex("送检时间")) = 1400: .ColHidden(.ColIndex("送检时间")) = False

33                mintCurrentCount = 0
34                mintCheckCount = 0
35            End With

              '已核对医嘱
36            strSQL = "Select 核对项目 From 标本送检记录 Where 核对时间 Is Not Null And 登记时间 Between [1] And [2]"
37            Set rsTemp = ComOpenSQL(Sel_Lis_DB, strSQL, "标本送检记录", CDate(mstrSDate), CDate(mstrEDate))
38            mstrAdvice = ""
39            Do While Not rsTemp.EOF
40                mstrAdvice = mstrAdvice & ";" & rsTemp!核对项目
41                rsTemp.MoveNext
42            Loop
43            If mstrAdvice <> "" Then mstrAdvice = Mid(mstrAdvice, 2)
44        Else
45            Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "医嘱信息", strCode, CDate(mstrSDate), CDate(mstrEDate))
46            If rsTemp.EOF Then MsgBox "未找到标本，或标本未采样，或已送检，或已登记！", vbInformation, "中联信息": Exit Sub

47            With vsfList
48                Do While Not rsTemp.EOF
49                    If Not FindAdvice("" & rsTemp!医嘱id, "" & rsTemp!试管编码, "" & rsTemp!样本条码) Then    '判断是否已存在列表中
50                        If .TextMatrix(.Rows - 1, .ColIndex("医嘱id")) <> "" Then
51                            .Rows = .Rows + 1
52                        End If

53                        If .TextMatrix(.Rows - 2, .ColIndex("医嘱内容")) <> "" & rsTemp!医嘱内容 And _
                             .TextMatrix(.Rows - 2, .ColIndex("紧急")) = "" & rsTemp!紧急 And _
                             .TextMatrix(.Rows - 2, .ColIndex("病人id")) = "" & rsTemp!病人ID And _
                             .TextMatrix(.Rows - 2, .ColIndex("标本类型")) = "" & rsTemp!标本类型 And _
                             .TextMatrix(.Rows - 2, .ColIndex("样本条码")) = "" & rsTemp!样本条码 And _
                             .TextMatrix(.Rows - 2, .ColIndex("开嘱科室id")) = "" & rsTemp!开嘱科室ID And _
                             .TextMatrix(.Rows - 2, .ColIndex("执行科室id")) = "" & rsTemp!执行科室id And _
                             .TextMatrix(.Rows - 2, .ColIndex("试管编码")) = "" & rsTemp!试管编码 Then
                              '医嘱合并
54                            .TextMatrix(.Rows - 2, .ColIndex("医嘱id")) = .TextMatrix(.Rows - 2, .ColIndex("医嘱id")) & "," & rsTemp!医嘱id
55                            .TextMatrix(.Rows - 2, .ColIndex("医嘱内容")) = .TextMatrix(.Rows - 2, .ColIndex("医嘱内容")) & "," & rsTemp!医嘱内容
56                            .Rows = .Rows - 1
57                        Else
58                            .TextMatrix(.Rows - 1, .ColIndex("id")) = "" & rsTemp!ID
59                            .TextMatrix(.Rows - 1, .ColIndex("病人id")) = "" & rsTemp!病人ID
60                            .TextMatrix(.Rows - 1, .ColIndex("医嘱id")) = "" & rsTemp!医嘱id
61                            .TextMatrix(.Rows - 1, .ColIndex("紧急")) = "" & rsTemp!紧急
62                            .TextMatrix(.Rows - 1, .ColIndex("病人来源")) = "" & rsTemp!病人来源
63                            .TextMatrix(.Rows - 1, .ColIndex("姓名")) = "" & rsTemp!姓名
64                            .TextMatrix(.Rows - 1, .ColIndex("性别")) = "" & rsTemp!性别
65                            .TextMatrix(.Rows - 1, .ColIndex("医嘱内容")) = "" & rsTemp!医嘱内容
66                            .TextMatrix(.Rows - 1, .ColIndex("标本类型")) = "" & rsTemp!标本类型
67                            .TextMatrix(.Rows - 1, .ColIndex("样本条码")) = "" & rsTemp!样本条码
68                            .TextMatrix(.Rows - 1, .ColIndex("采样人")) = "" & rsTemp!采样人
69                            .TextMatrix(.Rows - 1, .ColIndex("采样时间")) = "" & rsTemp!采样时间
70                            .TextMatrix(.Rows - 1, .ColIndex("送检人")) = "" & rsTemp!送检人
71                            .TextMatrix(.Rows - 1, .ColIndex("送检时间")) = "" & rsTemp!送检时间
72                            .TextMatrix(.Rows - 1, .ColIndex("开嘱科室id")) = "" & rsTemp!开嘱科室ID
73                            .TextMatrix(.Rows - 1, .ColIndex("执行科室id")) = "" & rsTemp!执行科室id
74                            .TextMatrix(.Rows - 1, .ColIndex("试管编码")) = "" & rsTemp!试管编码
75                            .Cell(flexcpChecked, .Rows - 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = "" & rsTemp!选择

76                            If "" & rsTemp!紧急 = "紧急" Then
77                                .Cell(flexcpPicture, .Rows - 1, .ColIndex("紧急"), .Rows - 1, .ColIndex("紧急")) = imgList.ListImages("紧急").ExtractIcon
78                            End If

79                            .TopRow = .Rows - 1
80                            .Row = .Rows - 1
81                        End If
82                    End If

83                    rsTemp.MoveNext
84                Loop
85            End With
86        End If


87        Exit Sub
GetAdvice_Error:
88        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "执行(GetAdvice)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
89        Err.Clear
End Sub

Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    With vsfList
        If .Col = .ColIndex("选择") Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If

        mintCheckCount = 0
        mintCurrentCount = 0
        If .Rows > 1 Then
            If .TextMatrix(1, 1) <> "" Then
                For lngRow = 1 To .Rows - 1
                    mintCurrentCount = mintCurrentCount + 1
                    If .Cell(flexcpChecked, lngRow, 0, lngRow, 0) = 1 Then
                        mintCheckCount = mintCheckCount + 1
                    End If
                Next
            End If
        End If

        Call ShowInfo
    End With
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 And Row > 0 Then
        If vsfList.TextMatrix(Row, 1) <> "" Then
            If vsfList.Cell(flexcpChecked, Row, 0, Row, 0) = 1 Then
                mintCheckCount = mintCheckCount + 1
            Else
                mintCheckCount = mintCheckCount - 1
            End If
        End If
    End If
    Call ShowInfo
End Sub

Private Sub vsfList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        '键盘按钮Delete可删除
        With vsfList
            If .TextMatrix(.Row, .ColIndex("医嘱id")) <> "" And .Rows > 1 Then
                .RemoveItem .Row
                Call vsfList_AfterRowColChange(0, 1, 0, 1)
            End If
        End With
    End If
End Sub

Private Function FindAdvice(ByVal strAdvice As String, ByVal strNO As String, ByVal strCodeBar As String) As Boolean
      '判断标本是否已在列表中，是否已核对过
          Dim lngRow As Long
          Dim arrAdvice As Variant
          Dim strTemp As String

1         On Error GoTo FindAdvice_Error

2         If mstrAdvice <> "" Then
3             arrAdvice = Split(mstrAdvice, ";")
4             For lngRow = 0 To UBound(arrAdvice)
5                 strTemp = arrAdvice(lngRow)
6                 If InStr("," & Split(strTemp, "^")(0) & ",", "," & strAdvice & ",") > 0 And Split(strTemp, "^")(1) = strNO Then
7                     FindAdvice = True
8                     Exit Function
9                 End If
10            Next
11        End If

12        With vsfList
13            If .Rows > 1 Then
14                For lngRow = 1 To .Rows - 1
15                    If InStr("," & .TextMatrix(lngRow, .ColIndex("医嘱id")) & ",", "," & strAdvice & ",") > 0 And .TextMatrix(lngRow, .ColIndex("试管编码")) = strNO And .TextMatrix(lngRow, .ColIndex("样本条码")) = strCodeBar Then
16                        FindAdvice = True
17                        .Row = lngRow
18                        Exit Function
19                    End If
20                Next
21            End If
22        End With


23        Exit Function
FindAdvice_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "执行(FindAdvice)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
25        Err.Clear
End Function

