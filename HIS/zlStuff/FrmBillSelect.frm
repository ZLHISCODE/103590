VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmBillSelect 
   Caption         =   "单据选择器"
   ClientHeight    =   6240
   ClientLeft      =   1548
   ClientTop       =   1896
   ClientWidth     =   10800
   Icon            =   "FrmBillSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10800
   StartUpPosition =   2  '屏幕中心
   Tag             =   "82"
   Begin VB.CommandButton cmd刷新 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   9660
      TabIndex        =   15
      Top             =   95
      Width           =   1100
   End
   Begin VB.CommandButton cmdDeptSel 
      Caption         =   "…"
      Height          =   300
      Left            =   7620
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   270
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Left            =   5940
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   5565
      TabIndex        =   12
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   6675
      TabIndex        =   11
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   8
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton Cmd保存 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8190
      TabIndex        =   7
      Top             =   5430
      Width           =   1100
   End
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9330
      TabIndex        =   6
      Top             =   5415
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp开始日期 
      Height          =   285
      Left            =   930
      TabIndex        =   3
      Top             =   128
      Width           =   1665
      _ExtentX        =   2942
      _ExtentY        =   508
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   111214595
      CurrentDate     =   36734
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   2595
      Left            =   30
      TabIndex        =   0
      Top             =   2730
      Width           =   10785
      _ExtentX        =   19029
      _ExtentY        =   4572
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtp结束日期 
      Height          =   285
      Left            =   3780
      TabIndex        =   4
      Top             =   128
      Width           =   1665
      _ExtentX        =   2942
      _ExtentY        =   508
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   111214595
      CurrentDate     =   36734
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   5880
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "FrmBillSelect.frx":0E42
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14012
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHead 
      Height          =   2055
      Left            =   30
      TabIndex        =   9
      Top             =   480
      Width           =   10785
      _ExtentX        =   19029
      _ExtentY        =   3620
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label LblDepartment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "部门"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5565
      TabIndex        =   5
      Top             =   180
      Width           =   360
   End
   Begin VB.Label Lbl结束日期 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2970
      TabIndex        =   2
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Lbl开始日期 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   720
   End
   Begin VB.Image ImgLine_S 
      Height          =   45
      Left            =   30
      MousePointer    =   7  'Size N S
      Top             =   2670
      Width           =   10755
   End
End
Attribute VB_Name = "FrmBillSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnBootUp As Boolean '启动成功标志
Private mblnFirst As Boolean
Private mstrFind As String '查找条件
Private mstrUserPart As String '部门条件
Private mintLastRow As Integer '上一行
Private mblnOpenCheckCbo As Boolean '是否在输入部门编码
Private mlngSelectCount As Long
Private mstrStart As String
Private mstrEnd As String
Private Const mlngModule = 1724
Private mintUnit As Integer                 '0：散装单位；1：包装单位
Private mblnSuccess As Boolean
Private mstr卫材分类 As String
Private mint计划类型 As Integer
Private mlng库房id As Long
Private mstrSelectNO As String
Private mstrStartDate As String, mstrEndDate As String
Private msngOldY As String
Private Const mstrCaption As String = "单据选择器"

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/12/27
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString

'----------------------------------------------------------------------------------------------------------

Public Function ShowCard(ByVal str材料分类 As String, ByVal lng库房ID As Long, _
        ByVal int计划类型 As Integer, ByRef strSelectNo As String, _
        ByRef strStartDate As String, ByRef strEndDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:对指定申购单进行选择.
    '参数:str材料分类-分类条件(以ID分离为准)
    '     int计划类型-计划类型
    '出参:strSelectNo-初选择的单据号
    '     strStartDate-选择的开始日期
    '     strEndDate-选择的结束日期
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '修改:2007/12/26
    '---------------------------------------------------------------------------------------------------------
    mlng库房id = lng库房ID
    mblnSuccess = False
    mstr卫材分类 = str材料分类
    mint计划类型 = int计划类型
    mstrSelectNO = ""
    
    Me.Show vbModal
    
    ShowCard = mblnSuccess
    strSelectNo = mstrSelectNO
    strStartDate = mstrStartDate
    strEndDate = mstrEndDate
End Function
  
Private Sub cmdAllCls_Click()
    Dim intRow As Integer, intCol As Integer
    
    intCol = GetCol(mshHead, "选择")
    
    mlngSelectCount = 0
    With mshHead
          For intRow = 1 To .Rows - 1
              If Trim(.TextMatrix(intRow, 0)) <> "" Then
                  .TextMatrix(intRow, intCol) = ""
              End If
          Next
      End With
   If mlngSelectCount = 0 Then
        Cmd保存.Enabled = False
    Else
        Cmd保存.Enabled = True
    End If
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer, intCol As Integer
    intCol = GetCol(mshHead, "选择")

    mlngSelectCount = 0
    With mshHead
          For intRow = 1 To .Rows - 1
              If Trim(.TextMatrix(intRow, 0)) <> "" Then
                  .TextMatrix(intRow, intCol) = "√"
                  mlngSelectCount = mlngSelectCount + 1
              End If
          Next
    End With
   If mlngSelectCount = 0 Then
        Cmd保存.Enabled = False
    Else
        Cmd保存.Enabled = True
    End If
End Sub

Private Sub cmdDeptSel_Click()
    If Select部门("") = False Then Exit Sub
    OS.PressKey vbKeyTab
 
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, 4)
End Sub

Private Sub Cmd保存_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    intCol = GetCol(mshHead, "选择")
    mstrSelectNO = ""
    With mshHead
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, intCol) <> "" Then
                mstrSelectNO = IIf(mstrSelectNO = "", "", mstrSelectNO & ",") _
                    & "'" & .TextMatrix(intRow, 0) & "'"
            End If
        Next
    End With
    mstrStartDate = Format(dtp开始日期.Value, "yyyy-mm-dd")
    mstrEndDate = Format(dtp结束日期.Value, "yyyy-mm-dd")
    mblnSuccess = True
    Unload Me
End Sub

Private Sub Cmd取消_Click()
    mblnSuccess = False
    mstrStartDate = "1991-01-01"
    mstrEndDate = "1991-01-01"
    Unload Me
End Sub

Private Sub cmd刷新_Click()
    Call GetList
End Sub

Private Sub Dtp结束日期_Change()
    If Me.dtp结束日期.Value < Me.dtp开始日期.Value Then Me.dtp结束日期.Value = Me.dtp开始日期.Value
    mstrEnd = Format(Me.dtp结束日期.Value, "yyyy-MM-dd")
End Sub
Private Sub Dtp开始日期_Change()
    If Me.dtp开始日期.Value > Me.dtp结束日期.Value Then Me.dtp开始日期.Value = Me.dtp结束日期.Value
    mstrStart = Format(Me.dtp开始日期.Value, "yyyy-MM-dd")
End Sub

Private Sub Form_Activate()
    
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub
Private Function Select部门(ByVal strSeach As String) As Boolean
    '--------------------------------------------------------------------------------------------
    '功能:选择部门
    '参数:strKey-输入条件
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/12/26
    '--------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    
    Set objCtl = txtDept
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSeach)
    strTittle = "申购部门选择"
    If strSeach = "" Then
        gstrSQL = "" & _
            "   Select ID,上级ID,编码,名称,简码,位置,to_char(建档时间,'yyyy-mm-dd') as 建档时间" & _
            "   From 部门表 " & _
            "   Where to_char(撤档时间,'yyyy-MM-dd')='3000-01-01' and (站点=[1] or 站点 is null) " & _
            "   start with 上级id is null connect by prior id=上级id "
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, gstrNodeNo)
    Else
        gstrSQL = "" & _
            "   Select ID,上级ID,编码,名称,简码,位置,to_char(建档时间,'yyyy-mm-dd') as 建档时间" & _
            "   From 部门表 " & _
            "   Where to_char(撤档时间,'yyyy-MM-dd')='3000-01-01' " & _
            "         and  (名称 like [1] or 编码  like [1] or  简码  like  [1]) and (站点=[2] or 站点 is null) " & _
            "   order by 编码"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, gstrNodeNo)
    End If
    
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "没有满足条件的申购部门,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    objCtl.Text = zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!名称)
    objCtl.Tag = zlStr.Nvl(rsTemp!Id)
     If objCtl.Enabled Then objCtl.SetFocus
    Select部门 = True
End Function
Private Sub Form_Load()
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    mintUnit = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    
    mblnBootUp = False
    Select Case mint计划类型
        Case 1
            mstrStart = Format(GetFirstDate(FirstDayOfMonth, sys.Currentdate), "yyyy-mm-dd")
        Case 2
            mstrStart = Format(GetFirstDate(FirstDayOfQuarter, sys.Currentdate), "yyyy-mm-dd")
        Case 3
            mstrStart = Format(GetFirstDate(FirstDayOfyear, sys.Currentdate), "yyyy-mm-dd")
        Case 4
            mstrStart = Format(GetFirstDate(FirstDayOfWeek, sys.Currentdate), "yyyy-mm-dd")
    End Select
    mblnFirst = True
    Me.dtp开始日期 = mstrStart
    Me.dtp开始日期.MaxDate = sys.Currentdate
    mstrEnd = Format(sys.Currentdate, "yyyy-MM-dd")
    Me.dtp结束日期 = sys.Currentdate
    Me.dtp结束日期.MaxDate = sys.Currentdate
    mblnBootUp = True
    Call SetDetal
    Call GetList
    RestoreWinState Me, App.ProductName, mstrCaption
    mblnFirst = False
End Sub

Public Function SetColWidth()
    Dim intCol As Integer
    With mshHead
                
        For intCol = 0 To .Cols - 1
            .ColAlignment(intCol) = flexAlignLeftCenter
            .ColAlignmentFixed(intCol) = 4
            If mblnFirst Then
                  .ColWidth(intCol) = 1000
            End If
        Next
        If mblnFirst Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1000
            .ColWidth(2) = 1200
            .ColWidth(3) = 2000
            .ColWidth(4) = 1000
            .ColWidth(6) = 1000
            
            .ColWidth(9) = 500
        End If
        .ColAlignment(9) = flexAlignCenterCenter
        
        .ColAlignment(GetCol(mshHead, "采购金额")) = flexAlignRightCenter
    End With
    
End Function

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height < 6380 Then
        Me.Height = 6380
    End If
    If Me.Width < 10000 Then
        Me.Width = 10000
    End If
    
    If Me.ImgLine_S.Top >= Me.ScaleHeight - cmdAllCls.Height * 3 - 100 Then
        Me.ImgLine_S.Top = Me.ScaleHeight - cmdAllCls.Height * 3 - 100
    End If
    If Me.ImgLine_S.Top <= 2000 Then
        Me.ImgLine_S.Top = 2000
    End If
    
    With cmd刷新
        .Left = Me.ScaleWidth - .Width - 50
    End With
    With ImgLine_S
        .Left = 0
        .Width = Me.ScaleWidth
    End With
   
    With mshHead
        .Left = 50
        .Width = Me.ScaleWidth - 100
        .Height = ImgLine_S.Top - .Top
    End With
    
    With Cmd取消
        .Top = Me.ScaleHeight - .Height - IIf(stbThis.Visible, stbThis.Height, 0) - 100
        .Left = Me.ScaleWidth - .Width - 50
        Cmd保存.Top = .Top
        Cmd保存.Left = .Left - Cmd保存.Width - 50
        cmdAllCls.Top = .Top
        cmdAllCls.Left = Cmd保存.Left - cmdAllCls.Width * 2
        cmdAllSel.Top = .Top
        cmdAllSel.Left = cmdAllCls.Left - cmdAllSel.Width - 50
        cmdHelp.Top = .Top
    End With
    
    With mshDetail
        .Top = ImgLine_S.Top + ImgLine_S.Height + 100
        .Left = 50
        If Cmd保存.Top - .Top - 50 <= 0 Then
            .Height = 0
        Else
            .Height = Cmd保存.Top - .Top - 50
        End If
        .Width = mshHead.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrCaption
End Sub

 Private Sub ImgLine_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
        msngOldY = y
End Sub

Private Sub ImgLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With ImgLine_S
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - msngOldY
    End With
    Call Form_Resize
End Sub

Private Sub ImgLine_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
        msngOldY = 0
End Sub

Private Sub GetList()
    If mblnBootUp = False Then Exit Sub
    Dim rsList As New Recordset
    Dim strSQL As String
    Dim lng部门ID As Long
    Dim dtStartDate As Date, dtEndDate As Date
    dtStartDate = Format(dtp开始日期.Value, "yyyy-mm-dd")
    dtEndDate = CDate(Format(dtp结束日期.Value, "yyyy-mm-dd") & " 23:59:59")
    lng部门ID = Val(txtDept.Tag)
    
    On Error GoTo ErrHandle
    If mstr卫材分类 <> "" Then
        strSQL = "" & _
            " Select /*+ Rule*/ distinct A.NO,B.名称 as 部门,decode(计划类型,1,'月度计划',2,'季度计划','年度计划') as 计划类型,rtrim(ltrim(to_char(Sum(nvl(c.金额,0))," & mOraFMT.FM_金额 & "))) as 采购金额," & _
            "       A.编制说明,A.编制人 as 填制人,to_char(A.编制日期,'yyyy-MM-dd') as 填制日期,A.审核人," & _
            "       to_char(A.审核日期,'yyyy-MM-dd') as 审核日期,'' as 选择 " & _
            " From  材料采购计划 A,部门表 B,材料计划内容 c,材料特性 d,诊疗项目目录 M," & _
            "       Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) J " & _
            " Where A.部门id=B.id and a.id=c.计划id and c.材料id=d.材料id and (b.站点=[6] or b.站点 is null)" & _
            "       And A.部门ID is not NULL  and a.单据=1 And d.诊疗id=M.id and (M.站点=[6] or M.站点 is null) and M.分类id=J.Column_Value"
    Else
        strSQL = "" & _
            " Select distinct A.NO,B.名称 as 部门, decode(计划类型,1,'月度计划',2,'季度计划','年度计划') as 计划类型,rtrim(ltrim(to_char(Sum(nvl(c.金额,0))," & mOraFMT.FM_金额 & "))) as 采购金额," & _
            "       A.编制说明,A.编制人 as 填制人,to_char(A.编制日期,'yyyy-MM-dd') as 填制日期,A.审核人," & _
            "       to_char(A.审核日期,'yyyy-MM-dd') as 审核日期,'' as 选择 " & _
            " From 材料采购计划 A,部门表 B,材料计划内容 c,材料特性 d " & _
            " Where A.部门id=B.id and a.id=c.计划id and c.材料id=d.材料id and (b.站点=[6] or b.站点 is null)" & _
            "       and a.单据=1 "
    End If
    strSQL = strSQL & _
    "       And (A.审核日期 between [2] and [3] )  And nvl(A.库房id,[5])=[5] " & _
            IIf(lng部门ID = 0, "", " And a.部门id=[4]") & _
    " Group by A.no,B.名称,A.编制说明,A.编制人,A.编制日期,A.审核人,A.审核日期,A.计划类型" & _
    " Order by to_char(A.审核日期,'yyyy-MM-dd') Desc,A.NO"
    
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, mstr卫材分类, dtStartDate, dtEndDate, lng部门ID, mlng库房id, gstrNodeNo)
    stbThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    
    With mshHead
        .Redraw = False
        Set mshHead.Recordset = rsList
        If .Rows = 1 Then
            .Rows = 2
            .Row = 1
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    Call SetColWidth
    Call SetDetal
    mshHead_EnterCell
    mintLastRow = 0
    mlngSelectCount = 0
    mshHead.Redraw = True
    Cmd保存.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshHead_EnterCell()
    Dim strSQL As String
    Dim rsDetail As New Recordset
    Dim lngLop As Integer
    If mintLastRow = mshHead.Row Then Exit Sub
    mintLastRow = mshHead.Row
    
    On Error GoTo ErrHandle
    If mstr卫材分类 <> "" Then
        strSQL = "" & _
                "Select /*+ Rule*/ A.序号, ('['|| q.编码 || ']' || q.名称) as 材料信息,q.规格,q.产地," & _
                "       ltrim(to_char(A.请购数量/" & IIf(mintUnit = 0, "1", "b.换算系数") & "," & mOraFMT.FM_数量 & "))  as 请购数量," & _
                "       ltrim(to_char(A.计划数量/" & IIf(mintUnit = 0, "1", "b.换算系数") & "," & mOraFMT.FM_数量 & ")) as 审批数量," & _
                        IIf(mintUnit = 0, "Q.计算单位", "B.包装单位") & " as 单位, " & _
                "       ltrim(to_char((" & IIf(mintUnit = 0, "1", "b.换算系数") & " * A.单价)," & mOraFMT.FM_成本价 & ")) as 单价," & _
                "       ltrim(to_char(A.金额," & mOraFMT.FM_金额 & ")) as 金额 " & _
                "   From 材料计划内容 A,材料特性 B,材料采购计划 c,收费项目目录 Q,诊疗项目目录 M, " & _
                "       Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) J " & _
                "   Where A.材料id=B.材料id and A.材料id=q.id And a.计划id=c.id and c.单据=1 And C.No =[2]  " & _
                "         And B.诊疗id=M.id and M.分类id=J.Column_Value" & _
                "   Order by A.序号"
    Else
        strSQL = "" & _
                "Select A.序号, ('['|| q.编码 || ']' || q.名称) as 材料信息,q.规格,q.产地," & _
                "       ltrim(to_char(A.请购数量/" & IIf(mintUnit = 0, "1", "b.换算系数") & "," & mOraFMT.FM_数量 & "))  as 请购数量," & _
                "       ltrim(to_char(A.计划数量/" & IIf(mintUnit = 0, "1", "b.换算系数") & "," & mOraFMT.FM_数量 & ")) as 审批数量," & _
                        IIf(mintUnit = 0, "Q.计算单位", "B.包装单位") & " as 单位, " & _
                "       ltrim(to_char((" & IIf(mintUnit = 0, "1", "b.换算系数") & " * A.单价)," & mOraFMT.FM_成本价 & ")) as 单价," & _
                "       ltrim(to_char(A.金额," & mOraFMT.FM_金额 & ")) as 金额 " & _
                "   From 材料计划内容 A,材料特性 B,材料采购计划 c,收费项目目录 Q " & _
                "   Where A.材料id=B.材料id and A.材料id=q.id And a.计划id=c.id and c.单据=1 And C.No =[2]  " & _
                "   Order by A.序号"
    End If
    Set rsDetail = zlDatabase.OpenSQLRecord(strSQL, mstrCaption, mstr卫材分类, mshHead.TextMatrix(mshHead.Row, 0))
    With rsDetail
        mshDetail.Rows = 2
        mshDetail.Redraw = False
        For lngLop = 0 To mshDetail.Cols - 1
            mshDetail.TextMatrix(1, lngLop) = ""
            mshDetail.ColAlignment(lngLop) = 1
        Next
        
        If Not .EOF Then
            For lngLop = 1 To .RecordCount
                mshDetail.TextMatrix(lngLop, 0) = zlStr.Nvl(!材料信息)
                mshDetail.TextMatrix(lngLop, 1) = zlStr.Nvl(!规格)
                mshDetail.TextMatrix(lngLop, 2) = zlStr.Nvl(!产地)
                mshDetail.TextMatrix(lngLop, 3) = zlStr.Nvl(!单位)
                mshDetail.TextMatrix(lngLop, 4) = zlStr.Nvl(!请购数量)
                mshDetail.TextMatrix(lngLop, 5) = zlStr.Nvl(!审批数量)
                mshDetail.TextMatrix(lngLop, 6) = zlStr.Nvl(!单价)
                mshDetail.TextMatrix(lngLop, 7) = zlStr.Nvl(!金额)
                If lngLop = mshDetail.Rows - 1 Then mshDetail.Rows = mshDetail.Rows + 1
                .MoveNext
            Next
            If .RecordCount > 0 Then
                .MoveFirst
                mshDetail.Row = 1
                mshDetail.Col = 0
                mshDetail.ColSel = mshDetail.Cols - 1
            End If
        End If
        mshDetail.Redraw = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function SetDetal()
    Dim i As Long
    With mshDetail
        .Clear
        .Rows = 2
        .Cols = 8
        .TextMatrix(0, 0) = "材料信息"
        .TextMatrix(0, 1) = "规格"
        .TextMatrix(0, 2) = "产地"
        .TextMatrix(0, 3) = "单位"
        .TextMatrix(0, 4) = "申购数量"
        .TextMatrix(0, 5) = "审批数量"
        .TextMatrix(0, 6) = "单价"
        .TextMatrix(0, 7) = "金额"
        For i = 0 To .Cols - 1
            .ColAlignment(i) = IIf(i <= 3, 1, 7)
            .ColAlignmentFixed(i) = 4
        Next
        If mblnFirst Then
            .ColWidth(0) = 2500
            .ColWidth(1) = 1000
            .ColWidth(2) = 1200
            .ColWidth(3) = 500
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 1000
        End If
    End With
End Function
Private Sub mshHead_DblClick()
    Dim lngCol As Long
    If mshHead.TextMatrix(mshHead.Row, 0) = "" Then Exit Sub
    lngCol = GetCol(mshHead, "选择")
    If mshHead.TextMatrix(mshHead.Row, lngCol) = "" Then
        mshHead.TextMatrix(mshHead.Row, lngCol) = "√"
        mlngSelectCount = mlngSelectCount + 1
    Else
        mshHead.TextMatrix(mshHead.Row, lngCol) = ""
        mlngSelectCount = mlngSelectCount - 1
    End If
   
    If mlngSelectCount = 0 Then
        Cmd保存.Enabled = False
    Else
        Cmd保存.Enabled = True
    End If
End Sub


Private Sub mshHead_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then mshHead_DblClick
End Sub

Private Sub txtDept_Change()
    txtDept.Tag = ""
End Sub

Private Sub txtDept_GotFocus()
    OS.OpenIme True
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtDept.Tag) <> "" Or Trim(txtDept.Text) = "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If Select部门(Trim(txtDept.Text)) = False Then Exit Sub
    OS.PressKey vbKeyTab
End Sub

Private Sub txtDept_LostFocus()
    OS.OpenIme False
End Sub
