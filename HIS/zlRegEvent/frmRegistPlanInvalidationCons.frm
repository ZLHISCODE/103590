VERSION 5.00
Begin VB.Form frmRegistPlanInvalidationCons 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "安排停用条件设置"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   -75
      TabIndex        =   15
      Top             =   3540
      Width           =   9345
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   0
      Left            =   15
      TabIndex        =   14
      Top             =   840
      Width           =   9345
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   315
      Index           =   3
      Left            =   5460
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2925
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   3
      Left            =   1095
      TabIndex        =   7
      Tag             =   "号类"
      Top             =   2880
      Width           =   4365
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   315
      Index           =   2
      Left            =   5460
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2445
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   2
      Left            =   1095
      TabIndex        =   5
      Tag             =   "号类"
      Top             =   2430
      Width           =   4365
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   315
      Index           =   1
      Left            =   5460
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1980
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   1
      Left            =   1095
      TabIndex        =   3
      Tag             =   "号类"
      Top             =   1965
      Width           =   4365
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   315
      Index           =   0
      Left            =   5460
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1485
      Width           =   345
   End
   Begin VB.TextBox txtEdit 
      Height          =   330
      Index           =   0
      Left            =   1095
      TabIndex        =   1
      Tag             =   "号类"
      Top             =   1530
      Width           =   4365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3990
      TabIndex        =   8
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5130
      TabIndex        =   9
      Top             =   3870
      Width           =   1100
   End
   Begin VB.Label Label9 
      Caption         =   "    详细设置指定停用日期的各挂号安排;以下为停用的指定各挂号安排的相关条件,它们之间的关系为且关系."
      Height          =   540
      Left            =   945
      TabIndex        =   16
      Top             =   420
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   300
      Picture         =   "frmRegistPlanInvalidationCons.frx":0000
      Top             =   330
      Width           =   480
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "医生"
      Height          =   180
      Index           =   3
      Left            =   705
      TabIndex        =   6
      Top             =   2985
      Width           =   360
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "项目"
      Height          =   180
      Index           =   2
      Left            =   705
      TabIndex        =   4
      Top             =   2505
      Width           =   360
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "科室"
      Height          =   180
      Index           =   1
      Left            =   705
      TabIndex        =   2
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      Caption         =   "号类"
      Height          =   180
      Index           =   0
      Left            =   705
      TabIndex        =   0
      Top             =   1605
      Width           =   360
   End
End
Attribute VB_Name = "frmRegistPlanInvalidationCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mblnOk As Boolean
Private mstrType As String, mstrDept As String, mstr项目 As String, mstr医生 As String
Public Function ShowCons(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    strType As String, strDept As String, str项目 As String, str医生 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示条件设置窗口(入口)
    '入参:lngModule -模块号
    '       strPrivs-权限串
    '出参:strType -号类(多个用逗号分隔)
    '       strDept-部门信息(多个用逗号分隔)
    '       str项目 -挂号项目(多个用逗号分隔)
    '       str医生-医生(格式:院内医生(ID:用逗号分隔)||院外医生(姓名:用逗号分隔)
    '返回:点确定,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-07 11:52:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrType = "": mstrDept = "": mstr项目 = "": mstr医生 = ""
    mlngModule = lngModule: mstrPrivs = strPrivs: mblnOk = False
    txtEdit(0).Tag = "": txtEdit(1).Tag = "": txtEdit(2).Tag = "": txtEdit(3).Tag = ""
    Me.Show 1, frmMain
    strType = mstrType: strDept = mstrDept: str项目 = mstr项目: str医生 = mstr医生
    ShowCons = mblnOk
End Function
Public Function SelectItem(ByVal intIndex As Integer, ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据输入的值，选择相关的数据(存在多选)
    '入参:intIndex-索引
    '       strInput-输入的值
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-07 10:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCode As String, blnCancel As Boolean, rsTemp As ADODB.Recordset
    Dim strDept As String, strDeptWhere As String, strTable As String
    Dim strLike As String, strWhere As String, bytCode As Byte
    Dim strTittle As String
    Dim vRect  As RECT
    On Error GoTo Hd
    bytCode = Val(zlDatabase.GetPara("简码方式", , , 0)) + 1
    strLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    '功能：多功能选择器,使用ADO.Command打开,允许使用[x]参数
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
    '     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等
    If strInput <> "" Then
        strCode = strLike & strInput & "%"
        If zlCommFun.IsCharAlpha(strInput) Then
                strWhere = "(A.编码 Like upper([1]) Or A.简码 Like upper([1]))"
        ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
            strWhere = "A.编码 Like upper([1])"
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strWhere = "A.名称 Like [1]"
        Else
            strWhere = "(A.名称 Like [1] Or A.编码 Like upper([1]) Or A.简码 Like upper([1]))"
        End If
    Else
        strWhere = ""
    End If
    
    Select Case intIndex
    Case 0   '号类
        If strWhere <> "" Then strWhere = " WHERE " & strWhere
        strSQL = "" & _
        "   Select rownum as ID,编码,名称,简码,缺省标志,说明 " & _
        "   From 号类 A" & _
            strWhere & _
        "   Order by 编码"
        strTittle = "号类"
    Case 1   ' 科室
        strTittle = "科室"
        '取出门诊临床科室
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质,B.服务对象 " & _
            " From 部门表 A,部门性质说明 B " & IIf(Not zlStr.IsHavePrivs(mstrPrivs, "所有科室"), ",部门人员 C", "") & _
            " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
            "           And B.部门ID=A.ID And Instr(',1,3,',',' || B.服务对象 || ',')>0 And B.工作性质 = '临床'" & _
                        IIf(Not zlStr.IsHavePrivs(mstrPrivs, "所有科室"), "  And A.id=C.部门ID and C.人员id =[2]", "") & _
            "           And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Case 2   '挂号项目
         strCode = strLike & strInput & "%"
         If strInput <> "" Then
                If zlCommFun.IsCharAlpha(strInput) Then
                        strWhere = "(A.编码 Like upper([1]) Or B.简码 Like upper([1]) and B.码类 in (3," & bytCode & "))"
                ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
                    strWhere = "A.编码 Like upper([1])"
                ElseIf zlCommFun.IsCharChinese(strInput) Then
                    strWhere = "A.名称 Like [1]"
                Else
                    strWhere = "(A.名称 Like [1] Or A.编码 Like upper([1]) Or B.简码 Like upper([1]) and B.码类 in (3," & bytCode & ") )"
                End If
                strWhere = " And " & strWhere
          Else
            strWhere = ""
          End If
            strSQL = "" & _
            "   Select Distinct A.ID, A.编码, B.名称 ,A.规格, A.产地, A.计算单位 " & _
            "   From 收费项目目录 A,收费项目别名 B " & _
            "   Where 类别='1' and A.id=B.收费细目ID  " & _
            "           And  (A.撤档时间>=to_date('3000-01-01','yyyy-mm-dd') Or A.撤档时间 Is Null)  " & strWhere & _
            "           And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            "           And rownum<101 " & _
            "   Order by 名称"
        strTittle = "挂号项目"
    Case 3  '医生
        strTittle = "医生"
        strDept = "": strTable = ""
        If txtEdit(1).Tag <> "" Then
            strDept = Trim(txtEdit(1).Tag)
            If InStr(1, strDept, ",") > 0 Then
                If zlCommFun.ActualLen(strDept) > 1990 Then
                    strTable = "Select Column_Value as ID from Table(Cast(f_Num2list([4]) As zlTools.t_Numlist))  "
                Else
                    strTable = " Select ID From 部门表 where id in (" & strDept & ") "
                End If
                strTable = ",(" & strTable & ") E"
                strDeptWhere = " C.部门ID=E.ID"
            Else
                strDeptWhere = " And C.部门id  =[3]"
            End If
        End If
        strWhere = Replace(strWhere, "名称", "姓名")
        If strWhere <> "" Then strWhere = " And " & strWhere
        strSQL = _
        "   Select /*+ rule */ distinct A.ID,A.编号 as 编码,A.姓名 as 名称,A.简码 " & _
        "   From 人员表 A ,人员性质说明 B, 部门人员 C" & strTable & vbCrLf & _
        "   Where A.ID=B.人员id And A.id=C.人员id  " & _
        "           And  B.人员性质='医生' And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & vbCrLf & _
        "           And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                    strDeptWhere & Replace(strWhere, "编码", "编号") & _
        "   Union ALL  " & _
        "   Select ID,编码,姓名 as 名称,简码  " & _
        "   From ( " & _
        "               Select Distinct -1*rownum  as ID,'' as 编码,A.医生姓名 as 姓名, zlspellcode(医生姓名) as 简码" & _
        "               From 挂号安排 A" & strTable & _
        "               where A.医生ID is null " & Replace(UCase(strDeptWhere), "C.部门ID", "A.科室ID") & _
        "           ) A " & _
        "  Where 1=1 " & strWhere
    Case Else
        Exit Function
    End Select
    
    vRect = zlcontrol.GetControlRect(txtEdit(intIndex).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, strTittle & "选择", False, "", "请选择", False, False, True, vRect.Left, vRect.Top, txtEdit(intIndex).Height, blnCancel, True, True, strCode, UserInfo.ID, Val(strDept), strDept)
    If blnCancel = True Then
        If txtEdit(intIndex).Enabled And txtEdit(intIndex).Visible Then txtEdit(intIndex).SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "没有找到满足条件的" & strTittle & "，请检查!", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(intIndex).Enabled And txtEdit(intIndex).Visible Then txtEdit(intIndex).SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "没有找到满足条件的" & strTittle & "，请检查!", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(intIndex).Enabled And txtEdit(intIndex).Visible Then txtEdit(intIndex).SetFocus
        Exit Function
    End If
    Dim strText As String, strValues As String, strValues1 As String
    With rsTemp
        Do While Not .EOF
            strText = strText & ";" & Nvl(rsTemp!名称)
            If intIndex <> 0 Then
                If intIndex = 3 And Val(Nvl(rsTemp!ID)) < 0 Then
                    strValues1 = strValues1 & "," & Nvl(rsTemp!名称)
                Else
                    strValues = strValues & "," & Nvl(rsTemp!ID)
                End If
            Else
                strValues = strValues & "," & Nvl(rsTemp!名称)
            End If
            .MoveNext
        Loop
        If strText <> "" Then strText = Mid(strText, 2)
        If strValues <> "" Then strValues = Mid(strValues, 2)
        If strValues1 <> "" Then strValues1 = "||" & Mid(strValues1, 2)
        txtEdit(intIndex).Text = strText: txtEdit(intIndex).Tag = strValues & strValues1
    End With
    zlCommFun.PressKey vbKeyTab
    SelectItem = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrType = txtEdit(0).Tag: mstrDept = txtEdit(1).Tag: mstr项目 = txtEdit(2).Tag: mstr医生 = txtEdit(3).Tag
    If mstrType = "" And mstrDept = "" And mstr项目 = "" And mstr医生 = "" Then
        MsgBox "未选择一个条件，不能继续!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(0).Text <> "" And mstrType = "" Then
        MsgBox "注意:" & vbCrLf & "    号类选择有误(可能你未按回车进行选择)，请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(1).Text <> "" And mstrDept = "" Then
        MsgBox "注意:" & vbCrLf & "    科室选择有误(可能你未按回车进行选择)，请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(2).Text <> "" And mstr项目 = "" Then
        MsgBox "注意:" & vbCrLf & "    挂号项目选择有误(可能你未按回车进行选择)，请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    If txtEdit(3).Text <> "" And mstr医生 = "" Then
        MsgBox "注意:" & vbCrLf & "    医生选择有误(可能你未按回车进行选择)，请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Sub
    End If
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdSel_Click(Index As Integer)
    If SelectItem(Index, "") = False Then
        Exit Sub
    End If
End Sub
Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlcontrol.TxtSelAll txtEdit(Index)
End Sub
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txtEdit(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txtEdit(Index).Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelectItem(Index, Trim(txtEdit(Index).Text)) = False Then
        Exit Sub
    End If
End Sub
