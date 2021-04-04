Attribute VB_Name = "Mdl通用"
Option Explicit
Public Enum gEditType
     g新增 = 0
     g修改 = 1
     g审核 = 2
     g取消 = 3
     g查看 = 4
     g预审 = 7
End Enum
Public Enum RecBillStatus  '记录状态信息
    正常记录 = 1
    冲销记录 = 2
    被冲销记录 = 3
End Enum
Public Enum ErrBillStatusInfor  '单据状态信息
    正常情况 = 1
    已经删除
    已经审核
    已经冲销
End Enum
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Public Const glngGetFocus As Long = &HA87B82                    '进入时的选择颜色
Public Const glngGetFocus_Font As Long = &H80000005             '进入时的字体颜色
Public Const glngLostFocus As Long = &HC0C0C0                   '离开时的选择色
Public Const glngLostFocus_Font As Long = &H80000008            '离开时的字体色

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '返回:返回加匹配串%dd%,并且是大写
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper = False Then
        GetMatchingSting = strLeft & strString & strRight
    Else
        GetMatchingSting = strLeft & UCase(strString) & strRight
    End If
End Function

Public Function MulitSelectPersion(ByVal frmParent As Form, ByVal objCtl As Object, _
    ByVal strKey As String, Optional lng部门ID As Long = 0, _
    Optional ByRef lng人员ID As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:选择指定的人员
    '入参:frmParent-调用的父窗口
    '     objCtl-控件(目前只支持文本框)
    '     strKey-输入的建值
    '     lng部门ID-如果不为零,找所有人员,否则, 找指定部门下的人员
    '出参:lng人员id-返回人员ID
    '返回:查找成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/23
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    'zlDatabase.ShowSQLSelect
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
    
    Err = 0: On Error GoTo ErrHand:
    
     
     If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        If lng部门ID = 0 Then
            gstrSQL = "" & _
                "   Select ID,编号,姓名,别名,简码,性别,民族,出生日期,办公室电话" & _
                "   From 人员表 " & _
                "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] or 别名 like [1]) " & zl_获取站点限制 & "" & _
                "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
                "   order by 编号"
        Else
            gstrSQL = "" & _
                "   Select distinct a.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话" & _
                "   From 人员表 a,部门人员 C " & _
                "   Where a.id=c.人员id and c.部门Id=[2]   " & zl_获取站点限制(True, "a") & " and (a.姓名 like [1] or a.编号 like [1] or a.简码 like [1] or a.别名 like [1]) " & _
                "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)" & _
                "   order by 编号"
        End If
     Else
        If lng部门ID = 0 Then
            gstrSQL = "" & _
                "   Select ID,编号,姓名,别名,简码,性别,民族,出生日期,办公室电话" & _
                "   From 人员表 " & _
                "   Where (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) " & zl_获取站点限制 & "" & _
                "   order by 编号"
        Else
            gstrSQL = "" & _
                "   Select distinct a.ID,a.编号,a.姓名,a.别名,a.简码,a.性别,a.民族,a.出生日期,a.办公室电话" & _
                "   From 人员表 a,部门人员 C " & _
                "   Where a.id=c.人员id and c.部门Id=[2] " & _
                "       and (a.撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)  " & zl_获取站点限制(True, "a") & "" & _
                "   order by 编号"
        End If
    End If
    
    If UCase(TypeName(objCtl)) = "TEXTBOX" Then
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(frmParent, gstrSQL, 0, "人员选择器", False, "", "人员选择", False, False, True, vRect.Left - 15, vRect.Top, objCtl.Height, blnCancel, False, False, strKey, lng部门ID)
    Else
        Dim sngX As Single, sngY As Single
        Call CalcPosition(sngX, sngY, objCtl.MsfObj)
        Set rsTemp = zlDatabase.ShowSQLSelect(frmParent, gstrSQL, 0, "人员选择器", False, "", "人员选择", False, False, True, sngX, sngY - objCtl.MsfObj.CellHeight, objCtl.MsfObj.CellHeight, blnCancel, False, False, strKey, lng部门ID)
    End If
    lng人员ID = 0
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        ShowMsgbox "未找到指定的人员,请检查!"
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = "TEXTBOX" Then
        objCtl.Text = Nvl(rsTemp!姓名)
    Else
        objCtl.TextMatrix(objCtl.Row, objCtl.Col) = Nvl(rsTemp!姓名)
        objCtl.Text = Nvl(rsTemp!姓名)
    End If
    lng人员ID = Val(Nvl(rsTemp!ID))
    rsTemp.Close
    MulitSelectPersion = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


