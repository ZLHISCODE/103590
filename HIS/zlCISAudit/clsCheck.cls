VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Enum NumberChkEnum
    ChkBlank = 0  '检测为空
    ChkBig = 1    '检测> 0        如果传值,则以传入的值为准
    ChkBigEq = 2  '检测>=0        如果传值,则以传入的值为准
    ChkNum = 3    '检测是否为数字
    ChkSml = 4    '检测< 0        如果传值,则以传入的值为准
    ChkSmlEq = 5  '检测<=0        如果传值,则以传入的值为准
End Enum

'ADO连接
Private mclsFormObjects()      As New clsFormObject

Dim strLanguage             As String
Dim strSystem               As String
Dim StrKey                  As String
Dim strSQL                  As String

'==================================================================================================
'=PUBLIC
'=============================================================
'=名称:锁定(But_Lock)
'=入口参数:
'=1.)将不Button名不属于butName的名锁定           类型:Form
'=出口参数:(无)
'=功能:清空窗体中文本框的数据
'=日期:2004-08-16
'=编程:谢荣
'==============================================================
Public Sub But_Lock(objFrm As Object, Optional butName As String)
    Dim mCtl As Control
On Error GoTo ErrH
    For Each mCtl In objFrm.Controls
        Debug.Print mCtl.Name
        If TypeOf mCtl Is CommandButton Then                      'VB的文本框
            If mCtl.Name <> butName Then mCtl.Enabled = False
        End If
    Next
    Exit Sub
ErrH:
    Err.Clear
End Sub

'=============================================================
'=名称:锁定(But_Open)
'=入口参数:
'=1.)将不Button名不属于butName的名锁定           类型:Form
'=出口参数:(无)
'=功能:清空窗体中文本框的数据
'=日期:2004-08-16
'=编程:谢荣
'==============================================================
Public Sub But_Open(objFrm As Object, Optional butName As String)
    Dim mCtl As Control
On Error GoTo ErrH
    For Each mCtl In objFrm.Controls
        If TypeOf mCtl Is CommandButton Then                    'VB的文本框
            If mCtl.Name <> butName Then mCtl.Enabled = True
        End If
    Next
    Exit Sub
ErrH:
    Err.Clear
End Sub

'========================================================================================
'=名稱:检测(ChkRsState)
'=入口参数:Rs               类型:ADODB.Recordset
'=出口参数:ChkRsState       类型:Boolean
'=功能:检测记录集的状态
'=日期:2004-07-08
'=編程:谢荣
'========================================================================================
Function Connection_ChkRsState(rs As ADODB.Recordset) As Boolean
On Error GoTo ErrH:
    With rs
        If rs Is Nothing Then
            Connection_ChkRsState = True
            Exit Function
        Else
            Connection_ChkRsState = False
        End If
        If rs.State = 0 Then
            Connection_ChkRsState = True
            Exit Function
        Else
            Connection_ChkRsState = False
        End If
        If .RecordCount < 1 Then
            Connection_ChkRsState = True
        Else
            Connection_ChkRsState = False
        End If
        If .EOF Or .BOF Then
            Connection_ChkRsState = True
        Else
            Connection_ChkRsState = False
        End If
    End With
    Exit Function
ErrH:
    Err.Clear
End Function

'==================================================================================================
'=记录定位
'==================================================================================================
Function Connection_GetBookMark(RsT As ADODB.Recordset, sFilter As String) As Variant
On Error GoTo ErrH
    If Connection_ChkRsState(RsT) Then
        Connection_GetBookMark = Null
    Else
        With RsT.Clone
            .Filter = sFilter
            If .EOF Or .BOF Then
                Connection_GetBookMark = Null
            Else
                Connection_GetBookMark = .Bookmark
            End If
        End With
    End If
    Exit Function
ErrH:
    Connection_GetBookMark = Null
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=删除定位
'==================================================================================================
Function Connection_DelLocate(rs As ADODB.Recordset, MPOS As Variant) As Variant
    Dim lrow As Variant
On Error GoTo ErrH
    Connection_DelLocate = Null
    With rs
        If Connection_ChkRsState(rs) Then
            Connection_DelLocate = Null
        Else
            .MoveLast
            If .AbsolutePosition <= 0 Then
                lrow = .RecordCount
            Else
                lrow = .AbsolutePosition
            End If
            If CInt(lrow) >= CInt(MPOS) Then
                Connection_DelLocate = MPOS
            Else
                If CInt(MPOS) > CInt(lrow) Then
                    Connection_DelLocate = .Bookmark
                End If
            End If
        End If
    End With
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=功    能:输入控件中的内容是为空
'=返 回 值:如果不为空,返回空字符串,如果为空返提示信息[并可以传入起先的值]
'=程    序:谢荣
'=日    期:2005-07-16
'==================================================================================================
Public Function Chk_CheckTxtNull(mDescription As String, objText As Object, Optional ctrSetF As Object, Optional strMsg As String) As String
On Error GoTo ErrH
    Chk_CheckTxtNull = strMsg
    Select Case TypeName(objText)
        Case "TextBox", "BetryNote", "BetrySelect", "BetryDate", "BetryNum", "ComboBox"
            If Trim(objText.Text) = "" Then
                If ctrSetF Is Nothing Then Set ctrSetF = objText
                Chk_CheckTxtNull = strMsg & mDescription & "不能为空!" & vbCrLf
            End If
    End Select
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'=============================================================
'=功    能:输入控件中的内容是否为数值
'=返 回 值:如果是,返回空字符串,如果不是,返提示信息
'=程    序:朱红军
'=日    期:2004-06-24
'=修改功能:传入定位控件
'=修 改 人:谢荣
'=日    期:2004-07-20
'=============================================================
Public Function Chk_CheckTxtNum(mDescription As String, objText As Object, Optional ctrSetF As Object, Optional strMsg As String, Optional mType As NumberChkEnum, Optional mChkVal As Double) As String
    Dim strTmp  As String
On Error GoTo ErrH
    If IsNull(mType) = True Then mType = ChkBlank   '检测默认为检测为空
    If IsNull(mChkVal) = True Then mChkVal = 0      '检测值默认为0
    strTmp = Trim(objText.Text)
    
    Chk_CheckTxtNum = ""
    Select Case mType
        Case ChkBlank  '检测为空
            If Trim(strTmp) = "" Then Chk_CheckTxtNum = mDescription & "不能为空!"
        Case ChkBig     '检测> 0        如果传值,则以传入的值为准
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "不能为空!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "必须为数据值!"
            Else
                If Val(strTmp) <= mChkVal Then Chk_CheckTxtNum = mDescription & "必须>" & mChkVal & "!"
            End If
        Case ChkBigEq   '检测>=0        如果传值,则以传入的值为准
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "不能为空!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "必须为数据值!"
            Else
                If Val(strTmp) < mChkVal Then Chk_CheckTxtNum = mDescription & "必须>=" & mChkVal & "!"
            End If
        Case ChkNum     '检测是否为数字
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "不能为空!"
            Else
                If IsNumeric(strTmp) = False Then Chk_CheckTxtNum = mDescription & "必须为数据值!"
            End If
        Case ChkSml     '检测< 0        如果传值,则以传入的值为准
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "不能为空!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "必须为数据值!"
            Else
                If Val(strTmp) >= mChkVal Then Chk_CheckTxtNum = mDescription & "必须<" & mChkVal & "!"
            End If
        Case ChkSmlEq   '检测<=0        如果传值,则以传入的值为准
            If strTmp = "" Then
                Chk_CheckTxtNum = mDescription & "不能为空!"
            ElseIf IsNumeric(strTmp) = False Then
                Chk_CheckTxtNum = mDescription & "不能为空!"
            Else
                If Val(strTmp) > mChkVal Then Chk_CheckTxtNum = mDescription & "必须<=" & mChkVal & "!"
            End If
    End Select
    If Chk_CheckTxtNum <> "" Then
        If ctrSetF Is Nothing Then Set ctrSetF = objText
        Chk_CheckTxtNum = Chk_CheckTxtNum & vbCrLf
        Chk_CheckTxtNum = strMsg & Chk_CheckTxtNum
    Else
        Chk_CheckTxtNum = strMsg
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=功    能:检测数据条件是否符合要求
'=返 回 值:BOOLEAN[Y.退出,N.继续]
'=入口参数:
'=       1:检测字符串
'=       2:定位的控件
'=程    序:谢荣
'=日    期:2005-07-17
'==================================================================================================
Public Function Chk_CheckMsg(strMsg As String, Optional ctrSetF As Object) As Boolean
On Error GoTo ErrH
    Chk_CheckMsg = True
    If strMsg = "" Then
        Chk_CheckMsg = False
    Else
        MsgBox strMsg, vbInformation, "中联提示"
        If (Not ctrSetF Is Nothing) Then
            If (ctrSetF.TabStop = True And ctrSetF.Enabled = True) Then ctrSetF.SetFocus
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=功    能:输入控件中的内容是为日期型
'=返 回 值:如果不为空,返回空字符串,如果为空返提示信息
'=程    序:朱红军
'=日    期:2004-06-24
'=修    改:传入字符数据,检测是否为正确的日期型数据
'=修 改 人:谢荣
'=日期:2004-01-12
'==================================================================================================
Function Chk_Date(strDate As String, Optional mTile As String) As String
    Dim strMsg    As String
On Error GoTo ErrH
    strMsg = ""
    Chk_Date = ""
    If strDate <> "" Then
        If Not IsDate(strDate) Then
            strMsg = mTile & "必须为日期型!" & vbCrLf
        ElseIf (Format(strDate, "yyyy-mm-dd")) < "1900" Then
            strMsg = mTile & "必须大于1900年!" & vbCrLf
        ElseIf (Format(strDate, "yyyy-mm-dd")) > "2099" Then
            strMsg = mTile & "必须小于2099年!" & vbCrLf
        End If
    End If
    Chk_Date = strMsg
    Exit Function
ErrH:
    Err.Clear
    Chk_Date = "必须为日期型!"
    Exit Function
End Function

'==================================================================================================
'=名  称:Chk_SrhNum
'=功  能:检测数据查询的起始条件
'=参  数:strTitle     消息名
'=       objTextQ   超始的控件
'=       objTextH   终止的控件
'=返回值:
'==================================================================================================
Public Function Chk_SrhNum(ByVal strTitle As String, objTextQ As Object, objTextH As Object, Optional ctrSetF As Object, Optional strMsg As String) As String
    Dim num1    As Double
    Dim num2    As Double
    Dim strQ    As String
    Dim strH    As String
On Error GoTo ErrH
    strQ = Trim(objTextQ.Text)
    strH = Trim(objTextH.Text)
    Chk_SrhNum = ""
    If strQ = "" Then
        num1 = 0
    Else
        If Not IsNumeric(strQ) Then
            Chk_SrhNum = Chk_SrhNum & strTitle & "的起始数量格式不正确!" & vbCrLf
            objTextQ.Text = ""
        Else
            num1 = CDbl(objTextQ.Text)
        End If
    End If
    If strH = "" Then
        num2 = 0
    Else
        If Not IsNumeric(strH) Then
            Chk_SrhNum = Chk_SrhNum & strTitle & "的终止数量格式不正确!" & vbCrLf
            objTextH.Text = ""
        Else
            num2 = CDbl(objTextH.Text)
        End If
    End If
    If num1 > num2 And Len(strH) > 0 Then
        Chk_SrhNum = Chk_SrhNum & strTitle & "范围错误。终止数不能大于起始数。" & vbCrLf
        objTextQ.Text = ""
    End If
    If Chk_SrhNum <> "" Then
        If ctrSetF Is Nothing Then Set ctrSetF = objTextQ
    End If
    Chk_SrhNum = Chk_SrhNum & strMsg
    Exit Function
ErrH:
    Chk_SrhNum = ""
    Err.Clear
    Resume Next
End Function

'==================================================================================================
'=名称:得到下拉列表框的数据(Cmb_List)
'=入口参数:
'=1).下拉列表框控件         类型:Control
'=2).原因类别ID             类型:String
'=3).类别(btyEditMode)      类型:Byte
'=      1.无参数:标准数据加载
'=      2.参数为2:标准数据查询
'=出口参数:空
'=功能:从SSQL中找到数据加载到下拉列表框中去
'=日期:2004-12-11
'=编程:谢荣
'=说明:在原因类别ID中的数据不能带"-"
'==================================================================================================
Sub Cmb_List(Combo As Object, rs As ADODB.Recordset, Optional btyEditMode As Byte, Optional btyOne As Byte)
    Dim lngStep     As Long
On Error GoTo ErrH
    Combo.Clear
 
    If Connection_ChkRsState(rs) Then Exit Sub
    lngStep = 0
    With rs
        .MoveFirst
        Do While Not (.EOF Or .BOF)
            If btyOne = 1 Then
                Combo.List(lngStep) = Trim(rs.Fields(0))
                lngStep = lngStep + 1
                rs.MoveNext
            Else
                Combo.List(lngStep) = Trim(rs.Fields(0)) + strSplitCmb + Trim(rs.Fields(1))
                lngStep = lngStep + 1
                rs.MoveNext
            End If
        Loop
        Combo.ListIndex = 0
    End With
    If btyEditMode = 2 Then
        Combo.List(lngStep) = ""
        Combo.ListIndex = -1
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'==================================================================================================
'=名称:得到下拉列表框的Text属性中取得ID(Cmb_ID)
'=入口参数:
'=1).下拉列表框控件         类型:Control
'=出口参数:空
'=功能:得到下拉列表框的Text属性中取得ID
'=日期:2004-12-11
'=编程:谢荣
'=说明:在原因类别ID中的数据不能带"-"
'==================================================================================================
Function Cmb_ID(Combo As Object, Optional Index As Byte = 1) As String
    Dim xx          As Variant
On Error GoTo ErrH
    If Combo.Text = "" Then
        Cmb_ID = ""
    Else
        xx = Split(Combo.Text, strSplitCmb)
        If Index - 1 <= UBound(xx) Then '最大下标值小于输入值[证明有截取值]
            Cmb_ID = xx(Index - 1)
        Else                        '最大下标值大于等于输入值[证明有无截取值]返回无
            Cmb_ID = "[无]"
        End If
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=名称:得到下拉列表框的Text属性中取得ID(Cmb_EditIndex)
'=入口参数:
'=1).下拉列表框控件         类型:Control
'=出口参数:空
'=功能:得到下拉列表框的Text属性中取得ID
'=日期:2004-12-11
'=编程:谢荣
'=说明:在原因类别ID中的数据不能带"-"
'==================================================================================================
Function Cmb_EditIndex(Combo As Object, sID As String) As Long
    Dim lngCount    As Long
    Dim lngStep     As Long
    Dim xx          As Variant
On Error GoTo ErrH
    lngCount = Combo.ListCount - 1
    For lngStep = 0 To lngCount
        xx = Split(Combo.List(lngStep), strSplitCmb)
        If sID = xx(0) Then
            Cmb_EditIndex = lngStep
            Exit For
        End If
    Next
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=名称:对文本赋值(Frm_GetControl)
'=入口参数:
'=1.)objFrm所需清空的窗体名             类型:Form
'=2.)Rs记录集                           ADODB.Recordset
'=出口参数:(无)
'=功能:清空窗体中文本框的数据
'=日期:2004-08-16
'=编程:谢荣
'=说明:对文本框DataField必需赋上对应的字段的值
'=修改:2004-09-13   程序简化,并添加CheckBox复选框数据
'=所用控件:1.TextBox--VB的文本框    2.BetryNote--备注输入控件   3.BetrySelect--数据选择控件
'=4.BetryDate--日期输入控件      5.BetryNum--数据输入控件    6.CheckBox--复选框数据
'==================================================================================================
Public Sub Frm_GetControl(objFrm As Object, rs As ADODB.Recordset)
    Dim mCtl As Control
On Error GoTo ErrH
    With rs.Clone
        For Each mCtl In objFrm.Controls
            Select Case TypeName(mCtl)
                Case "TextBox"
                    If Len(Trim(mCtl.DataField)) > 0 Then
                        If mCtl.Tag <> "" Then
                            If UCase(mCtl.Tag) <> "MAX" Or UCase(mCtl.Tag) <> "MIN" Then
                                mCtl.Text = "" & .Fields(mCtl.DataField)
                            Else
                                mCtl.Text = Format("" & .Fields(mCtl.DataField), mCtl.Tag)
                            End If
                        Else
                            mCtl.Text = "" & .Fields(mCtl.DataField)
                        End If
                    End If
                Case "ComboBox"
                    If Len(Trim(mCtl.DataField)) > 0 Then
                        If mCtl.STYLE <> 2 Then
                            mCtl.Text = "" & .Fields(mCtl.DataField)
                        Else
                            mCtl.ListIndex = Cmb_EditIndex(mCtl, "" & rs.Fields(mCtl.DataField))
                        End If
                    End If
                Case "CheckBox"
                    If Len(Trim(mCtl.DataField)) > 0 Then mCtl.Value = IIf("" & .Fields(mCtl.DataField) = "1", 1, 0)
            End Select
        Next
    End With
    Exit Sub
ErrH:
    Debug.Print mCtl.DataField
    Err.Clear
    Resume Next
End Sub

'========================================================================================
'=名称:清空(System)
'=入口参数:
'=1.)objFrm所需清空的窗体名           类型:Form
'=2.)SetMode                          类型:Byte
'=    (1.无参数,标准浏览窗体;2.参数为2,标准编辑,查询窗体)
'=出口参数:(无)
'=功能:清空窗体中文本框的数据(对系统数据窗体格式化)
'=日期:2004-08-16
'=编程:谢荣
'========================================================================================
Public Sub Sys_System(objFrm As Object)
    Dim mCtl                As Control
    Dim i                   As Long
    Dim intSetp             As Integer
    Dim strTabName          As String
    
On Error GoTo ErrH
    i = objFrm.hWnd
    ReDim Preserve mclsFormObjects(i + objFrm.Controls.count)
    
    For Each mCtl In objFrm.Controls
        With mCtl
            Select Case TypeName(mCtl)
                Case "SkinLabel"
                    .Caption = (.Caption)
                Case "TextBox"
                    Set mclsFormObjects(i).txt = mCtl
                    .Text = ""
                    .Appearance = 1
                    If .Locked Then .BackColor = &H80000000 Else .BackColor = vbWhite
                Case "CheckBox"
                    Set mclsFormObjects(i).chk = mCtl
                    .Value = 0
                    
                Case "ComboBox"
                    Set mclsFormObjects(i).Cmb = mCtl
                    .ListIndex = -1
                    .Appearance = 1
                    If .Locked Then .BackColor = &H80000000 Else .BackColor = vbWhite
                Case "CommandButton"
                    .Caption = (.Caption)
                    .ToolTipText = (.ToolTipText)
                Case "Label"
                    .Caption = (.Caption)
                Case "Frame"
                    .Caption = (.Caption)
                Case "VSFlexGrid"
                    RestoreFlexState mCtl, objFrm.Caption
            End Select
        End With
        i = i + 1
    Next
    Exit Sub
ErrH:
    Err.Clear
    Resume Next
End Sub

'========================================================================================
'=传入货币数字,得到货币的大写汉字
'========================================================================================
Function Sys_HBTran(money As String) As String '
    Dim X               As String
    Dim Y               As String
    Dim i               As Long
    Const zimu = ".sbqwsbqysbqwsbq" '定义位置代码
    Const letter = "0123456789sbqwy.zjf" '定义汉字缩写
    Const upcase = "零壹贰叁肆伍陆柒捌玖拾佰仟萬億圆整角分" '定义大写汉字
    Dim temp As String
On Error GoTo ErrH
    temp = money
    If InStr(temp, ".") > 0 Then temp = Left(temp, InStr(temp, ".") - 1)
    If Len(temp) > 16 Then MsgBox "数目太大，无法换算！请输入一亿亿以下的数字", 64, "错误提示": Exit Function '只能转换一亿亿元以下数目的货币！
    X = Format(money, "0.00")                                           '格式化货币
    Y = ""
    For i = 1 To Len(X) - 3
        Y = Y & Mid(X, i, 1) & Mid(zimu, Len(X) - 2 - i, 1)
    Next
    If Right(X, 3) = ".00" Then
        Y = Y & "z"                                                     '***元整
    Else
        Y = Y & Left(Right(X, 2), 1) & "j" & Right(X, 1) & "f"          '*元*角*分
    End If
    Y = Replace(Y, "0q", "0")                                           '避免零千(如：40200肆萬零千零贰佰)
    Y = Replace(Y, "0b", "0")                                           '避免零百(如：41000肆萬壹千零佰)
    Y = Replace(Y, "0s", "0")                                           '避免零十(如：204贰佰零拾零肆)
    Do While Y <> Replace(Y, "00", "0")
        Y = Replace(Y, "00", "0")                                       '避免双零(如：1004壹仟零零肆)
    Loop
    Y = Replace(Y, "0y", "y")                                           '避免零億(如：210億     贰佰壹十零億)
    Y = Replace(Y, "0w", "w")                                           '避免零萬(如：210萬     贰佰壹十零萬)
    Y = IIf(Len(X) = 5 And Left(Y, 1) = "1", Right(Y, Len(Y) - 1), Y)   '避免壹十(如：14壹拾肆；10壹拾)
    Y = IIf(Len(X) = 4, Replace(Y, "0.", ""), Replace(Y, "0.", "."))    '避免零元(如：20.00贰拾零圆；0.12零圆壹角贰分)
    For i = 1 To 19
        Y = Replace(Y, Mid(letter, i, 1), Mid(upcase, i, 1)) '大写汉字
    Next
    Sys_HBTran = Y
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'=名  称:Find_Txt
'=功  能:查询数字类型的返回SQL条件
'=参  数:mField     字段名
'=       objTextQ   超始的文本框
'=       objTextH   终止的文本框
'=返回值:
'==================================================================================================
Public Function Find_Txt(ByVal strField As String, objText As Object, sSql As String) As String
    Dim strSQL  As String
On Error GoTo ErrH
    Find_Txt = sSql
    If Trim(objText.Text) = "" Then Exit Function
    strSQL = strField & " like '" & GetPPFS & Trim(objText.Text) & "%'"
    If Find_Txt = "" Then
        Find_Txt = Trim(strSQL) & vbCrLf
    Else
        Find_Txt = Trim(Find_Txt & " AND " & strSQL) & vbCrLf
    End If
    Exit Function
ErrH:
    Err.Clear
    Resume Next
End Function

'==================================================================================================
'=名  称:Find_Cmb
'=功  能:查询下拉列表中的数据的返回SQL条件
'=参  数:mField     字段名
'=       objTextQ   超始的文本框
'=       objTextH   终止的文本框
'=返回值:
'==================================================================================================
Public Function Find_Cmb(ByVal strField As String, objCmb As Object, sSql As String, Optional BtyNoSP As Byte) As String
    Dim strSQL  As String
On Error GoTo ErrH
    Find_Cmb = sSql
    If Trim(objCmb.Text) = "" Then Exit Function
    If BtyNoSP = 0 Then
        strSQL = " nvl(" & strField & ",'0') = '" & Trim(Cmb_ID(objCmb)) & "'"
    Else
        strSQL = " nvl(" & strField & ",'') = '" & Trim(Cmb_ID(objCmb)) & "'"
    End If
    If Find_Cmb = "" Then
        Find_Cmb = Trim(strSQL) & vbCrLf
    Else
        Find_Cmb = Trim(Find_Cmb & " AND " & strSQL) & vbCrLf
    End If
    Exit Function
ErrH:
    Err.Clear
    Resume Next
End Function

'==================================================================================================
'=得到查询窗口的条件[1.查询主表条件,2.查询明细条件]
'==================================================================================================
Function Frm_GetFilter(ByRef objFrm As Object, Optional FindMode As Byte = 0) As String
    Dim mCtl        As Control
    Dim QCnt        As Control '前值
    Dim HCnt        As Control '后值
    Dim sTmp        As String
On Error GoTo ErrH
    sTmp = "1=1 "
    Select Case FindMode
        Case 0
            For Each mCtl In objFrm.Controls
                '名称2为后,则下次循环
                If Right(mCtl.Name, 1) <> "2" Then
                    If mCtl.Tag = "" Or mCtl.Tag = "M" Then
                        Select Case TypeName(mCtl)
                            Case "TextBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Txt(mCtl.DataField, mCtl, sTmp)
                             
                            Case "ComboBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Cmb(mCtl.DataField, mCtl, sTmp)
                        End Select
                    End If
                End If
            Next
        Case 1
            For Each mCtl In objFrm.Controls
                '名称2为后,则下次循环
                If Right(mCtl.Name, 1) <> "2" Then
                    If mCtl.Tag = "D" Then
                        Select Case TypeName(mCtl)
                            Case "TextBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Txt(mCtl.DataField, mCtl, sTmp)
                            Case "ComboBox"
                                If mCtl.DataField <> "" Then sTmp = Find_Cmb(mCtl.DataField, mCtl, sTmp)
                        End Select
                    End If
                End If
            Next
    End Select
    If Trim(sTmp) = "1=1" Then Frm_GetFilter = "" Else Frm_GetFilter = UCase(sTmp)
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function
'==================================================================================================
'确定/取消消息
'==================================================================================================
Public Function Msg_OKC(StrMessage As String, Optional MsgType As VbMsgBoxStyle = vbOKCancel + vbDefaultButton2 + vbQuestion, Optional strTitle As String = "中联提示") As Boolean
On Error GoTo ErrH
    Msg_OKC = IIf(MsgBox(StrMessage, MsgType, strTitle) <> 1, True, False)
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

'==================================================================================================
'确定消息
'==================================================================================================
Public Sub Msg_OK(StrOKMessage As String, Optional MsgType As VbMsgBoxStyle = vbOKOnly + vbInformation, Optional strTitle As String = "中联提示")
On Error GoTo ErrH
    MsgBox StrOKMessage, MsgType, strTitle
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

'==================================================================================================
'ORACLE 表结构
'==================================================================================================
Function GetRsFieldWidth(strTableName) As ADODB.Recordset
    On Error GoTo ErrH
    
    gstrSQL = "" & vbCrLf & _
        "SELECT USER_TAB_COLS.TABLE_NAME   as 表名," & vbCrLf & _
        "       USER_TAB_COLS.COLUMN_NAME  as 列名," & vbCrLf & _
        "       USER_TAB_COLS.DATA_TYPE    as 数据类型," & vbCrLf & _
        "       USER_TAB_COLS.DATA_LENGTH  as 长度," & vbCrLf & _
        "       USER_TAB_COLS.NULLABLE     as 是否为空," & vbCrLf & _
        "       USER_TAB_COLS.COLUMN_ID    as 列序号," & vbCrLf & _
        "       user_col_comments.comments as 备注" & vbCrLf & _
        "  From USER_TAB_COLS" & vbCrLf & _
        " INNER join user_col_comments" & vbCrLf & _
        "    ON user_col_comments.TABLE_NAME = USER_TAB_COLS.TABLE_NAME" & vbCrLf & _
        "   AND user_col_comments.COLUMN_NAME = USER_TAB_COLS.COLUMN_NAME" & vbCrLf & _
        " WHERE USER_TAB_COLS.Table_Name = [1]"
    
    Set GetRsFieldWidth = zlDatabase.OpenSQLRecord(gstrSQL, "clsCheck", strTableName)
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Function
End Function
