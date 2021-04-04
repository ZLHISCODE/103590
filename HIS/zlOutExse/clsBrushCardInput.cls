VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBrushCardInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*************************************************************************************************
'功能:刷卡对象
'编制:刘兴洪
'日期:2016-08-12 15:49:37
'说明:
'    1. 调用顺序:先调用"InitCompents"
'    2. 接收:ReadCardNoed（读取出来的卡号)
'    3.控制性说明:
'       1).开始接收:文本框或间隔2000毫秒时，识别第一个字符时认为开始刷卡
'       2).终止接收:特定字符和间隔大于0.04毫秒时。
'       3).接受有效字符(缺省):字母，数字
'       4).屏蔽字符(缺省): ?？:：;；
'*************************************************************************************************
Private WithEvents mobjBrush As BillEdit  '刷卡文本框对象
Attribute mobjBrush.VB_VarHelpID = -1
Private mstrInputText As String   '刷卡出来的字符串(全串，包含一些特殊字符)
Private mstrCardNO As String       '刷卡出来的有效卡号
Private msngBeginTime As Single '开始刷卡时间
Private msngPreInputCharTime As Single '上次输入字符时间
Private mstrPreChar As String   '上个字符
Private mobjMain As Object
Private mobjCard As Card
Private mintInputChar As Integer '字符接收类型:0-所有字符,1-数字,2-字母;3-数字或字母;4-指定字符
Private mstrSpecialChar As String '特定字符:Ascii码1，Ascii码2....
Private mblnEndReceiveChar As Boolean '结束接收字符
Private WithEvents mobjTime As Timer
Attribute mobjTime.VB_VarHelpID = -1

'事件声明
Public Event ReadCardNoed(ByVal strCardNo As String, ByVal blnBrushCard As Boolean)
'保持属性值的局部变量
Private mvarOnlyLegalCardNo As Boolean '局部复制:仅输入合法的卡号字符(不分手工输入还是刷卡输入)

Public Property Let OnlyLegalCardNo(ByVal vData As Boolean)
Attribute OnlyLegalCardNo.VB_Description = "仅输入合法的卡号字符"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.OnlyLegalCardNo = 5
    mvarOnlyLegalCardNo = vData
End Property

Public Property Get OnlyLegalCardNo() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.OnlyLegalCardNo
    OnlyLegalCardNo = mvarOnlyLegalCardNo
End Property

Public Sub InitCompents(ByRef objMain As Object, ByRef objBrush As Object, objCard As Card)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化部件
    '入参:objMain-调用的主窗口
    '    txtBrushObj-刷卡文本框
    '    objCard-卡类别对象
    '编制:刘兴洪
    '日期:2016-08-12 16:01:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '动态加载时间控件
    Set mobjTime = objMain.Controls.Add("VB.Timer", "mtimEnd")
    With mobjTime
        .Enabled = False
        .Interval = 500
    End With
   
    Set mobjCard = objCard
    Set mobjMain = objMain
    Set mobjBrush = objBrush
    '输入类型(0-所有字符,1-数字,2-字母;3-数字或字母;4-指定字符)|Ascii码1，Ascii码2....
'    mintInputChar = Val(Split(mobjCard.卡号有效字符 & "|", "|")(0))
'    mstrSpecialChar = Trim(Split(mobjCard.卡号有效字符 & "|", "|")(1))
    mblnEndReceiveChar = False
End Sub

Private Sub mobjBrush_EditChange(strInput As String)
    If Trim(mobjBrush.Text) <> "" Then Exit Sub
    mstrInputText = "": mstrCardNO = "": msngBeginTime = 0
    mblnEndReceiveChar = False
    mobjTime.Enabled = False
End Sub

Private Sub mobjBrush_EditKeyPress(KeyAscii As Integer)
    Dim sngNow As Single, blnCard As Boolean
    Dim blnAdd As Boolean
    Dim intKeyAscii As Integer
    
    
    sngNow = Timer: blnCard = False: mobjTime.Enabled = False
    
    intKeyAscii = KeyAscii
    If Format((sngNow - msngPreInputCharTime), "0.000") > 2 Or KeyAscii = 8 Or mstrInputText = "" Or msngBeginTime = 0 Then
        mstrInputText = "": mstrCardNO = "": msngBeginTime = 0
        mblnEndReceiveChar = False
    End If
    
    sngNow = 0
    msngPreInputCharTime = Timer
    
    If mstrInputText = "" Then
        msngBeginTime = Timer
       ' Debug.Print "msngBeginTime:" & msngBeginTime & "-->" & mstrInputText
    ElseIf KeyAscii > 32 Or KeyAscii = 13 Then
        sngNow = Timer
        blnCard = Format((sngNow - msngBeginTime) / IIf(mstrInputText = "", 1, Len(mstrInputText)), "0.000") < 0.04  '一般在0.014左右
        'Debug.Print "blnCard=" & blnCard & " ---->" & KeyAscii & " -->" & "sngNow:" & sngNow & " - " & msngBeginTime & "=" & (sngNow - msngBeginTime) & "  -->" & mstrInputText
    End If
    
    If Not blnCard Then  '当前不是刷卡，即手工输入
        If KeyAscii = 13 Or KeyAscii = 8 Then
            mstrCardNO = "": mstrInputText = "": msngBeginTime = 0
        Else
'            If InStr(1, ":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0
            If KeyAscii <> 8 And KeyAscii <> 0 Then
                mstrInputText = Chr(KeyAscii): mstrPreChar = Chr(KeyAscii)
            End If
            mstrCardNO = ""
        End If
        '可以做成属性，是否达到卡号长度时，手工输入达到长度时，也自动找（主要歉容以前的，如果卡号输入达到长度，就会自动查找）
        If (Len(mobjBrush.Text) = IIf(mobjCard.卡号长度 = 0, 100, mobjCard.卡号长度) - 1 And IsNumeric(mobjBrush.Text)) Or KeyAscii = 13 Then
            '达到长度，也去找,避免操作员需要按一下回车
            mstrCardNO = mobjBrush.Text & Chr(KeyAscii)
            RaiseEvent ReadCardNoed(mstrCardNO, False)
            mstrCardNO = "": mstrInputText = ""
        End If
        
        '只接收合法的卡号信息
'        If mvarOnlyLegalCardNo Then If CheckInputCharValied(KeyAscii) = False Then KeyAscii = 0
        Exit Sub
    End If
    
    '刷卡处理
     If mstrPreChar <> "" Then  '上一个有效卡号字符
        If CheckInputCharValied(Asc(mstrPreChar)) = False Then mstrPreChar = ""
        mstrCardNO = mstrPreChar: mstrPreChar = ""
    End If
    If KeyAscii = 13 Then    '有直接结束标志的，则直接确定
'        KeyAscii = 0
        mobjTime.Enabled = False
        Call mobjTime_Timer
        Exit Sub
    End If
    If KeyAscii <> 8 And KeyAscii > 32 Then
        mstrInputText = mstrInputText & Chr(KeyAscii)
    End If
    
    mobjTime.Enabled = True
    '只接收合法的卡号信息
'    If CheckInputCharValied(KeyAscii) = False Then
'        intKeyAscii = 0
'        If mvarOnlyLegalCardNo Then KeyAscii = 0
'        Exit Sub
'    End If
    If Len(mstrCardNO) > IIf(mobjCard.卡号长度 = 0, 100, mobjCard.卡号长度) - 1 Then
        intKeyAscii = 0: Exit Sub '到达卡号长度时，不在接收字符
    End If
    If KeyAscii <> 0 Then mstrCardNO = mstrCardNO & Chr(KeyAscii) '合法卡号
End Sub
Private Function CheckInputCharValied(ByRef KeyAscii As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的字符是否合法
    '返回:合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-08-17 16:42:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    CheckInputCharValied = False
    If KeyAscii = 8 Or KeyAscii = 13 Then CheckInputCharValied = True: Exit Function
    
    '1.特殊字符判断
    If InStr(1, ":：;；?？", Chr(KeyAscii)) > 0 Then Exit Function
    
    '2.卡号的结束符，比如:213442=222 如果结束符为:=,则取213442
'    If InStr(1, "," & mobjCard.卡号结束符 & ",", "," & KeyAscii & ",") > 0 Then mblnEndReceiveChar = True
    If mblnEndReceiveChar Then Exit Function
    
    '3.无效字符判断
'    If InStr(1, "," & mobjCard.卡号无效字符 & ",", "," & KeyAscii & ",") > 0 Then Exit Function


    '4.有效字符判断: mintInputChar:0-所有字符,1-数字,2-字母;3-数字或字母;4-指定字符
    If mintInputChar = 1 And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then Exit Function   '数字
    If mintInputChar = 2 And (Not ((KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")))) Then Exit Function        '字母
   
    If mintInputChar = 3 Then '数字或字母
       If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) _
            Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z"))) Then Exit Function
    End If

    '6.只接收指定字符
    If mstrSpecialChar <> "" And (InStr(1, "," & mstrSpecialChar & ",", "," & KeyAscii & ",") = 0) Then Exit Function
    
    CheckInputCharValied = True
End Function
Private Sub mobjTime_Timer()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:通过时间来判断是否刷卡结束
    '编制:刘兴洪
    '日期:2016-08-17 16:55:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mobjTime.Enabled = False
    
'    If zlCommFun.ActualLen(mstrCardNO) < mobjCard.卡号最小长度 Then Exit Sub
    
    '有效字符判断: mintInputChar:0-所有字符,1-数字,2-字母;3-数字或字母;4-指定字符
    If mintInputChar = 1 And IsNumeric(mstrCardNO) = False Then Exit Sub      '数字
    If mintInputChar = 2 And zlCommFun.IsCharAlpha(mstrCardNO) = False Then Exit Sub       '字母
    If mintInputChar = 3 And zlCommFun.IsNumOrChar(mstrCardNO) = False Then Exit Sub         '数字和字母
    
    
    If Not mobjBrush Is Nothing Then
        mobjBrush.Text = mstrCardNO
        mobjBrush.SelStart = Len(mstrCardNO)
    End If
    mobjTime.Enabled = False:   mblnEndReceiveChar = False
    If mstrCardNO = "" Then mstrInputText = "": Exit Sub
    RaiseEvent ReadCardNoed(mstrCardNO, True)
    mstrCardNO = "": mstrInputText = ""
End Sub
 

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mobjTime Is Nothing Then Unload mobjTime
    Set mobjTime = Nothing
End Sub
