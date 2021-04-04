Attribute VB_Name = "mdlCard"
Option Explicit

'消费卡种类和属性
Public gstrCardsAndProperty As String

'消费卡格式
Public Enum gCardFormat
    短名 = 0
    全名 = 1
    刷卡标志 = 2
    卡类别ID = 3
    卡号长度 = 4
    缺省标志 = 5
    是否存在帐户 = 6
    卡号密文 = 7
End Enum

Public Function zlfuncCard_Confirm(ByRef objSquareCard As Object, ByVal frmMain As Form, ByVal lngModule As Long, _
    ByVal strPrivs As String, ByVal lng病人ID As Long, _
    ByVal lngCardID As Long, ByVal intType As Integer, _
    ByVal strNos As String) As Boolean
    
    If objSquareCard.zlSquareAffirm(frmMain, lngModule, strPrivs, lng病人ID, lngCardID, False, intType, strNos) = False Then
        Exit Function
    End If
    zlfuncCard_Confirm = True
End Function

Public Function zlfuncCard_GetPatiName(ByRef objSquareCard As Object, ByVal lngCardID As Long, ByVal strCardNo As String) As String
    '一卡通功能：通过卡号取病人姓名
    Dim lng病人ID As Long
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not objSquareCard Is Nothing Then
        '通过卡ID和卡号查找病人ID
        objSquareCard.zlGetPatiID CStr(lngCardID), strCardNo, False, lng病人ID
        If lng病人ID > 0 Then
            gstrSQL = "Select 姓名 From 病人信息 Where 病人id = [1] "
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "FindSpecialRow", lng病人ID)
            If Not rsData.EOF Then
                zlfuncCard_GetPatiName = UCase(rsData!姓名)
            End If
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function zlfuncCard_GetPatiID(ByRef objSquareCard As Object, ByVal lngCardID As Long, ByVal strCardNo As String) As Long
    '一卡通功能：通过卡号取病人ID
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    If Not objSquareCard Is Nothing Then
        '通过卡ID和卡号查找病人ID
        objSquareCard.zlGetPatiID CStr(lngCardID), strCardNo, False, lng病人ID
        
        If lng病人ID > 0 Then
            zlfuncCard_GetPatiID = lng病人ID
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlfuncCard_Ini(ByRef objSquareCard As Object, ByVal frmMain As Form, ByVal lngModule As Long) As String
    '一卡通接口初始化，返回消费卡种类和属性
    Dim strCards As String
    
    On Error Resume Next
    
    Set objSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not objSquareCard Is Nothing Then
        If objSquareCard.zlInitComponents(frmMain, lngModule, glngSys, gstrDbUser, gcnOracle) = False Then
            Set objSquareCard = Nothing
        Else
            strCards = objSquareCard.zlGetIDKindStr
            
            '仅“就诊卡”类别及以后的为消费卡
            zlfuncCard_Ini = Mid(strCards, InStr(1, strCards, "就|就诊卡"))
        End If
    End If
End Function

Public Sub zlfuncCard_SetCardMenu(ByVal lngModule As Long, ByVal objMenu As Object, ByVal strCards As String)
    '设置消费卡菜单，
    
End Sub

Public Sub zlfuncCard_SetText(ByVal objTxt As TextBox, ByVal strCardProperty As String)
    '设置输入框属性
    '银行卡类别，格式：短名|全名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
    objTxt.Text = ""
    objTxt.Tag = ""
    objTxt.MaxLength = 0
    
    objTxt.Tag = strCardProperty
    objTxt.MaxLength = Val(Split(strCardProperty, "|")(gCardFormat.卡号长度))
    objTxt.PasswordChar = IIf(Trim(Split(strCardProperty, "|")(gCardFormat.卡号密文)) <> "", "*", "")
End Sub

Public Sub zlfuncCard_Unload(ByRef objSquareCard As Object)
    '卸载一卡通接口
    Set objSquareCard = Nothing
End Sub
