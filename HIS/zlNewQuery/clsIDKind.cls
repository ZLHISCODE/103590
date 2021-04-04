VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIDKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mintIDKind As Integer
Private mrsIDKind As New ADODB.Recordset
Private mblnCard As Boolean
Private mblnICCard As Boolean
Private mobjICCard As Object
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mcnOracle As ADODB.Connection
Private mbytApplyMode As Byte                                   '1-弹出菜单;2-下拉菜单
Private mobjFindKey As Object
Private mobjSquareCard As Object
Public Event AfterSelectSearchKey(ByVal Caption As String)
Public Event AfterInputComplete(ByVal CardNo As String)
Public Event AfterReadICCard(ByVal CardNo As String)
Public Event Location()

Public Event ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)

'接口
'######################################################################################################################
Public Function InitIDKind(ByVal bytApplyMode As Byte, ByVal cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal strUser As String, ByVal frmMain As Object, ByVal lngMoudal As Long, ByVal strIDKind As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    mbytApplyMode = bytApplyMode
    Set mcnOracle = cnOracle
    
    On Error GoTo errHand
    
    If mobjSquareCard Is Nothing Then
        Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    End If
        
    '取缺省的刷卡方式
    '短名|完成名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    '第7位后,就只能用索引,不然取不到数

    If mobjSquareCard.zlInitComponents(frmMain, lngMoudal, lngSysNo, strUser, mcnOracle) Then
        strIDKind = mobjSquareCard.zlGetIDKindStr(strIDKind)
        Set mrsIDKind = AnsyIDKind(strIDKind)
    End If
    
    InitIDKind = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsCard() As Boolean
    
    IsCard = (Val(mrsIDKind("卡类别ID").Value) > 0)
    
End Function

Public Function IsReadCard() As Boolean
    
    IsReadCard = (Val(mrsIDKind("卡类别ID").Value) > 0 And Val(mrsIDKind("刷卡标志").Value) = 1)
    
End Function


Public Function zlGetPatiIDByCardNo(ByVal strInputCardNo As String, ByRef lngPatientKey As Long, Optional ByRef strPassword As String, Optional ByRef strErrMsg As String) As Boolean
    
'    strCardPasword = zlCommFun.zlStringEncode(strCardPassword)
    If IsCard Then
        zlGetPatiIDByCardNo = mobjSquareCard.zlGetPatiID(Val(mrsIDKind("卡类别ID").Value), strInputCardNo, False, lngPatientKey, strPassword, strErrMsg)
    Else
        zlGetPatiIDByCardNo = mobjSquareCard.zlGetPatiID(GetValue("全名"), strInputCardNo, False, lngPatientKey, strPassword, strErrMsg)
    End If
    
End Function

Public Function GetPatientKey() As Long
    GetPatientKey = mobjSquareCard.zlGetPatiID
End Function

Public Property Get ShortName() As String
    '******************************************************************************************************************
    '功能:返回短名称
    '参数:
    '返回:
    '******************************************************************************************************************
    
    If mrsIDKind.EOF = False Then ShortName = mrsIDKind("短名").Value
    
End Property

Public Property Let ShortName(vData As String)
    '******************************************************************************************************************
    '功能:返回短名称
    '参数:
    '返回:
    '******************************************************************************************************************
    
    mrsIDKind.Filter = ""
    mrsIDKind.Filter = "短名='" & vData & "'"
    If mrsIDKind.RecordCount > 0 Then mintIDKind = mrsIDKind("序号").Value
    
End Property

Public Property Get IsPasswordCheck() As Boolean
    '******************************************************************************************************************
    '功能:返回短名称
    '参数:
    '返回:
    '******************************************************************************************************************
    
    If mrsIDKind.EOF = False Then LongName = mrsIDKind("全名").Value
    
End Property

Public Property Get LongName() As String
    '******************************************************************************************************************
    '功能:返回短名称
    '参数:
    '返回:
    '******************************************************************************************************************
    
    If mrsIDKind.EOF = False Then LongName = mrsIDKind("全名").Value
    
End Property

Public Property Let LongName(vData As String)
    '******************************************************************************************************************
    '功能:返回短名称
    '参数:
    '返回:
    '******************************************************************************************************************
    If IsOpen Then
        mrsIDKind.Filter = ""
        mrsIDKind.Filter = "全名='" & vData & "'"
        If mrsIDKind.RecordCount > 0 Then mintIDKind = mrsIDKind("序号").Value
    End If
    
End Property


Public Property Get CardLength() As Long
    If mrsIDKind.EOF = False Then CardLength = Val(mrsIDKind("卡号长度").Value)
End Property

Private Function IsOpen() As Boolean
    
    If Not (mrsIDKind Is Nothing) Then
        If mrsIDKind.State = adStateOpen Then
            IsOpen = True
        End If
    End If
    
End Function

Public Function GetCard(ByRef rs As ADODB.Recordset) As Boolean

    If IsOpen Then
        mrsIDKind.Filter = ""
        mrsIDKind.Filter = "卡类别id>0"
        If mrsIDKind.RecordCount > 0 Then
            Set rs = mrsIDKind
            GetCard = True
        End If
    Else
        GetCard = False
    End If
    
End Function

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)

    RaiseEvent ShowIDCardInfo(strID, strName, strSex, strNation, datBirthDay, strAddress)
    
End Sub


Public Sub EnterInputBox(Optional ByVal strIDKind As String)
    '******************************************************************************************************************
    '功能:进入输入框
    '参数:
    '返回:
    '******************************************************************************************************************
    If strIDKind = "" Then strIDKind = GetValue("全名")

    Select Case strIDKind
    Case "姓名"

        Call zlCommFun.OpenIme(True)
                    
    Case "门诊号"
        
        Call zlCommFun.OpenIme(False)
                
    Case "住院号"
        
        Call zlCommFun.OpenIme(False)
        
    Case "健康号"
        
        Call zlCommFun.OpenIme(False)
        
    Case "身份证号", "身份证"
        
        If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard
        If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(True)
        
    End Select
            
End Sub

Public Sub LeaveInputBox()
    If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(False)
End Sub


Public Sub InputKeyPress(KeyAscii As Integer, objInputTxt As Object)
    '******************************************************************************************************************
    '功能:正在输入键值
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim blnCard As Boolean
    Dim strKind As String
    
    strKind = GetValue("全名")
    
    objInputTxt.PasswordChar = IIf(GetValue("卡号密文") <> "", "*", "")
    Select Case strKind
    Case "姓名"
        
        blnCard = zlCommFun.InputIsCard(objInputTxt, KeyAscii, (GetValue("卡号密文") <> ""))
        
    Case "门诊号"
        
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    Case "住院号"
        
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    Case "健康号"
        
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    Case "身份证号"
        
        If InStr("0123456789X" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        
    End Select
    
     '刷卡完毕或输入号码后回车
    If blnCard And Len(objInputTxt.Text) = Val(GetValue("卡号长度")) - 1 And KeyAscii <> vbKeyBack Or KeyAscii = vbKeyReturn And Trim(objInputTxt.Text) <> "" Then
        
        If KeyAscii <> vbKeyReturn Then
            objInputTxt.Text = objInputTxt.Text & Chr(KeyAscii)
            objInputTxt.SelStart = Len(objInputTxt.Text)
            KeyAscii = 0
        End If

        RaiseEvent AfterInputComplete(Trim(objInputTxt.Text))
        
        zlControl.TxtSelAll objInputTxt
    End If
End Sub


Public Function ReadICCard() As String
    Dim strICCard As String
    
    If mobjICCard Is Nothing Then
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Set mobjICCard.gcnOracle = mcnOracle
    End If
    If Not mobjICCard Is Nothing Then
        strICCard = mobjICCard.Read_Card()
        If strICCard <> "" Then mblnICCard = True
        
        ReadICCard = strICCard
    End If
End Function

Public Sub ReadIDCard(strID As String, strName As String, strSex As String, strNation As String, datBirthDay As Date, strAddress As String, picPhoto As StdPicture)
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard
    Call mobjIDCard.ReadIDCard(strID, strName, strSex, strNation, datBirthDay, strAddress, picPhoto)
End Sub

Public Function GetPhotoAsStdPicture(picPhoto As StdPicture)
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard
    Call mobjIDCard.GetPhotoAsStdPicture(picPhoto)
End Function

'私有过程
'######################################################################################################################
Private Function GetValue(ByVal strName As String) As String
    GetValue = mrsIDKind(strName).Value
End Function

Private Function AnsyIDKind(ByVal strIDKind As String) As ADODB.Recordset
    '******************************************************************************************************************
    '功能:分解IDKind串
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rsIDKind As ADODB.Recordset
    Dim aryRow As Variant
    Dim aryCol As Variant
    Dim intRow As Integer
    Dim IntCount As Integer
    
    Set rsIDKind = New ADODB.Recordset
    With rsIDKind
        .Fields.Append "序号", adInteger
        .Fields.Append "短名", adVarChar, 10
        .Fields.Append "全名", adVarChar, 50
        .Fields.Append "刷卡标志", adInteger                '1-读卡;0-刷卡
        .Fields.Append "卡类别ID", adInteger
        .Fields.Append "卡号长度", adInteger
        .Fields.Append "卡号密文", adVarChar, 50
        .Fields.Append "参数", adVarChar, 50
        .Open
        .ActiveConnection = Nothing
    End With
    
    If strIDKind <> "" Then
        aryRow = Split(strIDKind, ";")
        IntCount = 0
        For intRow = 0 To UBound(aryRow)
            
            If Trim(aryRow(intRow)) <> "" Then
                aryCol = Split(aryRow(intRow), "|")
                
                IntCount = IntCount + 1
                rsIDKind.AddNew
                
                rsIDKind("序号").Value = IntCount
                rsIDKind("短名").Value = Trim(aryCol(0))
                
                rsIDKind("全名").Value = Trim(aryCol(1))
                If UBound(aryCol) >= 2 Then rsIDKind("刷卡标志").Value = Val(aryCol(2))
                If UBound(aryCol) >= 3 Then rsIDKind("卡类别ID").Value = Val(aryCol(3))
                If UBound(aryCol) >= 4 Then rsIDKind("卡号长度").Value = Val(aryCol(4))
                If UBound(aryCol) >= 7 Then rsIDKind("卡号密文").Value = Trim(aryCol(7))
                
                rsIDKind("参数").Value = Trim(aryRow(intRow))
                
            End If
            
        Next
    End If
    
    If rsIDKind.RecordCount > 0 Then rsIDKind.MoveFirst
    
    Set AnsyIDKind = rsIDKind
    
End Function

Private Sub Class_Terminate()
    If Not (mobjIDCard Is Nothing) Then Call mobjIDCard.SetEnabled(False)
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
    
    If Not (mobjSquareCard Is Nothing) Then Set mobjSquareCard = Nothing
End Sub
