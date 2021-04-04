Attribute VB_Name = "mdlNewQuery"
Option Explicit

'新定义数据类型
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门码 As String
    部门 As String
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'系统控制相关变量
Public gstrPrivs As String                  '当前用户具有的当前模块的功能
Public glngSys As Long
Public gfrmMain As Object
Public gbytCardNOLen As Long
Public UserInfo As TYPE_USER_INFO
Public gstrUnitName As String               '用户单位名称
Public gstrServerName As String
Public gblnBeginTrans As Boolean
'============医保参数=====================
Public gblnInsure As Boolean '是否连接医保
Public gintInsure As Integer
Public gclsInsure As New clsInsure '用于医保接口，周燕川添加
Public gstrConnect As String

'------------周燕川
'系统参数

Public gblnShowCard As Boolean '是否明文显示卡号
Public gblnDailyTime As Boolean '日报时间允许
Public gblnBill挂号 As Boolean '是否严格控制票据
Public gbyt挂号 As Byte '挂号票据号码长度
Public grs挂号诊室 As ADODB.Recordset   '67045
'------------周燕川


'本机参数
 '自动开启的输入法
Public gstrIme As String
 '挂号领用ID
Public glng挂号ID As Long

'其它参数
Public gint号长 As Integer '号别长度


'系统临时变量
Public gstrSQL As String
Public gRs As New ADODB.Recordset

Const MAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

'--------------
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

'------------------------



Public Sub DrawLine(pic As PictureBox, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional ByVal ForeColor As Long = 0, Optional ByVal DrawStyle As Byte, Optional ByVal LineWidth As Byte = 1)
    '在(X1,Y1),(X2,Y2)之间使用ForeColor色画一直线
    Dim lngSaveForeColor As Long
    Dim bytSaveLineWidth As Byte
    
    lngSaveForeColor = pic.ForeColor
    bytSaveLineWidth = pic.DrawWidth
    pic.ForeColor = ForeColor
    pic.DrawStyle = DrawStyle
    pic.DrawWidth = LineWidth
    pic.Line (X2, Y2)-(X1, Y1)
    pic.ForeColor = lngSaveForeColor
    pic.DrawWidth = bytSaveLineWidth
End Sub

Public Sub DrawText(pic As PictureBox, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0, Optional ByVal Rotate As Integer = 0)
    '在(X,Y)处输出Text文本
    Dim lngSaveForeColor As Long
    Dim objFont As New clsRotateFont '旋转字体对象
            
    With pic
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        .CurrentX = X
        .CurrentY = Y
        If Rotate <> 0 Then
            Set objFont = New clsRotateFont
            Set objFont.LogFont = New StdFont
            objFont.LogFont.Name = .FontName
            objFont.LogFont.Size = .FontSize
            objFont.Rotation = Rotate
            objFont.Output pic, .CurrentX, .CurrentY, Text
        Else
            pic.Print Text
        End If
        
        .ForeColor = lngSaveForeColor
    End With
    Set objFont = Nothing
End Sub

Public Sub ResizeControl(obj As Object, ByVal X As Single, ByVal Y As Single, ByVal cx As Single, ByVal cy As Single)
    On Error Resume Next
    obj.Left = X
    obj.Top = Y
    obj.Width = cx
    obj.Height = cy
End Sub


Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "日期"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "日期时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "可打印字符"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/."":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function

Public Function MaxValue(ByVal Table As String, ByVal Field As String) As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    MaxValue = ""
    Set rs = zlDatabase.OpenSQLRecord("select max(" & Field & ") from " & Table, "mdlNewQuery")
'    Set rs = OpenRecord(rs, "select max(" & Field & ") from " & Table, "mdlNewQuery")
    If rs.BOF = False Then MaxValue = IIf(IsNull(rs.Fields(0).Value), 0, rs.Fields(0).Value)
    CloseRecord rs
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub CloseRecord(rs As ADODB.Recordset)
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub

Public Function SaveLvwItem(lvwObj As Object) As String
    If Not (lvwObj.SelectedItem Is Nothing) Then SaveLvwItem = lvwObj.SelectedItem.Key
End Function

Public Sub RestoreLvwItem(lvwObj As Object, svrKey As String)
    On Error GoTo EndP
    If lvwObj.ListItems.Count > 0 Then
        If Not (lvwObj.ListItems(svrKey) Is Nothing) Then
            lvwObj.ListItems(svrKey).Selected = True
        End If
    End If
    Exit Sub
EndP:
    If lvwObj.ListItems.Count > 0 Then lvwObj.ListItems(1).Selected = True
End Sub

Public Function NextValue(ByVal Table As String, ByVal Field As String) As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    NextValue = 0
    
    Set rs = zlDatabase.OpenSQLRecord("select nvl(max(" & Field & "),0) from " & Table, "mdlNewQuery")
    If rs.BOF = False Then NextValue = IIf(IsNull(rs.Fields(0).Value), 1, rs.Fields(0).Value + 1)
    CloseRecord rs
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub DrawPicture(pic As Object, objPic As StdPicture, ByVal W As Long, ByVal H As Long)
'功能：在PictureBox中央按适当比例画一幅图
'参数：W,H=要作图的尺寸
    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    If W <= pic.ScaleWidth And H <= pic.ScaleHeight Then
        lngW = W: lngH = H
    Else
        sngW = W / pic.ScaleWidth
        sngH = H / pic.ScaleHeight
        If sngW > sngH Then
            lngW = W / sngW: lngH = H / sngW
        Else
            lngW = W / sngH: lngH = H / sngH
        End If
    End If
    
    pic.Cls
    On Error Resume Next
    pic.PaintPicture objPic, (pic.ScaleWidth - lngW) / 2, (pic.ScaleHeight - lngH) / 2, lngW, lngH
End Sub

Public Sub PlayFlash(pic As PictureBox, ObjPlay As Object, strFile As String, ByVal W As Long, ByVal H As Long)

    Dim lngW As Long, lngH As Long
    Dim sngW As Single, sngH As Single
    
    If W <= pic.ScaleWidth And H <= pic.ScaleHeight Then
        lngW = W: lngH = H
    Else
        sngW = W / pic.ScaleWidth
        sngH = H / pic.ScaleHeight
        If sngW > sngH Then
            lngW = W / sngW: lngH = H / sngW
        Else
            lngW = W / sngH: lngH = H / sngH
        End If
    End If
    
    On Error Resume Next
    ObjPlay.Left = (pic.ScaleWidth - lngW) / 2
    ObjPlay.Top = (pic.ScaleHeight - lngH) / 2
    ObjPlay.Width = lngW - 15
    ObjPlay.Height = lngH - 15
        
    ObjPlay.BackgroundColor = -1
    ObjPlay.BGColor = -1
    ObjPlay.Movie = strFile
    ObjPlay.Playing = True
    ObjPlay.Loop = True

End Sub

Public Function WriteBinByFile(strFile As String, objField As Field) As Boolean
    Const conChunkSize As Integer = 10240
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim intBolcks As Integer, intFile, i As Long
    Dim arrBin() As Byte
    
    On Error GoTo errH
    
    intFile = FreeFile
    Open strFile For Binary Access Read As intFile
    lngFileSize = LOF(intFile)
    
    lngModSize = lngFileSize Mod conChunkSize
    intBolcks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    objField.Value = Null
    For i = 0 To intBolcks
        If i = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        ReDim arrBin(lngCurSize - 1) As Byte
        Get intFile, , arrBin()
        objField.AppendChunk arrBin()
    Next
    Close intFile
    WriteBinByFile = True
    Exit Function
errH:
    Close intFile
End Function

Public Function ReadPicByFieldNew(ByVal lngNo As Long) As StdPicture
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    On Error GoTo errH
    
    
    
'    intFile = FreeFile
    strFile = CurDir & "\zlNewPicture" & Timer & ".pic"
    
    strFile = Sys.ReadLob(glngSys, 28, lngNo, strFile, 0, 0)
    
'    Open strFile For Binary As intFile
'
'    lngFileSize = objField.ActualSize
'    ReDim arrBin(lngFileSize - 1) As Byte
'    arrBin() = objField.GetChunk(lngFileSize)
'    Put intFile, , arrBin()
'    Close intFile
    
    Set ReadPicByFieldNew = VB.LoadPicture(strFile)
    
    If Dir(strFile) <> "" And strFile <> "" Then Call Kill(strFile)
    Exit Function
errH:
'    Close intFile
    If Dir(strFile) <> "" And strFile <> "" Then Call Kill(strFile)
End Function


Public Function CreateFileByFieldNew(ByVal lngNo As Long, strFile As String) As Boolean
    Dim lngFileSize As Long
    Dim arrBin() As Byte
    Dim intFile As Integer

    On Error GoTo errH
    
    strFile = Sys.ReadLob(glngSys, 28, lngNo, strFile, 0, 0)
    
'    intFile = FreeFile
'    Open strFile For Binary As intFile
'
'    lngFileSize = objField.ActualSize
'    ReDim arrBin(lngFileSize - 1) As Byte
'    arrBin() = objField.GetChunk(lngFileSize)
'    Put intFile, , arrBin()
'    Close intFile

    CreateFileByFieldNew = True
    Exit Function
errH:
    Close intFile
    Kill strFile
End Function

Public Function ReadFlashByFieldNew(ByVal lngNo As Long) As String
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    On Error GoTo errH
    
'    intFile = FreeFile
    strFile = CurDir & "\zlFlash" & Timer & ".swf"
    
    
    strFile = Sys.ReadLob(glngSys, 28, lngNo, strFile, 0, 0)
'
'    Open strFile For Binary As intFile
'
'    lngFileSize = objField.ActualSize
'    ReDim arrBin(lngFileSize - 1) As Byte
'    arrBin() = objField.GetChunk(lngFileSize)
'    Put intFile, , arrBin()
'    Close intFile
    
    ReadFlashByFieldNew = strFile
    Exit Function
errH:
    Close intFile
End Function

Public Function ReadMusicByField(objField As Field) As String
    Dim lngFileSize As Long, arrBin() As Byte
    Dim strFile As String, intFile As Integer
    
    On Error GoTo errH
    
    intFile = FreeFile
    strFile = CurDir & "\zlMusic" & Timer & ".mid"
    
    Open strFile For Binary As intFile
    
    lngFileSize = objField.ActualSize
    ReDim arrBin(lngFileSize - 1) As Byte
    arrBin() = objField.GetChunk(lngFileSize)
    Put intFile, , arrBin()
    Close intFile
    
    ReadMusicByField = strFile
    Exit Function
errH:
    Close intFile
End Function


Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
'功能: 清除指定网格的指定行指定列的数据
'参数: obj=要操作的网格控件
'      intRow=要清除的行号
'      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Public Sub AddColumn(obj As Object, ByVal Tital As String, ByVal ColWidth As Single, ByVal ColAlignment As Byte)
    obj.Cols = obj.Cols + 1
    obj.TextMatrix(0, obj.Cols - 1) = Tital
    obj.ColWidth(obj.Cols - 1) = ColWidth
    obj.ColAlignment(obj.Cols - 1) = ColAlignment
End Sub

Public Function SaveFlash(ByVal strFile As String, ByVal byt性质 As Byte, ByVal byt类型 As Byte, Optional ByVal vKey As Long = 0, Optional ByVal vOldName As String = "") As Long
    '保存ShowFlash图片文件到数据库中
    Dim lngKey As Long
    Dim strSQL As String
    Dim vFlashHeader As FLASHHEADER
    
    vFlashHeader = GetFlashHeader(strFile)
    Select Case vFlashHeader.intIsFlashMovie
    Case -1
        MsgBox "指定的FlashMovie文件找不到！", vbInformation, gstrSysName
        Exit Function
    Case 0
        MsgBox "指定的文件不是FlashMovie文件！", vbInformation, gstrSysName
        Exit Function
    Case 2
        MsgBox "读取FlashMovie文件时发生未知错误！", vbInformation, gstrSysName
        Exit Function
    End Select
            
    If vKey = 0 Then
        lngKey = NextValue("咨询图片元素", "序号")
    Else
        lngKey = vKey
    End If
    
    strSQL = "zl_咨询图片元素_Insert("
    strSQL = strSQL & lngKey
    strSQL = strSQL & ",'" & IIf(vOldName = "", StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), "") & "'"
    strSQL = strSQL & "," & byt性质
    strSQL = strSQL & "," & byt类型
    strSQL = strSQL & "," & vFlashHeader.lMWidth
    strSQL = strSQL & "," & vFlashHeader.lMHeight
    strSQL = strSQL & ",To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
    
    On Error GoTo errHand
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSQL, "mdlNewQuery")
    Call Sys.SaveLob(glngSys, 28, lngKey, strFile, 0)
    gcnOracle.CommitTrans

    SaveFlash = lngKey
    Exit Function
errHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = -1 Then Resume
    If ErrCenter() = 1 Then Resume
End Function

Public Function SavePicture(ByVal strFile As String, imgTmp As Object, ByVal byt性质 As Byte, ByVal byt类型 As Byte, Optional ByVal vKey As Long = 0, Optional ByVal vOldName As String = "") As Long
    Dim lngKey As Long
    Dim objMap As StdPicture
    Dim mvFile As WIN32_FIND_DATA
    Dim strSQL As String
    On Error Resume Next

    Set objMap = VB.LoadPicture(strFile)
    Call FindFirstFile(strFile, mvFile)
        
    imgTmp.ListImages.Clear
    imgTmp.ListImages.Add , , objMap
            
    If Not StrIsValid(StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), 30) Then Exit Function
        
    If vKey = 0 Then
        lngKey = NextValue("咨询图片元素", "序号")
    Else
        lngKey = vKey
    End If
    
    strSQL = "zl_咨询图片元素_Insert("
    strSQL = strSQL & lngKey
    strSQL = strSQL & ",'" & IIf(vOldName = "", StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), "") & "'"
    strSQL = strSQL & "," & byt性质
    strSQL = strSQL & "," & byt类型
    strSQL = strSQL & "," & imgTmp.ImageWidth
    strSQL = strSQL & "," & imgTmp.ImageHeight
    strSQL = strSQL & ",To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
    
    On Error GoTo errHand
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSQL, "mdlNewQuery")
    Call Sys.SaveLob(glngSys, 28, lngKey, strFile, 0)
    gcnOracle.CommitTrans

    SavePicture = lngKey
    Exit Function
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Function

Public Function SaveMidea(ByVal strFile As String, ByVal byt性质 As Byte, ByVal byt类型 As Byte, Optional ByVal vKey As Long = 0, Optional ByVal vOldName As String = "") As Long
    Dim lngKey As Long
    Dim strSQL As String
    
    If Not StrIsValid(StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), 30) Then Exit Function
        
    If vKey = 0 Then
        lngKey = NextValue("咨询图片元素", "序号")
    Else
        lngKey = vKey
    End If
    
    strSQL = "zl_咨询图片元素_Insert("
    strSQL = strSQL & lngKey
    strSQL = strSQL & ",'" & IIf(vOldName = "", StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), "") & "'"
    strSQL = strSQL & "," & byt性质
    strSQL = strSQL & "," & byt类型
    strSQL = strSQL & ",Null"
    strSQL = strSQL & ",Null"
    strSQL = strSQL & ",To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'))"
    
    On Error GoTo errHand
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSQL, "mdlNewQuery")
        
    Call Sys.SaveLob(glngSys, 28, lngKey, strFile, 0)
    
    gcnOracle.CommitTrans
    
    SaveMidea = lngKey
    Exit Function
errHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = -1 Then Resume
End Function

Public Function GetFileName(ByVal PicOrder As Long, W As Single, H As Single) As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "select B.类型,B.名称,B.宽度,B.高度 from 咨询图片元素 B where B.序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", PicOrder)
    If rs.BOF = False Then
        W = IIf(IsNull(rs!宽度), 0, rs!宽度) * Screen.TwipsPerPixelX
        H = IIf(IsNull(rs!高度), 0, rs!高度) * Screen.TwipsPerPixelY
        Select Case IIf(IsNull(rs!类型), 0, rs!类型)
        Case 0
            GetFileName = IIf(IsNull(rs!名称), "", App.Path & "\图形\" & rs!名称 & ".pic")
        Case 1
            GetFileName = IIf(IsNull(rs!名称), "", App.Path & "\图形\" & rs!名称 & ".ico")
        Case 2
            GetFileName = IIf(IsNull(rs!名称), "", App.Path & "\图形\" & rs!名称 & ".swf")
        End Select
    End If
    CloseRecord rs
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub InsertGrid(objDraw As Object, ByVal lngNo As Long, ByVal Aligment As Byte, ByVal NextY As Single, varWidth As Single, varHeight As Single)
'功能:显示用户自定义表格
'参数:objDraw           作图目标
'     lngNo             表格序号
'     NextX             表格左横坐标
'     NextY             表格上纵坐标
'     varWidth          返回表格所占的宽度
'     varHeight         返回表格所占的高度

    Dim j As Long
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    gstrSQL = "select 序号,名称,列数,列宽,行数,行高,合并列,合并行,颜色 from 咨询表格元素 where 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", lngNo)
    If rs.BOF = False Then
        j = objDraw.NextVsfIndex
        Call objDraw.AddPageItemGrid(j, Aligment, NextY, rs!行数, rs!列数, rs!行高, rs!列宽, IIf(IsNull(rs!合并行), "", rs!合并行), IIf(IsNull(rs!合并列), "", rs!合并列), varWidth, varHeight)
        
        gstrSQL = "select 表号,行号,列号,内容,对齐,颜色,字体 from 咨询表格内容 where 表号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", lngNo)
        If rs.BOF = False Then
            While Not rs.EOF
                objDraw.Row(j) = rs!行号 - 1
                objDraw.Col(j) = rs!列号 - 1
                objDraw.TextMatrix(j, objDraw.Row(j), objDraw.Col(j)) = IIf(IsNull(rs!内容), "", rs!内容)
                objDraw.CellAlignment(j) = IIf(IsNull(rs!对齐), 9, rs!对齐)
                objDraw.CellForeColor(j) = IIf(IsNull(rs!颜色), 0, rs!颜色)
                objDraw.CellFontName(j) = Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(0)
                objDraw.CellFontSize(j) = Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(1)
                objDraw.CellFontBold(j) = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(2) = True, True, False)
                objDraw.CellFontItalic(j) = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(3) = True, True, False)
                objDraw.CellFontStrikethru(j) = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(4) = True, True, False)
                objDraw.CellFontUnderline(j) = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(5) = True, True, False)
                rs.MoveNext
            Wend
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetInterval() As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    GetInterval = Val(GetPara("广告播放间隔", "5"))
    If GetInterval < 5 Then GetInterval = 5
    If GetInterval > 300 Then GetInterval = 300
    
    CloseRecord rs
    Exit Function
errHand:
    CloseRecord rs
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Function

Public Sub CalcAutoColWidth(msf As Object, ByVal Col As Long)
    Dim i As Long
    Dim W As Single
    
    On Error Resume Next
    W = 15
    For i = 0 To msf.Cols - 2
        If i <> Col And msf.ColHidden(i) = False Then
            W = W + msf.ColWidth(i) + 15
        End If
    Next
    msf.ColWidth(Col) = msf.Width - W + 30
End Sub

Public Sub EnablePageButton(msf As Object, ByVal CurPos As Long, ByVal DataRows As Long, UpButton As Object, DownButton As Object)
    Dim i As Long
    
    i = CalcVisibleRow(msf)
    
    UpButton.Enabled = False
    DownButton.Enabled = False
        
    If DataRows > i Then
        If CurPos + i < DataRows Then DownButton.Enabled = True
        If CurPos - i > 0 Then UpButton.Enabled = True
        If CurPos <> 1 Then UpButton.Enabled = True
    End If
    
End Sub

Private Function CalcVisibleRow(msf As Object) As Long
    Dim i As Long
    
    CalcVisibleRow = 0
    For i = 1 To msf.Rows - 1
        If msf.RowIsVisible(i) Then CalcVisibleRow = CalcVisibleRow + 1
    Next
    CalcVisibleRow = IIf(CalcVisibleRow > 0, CalcVisibleRow - 1, 0)
End Function

Public Sub TurnToPage(msf As Object, ByVal bytMode As Integer, CurPos As Long)
    Dim i As Long
    
    i = CalcVisibleRow(msf)
    CurPos = CurPos + i * bytMode
    CurPos = IIf(CurPos < 0, 1, CurPos)
    msf.TopRow = CurPos
End Sub

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Sub NextTvwPos(tvwObj As Object, ByVal vIndex As Long)
        
    If tvwObj.Nodes.Count > 0 Then
        vIndex = IIf(tvwObj.Nodes.Count > vIndex, vIndex, tvwObj.Nodes.Count)
        tvwObj.Nodes(vIndex).Selected = True
        tvwObj.Nodes(vIndex).EnsureVisible
    End If
End Sub

Public Sub AdjustOrder(lvwObject As Object, ByVal ItmxNo As Long)
    Dim i As Long
    
    For i = 1 To lvwObject.ListItems.Count
        If ItmxNo = 0 Then
            lvwObject.ListItems(i).Text = i
        Else
            lvwObject.ListItems(i).SubItems(1) = i
        End If
    Next
End Sub

Public Function CheckMenuLimit(ByVal UpKey As Long) As Boolean
    
    On Error GoTo errHand
    
    gstrSQL = "select nvl(count(*),0) from 咨询页面排列 where " & IIf(UpKey = 0, "父序号 is null or 父序号=0", "父序号=[1]")
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", UpKey)
    If gRs.BOF = False Then CheckMenuLimit = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub CheckPicture()
'功能:检查本地图片,并更新本地图片

    Dim strFileName As String
    Dim vFileData As New FileSystemObject
            
    '1.检查图形目录是否存在
    On Error Resume Next
    vFileData.CreateFolder App.Path & "\图形"
    
    On Error GoTo errHand
    '2.检查本系统中可能使用到的图片
    gstrSQL = "select 序号,性质,名称,类型,宽度,高度,固定,修改日期 from 咨询图片元素"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery")
    If gRs.BOF = False Then
        While Not gRs.EOF
            strFileName = IIf(IsNull(gRs!名称), "", gRs!名称)
            If strFileName <> "" Then
                Call CheckFileNew(strFileName, IIf(IsNull(gRs!类型), 0, gRs!类型), gRs!序号, gRs!修改日期, vFileData)
            End If
            gRs.MoveNext
        Wend
    End If
    
    '3.检查医生的照片
    gstrSQL = "select B.ID, B.姓名  from 咨询专家清单 A,人员表 B where A.人员id=B.id  And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) "
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery")
    If gRs.BOF = False Then
        While Not gRs.EOF
            strFileName = IIf(IsNull(gRs!姓名), "", gRs!姓名)
            If strFileName <> "" Then
                strFileName = App.Path & "\图形\" & strFileName & ".pic"
                If Dir(strFileName) <> "" Then vFileData.DeleteFile (strFileName)

                Call Sys.ReadLob(glngSys, 16, Val(Nvl(gRs!ID)), strFileName)
                Call SetFileDateTime(strFileName, Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS"))
            End If
            gRs.MoveNext
        Wend
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

'Private Sub CreateFile(ByVal strFileName As String, objField As Field, mvarDateTime As String)
''功能:创建本地图形文件
'    If CreateFileByField(objField, strFileName) Then Call SetFileDateTime(strFileName, mvarDateTime)
'End Sub

Public Sub InitInternal()
    frmMainQuery.mvarHomeInternal = 0
End Sub

Public Sub CheckFileNew(ByRef strName As String, ByVal byt类型 As Byte, lngNo As Long, objFieldDate As Field, flObj As FileSystemObject)
    
    Select Case byt类型
    Case 0              '所有图片
        strName = App.Path & "\图形\" & strName & ".pic"
    Case 1              '图标
        strName = App.Path & "\图形\" & strName & ".ico"
    Case 2              'FLASH
        strName = App.Path & "\图形\" & strName & ".swf"
    Case 3              'Media
        strName = App.Path & "\图形\" & strName & ".mid"
    End Select
    If Dir(strName) <> "" Then
        '图形文件存在,将检查和当前数据库中的图形是否一致(通过修改时间)
        If Format(flObj.GetFile(strName).DateLastModified, "yyyy-mm-dd hh:mm") <> Format(objFieldDate, "yyyy-mm-dd hh:mm") Then
            flObj.DeleteFile (strName)
            
            strName = Sys.ReadLob(glngSys, 28, lngNo, strName, 0, 0)
'            Call CreateFile(strName, objFieldMedia, Format(objFieldDate, "yyyy-mm-dd hh:mm:ss"))
        End If
    Else
        '图形文件不存在,将创建新的图形文件
        strName = Sys.ReadLob(glngSys, 28, lngNo, strName, 0, 0)
'        Call CreateFile(strName, objFieldMedia, Format(objFieldDate, "yyyy-mm-dd hh:mm:ss"))
    End If
End Sub

Public Sub SelectRow(mshObject As Object, Optional ByVal mvarRow As Long = -1, Optional ByVal mvarFore As Boolean = False)
    Dim i As Integer
    Dim blnPre As Boolean
    Dim intRow As Integer
    Dim intCol As Integer
    
    With mshObject
        blnPre = .Redraw
        intRow = .Row
        intCol = .Col
        .Redraw = False
        
        If mvarRow <> -1 Then mshObject.Row = mvarRow
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = .BackColorSel
            If mvarFore Then .CellForeColor = .ForeColorSel
        Next
        
        .Row = intRow
        .Col = intCol
        .Redraw = blnPre
    End With
End Sub

Public Sub UnSelectRow(mshObject As Object, Optional lngColorSave As Long = 0, Optional ByVal mvarRow As Long = -1, Optional ByVal mvarFore As Boolean = False)
    Dim i As Integer
    Dim blnPre As Boolean
    
    With mshObject
        blnPre = .Redraw
        .Redraw = False
                        
        If mvarRow <> -1 Then mshObject.Row = mvarRow
        For i = 0 To .Cols - 1
            .Col = i
            .CellBackColor = .BackColor
            If mvarFore Then .CellForeColor = lngColorSave
        Next
        
        .Redraw = blnPre
    End With
End Sub

Public Sub SortArray(ByRef objArray As Variant)
    '---------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '---------------------------------------------------------------------------------------
    
    Dim lngLoop As Long
    Dim objTmp As String
    Dim blnFlag As Boolean
    
    blnFlag = True
    Do While blnFlag
        blnFlag = False
        For lngLoop = LBound(objArray) To UBound(objArray) - 1
            If objArray(lngLoop) > objArray(lngLoop + 1) Then
                blnFlag = True
                objTmp = objArray(lngLoop)
                objArray(lngLoop) = objArray(lngLoop + 1)
                objArray(lngLoop + 1) = objTmp
            End If
        Next
    Loop
    
End Sub

Public Function CheckIdentify(ByRef lng病人ID As Long, ByRef lng主页id As Long) As Boolean
    '---------------------------------------------------------------------------------------------
    '
    '功能:验证身份
    '参数:strPsw医保卡密码
    '
    '---------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim strPass As String
    Dim strCard As String
    
    CheckIdentify = False
    
'    Select Case gintInsure
'    Case 80
        strCard = "0"
        CheckIdentify = frmIdentify泸州.ShowForm(strPass, strCard)
        
        '测试数据
'        CheckIdentify = True
'        strPass = "1111"
'        strCard = "0"

'    Case Else
'       '不支持此操作
'
'    End Select
    
    If CheckIdentify = False Then Exit Function
    
    CheckIdentify = False
    
    On Error GoTo errHand
    
    gintInsure = 0
    strTmp = gclsInsure.Identify2(strCard, strPass, 2, lng病人ID, gintInsure)
    
    If Trim(strTmp) = "" Then Exit Function
    
    '只能查住院病人
    If lng病人ID > 0 Then
        
        gstrSQL = "" & _
        " Select to_char(A.入院日期,'yyyy-mm-dd') as 入院日期,to_char(A.出院日期,'yyyy-mm-dd') as 出院日期,A.病人id,A.主页id " & _
        " From 病案主页 A,病人信息 B " & _
        " Where A.病人id=B.病人id and B.病人ID=[1] " & _
        " Union ALL " & _
        " Select '门诊费用' as 入院日期,'门诊费用' as 出院日期,病人id,0 as 主页id " & _
        " From 门诊费用记录 C " & _
        " Where C.病人ID=[1] AND rownum<2 "
        
        gstrSQL = "Select 入院日期,出院日期,病人id,主页id From (" & gstrSQL & " ) AA Order by AA.入院日期 asc,AA.主页id desc"
        
        gstrSQL = "" & _
        " SELECT rownum as No,D.入院日期,D.出院日期,D.病人id,D.主页id " & _
        " FROM ( " & gstrSQL & "  ) D"
        
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "费用查询", lng病人ID)
        If gRs.RecordCount > 0 Then
            CheckIdentify = frmSelect.ShowSelect(gRs, lng病人ID, lng主页id)
        Else
            lng病人ID = 0
        End If
        
        If CheckIdentify = False Then lng病人ID = 0
    End If
    
    CheckIdentify = (lng病人ID > 0)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatientInsure(ByVal lng病人ID As Long, Optional ByVal lng主页id As Long) As Long
    '------------------------------------------------------------------------------------------------------------------
    '功能：获取病人的险类
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strSQL = "Select Nvl(B.险类,A.险类) as 险类 " & _
        " From 病人信息 A,病案主页 B,医疗付款方式 C" & _
        " Where A.病人ID=" & lng病人ID & " And A.病人ID=B.病人ID(+)" & _
        " And B.主页ID(+)=" & lng主页id & " And A.医疗付款方式=C.名称(+)"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlNewQuery")
    If rs.BOF = False Then
        GetPatientInsure = zlCommFun.Nvl(rs("险类").Value, 0)
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：本级ID，表名
    '输出参数：成功返回 下级最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级序号 is null " & strWhere & " connect by prior 页面序号=上级序号"
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级序号=" & strID & strWhere & " connect by prior 页面序号=上级序号"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlnewQuery")
    If rsTemp.RecordCount = 0 Then
        GetDownCodeLength = 0
    Else
        GetDownCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetDownCodeLength = 0
End Function

Public Function GetLocalCodeLength(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：上级ID，表名
    '输出参数：成功返回 最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级序号 is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级序号=" & str上级ID & strWhere
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlnewQuery")
    
    If rsTemp.RecordCount = 0 Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Public Function GetParentCode(ByVal str上级ID As String, ByVal strTableName As String) As String
    '功能描述：读取上级编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 上级编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select 编码 from " & strTableName & " where 页面序号=" & str上级ID
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlnewQuery")
    
    If rsTemp.RecordCount = 0 Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("编码").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetMaxLocalCode(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '功能描述：根据指定表的上级ID 读取本级的最大编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 最大编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strCode As String, strAllCode As String
    Dim intLength   As Integer
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select max(to_number(编码))+1 as MaxCode from " & strTableName & " where 上级序号 is null" & strWhere
    Else
        strSQL = "select nvl(max(to_number(编码)),0)+1 as MaxCode from " & strTableName & " where 上级序号=" & str上级ID & strWhere
    End If
    intCode = GetLocalCodeLength(str上级ID, strTableName, strWhere)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlnewQuery")
    
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str上级ID, strTableName)) + 1)
    If GetMaxLocalCode = "" Then GetMaxLocalCode = "1"
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Sub EnterFocus(ByVal obj As Object)
    
    On Error GoTo errHand
    
    obj.SetFocus
    
    Exit Sub
    
errHand:
 
End Sub

Public Function GetPara(ByVal varPara As Variant, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************
    
    On Error GoTo errHand
    If blnNotCache Then Call zlDatabase.ClearParaCache
    GetPara = zlDatabase.GetPara(varPara, glngSys, 1536, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：设置指定的参数值
    '参数：varPara=参数号或参数名，以数字或字符类型传入区分
    '      strValue=要设置的参数值
    '      lngModual=使用该参数的模块号，如1230
    '      blnPrivate=该参数是否用户私有参数
    '返回：设置是否成功
    '******************************************************************************************************************

    On Error GoTo errH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, glngSys, 1536)

    Exit Function
    
errH:

End Function

Public Function GetNodeCheckSQL(Optional ByVal strNodeField As String = "站点") As String
    '******************************************************************************************************************
    '功能：获取站点限制的ＳＱＬ条件语句
    '参数：strNodeField=
    '返回：
    '******************************************************************************************************************
    Dim strNodeNo As String
    
    strNodeNo = gstrNodeNo
    If strNodeNo = "" Then strNodeNo = "-"
    
    GetNodeCheckSQL = " Nvl(" & strNodeField & ", '" & strNodeNo & "') = '" & strNodeNo & "' "
    
End Function

Public Function zlGetFeeFields(Optional strTableName As String = "门诊费用记录", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定表的值
    '入参：strTableName:如:门诊费用记录;住院费用记录;....
    '      blnReadDatabase-从数据库中读取
    '出参：
    '返回：字段集
    '编制：刘兴洪
    '日期：2010-03-10 10:41:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo errHand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "挂号安排"
        'zlGetFeeFields = "ID,号类,号码,科室ID,项目ID,医生姓名,医生ID,限号数,限约数,周日,周一,周二,周三,周四,周五,周六,病案必须,分诊方式,序号控制,开始时间,终止时间"
        zlGetFeeFields = "ID,号类,号码,科室ID,项目ID,医生姓名,医生ID,周日,周一,周二,周三,周四,周五,周六,病案必须,分诊方式,序号控制,开始时间,终止时间"
        Exit Function
    Case "门诊费用记录"
        zlGetFeeFields = "" & _
        "Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 记帐单id, 病人id, 医嘱序号, 门诊标志, 记帐费用, " & _
        "姓名, 性别, 年龄, 标识号, 付款方式, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, 付数, 发药窗口, 数次, " & _
        "加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, 开单部门id, 开单人, " & _
        "发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, 结帐id, 结帐金额, " & _
        "保险大类id, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
        Exit Function
    Case "住院费用记录"
        zlGetFeeFields = "" & _
         " Id, 记录性质, No, 实际票号, 记录状态, 序号, 从属父号, 价格父号, 多病人单, 记帐单id, 病人id, 主页id, 医嘱序号, " & _
         " 门诊标志, 记帐费用, 姓名, 性别, 年龄, 标识号, 床号, 病人病区id, 病人科室id, 费别, 收费类别, 收费细目id, 计算单位, " & _
         " 付数, 发药窗口, 数次, 加班标志, 附加标志, 婴儿费, 收入项目id, 收据费目, 标准单价, 应收金额, 实收金额, 划价人, " & _
         " 开单部门id, 开单人, 发生时间, 登记时间, 执行部门id, 执行人, 执行状态, 执行时间, 结论, 操作员编号, 操作员姓名, " & _
         " 结帐id , 结帐金额, 保险大类ID, 保险项目否, 保险编码, 费用类型, 统筹金额, 是否上传, 摘要, 是否急诊"
         Exit Function
    Case "病人结帐记录"
        zlGetFeeFields = "Id, No, 实际票号, 记录状态, 中途结帐, 病人id, 操作员编号, 操作员姓名, 收费时间, 开始日期, 结束日期, 备注"
        Exit Function
    Case "病人预交记录"
        zlGetFeeFields = "" & _
        " Id, 记录性质, No, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, " & _
        " 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交, 结帐id, 缴款, 找补"
        Exit Function
    Case "人员表"
        zlGetFeeFields = "" & _
        "Id, 编号, 姓名, 简码, 身份证号, 出生日期, 性别, 民族, 工作日期, 办公室电话, 电子邮件, 执业类别, 执业范围, " & _
        "管理职务, 专业技术职务, 聘任技术职务, 学历, 所学专业, 留学时间, 留学渠道, 接受培训, 科研课题, 个人简介, 建档时间, " & _
        "撤档时间, 撤档原因, 别名, 站点"
        Exit Function
    Case "票据领用记录"
        zlGetFeeFields = "ID,票种,领用人,前缀文本,开始号码,终止号码,使用方式,登记时间,使用时间,登记人,当前号码,剩余数量,批次,核对人,核对时间,核对结果,核对模式,备注"
        Exit Function
    Case "票据使用明细"
        zlGetFeeFields = "ID,票种,号码,性质,原因,领用ID,打印ID,回收次数,使用时间,使用人,核对人,核对时间,核对结果,备注"
        Exit Function
    Case "人员缴款记录"
        zlGetFeeFields = "ID,单据ID,收款员,收款部门ID,结算方式,结算号,金额,摘要,截止时间,登记时间,登记人"
        Exit Function
    Case "咨询图片元素"
        zlGetFeeFields = "序号,性质,名称,类型,宽度,高度,固定,修改日期"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo errHand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID;"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取列信息", strTableName)
    strFileds = ""
    With rsTemp
        Do While Not .EOF
            strFileds = strFileds & "," & Nvl(!column_name)
            .MoveNext
        Loop
        If strFileds <> "" Then strFileds = Mid(strFileds, 2)
    End With
    If strFileds = "" Then strFileds = "*"
    zlGetFeeFields = strFileds
    Exit Function
errHand:
  zlGetFeeFields = "*"
  If ErrCenter() = 1 Then
    Resume
  End If
End Function

Public Function zlGetFullFieldsTable(Optional strTableName As String = "门诊费用记录", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张数据表中的字段.类似于Select Id,....
    '入参：bytHistory-0-不包含历史数据,1-仅包含历史数据,2-两都都包含( select * from tablename Union select * from Htablename)
    '      strWhere-条件
    '      blnSubTable-是否子表
    '      strAliasName-别名
    '出参：
    '返回：select ID ... From tableName Union ALL
    '编制：刘兴洪
    '日期：2010-03-10 11:19:11
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '无
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '仅历史
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '两者都包含
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & " " & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
End Function

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub


Public Function Exist门诊号(ByVal str门诊号 As String, Optional ByVal lng病人ID As Long) As Boolean
'功能：判断指定门诊号是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病人ID From 病人信息 Where 门诊号=[1] And 病人ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl9NewQuery", str门诊号, lng病人ID)
    If rsTmp.RecordCount > 0 Then Exist门诊号 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Exist病人ID(ByVal lng病人ID As Long) As Boolean
'功能：判断指定病人ID是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 病人ID From 病人信息 Where   病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl9NewQuery", lng病人ID)
    If rsTmp.RecordCount > 0 Then Exist病人ID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'                                简易挂号相关
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
      
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetRegistParaFont(ByVal strParaName As String, ByRef strMsg As String, ByRef strFontName As String, ByRef dblSize As Double, _
                                    ByRef dblColor As Double, ByRef blnBold As Boolean, ByRef blnItalic As Boolean) As Boolean
    ' --------------------------------------------------
    '获取简易挂号提示内容以及字体
    ' 入参 strParaName -参数名 strMsg -提示信息, strFontName-字体名称 dblSize-字体大小 dblColor -颜色
    '--------------------------------------------------
     Dim strRows() As String
     Dim strTmp As String, X As Long
     On Error GoTo hErr
     strTmp = GetPara(strParaName, "", True)
     If strTmp = "" Then Exit Function
    ''提示信息|字体|颜色|字体大小'
    strRows = Split(strTmp, "|")
    strMsg = strRows(0)
    If UBound(strRows) < 1 Then Exit Function
    strFontName = strRows(1)
    If IsNumeric(strRows(3)) Then dblSize = CDbl(strRows(3))
    If IsNumeric(strRows(2)) Then dblColor = CDbl(strRows(2))
    If IsNumeric(strRows(4)) Then blnBold = Val(strRows(4)) = 1
    If IsNumeric(strRows(5)) Then blnItalic = Val(strRows(5)) = 1
    GetRegistParaFont = True
    Exit Function
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Public Function SetRegistParaFont(ByVal strParName As String, ByVal strMsg As String, strFontName As String, _
                        dblSize As Double, dblColor As Double, blnBold As Boolean, blnItalic As Boolean)
    '设置简易挂号的提示内容以及字体
     ' 入参 strMsg -提示信息, strFontName-字体名称 dblSize-字体大小 dblColor -颜色
     Dim strTmp As String
     strTmp = strMsg & "|" & strFontName & "|" & dblColor & "|" & dblSize & "|" & IIf(blnBold, 1, 0) & "|" & _
            IIf(blnItalic, 1, 0)
     On Error GoTo hErr
     SetRegistParaFont = SetPara(strParName, strTmp)
     Exit Function
hErr:
     If ErrCenter() = 1 Then Resume
     SaveErrLog
End Function


Public Function GetFreeRegistBGColor(ByRef dblUpBgColor As Double, ByRef dblDownBgColor As Double) As Boolean
    '---------------------------------------------------
    '获取简易挂号 渐变的背景色
    '---------------------------------------------------
        Dim strRows() As String
        Dim strTmp As String, X As Long
        On Error GoTo hErr
        strTmp = GetPara("简单挂号背景色", "")
        If strTmp = "" Then Exit Function
   
        ''提示信息|字体|颜色|字体大小'
        strRows = Split(strTmp, "|")
        If IsNumeric(strRows(0)) Then dblUpBgColor = CDbl(strRows(0))
        If UBound(strRows) < 1 Then Exit Function
        If IsNumeric(strRows(1)) Then dblDownBgColor = CDbl(strRows(1))
        GetFreeRegistBGColor = True
  Exit Function
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Public Function SetFreeRegistBGColor(ByVal dblUpBgColor As Double, ByVal dblDownBgColor As Double) As Boolean
    '---------------------------------------------------
    '设置简易挂号 渐变的背景色
    '---------------------------------------------------
        On Error GoTo hErr
        SetFreeRegistBGColor = SetPara("简单挂号背景色", dblUpBgColor & "|" & dblDownBgColor)
  Exit Function
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Public Function GetRs挂号诊室() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号安排诊室
    '返回:挂号安排诊室的记录集
    '编制:刘兴洪
    '日期:2013-11-01 16:36:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo ErrHandle
    If grs挂号诊室 Is Nothing Then
         strSQL = "Select 号表ID,门诊诊室 From 挂号安排诊室"
        Set grs挂号诊室 = zlDatabase.OpenSQLRecord(strSQL, "今日就诊")
        Set GetRs挂号诊室 = grs挂号诊室
        Exit Function
    End If
    If grs挂号诊室.State <> 1 Then
         strSQL = "Select 号表ID,门诊诊室 From 挂号安排诊室"
        Set grs挂号诊室 = zlDatabase.OpenSQLRecord(strSQL, "今日就诊")
        Set GetRs挂号诊室 = grs挂号诊室
        Exit Function
    End If
    Set GetRs挂号诊室 = grs挂号诊室
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ReplaseSpecial(strTmp As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能               替换特殊字符
    '参数
    '                   需替换的字符
    '返回               需替换了特殊字符后的字串
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strSpecial As String
    Dim astrTmp() As String
    strSpecial = "'^‘^’^;^；^:^：^?^？^|^,^，^.^。^"""
    astrTmp = Split(strSpecial, "^")
    For intLoop = 0 To UBound(astrTmp)
        strTmp = Replace$(strTmp, astrTmp(intLoop), "")
    Next
    ReplaseSpecial = strTmp
    
End Function







