Attribute VB_Name = "mdlNewQuery"
Option Explicit

'�¶�����������
Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ������ As String
    ���� As String
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'ϵͳ������ر���
Public gstrPrivs As String                  '��ǰ�û����еĵ�ǰģ��Ĺ���
Public glngSys As Long
Public gfrmMain As Object
Public gbytCardNOLen As Long
Public UserInfo As TYPE_USER_INFO
Public gstrUnitName As String               '�û���λ����
Public gstrServerName As String
Public gblnBeginTrans As Boolean
'============ҽ������=====================
Public gblnInsure As Boolean '�Ƿ�����ҽ��
Public gintInsure As Integer
Public gclsInsure As New clsInsure '����ҽ���ӿڣ����ന���
Public gstrConnect As String

'------------���ന
'ϵͳ����

Public gblnShowCard As Boolean '�Ƿ�������ʾ����
Public gblnDailyTime As Boolean '�ձ�ʱ������
Public gblnBill�Һ� As Boolean '�Ƿ��ϸ����Ʊ��
Public gbyt�Һ� As Byte '�Һ�Ʊ�ݺ��볤��
Public grs�Һ����� As ADODB.Recordset   '67045
'------------���ന


'��������
 '�Զ����������뷨
Public gstrIme As String
 '�Һ�����ID
Public glng�Һ�ID As Long

'��������
Public gint�ų� As Integer '�ű𳤶�


'ϵͳ��ʱ����
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
    '��(X1,Y1),(X2,Y2)֮��ʹ��ForeColorɫ��һֱ��
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
    '��(X,Y)�����Text�ı�
    Dim lngSaveForeColor As Long
    Dim objFont As New clsRotateFont '��ת�������
            
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
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "����ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "�ɴ�ӡ�ַ�"
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
'���ܣ���PictureBox���밴�ʵ�������һ��ͼ
'������W,H=Ҫ��ͼ�ĳߴ�
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
'���ܣ����ı���ĵ��ı�ѡ��
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
'����: ���ָ�������ָ����ָ���е�����
'����: obj=Ҫ����������ؼ�
'      intRow=Ҫ������к�
'      intCol=Ҫ������к��б���Array(1,2,3),������������Ա�ʾΪArray()
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

Public Function SaveFlash(ByVal strFile As String, ByVal byt���� As Byte, ByVal byt���� As Byte, Optional ByVal vKey As Long = 0, Optional ByVal vOldName As String = "") As Long
    '����ShowFlashͼƬ�ļ������ݿ���
    Dim lngKey As Long
    Dim strSQL As String
    Dim vFlashHeader As FLASHHEADER
    
    vFlashHeader = GetFlashHeader(strFile)
    Select Case vFlashHeader.intIsFlashMovie
    Case -1
        MsgBox "ָ����FlashMovie�ļ��Ҳ�����", vbInformation, gstrSysName
        Exit Function
    Case 0
        MsgBox "ָ�����ļ�����FlashMovie�ļ���", vbInformation, gstrSysName
        Exit Function
    Case 2
        MsgBox "��ȡFlashMovie�ļ�ʱ����δ֪����", vbInformation, gstrSysName
        Exit Function
    End Select
            
    If vKey = 0 Then
        lngKey = NextValue("��ѯͼƬԪ��", "���")
    Else
        lngKey = vKey
    End If
    
    strSQL = "zl_��ѯͼƬԪ��_Insert("
    strSQL = strSQL & lngKey
    strSQL = strSQL & ",'" & IIf(vOldName = "", StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), "") & "'"
    strSQL = strSQL & "," & byt����
    strSQL = strSQL & "," & byt����
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

Public Function SavePicture(ByVal strFile As String, imgTmp As Object, ByVal byt���� As Byte, ByVal byt���� As Byte, Optional ByVal vKey As Long = 0, Optional ByVal vOldName As String = "") As Long
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
        lngKey = NextValue("��ѯͼƬԪ��", "���")
    Else
        lngKey = vKey
    End If
    
    strSQL = "zl_��ѯͼƬԪ��_Insert("
    strSQL = strSQL & lngKey
    strSQL = strSQL & ",'" & IIf(vOldName = "", StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), "") & "'"
    strSQL = strSQL & "," & byt����
    strSQL = strSQL & "," & byt����
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

Public Function SaveMidea(ByVal strFile As String, ByVal byt���� As Byte, ByVal byt���� As Byte, Optional ByVal vKey As Long = 0, Optional ByVal vOldName As String = "") As Long
    Dim lngKey As Long
    Dim strSQL As String
    
    If Not StrIsValid(StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), 30) Then Exit Function
        
    If vKey = 0 Then
        lngKey = NextValue("��ѯͼƬԪ��", "���")
    Else
        lngKey = vKey
    End If
    
    strSQL = "zl_��ѯͼƬԪ��_Insert("
    strSQL = strSQL & lngKey
    strSQL = strSQL & ",'" & IIf(vOldName = "", StrReverse(Mid(StrReverse(strFile), 5, InStr(StrReverse(strFile), "\") - 5)), "") & "'"
    strSQL = strSQL & "," & byt����
    strSQL = strSQL & "," & byt����
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
    
    gstrSQL = "select B.����,B.����,B.���,B.�߶� from ��ѯͼƬԪ�� B where B.���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", PicOrder)
    If rs.BOF = False Then
        W = IIf(IsNull(rs!���), 0, rs!���) * Screen.TwipsPerPixelX
        H = IIf(IsNull(rs!�߶�), 0, rs!�߶�) * Screen.TwipsPerPixelY
        Select Case IIf(IsNull(rs!����), 0, rs!����)
        Case 0
            GetFileName = IIf(IsNull(rs!����), "", App.Path & "\ͼ��\" & rs!���� & ".pic")
        Case 1
            GetFileName = IIf(IsNull(rs!����), "", App.Path & "\ͼ��\" & rs!���� & ".ico")
        Case 2
            GetFileName = IIf(IsNull(rs!����), "", App.Path & "\ͼ��\" & rs!���� & ".swf")
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
'����:��ʾ�û��Զ�����
'����:objDraw           ��ͼĿ��
'     lngNo             ������
'     NextX             ����������
'     NextY             �����������
'     varWidth          ���ر����ռ�Ŀ��
'     varHeight         ���ر����ռ�ĸ߶�

    Dim j As Long
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    gstrSQL = "select ���,����,����,�п�,����,�и�,�ϲ���,�ϲ���,��ɫ from ��ѯ���Ԫ�� where ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", lngNo)
    If rs.BOF = False Then
        j = objDraw.NextVsfIndex
        Call objDraw.AddPageItemGrid(j, Aligment, NextY, rs!����, rs!����, rs!�и�, rs!�п�, IIf(IsNull(rs!�ϲ���), "", rs!�ϲ���), IIf(IsNull(rs!�ϲ���), "", rs!�ϲ���), varWidth, varHeight)
        
        gstrSQL = "select ���,�к�,�к�,����,����,��ɫ,���� from ��ѯ������� where ���=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", lngNo)
        If rs.BOF = False Then
            While Not rs.EOF
                objDraw.Row(j) = rs!�к� - 1
                objDraw.Col(j) = rs!�к� - 1
                objDraw.TextMatrix(j, objDraw.Row(j), objDraw.Col(j)) = IIf(IsNull(rs!����), "", rs!����)
                objDraw.CellAlignment(j) = IIf(IsNull(rs!����), 9, rs!����)
                objDraw.CellForeColor(j) = IIf(IsNull(rs!��ɫ), 0, rs!��ɫ)
                objDraw.CellFontName(j) = Split(IIf(IsNull(rs!����), "����;9;False;False;False;False", rs!����), ";")(0)
                objDraw.CellFontSize(j) = Split(IIf(IsNull(rs!����), "����;9;False;False;False;False", rs!����), ";")(1)
                objDraw.CellFontBold(j) = IIf(Split(IIf(IsNull(rs!����), "����;9;False;False;False;False", rs!����), ";")(2) = True, True, False)
                objDraw.CellFontItalic(j) = IIf(Split(IIf(IsNull(rs!����), "����;9;False;False;False;False", rs!����), ";")(3) = True, True, False)
                objDraw.CellFontStrikethru(j) = IIf(Split(IIf(IsNull(rs!����), "����;9;False;False;False;False", rs!����), ";")(4) = True, True, False)
                objDraw.CellFontUnderline(j) = IIf(Split(IIf(IsNull(rs!����), "����;9;False;False;False;False", rs!����), ";")(5) = True, True, False)
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
    
    GetInterval = Val(GetPara("��沥�ż��", "5"))
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
    
    gstrSQL = "select nvl(count(*),0) from ��ѯҳ������ where " & IIf(UpKey = 0, "����� is null or �����=0", "�����=[1]")
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery", UpKey)
    If gRs.BOF = False Then CheckMenuLimit = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub CheckPicture()
'����:��鱾��ͼƬ,�����±���ͼƬ

    Dim strFileName As String
    Dim vFileData As New FileSystemObject
            
    '1.���ͼ��Ŀ¼�Ƿ����
    On Error Resume Next
    vFileData.CreateFolder App.Path & "\ͼ��"
    
    On Error GoTo errHand
    '2.��鱾ϵͳ�п���ʹ�õ���ͼƬ
    gstrSQL = "select ���,����,����,����,���,�߶�,�̶�,�޸����� from ��ѯͼƬԪ��"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery")
    If gRs.BOF = False Then
        While Not gRs.EOF
            strFileName = IIf(IsNull(gRs!����), "", gRs!����)
            If strFileName <> "" Then
                Call CheckFileNew(strFileName, IIf(IsNull(gRs!����), 0, gRs!����), gRs!���, gRs!�޸�����, vFileData)
            End If
            gRs.MoveNext
        Wend
    End If
    
    '3.���ҽ������Ƭ
    gstrSQL = "select B.ID, B.����  from ��ѯר���嵥 A,��Ա�� B where A.��Աid=B.id  And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) "
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "mdlNewQuery")
    If gRs.BOF = False Then
        While Not gRs.EOF
            strFileName = IIf(IsNull(gRs!����), "", gRs!����)
            If strFileName <> "" Then
                strFileName = App.Path & "\ͼ��\" & strFileName & ".pic"
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
''����:��������ͼ���ļ�
'    If CreateFileByField(objField, strFileName) Then Call SetFileDateTime(strFileName, mvarDateTime)
'End Sub

Public Sub InitInternal()
    frmMainQuery.mvarHomeInternal = 0
End Sub

Public Sub CheckFileNew(ByRef strName As String, ByVal byt���� As Byte, lngNo As Long, objFieldDate As Field, flObj As FileSystemObject)
    
    Select Case byt����
    Case 0              '����ͼƬ
        strName = App.Path & "\ͼ��\" & strName & ".pic"
    Case 1              'ͼ��
        strName = App.Path & "\ͼ��\" & strName & ".ico"
    Case 2              'FLASH
        strName = App.Path & "\ͼ��\" & strName & ".swf"
    Case 3              'Media
        strName = App.Path & "\ͼ��\" & strName & ".mid"
    End Select
    If Dir(strName) <> "" Then
        'ͼ���ļ�����,�����͵�ǰ���ݿ��е�ͼ���Ƿ�һ��(ͨ���޸�ʱ��)
        If Format(flObj.GetFile(strName).DateLastModified, "yyyy-mm-dd hh:mm") <> Format(objFieldDate, "yyyy-mm-dd hh:mm") Then
            flObj.DeleteFile (strName)
            
            strName = Sys.ReadLob(glngSys, 28, lngNo, strName, 0, 0)
'            Call CreateFile(strName, objFieldMedia, Format(objFieldDate, "yyyy-mm-dd hh:mm:ss"))
        End If
    Else
        'ͼ���ļ�������,�������µ�ͼ���ļ�
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
    '���ܣ�
    '������
    '���أ�
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

Public Function CheckIdentify(ByRef lng����ID As Long, ByRef lng��ҳid As Long) As Boolean
    '---------------------------------------------------------------------------------------------
    '
    '����:��֤���
    '����:strPswҽ��������
    '
    '---------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim strPass As String
    Dim strCard As String
    
    CheckIdentify = False
    
'    Select Case gintInsure
'    Case 80
        strCard = "0"
        CheckIdentify = frmIdentify����.ShowForm(strPass, strCard)
        
        '��������
'        CheckIdentify = True
'        strPass = "1111"
'        strCard = "0"

'    Case Else
'       '��֧�ִ˲���
'
'    End Select
    
    If CheckIdentify = False Then Exit Function
    
    CheckIdentify = False
    
    On Error GoTo errHand
    
    gintInsure = 0
    strTmp = gclsInsure.Identify2(strCard, strPass, 2, lng����ID, gintInsure)
    
    If Trim(strTmp) = "" Then Exit Function
    
    'ֻ�ܲ�סԺ����
    If lng����ID > 0 Then
        
        gstrSQL = "" & _
        " Select to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,to_char(A.��Ժ����,'yyyy-mm-dd') as ��Ժ����,A.����id,A.��ҳid " & _
        " From ������ҳ A,������Ϣ B " & _
        " Where A.����id=B.����id and B.����ID=[1] " & _
        " Union ALL " & _
        " Select '�������' as ��Ժ����,'�������' as ��Ժ����,����id,0 as ��ҳid " & _
        " From ������ü�¼ C " & _
        " Where C.����ID=[1] AND rownum<2 "
        
        gstrSQL = "Select ��Ժ����,��Ժ����,����id,��ҳid From (" & gstrSQL & " ) AA Order by AA.��Ժ���� asc,AA.��ҳid desc"
        
        gstrSQL = "" & _
        " SELECT rownum as No,D.��Ժ����,D.��Ժ����,D.����id,D.��ҳid " & _
        " FROM ( " & gstrSQL & "  ) D"
        
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, "���ò�ѯ", lng����ID)
        If gRs.RecordCount > 0 Then
            CheckIdentify = frmSelect.ShowSelect(gRs, lng����ID, lng��ҳid)
        Else
            lng����ID = 0
        End If
        
        If CheckIdentify = False Then lng����ID = 0
    End If
    
    CheckIdentify = (lng����ID > 0)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetPatientInsure(ByVal lng����ID As Long, Optional ByVal lng��ҳid As Long) As Long
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ���˵�����
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    strSQL = "Select Nvl(B.����,A.����) as ���� " & _
        " From ������Ϣ A,������ҳ B,ҽ�Ƹ��ʽ C" & _
        " Where A.����ID=" & lng����ID & " And A.����ID=B.����ID(+)" & _
        " And B.��ҳID(+)=" & lng��ҳid & " And A.ҽ�Ƹ��ʽ=C.����(+)"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlNewQuery")
    If rs.BOF = False Then
        GetPatientInsure = zlCommFun.Nvl(rs("����").Value, 0)
    End If
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '�������������ID������
    '����������ɹ����� �¼�������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ���� is null " & strWhere & " connect by prior ҳ�����=�ϼ����"
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ����=" & strID & strWhere & " connect by prior ҳ�����=�ϼ����"
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

Public Function GetLocalCodeLength(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '����������ϼ�ID������
    '����������ɹ����� ������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ���� is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ����=" & str�ϼ�ID & strWhere
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

Public Function GetParentCode(ByVal str�ϼ�ID As String, ByVal strTableName As String) As String
    '������������ȡ�ϼ�����
    '����������ϼ�ID,����
    '����������ɹ����� �ϼ�����; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select ���� from " & strTableName & " where ҳ�����=" & str�ϼ�ID
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlnewQuery")
    
    If rsTemp.RecordCount = 0 Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("����").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetMaxLocalCode(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '��������������ָ������ϼ�ID ��ȡ������������
    '����������ϼ�ID,����
    '����������ɹ����� ������; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strCode As String, strAllCode As String
    Dim intLength   As Integer
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select max(to_number(����))+1 as MaxCode from " & strTableName & " where �ϼ���� is null" & strWhere
    Else
        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ����=" & str�ϼ�ID & strWhere
    End If
    intCode = GetLocalCodeLength(str�ϼ�ID, strTableName, strWhere)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlnewQuery")
    
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
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
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    If blnNotCache Then Call zlDatabase.ClearParaCache
    GetPara = zlDatabase.GetPara(varPara, glngSys, 1536, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************

    On Error GoTo errH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, glngSys, 1536)

    Exit Function
    
errH:

End Function

Public Function GetNodeCheckSQL(Optional ByVal strNodeField As String = "վ��") As String
    '******************************************************************************************************************
    '���ܣ���ȡվ�����Ƶģӣѣ��������
    '������strNodeField=
    '���أ�
    '******************************************************************************************************************
    Dim strNodeNo As String
    
    strNodeNo = gstrNodeNo
    If strNodeNo = "" Then strNodeNo = "-"
    
    GetNodeCheckSQL = " Nvl(" & strNodeField & ", '" & strNodeNo & "') = '" & strNodeNo & "' "
    
End Function

Public Function zlGetFeeFields(Optional strTableName As String = "������ü�¼", Optional blnReadDatabase As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ�����ֵ
    '��Σ�strTableName:��:������ü�¼;סԺ���ü�¼;....
    '      blnReadDatabase-�����ݿ��ж�ȡ
    '���Σ�
    '���أ��ֶμ�
    '���ƣ����˺�
    '���ڣ�2010-03-10 10:41:42
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strFileds As String
    
    Err = 0: On Error GoTo errHand:
    If blnReadDatabase Then GoTo ReadDataBaseFields:
    Select Case strTableName
    Case "�ҺŰ���"
        'zlGetFeeFields = "ID,����,����,����ID,��ĿID,ҽ������,ҽ��ID,�޺���,��Լ��,����,��һ,�ܶ�,����,����,����,����,��������,���﷽ʽ,��ſ���,��ʼʱ��,��ֹʱ��"
        zlGetFeeFields = "ID,����,����,����ID,��ĿID,ҽ������,ҽ��ID,����,��һ,�ܶ�,����,����,����,����,��������,���﷽ʽ,��ſ���,��ʼʱ��,��ֹʱ��"
        Exit Function
    Case "������ü�¼"
        zlGetFeeFields = "" & _
        "Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, " & _
        "����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, " & _
        "�Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, " & _
        "����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, " & _
        "���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
        Exit Function
    Case "סԺ���ü�¼"
        zlGetFeeFields = "" & _
         " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, " & _
         " �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, " & _
         " ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, " & _
         " ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, " & _
         " ����id , ���ʽ��, ���մ���ID, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���"
         Exit Function
    Case "���˽��ʼ�¼"
        zlGetFeeFields = "Id, No, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������, ��ע"
        Exit Function
    Case "����Ԥ����¼"
        zlGetFeeFields = "" & _
        " Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, " & _
        " ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�"
        Exit Function
    Case "��Ա��"
        zlGetFeeFields = "" & _
        "Id, ���, ����, ����, ���֤��, ��������, �Ա�, ����, ��������, �칫�ҵ绰, �����ʼ�, ִҵ���, ִҵ��Χ, " & _
        "����ְ��, רҵ����ְ��, Ƹ�μ���ְ��, ѧ��, ��ѧרҵ, ��ѧʱ��, ��ѧ����, ������ѵ, ���п���, ���˼��, ����ʱ��, " & _
        "����ʱ��, ����ԭ��, ����, վ��"
        Exit Function
    Case "Ʊ�����ü�¼"
        zlGetFeeFields = "ID,Ʊ��,������,ǰ׺�ı�,��ʼ����,��ֹ����,ʹ�÷�ʽ,�Ǽ�ʱ��,ʹ��ʱ��,�Ǽ���,��ǰ����,ʣ������,����,�˶���,�˶�ʱ��,�˶Խ��,�˶�ģʽ,��ע"
        Exit Function
    Case "Ʊ��ʹ����ϸ"
        zlGetFeeFields = "ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,���մ���,ʹ��ʱ��,ʹ����,�˶���,�˶�ʱ��,�˶Խ��,��ע"
        Exit Function
    Case "��Ա�ɿ��¼"
        zlGetFeeFields = "ID,����ID,�տ�Ա,�տ��ID,���㷽ʽ,�����,���,ժҪ,��ֹʱ��,�Ǽ�ʱ��,�Ǽ���"
        Exit Function
    Case "��ѯͼƬԪ��"
        zlGetFeeFields = "���,����,����,����,���,�߶�,�̶�,�޸�����"
        Exit Function
    End Select
ReadDataBaseFields:
    Err = 0: On Error GoTo errHand:
    strSQL = "Select  column_name From user_Tab_Columns Where Table_Name = Upper([1]) Order By Column_ID;"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ϣ", strTableName)
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

Public Function zlGetFullFieldsTable(Optional strTableName As String = "������ü�¼", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�����ݱ��е��ֶ�.������Select Id,....
    '��Σ�bytHistory-0-��������ʷ����,1-��������ʷ����,2-����������( select * from tablename Union select * from Htablename)
    '      strWhere-����
    '      blnSubTable-�Ƿ��ӱ�
    '      strAliasName-����
    '���Σ�
    '���أ�select ID ... From tableName Union ALL
    '���ƣ����˺�
    '���ڣ�2010-03-10 11:19:11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '��
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '����ʷ
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '���߶�����
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & " " & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
End Function

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '     blnNoBeginTrans:û������ʼ
    '����:���˺�
    '����:2008/01/09
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


Public Function Exist�����(ByVal str����� As String, Optional ByVal lng����ID As Long) As Boolean
'���ܣ��ж�ָ��������Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ������Ϣ Where �����=[1] And ����ID<>[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl9NewQuery", str�����, lng����ID)
    If rsTmp.RecordCount > 0 Then Exist����� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Exist����ID(ByVal lng����ID As Long) As Boolean
'���ܣ��ж�ָ������ID�Ƿ��Ѿ����������ݿ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select ����ID From ������Ϣ Where   ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zl9NewQuery", lng����ID)
    If rsTmp.RecordCount > 0 Then Exist����ID = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'                                ���׹Һ����
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
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
    '��ȡ���׹Һ���ʾ�����Լ�����
    ' ��� strParaName -������ strMsg -��ʾ��Ϣ, strFontName-�������� dblSize-�����С dblColor -��ɫ
    '--------------------------------------------------
     Dim strRows() As String
     Dim strTmp As String, X As Long
     On Error GoTo hErr
     strTmp = GetPara(strParaName, "", True)
     If strTmp = "" Then Exit Function
    ''��ʾ��Ϣ|����|��ɫ|�����С'
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
    '���ü��׹Һŵ���ʾ�����Լ�����
     ' ��� strMsg -��ʾ��Ϣ, strFontName-�������� dblSize-�����С dblColor -��ɫ
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
    '��ȡ���׹Һ� ����ı���ɫ
    '---------------------------------------------------
        Dim strRows() As String
        Dim strTmp As String, X As Long
        On Error GoTo hErr
        strTmp = GetPara("�򵥹Һű���ɫ", "")
        If strTmp = "" Then Exit Function
   
        ''��ʾ��Ϣ|����|��ɫ|�����С'
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
    '���ü��׹Һ� ����ı���ɫ
    '---------------------------------------------------
        On Error GoTo hErr
        SetFreeRegistBGColor = SetPara("�򵥹Һű���ɫ", dblUpBgColor & "|" & dblDownBgColor)
  Exit Function
hErr:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function
Public Function GetRs�Һ�����() As ADODB.Recordset
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�ҺŰ�������
    '����:�ҺŰ������ҵļ�¼��
    '����:���˺�
    '����:2013-11-01 16:36:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo ErrHandle
    If grs�Һ����� Is Nothing Then
         strSQL = "Select �ű�ID,�������� From �ҺŰ�������"
        Set grs�Һ����� = zlDatabase.OpenSQLRecord(strSQL, "���վ���")
        Set GetRs�Һ����� = grs�Һ�����
        Exit Function
    End If
    If grs�Һ�����.State <> 1 Then
         strSQL = "Select �ű�ID,�������� From �ҺŰ�������"
        Set grs�Һ����� = zlDatabase.OpenSQLRecord(strSQL, "���վ���")
        Set GetRs�Һ����� = grs�Һ�����
        Exit Function
    End If
    Set GetRs�Һ����� = grs�Һ�����
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function ReplaseSpecial(strTmp As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����               �滻�����ַ�
    '����
    '                   ���滻���ַ�
    '����               ���滻�������ַ�����ִ�
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strSpecial As String
    Dim astrTmp() As String
    strSpecial = "'^��^��^;^��^:^��^?^��^|^,^��^.^��^"""
    astrTmp = Split(strSpecial, "^")
    For intLoop = 0 To UBound(astrTmp)
        strTmp = Replace$(strTmp, astrTmp(intLoop), "")
    Next
    ReplaseSpecial = strTmp
    
End Function







