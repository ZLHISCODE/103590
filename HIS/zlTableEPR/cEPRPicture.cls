VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarKey As Long                     '集合中的关键字
Private mvarPicID As Double                   '内容图形ID,电子病历图形(对象ID)-电子病历内容ID
Private mvarOrigPic As StdPicture           '原始图片
Private mvarOrigWidth As Long               '图片原始宽度
Private mvarOrigHeight As Long              '图片原始高度
Private mvarDesWidth As Long                '目标宽度
Private mvarDesHeight As Long               '目标高度

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let PicID(ByVal vData As Double)
    mvarPicID = vData
End Property

Public Property Get PicID() As Double
    PicID = mvarPicID
End Property

Public Property Set OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
    If mvarOrigPic.Handle <> 0 Then
        mvarOrigHeight = mvarOrigPic.Height
        mvarOrigWidth = mvarOrigPic.Width
    End If
End Property

Public Property Let OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Get OrigPic() As StdPicture
    Set OrigPic = mvarOrigPic
End Property
Public Property Let DesWidth(ByVal vData As Long)
    mvarDesWidth = vData
End Property

Public Property Get DesWidth() As Long
    DesWidth = mvarDesWidth
End Property

Public Property Let DesHeight(ByVal vData As Long)
    mvarDesHeight = vData
End Property

Public Property Get DesHeight() As Long
    DesHeight = mvarDesHeight
End Property
Public Property Let OrigWidth(ByVal vData As Long)
    mvarOrigWidth = vData
End Property

Public Property Get OrigWidth() As Long
    OrigWidth = mvarOrigWidth
End Property

Public Property Let OrigHeight(ByVal vData As Long)
    mvarOrigHeight = vData
End Property

Public Property Get OrigHeight() As Long
    OrigHeight = mvarOrigHeight
End Property
'################################################################################################################
'## 功能：  复制图片对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Sub Clone(cNode As cTabPicture)
    With cNode
        Set .OrigPic = Me.OrigPic
        .OrigWidth = Me.OrigWidth
        .OrigHeight = Me.OrigHeight
    End With
End Sub
'################################################################################################################
'## 功能：  获取图片属性字符串
'################################################################################################################
Public Property Get 对象属性() As String
    Dim strR As String
    strR = Me.OrigWidth & ";" & Me.OrigHeight & ";" & Me.DesWidth & ";" & Me.DesHeight
    对象属性 = strR
End Property

'################################################################################################################
'## 功能：  通过一个字符串来设置图片属性
'################################################################################################################
Public Property Let 对象属性(vData As String)
On Error Resume Next
    Dim T As Variant, i As Long
    T = Split(vData, ";")
    If UBound(T) > 0 Then
        Me.OrigWidth = IIf(T(0) = "", 0, T(0))
        Me.OrigHeight = IIf(T(1) = "", 0, T(1))
        Me.DesWidth = IIf(T(0) = "", 0, T(2))
        Me.DesHeight = IIf(T(0) = "", 0, T(3))
    End If
End Property


'################################################################################################################
'## 功能：  从数据库直接获取图片内容
'##
'## 参数：  eEdtType    :当前编辑方式
'################################################################################################################
Public Function GetPictureFromDB(ByVal EM As TabEM, ByVal editType As TabET, ByVal blnMoved As Boolean, ByRef tmpMark As cTabPicMarks, ByRef tmpMarkKeys As String) As StdPicture
Dim strFile As String, strZipFile As String, rs As New ADODB.Recordset
    Set GetPictureFromDB = New StdPicture
    If EM = TabEm_新增 Then
        strZipFile = zlBlobRead(2, Me.PicID)
    Else
        Select Case editType
        Case TabET_病历文件定义
            strZipFile = zlBlobRead(2, Me.PicID)
        Case TabET_全文示范编辑
            gstrSQL = "Select ID from 病历范文内容 Where ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "", Me.PicID)
            If rs.EOF Then
                strZipFile = zlBlobRead(2, Me.PicID)
            Else
                strZipFile = zlBlobRead(4, Me.PicID)
            End If
        Case TabET_单病历编辑, TabET_单病历审核
            strZipFile = zlBlobRead(6, Me.PicID, , blnMoved)  '读出的是历史表时,图形数据也是转储
        End Select
    End If
    If strZipFile = "" Then Exit Function
    If gobjFSO.FileExists(strZipFile) Then
        strFile = zlFileUnzip(strZipFile, "Jpg")
        gobjFSO.DeleteFile strZipFile, True
        If strFile = "" Then Exit Function
        
        If gobjFSO.FileExists(strFile) Then
            Set Me.OrigPic = LoadPicture(strFile)
            gobjFSO.DeleteFile strFile, True
        End If
    End If


    Dim strTableName As String, strPicMark As String, i As Integer, strMarkKey As String, lngKey As Long, int保留 As Integer, int开始版 As Integer, int终止版 As Integer
    If EM = TabEm_新增 Then
        strTableName = "病历文件结构"
    Else
        Select Case editType
        Case TabET_病历文件定义
            strTableName = "病历文件结构"
        Case TabET_全文示范编辑
            gstrSQL = "Select ID from 病历范文内容 Where ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "", Me.PicID)
            If rs.EOF Then
                strTableName = "病历文件结构"
            Else
                strTableName = "病历范文内容"
            End If
        Case TabET_单病历编辑, TabET_单病历审核
            strTableName = IIf(blnMoved, "H电子病历内容", "电子病历内容")
        End Select
    End If
    
    If InStr(strTableName, "电子病历内容") > 0 Then '提取相同单元格终止版=0,表明该单元格是目前最新版
        gstrSQL = "Select *" & vbNewLine & _
                "From " & strTableName & vbNewLine & _
                "Where 父id = [1]" & vbNewLine & _
                "Order By 对象序号, 内容行次"
    Else
        gstrSQL = "Select * From " & strTableName & " Where 父id = [1] Order By 对象序号, 内容行次"
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", Me.PicID)
    Do Until rs.EOF
        strPicMark = strPicMark & rs!内容文本
        int保留 = Nvl(rs!保留对象, 0)
        If InStr(strTableName, "电子病历内容") > 0 Then
            int开始版 = Nvl(rs!开始版, 1)
            int终止版 = Nvl(rs!终止版, 0)
        End If
        rs.MoveNext
    Loop
    If strPicMark <> "" Then
        For i = 0 To UBound(Split(strPicMark, "||"))
            lngKey = tmpMark.Add
            tmpMarkKeys = tmpMarkKeys & "|" & lngKey
            With tmpMark("K" & lngKey)
                .父ID = Me.PicID
                .选中 = Split(Split(strPicMark, "||")(i), "|")(0)
                .类型 = Split(Split(strPicMark, "||")(i), "|")(1)
                .内容 = Split(Split(strPicMark, "||")(i), "|")(2)
                .点集 = Split(Split(strPicMark, "||")(i), "|")(3)
                .X1 = Split(Split(strPicMark, "||")(i), "|")(4)
                .Y1 = Split(Split(strPicMark, "||")(i), "|")(5)
                .X2 = Split(Split(strPicMark, "||")(i), "|")(6)
                .Y2 = Split(Split(strPicMark, "||")(i), "|")(7)
                .填充色 = Split(Split(strPicMark, "||")(i), "|")(8)
                .填充方式 = Split(Split(strPicMark, "||")(i), "|")(9)
                .线条色 = Split(Split(strPicMark, "||")(i), "|")(10)
                .字体色 = Split(Split(strPicMark, "||")(i), "|")(11)
                .线型 = Split(Split(strPicMark, "||")(i), "|")(12)
                .线宽 = Split(Split(strPicMark, "||")(i), "|")(13)
                .字体 = Split(Split(strPicMark, "||")(i), "|")(14)
                .保留对象 = IIf(int保留 = 1, True, False)
                .开始版 = int开始版
                .终止版 = int终止版
            End With
        Next
    End If
    Set GetPictureFromDB = mvarOrigPic
End Function
Public Function GetSaveSQL(ByVal editType As TabET, ByVal CellTmp As cTabCell, ByVal PicMarksTmp As cTabPicMarks, arrSQL As Variant) As Boolean
'功能:生成图片单元格,图片内容,图片标记保存SQL
    On Error GoTo errHand
    If editType = TabET_单病历审核 Then GetSaveSQL = True: Exit Function '审核不能编辑图片
    With CellTmp
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        Select Case editType
        Case TabET_病历文件定义
            arrSQL(UBound(arrSQL)) = "Zl_病历文件结构_Update(" & .ID & "," & .文件ID & _
                                        ",Null," & .对象序号 & "," & .对象类型 & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.对象属性 & "',Null,'" & .内容文本 & "')"
        Case TabET_全文示范编辑
            arrSQL(UBound(arrSQL)) = "Zl_病历范文内容_Update(" & .ID & "," & .文件ID & _
                                        ",Null," & .对象序号 & "," & .对象类型 & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.对象属性 & "',Null,'" & .内容文本 & "')"
        Case TabET_单病历编辑, TabET_单病历审核
            arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Update(" & .ID & "," & .文件ID & "," & .开始版 & "," & .终止版 & _
                                        ",Null," & .对象序号 & "," & .对象类型 & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                        .Row & "," & .Col & "|" & Me.对象属性 & "',Null,'" & .内容文本 & "')"
        End Select
        
        If mvarOrigPic.Handle <> 0 Then '保存图片'审核不能编辑图片
            Dim strFile As String, strZipFile As String
            strFile = App.Path & "\Tmp.Jpg"
            If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True

            Dim GpInput As GdiplusStartupInput
            Dim m_GDIpToken         As Long         ' 用于关闭 GDI+
            Dim oDIB As New cDIB
            Dim DIBDither As cDIBDither
            Dim DIBPal As New cDIBPal
            '-- 调入 GDI+ Dll
            GpInput.GdiplusVersion = 1
            If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
                '按照BMP格式保存！会增大图片体积
                SavePicture mvarOrigPic, strFile       '保存格式为BMP格式
            Else
                '采用JPEG压缩格式保存
                Call oDIB.CreateFromStdPicture(mvarOrigPic, DIBPal, DIBDither)
                '压缩存储
                Call mGdIpEx.SaveDIB(oDIB, strFile, [ImageJPEG], 100)           '90%的JPEG图片压缩质量
            End If
            Call mGdIpEx.GdiplusShutdown(m_GDIpToken)
            
            strZipFile = zlFileZip(strFile)     '压缩文件
            gobjFSO.DeleteFile strFile, True '删除原始文件
            If strZipFile <> "" Then
                If gobjFSO.FileExists(strZipFile) Then
                    Select Case editType
                        Case TabET_病历文件定义
                            Call zlBlobSql(2, .ID, strZipFile, arrSQL)
                        Case TabET_全文示范编辑
                            Call zlBlobSql(4, .ID, strZipFile, arrSQL)
                        Case TabET_单病历编辑, TabET_单病历审核
                            Call zlBlobSql(6, .ID, strZipFile, arrSQL)
                    End Select
                    gobjFSO.DeleteFile strZipFile, True      '删除临时文件
                End If
            End If
        End If
        
        If .PicMarkKey <> "" Then '保存标记
            
            Dim i As Long, strPicMark As String, strSub As String, lngLen As Long, lngSub行次 As Long, lngID As Double
            For i = 1 To UBound(Split(.PicMarkKey, "|"))
                With PicMarksTmp("K" & Split(.PicMarkKey, "|")(i))
                    strPicMark = strPicMark & "||" & .对象属性
                End With
            Next
            
            strPicMark = Mid(strPicMark, 3)
            lngLen = Len(strPicMark)
            i = 0    '按照4000为界分段存储。
            lngSub行次 = 1
            Do While (i * 4000 + 1 <= lngLen)
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                strSub = Mid(strPicMark, i * 4000 + 1, 4000)
                Select Case editType
                Case TabET_病历文件定义
                    lngID = zlDatabase.GetNextId("病历文件结构")
                    arrSQL(UBound(arrSQL)) = "Zl_病历文件结构_Update(" & lngID & "," & CellTmp.文件ID & "," & .ID & "," & _
                                            CellTmp.对象序号 & "," & cprCTPicMark & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                            CellTmp.Row & "," & CellTmp.Col & "|'," & lngSub行次 & ",'" & strSub & "')"
                Case TabET_全文示范编辑
                    lngID = zlDatabase.GetNextId("病历范文内容")
                    arrSQL(UBound(arrSQL)) = "Zl_病历范文内容_Update(" & lngID & "," & CellTmp.文件ID & "," & .ID & "," & _
                                            CellTmp.对象序号 & "," & cprCTPicMark & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                            CellTmp.Row & "," & CellTmp.Col & "|'," & lngSub行次 & ",'" & strSub & "')"
                Case TabET_单病历编辑, TabET_单病历审核
                    lngID = GetEPRContentNextId
                    arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Update(" & lngID & "," & CellTmp.文件ID & "," & .开始版 & "," & .终止版 & "," & .ID & "," & _
                                            CellTmp.对象序号 & "," & cprCTPicMark & ",Null," & IIf(.保留对象, 1, 0) & ",'" & _
                                            CellTmp.Row & "," & CellTmp.Col & "|'," & lngSub行次 & ",'" & strSub & "')"
                End Select
                lngSub行次 = lngSub行次 + 1
                i = i + 1
            Loop
        End If
    End With
    
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub Class_Initialize()
    Set mvarOrigPic = New StdPicture
End Sub

Private Sub Class_Terminate()
    Set mvarOrigPic = Nothing
End Sub
