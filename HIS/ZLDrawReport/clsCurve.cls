VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'######################################################################################################################

Private Type RANGE
    Min As Single
    Max As Single
End Type

Private mobjTextX As RANGE
Private mobjTextY As RANGE
Private mobjDrawX As RANGE
Private mobjDrawY As RANGE
Private mstrChar As String
Private mrsDataSource As New ADODB.Recordset
Private mstrFieldX As String
Private mstrFieldX2 As String
Private mstrCharField As String
Private mstrColorField As String
Private mstrFieldY As String
Private mlngColor As Long
Private mintPostionY As Integer
Private mobjDraw As Object
Private mbytLineType As Byte                                    '1-曲线;2-只画点;3-连续线;4-连续并以数据值作为点处的字符
Private mstrFontName As String
Private msglFontNameSize As Single
Private mblnBeginArraor As Boolean
Private mblnEndArraor As Boolean

'######################################################################################################################

Public Property Get LineType() As Byte
    LineType = mbytLineType
End Property

Public Property Let LineType(ByVal vData As Byte)
    mbytLineType = vData
End Property

Public Property Let PostionY(ByVal vData As Integer)
    mintPostionY = vData
End Property

Public Property Let Char(ByVal vData As String)
    mstrChar = vData
End Property

Public Property Let Color(ByVal vData As Long)
    mlngColor = vData
End Property

Public Property Let RangeMinX(ByVal vData As Single)
    mobjTextX.Min = vData
End Property

Public Property Let RangeMaxX(ByVal vData As Single)
    mobjTextX.Max = vData
End Property

Public Property Let RangeMinY(ByVal vData As Single)
    mobjTextY.Min = vData
End Property

Public Property Let RangeMaxY(ByVal vData As Single)
    mobjTextY.Max = vData
End Property

Public Property Let DrawMinX(ByVal vData As Single)
    mobjDrawX.Min = vData
End Property

Public Property Let DrawMaxX(ByVal vData As Single)
    mobjDrawX.Max = vData
End Property

Public Property Let DrawMinY(ByVal vData As Single)
    mobjDrawY.Min = vData
End Property

Public Property Let DrawMaxY(ByVal vData As Single)
    mobjDrawY.Max = vData
End Property

Public Property Let BeginArraor(ByVal vData As Boolean)
    mblnBeginArraor = vData
End Property

Public Property Let EndArraor(ByVal vData As Boolean)
    mblnEndArraor = vData
End Property

Public Function InitData(ByVal bytLineType As Byte, Optional ByVal strFontName As String = "宋体", Optional ByVal sglFontNameSize As Single = 9) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrFontName = strFontName
    msglFontNameSize = sglFontNameSize
    
    LineType = bytLineType
    
    InitData = True
    
End Function

Public Function DataSource(ByVal vData As ADODB.Recordset, ByVal strFieldX As String, ByVal strFieldY As String, _
                            ByVal strColorField As String, Optional ByVal strCharField As String, Optional ByVal strFieldX2 As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Set mrsDataSource = vData
    
    mstrFieldX = strFieldX
    mstrFieldY = strFieldY
    mstrFieldX2 = strFieldX2
    mstrCharField = strCharField
    mstrColorField = strColorField
    
    DataSource = True
End Function

Friend Function FormatData(Optional ByVal strKind As String = "数据", Optional ByVal strPrex As String = "A") As Boolean
    '******************************************************************************************************************
    '功能：格式化曲线数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim x As Single
    Dim y As Single
    Dim X2 As Single
    Dim objLine As USERRECT
    Dim objArraor As USERRECT
    Dim objChar As USERRECT
    Dim objSvrRect As USERRECT
    Dim objSvrFont As USERFONT
    Dim sglCharWidth As Single
    Dim sglCharHeight As Single
    Dim rsDraw As ADODB.Recordset
    Dim strFile As String
    
    On Error GoTo errHand
    
    Call SaveRect(gobjRect, objSvrRect)
    Call SaveFont(gobjFont, objSvrFont)
    
    If mstrFontName <> "" Then gobjFont.Name = mstrFontName
    If msglFontNameSize > 0 Then gobjFont.Size = msglFontNameSize
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    Set rsDraw = New ADODB.Recordset
    With rsDraw
        .Fields.Append "X", adDecimal
        .Fields.Append "Y", adDecimal
        .Fields.Append "X2", adDecimal
        .Fields.Append "Color", adVarChar, 30
        .Fields.Append "Char", adVarChar, 255
        .Fields.Append "BeginPoint", adTinyInt
        .Fields.Append "EndPoint", adTinyInt
        .Open
    End With
    
    '将数据转换为坐标值
    '------------------------------------------------------------------------------------------------------------------
    If mrsDataSource.RecordCount > 0 Then
        
        mrsDataSource.Sort = mstrFieldX
        
        mrsDataSource.MoveFirst
        
        Do While Not mrsDataSource.EOF

            x = mobjDrawX.Min + (mobjDrawX.Max - mobjDrawX.Min) * Val(mrsDataSource(mstrFieldX).Value) / (mobjTextX.Max - mobjTextX.Min)
            y = mobjDrawY.Min + (mobjDrawY.Max - mobjDrawY.Min) * (mobjTextY.Max - Val(mrsDataSource(mstrFieldY).Value)) / (mobjTextY.Max - mobjTextY.Min)
            
            If x <= mobjDrawX.Max And y <= mobjDrawY.Max Then
                If LineType = 3 Then
                    X2 = mobjDrawX.Min + (mobjDrawX.Max - mobjDrawX.Min) * Val(mrsDataSource(mstrFieldX2).Value) / (mobjTextX.Max - mobjTextX.Min)
                End If
                
                rsDraw.Filter = ""
                rsDraw.Filter = "X=" & x & " And Y=" & y
                If rsDraw.RecordCount = 0 Then
                    rsDraw.AddNew
                    rsDraw("X").Value = x
                    rsDraw("Y").Value = y
                    rsDraw("X2").Value = X2
                    rsDraw("Color").Value = 0
                    rsDraw("BeginPoint").Value = Val(mrsDataSource("起点").Value)
                    rsDraw("EndPoint").Value = Val(mrsDataSource("终点").Value)
                    
                    If Val(mrsDataSource("终点").Value) <> 1 Then
                        rsDraw("Char").Value = mrsDataSource(mstrCharField).Value
                    End If
                    
                    If mstrColorField <> "" Then
                        rsDraw("Color").Value = Val(mrsDataSource(mstrColorField).Value)
                    End If
                    
                End If
            End If
            
            mrsDataSource.MoveNext
        Loop
    End If
    
    rsDraw.Filter = ""
    If rsDraw.RecordCount > 0 Then
        
        Select Case LineType
        '--------------------------------------------------------------------------------------------------------------
        Case 3                                      '连续线

            Do While Not rsDraw.EOF
                
                objLine.X0 = rsDraw("X").Value
                objLine.Y0 = rsDraw("Y").Value
                
                objLine.X1 = rsDraw("X2").Value
                objLine.Y1 = rsDraw("Y").Value
                
                gobjFont.ForeColor = Val(rsDraw("Color").Value)
                Call SaveRect(objLine, gobjRect)
                
                Call AppendPrintData(strKind, "线条", , , , , , , , True, strPrex)
                
                rsDraw.MoveNext
            Loop
        '--------------------------------------------------------------------------------------------------------------
        Case Else
                                                
                                                
            If rsDraw("Char").Value <> "" Then
                
                If InStr(rsDraw("Char").Value, "<Image>") > 0 Then
                    sglCharWidth = mobjDraw.TextWidth("A") * 2
                    sglCharHeight = mobjDraw.TextHeight("A")
                Else
                    sglCharWidth = mobjDraw.TextWidth(rsDraw("Char").Value)
                    sglCharHeight = mobjDraw.TextHeight(rsDraw("Char").Value)
                End If
            End If
                                    
            rsDraw.Sort = "X,Y"
            rsDraw.MoveFirst
            
            objLine.X0 = -1
            objLine.Y0 = -1
            objLine.X1 = -1
            objLine.Y1 = -1
                
            Do While Not rsDraw.EOF
                
                If objLine.X0 = -1 Then
                    '第一个点
                    objLine.X0 = rsDraw("X").Value
                    objLine.Y0 = rsDraw("Y").Value
                    
                    If rsDraw("Char").Value <> "" Then
                        objChar.X0 = objLine.X0 - sglCharWidth / 2
                        objChar.Y0 = objLine.Y0 - sglCharHeight / 2
                        objChar.X1 = objLine.X0 + sglCharWidth / 2
                        objChar.Y1 = objLine.Y0 + sglCharHeight / 2
                        
                        gobjFont.ForeColor = Val(rsDraw("Color").Value)
                        Call SaveRect(objChar, gobjRect)
                        
                        If InStr(rsDraw("Char").Value, "<Image>") > 0 Then
                            
                            strFile = GetFullFilePath(rsDraw("Char").Value)
                            Call AppendPrintData(strKind, "图形", 2, False, strFile, 2, , , , , strPrex)
                            
                        Else
                            Call AppendPrintData(strKind, "文本", 2, , rsDraw("Char").Value, 2, , , , True, strPrex)
                        End If
                        
                    End If
                Else
                    '第二个及以上的点
                    objLine.X1 = rsDraw("X").Value
                    objLine.Y1 = rsDraw("Y").Value
                    
                    If rsDraw("Char").Value <> "" Then
                        objChar.X0 = objLine.X1 - sglCharWidth / 2
                        objChar.Y0 = objLine.Y1 - sglCharHeight / 2
                        objChar.X1 = objLine.X1 + sglCharWidth / 2
                        objChar.Y1 = objLine.Y1 + sglCharHeight / 2
                        
                        gobjFont.ForeColor = Val(rsDraw("Color").Value)
                        Call SaveRect(objChar, gobjRect)
                        
                        If InStr(rsDraw("Char").Value, "<Image>") > 0 Then
                            strFile = GetFullFilePath(rsDraw("Char").Value)
                            Call AppendPrintData(strKind, "图形", 2, False, strFile, 2, , , , , strPrex)
                        Else
                            Call AppendPrintData(strKind, "文本", 2, , rsDraw("Char").Value, 2, , , , True, strPrex)
                        End If
                        
                    End If
                End If
                
                If objLine.X0 <> -1 And objLine.X1 <> -1 And (LineType = 1 Or LineType = 4) Then
                    
                    gobjFont.ForeColor = Val(rsDraw("Color").Value)
                    Call SaveRect(objLine, gobjRect)
                    
                    If LineType = 4 Then
                        If Val(rsDraw("BeginPoint").Value) <> 1 Then
                            Call AppendPrintData(strKind, "线条", , , , , , , , True, strPrex)
                            If mblnBeginArraor Then Call DrawArraor(objLine.X0, objLine.Y0, strKind, strPrex)
                            If mblnEndArraor Then Call DrawArraor(objLine.X1, objLine.Y1, strKind, strPrex)
                        End If
                    Else
                        Call AppendPrintData(strKind, "线条", , , , , , , , True, strPrex)
                        If mblnBeginArraor Then Call DrawArraor(objLine.X0, objLine.Y0, strKind, strPrex)
                        If mblnEndArraor Then Call DrawArraor(objLine.X1, objLine.Y1, strKind, strPrex)
                    End If
                    
                End If
                
                objLine.X0 = rsDraw("X").Value
                objLine.Y0 = rsDraw("Y").Value
                
                rsDraw.MoveNext
            Loop
        
        End Select
        
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Call SaveRect(objSvrRect, gobjRect)
    Call SaveFont(objSvrFont, gobjFont)
    
    FormatData = True
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveRect(objSvrRect, gobjRect)
End Function

Private Function DrawArraor(ByVal x As Single, ByVal y As Single, Optional ByVal strKind As String = "数据", Optional ByVal strPrex As String = "A") As Boolean
    '******************************************************************************************************************
    '功能：格式化曲线数据
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Single
    Dim objArraor As USERRECT
    Dim objSvrRect As USERRECT
    Dim objSvrFont As USERFONT
        
    On Error GoTo errHand
    
    Call SaveRect(gobjRect, objSvrRect)
    Call SaveFont(gobjFont, objSvrFont)
        
    objArraor.X1 = x
    objArraor.Y1 = y - 180
    objArraor.Y0 = y - 110
    For intLoop = x - 40 To x + 40 Step 1
        objArraor.X0 = intLoop
        Call SaveRect(objArraor, gobjRect)
        Call AppendPrintData(strKind, "线条", , , , , , , , True, strPrex)
    Next

    objArraor.X0 = x
    objArraor.Y0 = y - 180
    objArraor.X1 = objArraor.X0
    objArraor.Y1 = y
    Call SaveRect(objArraor, gobjRect)
    Call AppendPrintData(strKind, "线条", , , , , , , , True, strPrex)
                                    
    '------------------------------------------------------------------------------------------------------------------
    Call SaveRect(objSvrRect, gobjRect)
    Call SaveFont(objSvrFont, gobjFont)
    
    DrawArraor = True
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveRect(objSvrRect, gobjRect)
    
End Function

Private Sub Class_Initialize()
    mbytLineType = 1
    Set mobjDraw = frmPubResource.picDraw
End Sub

Private Sub Class_Terminate()
    Set mobjDraw = Nothing
End Sub
