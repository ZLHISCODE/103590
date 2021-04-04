VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmLabMainImage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin C1Chart2D8.Chart2D ChartThis 
      Height          =   3645
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   5145
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   9075
      _ExtentY        =   6429
      _StockProps     =   0
      ControlProperties=   "frmLabMainImage.frx":0000
   End
End
Attribute VB_Name = "frmLabMainImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlngkey As Long
Public Function zlRefresh(ByVal lngKey As Long, Optional blnSave As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示数据
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    
    mlngkey = lngKey
    
    '初始仪器列表
    If ReadData(blnSave) = False Then Exit Function
    
    zlRefresh = True
End Function


Private Function ReadData(blnSave As Boolean) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''
    '功能               读入图像数据
    '返回               是否成功
    ''''''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim rsItem As New ADODB.Recordset
    Dim strImageType As String
    Dim strImageData As String
    Dim DrawIndex As Integer
    Dim intloop As Integer
    Dim lngStart As Long
    
    On Error GoTo errH

    gstrSql = "select id , 标本ID,图像类型,length(图像点) as 图像点长度 from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngkey)
    
    For intloop = Me.ChartThis.UBound To 1 Step -1
        Me.ChartThis(Me.ChartThis.UBound).ChartGroups(1).Data.NumSeries = 0
        Me.ChartThis(Me.ChartThis.UBound).Header.Text = ""
        If intloop <> 0 Then
            Unload Me.ChartThis(Me.ChartThis.UBound)
        End If
    Next
    '按图像数量的多少进行排版
    With Me.ChartThis
        Select Case rsTmp.RecordCount
            Case 0, 1
                Me.ChartThis(0).Top = 0
                Me.ChartThis(0).Left = 0
                Me.ChartThis(0).Width = Me.ScaleWidth
                Me.ChartThis(0).Height = Me.ScaleHeight
            Case 2
                '1
                Load Me.ChartThis(Me.ChartThis.UBound + 1)
                Me.ChartThis(Me.ChartThis.UBound).ChartGroups(1).Data.NumSeries = 0
                Me.ChartThis(Me.ChartThis.UBound).Header.Text = ""
                Me.ChartThis(Me.ChartThis.UBound).Visible = True
                Me.ChartThis(0).Width = Me.ScaleWidth / 2 - 50
                Me.ChartThis(0).Height = Me.ScaleHeight / 2 - 50
                Me.ChartThis(0).Top = Me.ScaleHeight / 2 - Me.ChartThis(0).Width / 2
                Me.ChartThis(0).Left = 0
                '2
                Me.ChartThis(1).Width = Me.ChartThis(0).Width
                Me.ChartThis(1).Height = Me.ChartThis(0).Height
                Me.ChartThis(1).Top = Me.ChartThis(0).Top
                Me.ChartThis(1).Left = Me.ChartThis(0).Left + Me.ChartThis(0).Width + 25
            Case 3, 4
                Me.ChartThis(0).Width = Me.ScaleWidth / 2 - 50
                Me.ChartThis(0).Height = Me.ScaleHeight / 2 - 50
                For intloop = 1 To 3
                    Load Me.ChartThis(Me.ChartThis.UBound + 1)
                    Me.ChartThis(Me.ChartThis.UBound).Visible = True
                    Me.ChartThis(Me.ChartThis.UBound).ChartGroups(1).Data.NumSeries = 0
                    Me.ChartThis(Me.ChartThis.UBound).Width = Me.ChartThis(0).Width
                    Me.ChartThis(Me.ChartThis.UBound).Height = Me.ChartThis(0).Height
                Next
                '1
                Me.ChartThis(0).Left = 0
                Me.ChartThis(0).Top = 0
                '2
                Me.ChartThis(1).Top = Me.ChartThis(0).Top
                Me.ChartThis(1).Left = Me.ChartThis(0).Left + Me.ChartThis(0).Width + 25
                '3
                Me.ChartThis(2).Left = 0
                Me.ChartThis(2).Top = Me.ChartThis(0).Top + Me.ChartThis(0).Height + 25
                '4
                Me.ChartThis(3).Left = Me.ChartThis(1).Left
                Me.ChartThis(3).Top = Me.ChartThis(2).Top
            Case 5, 6
                Me.ChartThis(0).Width = Me.ScaleWidth / 3 - 75
                Me.ChartThis(0).Height = Me.ScaleHeight / 2 - 50
                For intloop = 1 To 5
                    Load Me.ChartThis(Me.ChartThis.UBound + 1)
                    Me.ChartThis(Me.ChartThis.UBound).Visible = True
                    Me.ChartThis(Me.ChartThis.UBound).ChartGroups(1).Data.NumSeries = 0
                    Me.ChartThis(Me.ChartThis.UBound).Header.Text = ""
                    Me.ChartThis(Me.ChartThis.UBound).Width = Me.ChartThis(0).Width
                    Me.ChartThis(Me.ChartThis.UBound).Height = Me.ChartThis(0).Height
                Next
                '1
                Me.ChartThis(0).Left = 0
                Me.ChartThis(0).Top = 0
                '2
                Me.ChartThis(1).Top = Me.ChartThis(0).Top
                Me.ChartThis(1).Left = Me.ChartThis(0).Left + Me.ChartThis(0).Width + 25
                '3
                Me.ChartThis(2).Top = Me.ChartThis(1).Top
                Me.ChartThis(2).Left = Me.ChartThis(1).Left + Me.ChartThis(1).Width + 25
                '4
                Me.ChartThis(3).Left = 0
                Me.ChartThis(3).Top = Me.ChartThis(0).Top + Me.ChartThis(0).Height + 25
                '5
                Me.ChartThis(4).Top = Me.ChartThis(3).Top
                Me.ChartThis(4).Left = Me.ChartThis(3).Left + Me.ChartThis(3).Width + 25
                '6
                Me.ChartThis(5).Top = Me.ChartThis(3).Top
                Me.ChartThis(5).Left = Me.ChartThis(4).Left + Me.ChartThis(4).Width + 25
            Case 7, 8, 9
                Me.ChartThis(0).Width = Me.ScaleWidth / 3 - 50
                Me.ChartThis(0).Height = Me.ScaleHeight / 3 - 50
                For intloop = 1 To 8
                    Load Me.ChartThis(Me.ChartThis.UBound + 1)
                    Me.ChartThis(Me.ChartThis.UBound).Visible = True
                    Me.ChartThis(Me.ChartThis.UBound).ChartGroups(1).Data.NumSeries = 0
                    Me.ChartThis(Me.ChartThis.UBound).Header.Text = ""
                    Me.ChartThis(Me.ChartThis.UBound).Width = Me.ChartThis(0).Width
                    Me.ChartThis(Me.ChartThis.UBound).Height = Me.ChartThis(0).Height
                Next
                '1
                Me.ChartThis(0).Left = 0
                Me.ChartThis(0).Top = 0
                '2
                Me.ChartThis(1).Top = Me.ChartThis(0).Top
                Me.ChartThis(1).Left = Me.ChartThis(0).Left + Me.ChartThis(0).Width + 25
                '3
                Me.ChartThis(2).Top = Me.ChartThis(0).Top
                Me.ChartThis(2).Left = Me.ChartThis(1).Left + Me.ChartThis(1).Width + 25
                '4
                Me.ChartThis(3).Top = Me.ChartThis(0).Top + Me.ChartThis(0).Height + 25
                Me.ChartThis(3).Left = Me.ChartThis(0).Left
                '5
                Me.ChartThis(4).Top = Me.ChartThis(3).Top
                Me.ChartThis(4).Left = Me.ChartThis(1).Left
                '6
                Me.ChartThis(5).Top = Me.ChartThis(3).Top
                Me.ChartThis(5).Left = Me.ChartThis(2).Left
                '7
                Me.ChartThis(6).Top = Me.ChartThis(3).Top + Me.ChartThis(3).Height + 25
                Me.ChartThis(6).Left = Me.ChartThis(0).Left
                '8
                Me.ChartThis(7).Top = Me.ChartThis(6).Top
                Me.ChartThis(7).Left = Me.ChartThis(1).Left
                '9
                Me.ChartThis(8).Top = Me.ChartThis(6).Top
                Me.ChartThis(8).Left = Me.ChartThis(2).Left
        End Select
    End With
    
    Do Until rsTmp.EOF
        If Nvl(rsTmp("图像点长度"), 0) <> 0 And DrawIndex <= 9 Then
            If Nvl(rsTmp("图像点长度"), 0) > 4000 Then
                '长度超过4000需要单据处理
                strImageData = ""
                For intloop = 1 To rsTmp("图像点长度") / 2000 + 1
                    gstrSql = "select id , 标本ID,图像类型,to_char(substr(图像点," & intloop * 2000 - 1999 & ",2000)) as 图像点 from 检验图像结果 where id = [1] "
                    Set rsItem = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsTmp("ID")))
                    strImageType = Nvl(rsItem("图像类型"))
                    strImageData = strImageData & Nvl(rsItem("图像点"))
                Next
                DrawImg strImageType, strImageData, DrawIndex
            Else
                gstrSql = "select id , 标本ID,图像类型,to_char(substr(图像点,1,4000)) as 图像点 from 检验图像结果 where id = [1] "
                Set rsItem = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsTmp("ID")))
                strImageType = Nvl(rsItem("图像类型"))
                strImageData = Nvl(rsItem("图像点"))
                
                DrawImg strImageType, strImageData, DrawIndex
            End If
            If blnSave = True Then
                Me.ChartThis(DrawIndex).Save App.Path & "\" & rsTmp("ID") & ".cht"
            End If
        End If
        
        DrawIndex = DrawIndex + 1
        rsTmp.MoveNext
    Loop
    
    ReadData = True
    
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Sub DrawImg(strType As String, strData As String, intIndex As Integer)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能:              画图现支持(RBC,PLT,BAS,WBC)
    '参数:              strType  图像名称
    '                   strData  图像数据
    '                   IntIndex 当前画第几个Chart控件
    '其他               数据的第一位 0=直方图 1=散点图
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim aImage() As String
    Dim lngLoop As Long
    Dim lngLoopX As Long
    
    On Error GoTo errH
    
    If strType = "RBC" Then
        aImage = Split(strData, ";")
        If aImage(0) = 0 Then
            '直方图
            With Me.ChartThis(intIndex)
                .IsBatched = True
                .Reset
                .ChartGroups(1).Data.NumSeries = 0
                .Header.Adjust = oc2dAdjustCenter
                .Header.Text = strType
                .Header.Font.Bold = True
                .Header.Font.Size = 12
                .ChartGroups(1).Styles(1).Line.COLOR = vbBlack
                .ChartGroups(1).Styles(1).Line.Width = 1
                .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
                .ChartArea.Axes("Y").Min = 32
                .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
                .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
                .ChartArea.Axes("X").Max = 176
                .ChartArea.Axes("X").ValueLabels.RemoveAll
                .ChartArea.Axes("X").ValueLabels.Add 11, "50"
                .ChartArea.Axes("X").ValueLabels.Add 35, "100"
                .ChartArea.Axes("X").ValueLabels.Add 59, " "
                .ChartArea.Axes("X").ValueLabels.Add 83, "200"
                .ChartArea.Axes("X").ValueLabels.Add 104, " "
                .ChartArea.Axes("X").ValueLabels.Add 128, "300"
                .ChartArea.Axes("X").ValueLabels.Add 152, " "
                
                .ChartGroups(1).Data.NumSeries = 1
                .ChartGroups(1).Data.NumPoints(1) = UBound(aImage)
                For i = 1 To UBound(aImage)
                    .ChartGroups(1).Data.Y(1, i) = aImage(i)
                Next
                .IsBatched = False
            End With
        End If
    End If
    
    If strType = "PLT" Then
        aImage = Split(strData, ";")
        If aImage(0) = 0 Then
            '直方图
            With Me.ChartThis(intIndex)
                .IsBatched = True
                .Reset
                .ChartGroups(1).Data.NumSeries = 0
                .Header.Adjust = oc2dAdjustCenter
                .Header.Text = strType
                .Header.Font.Bold = True
                .Header.Font.Size = 12
                .ChartGroups(1).Styles(1).Line.COLOR = vbBlack
                .ChartGroups(1).Styles(1).Line.Width = 1
                .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
                .ChartArea.Axes("Y").Min = 32
                .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
                .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
                .ChartArea.Axes("X").Max = 170
                .ChartArea.Axes("X").ValueLabels.RemoveAll
                .ChartArea.Axes("X").ValueLabels.Add 7, "2"
                .ChartArea.Axes("X").ValueLabels.Add 27, " "
                .ChartArea.Axes("X").ValueLabels.Add 54, "20"
                .ChartArea.Axes("X").ValueLabels.Add 81, " "
                .ChartArea.Axes("X").ValueLabels.Add 108, "40"
                .ChartArea.Axes("X").ValueLabels.Add 135, " "
                .ChartArea.Axes("X").ValueLabels.Add 162, "60"
                .ChartGroups(1).Data.NumSeries = 1
                .ChartGroups(1).Data.NumPoints(1) = UBound(aImage)
                For i = 1 To UBound(aImage)
                    .ChartGroups(1).Data.Y(1, i) = aImage(i)
                Next
                .IsBatched = False
            End With
        End If
    End If
    
    If strType = "BAS" Then
        aImage = Split(strData, ";")
        If aImage(0) = 0 Then
            With Me.ChartThis(intIndex)
                .IsBatched = True
                .Reset
                .ChartGroups(1).Data.NumSeries = 0
                .Header.Adjust = oc2dAdjustCenter
                .Header.Text = strType
                .Header.Font.Bold = True
                .Header.Font.Size = 12
                .ChartGroups(1).Styles(1).Line.COLOR = vbBlack
                .ChartGroups(1).Styles(1).Line.Width = 1
                .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
                .ChartArea.Axes("Y").Min = 32
                .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
                .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
                .ChartArea.Axes("X").Max = 130
                .ChartArea.Axes("X").ValueLabels.RemoveAll
                .ChartArea.Axes("X").ValueLabels.Add 12, ""
                .ChartArea.Axes("X").ValueLabels.Add 25, "100"
                .ChartArea.Axes("X").ValueLabels.Add 38, " "
                .ChartArea.Axes("X").ValueLabels.Add 51, "200"
                .ChartArea.Axes("X").ValueLabels.Add 64, ""
                .ChartArea.Axes("X").ValueLabels.Add 77, "300"
                .ChartArea.Axes("X").ValueLabels.Add 90, ""
                .ChartArea.Axes("X").ValueLabels.Add 103, "400"
                .ChartArea.Axes("X").ValueLabels.Add 116, ""
                .ChartGroups(1).Data.NumSeries = 1
                .ChartGroups(1).Data.NumPoints(1) = UBound(aImage)
                For i = 1 To UBound(aImage)
                    .ChartGroups(1).Data.Y(1, i) = aImage(i)
                Next
                .IsBatched = False
            End With
        End If
    End If
    
    If strType = "WBC" Then
        aImage = Split(strData, ";")
        If aImage(0) = 1 Then
            With Me.ChartThis(intIndex)
                .IsBatched = True
                .Reset
                .ChartGroups(1).Data.NumSeries = 0
                .Header.Adjust = oc2dAdjustCenter
                .Header.Text = strType
                .Header.Font.Bold = True
                .Header.Font.Size = 12
                .ChartArea.PlotArea.IsBoxed = True
                .ChartGroups(1).Data.NumSeries = UBound(aImage) - 1
                
                .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
                .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
                For lngLoop = UBound(aImage) - 1 To 1 Step -1
                    .ChartGroups(1).ChartType = oc2dTypePlot
                    .ChartGroups(1).Styles(lngLoop).Line.Pattern = oc2dLineNone
                    .ChartGroups(1).Styles(lngLoop).Line.COLOR = vbBlack
                    .ChartGroups(1).Styles(lngLoop).Symbol.Shape = oc2dShapeBox
                    .ChartGroups(1).Styles(lngLoop).Symbol.Size = 2
                    .ChartGroups(1).Styles(lngLoop).Symbol.COLOR = vbBlack
                    .ChartGroups(1).Data.NumPoints(lngLoop) = Len(aImage(lngLoop)) + 1
                    For lngLoopX = 1 To Len(aImage(lngLoop)) + 1
                        .ChartGroups(1).Data.Y(lngLoop, lngLoopX) = IIf(Mid(aImage(lngLoop), lngLoopX, 1) = 0, .ChartGroups(1).Data.HoleValue, 128 - lngLoop + 1)
'                        Debug.Print "Y:" & lngLoop & "  x:" & lngLoopX
                    Next
                Next
                .IsBatched = False
            End With
        End If
    End If
    
    Exit Sub
errH:
'    Resume
End Sub
