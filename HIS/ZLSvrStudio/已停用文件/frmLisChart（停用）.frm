VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmLisChart 
   BorderStyle     =   0  'None
   Caption         =   "Chart"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin C1Chart2D8.Chart2D ChartThis 
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2760
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   4868
      _ExtentY        =   4868
      _StockProps     =   0
      ControlProperties=   "frmLisChart.frx":0000
   End
End
Attribute VB_Name = "frmLisChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcurYunti As Currency 'X�����һ�����൱�ڶ��ٸ�Y����ĵ�
Private Xpixe As Currency, Ypixe As Currency

Private Sub Form_Activate()
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.Hide
End Sub

Public Function DrawImg(ByVal strType As String, ByVal strData As String, ByVal strFileName As String, Optional intSaveType As Integer = 1) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '����:          ��ͼ��֧��(RBC,PLT,BAS,WBC)
        '����:          strType  ͼ������
        '                   strData  ͼ������
        '                   strFileName ���ɵ�ͼƬ�ļ���
        '                   intSaveType 0-cht��ʽ 1-jpg��ʽ 2-png��ʽ
        '
        '����               ���ݵĵ�һλ 0=ֱ��ͼ 1=ɢ��ͼ 2=Ѫ����ճ����������ͼ 3=Ѫ������ͼ
        '                   100=ͼƬ
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim aImage() As String
        Dim aFormat() As String          '��ʾ��ʽ
        Dim aLines() As String
        Dim lngLoop As Long
        Dim lngLoopX As Long
        Dim i As Integer, j As Integer
        Dim dblX As Double, dblY As Double
        Dim aLables() As String, aLable() As String, strLable As String
        Dim strFile As String, strImgData As String
        Dim killFile As String
        Dim FrmObj As frmLisChartPic
        
        On Error GoTo errH
    
     aImage = Split(strData, ";")
 
     If aImage(0) = 0 Then
         'ֱ��ͼ clsLISDev_ABX_P120
         With Me.ChartThis
             .IsBatched = True
             .Reset
             .ChartGroups(1).Data.NumSeries = 0
             .Header.Adjust = oc2dAdjustCenter
             .Header.Text = strType
             .Header.Font.Bold = True
             .Header.Font.Size = 12
             .ChartGroups(1).Styles(1).Line.Color = vbBlack
             .ChartGroups(1).Styles(1).Line.Width = 1
             .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
         
             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
         
             .ChartArea.Axes("X").ValueLabels.RemoveAll
     
             aFormat = Split(aImage(1), ",")
             If UBound(aFormat) > 1 Then
                 .ChartArea.Axes("Y").Min = aFormat(0)
                 .ChartArea.Axes("X").Max = aFormat(1)
                 '----2008-03-28 ��ʱ����ʾͼ��
                 .ChartArea.Axes("Y").Origin = 0
                 .ChartArea.Axes("Y").Min = 0
             
                 .ChartArea.Axes("Y").Max.IsDefault = True
                 .ChartArea.Axes("X").Max.IsDefault = True
                 .ChartArea.Axes("Y").Min.IsDefault = True
                 .ChartArea.Axes("X").Min.IsDefault = True
             
                 For i = 2 To UBound(aFormat)
                     .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
                 Next
             End If
             .ChartGroups(1).Data.NumSeries = 1
             .ChartGroups(1).Data.NumPoints(1) = UBound(aImage)
             For i = 2 To UBound(aImage)
                 .ChartGroups(1).Data.y(1, i) = Val(aImage(i))
             Next
             .IsBatched = False
             If intSaveType = 1 Then
                 DrawImg = .SaveImageAsJpeg(strFileName, 100, False, False, False)
             ElseIf intSaveType = 2 Then
                 DrawImg = .SaveImageAsPng(strFileName, False)
             Else
                 DrawImg = .Save(strFileName)
             End If
         End With
     
     End If
 
     '-- ɢ��ͼ clsLISDev_ABX_P120
     If aImage(0) = 1 Then
         With Me.ChartThis
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
                 .ChartGroups(1).Styles(lngLoop).Line.Color = vbBlack
                 .ChartGroups(1).Styles(lngLoop).Symbol.Shape = oc2dShapeBox
                 .ChartGroups(1).Styles(lngLoop).Symbol.Size = 1
                 .ChartGroups(1).Styles(lngLoop).Symbol.Color = vbBlack
                 .ChartGroups(1).Data.NumPoints(lngLoop) = Len(aImage(lngLoop)) + 1
                 For lngLoopX = 1 To Len(aImage(lngLoop)) + 1
                     .ChartGroups(1).Data.y(lngLoop, lngLoopX) = IIf(Mid(aImage(lngLoop), lngLoopX, 1) = 0, .ChartGroups(1).Data.HoleValue, 128 - lngLoop + 1)
                 Next
             Next
             .IsBatched = False
             If intSaveType = 1 Then
                 DrawImg = .SaveImageAsJpeg(strFileName, 100, False, False, False)
             ElseIf intSaveType = 2 Then
                 DrawImg = .SaveImageAsPng(strFileName, False)
             Else
                 DrawImg = .Save(strFileName)
             End If
         End With
     End If

     '--- Ѫ����ͼ  clsLISDev_File_LBYN6C
     If aImage(0) = 2 Then
         DrawImg = ChartDrawѪ����(strType, aImage(1), aImage(2), aImage(3), strFileName, intSaveType)
     End If
 
     '--Ѫ������ͼ clsLISDev_File_LBYN6C
     If aImage(0) = 3 Then
         DrawImg = ChartDrawѪ��(strType, aImage(1), aImage(2), aImage(3), strFileName, intSaveType)
     End If
 
     '--- �������ص����ߵ�PLTͼ clsLISDev_HMX
     If aImage(0) = 4 Then
         DrawImg = ChartDrawPLT(strType, aImage(1), aImage(2), strFileName, intSaveType)
     End If
 
     '--- �ڱ��ص�PIC�ؼ��ϻ��� ֱ��ͼ Ȼ����ʾ
     If aImage(0) = 5 Then
         DrawImg = PicShowChart(strType & ";" & strData, strFileName, intSaveType)
     
     End If
 
     '--- ֱ��ͼ����ͬһ��ͼ�ϻ��ƶ�������
     If aImage(0) = 6 Then
         'ֱ��ͼ WBC clsLISDev_MEDONIC_M20M
         With Me.ChartThis
             .IsBatched = True
             .Reset
             .ChartGroups(1).Data.NumSeries = 0
             .Header.Adjust = oc2dAdjustCenter
             .Header.Text = strType
             .Header.Font.Bold = True
             .Header.Font.Size = 12
         
         
             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
         
             .ChartArea.Axes("X").ValueLabels.RemoveAll
             aFormat = Split(aImage(1), ",")
             If UBound(aFormat) > 1 Then
                 .ChartArea.Axes("Y").Min = aFormat(0)
                 .ChartArea.Axes("X").Max = aFormat(1)
                 '----2008-03-28 ��ʱ����ʾͼ��
                 .ChartArea.Axes("Y").Origin = 0
                 .ChartArea.Axes("Y").Min = 0
             
                 For i = 2 To UBound(aFormat)
                     .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
                 Next
             End If
         
             aLines = Split(strData, "~")
             .ChartGroups(1).Data.NumSeries = UBound(aLines)
         
             For i = LBound(aLines) + 1 To UBound(aLines)
                 aImage = Split(aLines(i), ";")
                 .ChartGroups(1).Styles(i).Line.Color = vbBlack
                 .ChartGroups(1).Styles(i).Line.Width = 1
                 .ChartGroups(1).Styles(i).Symbol.Shape = oc2dShapeNone
                 .ChartGroups(1).Data.NumPoints(i) = UBound(aImage) + 1
                 For j = LBound(aImage) To UBound(aImage)
                     .ChartGroups(1).Data.y(i, j + 1) = Val(aImage(j))
                 Next
             Next
             .IsBatched = False
             If intSaveType = 1 Then
                 DrawImg = .SaveImageAsJpeg(strFileName, 100, False, False, False)
             ElseIf intSaveType = 2 Then
                 DrawImg = .SaveImageAsPng(strFileName, False)
             Else
                 DrawImg = .Save(strFileName)
             End If
         End With
     End If
     killFile = ""
     
    If aImage(0) >= 100 And aImage(0) <= 227 Then
        strFile = aImage(1)
             
        If UCase$(strFile) Like "*.ZIP" Then
            killFile = strFile
            If aImage(0) >= 200 And aImage(0) <= 207 Then
                strFile = zlFileUnzip(strFile)
            ElseIf aImage(0) >= 210 And aImage(0) <= 217 Then
                strFile = zlFileUnzip(strFile)
            ElseIf aImage(0) >= 220 And aImage(0) <= 227 Then
                strFile = zlFileUnzip(strFile)
            End If
            
            If killFile <> "" Then Kill killFile: killFile = "" '��ѹ���ԭʼZIPҪɾ��
        End If
         
            'Ҫ�ȱ����bmp�ſ�����Ϊchart2D�ı���...
            Set FrmObj = New frmLisChartPic
            If UCase(strFile) Like "*.JPG" Then
                FrmObj.picTmp.Picture = LoadPicture(strFile)
                strFile = Replace(UCase(strFile), ".JPG", ".BMP")
                SavePicture FrmObj.picTmp, strFile
                killFile = Replace(UCase(strFile), ".BMP", ".JPG")
            ElseIf UCase(strFile) Like "*.GIF" Then
                If CheckGif(strFile) Then
                    FrmObj.picTmp.Picture = LoadPicture(strFile)
                    strFile = Replace(UCase(strFile), ".GIF", ".BMP")
                    SavePicture FrmObj.picTmp, strFile
                Else
                    Exit Function
                End If
            End If
            Unload FrmObj
            
         '--- ֱ����ʾͼƬ clsLISDev_UF100_DY
            If CInt(Val(Right(aImage(0), 1))) = 0 Then
                DrawImg = ChartShowPic(strType, strFile, strFileName, , intSaveType)  '��Ĭ�ϵ�layOut
            Else
                DrawImg = ChartShowPic(strType, strFile, strFileName, CInt(Val(Right(aImage(0), 1))), intSaveType)   '��ָ����layout
            End If
            If strFile <> "" Then Kill strFile
            If killFile <> "" Then Kill killFile 'ɾ����ʱ�ļ�
    End If
    
    Exit Function
errH:
    MsgBox err.Description, vbExclamation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function ChartDrawѪ����(ByVal strType As String, ByVal strXYin As String, ByVal strLineIn As String, ByVal strLableIn As String, ByVal strFileName As String, Optional intSaveType As Integer) As Boolean
    Dim aFormat() As String
    Dim aLables() As String, aLable() As String, strLable As String
    Dim i As Integer
    Dim aCurves() As String '����������
    Dim aCurve() As String
    Dim intPoint As Integer
    Dim aPoint() As String '���������
    Dim lngLoop As Long
    Dim dblX As Double, dblY As Double
    Dim aAxes() As String
    
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        .ChartArea.Axes("Y").ValueLabels.RemoveAll
        
        '--- ��ʼ
        aFormat = Split(strXYin, "|") '������
        If UBound(aFormat) > 1 Then
            
            aAxes = Split(aFormat(0), ",") '���곤��
            .ChartArea.Axes("Y").Max = aAxes(0)
            .ChartArea.Axes("X").Max = aAxes(1)
            .ChartArea.Axes("Y").Origin = 0
            
            aAxes = Split(aFormat(1), ",") 'X�̶�
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("X").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
            aAxes = Split(aFormat(2), ",") 'Y�̶�
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("Y").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
        End If
        
        aFormat = Split(strLineIn, "~") '���߼����
        If UBound(aFormat) > 0 Then

            
            aCurves = Split(aFormat(0), "|") '��������
            
            For i = LBound(aCurves) To UBound(aCurves)
                aCurve = Split(aCurves(i), ",")
                
                .ChartGroups(1).Styles(i + 1).Line.Color = vbBlack
                .ChartGroups(1).Styles(i + 1).Line.Width = 1
                .ChartGroups(1).Styles(i + 1).Symbol.Shape = oc2dShapeNone
                
                .ChartGroups(1).Data.NumSeries = i + 1
                .ChartGroups(1).Data.NumPoints(i + 1) = 225
            
                    '����
                If UBound(aCurve) > 2 Then
                    For lngLoop = 1 To 200
                        dblY = GetNd(Val(aCurve(0)), Val(aCurve(1)), Val(aCurve(2)), Val(aCurve(3)), lngLoop)
                        If lngLoop > 2 Then
                            .ChartGroups(1).Data.y(i + 1, lngLoop + 1) = dblY
                        Else
                            .ChartGroups(1).Data.y(i + 1, lngLoop + 1) = .ChartGroups(1).Data.HoleValue
                        End If
                    Next
                End If
            Next
        
            aPoint = Split(aFormat(1), ",") '�������

            intPoint = UBound(aCurves) + 2
            .ChartGroups(1).Styles(intPoint).Line.Pattern = oc2dLineNone
            .ChartGroups(1).Styles(intPoint).Line.Color = vbBlack
            .ChartGroups(1).Styles(intPoint).Line.Width = 1
            .ChartGroups(1).Styles(intPoint).Symbol.Color = vbBlack
            .ChartGroups(1).Styles(intPoint).Symbol.Shape = oc2dShapeSquare
            .ChartGroups(1).Data.NumSeries = intPoint
            .ChartGroups(1).Data.NumPoints(intPoint) = 225

            For i = 1 To 200
                
                For lngLoop = LBound(aPoint) To UBound(aPoint)
                    '-- ���
                    dblX = Val(Mid(aPoint(lngLoop), 1, InStr(aPoint(lngLoop), "-") - 1))
                    dblY = Val(Mid(aPoint(lngLoop), InStr(aPoint(lngLoop), "-") + 1))
                    If dblX = i + 1 Then
                        .ChartGroups(1).Data.y(intPoint, i + 1) = dblY
                        Exit For
                    Else
                        .ChartGroups(1).Data.y(intPoint, i + 1) = .ChartGroups(1).Data.HoleValue
                    End If
                Next
                
            Next
        End If
        
        
        aLables = Split(strLableIn, "~")  'X���ǩ��Y���ǩ
        
        aLable = Split(aLables(0), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(1).Text = strLable
            .ChartLabels.Item(1).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(1).AttachDataCoord.y = Val(aLable(2))
        End If
        
        aLable = Split(aLables(1), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(2).Text = strLable
            .ChartLabels.Item(2).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(2).AttachDataCoord.y = Val(aLable(2))
        End If
        
        '---- ����

        .IsBatched = False
        If intSaveType = 1 Then
            ChartDrawѪ���� = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDrawѪ���� = .SaveImageAsPng(strFileName, False)
        Else
            ChartDrawѪ���� = .Save(strFileName)
        End If
    End With
End Function
Private Function ChartDrawѪ��(ByVal strType As String, ByVal strXYin As String, ByVal strLineIn As String, ByVal strLableIn As String, ByVal strFileName As String, Optional intSaveType As Integer) As Boolean
    '��Ѫ��ͼ
    Dim aFormat() As String
    Dim aLables() As String, aLable() As String, strLable As String
    Dim i As Integer
    Dim aCurves() As String '����������
    Dim aCurve() As String
    Dim intPoint As Integer
    Dim aPoint() As String '���������
    Dim lngLoop As Long
    Dim dblX As Double, dblY As Double
    Dim aAxes() As String
    
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        .ChartArea.Axes("Y").ValueLabels.RemoveAll
        
        .ChartArea.Axes("Y").Origin = 0
        .ChartArea.Axes("Y").Min = 0
        .ChartArea.Axes("X").Origin = 0
        .ChartArea.Axes("X").Min = 0
        '--- ��ʼ
        aFormat = Split(strXYin, "|") '������
        If UBound(aFormat) > 1 Then
                            
            aAxes = Split(aFormat(0), ",") '���곤��
            .ChartArea.Axes("Y").Max = aAxes(0)
            .ChartArea.Axes("X").Max = aAxes(1)
            .ChartArea.Axes("Y").Origin = 0
            
            aAxes = Split(aFormat(1), ",") 'X�̶�
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("X").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
            aAxes = Split(aFormat(2), ",") 'Y�̶�
            For i = LBound(aAxes) To UBound(aAxes)
                .ChartArea.Axes("Y").ValueLabels.Add Mid(aAxes(i), 1, InStr(aAxes(i), "-") - 1), Mid(aAxes(i), InStr(aAxes(i), "-") + 1)
            Next
            
        End If
        
        aFormat = Split(strLineIn, ",") '����
        If UBound(aFormat) > 0 Then

            .ChartGroups(1).Styles(1).Line.Color = vbBlack
            .ChartGroups(1).Styles(1).Line.Width = 1
            .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeSquare '���ķ���
            .ChartGroups(1).Styles(1).Symbol.Color = vbRed
            
            .ChartGroups(1).Styles(2).Symbol.Color = vbRed
            .ChartGroups(1).Styles(2).Symbol.Shape = oc2dShapeSquare
            .ChartGroups(1).Styles(2).Line.Color = vbBlack
            
            .ChartGroups(1).Styles(3).Symbol.Shape = oc2dShapeNone
            .ChartGroups(1).Styles(3).Line.Color = vbBlack
            
            .ChartGroups(1).Data.NumSeries = 3
            .ChartGroups(1).Data.NumPoints(1) = UBound(aFormat)
            
'            For i = 1 To UBound(aFormat) * 4
'                    '����
'                If (i Mod 2) = 0 Then
'                    .ChartGroups(1).Data.Y(1, i) = aFormat(i / 4)
'                End If
'                If aFormat(i / 4) - 0.5 >= 0 Then
'                .ChartGroups(1).Data.Y(2, i) = aFormat(i / 4) - 0.5
'                End If
'            Next
            For i = 1 To UBound(aFormat)
                If (i Mod 2) = 0 Then
                    .ChartGroups(1).Data.y(1, i) = aFormat(i)
                Else
                    .ChartGroups(1).Data.y(2, i) = aFormat(i)
                End If
                    '����
                .ChartGroups(1).Data.y(3, i) = aFormat(i) - 0.3
                
            Next
        
        End If
        
        aLables = Split(strLableIn, "~") 'X���ǩ��Y���ǩ
        
        aLable = Split(aLables(0), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(1).Text = strLable
            .ChartLabels.Item(1).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(1).AttachDataCoord.y = Val(aLable(2))
        End If
        
        aLable = Split(aLables(1), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(2).Text = strLable
            .ChartLabels.Item(2).AttachDataCoord.x = Val(aLable(1))
            .ChartLabels.Item(2).AttachDataCoord.y = Val(aLable(2))
        End If
        
        '---- ����

        .IsBatched = False

        If intSaveType = 1 Then
            ChartDrawѪ�� = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDrawѪ�� = .SaveImageAsPng(strFileName, False)
        Else
            ChartDrawѪ�� = .Save(strFileName)
        End If
    End With
End Function
Private Function ChartShowPic(ByVal strType As String, ByVal strImgName As String, ByVal strFileName As String, Optional ByVal intLayOut As Integer = oc2dImageFitted, Optional intSaveType As Integer = 0) As Boolean
    'Chart ��ʾͼƬ
    'strImgName-Դ�ļ� strFilename-�����ļ���
    Dim strImgFile As String
    
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        .ChartArea.Axes("Y").ValueLabels.RemoveAll
        If intSaveType = 0 Then
            strImgFile = Replace(strFileName, ".cht", ".bmp")
            gobjFile.CopyFile strImgName, strImgFile, True
        Else
            strImgFile = strFileName
            gobjFile.CopyFile strImgName, strImgFile, True
        End If
        .Interior.Image.Filename = strImgName
        .Interior.Image.Layout = intLayOut 'oc2dImageFitted
'        .ChartArea.Interior.Image.Filename = strImgName
'        .ChartArea.Interior.Image.Layout = intLayOut
        .IsBatched = False
        If intSaveType = 1 Then
            ChartShowPic = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartShowPic = .SaveImageAsPng(strFileName, False)
        Else
            ChartShowPic = .Save(strFileName)
        End If
    End With
    
End Function
Private Function GetNd(ByVal ND0 As Double, QB0 As Double, ND1 As Double, QB1 As Double, ByVal Qb As Double) As Double
    'Ѫ�����׼����������㺯��
    Dim k0 As Double, K1 As Double
    Dim sn As Double

    k0 = (Sqr(ND0) - Sqr(ND1)) / (1 / (Sqr(QB0)) - 1 / (Sqr(QB1)))
    K1 = Sqr(ND0) - k0 * (1 / (Sqr(QB0)))
    sn = k0 * (1 / (Sqr(Qb))) + K1
    GetNd = sn * sn

End Function

Private Function ChartDrawPLT(ByVal strType As String, ByVal str_���� As String, ByVal str_Lines As String, ByVal strFileName As String, Optional intSaveType As Integer = 0) As Boolean
    
    Dim aFormat() As String
    Dim i As Integer, y As Integer
    Dim varLine() As String
    Dim aLine() As String
    Dim Y�� As String, str����� As String
    Dim aLables() As String, aLable() As String, strLable As String
    Y�� = "": str����� = ""
     '2�����ص����Ƶ�ֱ��ͼ
    With Me.ChartThis
        .IsBatched = True
        .Reset
        .ChartGroups(1).Data.NumSeries = 0
        .Header.Adjust = oc2dAdjustCenter
        .Header.Text = strType
        .Header.Font.Bold = True
        .Header.Font.Size = 12
        .ChartGroups(1).Styles(1).Line.Color = vbBlack
        .ChartGroups(1).Styles(1).Line.Width = 1
        .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
        
        .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
        .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
        
        .ChartArea.Axes("X").ValueLabels.RemoveAll
        
        If InStr(str_����, "|") > 0 Then
            Y�� = Split(str_����, "|")(1)
            str_���� = Split(str_����, "|")(0)
        End If
        aFormat = Split(str_����, ",")
        If UBound(aFormat) > 1 Then
            .ChartArea.Axes("Y").Max = aFormat(0)
            .ChartArea.Axes("X").Max = aFormat(1)
            
            .ChartArea.Axes("Y").Min = 0
            .ChartArea.Axes("X").Min = 0
           
           .ChartArea.Axes("Y").Origin = 0
           .ChartArea.Axes("X").Origin = 0
            .ChartArea.Axes("Y").Max.IsDefault = True
            .ChartArea.Axes("X").Max.IsDefault = True
            .ChartArea.Axes("Y").Min.IsDefault = True
            .ChartArea.Axes("X").Min.IsDefault = True
            For i = 2 To UBound(aFormat)
                .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
            Next
        End If
        If Y�� <> "" Then
            aFormat = Split(Y��, ",")
            For i = 0 To UBound(aFormat)
                .ChartArea.Axes("Y").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
            Next
        End If
        
        If InStr(str_Lines, "~") > 0 Then
            str����� = Split(str_Lines, "~")(1)
            str_Lines = Split(str_Lines, "~")(0)
        End If
        varLine = Split(str_Lines, "|")
        For i = 0 To UBound(varLine)
            aLine = Split(varLine(i), ",")
            .ChartGroups(1).Data.NumSeries = i + 1
            .ChartGroups(1).Data.NumPoints(i + 1) = UBound(aLine) + 1
            
            .ChartGroups(1).Styles(i + 1).Line.Color = vbBlack
            .ChartGroups(1).Styles(i + 1).Line.Width = 1
            .ChartGroups(1).Styles(i + 1).Symbol.Shape = oc2dShapeNone
            
            For y = 0 To UBound(aLine)
                .ChartGroups(1).Data.y(i + 1, y + 1) = IIf(Val(aLine(y)) >= 1, aLine(y), .ChartGroups(1).Data.HoleValue)
            Next
        Next
        
        If str����� <> "" Then
            aLables = Split(str�����, "|") 'X���ǩ��Y���ǩ
            
            aLable = Split(aLables(0), ",")
            strLable = aLable(0)
            If strLable <> "" Then
                .ChartLabels.Add
                .ChartLabels.Item(1).Text = strLable
                .ChartLabels.Item(1).AttachDataCoord.x = Val(aLable(1))
                .ChartLabels.Item(1).AttachDataCoord.y = Val(aLable(2))
            End If
            
            aLable = Split(aLables(1), ",")
            strLable = aLable(0)
            If strLable <> "" Then
                .ChartLabels.Add
                .ChartLabels.Item(2).Text = strLable
                .ChartLabels.Item(2).AttachDataCoord.x = Val(aLable(1))
                .ChartLabels.Item(2).AttachDataCoord.y = Val(aLable(2))
            End If
        End If
        .IsBatched = False
    
        If intSaveType = 1 Then
            ChartDrawPLT = .SaveImageAsJpeg(strFileName, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDrawPLT = .SaveImageAsPng(strFileName, False)
        Else
            ChartDrawPLT = .Save(strFileName)
        End If
    End With

End Function


'------
Private Function PicShowChart(ByVal strData As String, strFileName As String, Optional intSaveType As Integer = 0) As Boolean
    '��PIC�ϻ�ͼȻ��ת�浽Chart�ؼ�����ʾ
    Dim strImgName As String
    Dim frmPic As New frmLisGraph
    
    If strData <> "" Then
        With frmPic
            If DrawGam(.Picture1, strData) Then
                strImgName = strFileName
                If InStr(strImgName, ".") > 0 Then
                    strImgName = Mid(strImgName, 1, InStr(strImgName, ".")) & "Bmp"
                    If Dir(strImgName) <> "" Then Kill strImgName
                    SavePicture .Picture1.Image, strImgName
                    Call ChartShowPic("", strImgName, strFileName, 5, intSaveType)
                    PicShowChart = True
                    'If mobjFSO.FileExists(strImgName) Then mobjFSO.DeleteFile strImgName
                End If
            End If
        End With
    
    End If
    If Dir(strImgName) <> "" Then Kill strImgName
    Unload frmPic
End Function

Private Function DrawGam(ByRef objPic As PictureBox, ByVal strData As String) As Boolean
        '���ݸ�ʽ:
        '����;ͼ������;Y�߶�,X����;�������ұ߿�;X��̶�[|Y�̶�];��������[;�������]
        ' ����:��������: ��y��������,��,�ָ�,��������������|�ָ�
        '     :�������: ��x��������,��,�ŷָ�
        Dim curX As Currency, curY As Currency, curLastY As Currency
        Dim intWidth As Integer, intHight As Integer
        Dim str�������ұ߿� As String
        Dim int�߿��� As Integer, int�߿��� As Integer, int�߿��� As Integer, int�߿��� As Integer
        Dim lngColor As Long
    
        Dim varData As Variant
        Dim str���� As String, str���� As String, strLineData As String, str�̶� As String
        Dim var����  As Variant, strK As String, curK As Currency, var�̶� As Variant
        Dim str��� As String
        Dim varLindData As Variant
        Dim lngLoop As Long
        Dim curOldW  As Currency, curOldH As Currency, curOldSW As Currency, curOldSH As Currency
        On Error GoTo errHandle
    
     varData = Split(strData, ";")
 
     If UBound(varData) < 3 Then Exit Function '---���ݲ�ȫ
 
     If varData(1) <> "5" Then Exit Function '---��ʽ����
     str���� = varData(0)
     str���� = varData(2)
     str�������ұ߿� = varData(3)
     str�̶� = varData(4)
     strLineData = varData(5)
     varLindData = Split(strLineData, "|")
 
     If UBound(varData) > 5 Then
         str��� = varData(6)
     End If
 
     '�����С
     var���� = Split(str����, ",")
 
     intHight = var����(0): intWidth = var����(1)
     mcurYunti = intHight / intWidth
     If str�������ұ߿� = "" Then
         int�߿��� = 20: int�߿��� = 10: int�߿��� = 10 * mcurYunti: int�߿��� = 50 * mcurYunti
     Else

         int�߿��� = Split(str�������ұ߿�, ",")(0) * mcurYunti
         int�߿��� = Split(str�������ұ߿�, ",")(1) * mcurYunti
         int�߿��� = Split(str�������ұ߿�, ",")(2)
         int�߿��� = Split(str�������ұ߿�, ",")(3)
     End If
 
     objPic.Cls
     objPic.BackColor = vbWhite
     curOldW = objPic.Width
     curOldH = objPic.Height
 
     objPic.Width = 3000
     objPic.Height = 1500
     objPic.DrawMode = vbCopyPen 'ȱʡ ����
     objPic.DrawStyle = vbSolid  'VbSolid -ʵ�� VbDash-����
     objPic.DrawWidth = 1.5        '�߿�
     objPic.AutoRedraw = True
 
     'objpic.Height = objpic.Width * (intHight / intWidth)
 
     Dim curTmp As Currency
     curOldSW = objPic.ScaleWidth
     curOldSH = objPic.ScaleHeight
 
     curTmp = objPic.ScaleWidth / (intWidth + int�߿��� + int�߿���)
     Xpixe = curTmp / Screen.TwipsPerPixelX  '����һ��X��=��������
 
     curTmp = objPic.ScaleHeight / (intHight + int�߿��� + int�߿���)
     Ypixe = curTmp / Screen.TwipsPerPixelY
 
     objPic.Scale (0, 0)-(intWidth + int�߿��� + int�߿���, intHight + int�߿��� + int�߿���)
     '������
     curX = int�߿���
     curLastY = 0
     For lngLoop = LBound(varLindData) To UBound(varLindData)
         strLineData = varLindData(lngLoop)
         Do While strLineData <> ""
             If InStr(strLineData, ",") > 0 Then
                 curK = Val(Mid(strLineData, 1, InStr(strLineData, ",") - 1))
                 strLineData = Mid(strLineData, InStr(strLineData, ",") + 1)
             Else
                 curK = Val(strLineData)
                 strLineData = ""
             End If
             curLastY = curY
         
             curX = curX + 1
         
             curY = (intHight - curK) + int�߿���
             If curX > int�߿��� + 1 And curLastY < intHight + int�߿��� - 2 * mcurYunti Then objPic.Line (curX, curY)-(curX - 1, curLastY), vbBlue
         Loop
     Next
     objPic.DrawWidth = 1        '�߿�
     '������
     lngColor = vbBlack
     objPic.Line (int�߿���, int�߿��� + intHight)-(int�߿��� + intWidth, int�߿��� + intHight), lngColor
     objPic.Line (int�߿���, int�߿��� + intHight)-(int�߿���, int�߿���), lngColor
 
     '�̶�
     With objPic
         .FontName = "����"
         .ForeColor = lngColor
         .FontBold = False
         .FontSize = 9
     End With
 
     If InStr(str�̶�, "|") > 0 Then
         var�̶� = Split(Split(str�̶�, "|")(0), ",")
     Else
         var�̶� = Split(str�̶�, ",")
     End If
     For lngLoop = LBound(var�̶�) To UBound(var�̶�)
         curK = Val(Split(var�̶�(lngLoop), "-")(0))
         strK = Split(var�̶�(lngLoop), "-")(1)
         Call DrawK_X(objPic, int�߿���, int�߿��� + intHight, curK, strK)
     Next
     If InStr(str�̶�, "|") > 0 Then
         var�̶� = Split(Split(str�̶�, "|")(1), ",")
         For lngLoop = LBound(var�̶�) To UBound(var�̶�)
             curK = Val(Split(var�̶�(lngLoop), "-")(0))
             strK = Split(var�̶�(lngLoop), "-")(1)
             Call DrawK_Y(objPic, int�߿���, int�߿��� + intHight, curK, strK, 5 / mcurYunti)
         Next
     End If

     '������
     objPic.DrawWidth = 1
     objPic.DrawStyle = vbDot
     Do While str��� <> ""
         If InStr(str���, ",") > 0 Then
             curK = Val(Mid(str���, 1, InStr(str���, ",") - 1))
             str��� = Mid(str���, InStr(str���, ",") + 1)
         Else
             curK = Val(str���)
             str��� = ""
         End If
         If curK <> 0 Then Call DrawK_X(objPic, int�߿���, int�߿��� + intHight - 10 * mcurYunti, curK, "", Val(intHight - 20 * mcurYunti))
     Loop

     '����
     If Trim(str����) <> "" Then
         With objPic
             .CurrentX = int�߿��� + intWidth - (Len(str����) * 12 / Xpixe)
             .CurrentY = int�߿��� - int�߿��� + 5 * mcurYunti
             .FontSize = 10
             .FontBold = True
         End With
         objPic.Print Trim(str����)
     End If
     objPic.Scale (0, 0)-(curOldSW, curOldSH)
     objPic.Width = curOldW
     objPic.Height = curOldH
     DrawGam = True
        Exit Function
errHandle:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub DrawK_X(ByRef objPic As PictureBox, ByVal curX As Currency, ByVal curY As Currency, ByVal curK As Currency, ByVal strK As String, Optional curHight As Currency = 10)
    '��X��̶�
    objPic.Line (curX + curK, curY)-(curX + curK, curY - curHight)
    If strK <> "" Then Call PrintRotText(objPic.hDC, strK, (curX + curK) * Xpixe, curY * Ypixe + 8, 0)
End Sub

Private Sub DrawK_Y(ByRef objPic As PictureBox, ByVal curX As Currency, ByVal curY As Currency, ByVal curK As Currency, ByVal strK As String, Optional curWidth As Currency = 10)
    '��Y��̶�
    objPic.Line (curX, curY - curK)-(curX + curWidth, curY - curK)
    If strK <> "" Then Call PrintRotText(objPic.hDC, strK, curX * Xpixe - 11, (curY - curK) * Ypixe, 0)
End Sub




