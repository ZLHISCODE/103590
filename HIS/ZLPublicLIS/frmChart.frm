VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmChart 
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
      Height          =   3000
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   3000
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   5292
      _ExtentY        =   5292
      _StockProps     =   0
      ControlProperties=   "frmChart.frx":0000
   End
   Begin C1Chart2D8.Chart2D ChartBlood 
      Height          =   5000
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   10583
      _ExtentY        =   8819
      _StockProps     =   0
      ControlProperties=   "frmChart.frx":0583
   End
End
Attribute VB_Name = "frmChart"
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

Public Function DrawImg(ByVal strType As String, ByVal strData As String, ByVal strFilename As String, Optional intSaveType As Integer = 0) As Boolean
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '����:              ��ͼ��֧��(RBC,PLT,BAS,WBC)
        '����:              strType  ͼ������
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
 
        On Error GoTo ErrH
    
100     aImage = Split(strData, ";")
    
102     If aImage(0) = 0 Then
            'ֱ��ͼ clsLISDev_ABX_P120
104         With Me.ChartThis
106             .IsBatched = True
108             .Reset
110             .ChartGroups(1).Data.NumSeries = 0
112             .Header.Adjust = oc2dAdjustCenter
114             .Header.Text = strType
116             .Header.Font.Bold = True
118             .Header.Font.Size = 12
120             .ChartGroups(1).Styles(1).Line.COLOR = vbBlack
122             .ChartGroups(1).Styles(1).Line.Width = 1
124             .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeNone
            
126             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
128             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
            
130             .ChartArea.Axes("X").ValueLabels.RemoveAll
        
132             aFormat = Split(aImage(1), ",")
134             If UBound(aFormat) > 1 Then
136                 .ChartArea.Axes("Y").Min = aFormat(0)
138                 .ChartArea.Axes("X").Max = aFormat(1)
                    '----2008-03-28 ��ʱ����ʾͼ��
140                 .ChartArea.Axes("Y").Origin = 0
142                 .ChartArea.Axes("Y").Min = 0
                
144                 .ChartArea.Axes("Y").Max.IsDefault = True
146                 .ChartArea.Axes("X").Max.IsDefault = True
148                 .ChartArea.Axes("Y").Min.IsDefault = True
150                 .ChartArea.Axes("X").Min.IsDefault = True
                
152                 For i = 2 To UBound(aFormat)
154                     .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
                    Next
                End If
156             .ChartGroups(1).Data.NumSeries = 1
158             .ChartGroups(1).Data.NumPoints(1) = UBound(aImage)
160             For i = 2 To UBound(aImage)
162                 .ChartGroups(1).Data.Y(1, i) = Val(aImage(i))
                Next
164             .IsBatched = False
166             If intSaveType = 1 Then
168                 DrawImg = .SaveImageAsJpeg(strFilename, 100, False, False, False)
170             ElseIf intSaveType = 2 Then
172                 DrawImg = .SaveImageAsPng(strFilename, False)
                Else
174                 DrawImg = .Save(strFilename)
                End If
            End With
        
        End If
    
        '-- ɢ��ͼ clsLISDev_ABX_P120
176     If aImage(0) = 1 Then
178         With Me.ChartThis
180             .IsBatched = True
182             .Reset
184             .ChartGroups(1).Data.NumSeries = 0
186             .Header.Adjust = oc2dAdjustCenter
188             .Header.Text = strType
190             .Header.Font.Bold = True
192             .Header.Font.Size = 12
194             .ChartArea.PlotArea.IsBoxed = True
196             .ChartGroups(1).Data.NumSeries = UBound(aImage) - 1
            
198             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
200             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
202             For lngLoop = UBound(aImage) - 1 To 1 Step -1
204                 .ChartGroups(1).ChartType = oc2dTypePlot
206                 .ChartGroups(1).Styles(lngLoop).Line.Pattern = oc2dLineNone
208                 .ChartGroups(1).Styles(lngLoop).Line.COLOR = vbBlack
210                 .ChartGroups(1).Styles(lngLoop).Symbol.Shape = oc2dShapeBox
212                 .ChartGroups(1).Styles(lngLoop).Symbol.Size = 1
214                 .ChartGroups(1).Styles(lngLoop).Symbol.COLOR = vbBlack
216                 .ChartGroups(1).Data.NumPoints(lngLoop) = Len(aImage(lngLoop)) + 1
218                 For lngLoopX = 1 To Len(aImage(lngLoop)) + 1
220                     .ChartGroups(1).Data.Y(lngLoop, lngLoopX) = IIf(Mid(aImage(lngLoop), lngLoopX, 1) = 0, .ChartGroups(1).Data.HoleValue, 128 - lngLoop + 1)
                    Next
                Next
222             .IsBatched = False
224             If intSaveType = 1 Then
226                 DrawImg = .SaveImageAsJpeg(strFilename, 100, False, False, False)
228             ElseIf intSaveType = 2 Then
230                 DrawImg = .SaveImageAsPng(strFilename, False)
                Else
232                 DrawImg = .Save(strFilename)
                End If
            End With
        End If

        '--- Ѫ����ͼ  clsLISDev_File_LBYN6C
234     If aImage(0) = 2 Then
236         DrawImg = ChartDrawѪ����(strType, aImage(1), aImage(2), aImage(3), strFilename, intSaveType)
        End If
    
        '--Ѫ������ͼ clsLISDev_File_LBYN6C
238     If aImage(0) = 3 Then
240         DrawImg = ChartDrawѪ��(strType, aImage(1), aImage(2), aImage(3), strFilename, intSaveType)
        End If
    
        '--- �������ص����ߵ�PLTͼ clsLISDev_HMX
242     If aImage(0) = 4 Then
244         DrawImg = ChartDrawPLT(strType, aImage(1), aImage(2), strFilename, intSaveType)
        End If
    
        '--- �ڱ��ص�PIC�ؼ��ϻ��� ֱ��ͼ Ȼ����ʾ
246     If aImage(0) = 5 Then
248         DrawImg = PicShowChart(strType & ";" & strData, strFilename, intSaveType)
        
        End If
    
        '--- ֱ��ͼ����ͬһ��ͼ�ϻ��ƶ�������
250     If aImage(0) = 6 Then
            'ֱ��ͼ WBC clsLISDev_MEDONIC_M20M
252         With Me.ChartThis
254             .IsBatched = True
256             .Reset
258             .ChartGroups(1).Data.NumSeries = 0
260             .Header.Adjust = oc2dAdjustCenter
262             .Header.Text = strType
264             .Header.Font.Bold = True
266             .Header.Font.Size = 12
            
            
268             .ChartArea.Axes("X").AnnotationMethod = oc2dAnnotateValueLabels
270             .ChartArea.Axes("Y").AnnotationMethod = oc2dAnnotateValueLabels
            
272             .ChartArea.Axes("X").ValueLabels.RemoveAll
274             aFormat = Split(aImage(1), ",")
276             If UBound(aFormat) > 1 Then
278                 .ChartArea.Axes("Y").Min = aFormat(0)
280                 .ChartArea.Axes("X").Max = aFormat(1)
                    '----2008-03-28 ��ʱ����ʾͼ��
282                 .ChartArea.Axes("Y").Origin = 0
284                 .ChartArea.Axes("Y").Min = 0
                
286                 For i = 2 To UBound(aFormat)
288                     .ChartArea.Axes("X").ValueLabels.Add Mid(aFormat(i), 1, InStr(aFormat(i), "-") - 1), Mid(aFormat(i), InStr(aFormat(i), "-") + 1)
                    Next
                End If
            
290             aLines = Split(strData, "~")
292             .ChartGroups(1).Data.NumSeries = UBound(aLines)
            
294             For i = LBound(aLines) + 1 To UBound(aLines)
296                 aImage = Split(aLines(i), ";")
298                 .ChartGroups(1).Styles(i).Line.COLOR = vbBlack
300                 .ChartGroups(1).Styles(i).Line.Width = 1
302                 .ChartGroups(1).Styles(i).Symbol.Shape = oc2dShapeNone
304                 .ChartGroups(1).Data.NumPoints(i) = UBound(aImage) + 1
306                 For j = LBound(aImage) To UBound(aImage)
308                     .ChartGroups(1).Data.Y(i, j + 1) = Val(aImage(j))
                    Next
                Next
310             .IsBatched = False
312             If intSaveType = 1 Then
314                 DrawImg = .SaveImageAsJpeg(strFilename, 100, False, False, False)
316             ElseIf intSaveType = 2 Then
318                 DrawImg = .SaveImageAsPng(strFilename, False)
                Else
320                 DrawImg = .Save(strFilename)
                End If
            End With
        End If
        killFile = ""
        
322     If aImage(0) >= 100 And aImage(0) <= 227 Then
324         strFile = aImage(1)
                
'326         If UCase$(strFile) Like "*.ZIP" Then
'328             killFile = strFile
'330             If aImage(0) >= 200 And aImage(0) <= 207 Then
'332                 strFile = zlFileUnzip(strFile)
'334             ElseIf aImage(0) >= 210 And aImage(0) <= 217 Then
'336                 strFile = zlFileUnzip(strFile)
'338             ElseIf aImage(0) >= 220 And aImage(0) <= 227 Then
'340                 strFile = zlFileUnzip(strFile)
'                End If
'342             If killFile <> "" Then Kill killFile: killFile = "" '��ѹ���ԭʼZIPҪɾ��
'            End If
        
344         If UCase(strFile) Like "*.JPG" Then
346             frmChartPic.picTmp.Picture = LoadPicture(strFile)
348             strFile = Replace(UCase(strFile), ".JPG", ".BMP")
350             SavePicture frmChartPic.picTmp, strFile
352             'killFile = Replace(strFile, ".BMP", ".JPG") '2012-05-24 ԭʼͼ�β�ɾ��
354         ElseIf UCase(strFile) Like "*.GIF" Then
356             If CheckGif(strFile) Then
358                 frmChartPic.picTmp.Picture = LoadPicture(strFile)
360                 strFile = Replace(UCase(strFile), ".GIF", ".BMP")
362                 SavePicture frmChartPic.picTmp, strFile
364                 'killFile = Replace(strFile, ".BMP", ".GIF") 'ԭʼͼ�β�ɾ��
                Else
366                 Call ErrLog("DrawImg", "Gif�ļ���ʽ����ȷ", strFile, "")
                    Exit Function
                End If

            End If
            '--- ֱ����ʾͼƬ clsLISDev_UF100_DY
370         If CInt(Val(Right(aImage(0), 1))) = 0 Then
372             DrawImg = ChartShowPic(strType, strFile, strFilename, , intSaveType)  '��Ĭ�ϵ�layOut
            Else
374             DrawImg = ChartShowPic(strType, strFile, strFilename, CInt(Val(Right(aImage(0), 1))), intSaveType)   '��ָ����layout
            End If
376         i = 0

        End If
    
        Exit Function
ErrH:
386     Call ErrLog("DrawImg", "��" & CStr(Erl()) & "��", err.Description, "")
End Function

Private Function ChartDrawѪ����(ByVal strType As String, ByVal strXYin As String, ByVal strLineIn As String, ByVal strLableIn As String, ByVal strFilename As String, Optional intSaveType As Integer) As Boolean
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
    
    With Me.ChartBlood
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
                
                .ChartGroups(1).Styles(i + 1).Line.COLOR = vbBlack
                .ChartGroups(1).Styles(i + 1).Line.Width = 1
                .ChartGroups(1).Styles(i + 1).Symbol.Shape = oc2dShapeNone
                
                .ChartGroups(1).Data.NumSeries = i + 1
                .ChartGroups(1).Data.NumPoints(i + 1) = 225
            
                    '����
                If UBound(aCurve) > 2 Then
                    For lngLoop = 1 To 200
                        dblY = GetNd(Val(aCurve(0)), Val(aCurve(1)), Val(aCurve(2)), Val(aCurve(3)), lngLoop)
                        If lngLoop > 2 Then
                            .ChartGroups(1).Data.Y(i + 1, lngLoop + 1) = dblY
                        Else
                            .ChartGroups(1).Data.Y(i + 1, lngLoop + 1) = .ChartGroups(1).Data.HoleValue
                        End If
                    Next
                End If
            Next
        
            aPoint = Split(aFormat(1), ",") '�������

            intPoint = UBound(aCurves) + 2
            .ChartGroups(1).Styles(intPoint).Line.Pattern = oc2dLineNone
            .ChartGroups(1).Styles(intPoint).Line.COLOR = vbBlack
            .ChartGroups(1).Styles(intPoint).Line.Width = 1
            .ChartGroups(1).Styles(intPoint).Symbol.COLOR = vbBlack
            .ChartGroups(1).Styles(intPoint).Symbol.Shape = oc2dShapeSquare
            .ChartGroups(1).Data.NumSeries = intPoint
            .ChartGroups(1).Data.NumPoints(intPoint) = 225

            For i = 1 To 200
                
                For lngLoop = LBound(aPoint) To UBound(aPoint)
                    '-- ���
                    dblX = Val(Mid(aPoint(lngLoop), 1, InStr(aPoint(lngLoop), "-") - 1))
                    dblY = Val(Mid(aPoint(lngLoop), InStr(aPoint(lngLoop), "-") + 1))
                    If dblX = i + 1 Then
                        .ChartGroups(1).Data.Y(intPoint, i + 1) = dblY
                        Exit For
                    Else
                        .ChartGroups(1).Data.Y(intPoint, i + 1) = .ChartGroups(1).Data.HoleValue
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
            .ChartLabels.Item(1).AttachDataCoord.X = Val(aLable(1))
            .ChartLabels.Item(1).AttachDataCoord.Y = Val(aLable(2))
        End If
        
        aLable = Split(aLables(1), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(2).Text = strLable
            .ChartLabels.Item(2).AttachDataCoord.X = Val(aLable(1))
            .ChartLabels.Item(2).AttachDataCoord.Y = Val(aLable(2))
        End If
        
        '---- ����

        .IsBatched = False
        If intSaveType = 1 Then
            ChartDrawѪ���� = .SaveImageAsJpeg(strFilename, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDrawѪ���� = .SaveImageAsPng(strFilename, False)
        Else
            ChartDrawѪ���� = .Save(strFilename)
        End If
    End With
End Function

Private Function ChartDrawѪ��(ByVal strType As String, ByVal strXYin As String, ByVal strLineIn As String, ByVal strLableIn As String, ByVal strFilename As String, Optional intSaveType As Integer) As Boolean
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
    
    With Me.ChartBlood
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

            .ChartGroups(1).Styles(1).Line.COLOR = vbBlack
            .ChartGroups(1).Styles(1).Line.Width = 1
            .ChartGroups(1).Styles(1).Symbol.Shape = oc2dShapeSquare '���ķ���
            .ChartGroups(1).Styles(1).Symbol.COLOR = vbRed
            
            .ChartGroups(1).Styles(2).Symbol.COLOR = vbRed
            .ChartGroups(1).Styles(2).Symbol.Shape = oc2dShapeSquare
            .ChartGroups(1).Styles(2).Line.COLOR = vbBlack
            
            .ChartGroups(1).Styles(3).Symbol.Shape = oc2dShapeNone
            .ChartGroups(1).Styles(3).Line.COLOR = vbBlack
            
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
                    .ChartGroups(1).Data.Y(1, i) = aFormat(i)
                Else
                    .ChartGroups(1).Data.Y(2, i) = aFormat(i)
                End If
                    '����
                .ChartGroups(1).Data.Y(3, i) = aFormat(i) - 0.3
                
            Next
        
        End If
        
        aLables = Split(strLableIn, "~") 'X���ǩ��Y���ǩ
        
        aLable = Split(aLables(0), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(1).Text = strLable
            .ChartLabels.Item(1).AttachDataCoord.X = Val(aLable(1))
            .ChartLabels.Item(1).AttachDataCoord.Y = Val(aLable(2))
        End If
        
        aLable = Split(aLables(1), ",")
        strLable = aLable(0)
        If strLable <> "" Then
            .ChartLabels.Add
            .ChartLabels.Item(2).Text = strLable
            .ChartLabels.Item(2).AttachDataCoord.X = Val(aLable(1))
            .ChartLabels.Item(2).AttachDataCoord.Y = Val(aLable(2))
        End If
        
        '---- ����

        .IsBatched = False

        If intSaveType = 1 Then
            ChartDrawѪ�� = .SaveImageAsJpeg(strFilename, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDrawѪ�� = .SaveImageAsPng(strFilename, False)
        Else
            ChartDrawѪ�� = .Save(strFilename)
        End If
    End With
End Function

Private Function ChartShowPic(ByVal strType As String, ByVal strImgName As String, ByVal strFilename As String, Optional ByVal intLayOut As Integer = oc2dImageFitted, Optional intSaveType As Integer = 0) As Boolean
    'Chart ��ʾͼƬ
    Dim strImgFile  As String
    Dim objFso      As New FileSystemObject
    
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
            strImgFile = Replace(strFilename, ".cht", ".bmp")
            objFso.CopyFile strImgName, strImgFile, True
        Else
            strImgFile = strFilename
            objFso.CopyFile strImgName, strImgFile, True
        End If
        .Interior.Image.Filename = strImgFile
        .Interior.Image.Layout = intLayOut 'oc2dImageFitted
'        .ChartArea.Interior.Image.Filename = strImgName
'        .ChartArea.Interior.Image.Layout = intLayOut
        .IsBatched = False
 
        If intSaveType = 1 Then
            ChartShowPic = .SaveImageAsJpeg(strFilename, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartShowPic = .SaveImageAsPng(strFilename, False)
        Else
            ChartShowPic = .Save(strFilename)
        End If
    End With
    Set objFso = Nothing
End Function

Private Function GetNd(ByVal ND0 As Double, QB0 As Double, ND1 As Double, QB1 As Double, ByVal Qb As Double) As Double
    'Ѫ�����׼����������㺯��
    Dim k0 As Double, k1 As Double
    Dim sn As Double

    k0 = (Sqr(ND0) - Sqr(ND1)) / (1 / (Sqr(QB0)) - 1 / (Sqr(QB1)))
    k1 = Sqr(ND0) - k0 * (1 / (Sqr(QB0)))
    sn = k0 * (1 / (Sqr(Qb))) + k1
    GetNd = sn * sn

End Function

Private Function ChartDrawPLT(ByVal strType As String, ByVal str_���� As String, ByVal str_Lines As String, ByVal strFilename As String, Optional intSaveType As Integer = 0) As Boolean
    
    Dim aFormat() As String
    Dim i As Integer, Y As Integer
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
        .ChartGroups(1).Styles(1).Line.COLOR = vbBlack
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
            
            .ChartGroups(1).Styles(i + 1).Line.COLOR = vbBlack
            .ChartGroups(1).Styles(i + 1).Line.Width = 1
            .ChartGroups(1).Styles(i + 1).Symbol.Shape = oc2dShapeNone
            
            For Y = 0 To UBound(aLine)
                .ChartGroups(1).Data.Y(i + 1, Y + 1) = IIf(Val(aLine(Y)) >= 1, aLine(Y), .ChartGroups(1).Data.HoleValue)
            Next
        Next
        
        If str����� <> "" Then
            aLables = Split(str�����, "|") 'X���ǩ��Y���ǩ
            
            aLable = Split(aLables(0), ",")
            strLable = aLable(0)
            If strLable <> "" Then
                .ChartLabels.Add
                .ChartLabels.Item(1).Text = strLable
                .ChartLabels.Item(1).AttachDataCoord.X = Val(aLable(1))
                .ChartLabels.Item(1).AttachDataCoord.Y = Val(aLable(2))
            End If
            
            aLable = Split(aLables(1), ",")
            strLable = aLable(0)
            If strLable <> "" Then
                .ChartLabels.Add
                .ChartLabels.Item(2).Text = strLable
                .ChartLabels.Item(2).AttachDataCoord.X = Val(aLable(1))
                .ChartLabels.Item(2).AttachDataCoord.Y = Val(aLable(2))
            End If
        End If
        .IsBatched = False
    
        If intSaveType = 1 Then
            ChartDrawPLT = .SaveImageAsJpeg(strFilename, 100, False, False, False)
        ElseIf intSaveType = 2 Then
            ChartDrawPLT = .SaveImageAsPng(strFilename, False)
        Else
            ChartDrawPLT = .Save(strFilename)
        End If
    End With

End Function

'------
Private Function PicShowChart(ByVal strData As String, strFilename As String, Optional intSaveType As Integer = 0) As Boolean
    '��PIC�ϻ�ͼȻ��ת�浽Chart�ؼ�����ʾ
    Dim frmPic As New frmGraph
    Dim strImgName As String
    
    If strData <> "" Then
        With frmPic
            If DrawGam(.picImg, strData) Then
                strImgName = strFilename
                If InStr(strImgName, ".") > 0 Then
                    strImgName = Mid(strImgName, 1, InStr(strImgName, ".")) & "Bmp"
                    If Dir(strImgName) <> "" Then Kill strImgName
                    SavePicture .picImg.Image, strImgName
                    Call ChartShowPic("", strImgName, strFilename, 5, intSaveType)
                    PicShowChart = True
                    'If objFso.FileExists(strImgName) Then objFso.DeleteFile strImgName
                End If
            End If
        End With
    
    End If
    'If Dir(strImgName) <> "" Then Kill strImgName
    Set frmPic = Nothing
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
    
100     varData = Split(strData, ";")
    
102     If UBound(varData) < 3 Then Exit Function '---���ݲ�ȫ
    
104     If varData(1) <> "5" Then Exit Function '---��ʽ����
106     str���� = varData(0)
108     str���� = varData(2)
110     str�������ұ߿� = varData(3)
112     str�̶� = varData(4)
114     strLineData = varData(5)
116     varLindData = Split(strLineData, "|")
    
118     If UBound(varData) > 5 Then
120         str��� = varData(6)
        End If
    
        '�����С
122     var���� = Split(str����, ",")
    
124     intHight = var����(0): intWidth = var����(1)
126     mcurYunti = intHight / intWidth
128     If str�������ұ߿� = "" Then
130         int�߿��� = 20: int�߿��� = 10: int�߿��� = 10 * mcurYunti: int�߿��� = 50 * mcurYunti
        Else

132         int�߿��� = Split(str�������ұ߿�, ",")(0) * mcurYunti
134         int�߿��� = Split(str�������ұ߿�, ",")(1) * mcurYunti
136         int�߿��� = Split(str�������ұ߿�, ",")(2)
138         int�߿��� = Split(str�������ұ߿�, ",")(3)
        End If
    
140     objPic.Cls
142     objPic.BackColor = vbWhite
144     curOldW = objPic.Width
146     curOldH = objPic.Height
    
148     objPic.Width = 3000
150     objPic.Height = 1500
152     objPic.DrawMode = vbCopyPen 'ȱʡ ����
154     objPic.DrawStyle = vbSolid  'VbSolID -ʵ�� VbDash-����
156     objPic.DrawWidth = 1.5        '�߿�
158     objPic.AutoRedraw = True
    
        'objpic.Height = objpic.Width * (intHight / intWidth)
    
        Dim curTmp As Currency
160     curOldSW = objPic.ScaleWidth
162     curOldSH = objPic.ScaleHeight
    
164     curTmp = objPic.ScaleWidth / (intWidth + int�߿��� + int�߿���)
166     Xpixe = curTmp / Screen.TwipsPerPixelX  '����һ��X��=��������
    
168     curTmp = objPic.ScaleHeight / (intHight + int�߿��� + int�߿���)
170     Ypixe = curTmp / Screen.TwipsPerPixelY
    
172     objPic.Scale (0, 0)-(intWidth + int�߿��� + int�߿���, intHight + int�߿��� + int�߿���)
        '������
174     curX = int�߿���
176     curLastY = 0
178     For lngLoop = LBound(varLindData) To UBound(varLindData)
180         strLineData = varLindData(lngLoop)
182         Do While strLineData <> ""
184             If InStr(strLineData, ",") > 0 Then
186                 curK = Val(Mid(strLineData, 1, InStr(strLineData, ",") - 1))
188                 strLineData = Mid(strLineData, InStr(strLineData, ",") + 1)
                Else
190                 curK = Val(strLineData)
192                 strLineData = ""
                End If
194             curLastY = curY
            
196             curX = curX + 1
            
198             curY = (intHight - curK) + int�߿���
200             If curX > int�߿��� + 1 And curLastY < intHight + int�߿��� - 2 * mcurYunti Then objPic.Line (curX, curY)-(curX - 1, curLastY), vbBlue
            Loop
        Next
202     objPic.DrawWidth = 1        '�߿�
        '������
204     lngColor = vbBlack
206     objPic.Line (int�߿���, int�߿��� + intHight)-(int�߿��� + intWidth, int�߿��� + intHight), lngColor
208     objPic.Line (int�߿���, int�߿��� + intHight)-(int�߿���, int�߿���), lngColor
    
        '�̶�
210     With objPic
212         .FontName = "����"
214         .ForeColor = lngColor
216         .FontBold = False
218         .FontSize = 9
        End With
    
220     If InStr(str�̶�, "|") > 0 Then
222         var�̶� = Split(Split(str�̶�, "|")(0), ",")
        Else
224         var�̶� = Split(str�̶�, ",")
        End If
226     For lngLoop = LBound(var�̶�) To UBound(var�̶�)
228         curK = Val(Split(var�̶�(lngLoop), "-")(0))
230         strK = Split(var�̶�(lngLoop), "-")(1)
232         Call DrawK_X(objPic, int�߿���, int�߿��� + intHight, curK, strK)
        Next
234     If InStr(str�̶�, "|") > 0 Then
236         var�̶� = Split(Split(str�̶�, "|")(1), ",")
238         For lngLoop = LBound(var�̶�) To UBound(var�̶�)
240             curK = Val(Split(var�̶�(lngLoop), "-")(0))
242             strK = Split(var�̶�(lngLoop), "-")(1)
244             Call DrawK_Y(objPic, int�߿���, int�߿��� + intHight, curK, strK, 5 / mcurYunti)
            Next
        End If
   
        '������
246     objPic.DrawWidth = 1
248     objPic.DrawStyle = vbDot
250     Do While str��� <> ""
252         If InStr(str���, ",") > 0 Then
254             curK = Val(Mid(str���, 1, InStr(str���, ",") - 1))
256             str��� = Mid(str���, InStr(str���, ",") + 1)
            Else
258             curK = Val(str���)
260             str��� = ""
            End If
262         If curK <> 0 Then Call DrawK_X(objPic, int�߿���, int�߿��� + intHight - 10 * mcurYunti, curK, "", Val(intHight - 20 * mcurYunti))
        Loop

        '����
264     If Trim(str����) <> "" Then
266         With objPic
268             .CurrentX = int�߿��� + intWidth - (Len(str����) * 12 / Xpixe)
270             .CurrentY = int�߿��� - int�߿��� + 5 * mcurYunti
272             .FontSize = 10
274             .FontBold = True
            End With
276         objPic.Print Trim(str����)
        End If
278     objPic.Scale (0, 0)-(curOldSW, curOldSH)
280     objPic.Width = curOldW
282     objPic.Height = curOldH
284     DrawGam = True
        Exit Function
errHandle:
286     WriteLog "frmChart DrawGam ", CStr(Erl()) & "��", ""
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




