Attribute VB_Name = "mdlImage"
Option Explicit
'--------------------------------------------------------
'��  �ܣ���ģ��Ϊͼ����ĺ������̵�
'�����ˣ����Σ��ƽ�
'�������ڣ�2004.6.12
'���̺����嵥��
'    subMPRLinenPhase():        ��ָ��Viewer������ͼ���ʸ��״�ؽ����Ƶ�Ϳ����ߣ�����ָ��ͼ���е����ý���ͬ��
'    subInitMPRLine()��         �Ե�ǰ�����б�ѡ��Viewer��ȫ��ͼ�񣬳�ʼ��ʸ��״�ؽ����Ƶ�Ϳ�����
'    funcPlaneRestructInit():   ʸ��״�ؽ���ʼ����ͬʱ��дͼ��Ĳ���ܸ߶Ⱥ���������
'    LeagelToACRebuild():       �ж�ͼ���Ƿ�����ʸ��״�ؽ�������
'    subGetArray():             ����Line1�����ߣ���Viewer�ĵ�һ��ͼ���ϵ�λ�ã��������ֱ����ÿһ�����������������
'    subACRebuild():            �ԻҶ�ֵ������в�ֵ��ƽ������
'    subGetLabelStoreToVar():   �����ݿ��ж�ȡ��ע������ʹ�õ�TAG��ϵͳ����
'    subaCorrectCursor():       ����ƶ��������ͼ��Χ�����������λ��
'    funAutoWinWL():            ����Ӧ����
'    subStackEnd():             �������
'    subLabelCopyRebuild():     �ؽ�ͼ��ı�ע������ϵ
'    ResizeRegion():            �Զ�����ָ�������ڣ�һ����Ŀͼ������е�������Ŀ
'    subSetWidthLevelF():       ���ô���λ���ܼ������˵�
'    GetAngle():                ����ͨ�����������ɵ�������֮��ĽǶ�
'    Max7InArray():             ����������ȡֵ����7���±꣬������ƽ��ֵ
'    funIsShutter():            �ж������Ӱ������Ƿ���Ҫ����ͼ����������
'    subDrawImgShutter():       ����ϵͳ���õ�Ӱ����𣬸������ͼ��ͼ������
'    funGetLinePoints():        ��ͼ��ĸ������ͱ�ע��ֱ�ߡ����ߣ�����ȡ�Ҷ�ֵ�������㡢�յ�����
'    funGetVasEdge():           ����Ѫ����խ����������ֱ�߱�ע��Ԥ�����ֵ������Ѫ�ܱڵ����ꡣ
'    subDrawVasEdgeLine():      ����Ѫ����խ����������ֱ�߱�ע��Ѫ�ܱڵ����꣬ȷ��������Ѫ�ܱڶ�ֱ�ߵ�λ�á�
'    subCenterZoom()��          ��ͼ��������š��Ե�ǰviewer���ĵ�Ϊ�������ĵ㡣
'�޸ļ�¼��
'    2005.7.07    �ƽ�
'    2005.8.19    �ƽ�
'    2005.9.15    �ƽ�
'    2006-2-10    �ƽ�
'-------------------------------------------------------

Public ToltalHeight As Integer                         ''�ؽ����ܸ߶�
Public aPixels() As Integer                                  ''�����ؽ�����ֵ������

Public Sub subMPRLinenPhase(v As DicomViewer, im As DicomImage)
'------------------------------------------------
'���ܣ���ָ��Viewer������ͼ���ʸ��״�ؽ����Ƶ�Ϳ����ߣ�����ָ��ͼ���е����ý���ͬ��
'������v--����ͼ����ʸ��״���Ʊ�עͬ����Viewer��im--��Ϊͬ����׼��ͼ��
'���أ��ޣ�ֱ�ӽ�v������ͼ���ʸ��״���Ƶ��߽������á�
'2009 Ҫ�޸ģ��ĳ�ֻ����ʾ��ͼ������޸�
'------------------------------------------------
    Dim img As DicomImage, i As Integer
    For Each img In v.Images
        For i = G_INT_SYS_LABEL_MPRV To G_INT_SYS_LABEL_MPR_POINT_O
            img.Labels(i).Visible = im.Labels(i).Visible
            img.Labels(i).left = im.Labels(i).left
            img.Labels(i).width = im.Labels(i).width
            img.Labels(i).top = im.Labels(i).top
            img.Labels(i).height = im.Labels(i).height
        Next
        img.Refresh False
    Next
End Sub

Public Sub subInitMPRLine(thisViewer As DicomViewer)
'------------------------------------------------
'���ܣ��Ե�ǰ�����б�ѡ��Viewer��ȫ��ͼ�񣬳�ʼ��ʸ��״�ؽ����Ƶ�Ϳ�����
'������     thisViewer--����ʸ��״�ؽ���Viewer
'���أ��ޣ�ֱ�Ӷ�imͼ���ϵ�ʸ��״�ؽ���ע����ʼ����
'2009��
'------------------------------------------------
     Dim im As DicomImage
     
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     If thisViewer.Images.Count = 0 Then Exit Sub
     
     Set im = thisViewer.Images(1)
     
     Call funInitMPRControlLines(im, True)
     
     subMPRLinenPhase thisViewer, im
End Sub

Public Function funInitMPRControlLines(im As DicomImage, blnVisible As Boolean)
'------------------------------------------------
'���ܣ���ʼ��ָ��ͼ���ʸ��״�ؽ����Ƶ�Ϳ�����
'������     im--����ʸ��״�ؽ�����λͼ��
'���أ��ޣ�ֱ�Ӷ�imͼ���ϵ�ʸ��״�ؽ���ע����ʼ����
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    For i = G_INT_SYS_LABEL_MPRV To G_INT_SYS_LABEL_MPR_POINT_O
        im.Labels(i).Visible = blnVisible
        If i >= G_INT_SYS_LABEL_MPR_POINT_V1 Then
            im.Labels(i).width = G_INT_MPR_RADIUS
            im.Labels(i).height = G_INT_MPR_RADIUS
        End If
    Next
    
    im.Labels(G_INT_SYS_LABEL_MPRV).left = im.sizex / 2
    im.Labels(G_INT_SYS_LABEL_MPRV).top = 0
    im.Labels(G_INT_SYS_LABEL_MPRV).height = im.sizey
    im.Labels(G_INT_SYS_LABEL_MPRV).width = 0
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPRH).left = 0
    im.Labels(G_INT_SYS_LABEL_MPRH).top = im.sizey / 2
    im.Labels(G_INT_SYS_LABEL_MPRH).height = 0
    im.Labels(G_INT_SYS_LABEL_MPRH).width = im.sizex
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).left = im.sizex / 2 - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).top = -G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).left = im.sizex / 2 - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).top = im.sizey - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).left = -G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).top = im.sizey / 2 - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).left = im.sizex - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).top = im.sizey / 2 - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = im.sizex / 2 - G_INT_MPR_RADIUS / 2
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = im.sizey / 2 - G_INT_MPR_RADIUS / 2
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Refresh False
     
    funInitMPRControlLines = True
    Exit Function
err:
    If ErrCenter = 1 Then Resume
End Function

Public Function funcPlaneRestructInit(Viewer As DicomViewer, thisForm As frmViewer) As Boolean
'------------------------------------------------
'���ܣ�ʸ��״�ؽ���ʼ����ͬʱ��дͼ��Ĳ���ܸ߶Ⱥ���������
'������ viewer--����ʸ��״�ؽ���viewer
'       thisForm -- ��ʾͼ��Ĵ���
'���أ�True--��ʼ���ɹ����Խ����ؽ���False--��ʼ��ʧ�ܣ����ܹ������ؽ�
'2009��
'------------------------------------------------
    Dim iHeight As Integer
    Dim iPixSpacing As Double
    Dim v As Variant
    Dim i As Integer
    Dim ix As Integer
    Dim iy As Integer
    
    funcPlaneRestructInit = False
    
    On Error GoTo err
    ''''��ȡͼ��Ĳ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''��ǰ���Ѿ������µ� ���λ�ã�ͼ�����������ؾ������˼�飬�������в����ټ���ˡ�
    v = Viewer.Images(1).Attributes(&H28, &H30).Value
    iPixSpacing = v(1)
    iHeight = Viewer.Images.Count
    ''''''ȷ��ͼ����ܸ߶ȣ�������Slice Location�Ĳ�Ͳ������Ľ��֮��ȡ���ֵ'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ToltalHeight = Abs(Viewer.Images(Viewer.Images.Count).Attributes(&H20, &H1041).Value - Viewer.Images(1).Attributes(&H20, &H1041).Value) / iPixSpacing
    
    '''''''''''''''''''�ض�����������ֵ����''''''''''''''''''''''''''''''''''''''
    zl9ComLib.zlCommFun.ShowFlash "���ڳ�ʼ��MPR�ؽ�����ȴ���", thisForm
    zl9ComLib.zlCommFun.ShowFlash
    
    '�������MPR����ά���鳬���ڴ���ɷ�Χ������֡��ڴ����������aPixelsά��=0��������ֱ����ͼ��������MPR
    ReDim aPixels(Viewer.Images(1).sizex, Viewer.Images(1).sizey, Viewer.Images.Count) As Integer
    For i = 1 To Viewer.Images.Count
        v = Viewer.Images(i).Pixels
        For ix = 1 To Viewer.Images(i).sizex
            For iy = 1 To Viewer.Images(i).sizey
                aPixels(ix, iy, i) = v(ix, iy, 1)
            Next
        Next
    Next
    funcPlaneRestructInit = True
    zl9ComLib.zlCommFun.StopFlash
    Exit Function
err:
    funcPlaneRestructInit = False
    zl9ComLib.zlCommFun.StopFlash
End Function

Private Function LeagelToMPR(imgs As DicomImages) As Long
'------------------------------------------------
'���ܣ��ж�ͼ���Ƿ�����ʸ��״�ؽ�������
'������imgs����ʸ��״�ؽ���ͼ��
'���أ�0--���Խ����ؽ���1--���ܽ����ؽ�

'------------------------------------------------
    Dim i As Long
    Dim j As Long
    
    LeagelToMPR = 1
    
    ''''''ͼ�������Ƿ�3��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (imgs.Count < 3) Then
       Exit Function
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim SeriesUID As String
    Dim thickness As Double
    Dim location() As Double
    ReDim location(imgs.Count) As Double
    Dim v As Variant
    Dim PixelSpacing As Double
    '''''''�����һ��ͼ�������UID''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H20, &HE)) Then
        SeriesUID = imgs(1).Attributes(&H20, &HE).Value
    Else
        Exit Function
    End If
    '''''''�����һ��ͼ��Ĳ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H18, &H50)) Then
        thickness = imgs(1).Attributes(&H18, &H50).Value
    Else
        Exit Function
    End If
    '''''''''�����һ��ͼ�����Ƭλ��Slice Location''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H20, &H1041)) Then
           location(1) = imgs(1).Attributes(&H20, &H1041).Value
    Else
       Exit Function
    End If
    ''''''''�����һ��ͼ������ؼ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(imgs(1).Attributes(&H28, &H30)) Then
        v = imgs(1).Attributes(&H28, &H30).Value
        PixelSpacing = v(1)
    Else
        Exit Function
    End If
    '''''''������ͼ����ѭ�����ж��Ƿ���������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 2 To imgs.Count
        '''''�ж��Ƿ�����ͬ������UID''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H20, &HE)) Then
            If SeriesUID <> imgs(i).Attributes(&H20, &HE).Value Then
                Exit Function
            End If
        Else
            Exit Function
        End If
        ''''''''�ж��Ƿ�����ͬ�Ĳ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H18, &H50)) Then
            If thickness <> imgs(i).Attributes(&H18, &H50).Value Then
                Exit Function
            End If
        Else
            Exit Function
        End If
        ''''''''����ͼ���λ��SliceLocation'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H20, &H1041)) Then
           location(i) = imgs(i).Attributes(&H20, &H1041).Value
        Else
           Exit Function
        End If
        '''''''�ж��Ƿ�����ͬ�����ؼ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not IsNull(imgs(i).Attributes(&H28, &H30)) Then
            v = imgs(i).Attributes(&H28, &H30).Value
            If PixelSpacing <> v(1) Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    Next
    ''''''''�ж��Ƿ��в���ͬ����Ƭλ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To imgs.Count
       For j = 1 To imgs.Count - i
          If location(i) = location(i + j) Then
             Exit Function
          End If
       Next
    Next
    '''''���������򷵻���''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    LeagelToMPR = 0
End Function

Public Function LeagelToACRebuild(imgs As DicomImages) As Long
'------------------------------------------------
'���ܣ��ж�ͼ���Ƿ�����ʸ��״�ؽ�������
'������imgs����ʸ��״�ؽ���ͼ��
'���أ�0--���Խ����ؽ���1--���ܽ����ؽ�������ʾ
'------------------------------------------------
    
    LeagelToACRebuild = LeagelToMPR(imgs)
    
    If LeagelToACRebuild = 1 Then
        MsgBox "ͼ���ܽ���ʸ��״�ؽ�����������������֮һ��" & vbCrLf & vbCrLf & _
              "3��ͼ�����ϣ�ͬһ���У���ͬ��񣻲�ͬλ�ã���ͬ���ؾ��롣", vbInformation, gstrSysName
    End If
End Function

Public Sub subGetArray(Line1 As DicomLabel, Image As DicomImage, LineLong() As POINTAPI)
'------------------------------------------------
'���ܣ�����Line1�����ߣ���Viewer�ĵ�һ��ͼ���ϵ�λ�ã��������ֱ����ÿһ�����������������
'������Line1--ʸ��״�����ߣ�Image--����ʸ��״�ؽ���ͼ��LineLong()--��Ϊ����ֵ�ã�����ֱ�������е�����ꡣ
'���أ��ޣ�ֱ�ӽ�ֱ����ÿһ�������������ŵ�LineLong�����С�
'2009��
'------------------------------------------------
    Dim beginx As Integer, beginy As Integer
    Dim endx As Integer, endy As Integer
    Dim i As Integer
    Dim j As Integer
    Dim iH As Long
    Dim iW As Long
    Dim sizex As Integer
    Dim sizey As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sizex = Image.sizex
    sizey = Image.sizey
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Line1.width > 0 Then
        beginx = Line1.left
        beginy = Line1.top
        endx = Line1.left + Line1.width
        endy = Line1.top + Line1.height
    Else
        endx = Line1.left
        endy = Line1.top
        beginx = Line1.left + Line1.width
        beginy = Line1.top + Line1.height
    End If
    ''''����''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If beginx <= 0 Then beginx = 1
    If beginx > sizex Then beginx = sizex
    If beginy <= 0 Then beginy = 1
    If beginy > sizey Then beginy = sizey
    If endx <= 0 Then endx = 1
    If endx > sizex Then endx = sizex
    If endy <= 0 Then endy = 1
    If endy > sizey Then endy = sizey
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iH = Abs(endy - beginy) + 1
    iW = Abs(endx - beginx) + 1
    ''''''''''''����ڸ�''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If iW > iH Then
        ReDim LineLong(iW)
        j = 1
        For i = beginx To endx
            LineLong(j).x = i
            LineLong(j).y = IIf(((endy - beginy) / (endx - beginx)) * (i - beginx) + beginy > sizey, _
                    sizey, ((endy - beginy) / (endx - beginx)) * (i - beginx) + beginy)
            j = j + 1
        Next
     '''''''''�߶ȴ��ڿ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        ReDim LineLong(iH)
        j = 1
        If beginy > endy Then
        ''''''����begin��end��λ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim iTmep As Integer
            iTmep = beginx
            beginx = endx
            endx = iTmep
            iTmep = beginy
            beginy = endy
            endy = iTmep
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = beginy To endy
            LineLong(j).y = i
            LineLong(j).x = IIf(((endx - beginx) / (endy - beginy)) * (i - beginy) + beginx > sizex, _
                  sizex, ((endx - beginx) / (endy - beginy)) * (i - beginy) + beginx)
            j = j + 1
        Next
    End If
End Sub

Public Sub subACRebuild(a() As Integer, b() As Integer)
'------------------------------------------------
'���ܣ��ԻҶ�ֵ������в�ֵ��ƽ������
'������a()--����ͼ��ԭ���Ҷ�ֵ�Ķ�ά���飬��һά���У��ڶ�ά���У�
'      b()--����ͼ���ؽ����»Ҷ�ֵ�Ķ�ά���飬��һά���У��ڶ�ά���У�
'���أ��ޣ�ֱ��ʹ��b()�������ؽ����ɵ��»Ҷ����顣
'2009��
'------------------------------------------------

    '''''''''''''aΪ����ͼ��Ҷ�ֵ�Ķ�ά���飬��һά���У��ڶ�ά����'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim TolHeight As Long
   Dim lngWidth As Long
   Dim lngHeight As Long
   Dim dblThick As Double
   Dim dblResidua As Double
   Dim dblAccResidua As Double
   Dim lngThick As Long
   Dim lngThickAddOne As Long
   Dim intRealRows As Integer
   
   On Error GoTo err
   
   lngWidth = UBound(a, 1)      'ͼ��Ҷ�ֵ�����һά�ĳ��ȣ�ͼ���п����ߵĳ���
   lngHeight = UBound(a, 2)     'ͼ��Ҷ�ֵ����ڶ�ά�ĳ��ȣ�ͼ�������
   TolHeight = UBound(b, 2)
   
    ''''''''''��a�ж�ȡһ�У���b�б任��������ΪSThicknessָ��������,��a��ÿһ����ѭ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Dim i As Long, j As Long, k As Long
   dblThick = TolHeight / lngHeight
   lngThick = Int(dblThick)             'ȥβȡ����������������
   lngThickAddOne = lngThick + 1
   dblResidua = dblThick - lngThick     'ȡ����
   dblAccResidua = 0
   dblAccResidua = dblAccResidua + dblResidua     '�ۼ�����
   intRealRows = 0
   For i = 0 To lngHeight - 1
        dblAccResidua = dblAccResidua + dblResidua
        If dblAccResidua >= 1 Then
            dblAccResidua = dblAccResidua - 1
            For j = 1 To lngWidth
                For k = 1 To lngThickAddOne
                    b(j, intRealRows + k) = a(j, i + 1)
                Next
            Next
            intRealRows = intRealRows + lngThickAddOne
        Else
            For j = 1 To lngWidth
                For k = 1 To lngThick
                    b(j, intRealRows + k) = a(j, i + 1)
                Next
            Next
            intRealRows = intRealRows + lngThick
        End If
        
   Next
    ''''''������b�еĵ���ƽ������,���ü�Ȩģ�壬������ƽ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call funImageSmoothing(b, Int(IIf(dblThick / 2 > 5, 5, dblThick / 2)))
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub subGetLabelStoreToVar()
'------------------------------------------------
'���ܣ������ݿ��ж�ȡ��ע������ʹ�õ�TAG��ϵͳ����
'��������
'���أ���
'2009��
'------------------------------------------------
   Dim strSQL As String
   Dim cOneAttr As clsLabelAttr
   Dim i As Integer
   '''''''��ȡϵͳ����������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If blLocalRun = True Then
      strSQL = "SELECT VGroup,Element,VR,��ע���� FROM Ӱ���ע�洢��"
      Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
   Else
      strSQL = "SELECT VGroup,Element,VR,��ע���� FROM Ӱ���ע�洢��"
      Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
   End If
   '''''��ռ������������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   For i = 1 To cLabelStore.Count
       cLabelStore.Remove 1
   Next i
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   On Error GoTo err
   
   With rsTemp
       .MoveFirst
       While Not .EOF
           Set cOneAttr = New clsLabelAttr
           cOneAttr.AttrName = !��ע����
           cOneAttr.Group = !VGroup
           cOneAttr.Element = !Element
           cOneAttr.VR = !VR
           ''''''�ӵ���������'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           cLabelStore.Add cOneAttr, cOneAttr.AttrName
           .MoveNext
       Wend
   End With
   Exit Sub
err:
   If blLocalRun = True Then
        MsgBox "��������:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
End Sub

Public Sub subaCorrectCursor(v As DicomViewer, im As DicomImage, xx As Long, Yy As Long)
'------------------------------------------------
'���ܣ�����ƶ��������ͼ��Χ�����������λ��
'������v--ͼ�����ڵ�viewer��im--������ڵ�ͼ��xx--������ڵ�x����λ�ã������곬��ͼ���򽫴�ֵ�޸ĵ�ͼ��֮�ڣ�
'      yy--������ڵ�y����λ�ã������곬��ͼ���򽫴�ֵ�޸ĵ�ͼ��֮�ڣ�
'���أ���
'2009��
'------------------------------------------------
    Dim x As Integer, y As Integer, w As Long, h As Long
    Dim i As DicomImage
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    w = v.width / v.MultiColumns / Screen.TwipsPerPixelX - v.CellSpacing * 2
    h = v.height / v.MultiRows / Screen.TwipsPerPixelY - v.CellSpacing * 2
    x = im.OriginX + v.CellSpacing
    y = im.OriginY + v.CellSpacing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If xx < x Then xx = x
    If xx > x + w Then xx = x + w
    If Yy < y Then Yy = y
    If Yy > y + h Then Yy = y + h
End Sub

Function funAutoWinWL(img As DicomImage, left As Long, top As Long, width As Long, _
                      height As Long, ByRef ww As Long, ByRef wl As Long) As Boolean
'------------------------------------------------
'���ܣ�����Ӧ������
'�㷨˵�������õķ����ǴӸ����ľ��������У���ȡȫ�����ص�ĻҶ�ֵ������Ϊ�����������Ҷ�ֵ����С�Ҷ�ֵ֮����90%��
'         ��λΪ��ǰ�Ҷ�ֵ����7����ĻҶ�ƽ��ֵ��
'������img--��Ҫ��������Ӧ������ͼ��(Left,Top,Width ,Height)--��ͼ������Ҫ��������Ӧ�����ľ�������
'      ww--���ش���ֵ��wl--���ش�λֵ��
'���أ�True--ִ�гɹ���Fasle--ִ��ʧ�ܡ�
'2009��
'------------------------------------------------
    Dim iBitWidth As Integer
    Dim tImg As DicomImage
    Dim lngMax As Long
    Dim lngMin As Long
    Dim aImg As Variant
    ''''''��ʼ������ֵ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    funAutoWinWL = False
    ww = img.width
    wl = img.Level
    '''''''''��ȡͼ��Ĵ洢λ����Ϣ,������Ϣ�����ڣ��򷵻ش���'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not IsNull(img.Attributes(&H28, &H100)) Then
        iBitWidth = img.Attributes(&H28, &H100).Value
        iBitWidth = iBitWidth / 8
    Else
        Exit Function
    End If
    ''''''''���ڿ�Ⱥ͸߶�ͬʱΪ1��ͼ�����򣬲�����������Ӧ����λ����Ϊ��ʱͨ����ͼ�ò���������С����ֵ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Abs(width) <= 1 And Abs(height) <= 1 Then
        Exit Function
    End If
    ''''''''�ж�ͼ�������Ƿ�ԭͼ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (left = 0) And (top = 0) And (width = img.sizex) And (height = img.sizey) Then
        Set tImg = img
    Else
        ''''''''�����������ѡȡ������������ϽǺ͸߿���ʱ�߿���Ҫ������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If width < 0 Then
            left = left + width
            width = -width
        End If
        If height < 0 Then
            top = top + height
            height = -height
        End If
        Set tImg = img.SubImage(left, top, width, height, 1, 1)
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If tImg.MinimumPixelValue(False) > 0 Then
        aImg = tImg.Histogram(tImg.MinimumPixelValue(False), tImg.MaximumPixelValue(False), iBitWidth)
        wl = tImg.MinimumPixelValue(False) + Max7InArray(aImg, lngMax, lngMin) * iBitWidth
    End If
    ww = Abs(tImg.MaximumPixelValue(False) - tImg.MinimumPixelValue(False)) * 0.9
    funAutoWinWL = True
End Function

Public Sub subStackEnd(v As DicomViewer, f As frmViewer)
'------------------------------------------------
'���ܣ��������
'������v--���д����viewer��f--����Ĵ��塣
'���أ���
'2009��
'------------------------------------------------
    Dim i As Integer
    i = f.MSFViewer.TextMatrix(f.intSelectedSerial, 3)
    v.Images.Add f.objStackOldImage
    v.Images.Move v.Images.Count, i
    subLabelCopyRebuild f.objStackOldImage, v.Images(i)
    v.Images.Remove i + 1
End Sub

Public Sub subLabelCopyRebuild(Simg As DicomImage, oImg As DicomImage)
'------------------------------------------------
'���ܣ��ؽ�ͼ��ı�ע������ϵ
'������sImg--Դͼ��oImg--Ŀ��ͼ��
'���أ���
'2009��
'------------------------------------------------
    Dim l As DicomLabel
    For Each l In oImg.Labels
        If Not l.TagObject Is Nothing Then
            If Simg.Labels.IndexOf(l.TagObject) <> 0 Then
                Set l.TagObject = oImg.Labels(Simg.Labels.IndexOf(l.TagObject))
            End If
        End If
    Next
End Sub

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
'-----------------------------------------------------------------------------
'���ܣ����������ͼ��������ͼ������Ŀ�Ⱥ͸߶ȣ�������ѵ�ͼ����������������
'������ ImageCount����ͼ������
'       RegionWidth--ͼ����ʾ����Ŀ��
'       RegionHeight--ͼ����ʾ����ĸ߶�
'       Rows����[����]�������
'       Cols����[����]�������
'       MaxRows ������ѡ���������
'       MaxCols ������ѡ���������
'���أ������������Rows���������Cols
'2009��
'-----------------------------------------------------------------------------
    Dim iCols As Integer, iRows As Integer
    
    On Error GoTo err
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))

    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols > ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols - 1
        Else
            iRows = iRows - 1
        End If
    Loop
    
    If iRows < 1 Then iRows = 1
    If iCols < 1 Then iCols = 1
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    
    '�������ֵ����ȷ������������
    If (MaxRows <> 0) And (MaxRows < iRows) Then
        Rows = MaxRows
    Else
        Rows = iRows
    End If
    If (MaxCols <> 0) And (MaxCols < iCols) Then
        Cols = MaxCols
    Else
        Cols = iCols
    End If
'����������
err:
End Sub

Public Sub subSetFilterF(im As DicomImage, f As frmViewer, Optional cbrPopup As CommandBarPopup)
'------------------------------------------------
'���ܣ������˾�ģ�幦�ܼ������˵�
'������ im--�����˾�ģ��Ļ�׼ͼ����ȡͼ���Modality
'       f--���õ����˵��Ĵ��壻
'       cbrPopup -- ������˵��������ӵ����˵�
'���أ���
'-----------------------------------------------
    Dim strModality As String
    Dim ControlPopup As CommandBarPopup
    Dim cbrToolBar As CommandBarControl
    Dim i As Integer
    Dim MenuPopup As CommandBarPopup    '���˵��еĵ����˵���
    Dim cbrMenuBar As CommandBarControl '���˵��еĲ˵���
    
    If im Is Nothing Then Exit Sub
    If IsNull(im.Attributes(&H8, &H60).Value) Then Exit Sub         '��ȡModality
    strModality = UCase(im.Attributes(&H8, &H60).Value)
    If cbrPopup Is Nothing Then
        Set ControlPopup = f.ComToolBar.Item(toolBar_PhotoStrong).FindControl(, ID_Active_SieveLens_Model, , True)
        ControlPopup.CommandBar.Controls.DeleteAll      '���ԭ�е����˵�������
        
        '���ԭ�����˵��еĵ����˵���
        Set MenuPopup = f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_SieveLens_Model, , True)
        MenuPopup.CommandBar.Controls.DeleteAll      '���ԭ�е����˵�������
    Else
        Set ControlPopup = cbrPopup
    End If
    
    '�����µĵ����˵�
    For i = 1 To UBound(aPresetFilter)
        If UCase(aPresetFilter(i - 1).strModality) = strModality Then
            Set cbrToolBar = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_Model + i, aPresetFilter(i - 1).strname)
            cbrToolBar.Category = i - 1
            
            If Not MenuPopup Is Nothing Then
                Set cbrMenuBar = MenuPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_Model + i, aPresetFilter(i - 1).strname)
                cbrMenuBar.Category = i - 1
            End If
        End If
    Next i
End Sub

Public Sub subSetWidthLevelF(im As DicomImage, f As Form, Optional cbrPopup As CommandBarPopup)
'------------------------------------------------
'���ܣ����ô���λ���ܼ������˵�
'������ im--���ô���λ�Ļ�׼ͼ��
'       f--���õ����˵��Ĵ��壻
'       cbrPopup -- Ϊ�������ù����������˵��Ĵ���λ�˵����Ϊ�գ���������Ҽ������˵���������cbrPopup�еĴ���λ�˵��
'���أ���
'2009��
'-----------------------------------------------
    Dim strDriverType As String, intDriverType As Integer
    Dim i As Integer, j As Integer
    Dim ControlPopup As CommandBarPopup
    Dim cbrToolBar As CommandBarControl
    Dim cbrToolBarF2 As CommandBarControl
    Dim MenuPopup As CommandBarPopup    '���˵��еĵ����˵���
    Dim cbrMenuBar As CommandBarControl '���˵��еĲ˵���
    Dim cbrMenuBarF2 As CommandBarControl   '���˵��еĲ˵���
    Dim blnIsMainViewer As Boolean          '�Ƿ������壬������������ǽ�Ƭ��ӡ����
    
    On Error GoTo err
    
    If im Is Nothing Then Exit Sub
    If IsNull(im.Attributes(&H8, &H60).Value) Then Exit Sub         '��ȡModality
    If f.Name = "frmFilm" Then    '��Ƭ��ӡ����
        blnIsMainViewer = False
    Else
        blnIsMainViewer = True
    End If
    
    strDriverType = im.Attributes(&H8, &H60).Value
    If cbrPopup Is Nothing Then
        If blnIsMainViewer = False Then   '��Ƭ��ӡ����
            '��չ������еĵ�������
            Set ControlPopup = f.CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
            ControlPopup.CommandBar.Controls.DeleteAll
        Else
            '��չ������еĵ����˵�����
            Set ControlPopup = f.ComToolBar.Item(ToolBar_Comm).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
            ControlPopup.CommandBar.Controls.DeleteAll
            
            '������˵��еĵ����˵�����
            Set MenuPopup = f.ComToolBar.Item(ToolBar_Menu).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True)
            MenuPopup.CommandBar.Controls.DeleteAll
        End If
    Else
        Set ControlPopup = cbrPopup
    End If
    
    intDriverType = 0
    
    For i = 1 To UBound(aPresetWinWL, 2)        '[�ҵ�ͼ���Ӧ�豸]
        If UCase(aPresetWinWL(3, i).strModality) = UCase(strDriverType) Then
            intDriverType = i
            Exit For
        End If
    Next
    '''''''''''''''''''''''''''''''[����F2�˵�]''''''''''''''''''''''''''''''''''''''''
    Set cbrToolBarF2 = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_ReSet, "F2 �Զ�")
    cbrToolBarF2.Checked = True
    If blnIsMainViewer Then
        If Not MenuPopup Is Nothing Then
            Set cbrMenuBarF2 = MenuPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_ReSet, "F2 �Զ�")
            cbrMenuBarF2.Checked = True
        End If
        f.ComToolBar.KeyBindings.Add 0, VK_F2, ID_Active_AdjustWindow_HandAdjustWindow_ReSet
    Else
        f.CommBar_Film.KeyBindings.Add 0, VK_F2, ID_Active_AdjustWindow_HandAdjustWindow_ReSet
    End If
    
    ''''''''''''''''''''''''''''''[�����Զ��尴ť]'''''''''''''''''''''''''''''''''''''''''
    ControlPopup.CommandBar.Controls.Add xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_Custom, "�Զ���"
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
    If intDriverType > 0 Then
        For j = 3 To 12
            If aPresetWinWL(j, intDriverType).bInUse Then
                '���Ӵ���λ��ť
                Set cbrToolBar = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3, aPresetWinWL(j, intDriverType).strWinWLCName)
                cbrToolBar.Category = aPresetWinWL(j, intDriverType).lngWinWidth & "-" & aPresetWinWL(j, intDriverType).lngWinLevel
                If blnIsMainViewer Then
                    f.ComToolBar.KeyBindings.Add 0, VK_F3 + (j - 3), ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3
                Else
                    f.CommBar_Film.KeyBindings.Add 0, VK_F3 + (j - 3), ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3
                End If
                
                '�������˵��ĵ����˵���
                If Not MenuPopup Is Nothing Then
                    Set cbrMenuBar = MenuPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow_F3 + j - 3, aPresetWinWL(j, intDriverType).strWinWLCName)
                    cbrMenuBar.Category = aPresetWinWL(j, intDriverType).lngWinWidth & "-" & aPresetWinWL(j, intDriverType).lngWinLevel
                End If
                '����Ĭ�ϰ�ť
                If aPresetWinWL(j, intDriverType).intDefault = 1 Then
                    cbrToolBarF2.Checked = False
                    cbrToolBar.Checked = True
                    
                    If Not MenuPopup Is Nothing Then
                        cbrMenuBarF2.Checked = False
                        cbrMenuBar.Checked = True
                    End If
                End If
            End If
        Next
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal X3 As Double, ByVal Y3 As Double) As Double
'------------------------------------------------
'���ܣ�����ͨ�����������ɵ�������֮��ĽǶ�
'��������X1,Y1��������ֱ�߽����X,Y���ꣻ��X2,Y2������ֱ��1�ϵ�һ�㣻��X3,Y3������ֱ��2�ϵ�һ��
'���أ�GetAngle��������ֱ��֮��ĽǶȣ���λΪ����
'2009��
'------------------------------------------------
    Dim Pi As Double
    Dim dblCos As Double, dblAngle1 As Double, dblAngle2 As Double
    Pi = 3.14159265358979
    If x1 = x2 And y1 = y2 Then
        dblAngle1 = 0
    Else
        dblCos = (x2 - x1) / Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
        If Abs(dblCos) = 1 Then
            dblAngle1 = IIf(dblCos = 1, 0, Pi)
        Else
            dblAngle1 = Atn(-dblCos / Sqr(-dblCos * dblCos + 1)) + 2 * Atn(1)
        End If
    End If
    If x1 = X3 And y1 = Y3 Then
        dblAngle2 = 0
    Else
        dblCos = (X3 - x1) / Sqr((X3 - x1) ^ 2 + (Y3 - y1) ^ 2)
        If Abs(dblCos) = 1 Then
            dblAngle2 = IIf(dblCos = 1, 0, Pi)
        Else
            dblAngle2 = Atn(-dblCos / Sqr(-dblCos * dblCos + 1)) + 2 * Atn(1)
        End If
    End If
    GetAngle = IIf((y2 - y1) * (Y3 - y1) > 0, Abs(dblAngle1 - dblAngle2), Abs(dblAngle1 + dblAngle2)) * 180 / Pi
    If GetAngle > 180 Then GetAngle = 360 - GetAngle
End Function

Function Max7InArray(a As Variant, ByRef lMax As Long, ByRef lMin As Long) As Long
'------------------------------------------------
'���ܣ�����������ȡֵ����7���±꣬������ƽ��ֵ
'������a--���в��������飻lMax--������������±�   lMin--����������С�±ꡣ
'���أ�����ֵΪ��ƽ��ֵ����ͨ����lMax��lMin������������С�±ꡣ
'2009��
'------------------------------------------------
    Dim m1 As Long, m2 As Long, m3 As Long, m4 As Long, m5 As Long, m6 As Long, m7 As Long
    Dim c1 As Long, c2 As Long, c3 As Long, c4 As Long, c5 As Long, c6 As Long, c7 As Long
    Dim s As Long
    Dim cMax As Long
    Dim cMin As Long
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cMax = 1
    cMin = 1
    Max7InArray = 0
    m1 = 0
    m2 = 0
    m3 = 0
    m4 = 0
    m5 = 0
    m6 = 0
    m7 = 0
    s = a(1)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim lCount As Long
    lCount = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    While lCount <= UBound(a)
        If a(lCount) > m7 Then
           If a(lCount) > m6 Then
               If a(lCount) > m5 Then
                    If a(lCount) > m4 Then
                       If a(lCount) > m3 Then
                           If a(lCount) > m2 Then
                               If a(lCount) > m1 Then
                                   m1 = a(lCount)
                                   cMax = lCount
                                   c1 = lCount
                               Else
                                   m2 = a(lCount)
                                   c2 = lCount
                               End If
                           Else
                               m3 = a(lCount)
                               c3 = lCount
                           End If
                       Else
                           m4 = a(lCount)
                           c4 = lCount
                       End If
                    Else
                       m5 = a(lCount)
                       c5 = lCount
                    End If
               Else
                    m6 = a(lCount)
                    c6 = lCount
               End If
           Else
                m7 = a(lCount)
                c7 = lCount
           End If
        End If
        ''''''�ж�Сֵ'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If a(lCount) < s Then
            s = a(lCount)
            cMin = lCount
        End If
        lCount = lCount + 1
    Wend
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Max7InArray = (c1 + c2 + c3 + c4 + c5 + c6 + c7) / 7
    lMax = cMax
    lMin = cMin
End Function

Private Function funIsShutter(strModality As String) As Integer
'------------------------------------------------
'���ܣ��ж������Ӱ������Ƿ���Ҫ����ͼ����������
'������
'     strModality ����ͼ��������Ӱ�����
'���أ�0-û�����ã����ô���1����������2����ͼ��������
'2009��
'------------------------------------------------
    Dim i As Integer
    funIsShutter = 0
    For i = 1 To UBound(aImageShutter)
        If UCase(aImageShutter(i).strModality) = UCase(strModality) Then
            If aImageShutter(i).intShutterType > 0 And aImageShutter(i).intShutterType < 8 Then
                funIsShutter = 2    '��ͼ������
            Else
                funIsShutter = 1    '��ͼ������
            End If
        End If
    Next i
End Function

Public Sub subDrawImgShutter(img As DicomImage, Optional isForce As Boolean = False)
'------------------------------------------------
'���ܣ�����ϵͳ���õ�Ӱ����𣬸������ͼ��ͼ������
'������
'       img ����ͼ��������ͼ��
'       isForce ����������Ϊ��ʹ��������ʱ���Ƿ�ǿ��ɾͼ���г����е�������Ϣ
'���أ���
'2009��
'------------------------------------------------
    Dim iResult As Integer
    Dim intModality As Integer
    Dim strModality As String
    Dim i As Integer
    Dim strArray() As String
    Dim intCount As Integer
    Dim intShutterType As Integer
    Dim strCenter(2) As String
    Dim strVertices() As String
    
    If IsNull(img.Attributes(&H8, &H60).Value) Then Exit Sub
    strModality = img.Attributes(&H8, &H60).Value
    iResult = funIsShutter(strModality)
    If isForce = False And iResult = 0 Then Exit Sub
    For i = 1 To UBound(aImageShutter)
        If UCase(aImageShutter(i).strModality) = UCase(strModality) Then
            intModality = i
            Exit For
        End If
    Next i
    
    '����ͼ������
    If aImageShutter(intModality).intShutterType > 0 And aImageShutter(intModality).intShutterType < 8 Then
        '������������
        intShutterType = aImageShutter(intModality).intShutterType
        intCount = 0
        If intShutterType >= 4 Then     '���������
            intShutterType = intShutterType - 4
            intCount = intCount + 1
            ReDim Preserve strArray(intCount) As String
            strArray(intCount) = "POLYGONAL"
            '���������㣬����ż�����򽫶���ζ�����ӵ�ͼ����
            strVertices = Split(aImageShutter(intModality).strVertices, ":")
            If UBound(strVertices) >= 5 And UBound(strVertices) Mod 2 = 1 Then
                ReDim Preserve strVertices(UBound(strVertices) + 1) As String
                For i = UBound(strVertices) To 1 Step -1
                    strVertices(i) = strVertices(i - 1)
                Next i
                img.Attributes.Add &H18, &H1620, strVertices
            End If
        End If
        If intShutterType >= 2 Then     '��������
            intShutterType = intShutterType - 2
            intCount = intCount + 1
            ReDim Preserve strArray(intCount) As String
            strArray(intCount) = "RECTANGULAR"
            img.Attributes.Add &H18, &H1602, aImageShutter(intModality).intRectLeft
            img.Attributes.Add &H18, &H1604, aImageShutter(intModality).intRectRight
            img.Attributes.Add &H18, &H1606, aImageShutter(intModality).intRectUpper
            img.Attributes.Add &H18, &H1608, aImageShutter(intModality).intRectLower
        End If
        If intShutterType >= 1 Then     'Բ������
            intCount = intCount + 1
            ReDim Preserve strArray(intCount) As String
            strArray(intCount) = "CIRCULAR"
            '���Բ�ĺͰ뾶
            strCenter(1) = aImageShutter(intModality).intCenterX
            strCenter(2) = aImageShutter(intModality).intCenterY
            img.Attributes.Add &H18, &H1610, strCenter
            img.Attributes.Add &H18, &H1612, aImageShutter(intModality).intRadius
        End If
        img.Attributes.Add &H18, &H1600, strArray
        img.Attributes.Add &H18, &H1622, aImageShutter(intModality).lngColor
    Else        '������������
        img.Attributes.Remove &H18, &H1600
        img.Attributes.Remove &H18, &H1602
        img.Attributes.Remove &H18, &H1604
        img.Attributes.Remove &H18, &H1606
        img.Attributes.Remove &H18, &H1608
        img.Attributes.Remove &H18, &H1610
        img.Attributes.Remove &H18, &H1612
        img.Attributes.Remove &H18, &H1620
        img.Attributes.Remove &H18, &H1622
    End If
    img.Refresh False
End Sub

Public Function funGetLinePoints(img As DicomImage, la As DicomLabel, aGrey() As Integer, intBeginX As Integer _
                , intBeginY As Integer, intEndX As Integer, intEndY As Integer) As Boolean
'------------------------------------------------
'���ܣ���ͼ��ĸ������ͱ�ע��ֱ�ߡ����ߣ�����ȡ�Ҷ�ֵ�������㡢�յ�����
'������
'       img ��������ע��ͼ��
'       la �� ��ȡ�Ҷ�ֵ�����ͱ�ע��
'       aGrey������Ҷ�ֵ�����飬����ֵ��
'       intBeginX������X���꣬����ֵ��
'       intBeginY ������Y���꣬����ֵ��
'       intEndX���յ��X���꣬����ֵ��
'       intEndY���յ��X���꣬����ֵ��
'���أ��Ƿ�ɹ������˻Ҷ�ֵ���顣True���������ء�Fasle��ִ��ʧ�ܣ������Ǳ�ע�������ͱ�ע��
'2009��
'------------------------------------------------
    '��ȡֱ���ϻҶ�ֵ����ŵ�������
    Dim vPixels As Variant
    Dim i As Integer
    Dim iFrame As Integer
    Dim lngCount As Long        '�������
    Dim iSizex As Integer       'ͼ���x�������
    Dim iSizey As Integer       'ͼ���y�������
    Dim iTempx As Integer       '��ǰ����ʱx��
    Dim iTempy As Integer       '��ǰ����ʱy��
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iFrame = img.Frame          '��¼��ǰͼ�������
    vPixels = img.Pixels        '��ȡ��ǰͼ������ص�
    iSizex = img.sizex          '��ȡ��ǰͼ���x�������
    iSizey = img.sizey          '��ȡ��ǰͼ���y�������
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If la.width = 0 And la.height = 0 Then  '��ȡ��߶ȶ�Ϊ�㣬�����ÿ��Ϊ1
        intBeginX = la.left
        intEndX = la.left + 1
        intBeginY = la.top
        intEndY = la.top
        lngCount = intEndX - intBeginX + 1
        ReDim aGrey(lngCount) As Integer
        aGrey(lngCount - 1) = vPixels(intBeginX, intBeginY, iFrame)
        aGrey(lngCount) = vPixels(intEndX, intBeginY, iFrame)
        funGetLinePoints = True
        Exit Function
    End If
    '�ֳ�����������Ҷ�����
    If la.LabelType = doLabelLine Then      ' ����ֱ�ߵĲ���
        Dim lngW As Long
        Dim lngH As Long
        Dim iCount As Long
        iCount = 1
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        lngW = Abs(la.width) + 1
        lngH = Abs(la.height) + 1
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If lngW > lngH Then             '��ȴ��ڸ߶ȣ�����x��������������Ҽ���ֱ���ϵĵ�
            If la.width < 0 Then        'left���ұߣ���Ҫ����˳��
                intEndX = la.left
                intEndY = la.top
                intBeginX = la.left + la.width
                intBeginY = la.top + la.height
            Else                        'left����ߣ�begin��ֱ��ȡleft,top��
                intBeginX = la.left
                intBeginY = la.top
                intEndX = la.left + la.width
                intEndY = la.top + la.height
            End If
    
            'ȷ��intBeginX��intEndX��ֵ��1��ͼ���sizex֮��
'            If intBeginX < 1 Then intBeginX = 1
'            If intBeginX > iSizex Then intBeginX = iSizex
'            If intEndX < 1 Then intEndX = 1
'            If intEndX > iSizex Then intEndX = iSizex
    
            lngCount = intEndX - intBeginX + 1
            ReDim aGrey(lngCount) As Integer
    
            For i = intBeginX To intEndX
                iTempx = i
                iTempy = la.height / la.width * (i - intBeginX) + intBeginY
                'ȷ��iTempx��ֵ��1��ͼ���sizex֮��
                If iTempx < 1 Then iTempx = 1
                If iTempx > iSizex Then iTempx = iSizex
                'ȷ��iTempy��ֵ��1��ͼ���sizey֮��
                If iTempy < 1 Then iTempy = 1
                If iTempy > iSizey Then iTempy = iSizey
                aGrey(iCount) = vPixels(iTempx, iTempy, iFrame)
                iCount = iCount + 1
            Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Else                            '�߶ȴ��ڿ�ȣ�����y������������ϵ��¼���ֱ���ϵĵ�
            If la.height < 0 Then       'top���±ߣ���Ҫ����˳��
                intEndX = la.left
                intEndY = la.top
                intBeginX = la.left + la.width
                intBeginY = la.top + la.height
            Else                        'top���ϱߣ�begin��ֱ��ȡleft,top��
                intBeginX = la.left
                intBeginY = la.top
                intEndX = la.left + la.width
                intEndY = la.top + la.height
            End If
    
            'ȷ��intBeginY��intEndY��ֵ��1��ͼ���sizey֮��
            If intBeginY < 1 Then intBeginY = 1
            If intBeginY > iSizey Then intBeginY = iSizey
            If intEndY < 1 Then intEndY = 1
            If intEndY > iSizey Then intEndY = iSizey
    
            lngCount = intEndY - intBeginY + 1
            ReDim aGrey(lngCount) As Integer
            For i = intBeginY To intEndY
                iTempx = la.width / la.height * (i - intBeginY) + intBeginX
                iTempy = i
                'ȷ��iTempx��ֵ��1��ͼ���sizex֮��
                If iTempx < 1 Then iTempx = 1
                If iTempx > iSizex Then iTempx = iSizex
                'ȷ��iTempy��ֵ��1��ͼ���sizey֮��
                If iTempy < 1 Then iTempy = 1
                If iTempy > iSizey Then iTempy = iSizey
                aGrey(iCount) = vPixels(iTempx, iTempy, iFrame)
                iCount = iCount + 1
            Next
        End If
        funGetLinePoints = True
    Else                            '���ڶ���ߣ�ֱ�Ӷ�Points����
        Dim vPoints As Variant

        vPoints = la.Points
        lngCount = UBound(vPoints) / 2
        ReDim aGrey(lngCount) As Integer
        For i = 1 To lngCount
            iTempx = vPoints(2 * i - 1)
            iTempy = vPoints(2 * i)
            
            'ȷ��iTempx��ֵ��1��ͼ���sizex֮��
            If iTempx < 1 Then iTempx = 1
            If iTempx > iSizex Then iTempx = iSizex
            
            'ȷ��iTempy��ֵ��1��ͼ���sizey֮��
            If iTempy < 1 Then iTempy = 1
            If iTempy > iSizey Then iTempy = iSizey
            
            aGrey(i) = vPixels(iTempx, iTempy, iFrame)
        Next
        intBeginX = vPoints(1)
        intBeginY = vPoints(2)
        intEndX = vPoints(lngCount * 2 - 1)
        intEndY = vPoints(lngCount * 2)
        funGetLinePoints = True
    End If
End Function

Public Function funGetVasEdge(img As DicomImage, lblLine As DicomLabel, intThreshold As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean
'------------------------------------------------
'���ܣ�����Ѫ����խ����������ֱ�߱�ע��Ԥ�����ֵ������Ѫ�ܱڵ����ꡣ
'������
'       img ��������ע��ͼ��
'       lblLine �� ��ֱ��Ѫ�ܵ�ֱ�߱�ע��
'       (x1,y1)�����Ѫ�ܱڸ�Ѫ�ܴ�ֱ�߽�������꣬����ֵ��
'       (x2,y2)���ұ�Ѫ�ܱڸ�Ѫ�ܴ�ֱ�߽�������꣬����ֵ��
'���أ��Ƿ�ɹ������������Ѫ�ܱ����ꡣTrue���������ء�Fasle��ִ��ʧ�ܣ����������ֱ�߱�ע�������ͱ�ע��
'2009��
'------------------------------------------------
    Dim aGrey() As Integer
    Dim intBeginX As Integer, intBeginY As Integer, intEndX As Integer, intEndY As Integer
    Dim lngCount As Long
    Dim i As Integer
    Dim lngCenter As Long
    Dim intLower As Integer
    Dim intUpper As Integer
    
    If lblLine.LabelType <> doLabelLine Then Exit Function
    'If Abs(lblLine.width) < 2 And Abs(lblLine.height) < 2 Then Exit Function
    If funGetLinePoints(img, lblLine, aGrey, intBeginX, intBeginY, intEndX, intEndY) = False Then Exit Function
    lngCount = UBound(aGrey)
    lngCenter = lngCount \ 2
    '��ͼ�����Ͻ���Ѫ�ܱ�
    intLower = 1        '��ʼ�����Ͻ�Ѫ�ܱ�
    For i = lngCenter To 1 Step -1
        If Abs(aGrey(i) - aGrey(lngCenter)) > intThreshold Then
            intLower = i
            Exit For
        End If
    Next i
    '��ͼ�����½���Ѫ�ܱ�
    intUpper = lngCount     '��ʼ�����½�Ѫ�ܱ�
    For i = lngCenter + 1 To lngCount Step 1
        If Abs(aGrey(i) - aGrey(lngCenter)) > intThreshold Then
            intUpper = i
            Exit For
        End If
    Next i
    '�ж�ֱ�ߵ�б�ʣ��ǰ���X�����㣬���ǰ���Y������
    If lngCount = intEndY - intBeginY + 1 Then '����Y������
        y1 = intBeginY + intLower
        y2 = intBeginY + intUpper
        x1 = (y1 - intEndY) / (intBeginY - intEndY) * (intBeginX - intEndX) + intEndX
        x2 = (y2 - intEndY) / (intBeginY - intEndY) * (intBeginX - intEndX) + intEndX
    Else        '����X������
        x1 = intBeginX + intLower
        x2 = intBeginX + intUpper
        y1 = (x1 - intEndX) / (intBeginX - intEndX) * (intBeginY - intEndY) + intEndY
        y2 = (x2 - intEndX) / (intBeginX - intEndX) * (intBeginY - intEndY) + intEndY
    End If
    funGetVasEdge = True
End Function

Public Sub subDrawVasEdgeLine(lblLine As DicomLabel, lblShortLine As DicomLabel, intCenterX As Long, intCenterY As Long)
'------------------------------------------------
'���ܣ�����Ѫ����խ����������ֱ�߱�ע��Ѫ�ܱڵ����꣬ȷ��������Ѫ�ܱڶ�ֱ�ߵ�λ�á�
'������
'       lblLine �� ��ֱ��Ѫ�ܵ�ֱ�߱�ע��
'       lblShortLine �� Ѫ�ܱڶ�ֱ�ߡ�
'       (intCenterX,intCenterY)��Ѫ�ܱڸ�Ѫ�ܴ�ֱ�߽�������ꡣ
'���أ��ޣ�ֱ���ƶ�Ѫ�ܱڶ�ֱ�ߵ�λ�á�
'2009��
'------------------------------------------------
    Dim lngLineWidth As Long
    Dim intNewX As Integer
    Dim intNewY As Integer
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    lngWidth = Abs(lblLine.width)
    lngHeight = Abs(lblLine.height)
    lngLineWidth = Sqr(lngHeight * lngHeight + lngWidth * lngWidth)
    If lngLineWidth = 0 Then
        intNewX = 0
        intNewY = 0
    Else
        intNewX = lngHeight / lngLineWidth * intVasEdgeWidth / 2
        intNewY = lngWidth / lngLineWidth * intVasEdgeWidth / 2
    End If
    
    If (lblLine.width > 0 And lblLine.height > 0) Or (lblLine.width < 0 And lblLine.height < 0) Then
        intNewY = -intNewY
    End If
    lblShortLine.left = intCenterX - intNewX
    lblShortLine.top = intCenterY - intNewY
    lblShortLine.height = intNewY * 2
    lblShortLine.width = intNewX * 2
End Sub

Public Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double)
'------------------------------------------------
'���ܣ���ͼ��������š��Ե�ǰviewer���ĵ�Ϊ�������ĵ㡣
'������
'       img -- �������ŵ�ͼ��
'       viewer ���� ͼ�����ڵ�viewer
'       dblZoom ����ͼ���µ����ű���
'���أ��ޣ�ֱ�ӵ���ͼ������ű���
'2009��
'------------------------------------------------
    Dim dblOldZoom As Double
    Dim lngOldScroX As Long
    Dim lngOldScroY As Long
    Dim dblZoomRatio As Double
    
    On Error GoTo err
    
    If img Is Nothing Then Exit Sub
    If img.ActualZoom = 0 Then Exit Sub
    
    dblOldZoom = img.ActualZoom
    lngOldScroX = img.ActualScrollX
    lngOldScroY = img.ActualScrollY
    img.Zoom = dblZoom
    img.StretchToFit = False
    
    dblZoomRatio = 1 - img.ActualZoom / dblOldZoom
    img.ScrollX = lngOldScroX - (lngOldScroX + Viewer.width / Viewer.MultiColumns / Screen.TwipsPerPixelX / 2) * dblZoomRatio
    img.ScrollY = lngOldScroY - (lngOldScroY + Viewer.height / Viewer.MultiRows / Screen.TwipsPerPixelY / 2) * dblZoomRatio
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub RectangleZoom(FilmViewer As DicomViewer, img As DicomImage, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
'------------------------------------------------
'���ܣ���ͼ��������š�����Viewer��Ĵ�С����
'������
'       img -- �������ŵ�ͼ��
'       viewer ���� ͼ�����ڵ�viewer
'       dblZoom ����ͼ���µ����ű���
'���أ��ޣ�ֱ�ӵ���ͼ������ű���
'2009��
'------------------------------------------------
    Dim newZoom As Double
    Dim dblRatio As Double
    Dim oldZoom As Double
    
    If lngWidth > 0 And lngHeight > 0 Then
        oldZoom = img.ActualZoom
        img.StretchToFit = False

        dblRatio = FilmViewer.width / FilmViewer.MultiColumns / Screen.TwipsPerPixelX / lngWidth
        If dblRatio > FilmViewer.height / FilmViewer.MultiRows / Screen.TwipsPerPixelY / lngHeight Then
            dblRatio = FilmViewer.height / FilmViewer.MultiRows / Screen.TwipsPerPixelY / lngHeight
        End If
        
        newZoom = oldZoom * dblRatio
        img.Zoom = newZoom
        
        img.ScrollX = lngLeft * dblRatio
        img.ScrollY = lngTop * dblRatio
    End If
End Sub

Public Function CutOutAImage(img As DicomImage)
    Dim Simg As New DicomImage
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lblFrame As DicomLabel
    Dim lngTemp As Long
    
    Set CutOutAImage = Simg
    On Error GoTo err
    
    Set lblFrame = img.Labels(img.Labels.Count)
    
    'ͼ����λ��
    If lblFrame.width >= 0 Then
        lngLeft = lblFrame.left
    Else
        lngLeft = lblFrame.left + lblFrame.width
    End If
    lngWidth = Abs(lblFrame.width)
    
    If lblFrame.height >= 0 Then
        lngTop = lblFrame.top
    Else
        lngTop = lblFrame.top + lblFrame.height
    End If
    lngHeight = Abs(lblFrame.height)
    
    lblFrame.Visible = False
    
    '�ü�ͼ��
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) Then
        '����,����+ȫ��
        lngLeft = img.sizex - lngLeft - lngWidth
        
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) Then
        '���� ,����+ȫ��
        lngTop = img.sizey - lngTop - lngHeight
        
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    ElseIf (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
         Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
         '180�� �� ȫ����
        lngLeft = img.sizex - lngLeft - lngWidth
        lngTop = img.sizey - lngTop - lngHeight
    ElseIf (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) Then
        '���Ҿ���,180+���µ���
        lngLeft = img.sizex - lngLeft - lngWidth
    ElseIf (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) Then
        '���µ���,180+����
        lngTop = img.sizey - lngTop - lngHeight
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) Then
        '����+���Ҿ������������Ƿ���
        '����+���µ��ã����������Ƿ���
        lngTop = img.sizey - lngTop - lngHeight
        lngLeft = img.sizex - lngLeft - lngWidth
        
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipVertical) Then
        '����+���Ҿ��� ���������Ƿ���
        '����+���µ��ã� ���������Ƿ���
        lngTemp = lngLeft
        lngLeft = lngTop
        lngTop = lngTemp
    End If
    Set Simg = img.SubImage(lngLeft, lngTop, lngWidth, lngHeight, 1, 1)
    
    Set CutOutAImage = Simg
    Exit Function
err:

End Function


Public Sub subWriteDicomPara(imgSource As DicomImage, imgDest As DicomImage)
'------------------------------------------------
'���ܣ��������ͼ����дDICOM�ļ�ͷ��Ϣ
'������img���������DICOM�ļ�,lngAdviceID����ҽ��ID
'���أ��ޣ�ֱ���ļ�ͷ��Ϣд��img���ļ�ͷ
'------------------------------------------------
    Dim curDate As Date
    Dim attr As DicomAttribute
    Dim Dicomglb As New DicomGlobal

    curDate = zlDatabase.Currentdate
    
    imgDest.InstanceUID = Dicomglb.NewUID
    imgDest.StudyUID = imgSource.StudyUID
    imgDest.SeriesUID = imgSource.SeriesUID
    
    imgDest.Attributes.Add &H8, &H8, ""                             'ImageType  ��
    imgDest.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"   'SOP Class  UID�����β�׽
    imgDest.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date �������
    imgDest.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date ��������
    imgDest.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date �ɼ�����
    imgDest.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   ͼ������
    imgDest.Attributes.Add &H8, &H30, Format(curDate, "HH24:MI:SS")     'Study Time   ���ʱ��
    imgDest.Attributes.Add &H8, &H31, Format(curDate, "HH24:MI:SS")     'Series Time  ����ʱ��
    imgDest.Attributes.Add &H8, &H32, Format(curDate, "HH24:MI:SS")     'Acquisition Time  �ɼ�ʱ��
    imgDest.Attributes.Add &H8, &H33, Format(curDate, "HH24:MI:SS")     'Image Time  ͼ��ʱ��
    imgDest.Attributes.Add &H8, &H50, ""                            'Accession Number ��
    imgDest.Attributes.Add &H8, &H60, imgSource.Attributes(&H8, &H60).Value                  'Modality Ӱ�����
    imgDest.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer ����
    imgDest.Attributes.Add &H8, &H80, "ZLSOFT"                'Institution Name ��λ����
    imgDest.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name ��
'    imgDest.Attributes.Add &H8, &H1030, ""                          'Study Description ������� ��
    imgDest.Attributes.Add &H10, &H10, imgSource.Name                       'Name ����
    imgDest.Attributes.Add &H10, &H20, imgSource.PatientID                 'Patient ID ����ID
    imgDest.Attributes.Add &H10, &H30, imgSource.DateOfBirth                  'BirthDate ����
    imgDest.Attributes.Add &H10, &H40, imgSource.Sex                        'Sex �Ա�
    imgDest.Attributes.Add &H10, &H4000, ""                         'Patient Comment ����ע��
    imgDest.Attributes.Add &H20, &H10, "1"                   'Study ID ���ID
    imgDest.Attributes.Add &H20, &H11, "1"                          'Series Number ���к�
    imgDest.Attributes.Add &H20, &H13, "1"                          'ImageNumber ͼ���
    imgDest.Attributes.Add &H20, &H20, ""                           'Orientation ��
    
'    ��ӱ����Ϣ
    If imgSource.Attributes(&H28, &H30).Exists Then
        imgDest.Attributes.Add &H28, &H30, imgSource.Attributes(&H28, &H30).Value
    End If
    'KODAK CR800 ʹ�����µı����Ϣ
    If imgSource.Attributes(&H18, &H1164).Exists Then
        imgDest.Attributes.Add &H18, &H1164, imgSource.Attributes(&H18, &H1164).Value
    End If
End Sub

Public Function funCopyMPRControlLines(im As DicomImage, oldImage As DicomImage)
'------------------------------------------------
'���ܣ���ʼ��ָ��ͼ���ʸ��״�ؽ����Ƶ�Ϳ�����
'������     im--����ʸ��״�ؽ�����λͼ��
'           oldImage -- ��Ҫ���������ߵ�ԭͼ�������ǿ�
'���أ��ޣ�ֱ�Ӷ�imͼ���ϵ�ʸ��״�ؽ���ע����ʼ����
'------------------------------------------------
    On Error GoTo err
    
    If oldImage Is Nothing Then
        Exit Function
    End If
    
    im.Labels(G_INT_SYS_LABEL_MPRV).left = oldImage.Labels(G_INT_SYS_LABEL_MPRV).left
    im.Labels(G_INT_SYS_LABEL_MPRV).top = oldImage.Labels(G_INT_SYS_LABEL_MPRV).top
    im.Labels(G_INT_SYS_LABEL_MPRV).height = oldImage.Labels(G_INT_SYS_LABEL_MPRV).height
    im.Labels(G_INT_SYS_LABEL_MPRV).width = oldImage.Labels(G_INT_SYS_LABEL_MPRV).width
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPRH).left = oldImage.Labels(G_INT_SYS_LABEL_MPRH).left
    im.Labels(G_INT_SYS_LABEL_MPRH).top = oldImage.Labels(G_INT_SYS_LABEL_MPRH).top
    im.Labels(G_INT_SYS_LABEL_MPRH).height = oldImage.Labels(G_INT_SYS_LABEL_MPRH).height
    im.Labels(G_INT_SYS_LABEL_MPRH).width = oldImage.Labels(G_INT_SYS_LABEL_MPRH).width
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left
    im.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = oldImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Refresh False
     
    funCopyMPRControlLines = True
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function funImageSmoothing(ByRef img() As Integer, intTimes As Integer) As Boolean
'------------------------------------------------
'���ܣ��Զ�ά�����е�ͼ����ƽ������
'������ img() -- ͼ���ά����
'       intTimes -- ƽ��������һ����1-2��
'���أ��ޣ������ؽ����ͼ��
'------------------------------------------------
    Dim i As Long, j As Long, k As Long
    Dim TolHeight As Long       'ͼ��߶�
    Dim lngWidth As Long
    Dim lngTemp As Long          'Ϊ�˷�ֹinteger���͵�a,b����������ƽ�������е������ʹ��long�͵��м����
    
    On Error GoTo err
    
    TolHeight = UBound(img, 2)
    lngWidth = UBound(img, 1)
    
    For k = 1 To intTimes
        For i = 3 To TolHeight - 2
            For j = 2 To lngWidth - 1
                '���ģ��ܺã���ϸ�ڴ��������
                lngTemp = CLng(img(j - 1, i - 1)) + CLng(img(j - 1, i + 1)) + _
                          CLng(2 * img(j, i - 2)) + CLng(2 * img(j, i - 1)) + CLng(img(j, i))
                lngTemp = (CLng(lngTemp) + CLng(2 * img(j, i + 1)) + CLng(2 * img(j, i + 2)) _
                         + CLng(img(j + 1, i - 1)) + CLng(img(j + 1, i + 1)) _
                         ) / 13
                img(j, i) = lngTemp
            Next
        Next
        '''''������������������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 2 To lngWidth - 1
            lngTemp = CLng(img(j - 1, 1)) + CLng(img(j - 1, 2)) + CLng(2 * img(j - 1, 3)) + CLng(img(j, 1)) + CLng(img(j, 2))
            lngTemp = (CLng(lngTemp) + CLng(2 * img(j, 3)) + CLng(img(j + 1, 1)) + CLng(img(j + 1, 2)) + CLng(2 * img(j + 1, 3))) / 12
            img(j, 2) = lngTemp
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 1 To lngWidth - 1
            lngTemp = CLng(img(j - 1, 1)) + CLng(img(j - 1, 2)) + CLng(img(j - 1, 3)) + CLng(img(j, 1))
            lngTemp = (CLng(lngTemp) + CLng(img(j, 2)) + CLng(img(j, 3)) + CLng(img(j + 1, 1)) + CLng(img(j + 1, 2)) + CLng(img(j + 1, 3))) / 9
            img(j, 1) = lngTemp
        Next
        ''''''������������������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 2 To lngWidth - 1
            lngTemp = CLng(2 * img(j - 1, TolHeight - 2)) + CLng(img(j - 1, TolHeight - 1)) + CLng(img(j - 1, TolHeight)) + CLng(2 * img(j, TolHeight - 2))
            lngTemp = (CLng(lngTemp) + CLng(img(j, TolHeight - 1)) + CLng(img(j, TolHeight)) + CLng(2 * img(j + 1, TolHeight - 2)) + CLng(img(j + 1, TolHeight - 1)) + CLng(img(j + 1, TolHeight))) / 12
            img(j, TolHeight - 1) = lngTemp
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j = 1 To lngWidth - 1
            lngTemp = CLng(img(j - 1, TolHeight)) + CLng(img(j - 1, TolHeight - 1)) + CLng(img(j - 1, TolHeight - 2)) + CLng(img(j, TolHeight)) + CLng(img(j, TolHeight - 1))
            lngTemp = (CLng(lngTemp) + CLng(img(j, TolHeight - 2)) + CLng(img(j + 1, TolHeight)) + CLng(img(j + 1, TolHeight - 1)) + CLng(img(j + 1, TolHeight - 2))) / 9
            img(j, TolHeight) = lngTemp
        Next
   Next
   
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

