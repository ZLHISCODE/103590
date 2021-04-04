Attribute VB_Name = "mdlLabel"
Option Explicit
'--------------------------------------------------------
'��  �ܣ���ģ��Ϊ��ע����ĺ������̵�
'�����ˣ�����
'�������ڣ�2005.06.12
'���̺����嵥��
'    funAccountMiddlePoint():   ֱ����������͵��������X��Y���꣬�������Ӧ��Y��X���ꡣ
'    subPeriodMovee5X():        ʸ��״���ĵ���ƶ�
'    subPeriodMove():           ʸ��״���Ƶ����ĸ��ߵ���ƶ�����
'    funROIResultString():      ���ɲ�������ַ���,����ϵͳ���������õ��Ƿ���ʾ�����ƽ��ֵ�������������
'    subMove25():               �ü�����ϵͳ��ע1���ƶ���ע2-5��
'    subTakeOut1():             ��ѡ�еı�ע��ȥ��ϵͳ��ע���ڲü������Ϊ2-5�ı�ע���������1�ı�ע�Ƿ�ɾ����ͨ������isTakeOut1���ж�
'    funLabelType():            �����ڲ�ʹ�õı�ע��������
'    SubDispPeriod():           Ϊָ��ͼ���е�ָ����ע����ʾ��עѡ����
'    SubDispLinePeriod():       Ϊͼ���е�ָ��ֱ�ߺͼ�ͷ��ע����ʾ��עѡ����
'    subCutOutInphase():        �ڲü�״̬�¶Բü���������ͼ��ͬ������
'    funMouseOverPeriod():      ���������Խ���ľ�����
'    subMoveMPRLabel():         �ƶ�ʸ��״�ؽ����Ƶ㡢�ߣ��������µ��ؽ�ͼ��
'    subMoveLable():            �ƶ�һ����ע,����ʸ��״�ؽ���ע���û���ע�Ͳü���ע
'    subChangeLableSize():      �ı�һ����ע�Ĵ�С,���޸�����ز�����Ϣ����ʾֵ
'    SubNoDispPeriod():         Ϊָ��ͼ�����ر�עѡ����
'    subTextCoordinate():       ����ͼ��ķ�ת����������ֵ�����ת��
'    SubChangeColor():          �ı�ѡ��LABEL����ɫ
'    GetNewLabel():             ����һ��LABEL���󣬲���������ʼ����
'    subDeleteAppointLabel():   ɾ��ָ�����͵ı�ע
'    SubInitPeriod():           Ϊÿһ��ͼ����ǰn��ϵͳ�����n�������ɳ���G_INT_SYS_LABEL_COUNT����
'    UpdateMarkers():           ����ͼ����ʾ�����ز�����λ��Ϣ
'    UpdateRuler():             ��ʾͼ����
'    subDispImageInfo():        ��ʾ�����ز���ͼ���Ľ���Ϣ�ʹ���λ��ʾ
'    subGetImgInfoLabel():      ��ͼ������ȡ���˵��ĸ�����Ϣ��ע�����ϵͳ�����������ĸ��Ǳ�ע������ʹ��
'    subInitImageLabels():      ��ʼ������ʾ������ָ��ͼ��ı�ע��Ϣ:ϵͳ��ע����λ��ע����ߣ��Ľ���Ϣ������λ
'    funcCalImgInfoLabel():     ���ݴ�������ļ�ƣ��������Ӧ�ĽǱ�ע����ʾֵ��
'    subSaveLabelToImg():       ����ע���浽DICOMͼ���ͷ��Ϣ����
'    subReadLabelFromImg():     ��ͼ���ͷ�ļ��ж�ȡ��ע������ʾ��ע
'    funDrawVas():              ����lblLine���Զ�Ѫ�ܲ���
'�޸ļ�¼��
'    2005.06.30    �ƽ�      �����Ż�
'-------------------------------------------------------

Private Function funAccountMiddlePoint(x1, y1, x2, y2, X3)
'------------------------------------------------
'���ܣ�ֱ����������͵��������X��Y���꣬�������Ӧ��Y��X���ꡣ
'������(X1,Y1)--ֱ���ϵ�һ��������ꣻ��X2��Y2��--ֱ���ϵڶ���������ꣻX3--ֱ���ϵ��������X����
'���أ����������Y��X����
'2009��
'------------------------------------------------
    funAccountMiddlePoint = 0
    If x1 = x2 Then Exit Function
    If y1 = y2 Then
        funAccountMiddlePoint = y1
        Exit Function
    End If
    funAccountMiddlePoint = (X3 - x2) * (y2 - y1) / (x2 - x1) + y2
    If funAccountMiddlePoint > 30000 Then funAccountMiddlePoint = 30000
    If funAccountMiddlePoint < -30000 Then funAccountMiddlePoint = -30000
End Function

Public Sub subPeriodMovee5X(la, lb, ll, xx, Yy, E5, basex, baseY, im As DicomImage)
'------------------------------------------------
'���ܣ�ʸ��״���ĵ���ƶ�
'������la--�����Ŀ��Ƶ���ͬһֱ���ϵı߿��Ƶ㣻lb--��߿��Ƶ�la��ͬһֱ�ߵ���һ���߿��Ƶ㣻
'      ll--la��lb�����ӵĿ����ߣ�xx--���Ŀ��Ƶ����ĵ����Xλ�ã�Yy--���Ŀ��Ƶ����ĵ����Yλ�ã�
'      E5--���Ŀ��Ƶ㣻basex--���Ŀ��Ƶ����ĵ�ľ�Xλ�ã�baseY--���Ŀ��Ƶ����ĵ�ľ�Yλ�ã�
'      im--�����������Ƶ��ߵ�ͼ��
'���أ��ޡ�ֱ���ƶ���la��lb��ll�������Ƶ����
'2009��
'------------------------------------------------
    Dim x, y
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If ll.height = 0 Then
        la.top = E5.top
        lb.top = E5.top
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf ll.width = 0 Then
        la.left = E5.left
        lb.left = E5.left
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.top = -G_INT_MPR_RADIUS / 2 Then   '''''���la�ڶ���
        x = funAccountMiddlePoint(la.top + Yy - baseY, la.left, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
        x = x + xx - basex
        la.left = x
        If la.left + G_INT_MPR_RADIUS / 2 < 0 Then   '''la�������
            la.top = funAccountMiddlePoint(la.left, -G_INT_MPR_RADIUS / 2, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
            la.left = -G_INT_MPR_RADIUS / 2
        ElseIf la.left + G_INT_MPR_RADIUS / 2 > im.sizex Then   '''la�����ұ�
            la.top = funAccountMiddlePoint(la.left, -G_INT_MPR_RADIUS / 2, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            la.left = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.top = im.sizey - G_INT_MPR_RADIUS / 2 Then  ''���la�ڵ���
        x = funAccountMiddlePoint(la.top + Yy - baseY, la.left, E5.top, E5.left, im.sizey - G_INT_MPR_RADIUS / 2)
        x = x + xx - basex
        la.left = x
        If la.left + G_INT_MPR_RADIUS / 2 < 0 Then
            la.top = funAccountMiddlePoint(la.left, im.sizey - G_INT_MPR_RADIUS / 2, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
            la.left = -G_INT_MPR_RADIUS / 2
        ElseIf la.left + G_INT_MPR_RADIUS / 2 > im.sizex Then
            la.top = funAccountMiddlePoint(la.left, im.sizey - G_INT_MPR_RADIUS / 2, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            la.left = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.left = -G_INT_MPR_RADIUS / 2 Then    ''���la�����
        y = funAccountMiddlePoint(la.left + xx - basex, la.top, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
        y = y + Yy - baseY
        la.top = y
        If la.top + G_INT_MPR_RADIUS / 2 < 0 Then   '''la�����ϱ�
            la.left = funAccountMiddlePoint(la.top, -G_INT_MPR_RADIUS / 2, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
            la.top = -G_INT_MPR_RADIUS / 2
        ElseIf la.top + G_INT_MPR_RADIUS / 2 > im.sizex Then  '''la�����±�
            la.left = funAccountMiddlePoint(im.sizex - G_INT_MPR_RADIUS / 2, la.left, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
            la.top = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.left = im.sizey - G_INT_MPR_RADIUS / 2 Then   ''���la���ұ�
        y = funAccountMiddlePoint(la.left + xx - basex, la.top, E5.left, E5.top, im.sizey - G_INT_MPR_RADIUS / 2)
        y = y + Yy - baseY
        la.top = y
        If la.top + G_INT_MPR_RADIUS / 2 < 0 Then   '''la�����ϱ�
            la.left = funAccountMiddlePoint(la.top, -G_INT_MPR_RADIUS / 2, E5.top, E5.left, im.sizex - G_INT_MPR_RADIUS / 2)
            la.top = -G_INT_MPR_RADIUS / 2
        ElseIf la.top + G_INT_MPR_RADIUS / 2 > im.sizex Then  '''la�����±�
            la.left = funAccountMiddlePoint(im.sizex - G_INT_MPR_RADIUS / 2, la.left, E5.top, E5.left, im.sizex - G_INT_MPR_RADIUS / 2)
            la.top = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    End If
    ''''''����Ե�'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If la.top < E5.top Then
        lb.left = funAccountMiddlePoint(la.top, la.left, E5.top, E5.left, im.sizey - G_INT_MPR_RADIUS / 2)
        lb.top = im.sizey - G_INT_MPR_RADIUS / 2
        If lb.left + G_INT_MPR_RADIUS / 2 < 0 Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, G_INT_MPR_RADIUS / 2)
            lb.left = -G_INT_MPR_RADIUS / 2
        ElseIf lb.left + G_INT_MPR_RADIUS / 2 > im.sizex Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            lb.left = im.sizex - G_INT_MPR_RADIUS / 2
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf la.top > E5.top Then
        lb.left = funAccountMiddlePoint(la.top, la.left, E5.top, E5.left, -G_INT_MPR_RADIUS / 2)
        lb.top = -G_INT_MPR_RADIUS / 2
        If lb.left + G_INT_MPR_RADIUS / 2 < 0 Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, -G_INT_MPR_RADIUS / 2)
            lb.left = -G_INT_MPR_RADIUS / 2
        ElseIf lb.left + G_INT_MPR_RADIUS / 2 > im.sizex Then
            lb.top = funAccountMiddlePoint(lb.left, lb.top, E5.left, E5.top, im.sizex - G_INT_MPR_RADIUS / 2)
            lb.left = im.sizex - G_INT_MPR_RADIUS / 2 - G_INT_MPR_RADIUS / 2
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ll.left = la.left + G_INT_MPR_RADIUS / 2
    ll.top = la.top + G_INT_MPR_RADIUS / 2
    ll.width = lb.left - la.left
    ll.height = lb.top - la.top
End Sub


Public Sub subPeriodMove(ByVal la As DicomLabel, ByVal x As Long, ByVal y As Long, ByVal lb As DicomLabel, _
                  ByVal ll As DicomLabel, E5 As DicomLabel, im As DicomImage)
'------------------------------------------------
'���ܣ�ʸ��״���Ƶ����ĸ��ߵ���ƶ�����
'������ la--���ƶ��Ŀ��Ʊߵ��ע ��x--��ע��λ����ͼ���ϵ�X����  ��y--��ע��λ����ͼ���ϵ�Y����  ��
'       lb--�����ƶ���ע��ͬһֱ���ϵ���һ�����Ʊߵ㣻ll--��la������ʸ��״�����ߣ�
'       E5--ʸ��״���Ƶ��е����ĵ㣻im--��ʸ��״�ؽ���ͼ��
'���أ��ޣ�ֱ���ƶ���ע���ı���la,lb,ll��λ��
'2009��
'------------------------------------------------
    Dim x1 As Long, y1 As Long
    Dim x2 As Long, y2 As Long  '��X2,Y2����¼ʸ��״�ؽ����Ŀ��Ƶ��ע����������
    Dim X3 As Long, Y3 As Long, movex As Long
    x2 = E5.left + G_INT_MPR_RADIUS / 2    '''���ĵ�λ��
    y2 = E5.top + G_INT_MPR_RADIUS / 2
    '''''''''''''''''�����λ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If x = x2 Then             ''''���Xλ�ú����ĵ�ƽ��
        X3 = x
        Y3 = IIf(y > y2, 0, 0 + im.sizey)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y = y2 Then           ''''���Yλ�ú����ĵ�ƽ��
        Y3 = y
        X3 = IIf(x > x2, 0, 0 + im.sizex)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y > y2 Then              ''''���λ�����ĵ��Ϸ�
        Y3 = 0 + im.sizey
        X3 = (Y3 - y2) * (x2 - x) / (y2 - y) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y < y2 Then              '''''���λ�����ĵ��·�
        Y3 = 0
        X3 = (Y3 - y2) * (x2 - x) / (y2 - y) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�������һ����ߵ�λ��
    la.left = X3 - G_INT_MPR_RADIUS / 2
    la.top = Y3 - G_INT_MPR_RADIUS / 2
    '''''''''''''''''''''''''''''''''''''''''''''''''''����Ե�λ��
    x1 = X3
    y1 = Y3
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If x1 = x2 Then
        X3 = x1
        Y3 = IIf(y1 > y2, 0, 0 + im.sizey)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y1 = y2 Then
        Y3 = y1
        X3 = IIf(x1 > x2, 0, 0 + im.sizex)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y1 < y2 Then
        Y3 = 0 + im.sizey
        X3 = (Y3 - y2) * (x2 - x1) / (y2 - y1) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf y1 > y2 Then
        Y3 = 0
        X3 = (Y3 - y2) * (x2 - x1) / (y2 - y1) + x2
        If X3 < 0 Then
            Y3 = (y2 - Y3) * (0 - X3) / (x2 - X3) + Y3
            X3 = 0
        ElseIf X3 > 0 + im.sizex Then
            Y3 = (y2 - Y3) * (0 + im.sizex - X3) / (x2 - X3) + Y3
            X3 = 0 + im.sizex
        End If

    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lb.left = X3 - G_INT_MPR_RADIUS / 2
    lb.top = Y3 - G_INT_MPR_RADIUS / 2
'    '''''''''''''''''''''''''''''''''''''''''''''
    ll.left = la.left + G_INT_MPR_RADIUS / 2
    ll.top = la.top + G_INT_MPR_RADIUS / 2
    ll.width = X3 - ll.left
    ll.height = Y3 - ll.top
End Sub

Public Function funROIResultString(la As DicomLabel, img As DicomImage) As String
'------------------------------------------------
'���ܣ����ɲ���������Ļ�Ӣ���ַ���,����ϵͳ���������õ��Ƿ���ʾ�����ƽ��ֵ���������������Ϊ����ı�ע������������
'������la--Ϊ��Ҫ��������ı�ע�������ڲ����ݲ�ͬ�ı�ע���ͷ��ز�ͬ�Ľ��������ֱ�ߣ�����߲��������ز����ĳ���
'���أ�Ϊ��������ַ���
'2009��
'------------------------------------------------
    funROIResultString = ""
    Dim strROIArea As String
    Dim strROIMean As String
    Dim strROIStdDev As String
    Dim strROILength As String
    Dim strROIMax As String
    Dim strROIMin As String
    Dim strAngle As String
    Dim lTemp As DicomLabel
    If bROITextChinese Then        ''ʹ�����ı�ʾ������Ϣ
        strROIArea = "�����"
        strROIMean = "ƽ��ֵ��"
        strROIMax = "���ֵ��"
        strROIMin = "��Сֵ��"
        If img.Attributes(&H8, &H60).Exists And Not IsNull(img.Attributes(&H8, &H60).Value) Then
            If UCase(img.Attributes(&H8, &H60).Value) = "CT" Then
                strROIMean = "ƽ��CTֵ��"
                strROIMax = "���CTֵ��"
                strROIMin = "��СCTֵ��"
            End If
        End If
        
        strROIStdDev = "��׼�"
        strROILength = "�ܳ���"
        strAngle = "�Ƕȣ�"
    Else
        strROIArea = "Area: "
        strROIMean = "Mean: "
        strROIStdDev = "Std.Dev: "
        strROILength = "Length:"
        strAngle = "Angle:"
        strROIMax = "Max: "
        strROIMin = "Min:"
    End If
    
    '���δ���,��Ҫ�Ƿ�ֹα��ͼ����ִ���
    On Error Resume Next
    
    If left(la.Tag, 2) = "JD" Then
        '����Ƕ�
        If bROIArea Then
            Set lTemp = la
            If lTemp.Tag = "JD1" Then
                If Not lTemp.TagObject.TagObject Is Nothing Then
                    Set lTemp = lTemp.TagObject.TagObject
                End If
            End If
            funROIResultString = strAngle & Int(GetAngle(lTemp.left, lTemp.top, lTemp.left + lTemp.width, lTemp.top + lTemp.height, lTemp.TagObject.left, lTemp.TagObject.top) * 100) / 100
        End If
    ElseIf la.LabelType = doLabelLine Or la.LabelType = doLabelPolyLine Then
        If bROIArea Then funROIResultString = Int(la.ROILength) & la.ROIDistanceUnits
    Else
        If la.LabelType = doLabelEllipse Or la.LabelType = doLabelPolygon Or _
          la.LabelType = doLabelRectangle Then
                If bROIArea Then funROIResultString = strROIArea & Int(la.ROIArea) & la.ROIDistanceUnits
                If bROIMean Then
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIMean
                        If blnSelectedImageIfColor = True Then
                            funROIResultString = funROIResultString & "0"
                        Else
                            funROIResultString = funROIResultString & Int(la.ROIMean)
                        End If
                    Else
                        funROIResultString = strROIMean & Int(la.ROIMean)
                    End If
                End If
                If bROIStandardDeviation Then       '��ʾ��׼��
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIStdDev
                        If blnSelectedImageIfColor = True Then
                            funROIResultString = funROIResultString & "0"
                        Else
                            funROIResultString = funROIResultString & Int(la.ROIStandardDeviation)
                        End If
                    Else
                        funROIResultString = strROIStdDev & Int(la.ROIStandardDeviation)
                    End If
                End If
                If bROILength Then          '��ʾ�ܳ�
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROILength & Int(la.ROILength)
                    Else
                        funROIResultString = strROILength & Int(la.ROILength)
                    End If
                End If
                If bROIMax Then             '��ʾ���ֵ
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIMax & Int(la.ROIMax)
                    Else
                        funROIResultString = strROIMax & Int(la.ROIMax)
                    End If
                End If
                If bROIMin Then             '��ʾ��Сֵ
                    If funROIResultString <> "" Then
                        funROIResultString = funROIResultString & vbCrLf & strROIMin & Int(la.ROIMin)
                    Else
                        funROIResultString = strROIMin & Int(la.ROIMin)
                    End If
                End If
        End If
    End If
End Function

Public Sub subMove25(im As DicomImage, f As frmViewer)
'------------------------------------------------
'���ܣ��ü�����ϵͳ��ע1���ƶ���ע2-5������2����ߣ�3-�±ߣ�4-�ұߣ�5-�ϱߡ�
'������im--��Ҫ�ƶ�ϵͳ��ע��ͼ��f--��Ҫ�ƶ�ϵͳ��ע�Ĵ���
'���أ��ޣ�ֱ���ƶ���ע2,3,4,5��λ��
'------------------------------------------------
    Dim i As Integer
    For i = 2 To 5
        '�����ĸ��ڵ����εĿ�Ⱥ͸߶�
        im.Labels(i).height = 32766 \ IIf(i Mod 2 = 0, 1, 2)
        im.Labels(i).width = 32766 \ IIf(i Mod 2 = 0, 2, 1)
    Next
    If im.Labels(1).width > 0 Then
        im.Labels(2).left = im.Labels(1).left - im.Labels(2).width
        im.Labels(4).left = im.Labels(1).left + im.Labels(1).width
    Else
        im.Labels(2).left = im.Labels(1).left + im.Labels(1).width - im.Labels(4).width
        im.Labels(4).left = im.Labels(1).left
    End If
    
    If im.Labels(1).height > 0 Then
        im.Labels(3).top = im.Labels(1).top + im.Labels(1).height
        im.Labels(5).top = im.Labels(1).top - im.Labels(5).height
    Else
        im.Labels(3).top = im.Labels(1).top
        im.Labels(5).top = im.Labels(1).top + im.Labels(1).height - im.Labels(5).height
    End If
    im.Labels(2).top = im.Labels(5).top
    im.Labels(4).top = im.Labels(5).top
    im.Labels(3).left = im.Labels(2).left
    im.Labels(5).left = im.Labels(2).left
End Sub

Public Sub subTakeOut1(ls As DicomLabels, im As DicomImage, isTakeOut1 As Boolean)
'------------------------------------------------
'���ܣ���ѡ�еı�ע��ȥ��ϵͳ��ע���ڲü������Ϊ2-5�ı�ע���������1�ı�ע�Ƿ�ɾ����ͨ������isTakeOut1���жϡ�
'������ls--��Ҫ����ɾ���ı�ע����im--��ע���ڵ�ͼ��isTakeOut1--�Ƿ�ɾ�����Ϊ1�ı�ע��Trueɾ����Fasle��ɾ����
'���أ��ޣ�ֱ�Ӵ����ע��ls�����ݡ�
'------------------------------------------------
    Dim l As DicomLabel
    For Each l In ls
        If im.Labels.IndexOf(l) < 6 Then
            If isTakeOut1 Or im.Labels.IndexOf(l) <> 1 Then ls.Remove (ls.IndexOf(l))
        End If
    Next
End Sub

Public Function funLabelType(la As DicomLabel) As String
'------------------------------------------------
'���ܣ������ڲ�ʹ�õı�ע��������
'������la--��Ҫ�жϱ�ע���͵ı�ע��
'���أ���ע���͵���������
'�ϼ���������̣�frmLabelObject.load
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ�����
'------------------------------------------------
    funLabelType = ""
    Select Case la.LabelType
        Case 0:
            funLabelType = "����"
        Case 1:
            funLabelType = "��Բ"
        Case 2:
            funLabelType = "����"
        Case 3:
            If la.Tag = "JD1" Then
                funLabelType = "�Ƕ���(1)"
            ElseIf la.Tag = "JD2" Then
                funLabelType = "�Ƕ���(2)"
            ElseIf la.Tag = "JDT" Then
                funLabelType = "�Ƕ�����"
            ElseIf la.Tag = "RLL" Then
                funLabelType = "��λ��"
            Else
                funLabelType = "ֱ��"
            End If
        Case 4:
            funLabelType = "�����"
        Case 5:
            funLabelType = "�����"
        Case 6:
            funLabelType = "ͼ��"
        Case 7:
            funLabelType = "��λ��ע"
        Case 8:
             funLabelType = "Բ��"
        Case 9:
             funLabelType = "�ڲ�ֵ�㷨�Ķ����"
        Case 10:
             funLabelType = "��ͷ"
        Case 11:
             funLabelType = "���"
    End Select
End Function

Public Sub SubDispPeriod(la As DicomLabel, im As DicomImage, f As frmViewer)
'------------------------------------------------
'���ܣ�Ϊָ��ͼ���е�ָ����ע����ʾ��עѡ����
'������la--����עѡ������Χ�ı�ע��im--��ʾ��עѡ������ͼ��f--��ʾ��עѡ�����Ĵ��塣
'���أ��ޣ�ֱ����ʾ��עѡ������
'2009��
'------------------------------------------------
    Dim intZoom As Double
    Dim i As Integer        '��ѭ���õ���ʱ����
    Dim img As DicomImage
    Dim lblTemp As DicomLabel
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''����������ʾ��С
    If im.ActualZoom <> 0 Then
        For i = 11 To 20
            im.Labels(i).height = IIf(intPeriodSize / im.ActualZoom >= 1, intPeriodSize / im.ActualZoom, 1)
            im.Labels(i).width = im.Labels(i).height
        Next
    End If
    SubNoDispPeriod im, f               ''''Ϊָ��ͼ�����ر�עѡ����
    Set im.Labels(11).TagObject = la    ''''''������1�ž��ָ��ǰ��ע''����һ����Ҫ��,�Ժ�ܶ�ط�Ҫ�õ��˵�ļ�¼
    
    If la.LabelType = doLabelLine Or la.LabelType = doLabelArrow Then
     ''�ߺͼ�ͷ���ͱ�ע�Ĵ�������ֱ�ߡ���ͷ�������Ρ�Ѫ����խ�����رȲ�����
     ''ʹ�õ�ѡ�������Ϊ:ֱ�ߺͼ�ͷ��11��15��,�Ƕȣ�11,15,18��
     ''Ѫ����խ������11,12,13,14,15,16,17��18)�����رȣ�11,14,15,18��
        
        If la.Tag = "VAS1L" Or la.Tag = "VAS2L" Then    '��Ѫ����խ������ע���д���
            Set lblTemp = la
            Set im.Labels(11).TagObject = lblTemp    '��1�ž��ָ��ֱ��
            '����ֱ�߼�ѡ����
            SubDispLinePeriod lblTemp, im, 11, 14
            '����ֱ�߼�ѡ����
            For i = 1 To 4
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                End If
            Next i
            If Right(lblTemp.Tag, 1) = "L" Then
                SubDispLinePeriod lblTemp, im, 15, 18
            Else    '�������Ϊ15,18��ѡ����
                
            End If
        ElseIf la.Tag = "CTR1L" Or la.Tag = "CTR2L" Then  '�����رȲ����ı�ע���д���
            Set lblTemp = la
            Set im.Labels(11).TagObject = lblTemp   '��1�ž��ָ��ֱ��
            Call SubDispLinePeriod(lblTemp, im, 11, 14)
            If Not lblTemp.TagObject Is Nothing Then
                Set lblTemp = lblTemp.TagObject
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                End If
            End If
            If Right(lblTemp.Tag, 1) = "L" Then
                Call SubDispLinePeriod(lblTemp, im, 15, 18)
            End If
        Else
            SubDispLinePeriod la, im, 11, 15
        End If
        
        '''''''''�ǶȵĴ���'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Mid(la.Tag, 1, 3) = "JD1" Or Mid(la.Tag, 1, 3) = "JD2" Then
            Dim laTagObject As New DicomLabel
            If Mid(la.Tag, 1, 3) = "JD1" Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set laTagObject = la.TagObject.TagObject
                im.Labels(18).left = (laTagObject.left + laTagObject.width)
                im.Labels(18).top = (laTagObject.top + laTagObject.height)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If laTagObject.width > 0 And laTagObject.height > 0 Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width > 0 And laTagObject.height < 0 Then
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height < 0 Then
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height > 0 Then
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width = 0 Then  ''��������
                    If laTagObject.height > 0 Then
                    Else
                        im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                    End If
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width / 2
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.height = 0 Then   ''�������
                    If laTagObject.width > 0 Then
                    Else
                        im.Labels(18).left = im.Labels(18).left - im.Labels(18).height
                    End If
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).width / 2
                End If
            Else
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Set laTagObject = la.TagObject
                im.Labels(18).left = laTagObject.left
                im.Labels(18).top = laTagObject.top
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If laTagObject.width > 0 And laTagObject.height > 0 Then
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width > 0 And laTagObject.height < 0 Then
                    im.Labels(18).left = laTagObject.left - im.Labels(18).width
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height < 0 Then
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width < 0 And laTagObject.height > 0 Then
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.width = 0 Then  ''��������
                    If laTagObject.height > 0 Then
                        im.Labels(18).top = im.Labels(18).top - im.Labels(18).height
                    Else
                    End If
                    im.Labels(18).left = im.Labels(18).left - im.Labels(18).width / 2
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ElseIf laTagObject.height = 0 Then   ''�������
                    If laTagObject.width > 0 Then
                        im.Labels(18).left = im.Labels(18).left - im.Labels(18).height
                    Else
                    End If
                    im.Labels(18).top = im.Labels(18).top - im.Labels(18).width / 2
                End If
            End If
        End If          '�Ƕȴ������
        
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 11 To 18
          im.Labels(i).Visible = True
        Next
    ElseIf la.LabelType = doLabelEllipse Or la.LabelType = doLabelRectangle Then
    '''''''''''''''''''''''''''''''''''' ''''���κ���Բ''''''''''''''''''''''''''''''''''''''''''''''''''''''
        im.Labels(11).left = la.left
        im.Labels(11).top = la.top
        im.Labels(12).top = (la.top + (la.height - im.Labels(11).height) / 2)
        im.Labels(13).top = (la.top + la.height)
        im.Labels(14).left = (la.left + (la.width - im.Labels(11).height) / 2)
        im.Labels(15).left = (la.left + la.width)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If la.width > 0 And la.height > 0 Then
            im.Labels(11).left = im.Labels(11).left - im.Labels(11).width
            im.Labels(11).top = im.Labels(11).top - im.Labels(11).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width > 0 And la.height < 0 Then
            im.Labels(11).left = la.left - im.Labels(11).width
            im.Labels(13).top = im.Labels(13).top - im.Labels(11).width
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height < 0 Then
            im.Labels(13).left = im.Labels(13).left - im.Labels(11).width
            im.Labels(13).top = im.Labels(13).top - im.Labels(11).width
            im.Labels(15).left = im.Labels(15).left - im.Labels(11).width
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height > 0 Then
            im.Labels(15).left = im.Labels(15).left - im.Labels(11).width
            im.Labels(13).left = im.Labels(13).left - im.Labels(11).width
            im.Labels(11).top = im.Labels(11).top - im.Labels(11).height
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        im.Labels(12).left = im.Labels(11).left
        im.Labels(13).left = im.Labels(11).left
        im.Labels(14).top = im.Labels(13).top
        im.Labels(15).top = im.Labels(13).top
        im.Labels(16).left = im.Labels(15).left
        im.Labels(16).top = im.Labels(12).top
        im.Labels(17).left = im.Labels(15).left
        im.Labels(17).top = im.Labels(11).top
        im.Labels(18).left = im.Labels(14).left
        im.Labels(18).top = im.Labels(11).top
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 11 To 18
          im.Labels(i).Visible = True
        Next
    ElseIf la.LabelType = 4 Or la.LabelType = 5 Then
        '''''''''''''''''''''''''''''''''''' ''''����κͶ����'''''''''''''''''''''''''''''''''''''''''''''''''''''
        la.SelectMode = 4
    End If
End Sub

Private Sub SubDispLinePeriod(la As DicomLabel, im As DicomImage, intEnd1 As Integer, intEnd2 As Integer)
'------------------------------------------------
'���ܣ�Ϊͼ���е�ָ��ֱ�ߺͼ�ͷ��ע����ʾ��עѡ����
'������la--����עѡ������Χ�ı�ע��im--��ʾ��עѡ������ͼ��intEnd1-��һ�������ţ�intEnd2-�ڶ���������
'���أ��ޣ�ֱ����ʾ��עѡ������
'2009��
'------------------------------------------------
    If la.LabelType = doLabelLine Or la.LabelType = doLabelArrow Then
        im.Labels(intEnd1).left = la.left
        im.Labels(intEnd1).top = la.top
        im.Labels(intEnd2).left = (la.left + la.width)
        im.Labels(intEnd2).top = (la.top + la.height)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If la.width > 0 And la.height > 0 Then
            im.Labels(intEnd1).left = im.Labels(intEnd1).left - im.Labels(intEnd1).width
            im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width > 0 And la.height < 0 Then
            im.Labels(intEnd1).left = la.left - im.Labels(intEnd1).width
            im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height < 0 Then
            im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).width
            im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).height
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width < 0 And la.height > 0 Then
            im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).height
            im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).width
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.width = 0 Then  ''��������
            If la.height > 0 Then
                im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).height
            Else
                im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).height
            End If
            im.Labels(intEnd1).left = im.Labels(intEnd1).left - im.Labels(intEnd1).width / 2
            im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).width / 2
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ElseIf la.height = 0 Then   ''�������
            If la.width > 0 Then
                im.Labels(intEnd1).left = im.Labels(intEnd1).left - im.Labels(intEnd1).height
            Else
                im.Labels(intEnd2).left = im.Labels(intEnd2).left - im.Labels(intEnd2).height
            End If
            im.Labels(intEnd1).top = im.Labels(intEnd1).top - im.Labels(intEnd1).width / 2
            im.Labels(intEnd2).top = im.Labels(intEnd2).top - im.Labels(intEnd2).width / 2
        End If
    End If
End Sub

Public Sub subCutOutInphase(v As DicomViewer, im As DicomImage, f As frmViewer)
'------------------------------------------------
'���ܣ��ڲü�״̬�¶Բü���������ͼ��ͬ������
'������v--����ͼ��ͬ����viewer��im--��Ϊͬ�����յ�ͼ��f--����ͬ���Ĵ���
'���أ��ޣ�ֱ�Ӹı�ü���ע��λ�úʹ�С
'------------------------------------------------
    Dim img As DicomImage
    Dim i As Integer
    For Each img In v.Images
        For i = 1 To 5
            img.Labels(i).Visible = im.Labels(i).Visible
            img.Labels(i).left = im.Labels(i).left
            img.Labels(i).top = im.Labels(i).top
            img.Labels(i).width = im.Labels(i).width
            img.Labels(i).height = im.Labels(i).height
        Next
        SubNoDispPeriod img, f          'Ϊָ��ͼ�����ر�עѡ����
        If img.Labels(1).Visible Then SubDispPeriod img.Labels(1), img, f   'Ϊָ��ͼ���е�ָ����ע����ʾ��עѡ����
    Next
    v.Refresh
End Sub

Public Function funMouseOverPeriod(v As DicomViewer, im As DicomImage, ByVal x As Long, ByVal y As Long) As Integer
'------------------------------------------------
'���ܣ����������Խ���ľ�����
'������v--������ڵ�viewer��im--������ڵ�ͼ��x--����Xλ�ã�y--����Yλ�á�
'���أ�0--��겻�ھ���ϣ�11��18-����ڸ����������ľ���ϡ�
'2009��
'------------------------------------------------
    Dim xx As Long, Yy As Long
    xx = v.ImageXPosition(x, y)
    Yy = v.ImageYPosition(x, y)
    funMouseOverPeriod = 0
    Dim i As Integer
    With im
        For i = 11 To 18
            If .Labels(i).Visible And .Labels(i).left <= xx And .Labels(i).top <= Yy And .Labels(i).top + .Labels(i).height >= Yy And .Labels(i).left + .Labels(i).width >= xx Then
                funMouseOverPeriod = i
                Exit For
            End If
        Next
    End With
End Function

Private Sub subMoveMPRLabel(f As frmViewer, la As DicomLabel, xx As Integer, Yy As Integer, basex As Long, baseY As Long)
'------------------------------------------------
'���ܣ��ƶ�ʸ��״�ؽ����Ƶ㡢�ߣ��������µ��ؽ�ͼ��
'������ f--����ʸ��״�ؽ��Ĵ��壻
'       la--���ƶ���ʸ��״�ؽ����Ƶ������ߣ�
'       xx --��ע��λ����ͼ���ϵ�X���ꣻ
'       yy --��ע��λ����ͼ���ϵ�Y���ꣻ
'       basex--��λ�õ�ͼ������x���ꣻ
'       baseY--��λ�õ�ͼ������y���ꡣ
'���أ��ޣ�ֱ���ƶ�ʸ��״�ؽ��Ŀ��Ƶ���ߣ��������ؽ����ͼ��
'2009��
'------------------------------------------------
    Dim intIndex As Integer
    
    On Error GoTo err
    
    ''''''''''''''''''''''[���Ľǵ���ƶ�]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    intIndex = f.SelectedImage.Labels.IndexOf(la)
    'ʸ��״���Ƶ����ĸ��ߵ���ƶ�����
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 Then
        Call subPeriodMove(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), xx, Yy, _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), f.SelectedImage)
    '''''''''''''''''''''''''''���ĵ���ƶ�'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_O Then
        f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top + Yy - baseY
        f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left + xx - basex
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top < -G_INT_MPR_RADIUS / 2 + 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left < -G_INT_MPR_RADIUS / 2 + 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top > f.SelectedImage.sizey - G_INT_MPR_RADIUS / 2 - 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = f.SelectedImage.sizey - G_INT_MPR_RADIUS - 1
        End If
        If f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left > f.SelectedImage.sizex - G_INT_MPR_RADIUS / 2 - 1 Then
            f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = f.SelectedImage.sizex - G_INT_MPR_RADIUS - 1
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'ʸ��״���ĵ���ƶ�
        If xx <> basex Then
            Call subPeriodMovee5X(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), xx, Yy, _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), basex, baseY, f.SelectedImage)
        End If
        
        If Yy <> baseY Then
            Call subPeriodMovee5X(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), xx, Yy, _
                                f.SelectedImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), basex, baseY, f.SelectedImage)
        End If
        
    End If
    
    ''''''''''�����ؽ�''''''''''''''''''''''''''''''''''''''''''''''
    '��ע��MPR���������ߵ������˵㣬������MPR���������ߵ����ĵ㣬��ʱ��Ҫ�ƶ�����MPR����
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And xx <> basex) Then
        If funGetMPRImageAndShow(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), f, _
                                    f.Viewer(ZLMPRCube(1).intViewerIndex), f.SelectedImageIndex, _
                                    ZLMPRCube(2).intViewerIndex, ToltalHeight, 1, False, True) = False Then
            Call funMPR(f, True)
            Exit Sub
        End If
    End If
    
    '��ע��MPR�����ߺ��ߵ������˵㣬������MPR�����ߵĺ������ĵ㣬��ʱ��Ҫ�ƶ�����MPR����
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And Yy <> baseY) Then
        If funGetMPRImageAndShow(f.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), f, _
                                    f.Viewer(ZLMPRCube(1).intViewerIndex), f.SelectedImageIndex, _
                                    ZLMPRCube(3).intViewerIndex, ToltalHeight, 2, False, True) = False Then
            Call funMPR(f, True)
            Exit Sub
        End If
    End If
    '''''''''ʸ��״����ͬ��'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subMPRLinenPhase f.Viewer(f.intSelectedSerial), f.SelectedImage
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subMoveMPRReslutLabel(f As frmViewer, la As DicomLabel, x As Long, y As Long)
'------------------------------------------------
'���ܣ��ƶ�ʸ��״�ؽ�����ߣ����߿�����λͼ����Զ���ҳ�����߿��ƽ��ͼ
'������ f--����ʸ��״�ؽ��Ĵ��壻
'       la--���ƶ���ʸ��״�ؽ����Ƶ������ߣ�
'       x--x�����ƶ���ͼ�����ؾ��룻
'       y--y�����ƶ���ͼ�����ؾ��룻
'���أ��ޣ�ֱ���ƶ�ʸ��״�ؽ��Ľ���ߣ�����λͼ��ҳ��
'------------------------------------------------
    Dim iImageIndex As Integer
    Dim OldIntSelectedSeries As Integer
    Dim OldSelectedImage As DicomImage
    Dim oldSelectedImageIndex As Integer
    Dim intIndex As Integer
    Dim lngNewPosLeft As Long   '��λ��LEFT
    Dim lngNewPosTop As Long    '��λ��TOP
    Dim dblH As Double
    Dim dblW As Double
    Dim iViewerIndex As Integer
    Dim OldLeft As Long
    Dim dblXieBian As Double
    Dim dblSin As Double
    Dim dblCos As Double
    Dim dblDistance As Double
    
    On Error GoTo err
    
    intIndex = f.SelectedImage.Labels.IndexOf(la)

    
    If intIndex = G_INT_SYS_LABEL_MPR_RESULT_H Then '���ߣ��Ƿ�ҳ
        '���ƶ�ʸ��״�ؽ������
        la.top = la.top + y
        
        'ȷ������߲����뿪ͼ��
        If la.top < 0 Then
            la.top = 0
        ElseIf la.top > f.SelectedImage.sizey Then
            la.top = f.SelectedImage.sizey
        End If
        
        '���ݽ���ߵ�λ�ã���ҳ
        iImageIndex = la.top / f.SelectedImage.sizey * f.Viewer(ZLMPRCube(1).intViewerIndex).Images.Count
        If iImageIndex > 0 And iImageIndex <= f.Viewer(ZLMPRCube(1).intViewerIndex).Images.Count Then
            OldIntSelectedSeries = f.intSelectedSerial
            Set OldSelectedImage = f.SelectedImage
            
            f.VScro(ZLMPRCube(1).intViewerIndex).Value = iImageIndex
            
            'ԭͼ��ҳ�󣬻���������ı䣬��Ҫ�ָ�
            f.intSelectedSerial = OldIntSelectedSeries
            Set f.SelectedImage = OldSelectedImage
        End If
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_RESULT_V Then '���ߣ����ؽ�ͼ��
        OldLeft = la.left
        '���ƶ�ʸ��״�ؽ������
        la.left = la.left + x
        
        'ȷ������߲����뿪ͼ��
        If la.left < 0 Then
            la.left = 0
        ElseIf la.left > f.SelectedImage.sizex Then
            la.left = f.SelectedImage.sizex
        End If
        
        '���ݽ���ߵ�λ�ã��ƶ���λͼ�Ķ�Ӧ������
        
        '������λ��
        iViewerIndex = ZLMPRCube(1).intViewerIndex
        iImageIndex = f.VScro(iViewerIndex).Value
        
        '�ȳ�����ƽ��֮���ٳ���10������ƽ��ʱ�����
        '���ҵ���ǰ��ͼ���ǵڶ���ͼ�����ǵ�����ͼ
        If f.intSelectedSerial = ZLMPRCube(2).intViewerIndex Then
            '�ڶ���ͼ���ƶ���λͼ�е�����
            dblH = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRV).height / 10
            dblW = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRV).width / 10
        ElseIf f.intSelectedSerial = ZLMPRCube(3).intViewerIndex Then
            '������ͼ���ƶ���λͼ�е�����
            dblH = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRH).height / 10
            dblW = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPRH).width / 10
        End If
        
        dblXieBian = Sqr((dblW * dblW) + (dblH * dblH)) * 10
        dblSin = dblH * 10 / dblXieBian
        dblCos = dblW * 10 / dblXieBian
        dblDistance = (la.left - OldLeft) / f.SelectedImage.sizex * dblXieBian
        lngNewPosLeft = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left + dblDistance * dblCos
        lngNewPosTop = f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top + dblDistance * dblSin
            
        '��¼ͼ�����������ǰͼ�����ó���λͼ��
        OldIntSelectedSeries = f.intSelectedSerial
        oldSelectedImageIndex = f.SelectedImageIndex
        Set OldSelectedImage = f.SelectedImage
        
        f.intSelectedSerial = iViewerIndex
        Set f.SelectedImage = f.Viewer(iViewerIndex).Images(iImageIndex)
        f.SelectedImageIndex = iImageIndex
        
        Call subMoveMPRLabel(f, f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O), _
            CInt(lngNewPosLeft), CInt(lngNewPosTop), _
            f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).left, _
            f.Viewer(iViewerIndex).Images(iImageIndex).Labels(G_INT_SYS_LABEL_MPR_POINT_O).top)
            
        'ԭͼ�󣬻���������ı䣬��Ҫ�ָ�
        f.intSelectedSerial = OldIntSelectedSeries
        f.SelectedImageIndex = oldSelectedImageIndex
        Set f.SelectedImage = OldSelectedImage
        
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subMPRChanegImage(thisForm As frmViewer)
'MPRʸ��״λ�ؽ�ʱ���ı��˵�ǰѡ�е�ͼ��
'������ thisForm --- MPR���ڵĴ���
    
    On Error GoTo err
    
    If thisForm.SelectedImage Is Nothing Then Exit Sub
    If thisForm.intSelectedSerial <> ZLMPRCube(1).intViewerIndex Then Exit Sub
    
    '������ʾMPR������-���߶�Ӧ�Ľ��ͼ
    If funGetMPRImageAndShow(thisForm.SelectedImage.Labels(G_INT_SYS_LABEL_MPRV), thisForm, _
                                thisForm.Viewer(ZLMPRCube(1).intViewerIndex), thisForm.SelectedImageIndex, _
                                ZLMPRCube(2).intViewerIndex, ToltalHeight, 1, False, False) = False Then
        Call funMPR(thisForm, True)
        Exit Sub
    End If
    
    '������ʾMPR������-���߶�Ӧ�Ľ��ͼ
    If funGetMPRImageAndShow(thisForm.SelectedImage.Labels(G_INT_SYS_LABEL_MPRH), thisForm, _
                                thisForm.Viewer(ZLMPRCube(1).intViewerIndex), thisForm.SelectedImageIndex, _
                                ZLMPRCube(3).intViewerIndex, ToltalHeight, 2, False, False) = False Then
        Call funMPR(thisForm, True)
        Exit Sub
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subMoveLable(la As DicomLabel, x As Long, y As Long, f As frmViewer, xxx As Long, yyy As Long, basex As Long, baseY As Long)
'------------------------------------------------
'���ܣ��ƶ�һ����ע,����ʸ��״�ؽ���ע�����رȲ�����ע���û���ע�Ͳü���ע
'������la--���ƶ��ı�ע��x--x�����ƶ���ͼ�����ؾ��룻y--y�����ƶ���ͼ�����ؾ��룻f--�ƶ���ע�Ĵ��壻
'      xxx--��λ�õ���Ļ����x���ꣻyyy--��λ�õ���Ļ����y���ꣻbasex--��λ�õ�ͼ������x���ꣻ
'      baseY--��λ�õ�ͼ������y���ꡣ
'���أ���
'2009��
'------------------------------------------------
    Dim aa As Variant
    Dim lat As DicomLabel
    Dim i As Integer
    Dim pyX, pyY As Integer
    Dim lblTemp As DicomLabel
    
    
    If f.SelectedImage.Labels.IndexOf(la) >= G_INT_SYS_LABEL_MPRV And f.SelectedImage.Labels.IndexOf(la) <= G_INT_SYS_LABEL_MPR_POINT_O Then ''[ʸ��״�ߵ��ƶ�]
        '�ƶ�ʸ��״�ؽ����Ƶ㡢�ߣ��������µ��ؽ�ͼ��
        Dim xx As Integer
        Dim Yy As Integer
        
        xx = f.Viewer(f.intSelectedSerial).ImageXPosition(xxx, yyy)
        Yy = f.Viewer(f.intSelectedSerial).ImageYPosition(xxx, yyy)
        
        Call subMoveMPRLabel(f, la, xx, Yy, basex, baseY)
    ElseIf (f.SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_H) _
        Or (f.SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_V) Then
        '�ƶ�ʸ��״�ؽ������
        Call subMoveMPRReslutLabel(f, la, x, y)
    Else                                '�û���ע���ƶ�
        la.left = la.left + x
        la.top = la.top + y
        
        If la.LabelType = 4 Or la.LabelType = 5 Then    '�������κͶ����
            aa = la.Points
             For i = 1 To UBound(aa) Step 2
                 aa(i) = aa(i) + x
                 aa(i + 1) = aa(i + 1) + y
             Next
            la.Points = aa
            If la.LabelType = doLabelPolygon And Not la.TagObject Is Nothing Then la.TagObject.Text = funROIResultString(la, f.SelectedImage)
        End If
        ''''''''''''''''''���ڽǶ��ߵĴ���'''
        If Mid(la.Tag, 1, 2) = "JD" And Mid(la.Tag, 1, 3) <> "JDT" Then
            la.TagObject.left = la.TagObject.left + x
            la.TagObject.top = la.TagObject.top + y
            la.TagObject.TagObject.left = la.TagObject.TagObject.left + x
            la.TagObject.TagObject.top = la.TagObject.TagObject.top + y
            If Mid(la.Tag, 1, 3) = "JD1" Then
                la.TagObject.AnchorX = la.left '
                la.TagObject.AnchorY = la.top '
            Else
                la.TagObject.TagObject.AnchorX = la.TagObject.left '
                la.TagObject.TagObject.AnchorY = la.TagObject.top '
            End If
        ElseIf left(la.Tag, 3) = "VAS" Then         '����Ѫ����խ����"
            Dim iVasCount As Integer        '��¼Ѫ����խ������ע��������ȫ����ע��8����ֻ��������ע��4��
            '�Ȼָ�la��λ��
            iVasCount = 8
            If la.Tag = la.TagObject.TagObject.TagObject.TagObject.Tag Then iVasCount = 4
            la.left = la.left - x
            la.top = la.top - y
            '�ƶ�ʣ�µ�7����ע
            Set lblTemp = la
            If lblTemp.Tag = "VAS1L" Or lblTemp.Tag = "VAS2L" Then
                Set lblTemp = lblTemp.TagObject.TagObject.TagObject
            ElseIf lblTemp.Tag = "VAS1T" Or lblTemp.Tag = "VAS2T" Then
                Set lblTemp = lblTemp.TagObject.TagObject
            ElseIf lblTemp.Tag = "VAS1E1" Or lblTemp.Tag = "VAS2E1" Then
                Set lblTemp = lblTemp.TagObject
            End If
            For i = 1 To iVasCount
                Set lblTemp = lblTemp.TagObject
                lblTemp.left = lblTemp.left + x
                lblTemp.top = lblTemp.top + y
                If lblTemp.Tag = "VAS1L" Or lblTemp.Tag = "VAS2L" Then
                    lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
                    lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
                End If
                If Mid(lblTemp.Tag, 5, 1) = "E" Then
                    lblTemp.Text = Val(left(lblTemp.Text, InStr(lblTemp.Text, ",") - 1)) + x & "," _
                                   & Val(Right(lblTemp.Text, Len(lblTemp.Text) - InStr(lblTemp.Text, ","))) + y
                End If
            Next i
        ElseIf left(la.Tag, 3) = "CTR" Then     '�������رȲ�����ע
            Dim iCtrCount As Integer    '��¼���رȲ�����ע������������ȫ����ע��4����ֻ������������ֻ��2��
            iCtrCount = 4
            If la.Tag = la.TagObject.TagObject.Tag Then iCtrCount = 2
            
            la.TagObject.AnchorX = la.left + la.width / 2
            la.TagObject.AnchorY = la.top + la.height / 2
            Set lblTemp = la
            For i = 1 To iCtrCount - 1
                Set lblTemp = lblTemp.TagObject
                lblTemp.left = lblTemp.left + x
                lblTemp.top = lblTemp.top + y
                If Right(lblTemp.Tag, 1) = "L" Then     '���ֱ�עָ��L��ע
                    lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
                    lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
                End If
            Next i
        Else            '������ע�Ĵ������֡���ͷ��ֱ�ߡ�������Բ�����ε�
            If la.LabelType <> doLabelText Then         ''�������ֵĴ���
                If Not la.TagObject Is Nothing Then
                    Set lat = la.TagObject
                    If la.LabelType <> doLabelArrow Then        '''''���Ǽ�ͷ
                        lat.AnchorX = la.left + la.width / 2
                        lat.AnchorY = la.top + la.height / 2
                        If la.LabelType = doLabelLine Or la.LabelType = doLabelPolyLine Then
                            lat.Text = funROIResultString(la, f.SelectedImage)
                        Else
                            lat.Text = ""
                        End If
                    Else
                        lat.AnchorX = la.left + la.width
                        lat.AnchorY = la.top + la.height
                    End If
                    lat.left = lat.left + x
                    lat.top = lat.top + y
                End If
            End If
        End If
        ''''''''''����ǲü����ο���''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If f.SelectedImage.Labels.IndexOf(la) = 1 Then
            subMove25 f.SelectedImage, f        '�ü�����ϵͳ��ע1 �ƶ�2-4
            If Button_miImageInPhase = True Then subCutOutInphase f.Viewer(f.intSelectedSerial), f.SelectedImage, f
        End If
    End If
End Sub

Public Sub subChangeLableSize(la As DicomLabel, x As Long, y As Long, iR As Integer, f As frmViewer)
'------------------------------------------------
'���ܣ��ı�һ����ע�Ĵ�С,���޸�����ز�����Ϣ����ʾֵ��
'������la--��Ҫ�ı��С�ı�ע��x--��עX�����ƶ��ľ��룻y--��עY�����ƶ��ľ��룻
'      iR--���ͨ���ĸ�������ƶ���ע����11��18�ž���ֱ��в�ͬ�Ĵ�������11-18�ž���ֱ��ʾ���ϡ����С����¡����С����¡����С����ϣ����С�
'      f--�ƶ���ע�Ĵ��塣
'���أ��ޣ�ֱ���޸ı�ע��
'2009��
'------------------------------------------------
    Dim lat As DicomLabel
    Dim lblTemp As DicomLabel
    '�ƶ���ע��λ��
    'Ѫ����խ����ʹ�ñ�ע(11��14)--VAS1L��(15��18)--VAS2L����ϵ��VAS1L-VAS1T-VAS1E1-VAS1E2-VAS2L-VAS2T-VAS2E1-VAS2E2
    '���رȲ���ʹ�ñ�ע(11,14)--CTR1L,(15,18)--CTR2L,��ϵ��CTR1L-CTR1T-CTR2L-CTR2T
    '�ǶȲ���ʹ�ñ�ע��11,15,18��,��ϵ��JD2-JD1-JDT
    
    If iR = 11 Then         '���Ͻǵľ��
        la.left = la.left + x
        la.width = la.width - x
        la.top = la.top + y
        la.height = la.height - y
        If Mid(la.Tag, 1, 3) = "JD2" Then       '����Ƕȱ�ע
            la.TagObject.width = la.TagObject.width + x
            la.TagObject.height = la.TagObject.height + y
        End If
    ElseIf iR = 12 Then     '���еľ��
        la.left = la.left + x
        la.width = la.width - x
    ElseIf iR = 13 Then     '���½ǵľ��
        la.left = la.left + x
        la.width = la.width - x
        la.height = la.height + y
    ElseIf iR = 14 Then     '���еľ��
        If left(la.Tag, 3) = "VAS" Then     'Ѫ����խ����
            la.height = la.height + y
            la.width = la.width + x
        ElseIf left(la.Tag, 3) = "CTR" Then '���رȲ���
            la.height = la.height + y
            la.width = la.width + x
        Else    '������������λ�Բ��
            la.height = la.height + y
        End If
    ElseIf iR = 15 Then     '���½ǵľ��
        If Mid(la.Tag, 1, 3) = "JD1" Then   '����Ƕȱ�ע
            la.width = la.width + x
            la.height = la.height + y
            la.TagObject.TagObject.left = la.TagObject.TagObject.left + x
            la.TagObject.TagObject.width = la.TagObject.TagObject.width - x
            la.TagObject.TagObject.top = la.TagObject.TagObject.top + y
            la.TagObject.TagObject.height = la.TagObject.TagObject.height - y
        ElseIf left(la.Tag, 3) = "VAS" Then 'Ѫ����խ����
            Set lblTemp = la.TagObject.TagObject.TagObject.TagObject
            lblTemp.left = lblTemp.left + x
            lblTemp.width = lblTemp.width - x
            lblTemp.top = lblTemp.top + y
            lblTemp.height = lblTemp.height - y
        ElseIf left(la.Tag, 3) = "CTR" Then '���رȲ�������Ҫר�Ŵ������ֱ�ע��λ��
            Set lblTemp = la.TagObject.TagObject
            lblTemp.left = lblTemp.left + x
            lblTemp.width = lblTemp.width - x
            lblTemp.top = lblTemp.top + y
            lblTemp.height = lblTemp.height - y
        Else                                '����������ע
            la.width = la.width + x
            la.height = la.height + y
        End If
    ElseIf iR = 16 Then     '���еľ��
        la.width = la.width + x
    ElseIf iR = 17 Then     '���Ͻǵľ��
        la.top = la.top + y
        la.height = la.height - y
        la.width = la.width + x
    ElseIf iR = 18 Then     '���еľ��
        If Mid(la.Tag, 1, 2) = "JD" Then        '����Ƕȱ�ע
            If Mid(la.Tag, 1, 3) = "JD1" Then
                la.TagObject.TagObject.width = la.TagObject.TagObject.width + x
                la.TagObject.TagObject.height = la.TagObject.TagObject.height + y
            Else
                '����JD1����
                la.TagObject.left = la.TagObject.left + x
                la.TagObject.width = la.TagObject.width - x
                la.TagObject.top = la.TagObject.top + y
                la.TagObject.height = la.TagObject.height - y
            End If
        ElseIf left(la.Tag, 3) = "VAS" Then     'Ѫ����խ����
            Set lblTemp = la.TagObject.TagObject.TagObject.TagObject
            lblTemp.height = lblTemp.height + y
            lblTemp.width = lblTemp.width + x
        ElseIf left(la.Tag, 3) = "CTR" Then     '���رȲ���
            Set lblTemp = la.TagObject.TagObject
            lblTemp.height = lblTemp.height + y
            lblTemp.width = lblTemp.width + x
        Else                                    '����������ע
            la.top = la.top + y
            la.height = la.height - y
        End If
    End If
    
    '�������ǰ��ע�������������ױ�ע�������������Ϣ�ȵ�λ�ú���ʾֵ
    '�����ѡ�еı�ע�ǲü���ע���������Ӧ����
    If f.SelectedImage.Labels.IndexOf(la) = 1 Then
        subMove25 f.SelectedImage, f            '�ü�����ϵͳ��ע1���ƶ���ע2-5
        '�ڲü�״̬�¶Բü���������ͼ��ͬ������
        If Button_miImageInPhase = True Then subCutOutInphase f.Viewer(f.intSelectedSerial), f.SelectedImage, f
    Else
        If Mid(la.Tag, 1, 2) = "JD" Then           '����Ƕȱ�ע
            If Mid(la.Tag, 1, 3) = "JD1" Then
                Set lat = la.TagObject
                la.TagObject.left = la.left
                la.TagObject.top = la.top
                la.TagObject.AnchorX = la.left
                la.TagObject.AnchorY = la.top
                f.lblChange = funROIResultString(la, f.SelectedImage)
                lat.Text = f.lblChange
            Else
                Set lat = la.TagObject.TagObject
                la.TagObject.TagObject.left = la.TagObject.left
                la.TagObject.TagObject.top = la.TagObject.top
                la.TagObject.TagObject.AnchorX = la.TagObject.left
                la.TagObject.TagObject.AnchorY = la.TagObject.top
                f.lblChange = funROIResultString(la, f.SelectedImage)
                lat.Text = f.lblChange
            End If
        ElseIf left(la.Tag, 3) = "VAS" Then     '����Ѫ����խ����
            If iR = 15 Or iR = 18 Then
                Set lblTemp = la.TagObject.TagObject.TagObject.TagObject
            Else
                Set lblTemp = la
            End If
            
            lblTemp.TagObject.left = lblTemp.left + lblTemp.width + intTextoOffX
            lblTemp.TagObject.top = lblTemp.top + lblTemp.height + intTextoOffY
            lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
            lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
        
            Call funDrawVas(lblTemp, f.SelectedImage, IIf(lblTemp.Tag = "VAS1L", 1, 2))
        ElseIf left(la.Tag, 3) = "CTR" Then     '�������رȲ���
            If iR = 15 Or iR = 18 Then
                Set lblTemp = la.TagObject.TagObject
            Else
                Set lblTemp = la
            End If
            lblTemp.TagObject.left = lblTemp.left + lblTemp.width + intTextoOffX
            lblTemp.TagObject.top = lblTemp.top + lblTemp.height + intTextoOffY
            lblTemp.TagObject.AnchorX = lblTemp.left + lblTemp.width / 2
            lblTemp.TagObject.AnchorY = lblTemp.top + lblTemp.height / 2
            
            Call funcGetCadioThoracicRatio(la, f.SelectedImage)
        Else                                    '����������ע
            Set lat = la.TagObject
            If la.LabelType = doLabelArrow Then   '''��ͷ
                lat.left = la.left + la.width
                lat.top = la.top + la.height
                lat.AnchorX = la.left + la.width
                lat.AnchorY = la.top + la.height
            Else                'ֱ�ߡ���Բ�����Ρ��������ߵȱ�ע
                '�ԷǷ�������ע�����ɲ�������ַ���,����ϵͳ���������õ��Ƿ���ʾ�����ƽ��ֵ�������������
                If la.LabelType = doLabelEllipse Or la.LabelType = doLabelPolygon Or la.LabelType = doLabelRectangle Then
                    lat.Text = ""
                Else
                    lat.Text = funROIResultString(la, f.SelectedImage)
                End If
                lat.left = la.left + la.width + intTextoOffX
                lat.top = la.top + la.height + intTextoOffY
                lat.AnchorX = la.left + la.width / 2
                lat.AnchorY = la.top + la.height / 2
            End If
        End If
    End If
End Sub

Sub SubNoDispPeriod(im As DicomImage, f As frmViewer)
'------------------------------------------------
'���ܣ�Ϊָ��ͼ�����ر�עѡ����
'������im--��Ҫ���ر�עѡ������ͼ��f--���ر�עѡ�����Ĵ���
'���أ��ޣ�ֱ�����ر�עѡ����
'2009��
'------------------------------------------------
    Dim i As Integer
    For i = 11 To 20
      im.Labels(i).Visible = False
      im.Labels(i).left = G_INT_SYS_LABEL_HIDE_LEFT
    Next
    im.Refresh False
    If Not f.DLblOld Is Nothing Then f.DLblOld.SelectMode = doSelectNone
End Sub

Public Sub subTextCoordinate(im As DicomImage, x, y, lb As Label)
'------------------------------------------------
'���ܣ�����ͼ��ķ�ת����������ֵ�����ת��
'������im--������ת��ͼ��x--   y--   lb--
'���أ�
'2009��
'------------------------------------------------
    Dim xx As Long, Yy As Long
    Dim TXY As Single
    TXY = im.sizex / im.sizey
    xx = x
    Yy = y
    
    '������ת����������¼����µ�x,y����
    If im.RotateState = doRotateNormal Then         '����
        '���ô���
    ElseIf im.RotateState = doRotateLeft Then       '��90��
        x = Yy
        y = im.sizey * TXY - xx - lb.height / Screen.TwipsPerPixelX / im.ActualZoom
    ElseIf im.RotateState = doRotate180 Then        '��ת180��
        x = im.sizex - xx - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
        y = im.sizey - Yy - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
    ElseIf im.RotateState = doRotateRight Then      '��ת90��
        x = im.sizex / TXY - Yy - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
        y = xx
    End If
    
    '�������Ҿ�������µ��õ���������¼���x,y����
    If im.FlipState = 1 Then            '���Ҿ���
        If im.RotateState = 0 Or im.RotateState = doRotate180 Then
             x = im.sizex - x - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
        Else
             y = im.sizey * TXY - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        End If
    ElseIf im.FlipState = 2 Then        '���µ���
        If im.RotateState = 0 Or im.RotateState = doRotate180 Then
             y = im.sizey - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        Else
            x = im.sizex / TXY - x - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        End If
    ElseIf im.FlipState = 3 Then        '���Ҿ�������µ���
        If im.RotateState = 0 Or im.RotateState = doRotate180 Then
            x = im.sizex - x - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
            y = im.sizey - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        Else
            x = im.sizex / TXY - x - lb.width / Screen.TwipsPerPixelX / im.ActualZoom
            y = im.sizey * TXY - y - lb.height / Screen.TwipsPerPixelY / im.ActualZoom
        End If
    End If
End Sub

Public Sub SubChangeColor(la As DicomLabel, f As frmViewer)
'------------------------------------------------
'���ܣ��ı�ѡ��LABEL����ɫ
'������la--��Ҫ�ı���ɫ�ı�ע��f--�ı��ע��ɫ�Ĵ��塣
'���أ��ޣ�ֱ���޸ı�ע����ɫ��
'2009��
'------------------------------------------------
    Dim lblTemp As DicomLabel
    Dim i As Integer
    '''''''''''''''''''''''''''[�Ȼָ���һ����ѡ�б�ע����ɫ]'''''''''''''''''''''''''''''
    If Not f.DLblOld Is Nothing Then
        f.DLblOld.ForeColour = f.LngOldColor
        If Mid(f.DLblOld.Tag, 1, 2) = "JD" Then    ''����ǽǶ��ߵĴ���
            f.DLblOld.TagObject.ForeColour = f.LngOldColor
            If Not f.DLblOld.TagObject.TagObject Is Nothing Then f.DLblOld.TagObject.TagObject.ForeColour = f.LngOldColor
        ElseIf left(f.DLblOld.Tag, 3) = "VAS" Then      ''����Ѫ����խ������ע
            Set lblTemp = f.DLblOld
            For i = 1 To 7
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = f.LngOldColor
                End If
            Next i
        ElseIf left(f.DLblOld.Tag, 3) = "CTR" Then      ''�������رȲ�����ע
            Set lblTemp = f.DLblOld
            For i = 1 To 3
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = f.LngOldColor
                End If
            Next i
        Else
            If Not f.DLblOld.TagObject Is Nothing Then f.DLblOld.TagObject.ForeColour = f.LngOldColor
        End If
    End If
    '''''''''''''''''''''''''''[��¼��ǰ��ע]'''''''''''''''''''''''''''''
    f.LngOldColor = la.ForeColour
    Set f.DLblOld = la
    '''''''''''''''''''''''''''''[�ı䵱ǰ��ѡ�б�ע����ɫ]'''''''''''''''''''''''''''
    la.ForeColour = lngLabelSelectedColor
    If la.LabelType <> doLabelText Then
        If Mid(la.Tag, 1, 2) = "JD" Then
            la.TagObject.ForeColour = lngLabelSelectedColor
            If Not la.TagObject.TagObject Is Nothing Then la.TagObject.TagObject.ForeColour = lngLabelSelectedColor
        ElseIf left(la.Tag, 3) = "VAS" Then
            Set lblTemp = la
            For i = 1 To 7
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = lngLabelSelectedColor
                End If
            Next i
        ElseIf left(la.Tag, 3) = "CTR" Then
            Set lblTemp = la
            For i = 1 To 3
                If Not lblTemp.TagObject Is Nothing Then
                    Set lblTemp = lblTemp.TagObject
                    lblTemp.ForeColour = lngLabelSelectedColor
                End If
            Next i
        Else
            If Not f.DLblOld.TagObject Is Nothing Then la.TagObject.ForeColour = lngLabelSelectedColor
        End If
    End If
End Sub

Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'���ܣ�����һ��LABEL���󣬲���������ʼ����
'������lType--��ע�����ͣ�lLeft--��ע��Leftֵ��lTop--��ע��Topֵ��lWidth--��ע��Widthֵ��lHeight--��ע��Heightֵ��
'���أ������ɵı�ע��
'2009��
'------------------------------------------------
    Dim l As New DicomLabel
    l.LabelType = lType
    l.Transparent = True
    l.XOR = True
    l.ImageTied = True
    l.left = lLeft
    l.top = lTop
    l.width = lWidth
    l.height = lHeight
    l.Margin = 0
    l.ScaleFontSize = blnLabelTextScaleFontSize
    l.AutoSize = True
    l.FontSize = lngLabelFontSize
    l.LineStyle = lngLabelLineStyleNorm
    l.LineWidth = lngLabelLineWidthNorm
    l.ForeColour = lngLabelColor
    If l.LabelType = 0 Then
        l.Transparent = False
    Else
        If Button_mi3dCursor <> True Then
'            l.Outline = True
        End If
    End If
    Set GetNewLabel = l
End Function

Public Sub subDeleteAppointLabel(im As DicomImage, strL As String)
'------------------------------------------------
'���ܣ�ɾ��ָ�����͵ı�ע
'������im--��Ҫɾ��ָ�����ͱ�ע��ͼ��strL--��Ҫɾ���ı�עtag�а�����ָ�����ݡ�
'���أ��ޣ�ֱ��ɾ��ͼ���ָ����ע
'2009��
'------------------------------------------------
    Dim i  As Integer
    If strL = "" Then Exit Sub
    For i = im.Labels.Count To G_INT_SYS_LABEL_COUNT Step -1
        If Mid(im.Labels(i).Tag, 1, Len(strL)) = strL Then im.Labels.Remove i
    Next
End Sub

Public Sub SubInitPeriod(im As DicomImage)
'------------------------------------------------
'���ܣ�Ϊÿһ��ͼ����ǰn��ϵͳ�����n�������ɳ���G_INT_SYS_LABEL_COUNT����
'������im--��Ҫ����ϵͳ��ע��ϵͳ�������ͼ��
'���أ��ޣ�ֱ����ͼ��������n��ϵͳ���
'2009��
'------------------------------------------------
    Dim CurrentLabel As DicomLabel
    Dim i As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If im.Labels.Count > 0 Then
        MsgBox "����ͼ���Ѿ��ж��󣬲��ܳ�ʼ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To G_INT_SYS_LABEL_COUNT
        Set CurrentLabel = New DicomLabel
        CurrentLabel.LabelType = doLabelRectangle
        CurrentLabel.Transparent = False
        CurrentLabel.XOR = False
        CurrentLabel.BackColour = lngPeriodColor
        CurrentLabel.ForeColour = 0
        CurrentLabel.height = intPeriodSize
        CurrentLabel.width = CurrentLabel.height
        CurrentLabel.Visible = False
        CurrentLabel.left = G_INT_SYS_LABEL_HIDE_LEFT
        CurrentLabel.ImageTied = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If i > 1 And i < 6 Then  '''''��Ϊ�ü��ڸ��õľ��
            CurrentLabel.BackColour = vbBlack
            CurrentLabel.ForeColour = vbBlack
        End If
        '''''''''''''''�ü��þ��ο�'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If i = 1 Then
            CurrentLabel.Transparent = True
            CurrentLabel.ForeColour = vbBlack 'vbBlue
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''��λ��ער�þ��'''''''''''''''''
        If i >= G_INT_SYS_LABEL_TIWEI And i <= G_INT_SYS_LABEL_TIWEI + 3 Then
            CurrentLabel.LabelType = doLabelSpecial
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.width = 200
            CurrentLabel.height = 200
'            CurrentLabel.BackColour = lngViewerBackColor
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.Margin = 2
            CurrentLabel.ImageTied = False
            CurrentLabel.Transparent = True
        End If
        
        If i >= G_INT_SYS_LABEL_MPRV And i <= G_INT_SYS_LABEL_MPR_RESULT_V Then   '''ʸ��״�ؽ�ʹ�þ��
            If i = G_INT_SYS_LABEL_MPRV Or i = G_INT_SYS_LABEL_MPRH _
                Or i = G_INT_SYS_LABEL_MPR_RESULT_H Or i = G_INT_SYS_LABEL_MPR_RESULT_V Then    '���������ߺͽ��ͼ�е�����ͶӰ��
                CurrentLabel.LabelType = doLabelLine
                CurrentLabel.ForeColour = vbRed
                CurrentLabel.LineWidth = 2
            Else        '�������ϵ�������Ƶ�
                CurrentLabel.Transparent = False
                CurrentLabel.LabelType = doLabelEllipse
                CurrentLabel.LineWidth = 1
                CurrentLabel.ForeColour = RGB(255, 255, 255)
                CurrentLabel.width = G_INT_MPR_RADIUS
                CurrentLabel.height = G_INT_MPR_RADIUS
            End If
            CurrentLabel.ImageTied = True
        End If
        
        '''''''''''''30�ű�ע,������ʾ����λ'''''''''''''''''''''''''''''''''''''''''''''''''''''
        If i = G_INT_SYS_LABEL_WWWL Then
            CurrentLabel.LabelType = doLabelText
            CurrentLabel.width = 0          '��Ⱥ͸߶ȵ����û�Ӱ��AutoSize
            CurrentLabel.height = 0
            CurrentLabel.ImageTied = False  '�����û�Ӱ��ScaleWithCell
            CurrentLabel.Transparent = True
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.AutoSize = True
'            CurrentLabel.BackColour = lngViewerBackColor
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.left = 0
            CurrentLabel.Text = "WL"
            CurrentLabel.Visible = True
            CurrentLabel.Alignment = doAlignCentre
        End If
        
        '�������Ľ���Ϣ
        If i >= G_INT_SYS_LABEL_PAT_INFO And i <= G_INT_SYS_LABEL_PAT_INFO + 3 Then
            CurrentLabel.LabelType = doLabelText
            CurrentLabel.width = 0          '��Ⱥ͸߶ȵ����û�Ӱ��AutoSize
            CurrentLabel.height = 0
            CurrentLabel.ImageTied = False  '�����û�Ӱ��ScaleWithCell
            CurrentLabel.Transparent = True
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.Font.Name = strPatientInfoFontName
            CurrentLabel.Font.Size = lngPatientInfoFontSize
            CurrentLabel.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            CurrentLabel.Font.Italic = blnPatientInfoFontItalic
            CurrentLabel.AutoSize = True
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.left = 0
            CurrentLabel.top = 0
            Select Case i
                Case G_INT_SYS_LABEL_PAT_INFO
                    CurrentLabel.Tag = "PAT1"
                    CurrentLabel.Alignment = doAlignLeft
                Case G_INT_SYS_LABEL_PAT_INFO + 1
                    CurrentLabel.Tag = "PAT2"
                    CurrentLabel.Alignment = doAlignBottomLeft
                Case G_INT_SYS_LABEL_PAT_INFO + 2
                    CurrentLabel.Tag = "PAT3"
                    CurrentLabel.Alignment = doAlignBottomRight
                Case G_INT_SYS_LABEL_PAT_INFO + 3
                    CurrentLabel.Tag = "PAT4"
                    CurrentLabel.Alignment = doAlignRight
            End Select
        End If
        
        '���˱����Ϣ�ͱ�ߵ�λ
        If i >= G_INT_SYS_LABEL_RULLER And i <= G_INT_SYS_LABEL_RULLER + 7 Then
            If i >= G_INT_SYS_LABEL_RULLER + 4 Then
                CurrentLabel.LabelType = doLabelText    '��ߵ�λ
                CurrentLabel.AutoSize = True
                CurrentLabel.width = 0          '��Ⱥ͸߶ȵ����û�Ӱ��AutoSize
                CurrentLabel.height = 0
                CurrentLabel.Transparent = True
                CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
                CurrentLabel.Font.Name = strPatientInfoFontName
                CurrentLabel.Font.Size = lngPatientInfoFontSize
                CurrentLabel.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
                CurrentLabel.Font.Italic = blnPatientInfoFontItalic
                CurrentLabel.AutoSize = True
                CurrentLabel.left = 0
                CurrentLabel.top = 0
            Else
                CurrentLabel.LabelType = doLabelRuler   '���
            End If
            CurrentLabel.ImageTied = False  '�����û�Ӱ��ScaleWithCell
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = False
            CurrentLabel.ForeColour = lngRulerLeftColor
            CurrentLabel.LineWidth = intRulerLineWidth
        End If
        
        '��ӡ���
        If i = G_INT_SYS_LABEL_PRINT_TAG Then
            CurrentLabel.LabelType = doLabelText
            CurrentLabel.width = 400
            CurrentLabel.height = lngPatientInfoFontSize * 4
            CurrentLabel.ImageTied = False  '�����û�Ӱ��ScaleWithCell,1000,1000
            CurrentLabel.Transparent = True
            CurrentLabel.ScaleWithCell = True
            CurrentLabel.ScaleFontSize = blnpatientInfoScaleFontSize
            CurrentLabel.Font.Name = strPatientInfoFontName
            CurrentLabel.Font.Size = lngPatientInfoFontSize
            CurrentLabel.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            CurrentLabel.Font.Italic = blnPatientInfoFontItalic
            CurrentLabel.ForeColour = lngpatientInfoColor
            CurrentLabel.BackColour = vbRed
            CurrentLabel.left = 300
            CurrentLabel.top = 50
            CurrentLabel.Text = "�Ѵ�ӡ"
            CurrentLabel.ShowTextBox = True
            CurrentLabel.Shadow = doShadowBottomRight
            CurrentLabel.Alignment = doAlignCentre
        End If
        
        im.Labels.Add CurrentLabel
    Next
    im.Labels(1).TagObject = im.Labels(6)
    im.Labels(6).TagObject = im.Labels(1)
End Sub

Public Sub UpdateMarkers(Image As DicomImage, Optional blnShow As Boolean = True)
'------------------------------------------------
'���ܣ�����ͼ����ʾ�����ز�����λ��Ϣ
'������Image--��Ҫ��ʾ��λ��Ϣ��ͼ��blnShow--�Ƿ���ʾ������Ϣ��
'���أ��ޣ�ֱ����ͼ������ʾ��������λ��Ϣ��
'------------------------------------------------
    Dim DG As New DicomGlobal
    Dim l As DicomLabel, i As Integer
    DG.DirectionStrings = IIf(blnChinaMark, "��\��\ǰ\��\��\ͷ", "R\L\A\P\I\S")
    If blnShow Then
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI)
        If blnAnatomicMarkersLeft Then
            l.left = 0
            l.top = 500
            l.Text = "LEFT"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
        
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI + 1)
        If blnAnatomicMarkersTop Then
            l.left = 500 - l.width / 2
            l.top = 0
            l.Alignment = doAlignCentre
            l.Text = "TOP"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
        
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI + 2)
        If blnAnatomicMarkersRight Then
            l.left = 1000 - l.width
            l.top = 500
            l.Alignment = doAlignRight
            l.Text = "RIGHT"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
        
        Set l = Image.Labels(G_INT_SYS_LABEL_TIWEI + 3)
        If blnAnatomicMarkersBottom Then
            l.left = 500 - l.width / 2
            l.top = 1000 - l.height
            l.Alignment = doAlignBottomCentre
            l.Text = "BOTTOM"
            l.Font.Name = strPatientInfoFontName
            l.Font.Size = lngPatientInfoFontSize
            l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
            l.Font.Italic = blnPatientInfoFontItalic
            l.ScaleFontSize = blnpatientInfoScaleFontSize
            l.ForeColour = lngpatientInfoColor
            l.Visible = True
        Else
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        End If
    Else        '������λ��ע
        For i = G_INT_SYS_LABEL_TIWEI To G_INT_SYS_LABEL_TIWEI + 3
            Set l = Image.Labels(i)
            l.Visible = False
            l.left = G_INT_SYS_LABEL_HIDE_LEFT
        Next i
    End If
    Image.Refresh True
End Sub


Public Function UpdateRuler(im As DicomImage, blnDisp As Boolean) As Long
'------------------------------------------------
'���ܣ���ʾͼ����,ֱ����ʾ�����ر��
'������im--��ʾ��ߵ�ͼ��blnDisp--�Ƿ���ʾ��ߣ�True��ʾ��ߣ�False����ʾ���
'���أ� 0---������1--��߱�ע�������ԣ�2-��������
'------------------------------------------------
    Dim l As DicomLabel
    Dim lUnit As DicomLabel
    
    On Error GoTo err
    
    '���ͼ��ı�ע����Ƿ����
    If im.Labels.Count < G_INT_SYS_LABEL_RULLER + 4 Then
        UpdateRuler = 1
        Exit Function
    End If
    
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER)           '�������ߺͱ�ߵ�λ
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 4)
    l.left = IIf(blnDisp, intRulerLeft, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = intRulerTop
    l.width = intRulerWidth
    l.height = intRulerHeight
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipLeft And blnDisp, True, False)
    '������Ϣʹ�õ���ͷ��0--��ʹ����ͷ��1--������ͷ��2--Ӣ����ͷ
    If lngPatientInfoTitle = 0 Then
        lUnit.Text = l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    ElseIf lngPatientInfoTitle = 2 Then
        lUnit.Text = "Unit:" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    Else
        lUnit.Text = "��λ��" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    End If
    lUnit.left = l.left
    lUnit.top = l.top + l.height
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER + 1)               '�����ϱ�ߺ͵�λ
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 5)
    l.left = IIf(blnDisp, intRulerTop, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = intRulerLeft
    l.width = intRulerHeight
    l.height = intRulerWidth
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipTop And blnDisp, True, False)
    lUnit.Text = "��λ��" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    lUnit.left = l.left + l.width
    lUnit.top = l.top
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER + 2)           '�����ұ�ߺ͵�λ
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 6)
    l.left = IIf(blnDisp, 1000 - intRulerLeft, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = intRulerTop
    l.width = -intRulerWidth
    l.height = intRulerHeight
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipRight And blnDisp, True, False)
    lUnit.Text = "��λ��" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    lUnit.left = l.left + l.width
    lUnit.top = l.top + l.height
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set l = im.Labels(G_INT_SYS_LABEL_RULLER + 3)           '�����±�ߺ͵�λ
    Set lUnit = im.Labels(G_INT_SYS_LABEL_RULLER + 7)
    l.left = IIf(blnDisp, intRulerTop, G_INT_SYS_LABEL_HIDE_LEFT)
    l.top = 1000 - intRulerLeft
    l.width = intRulerHeight
    l.height = -intRulerWidth
    l.ForeColour = lngRulerLeftColor
    l.LineWidth = intRulerLineWidth
    l.Visible = IIf(blnRulerDsipBottom And blnDisp, True, False)
    lUnit.Text = "��λ��" & l.TickSpacing(False) & im.Labels(11).ROIDistanceUnits
    lUnit.left = l.left + l.width
    lUnit.top = l.top + l.height
    lUnit.Font.Name = strPatientInfoFontName
    lUnit.Font.Size = lngPatientInfoFontSize
    lUnit.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    lUnit.Font.Italic = blnPatientInfoFontItalic
    lUnit.ScaleFontSize = blnpatientInfoScaleFontSize
    lUnit.ForeColour = l.ForeColour
    lUnit.LineWidth = l.LineWidth
    lUnit.Visible = l.Visible
    If im.Labels(11).ROIDistanceUnits = "Pixels" Then lUnit.Visible = False
    Exit Function
    
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    UpdateRuler = 2
End Function


Public Sub subDispImageInfo(im As DicomImage, blnDisp As Boolean, blnRefreshPatiIn As Boolean, blnRefreshWL As Boolean, Optional strPatientInfo1 As String = "", Optional strPatientInfo2 As String = "", _
                     Optional strPatientInfo3 As String = "", Optional strPatientInfo4 As String = "")
'------------------------------------------------
'���ܣ���ʾ�����ز���ͼ���Ľ���Ϣ�ʹ���λ��ʾ
'������im--��ʾ������Ϣ��ͼ��blnDisp--��ʾ�����ز����Ľ���Ϣ�ʹ���λ��TrueΪ��ʾ��FalseΪ���أ�
'      blnRefreshPatiIn--�Ƿ��մ�����ĸ��Ľ���Ϣ�ַ�����ˢ�²����Ľ���Ϣ��TrueΪˢ�£�FalseΪ��ˢ�£�
'      blnRefreshWL -- �Ƿ�ˢ��ͼ��Ĵ���λ
'      strPatientInfo1--���ϽǵĲ�����Ϣ��strPatientInfo2--���½ǵĲ�����Ϣ��
'      strPatientInfo3--���½ǵĲ�����Ϣ��strPatientInfo4--���ϽǵĲ�����Ϣ��
'���أ�
'------------------------------------------------
    Dim i, j, intTop, intLeft As Integer
    Dim l As DicomLabel
    
    On Error GoTo err
    
    '����
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.ForeColour = lngpatientInfoColor
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo1, l.Text)
    '����
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO + 1)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.ForeColour = lngpatientInfoColor
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo2, l.Text)
    '����
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO + 2)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.ForeColour = lngpatientInfoColor
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo3, l.Text)
    '����
    Set l = im.Labels(G_INT_SYS_LABEL_PAT_INFO + 3)
    l.Visible = blnDisp
    l.left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    l.ScaleFontSize = blnpatientInfoScaleFontSize
    l.Font.Name = strPatientInfoFontName
    l.Font.Size = lngPatientInfoFontSize
    l.Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    l.Font.Italic = blnPatientInfoFontItalic
    l.ForeColour = lngpatientInfoColor
    l.Text = IIf(blnRefreshPatiIn, strPatientInfo4, l.Text)
    ''''''����λ��ע����''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    im.Labels(G_INT_SYS_LABEL_WWWL).Visible = blnDisp
    im.Labels(G_INT_SYS_LABEL_WWWL).left = IIf(blnDisp, 0, G_INT_SYS_LABEL_HIDE_LEFT)
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Name = strPatientInfoFontName
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Size = lngPatientInfoFontSize
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Bold = IIf(blnPatientInfoFontBold, True, False)
    im.Labels(G_INT_SYS_LABEL_WWWL).Font.Italic = blnPatientInfoFontItalic
    im.Labels(G_INT_SYS_LABEL_WWWL).ScaleFontSize = blnpatientInfoScaleFontSize
    im.Labels(G_INT_SYS_LABEL_WWWL).ForeColour = lngpatientInfoColor
    im.Labels(G_INT_SYS_LABEL_WWWL).Text = IIf(blnRefreshWL, "W:" & im.width & "-L:" & im.Level, im.Labels(G_INT_SYS_LABEL_WWWL).Text)
    If lngWinWidthLevelLocation = 1 Then  '''1-�ϱߣ�2-�±ߣ�3-��ߣ�4-�ұ�
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 0
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignCentre
    ElseIf lngWinWidthLevelLocation = 2 Then
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 0
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignBottomCentre
    ElseIf lngWinWidthLevelLocation = 3 Then
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 500
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignLeft
    ElseIf lngWinWidthLevelLocation = 4 Then
        im.Labels(G_INT_SYS_LABEL_WWWL).top = 500
        im.Labels(G_INT_SYS_LABEL_WWWL).Alignment = doAlignRight
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub subGetImgInfoLabel(intSeriesIndex As Integer, intIndexType As Integer, img As DicomImage, strInfoLabel() As String, Optional lngPrefix As Long = 0, Optional blnIsOnlyExport As Boolean = False)
'------------------------------------------------
'���ܣ���ͼ������ȡ���˵��ĸ�����Ϣ��ע�����ϵͳ�����������ĸ��Ǳ�ע������ʹ��
'������ intSeriesIndex -- ͼ���������е�����
'       intIndexType -- �������������ͣ�0--��ZLSeriesInfos��ȡ��1 -- ��ZLShowSeriesInfos��ȡ
'       img--��ȡ������Ϣ��ͼ��
'       strInfoLabel()--�ŷ���ֵ�����飻
'       lngPrefix--��ʾʹ��ǰ׺�����͡�0-����ǰ׺��1-ʹ������ǰ׺��2-ʹ��Ӣ��ǰ׺
'���أ��ޣ�ֱ����д��strInfoLabel()�������档
'2009��
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim iLocation As Integer
    Dim v As Variant
    Dim iCount(4) As Integer
    Dim iMax As Integer
    Dim strInfo() As String
    Dim StrTmp As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If UBound(strInfoLabel) <> 4 Then
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To 4
        strInfoLabel(i) = ""
        iCount(i) = 0
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To lngInfoLabelCount
        If aInfoLabelLocate(i).bUsed Then
            iLocation = aInfoLabelLocate(i).lngLocation
            iCount(iLocation) = iCount(iLocation) + 1
        End If
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iMax = iCount(1)
    If iMax < iCount(2) Then iMax = iCount(2)
    If iMax < iCount(3) Then iMax = iCount(3)
    If iMax < iCount(4) Then iMax = iCount(4)
            
    ReDim strInfo(4, iMax) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To lngInfoLabelCount
        If aInfoLabelLocate(i).bUsed And (Not blnIsOnlyExport Or aInfoLabelLocate(i).blnIsExport) Then
            iLocation = aInfoLabelLocate(i).lngLocation
            If (img.Attributes(Val("&H" & aInfoLabelLocate(i).strGroup), Val("&H" & aInfoLabelLocate(i).strElement)).Exists) _
                Or (aInfoLabelLocate(i).strGroup = "1" And aInfoLabelLocate(i).strElement = "1") _
                Or (aInfoLabelLocate(i).strGroup = "2" And aInfoLabelLocate(i).strElement = "2") _
                Or (aInfoLabelLocate(i).strGroup = "3" And aInfoLabelLocate(i).strElement = "3") _
                Or (aInfoLabelLocate(i).strGroup = "0010" And aInfoLabelLocate(i).strElement = "1010") Then
                
                If aInfoLabelLocate(i).strGroup = "1" And aInfoLabelLocate(i).strElement = "1" Then
                    '����TagΪ��1,1����ͼ�����ԣ���Ҫ���������ļ��Ϊ��ʶ�����м��㡣
                    v = funcCalImgInfoLabel(img, aInfoLabelLocate(i).strCName)
                ElseIf aInfoLabelLocate(i).strGroup = "2" And aInfoLabelLocate(i).strElement = "2" Then
                    '����TagΪ��2,2)��ͼ�����ԣ����û�����ģ�ֱ����ʾ������
                    v = aInfoLabelLocate(i).strCName
                ElseIf aInfoLabelLocate(i).strGroup = "3" And aInfoLabelLocate(i).strElement = "3" Then
                    '����TagΪ��3,3����ͼ�����ԣ������ݿ��ֶΣ�������Ϣ����ȡԤ�ȴ洢�����ݿ���Ϣ
                    v = funGetDBInfoLabel(intSeriesIndex, aInfoLabelLocate(i).strCName, intIndexType)
                ElseIf aInfoLabelLocate(i).strGroup = "0020" And aInfoLabelLocate(i).strElement = "0013" Then
                    'ͼ��ţ�20��13�����⴦������Ƕ�֡ͼ����ʾ֡��
                    If img.FrameCount > 1 Then
                        v = img.Attributes(&H20, &H13).Value & "-" & img.Frame
                    Else
                        v = img.Attributes(&H20, &H13).Value
                    End If
                ElseIf aInfoLabelLocate(i).strGroup = "0010" And aInfoLabelLocate(i).strElement = "1010" Then
                    '���䣨0010,1010�����⴦���������Ϊ�գ���ͨ���������ڼ�������
                    If Not img.Attributes(&H10, &H1010).Exists Or IsNull(img.Attributes(&H10, &H1010).Value) Then
                        If Not IsNull(img.DateOfBirth) Then
                            v = DateDiff("yyyy", img.DateOfBirth, Now)
                        Else
                            v = ""
                        End If
                    Else
                        v = img.Attributes(&H10, &H1010).Value
                    End If
                Else
                    v = img.Attributes(Val("&H" & aInfoLabelLocate(i).strGroup), Val("&H" & aInfoLabelLocate(i).strElement)).Value
                End If
                If TypeName(v) = "String()" Then
                    StrTmp = v(1)
                Else
                    StrTmp = IIf(IsNull(v), "", v)
                End If
                '����������ֵ���뵽�ĸ��ǵ�������
                If aInfoLabelLocate(i).strGroup = "2" And aInfoLabelLocate(i).strElement = "2" Then
                    strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = StrTmp
                Else
                    If IsNull(v) Or StrTmp = "" Then
                        strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = "--"
                    Else
                        Select Case lngPrefix
                        Case 0          ''��ʹ��ǰ׺
                            strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = StrTmp
                        Case 1          ''ʹ������ǰ׺
                            strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = aInfoLabelLocate(i).strCName & " " & StrTmp
                        Case 2          ''ʹ��Ӣ��ǰ׺
                            strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = aInfoLabelLocate(i).strEName & " " & StrTmp
                        End Select
                    End If
                End If
            Else
                strInfo(iLocation, aInfoLabelLocate(i).lngOrder) = "--"
            End If
        End If
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '���������������������ʾ
    For i = 1 To 4
        For j = 0 To iCount(i) - 1
            If strInfo(i, j) <> "--" Then
                If strInfoLabel(i) = "" Then
                    strInfoLabel(i) = strInfo(i, j)
                Else
                    strInfoLabel(i) = strInfoLabel(i) & vbCrLf & strInfo(i, j)
                End If
            End If
        Next
    Next
End Sub


Public Sub subInitImageLabels(intSeriesIndex As Integer, intIndexType As Integer, img As DicomImage, blnShowLabel As Boolean, _
           Optional blnGetImgInfo As Boolean = False, Optional blnCreateSysLabel As Boolean = False, Optional blnIsOnlyExport As Boolean = False)
'------------------------------------------------
'���ܣ���ʼ������ʾ������ָ��ͼ��ı�ע��Ϣ:ϵͳ��ע����λ��ע����ߣ��Ľ���Ϣ������λ��
'      ֻ��һ��ͼ����в�����
'������img--��Ҫ����ͼ���ע��Ϣ��ͼ��blnShowLabel--��ʾ�����ر�ע��True-��ʾ��ע��False-���ر�ע��
'      blnGetImgInfo-�Ƿ��ȡͼ���Ľ���Ϣ��True-��ͼ���ȡ�Ľ���Ϣ��False-����ȡ�Ľ���Ϣ��
'      blnCreateSysLabel-�Ƿ񴴽�ϵͳ��ע��True-����ϵͳ��ע��False-������ϵͳ��ע��
'���أ��ޣ�ֱ�Ӹı�ͼ��
'2009��
'------------------------------------------------
    Dim strInfo(4) As String
    If blnCreateSysLabel Then SubInitPeriod img      ''��ʼ�� G_INT_SYS_LABEL_COUNT ��ϵͳ���
    
    'If Not blnIsOnlyExport Then
        UpdateMarkers img, blnShowLabel   ''��ʾ������λ��Ϣ
        UpdateRuler img, blnShowLabel     ''��ʾ���˱��
    'End If
    
    If blnGetImgInfo Then
        subGetImgInfoLabel intSeriesIndex, intIndexType, img, strInfo, lngPatientInfoTitle, blnIsOnlyExport    ''�����ݿ��ȡ�����Ľ���Ϣ
    End If
    
    subDispImageInfo img, blnShowLabel, blnGetImgInfo, blnGetImgInfo, strInfo(1), strInfo(2), strInfo(3), strInfo(4)   ''��ʾ�����Ľ���Ϣ�ʹ���λ��Ϣ
End Sub


Private Function funGetDBInfoLabel(intSeriesIndex As Integer, strFieldName As String, intIndexType As Integer)
'------------------------------------------------
'���ܣ����ݴ�������ļ�ƣ���������Ϣ�в��Ҷ�Ӧ�ĽǱ�ע����ʾֵ��
'������ intSeriesIndex -- ͼ���������е�����
'       strFieldName -- ��Ҫ��ʾ�����ݿ���Ϣ������
'       intIndexType -- �������������ͣ�0--��ZLSeriesInfos��ȡ��1 -- ��ZLShowSeriesInfos��ȡ
'���أ��������ļ����ȡ��������ʾֵ��
'------------------------------------------------
    funGetDBInfoLabel = Null
    
    On Error GoTo err
    
    If intIndexType = 0 Then
        If intSeriesIndex > 0 And intSeriesIndex <= ZLSeriesInfos.Count Then
            If strFieldName = "[����]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strCName
            ElseIf strFieldName = "[Ӣ����]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strEName
            ElseIf strFieldName = "[�Ա�]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strSex
            ElseIf strFieldName = "[����]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strAge
            ElseIf strFieldName = "[����]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strStudyID
            ElseIf strFieldName = "[ҽ��ID]" Then
                funGetDBInfoLabel = ZLSeriesInfos(intSeriesIndex).strOrderID
            End If
        End If
    Else
        If intSeriesIndex > 0 And intSeriesIndex <= ZLShowSeriesInfos.Count Then
            If strFieldName = "[����]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strCName
            ElseIf strFieldName = "[Ӣ����]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strEName
            ElseIf strFieldName = "[�Ա�]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strSex
            ElseIf strFieldName = "[����]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strAge
            ElseIf strFieldName = "[����]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strStudyID
            ElseIf strFieldName = "[ҽ��ID]" Then
                funGetDBInfoLabel = ZLShowSeriesInfos(intSeriesIndex).strOrderID
            End If
        End If
    End If
    Exit Function
err:
    '�����κδ������ؿ�
    funGetDBInfoLabel = Null
End Function

Private Function funcCalImgInfoLabel(img As DicomImage, strCalName As String) As Variant
'------------------------------------------------
'���ܣ����ݴ�������ļ�ƣ��������Ӧ�ĽǱ�ע����ʾֵ��
'������img--��Ҫ��ʾ�ĽǱ�ע��ͼ��aInfo--��ע��Ϣ���͡�
'���أ��������ļ�Ƽ����������ʾֵ��
'2009��
'------------------------------------------------
    funcCalImgInfoLabel = Null
    Dim v1 As Variant
    Dim v2 As Variant
    Dim v3 As Variant
    If strCalName = "�����" Then
        If img.Attributes(&H18, &H50).Exists Then
            v1 = img.Attributes(&H18, &H50).Value   'slice thickness
            v2 = Null
            If IsNull(v1) Then Exit Function
            If img.Attributes(&H18, &H88).Exists Then
                v2 = img.Attributes(&H18, &H88).Value   'spacing between slices
            End If
            If IsNull(v2) Then
                funcCalImgInfoLabel = v1 & "thk"
            Else
                funcCalImgInfoLabel = v1 & "thk/" & v2 - v1 & "sp"
            End If
        End If
    ElseIf strCalName = "��ҰFOV" Then
        If img.Attributes(&H28, &H10).Exists And img.Attributes(&H28, &H11).Exists And img.Attributes(&H28, &H30).Exists Then
            v1 = img.Attributes(&H28, &H10).Value   'rows
            v2 = img.Attributes(&H28, &H11).Value   'columns
            v3 = img.Attributes(&H28, &H30).Value  'pixel spacing
            If IsNull(v1) Or IsNull(v2) Or IsNull(v3) Then
                Exit Function
            End If
            '��Թ����ο�ҽԺ��������ҩ����������Ϣ������˾��DR���⴦���������ؾ����ֶ�ֻ��һάֵ
            If TypeName(v3) = "String()" Then
                If UBound(v3) < 2 Then
                    Exit Function
                End If
            End If
            funcCalImgInfoLabel = Format(v1 * v3(1) / 10, "#00.0") & " CM X " & Format(v2 * v3(2) / 10, "#00.0") & " CM"
        End If
    End If
End Function


Public Sub subSaveLabelToImg(img As DicomImage)
'------------------------------------------------
'���ܣ�����ע���浽DICOMͼ���ͷ��Ϣ����
'������img--��Ҫ�����ע��ͼ��
'���أ��ޣ�ֱ�ӽ���ע��Ϣ��д��ͼ���ͷ��Ϣ���档
'2009��
'------------------------------------------------
    Dim la As DicomLabel
    Dim ds As DicomDataSet
    Dim dssAll As DicomDataSets
    Dim i As Integer
    Dim iIncrease As Integer
    Dim lngTemp As Long
    Dim strPoints As String
    Dim j As Integer
    Dim vPoints As Variant
    Dim lngPointsCount As Long
    Dim aSaveTagObject() As Integer
    ReDim aSaveTagObject(img.Labels.Count) As Integer
    
    Dim v As Variant
    'ͼ���б�עȫ�����Ǳ�����ע��ǰ��ʮ����ע��ϵͳ�����ı�ע
    If img.Labels.Count <= G_INT_SYS_LABEL_COUNT Then
        Exit Sub
    End If
    Set dssAll = New DicomDataSets
    iIncrease = 0
    For i = G_INT_SYS_LABEL_COUNT + 1 To img.Labels.Count
        '�����ע��img��
        Set la = img.Labels(i)
        Set ds = New DicomDataSet
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TINDEX").Group) + iIncrease)), Val("&h" & cLabelStore("TINDEX").Element), cLabelStore("TINDEX").VR, iIncrease
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Left").Group) + iIncrease)), Val("&h" & cLabelStore("Left").Element), cLabelStore("Left").VR, la.left
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Top").Group) + iIncrease)), Val("&h" & cLabelStore("Top").Element), cLabelStore("Top").VR, la.top
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Width").Group) + iIncrease)), Val("&h" & cLabelStore("Width").Element), cLabelStore("Width").VR, la.width
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Height").Group) + iIncrease)), Val("&h" & cLabelStore("Height").Element), cLabelStore("Height").VR, la.height
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("LabelType").Group) + iIncrease)), Val("&h" & cLabelStore("LabelType").Element), cLabelStore("LabelType").VR, la.LabelType
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("ImageTied").Element), cLabelStore("ImageTied").VR, la.ImageTied
        
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Alignment").Group) + iIncrease)), Val("&h" & cLabelStore("Alignment").Element), cLabelStore("Alignment").VR, la.Alignment
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AnchorImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorImageTied").Element), cLabelStore("AnchorImageTied").VR, la.AnchorImageTied
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AnchorX").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorX").Element), cLabelStore("AnchorX").VR, la.AnchorX
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AnchorY").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorY").Element), cLabelStore("AnchorY").VR, la.AnchorY
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Angle").Group) + iIncrease)), Val("&h" & cLabelStore("Angle").Element), cLabelStore("Angle").VR, la.Angle
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("AutoSize").Group) + iIncrease)), Val("&h" & cLabelStore("AutoSize").Element), cLabelStore("AutoSize").VR, la.AutoSize
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("BackColour").Group) + iIncrease)), Val("&h" & cLabelStore("BackColour").Element), cLabelStore("BackColour").VR, la.BackColour
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("BackStyle").Group) + iIncrease)), Val("&h" & cLabelStore("BackStyle").Element), cLabelStore("BackStyle").VR, la.BackStyle
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("FontName").Group) + iIncrease)), Val("&h" & cLabelStore("FontName").Element), cLabelStore("FontName").VR, la.FontName
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("FontSize").Group) + iIncrease)), Val("&h" & cLabelStore("FontSize").Element), cLabelStore("FontSize").VR, la.FontSize
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ForeColour").Group) + iIncrease)), Val("&h" & cLabelStore("ForeColour").Element), cLabelStore("ForeColour").VR, la.ForeColour
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("LineStyle").Group) + iIncrease)), Val("&h" & cLabelStore("LineStyle").Element), cLabelStore("LineStyle").VR, la.LineStyle
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("LineWidth").Group) + iIncrease)), Val("&h" & cLabelStore("LineWidth").Element), cLabelStore("LineWidth").VR, la.LineWidth
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Margin").Group) + iIncrease)), Val("&h" & cLabelStore("Margin").Element), cLabelStore("Margin").VR, la.Margin
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Outline").Group) + iIncrease)), Val("&h" & cLabelStore("Outline").Element), cLabelStore("Outline").VR, la.Outline
        
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("RotateTextWithImage").Group) + iIncrease)), Val("&h" & cLabelStore("RotateTextWithImage").Element), cLabelStore("RotateTextWithImage").VR, la.RotateTextWithImage
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ScaleFontSize").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleFontSize").Element), cLabelStore("ScaleFontSize").VR, la.ScaleFontSize
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ScaleWithCell").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleWithCell").Element), cLabelStore("ScaleWithCell").VR, la.ScaleWithCell
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Shadow").Group) + iIncrease)), Val("&h" & cLabelStore("Shadow").Element), cLabelStore("Shadow").VR, la.Shadow
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ShowAnchor").Group) + iIncrease)), Val("&h" & cLabelStore("ShowAnchor").Element), cLabelStore("ShowAnchor").VR, la.ShowAnchor
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("ShowTextBox").Group) + iIncrease)), Val("&h" & cLabelStore("ShowTextBox").Element), cLabelStore("ShowTextBox").VR, la.ShowTextBox
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Tag").Group) + iIncrease)), Val("&h" & cLabelStore("Tag").Element), cLabelStore("Tag").VR, la.Tag
        
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Text").Group) + iIncrease)), Val("&h" & cLabelStore("Text").Element), cLabelStore("Text").VR, la.Text
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Transparent").Group) + iIncrease)), Val("&h" & cLabelStore("Transparent").Element), cLabelStore("Transparent").VR, la.Transparent
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Visible").Group) + iIncrease)), Val("&h" & cLabelStore("Visible").Element), cLabelStore("Visible").VR, la.Visible
        ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("XOR").Group) + iIncrease)), Val("&h" & cLabelStore("XOR").Element), cLabelStore("XOR").VR, la.XOR
        
        '��Ҫ���⴦�������
        'Points����
        lngPointsCount = UBound(la.Points)
        If lngPointsCount = 0 Then
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element), cLabelStore("Points").VR, 0
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element), cLabelStore("PointsCount").VR, lngPointsCount
        Else
            vPoints = la.Points
            strPoints = vPoints(1)
            For j = 2 To lngPointsCount
                strPoints = strPoints & ";" & vPoints(j)
            Next
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element), cLabelStore("Points").VR, strPoints
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element), cLabelStore("PointsCount").VR, lngPointsCount
        End If
        
        'TagObject����
        If la.TagObject Is Nothing Then
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + iIncrease)), Val("&h" & cLabelStore("TagObject").Element), cLabelStore("TagObject").VR, 0
        Else
            lngTemp = img.Labels.IndexOf(la.TagObject)
            aSaveTagObject(i) = iIncrease + 1
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + iIncrease)), Val("&h" & cLabelStore("TagObject").Element), cLabelStore("TagObject").VR, lngTemp
        End If
        
        '��һ����ע��ӵ����ݼ���
        dssAll.Add ds
        iIncrease = iIncrease + 1
    Next
    '��������ı�ע����
    'TagObject����
    For i = 1 To dssAll.Count
            Set ds = dssAll(i)
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + i - 1)), Val("&h" & cLabelStore("TagObject").Element)).Value
            lngTemp = v(1)
            ds.Attributes.AddExplicit Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + i - 1)), Val("&h" & cLabelStore("TagObject").Element), cLabelStore("TagObject").VR, aSaveTagObject(lngTemp)
    Next
    If dssAll.Count > 0 Then    '��ͼ��������ӱ�ע
        img.Attributes.AddExplicit Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element), cLabelStore("TPRODUCER").VR, cProducer
        img.Attributes.AddExplicit Val("&h" & cLabelStore("TSUM").Group), Val("&h" & cLabelStore("TSUM").Element), cLabelStore("TSUM").VR, iIncrease
        img.Attributes.AddExplicit Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element), cLabelStore("TALL").VR, dssAll
    Else                        '��ͼ��û�б�ע����ԭ�б�ע���
        img.Attributes.Remove Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element)
        img.Attributes.Remove Val("&h" & cLabelStore("TSUM").Group), Val("&h" & cLabelStore("TSUM").Element)
        img.Attributes.Remove Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element)
    End If
End Sub


Public Sub subReadLabelFromImg(img As DicomImage)
'------------------------------------------------
'���ܣ���ͼ���ͷ�ļ��ж�ȡ��ע������ʾ��ע
'������img--��Ҫ��ȡ��ע��ͼ��
'���أ��ޣ�ֱ�ӽ�ͼ���еı�ע��ȡ��������������ʾ��
'2009��
'------------------------------------------------
    Dim ds As DicomDataSet
    Dim dss As DicomDataSets
    Dim las As New DicomLabels
    Dim la As DicomLabel
    Dim v As Variant
    Dim i As Integer
    Dim iCount As Integer
    Dim lngTemp As Long
    Dim aTagObject() As Long
    Dim iIncrease As Integer
    Dim strPoints As String
    Dim lngPointsCount As Long
    Dim aReadTagObject() As Long
    Dim j As Long
    Dim strX As String
    Dim strY As String
    Dim iOldCount As Integer
    ReDim aReadTagObject(cLabelStore.Count) As Long
    aReadTagObject(0) = 0
    
    If (img.Attributes(Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element)).Exists) Then
        v = img.Attributes(Val("&h" & cLabelStore("TPRODUCER").Group), Val("&h" & cLabelStore("TPRODUCER").Element)).Value
        If IsNull(v) Or v <> cProducer Then
            Exit Sub
        End If
        
        If IsNull(img.Attributes(Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element))) Then
            Exit Sub
        End If
        Set dss = img.Attributes(Val("&h" & cLabelStore("TALL").Group), Val("&h" & cLabelStore("TALL").Element)).Value
        iCount = dss.Count
        iIncrease = 0
        For i = 1 To iCount
            Set ds = dss(i)
            Set la = New DicomLabel
                        
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Left").Group) + iIncrease)), Val("&h" & cLabelStore("Left").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Left").Group) + iIncrease)), Val("&h" & cLabelStore("Left").Element))
            la.left = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Top").Group) + iIncrease)), Val("&h" & cLabelStore("Top").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Top").Group) + iIncrease)), Val("&h" & cLabelStore("Top").Element))
            la.top = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Width").Group) + iIncrease)), Val("&h" & cLabelStore("Width").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Width").Group) + iIncrease)), Val("&h" & cLabelStore("Width").Element))
            la.width = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Height").Group) + iIncrease)), Val("&h" & cLabelStore("Height").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Height").Group) + iIncrease)), Val("&h" & cLabelStore("Height").Element))
            la.height = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LabelType").Group) + iIncrease)), Val("&h" & cLabelStore("LabelType").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LabelType").Group) + iIncrease)), Val("&h" & cLabelStore("LabelType").Element))
            la.LabelType = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("ImageTied").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("ImageTied").Element))
            la.ImageTied = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Alignment").Group) + iIncrease)), Val("&h" & cLabelStore("Alignment").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Alignment").Group) + iIncrease)), Val("&h" & cLabelStore("Alignment").Element))
            la.Alignment = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorImageTied").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorImageTied").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorImageTied").Element))
            la.AnchorImageTied = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorX").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorX").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorX").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorX").Element))
            la.AnchorX = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorY").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorY").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AnchorY").Group) + iIncrease)), Val("&h" & cLabelStore("AnchorY").Element))
            la.AnchorY = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Angle").Group) + iIncrease)), Val("&h" & cLabelStore("Angle").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Angle").Group) + iIncrease)), Val("&h" & cLabelStore("Angle").Element))
            la.Angle = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AutoSize").Group) + iIncrease)), Val("&h" & cLabelStore("AutoSize").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("AutoSize").Group) + iIncrease)), Val("&h" & cLabelStore("AutoSize").Element))
            la.AutoSize = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackColour").Group) + iIncrease)), Val("&h" & cLabelStore("BackColour").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackColour").Group) + iIncrease)), Val("&h" & cLabelStore("BackColour").Element))
            la.BackColour = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackStyle").Group) + iIncrease)), Val("&h" & cLabelStore("BackStyle").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("BackStyle").Group) + iIncrease)), Val("&h" & cLabelStore("BackStyle").Element))
            la.BackStyle = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontName").Group) + iIncrease)), Val("&h" & cLabelStore("FontName").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontName").Group) + iIncrease)), Val("&h" & cLabelStore("FontName").Element))
            la.FontName = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontSize").Group) + iIncrease)), Val("&h" & cLabelStore("FontSize").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("FontSize").Group) + iIncrease)), Val("&h" & cLabelStore("FontSize").Element))
            la.FontSize = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ForeColour").Group) + iIncrease)), Val("&h" & cLabelStore("ForeColour").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ForeColour").Group) + iIncrease)), Val("&h" & cLabelStore("ForeColour").Element))
            la.ForeColour = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineStyle").Group) + iIncrease)), Val("&h" & cLabelStore("LineStyle").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineStyle").Group) + iIncrease)), Val("&h" & cLabelStore("LineStyle").Element))
            la.LineStyle = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineWidth").Group) + iIncrease)), Val("&h" & cLabelStore("LineWidth").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("LineWidth").Group) + iIncrease)), Val("&h" & cLabelStore("LineWidth").Element))
            la.LineWidth = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Margin").Group) + iIncrease)), Val("&h" & cLabelStore("Margin").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Margin").Group) + iIncrease)), Val("&h" & cLabelStore("Margin").Element))
            la.Margin = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Outline").Group) + iIncrease)), Val("&h" & cLabelStore("Outline").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Outline").Group) + iIncrease)), Val("&h" & cLabelStore("Outline").Element))
            la.Outline = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("RotateTextWithImage").Group) + iIncrease)), Val("&h" & cLabelStore("RotateTextWithImage").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("RotateTextWithImage").Group) + iIncrease)), Val("&h" & cLabelStore("RotateTextWithImage").Element))
            la.RotateTextWithImage = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleFontSize").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleFontSize").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleFontSize").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleFontSize").Element))
            la.ScaleFontSize = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleWithCell").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleWithCell").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ScaleWithCell").Group) + iIncrease)), Val("&h" & cLabelStore("ScaleWithCell").Element))
            la.ScaleWithCell = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Shadow").Group) + iIncrease)), Val("&h" & cLabelStore("Shadow").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Shadow").Group) + iIncrease)), Val("&h" & cLabelStore("Shadow").Element))
            la.Shadow = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowAnchor").Group) + iIncrease)), Val("&h" & cLabelStore("ShowAnchor").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowAnchor").Group) + iIncrease)), Val("&h" & cLabelStore("ShowAnchor").Element))
            la.ShowAnchor = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowTextBox").Group) + iIncrease)), Val("&h" & cLabelStore("ShowTextBox").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("ShowTextBox").Group) + iIncrease)), Val("&h" & cLabelStore("ShowTextBox").Element))
            la.ShowTextBox = v(1)
            
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Tag").Group) + iIncrease)), Val("&h" & cLabelStore("Tag").Element))
            If IsNull(v) Then
                la.Tag = ""
            Else
                la.Tag = v(1)
            End If
            
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Text").Group) + iIncrease)), Val("&h" & cLabelStore("Text").Element))
            If IsNull(v) Then
                la.Text = ""
            Else
                la.Text = v(1)
            End If
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Transparent").Group) + iIncrease)), Val("&h" & cLabelStore("Transparent").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Transparent").Group) + iIncrease)), Val("&h" & cLabelStore("Transparent").Element))
            la.Transparent = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Visible").Group) + iIncrease)), Val("&h" & cLabelStore("Visible").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Visible").Group) + iIncrease)), Val("&h" & cLabelStore("Visible").Element))
            la.Visible = v(1)
            
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("XOR").Group) + iIncrease)), Val("&h" & cLabelStore("XOR").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("XOR").Group) + iIncrease)), Val("&h" & cLabelStore("XOR").Element))
            la.XOR = v(1)
            
            '��Ҫ���⴦�������
            'Points����
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("Points").Group) + iIncrease)), Val("&h" & cLabelStore("Points").Element))
            
            strPoints = v(1)
            If IsNull(ds.Attributes(Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element))) Then Exit Sub
            v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("PointsCount").Group) + iIncrease)), Val("&h" & cLabelStore("PointsCount").Element))
            lngPointsCount = v(1) / 2
            For j = 1 To lngPointsCount - 1
                strX = left(strPoints, InStr(strPoints, ";") - 1)
                strPoints = Right(strPoints, Len(strPoints) - InStr(strPoints, ";"))
                strY = left(strPoints, InStr(strPoints, ";") - 1)
                strPoints = Right(strPoints, Len(strPoints) - InStr(strPoints, ";"))
                la.AddPoint Val(strX), Val(strY)
            Next
            If lngPointsCount > 0 Then
                strX = left(strPoints, InStr(strPoints, ";") - 1)
                strPoints = Right(strPoints, Len(strPoints) - InStr(strPoints, ";"))
                strY = strPoints
                la.AddPoint Val(strX), Val(strY)
            End If
            
            las.Add la
            iIncrease = iIncrease + 1
        Next
    End If
    '����ע�ŵ�ͼ������
    '�Ƚ�ͼ��ԭ���ı�ע����������
    iOldCount = img.Labels.Count
    For i = 1 To las.Count
        img.Labels.Add las(i)
    Next
    
    '�����������������
    '����TagObject
    For i = 1 To las.Count
        Set ds = dss(i)
        Set la = img.Labels(iOldCount + i)
        v = ds.Attributes(Val("&h" & CStr(Val(cLabelStore("TagObject").Group) + i - 1)), Val("&h" & cLabelStore("TagObject").Element))
        
        lngTemp = v(1)
        If lngTemp = 0 Then
            Set la.TagObject = Nothing
        Else
            la.TagObject = img.Labels(iOldCount + lngTemp)
        End If
    Next
End Sub

Public Sub subDrawRefLine(imgSource As DicomImage, imgDest As DicomImage, blnCheckSpacing As Boolean, _
    strLineTag As String, blnShowNum As Boolean)
'------------------------------------------------
'���ܣ�����λ��
'������ imgSource--��λ�ߵ�ͶӰͼ
'       imgDest -- ��λ�����ڵ�ͼ��
'       blnCheckSpacing -- �Ƿ��ⶨλ��֮��ľ���
'       strLineTag -- ��λ�ߵ�Tag������
'       blnShowNum -- �Ƿ���ʾ����
'���أ���
'2009��
'------------------------------------------------
    Dim l As DicomLabel
    Dim dlNum As DicomLabel
    Dim iXoffset As Integer, iYoffset As Integer
    Dim strIOPSource As String
    Dim strIOPDest As String

    '��0020,0052���ж�Frame of Reference UID�Ƿ���ͬ��ֻ�ܶԲο�֡UID��ͬ��ͼ������λ�߲���
    If Not IsNull(imgDest.Attributes(&H20, &H52).Value) And Not IsNull(imgSource.Attributes(&H20, &H52).Value) Then
        If imgDest.Attributes(&H20, &H52).Value = imgSource.Attributes(&H20, &H52).Value Then
            
            '��ͬһ�������ͼ������λ��
            If imgSource.Attributes(&H20, &H37).VM = 6 And imgDest.Attributes(&H20, &H37).VM = 6 Then
                strIOPSource = CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(1)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(2)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(3)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(4)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(5)) & "," _
                                & CInt(imgSource.Attributes(&H20, &H37).ValueByIndex(6))
                strIOPDest = CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(1)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(2)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(3)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(4)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(5)) & "," _
                                & CInt(imgDest.Attributes(&H20, &H37).ValueByIndex(6))
                If strIOPSource <> strIOPDest Then
                    Set l = imgDest.ReferenceLine(imgSource, True)
                    If l.LabelType = 3 Then
                        
                        If blnCheckSpacing = True Then
                            '�жϵ�ǰ��λ�ߺ���һ����λ��֮��ľ����Ƿ�С�ڶ�λ�߼�࣬���С�ڣ�����ʾ
                            If imgDest.Labels(imgDest.Labels.Count - 1).Tag <> "RLL" _
                                Or Abs(imgDest.Labels(imgDest.Labels.Count - 1).left - l.left) >= lngReferenceLineSpacing _
                                Or Abs(imgDest.Labels(imgDest.Labels.Count - 1).top - l.top) >= lngReferenceLineSpacing Then
                                '���Ի���λ�ߣ����˳�
                            Else
                                Exit Sub
                            End If
                        End If
                        
                        l.ForeColour = lngReferenceLineColor
                        l.Tag = strLineTag
                        l.LineStyle = lngReferenceLineStyle
                        imgDest.Labels.Add l
                        
                        If blnShowNum Then
                            Set dlNum = New DicomLabel
                            If Abs(l.width) > Abs(l.height) Then
                                iXoffset = 10
                                iYoffset = 0
                            Else
                                iXoffset = 0
                                iYoffset = 20
                            End If
                            dlNum.left = IIf(l.width > 0, l.left - iXoffset, l.left + iXoffset)
                            If dlNum.left < 0 Then
                                dlNum.left = 0
                            ElseIf dlNum.left > imgDest.sizex Then
                                dlNum.left = imgDest.sizex
                            End If
                            dlNum.top = IIf(l.height > 0, l.top - iYoffset, l.top + iYoffset)
                            If dlNum.top < 0 Then
                                dlNum.top = 0
                            ElseIf dlNum.top > imgDest.sizey Then
                                dlNum.top = imgDest.sizey
                            End If
                            
                            dlNum.LabelType = doLabelText
                            dlNum.Tag = strLineTag
                            dlNum.ForeColour = lngReferenceLineColor
                            dlNum.Text = IIf(Not IsNull(imgSource.Attributes(&H20, &H13).Value), imgSource.Attributes(&H20, &H13).Value, "")
                            dlNum.ImageTied = True
                            dlNum.FontSize = 12
                            imgDest.Labels.Add dlNum
                        End If
                    End If
                
                End If
            End If
        End If
    End If
End Sub


Public Function funDrawVas(lblLine As DicomLabel, img As DicomImage, intVasType As Integer) As Boolean
'------------------------------------------------
'���ܣ�����lblLine���Զ�Ѫ�ܲ���
'������lblLine--����Ѫ�ܲ�����Ѫ�ܴ�ֱ�ߣ�img--����Ѫ�ܲ�����ͼ��intVasType--Ѫ�ܲ������ͣ�1Ϊ����Ѫ�ܣ�2Ϊ��խѪ�ܡ�
'���أ���
'2009��
'------------------------------------------------
    '����Ѫ�ܱ�
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    Dim lngRadius As Long
    Dim lngArea As Long
    Dim lblVas1 As DicomLabel
    Dim lblVas2 As DicomLabel
    Dim lblText As DicomLabel
    
    If lblLine.TagObject Is Nothing Or lblLine.TagObject.TagObject Is Nothing _
        Or lblLine.TagObject.TagObject.TagObject Is Nothing Then
       Exit Function
    End If
    Set lblText = lblLine.TagObject
    Set lblVas1 = lblText.TagObject
    Set lblVas2 = lblVas1.TagObject
    If funGetVasEdge(img, lblLine, IIf(intVasType = 1, intStandardThreshold, intNarrowThreshold), x1, y1, x2, y2) = True Then
        '����Ѫ�ܱڶ�ֱ��
        
        subDrawVasEdgeLine lblLine, lblVas1, x1, y1
        lblVas1.Text = x1 & "," & y1
        subDrawVasEdgeLine lblLine, lblVas2, x2, y2
        lblVas2.Text = x2 & "," & y2
        lngRadius = Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
        lngArea = 3.14 * lngRadius * lngRadius / 4
        lblText.Text = IIf(intVasType = 1, "����Ѫ��ֱ����", "��խѪ��ֱ����") & lngRadius & _
                        "(" & lblLine.ROIDistanceUnits & ")" & vbCrLf & "Ѫ�������" _
                        & lngArea & "(sq " & lblLine.ROIDistanceUnits & ")"
        lblLine.Text = lngRadius & ":" & IIf(intVasType = 1, intStandardThreshold, intNarrowThreshold)
        funDrawVas = True
    End If
End Function
                       
Public Sub subChangeLabelForPrint(img As DicomImage, intType As Integer)
'------------------------------------------------
'���ܣ��޸�ͼ�����ĽǱ�ע����λ��ע������λ��ע�ɸ�ͼ��һ�����ţ�Ϊ��Ƭ��ӡ��׼��
'������img������Ҫ�޸ı�ע��ͼ��,intType -- 0������ʾ��1���ڴ�ӡ
'���أ���
'------------------------------------------------
    Dim dlLabel As DicomLabel
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strImageType As String
    Dim strImageFontSize As Integer
    Dim strImageAutoZoom As Boolean
    Dim strPostureFontSize As Integer
    Dim strPostureAutoZoom As Boolean
    Dim blnFontInverse As Boolean
    Dim blnFontShadow As Boolean
    Dim blnFontTransparent As Boolean
    
    strImageType = IIf(IsNull(img.Attributes(&H8, &H60).Value), "OT", img.Attributes(&H8, &H60).Value)
    
    If blLocalRun = True Then
        strSQL = "select Ӱ�����,�����С,�Ƿ���ͼ������,��λ��ע�����С,��λ��ע��ͼ������ from Ӱ��Ƭ��ӡ���� where Ӱ����� = '" & strImageType & "'"
        Set rsTmp = cnAccess.Execute(strSQL)
    Else
        strSQL = "select Ӱ�����,�����С,�Ƿ���ͼ������,��λ��ע�����С,��λ��ע��ͼ������,���巴ɫ,������Ӱ,���屳��͸�� from Ӱ��Ƭ��ӡ���� where Ӱ����� = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strImageType)
    End If
    
    If rsTmp.EOF = True Then
        '����Ĭ�����壺�Ľ���Ϣ��ע����ͨ������ע
        Select Case strImageType
            Case "CR"
                strImageFontSize = 14
                strImageAutoZoom = True
            Case "CT"
                strImageFontSize = 20
                strImageAutoZoom = True
            Case "MR"
                strImageFontSize = 8
                strImageAutoZoom = False
            Case "RF"
                strImageFontSize = 18
                strImageAutoZoom = False
            Case Else
                strImageFontSize = 14
                strImageAutoZoom = True
        End Select
        
        '����Ĭ�����壺��λ��ע
        If img.sizex > 1024 Then
            strPostureFontSize = 40
        ElseIf img.sizex > 512 Then
            strPostureFontSize = 25
        ElseIf img.sizex > 400 Then
            strPostureFontSize = 18
        Else
            strPostureFontSize = 10
        End If
        strPostureAutoZoom = True
        
        blnFontInverse = False
        blnFontShadow = False
        blnFontTransparent = True
    Else
        strImageFontSize = NVL(rsTmp("�����С"), 14)
        strImageAutoZoom = NVL(rsTmp("�Ƿ���ͼ������"), "True")
        strPostureFontSize = NVL(rsTmp("��λ��ע�����С"), 25)
        strPostureAutoZoom = NVL(rsTmp("��λ��ע��ͼ������"), "True")
        blnFontInverse = NVL(rsTmp("���巴ɫ"), "False")
        blnFontShadow = NVL(rsTmp("������Ӱ"), "False")
        blnFontTransparent = NVL(rsTmp("���屳��͸��"), "True")
    End If
    
    
    For Each dlLabel In img.Labels
        '�����±�ע����Ҫ���������ִ�С
        '1������ΪdoLabelSpecial���ı���λ��ע��
        '2��Tag =��PAT��Ϊ�����Ľ���Ϣ��
        '3������λ��ע
        '4�����
        '5���û��Լ����ı�ע���Ҳ�������λ��ע
        
        If dlLabel.LabelType = doLabelSpecial Or Mid(dlLabel.Tag, 1, 3) = "PAT" Or _
            img.Labels.IndexOf(dlLabel) = G_INT_SYS_LABEL_WWWL Or _
            (img.Labels.IndexOf(dlLabel) >= G_INT_SYS_LABEL_RULLER And img.Labels.IndexOf(dlLabel) <= G_INT_SYS_LABEL_RULLER + 5) Or _
            img.Labels.IndexOf(dlLabel) > G_INT_SYS_LABEL_COUNT Then
            
            If intType = 0 Then     '������ʾ
                dlLabel.ScaleFontSize = True
                dlLabel.ForeColour = vbWhite
                dlLabel.Shadow = doShadowNone
                dlLabel.Transparent = True
                dlLabel.XOR = False
            ElseIf intType = 1 Then     '���ڴ�ӡ
                If InStr(dlLabel.Tag, POSTURE_LABEL) = 0 Then
                    '�����ĽǱ�ע����ͨ������ע�����С
                    dlLabel.FontSize = strImageFontSize
                    dlLabel.ScaleFontSize = strImageAutoZoom
                    
                Else
                    '������λ��ע�����С
                    dlLabel.FontSize = strPostureFontSize
                    dlLabel.ScaleFontSize = strPostureAutoZoom
                End If
                If blnFontShadow = True Then
                    dlLabel.Shadow = doShadowTopLeft
                End If
                If blnFontTransparent = False Then
                    dlLabel.BackColour = vbBlack
                    dlLabel.Transparent = False
                End If
                If blnFontInverse = True Then
                    dlLabel.XOR = True
                End If
            End If
        Else
            If intType = 0 Then
                dlLabel.ScaleFontSize = True
            End If
        End If
    Next
    
    '���ص�ǰ��ʾ�ı�עѡ����
    For i = 11 To 20
        img.Labels(i).Visible = False
    Next i
        
    '���ش�ӡ���
    img.Labels(G_INT_SYS_LABEL_PRINT_TAG).Visible = False
    
    '���ͼ����CT�������ش���λ
    If UCase(strImageType) = "CT" Then
        img.Labels(G_INT_SYS_LABEL_WWWL).Visible = True
    Else
        img.Labels(G_INT_SYS_LABEL_WWWL).Visible = False
    End If
End Sub

Public Sub funcGetCadioThoracicRatio(thisLabel As DicomLabel, thisImage As DicomImage)
'���㲢�ҷ������رȣ�����ֵ�ĸ�ʽ��"0.xx"
'������ thisLabel---���رȵĲ�����ע
'       thisImage---�������رȵ�ͼ��
'���رȲ����ı�ע�ǡ�CTR1L��+��CTR1T��+��CTR2L��+��CTR2T�����ĸ���ע���������ġ�
'thisLabelָ��CTR1L�����ߡ�CTR2L��
    
    If thisLabel Is Nothing Then Exit Sub
    If thisLabel.TagObject Is Nothing Then Exit Sub
    If thisLabel.TagObject.TagObject Is Nothing Then Exit Sub
    If thisLabel.TagObject.TagObject.TagObject Is Nothing Then Exit Sub
    
    Dim intLine1 As Integer
    Dim intLine2 As Integer
    Dim otherLabel As DicomLabel
    
    On Error GoTo err
        
    If thisImage.RotateState = doRotateLeft Or thisImage.RotateState = doRotateRight Then
        intLine1 = Abs(thisLabel.height)
        intLine2 = Abs(thisLabel.TagObject.TagObject.height)
    Else
        intLine1 = Abs(thisLabel.width)
        intLine2 = Abs(thisLabel.TagObject.TagObject.width)
    End If
        
    If intLine1 = 0 Or intLine2 = 0 Then Exit Sub
        
    Set otherLabel = thisLabel.TagObject.TagObject
    If thisLabel.Tag = "CTR1L" Then     'intLine1��������
        thisLabel.TagObject.Text = funROIResultString(thisLabel, thisImage)
        otherLabel.TagObject.Text = funROIResultString(otherLabel, thisImage) & vbCrLf _
            & "���رȣ� " & Format(intLine1 / intLine2, "0.00")
    ElseIf thisLabel.Tag = "CTR2L" Then 'intLine1��������
        thisLabel.TagObject.Text = funROIResultString(thisLabel, thisImage) & vbCrLf _
            & "���رȣ� " & Format(intLine2 / intLine1, "0.00")
        otherLabel.TagObject.Text = funROIResultString(otherLabel, thisImage)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub ShowOverlay(f As frmViewer)
'------------------------------------------------
'���ܣ���ʾ��������Overlay��Ϣ
'������ f - ��Ƭ������
'���أ���
'------------------------------------------------
    Dim v As DicomViewer
    Dim img As DicomImage
    
    On Error GoTo err
    
    For Each v In f.Viewer
        If v.Index <> 0 Then
            For Each img In v.Images
                If img.Attributes(&H6000, &H10).Exists = True Then
                    img.OverlayVisible(0) = Button_miShowOverlay
                End If
            Next
        End If
        v.Refresh
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
