Attribute VB_Name = "MdlSerial"
Option Explicit

'--------------------------------------------------------
'��  �ܣ���ģ��Ϊ����������ݴ���
'�������ڣ�2004.6
'���̺����嵥��
'funSliceLocation   ():��һ��Viewer��Ѱ����Ƭλ�����ָ��ֵ��ͼ��
'subSerialPlaceInPhase  ():����֮��λ��ͬ��
'FunImageIsX        ():�ж�һ��ͼ�����ڵ���
'FunImageIsY        ():�ж�һ��ͼ�����ڵ���
'subIsSerialXY      ():�жϵ�ǰ�����ĸ����е�λ����
'subDispframe       ():��ʾ��ǰ����ָ��viewer��ͼ�����ѡ���ǵ�
'subInitSerial      ():ɾ��ԭ���϶�������������Ӻ�������˫���϶�����'
'
'�޸ļ�¼��
'    2005.6     �ƽ�
'-------------------------------------------------------

Private Function funSliceLocation(intViewerIndex As Integer, s As Double) As Integer
'------------------------------------------------
'���ܣ���һ��Viewer��Ѱ����Ƭλ�����ָ��ֵ��ͼ��
'������v--Ѱ�������Ƭͼ���viewer��s--���бȽϵ�Ŀ����Ƭλ�á�
'���أ������ͼ����š�
'------------------------------------------------
    Dim dt As Double
    Dim i As Integer
    
    dt = intSliceOffset
    funSliceLocation = 0
    For i = 1 To ZLShowSeriesInfos(intViewerIndex).ImageInfos.Count
        '������ӽ�s��λ��
        If Abs(Val(ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceLocation) - s) < dt Then
            dt = Abs(Val(ZLShowSeriesInfos(intViewerIndex).ImageInfos(i).SliceLocation) - s)
            funSliceLocation = i
        End If
    Next
End Function

Public Sub subManualSeriesSyn(f As frmViewer, iMove As Integer, vIndex As Integer)
'------------------------------------------------
'���ܣ��ֹ����м�λ��ͬ��
'������ f--��������ͬ���Ĵ���
'       iMove--ͼ�񷭶��ķ����������������ǰ������������󷭶���
'       vIndex ���ĸ�viewer�����ֹ�����ͬ��
'���أ��ޣ�ֱ���޸�ͼ�����ʾ��
'------------------------------------------------
    Dim v As DicomViewer
    Dim i As Integer
    Dim intOldIndex As Integer
    Dim intCurrentIndex As Integer
    
    f.blnVscroInvoked = True
    For Each v In f.Viewer
        If v.Visible And v.Index <> vIndex Then
            If ZLShowSeriesInfos(v.Index).Selected = True Then
                '�������ڵ�ͼ��λ��
                intOldIndex = f.VScro(v.Index).Value
                
                '����ͬ����ͼ�����λ��
                If (intOldIndex + iMove) <= (ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1) And (intOldIndex + iMove) > 0 Then
                    intCurrentIndex = intOldIndex + iMove
                ElseIf iMove > 0 And intOldIndex < ZLShowSeriesInfos(v.Index).ImageInfos.Count Then
                    intCurrentIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1
                ElseIf iMove < 0 And intOldIndex > 1 Then
                    intCurrentIndex = 1
                ElseIf iMove > 0 And intOldIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count Then
                    intCurrentIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count
                End If
                
                If intCurrentIndex <= 0 Or intCurrentIndex > ZLShowSeriesInfos(v.Index).ImageInfos.Count Then
                    intCurrentIndex = 1
                End If
                
                If intOldIndex <> intCurrentIndex Then       'ͼ��������
                    f.MSFViewer.TextMatrix(v.Index, 3) = intCurrentIndex
                    f.VScro(v.Index).Value = intCurrentIndex
                    Call subShowALLImage(f, v, intCurrentIndex, True)
                End If
            End If
        End If
        
    Next
    f.blnVscroInvoked = False
End Sub

Public Sub subSerialPlaceInPhase(dubPlace As Double, f As frmViewer)
'------------------------------------------------
'���ܣ�����֮��λ��ͬ��
'������dubPlace--��������ͬ����Ŀ����Ƭλ�ã�f--��������֮��λ��ͬ���Ĵ��塣
'���أ��ޣ�ֱ�ӵ���ͼ��λ�á�
'------------------------------------------------
    Dim v As DicomViewer, i As Integer
    Dim m As Integer
    Dim intCurrentIndex As Integer
    
    f.blnVscroInvoked = True
    For Each v In f.Viewer
        If v.Index <> f.intSelectedSerial And v.Visible Then
            If ZLShowSeriesInfos(v.Index).Selected = True Then
            intCurrentIndex = funSliceLocation(v.Index, dubPlace)
                If intCurrentIndex <> 0 Then
                    m = intCurrentIndex - f.MSFViewer.TextMatrix(v.Index, 3)
                    f.MSFViewer.TextMatrix(v.Index, 3) = intCurrentIndex
                    '�й����������
                    If ZLShowSeriesInfos(v.Index).ImageInfos.Count > v.MultiColumns * v.MultiRows Then
                        If intCurrentIndex > ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1 Then
                            intCurrentIndex = ZLShowSeriesInfos(v.Index).ImageInfos.Count - v.MultiColumns * v.MultiRows + 1
                        End If
                        If intCurrentIndex < 1 Then intCurrentIndex = 1
                        
                        'ͼ��������
                        If f.VScro(v.Index).Value <> intCurrentIndex Then
                            f.MSFViewer.TextMatrix(v.Index, 3) = intCurrentIndex
                            f.VScro(v.Index).Value = intCurrentIndex
                            Call subShowALLImage(f, v, intCurrentIndex, True)
                        End If
                    End If
                End If
            End If
        End If
    Next
    f.blnVscroInvoked = False
End Sub

Function FunImageIsX(Index As Integer, v As DicomViewer) As Integer
'''''�ж�һ��ͼ�����ڵ���
'2009��
    FunImageIsX = Index - v.CurrentIndex + 1
    FunImageIsX = FunImageIsX Mod v.MultiColumns
    If FunImageIsX = 0 Then FunImageIsX = v.MultiColumns
End Function

Function FunImageIsY(Index As Integer, v As DicomViewer) As Integer
'''''�ж�һ��ͼ�����ڵ���
'2009��
    FunImageIsY = Index - v.CurrentIndex + 1
    FunImageIsY = Int(FunImageIsY / v.MultiColumns - 0.5) + 1
End Function


Public Sub subIsSerialXY(f As frmViewer, x, y, intSerialX As Integer, intSerialY As Integer)
'------------------------------------------------
'���ܣ��жϵ�ǰ�����ĸ����е�λ����
'������f--�����жϵĴ��壻(x,y)--��Ҫ�жϵĵ�����ꣻ
'      intSerialX--���ظõ�����x�������������intSerialY--���ظõ�����y�������������
'���أ���
'2009��
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim tx1 As Integer, tx2 As Integer, ty1 As Integer, ty2 As Integer
    intSerialX = 0
    intSerialY = 0
    For i = 1 To intMaxAreaX                          ''�������ɻ��ֵ�����
        If i = 1 Then
            tx1 = 0
        Else
            tx1 = f.PicX(i - 1).left
        End If
        If i = intMaxAreaX Then
            tx2 = f.picViewer.ScaleWidth
        Else
            tx2 = f.PicX(i).left
        End If
        For j = 1 To intMaxAreaY
            If j = 1 Then
                ty1 = 0
            Else
                ty2 = f.PicY(j - 1).top
            End If
            If j = intMaxAreaY Then
                ty2 = f.picViewer.ScaleHeight
            Else
                ty2 = f.PicY(j).top
            End If
            If x >= tx1 And x <= tx2 And y >= ty1 And y <= ty2 Then
                intSerialX = i
                intSerialY = j
                Exit Sub
            End If
        Next
    Next
End Sub

Public Sub subDispframe(f As frmViewer, v As DicomViewer)
'------------------------------------------------
'���ܣ���ʾ��ǰ����ָ��viewer��ͼ�����ѡ���ǵ�
'������f--��Ҫ��ʾviewerͼ�����Ĵ��壻v--��Ҫ��ʾͼ������viewer��
'���أ��ޣ�ֱ����ʾָ��viewer��ͼ�����
'2009��
'------------------------------------------------
    Dim l, lx, ly, lb As DicomLabel
    Dim x, y As Integer
    Dim w, h As Single
    Dim i As Integer
    Dim iTempIndex As Integer
    
    If v.Index = 0 Then Exit Sub
    If v.Images.Count = 0 Then Exit Sub
    
    On Error GoTo err
    
    '���ԭ�еı�ע
    v.Labels.Clear
    
    '����ÿһ��ͼ������Ŀ�Ⱥ͸߶�
    w = v.width / v.MultiColumns / Screen.TwipsPerPixelX
    h = v.height / v.MultiRows / Screen.TwipsPerPixelY
    
    '��ʾ��ӡ���
    For i = v.CurrentIndex To v.CurrentIndex + v.MultiColumns * v.MultiRows
        If i > v.Images.Count Then Exit For
        v.Images(i).Labels(G_INT_SYS_LABEL_PRINT_TAG).Visible = (blnShowPrintTag And ZLShowSeriesInfos(v.Index).ImageInfos(v.Images(i).Tag).blnPrinted)
    Next i
    
    'ѭ��ÿһ��ͼ�񣬻���ע
    For y = 1 To v.MultiRows
        For x = 1 To v.MultiColumns
            '���㵱ǰ��ͼ��Index
            iTempIndex = x + (y - 1) * v.MultiColumns + v.CurrentIndex - 1
            ''''''''''''''''�߿�''''''''''''''
            Set l = New DicomLabel
            l.LabelType = 2         '���ο�
            l.width = w - lngCellSpacing * 2
            l.height = h - lngCellSpacing * 2
            l.left = (x - 1) * w + lngCellSpacing
            l.top = (y - 1) * h + lngCellSpacing
            l.Tag = "L" & x + (y - 1)
            '''�ж��Ƿ�ǰѡ���ͼ��������ͼ����ο����ɫ�����ͣ��߿�
            '��ǰͼ��߿���ɫ lngCurrentImageBorderColor
            'ѡ��ͼ��߿���ɫ lngSelectedImageBorderColor
            '��ǰ��δѡ�У����б߿���ɫ lngCurrentSeriesBorderColor
            '���ж���������Ƿ�ѡ��
            If ZLShowSeriesInfos(v.Index).Selected = True Then   '��ѡ�е�����
                '���жϵ�ǰͼ���Ƿ��ǵ�ǰͼ��
                If v.Index = f.intSelectedSerial And iTempIndex = f.MSFViewer.TextMatrix(v.Index, 3) Then
                    l.ForeColour = lngCurrentImageBorderColor
                    l.LineStyle = lngCurrentImageBorderLineStyle
                    l.LineWidth = lngCurrentImageBorderLineWidth
                Else
                    l.ForeColour = lngSelectedImageBorderColor
                    l.LineStyle = lngSelectedImageBorderLineStyle
                    l.LineWidth = lngSelectedImageBorderLineWidth
                End If
            ElseIf v.Index = f.intSelectedSerial Then   'û�б�ѡ�У����ж��Ƿ��ǵ�ǰ����
                '���жϵ�ǰͼ���Ƿ��ǵ�ǰͼ��
                If iTempIndex = f.MSFViewer.TextMatrix(v.Index, 3) Then
                    l.ForeColour = lngCurrentImageBorderColor
                Else
                    l.ForeColour = lngCurrentSeriesBorderColor
                End If
                l.LineStyle = lngCurrentImageBorderLineStyle
                l.LineWidth = lngCurrentImageBorderLineWidth
            Else        '��û�б�ѡ�У�Ҳ���ǵ�ǰ���У�����ʾĬ�ϱ߿�
                l.ForeColour = lngDefaultImageBorderColor
                l.LineStyle = lngDefaultImageBorderLineStyle
                l.LineWidth = lngDefaultImageBorderLineWidth
            End If
            v.Labels.Add l
            
            '����ǵ�ǰ���еĵ�ǰͼ����Ϊ���ͼ���б�ѡ�еı�ע��ʾ8����עѡ����
            If iTempIndex <= v.Images.Count And iTempIndex > 0 Then
                If v.Images(iTempIndex).Labels(11).Visible Then
                    'Ϊָ��ͼ���е�ָ����ע����ʾ��עѡ����
                    SubDispPeriod v.Images(iTempIndex).Labels(11).TagObject, v.Images(iTempIndex), f
                End If
            End If
            
            ''''''''''''''''Ϊÿһ��ͼ������ѡ���ǣ��̺��ߣ������ߣ�ѡ����''''''''''''''
            '''''''''''''''''''����'''''''''''''''''''''''''''''''''
            Set lx = New DicomLabel
            lx.LabelType = 3            'ֱ��
            lx.width = lngImageIdentifierSize
            lx.height = 0
            lx.left = l.left + l.width - lngImageIdentifierSize
            lx.top = l.top + l.height - lngImageIdentifierSize
            lx.Tag = "X" & x + (y - 1) * v.MultiColumns
            lx.TagObject = l
            lx.ForeColour = l.ForeColour
            lx.LineStyle = l.LineStyle
            lx.LineWidth = l.LineWidth
            v.Labels.Add lx
            ''''''''''''''''����''''''''''''''
            Set ly = New DicomLabel
            ly.LabelType = 3        'ֱ��
            ly.width = 0
            ly.height = lngImageIdentifierSize
            ly.left = l.left + l.width - lngImageIdentifierSize
            ly.top = l.top + l.height - lngImageIdentifierSize
            ly.Tag = "Y" & x + (y - 1) * v.MultiColumns
            ly.TagObject = lx
            ly.ForeColour = l.ForeColour
            ly.LineStyle = l.LineStyle
            ly.LineWidth = l.LineWidth
            v.Labels.Add ly
            ''''''''''''''''ѡ����''''''''''''''
            Set lb = New DicomLabel
            lb.LabelType = 2            '����
            lb.width = lngImageIdentifierSize - IIf(l.LineWidth / 2 >= 2, l.LineWidth / 2, 2)
            lb.height = lngImageIdentifierSize - IIf(l.LineWidth / 2 >= 2, l.LineWidth / 2, 2)
            lb.left = l.left + l.width - lngImageIdentifierSize + 1
            lb.top = l.top + l.height - lngImageIdentifierSize + 1
            lb.Transparent = False
            lb.ForeColour = lngSelectImageForeColour
            lb.BackColour = lngSelectImageForeColour
            lb.Tag = "B" & x + (y - 1) * v.MultiColumns
            lb.TagObject = ly
            lb.Visible = False
            v.Labels.Add lb
            l.TagObject = lb
            If (iTempIndex <= v.Images.Count) Then
                If v.Images.Count > 0 And (iTempIndex) > 0 Then
                    lb.Visible = ZLShowSeriesInfos(v.Index).ImageInfos(v.Images(iTempIndex).Tag).blnSelected
                    If iTempIndex <> v.Images(iTempIndex).Tag Then
                        Debug.Print "ii"
                    End If
                Else
                    lb.Visible = False
                End If
            End If
            If (iTempIndex = v.Images.Count) And Not blnDsipSpilthBorder Then Exit Sub
        Next
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub subDispWWWL(img As DicomImage)
'------------------------------------------------
'���ܣ�����Ԥ��Ĵ���λֵ��ʾͼ��
'������img--��Ҫ������ʾ��ͼ��
'���أ��ޣ�ֱ���޸�ͼ��Ĵ���λֵ���ʹ���λ��ע
'�ϼ���������̣�
'�¼���������̣���
'���õ��ⲿ������aPresetWinWL
'�����ˣ��ƽ�
'------------------------------------------------
    Dim strDriverType As String
    Dim intModality As Integer
    Dim i As Integer
    Dim im As DicomImage
    
    If IsNull(img.Attributes(&H8, &H60).Value) Then Exit Sub         '��ȡModality
    strDriverType = img.Attributes(&H8, &H60).Value
    
    For i = 1 To UBound(aPresetWinWL, 2)
        If UCase(aPresetWinWL(3, i).strModality) = UCase(strDriverType) Then
            intModality = i
            Exit For
        End If
    Next i
    
    For i = 3 To 12
        If aPresetWinWL(i, intModality).bInUse And aPresetWinWL(i, intModality).intDefault = 1 Then
            img.width = aPresetWinWL(i, intModality).lngWinWidth
            img.Level = aPresetWinWL(i, intModality).lngWinLevel
            Exit For
        End If
    Next i
    img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
End Sub

Public Sub subScaleImage(img As DicomImage, v As DicomViewer, lngOldX As Long, lngOldY As Long)
'------------------------------------------------
'���ܣ�����ͼ���λ�ú����ű�������ͼ�񲼾������ı��ʱ�����
'������ img         --- ��Ҫ������ͼ��
'       v           --- ͼ�����ڵ��°ڷźõ�Viewer
'       lngOldX     --- ͼ��ԭ�����ڵ�Viewer�е���ͼ����ռ�õĿ��
'       lngOldY     --- ͼ��ԭ�����ڵ�Viewer�е���ͼ����ռ�õĸ߶�
'���أ���
'ʱ�䣺2009-7
'------------------------------------------------
    Dim dblScale As Double
    Dim dblNewZoom As Double
    Dim lngNewX As Long
    Dim lngNewY As Long
    Dim dblScaleX As Double
    Dim dblScaleY As Double
    Dim dblOldXY As Double
    Dim dblNewXY As Double
    
    If lngOldX = 0 Or lngOldY = 0 Then Exit Sub
    
    '������Viewer�е���ͼ����ռ�õĿ�Ⱥ͸߶�
    lngNewX = v.width / v.MultiColumns
    lngNewY = v.height / v.MultiRows
    
    '�������ŵı���
'    dblScale = lngNewX / lngOldX
'    If Abs(dblScale - 1) > Abs(lngNewY / lngOldY - 1) Then
'        dblScale = lngNewY / lngOldY
'    End If
    '�������ű���
    '���Sx��Sy������1����ȡС�ġ�
    '���Sx��Sy��С��1��ȡ��ġ�
    '���Sx����1��SyС��1����SxС��1��Sy����1��ȡOX/OY��NX/NY�б�����Ķ̱�Ϊ��׼
    dblScaleX = lngNewX / lngOldX
    dblScaleY = lngNewY / lngOldY
    If dblScaleX > 1 And dblScaleY > 1 Then
        dblScale = IIf(dblScaleX < dblScaleY, dblScaleX, dblScaleY)
    ElseIf dblScaleX < 1 And dblScaleY < 1 Then
        dblScale = IIf(dblScaleX > dblScaleY, dblScaleY, dblScaleX)
    Else
        If lngOldX >= lngOldY Then
            dblOldXY = lngOldX / lngOldY
        Else
            dblOldXY = lngOldY / lngOldX
        End If
        If lngNewX >= lngNewY Then
            dblNewXY = lngNewX / lngNewY
        Else
            dblNewXY = lngNewY / lngNewX
        End If
        
        If dblOldXY >= dblNewXY Then
            dblScale = IIf(lngOldX < lngOldY, dblScaleX, dblScaleY)
        Else
            dblScale = IIf(lngNewX < lngNewY, dblScaleX, dblScaleY)
        End If
    End If
    
    '�����µ�Zoom
    dblNewZoom = dblScale * img.ActualZoom
    
    '�ȵ���Scroll
    img.ScrollX = img.ActualScrollX - (lngNewX - lngOldX) / 2 / Screen.TwipsPerPixelX
    img.ScrollY = img.ActualScrollY - (lngNewY - lngOldY) / 2 / Screen.TwipsPerPixelY
    
    '�þ������ŵķ�ʽ����Zoom
    Call subCenterZoom(img, v, dblNewZoom)
End Sub

Public Sub subInitSerial(f As frmViewer)
'------------------------------------------------
'���ܣ�unloadԭ�еķָ�����Ȼ������load�ָ������������ǰڷŵ���ʼλ��
'������f--��Ҫ��ʼ���϶����Ĵ���
'���أ��ޣ�ֱ�ӶԴ����е��϶������г�ʼ����
'2009��
'------------------------------------------------
    Dim i, j, k As Long
    With f
        '���ԭ�е������϶�����������˫��
        For i = 1 To .PicX.Count - 1
            Unload .PicX(i)
        Next
        For i = 1 To .PicY.Count - 1
            Unload .PicY(i)
        Next
        For i = 1 To .PicXY.Count - 1
            Unload .PicXY(i)
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .PicX(0).width = intSpaceSize
        .PicX(0).height = .picViewer.height
        .PicX(0).top = 0
        .PicXX.width = intSpaceSize
        .PicXX.height = .picViewer.height
        .PicXX.top = 0
        .PicY(0).height = intSpaceSize
        .PicY(0).width = .picViewer.width
        .PicY(0).left = 0
        .PicYY.height = intSpaceSize
        .PicYY.width = .picViewer.width
        .PicYY.left = 0
        .PicXY(0).height = intSpaceSize
        .PicXY(0).width = intSpaceSize
        .PicXY(0).top = .PicY(0).top
        .PicXY(0).left = .PicX(0).left
        .PicY(0).AutoRedraw = True
        .PicX(0).AutoRedraw = True
        .PicXY(0).AutoRedraw = True
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaX - 1 ''''��ʼ��������
            load .PicX(i)
            .PicX(i).left = .picViewer.width - intSpaceSize
            .PicX(i).AutoRedraw = True
            Call zlControl.PicShowFlat(.PicX(i), 1)     '��PictureBoxģ���3Dƽ�水ť
            .PicX(i).Visible = True
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = 1 To intMaxAreaY - 1 ''''��ʼ��������
            load .PicY(i)
            .PicY(i).top = .picViewer.height - intSpaceSize
            .PicY(i).AutoRedraw = True
             Call zlControl.PicShowFlat(.PicY(i), 1)        '��PictureBoxģ���3Dƽ�水ť
            .PicY(i).Visible = True
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For i = intMaxAreaX - 1 To 1 Step -1 ''''��ʼ����������
            For j = intMaxAreaY - 1 To 1 Step -1
                k = (j - 1) * (intMaxAreaX - 1) + i
                load .PicXY(k)
                .PicXY(k).top = .PicY(j).top
                .PicXY(k).left = .PicX(i).left
                .PicXY(k).Visible = True
                .PicXY(k).ZOrder
            Next
        Next
    End With
End Sub

