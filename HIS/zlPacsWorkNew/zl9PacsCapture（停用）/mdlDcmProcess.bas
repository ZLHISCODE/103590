Attribute VB_Name = "mdlDcmProcess"
Option Explicit

'����������
Public Type TPoint
  X As Integer
  Y As Integer
End Type


Public Sub subSetSharp(objDcmImage As DicomImage, blnSharp As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ���ƽ������
'������blnSharp��ʾͼ����ķ���True=�񻯣�False=ƽ��
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If Not objDcmImage Is Nothing Then
        If blnSharp = True Then
            '�񻯴���
            If objDcmImage.FilterLength <= 0 Then
                objDcmImage.FilterLength = 0
                '��ǰû��ƽ������ֱ�ӽ����񻯴���
                objDcmImage.UnsharpEnhancement = objDcmImage.UnsharpEnhancement + 0.1
            Else
                '�����ǰ�Ѿ���ƽ���������ȵ���ƽ��Ч��
                objDcmImage.FilterLength = objDcmImage.FilterLength - 1
            End If
        Else
            'ƽ������
            '�ж�Zoom�Ƿ�1������ǣ����޸�Ϊ0.9999
            If objDcmImage.ActualZoom = 1 Then
                objDcmImage.Zoom = 0.9999
            End If
            
            If objDcmImage.UnsharpEnhancement <= 0 Then
                objDcmImage.UnsharpEnhancement = 0
                '��ǰû���񻯴���ֱ�ӿ�ʼƽ��
                '�ж�FilterLength�Ƿ�0����ǣ�����2/ActualZoom��2��FilterLength֮����е���
                If objDcmImage.FilterLength = 0 Then
                    objDcmImage.FilterLength = 2 / objDcmImage.ActualZoom + 1
                Else    '���������FilterLength��1
                    objDcmImage.FilterLength = objDcmImage.FilterLength + 1
                End If
            Else
                '��ǰ�Ѿ������񻯴����ȵ����񻯵�Ч��
                objDcmImage.UnsharpEnhancement = objDcmImage.UnsharpEnhancement - 0.1
            End If
        End If
    End If
End Sub



Public Sub subSetRotate(objDcmImage As DicomImage, blnClockwise As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ�����ת
'������blnClockwise��ת�ķ���,True=˳ʱ����ת��False=��ʱ����ת
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If Not objDcmImage Is Nothing Then
        Dim iRotateState As Integer
        
        iRotateState = objDcmImage.RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        
        If iRotateState = -1 Then iRotateState = 3
        
        iRotateState = iRotateState Mod 4
        objDcmImage.RotateState = iRotateState
    End If
End Sub


'DicomViewer�ü���ɼ�ͼ��
Public Function CutImage(objDcmImage As DicomImage) As DicomImage
    Dim imgResult As DicomImage
    Dim imgs As New DicomImages
    
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim iMax As Integer
    Dim img As DicomImage
    Dim lblFrame As DicomLabel
    
    Set CutImage = Nothing
    
    If objDcmImage Is Nothing Then Exit Function
    If objDcmImage.Labels.Count < 1 Then Exit Function
    
    Set img = objDcmImage
    Set lblFrame = objDcmImage.Labels(objDcmImage.Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgboxCus "��ѡ��ͼ��������ٱ���", vbExclamation, G_STR_HINT_TITLE
        Exit Function
    End If
    
    'ͼ�������=300
    iMax = 300
    
    '����label����ȡ����ѡ�е�ͼ��
    'ͼ��λ��,�ڰ�ͼ��Ϊ1����ɫͼ��Ϊ3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).value = "RGB" Then
            iPlane = 3
        End If
    End If
    
    'ͼ����λ��
    If lblFrame.Width >= 0 Then
        iLeft = lblFrame.Left
        iRight = iLeft + lblFrame.Width
    Else
        iLeft = lblFrame.Left + lblFrame.Width
        iRight = lblFrame.Left
    End If
    
    If lblFrame.Height >= 0 Then
        iTop = lblFrame.Top
        iBottom = iTop + lblFrame.Height
    Else
        iTop = lblFrame.Top + lblFrame.Height
        iBottom = lblFrame.Top
    End If
    
    '���ƽ��ͼ��Ĵ�С��300*300֮��
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.SizeY - iBottom, img.SizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X��Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, img.SizeY - iBottom, img.SizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    Set CutImage = imgResult
End Function


Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'���ܣ�����һ��LABEL���󣬲���������ʼ����
'������lType--��ע�����ͣ�lLeft--��ע��Leftֵ��lTop--��ע��Topֵ��lWidth--��ע��Widthֵ��lHeight--��ע��Heightֵ��
'���أ������ɵı�ע��
'�����ˣ��ƽ�
'------------------------------------------------
    Dim l As New DicomLabel
    
    l.LabelType = lType
    l.XOR = True
    l.ImageTied = True
    l.Left = lLeft
    l.Top = lTop
    l.Width = lWidth
    l.Height = lHeight
    l.Margin = 0
    l.AutoSize = True
    l.FontSize = 12
    l.LineWidth = 1
    
    If l.LabelType = 0 Then     '����
        l.Transparent = False
        l.Width = 200
        l.Height = 10
    End If
    
    Set GetNewLabel = l
End Function
   
   
Public Sub subCenterZoom(frmWindow As Form, img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'���ܣ���ͼ��������š��Ե�ǰviewer���ĵ�Ϊ�������ĵ㡣
'������
'       img -- �������ŵ�ͼ��
'       viewer ���� ͼ�����ڵ�viewer
'       dblZoom ����ͼ���µ����ű���
'���أ��ޣ�ֱ�ӵ���ͼ������ű���
'�ϼ���������̣�frmViewer.Viewer_MouseMove
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ� �ƽ� 2006-2-10
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False
            
    img.ScrollX = (img.SizeX * img.ActualZoom - frmWindow.ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - frmWindow.ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub


Public Sub RectangleZoom(Viewer As DicomViewer, img As DicomImage, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
    Dim newZoom As Double
    Dim dblRatio As Double
    Dim sX As Long
    Dim sY As Long
    Dim oldZoom As Double
    
    If lngWidth > 0 And lngHeight > 0 Then
        oldZoom = img.ActualZoom
        sX = img.ActualScrollX
        sY = img.ActualScrollY
        
        img.StretchToFit = False
        
        dblRatio = Viewer.Width / Screen.TwipsPerPixelX / lngWidth
        If dblRatio > Viewer.Height / Screen.TwipsPerPixelY / lngHeight Then
            dblRatio = Viewer.Height / Screen.TwipsPerPixelY / lngHeight
        End If
        
        newZoom = oldZoom * dblRatio
        img.Zoom = newZoom
        
        img.ScrollX = (sX + lngLeft) * dblRatio
        img.ScrollY = (sY + lngTop) * dblRatio
    End If
End Sub
