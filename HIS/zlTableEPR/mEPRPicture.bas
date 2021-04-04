Attribute VB_Name = "mEPRPicture"
Option Explicit

Private Type Size
    cx As Long
    cy As Long
End Type

Public glngPen As Long          '��ǰ���ʶ���
Public glngBrush As Long        '��ǰˢ�Ӷ���

'���±���ȡֵ��API������Ӧ
Public gcurPenColor As Long     '��ǰʹ�õ�����ɫ
Public gcurPenStyle As Byte     '��ǰʹ�õ�����
Public gcurPenWidth As Byte     '��ǰʹ�õ��߿�
Public gcurFillColor As Long    '��ǰʹ�õ����ɫ
Public gcurFillStyle As Integer '��ǰʹ�õ������ʽ

'################################################################################################################
'## ���ܣ�  ����ָ��ֵ���õ�ǰ�Ļ��ʵĻ�ˢ
'##
'## ������  lngHDc          :   IN  ���༭�ؼ�
'##         PenColor        :   IN  ������ɫ
'##         PenStyle        :   IN  ��0-ʵ��,1-����,2-����,3-�㻮��,4-˫�㻮��
'##         PenWidth        :   IN  ���߿�
'##         FillColor       :   IN  �����ɫ
'##         FillStyle       :   IN  ��-1-�����,-2-ʵ��,0-ˮƽ��,1-��ֱ��,2-��б��,3-��б��,4-ˮƽ�ʹ�ֱ��,5-������
'################################################################################################################
Public Sub SetDrawStyleFromValue(lngHDc As Long, _
    PenColor As OLE_COLOR, _
    PenStyle As Byte, _
    PenWidth As Byte, _
    FillColor As OLE_COLOR, _
    FillStyle As Integer)
    
    Dim vBrush As LOGBRUSH
    Dim lngPen As Long, lngBrush As Long
    
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
    
    '����
    lngPen = CreatePen(PenStyle, IIf(PenWidth < 1, 1, PenWidth), PenColor)
    glngPen = SelectObject(lngHDc, lngPen)
    
    '��ˢ
    vBrush.lbColor = FillColor
    If FillStyle = -1 Then
        vBrush.lbStyle = BS_NULL
    ElseIf FillStyle = -2 Then
        vBrush.lbStyle = BS_SOLID
    Else
        vBrush.lbStyle = BS_HATCHED
        vBrush.lbHatch = FillStyle
    End If
    lngBrush = CreateBrushIndirect(vBrush)
    glngBrush = SelectObject(lngHDc, lngBrush)
End Sub

'################################################################################################################
'## ���ܣ�  ��������ű��浽��һ��PicMarks������
'##
'## ������  picMarksSource  ��Դ
'##         picMarksDest    ��Ŀ��
'##         zoomFactor      ����������
'################################################################################################################
Public Function ScalePicMarks(picMarksSource As cTabPicMarks, ZoomFactor As Double) As cTabPicMarks
    Dim i As Long, j As Long, T As Variant, x As Long, y As Long
    Dim strTmp As String
    Dim picMarksDest As cTabPicMarks
    
    Set picMarksDest = picMarksSource.Clone
    For i = 1 To picMarksDest.Count
        picMarksDest(i).X1 = picMarksDest(i).X1 * ZoomFactor
        picMarksDest(i).Y1 = picMarksDest(i).Y1 * ZoomFactor
        picMarksDest(i).X2 = picMarksDest(i).X2 * ZoomFactor
        picMarksDest(i).Y2 = picMarksDest(i).Y2 * ZoomFactor
        strTmp = ""
        T = Split(picMarksDest(i).�㼯, ";")
        For j = 0 To UBound(T)
            x = CLng(Split(T(j), ",")(0))
            y = CLng(Split(T(j), ",")(1))
            x = x * ZoomFactor
            y = y * ZoomFactor
            If j = 0 Then
                strTmp = CStr(x) & "," & CStr(y)
            Else
                strTmp = strTmp & ";" & CStr(x) & "," & CStr(y)
            End If
        Next
        picMarksDest(i).�㼯 = strTmp
    Next
    Set ScalePicMarks = picMarksDest
End Function

'################################################################################################################
'## ���ܣ�  ��ʾ���ͼ�������
'##
'## ������  objPic          :   IN  ����ͼ���ʣ�ͼƬ��ؼ�
'##         objMarks        :   IN  ��ͼ�α�Ǽ���
'################################################################################################################
Public Function ShowPicMark(objPic As PictureBox, objMark As cTabPicMark) As StdPicture
    Dim arrTmp() As String, arrXY() As POINTAPI
    Dim i As Integer, j As Integer, objFnt As New StdFont
    
    objPic.DrawMode = vbCopyPen
    objFnt.Name = IIf(objMark.���� = "", "����", objMark.����) 'Ŀǰ�˲�������δʹ��
            
    With objMark '������Ԫ��
        Call SetDrawStyleFromValue(objPic.hdc, .����ɫ, .����, .�߿�, .���ɫ, .��䷽ʽ)
        Select Case .����
            Case 0 '�ı�
                Call TextOut(objPic, .����, .X1, .Y1, .X2, .Y2, objFnt)
            Case 1 '����
                MoveToEx objPic.hdc, .X1, .Y1, 0
                LineTo objPic.hdc, .X2, .Y2
            Case 2 '����
                arrTmp = Split(.�㼯, ";")
                For j = 0 To UBound(arrTmp)
                    ReDim Preserve arrXY(j)
                    arrXY(j).x = CLng(Split(arrTmp(j), ",")(0))
                    arrXY(j).y = CLng(Split(arrTmp(j), ",")(1))
                Next
                Polyline objPic.hdc, arrXY(0), UBound(arrXY) + 1
            Case 3 '����
                Rectangle objPic.hdc, .X1, .Y1, .X2, .Y2
            Case 4 '�����
                arrTmp = Split(.�㼯, ";")
                For j = 0 To UBound(arrTmp)
                    ReDim Preserve arrXY(j)
                    arrXY(j).x = CLng(Split(arrTmp(j), ",")(0))
                    arrXY(j).y = CLng(Split(arrTmp(j), ",")(1))
                Next
                Polygon objPic.hdc, arrXY(0), UBound(arrXY) + 1
            Case 5 'Բ
                Ellipse objPic.hdc, .X1, .Y1, .X2, .Y2
            Case 6 '���б��
                If .���ɫ = 0 Then
                    Call SetDrawStyleFromValue(objPic.hdc, RGB(255, 255, 0), 0, 1, RGB(255, 255, 0), -2)
                Else
                    Call SetDrawStyleFromValue(objPic.hdc, RGB(255, 255, 0), 0, 1, .���ɫ, -2)
                End If
                Ellipse objPic.hdc, .X1 - 7, .Y1 - 7, .X1 + 7, .Y1 + 7
                If .����ɫ = 0 Then
                    Call SetDrawStyleFromValue(objPic.hdc, vbBlack, 0, 1, vbBlack, -1)
                Else
                    Call SetDrawStyleFromValue(objPic.hdc, .����ɫ, 0, 1, .����ɫ, -1)
                End If
                Ellipse objPic.hdc, .X1 - 7, .Y1 - 7, .X1 + 7, .Y1 + 7
                objFnt.Bold = True
                Call TextOut(objPic, .����, IIf(Len(.����) > 1, .X1 - 6, .X1 - 2), .Y1 - 6, .X1 + 14, .Y1 + 14, objFnt)
        End Select
    End With
    objPic.Refresh

    
'    Set ShowPicMark = objPic.Image
    
End Function

'################################################################################################################
'## ���ܣ�  �жϾ�������Բ�ཻ���
'##
'## ������  (X1,Y1),(X2,Y2) :�������Ͻ������½ǵ�����
'##         (X3,Y3),(X4,Y4) :��Բ���Ͻ������½ǵ�����
'################################################################################################################
Public Function ��������Բ�ཻ(X1 As Long, Y1 As Long, _
    X2 As Long, Y2 As Long, _
    X3 As Long, Y3 As Long, _
    X4 As Long, Y4 As Long) As Boolean
    
    Dim MyRgn As Long, OutRgn As Long, InRgn As Long, R As Long
    MyRgn = CreateRectRgn(0, 0, 0, 0) '����
    OutRgn = CreateRectRgn(X1, Y1, X2, Y2)       '��Բ
    InRgn = CreateEllipticRgn(X3, Y3, X4, Y4)
    R = CombineRgn(MyRgn, OutRgn, InRgn, RGN_AND)

    If R = NULLREGION Or R = 0 Then  '0��ʧ�ܣ�NULLREGION���޽���
'        If (X3 > X1 And X3 < X2 And Y3 > Y1 And Y3 < Y2) Or (X4 > X1 And X4 < X2 And Y4 > Y1 And Y4 < Y2) Then
'            ��������Բ�ཻ = True
'        Else
            ��������Բ�ཻ = False
'        End If
    Else
        ��������Բ�ཻ = True
    End If
End Function

'################################################################################################################
'## ���ܣ�  �жϾ���������ཻ���
'##
'## ������  (X1,Y1),(X2,Y2) :����1���Ͻ������½ǵ�����
'##         (X3,Y3),(X4,Y4) :����2���Ͻ������½ǵ�����
'################################################################################################################
Public Function ����������ཻ(X1 As Long, Y1 As Long, _
    X2 As Long, Y2 As Long, _
    X3 As Long, Y3 As Long, _
    X4 As Long, Y4 As Long) As Boolean
    
    Dim MyRgn As Long, OutRgn As Long, InRgn As Long, R As Long
    MyRgn = CreateRectRgn(0, 0, 0, 0) '����
    OutRgn = CreateRectRgn(X1, Y1, X2, Y2)       '��Բ
    InRgn = CreateRectRgn(X3, Y3, X4, Y4)
    R = CombineRgn(MyRgn, OutRgn, InRgn, RGN_AND)

    If R = NULLREGION Or R = 0 Then  '0��ʧ�ܣ�NULLREGION���޽���
        ����������ཻ = False
    Else
        ����������ཻ = True
    End If
End Function

'################################################################################################################
'## ���ܣ�  �жϾ����������ཻ���
'##
'## ������  (X1,Y1),(X2,Y2) :�������Ͻ������½ǵ�����
'##         Points()        :����ζ������꼯��
'################################################################################################################
Public Function �����������ཻ(X1 As Long, Y1 As Long, _
    X2 As Long, Y2 As Long, _
    Points() As POINTAPI) As Boolean
    
    Dim MyRgn As Long, OutRgn As Long, InRgn As Long, R As Long
    MyRgn = CreateRectRgn(0, 0, 0, 0)           '����
    OutRgn = CreateRectRgn(X1, Y1, X2, Y2)      '��Բ
    InRgn = CreatePolygonRgn(Points(0), UBound(Points), WINDING) '���ݶ���ζ������ݴ��������
    R = CombineRgn(MyRgn, OutRgn, InRgn, RGN_AND)

    If R = NULLREGION Or R = 0 Then  '0��ʧ�ܣ�NULLREGION���޽���
        �����������ཻ = False
    Else
        �����������ཻ = True
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ָ���豸��ָ����Χ���������
'##
'## ������  objOut          :��ͼ����--ͼƬ��ؼ�
'##         strOut          :�ı�����
'##         (X1,Y1),(X2,Y2) :��������
'##         sFont           :�������
'################################################################################################################
Public Sub TextOut(objOut As Object, _
    ByVal strOut As String, _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, _
    ByRef sFont As StdFont)
    
    Dim R As RECT
    
    If Trim(Replace(strOut, vbCrLf, "")) = "" Then Exit Sub
    
    R.Left = X1: R.Right = X2
    R.Top = Y1: R.Bottom = Y2
    
    DrawTextEx objOut.hdc, strOut, LenB(StrConv(strOut, vbFromUnicode)), R, DT_EDITCONTROL Or DT_WORDBREAK, 0&
    
    objOut.Refresh
End Sub

'################################################################################################################
'## ���ܣ�  ��ָ����������ǿ�е�����������
'##
'## ������  (X1,Y1)     :IN         ԭʼ�������Ͻ�����
'##         (X2,Y2)     :IN/OUT     �µľ������½�����
'################################################################################################################
Public Sub ForceSquare(ByVal X1 As Long, ByVal Y1 As Long, X2 As Long, Y2 As Long)
    If Abs(Y2 - Y1) > Abs(X2 - X1) Then
        If X2 < X1 Then
            X2 = X1 - Abs(Y2 - Y1)
        Else
            X2 = X1 + Abs(Y2 - Y1)
        End If
        Y2 = Y2
    Else
        If Y2 < Y1 Then
            Y2 = Y1 - Abs(X2 - X1)
        Else
            Y2 = Y1 + Abs(X2 - X1)
        End If
        X2 = X2
    End If
End Sub




