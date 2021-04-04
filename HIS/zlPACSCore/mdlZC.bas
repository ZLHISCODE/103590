Attribute VB_Name = "mdlPopupMenu"
Option Explicit
'--------------------------------------------------------
'��  �ܣ���ģ����ݵ�ǰ��ͼ���������ͼ�����õ����˵�����Ҫ�Ż�Ϊ����ͼ��ʱ�����˵�����Ҫʱ����ã�
'�����ˣ�����
'�������ڣ�2004.6.12
'���̺����嵥��
'    BRName():       ��ͼ����ȡ��������
'    BRUID():        ��ͼ����ȡ����ID
'    CKUID():        ��ͼ����ȡ���UID
'    CheckDate():        ��ͼ����ȡ����
'    CheckTime():        ��ͼ����ȡʱ��
'    SeriesNum():        ��ͼ����ȡ���к�
'    CheckPart():        ��ͼ����ȡ��λ
'    CheckMenuClass():   ������߲˵����Ƿ��ظ�
'    PopMenu():      �����Ҽ������Ĵ�ͼ��˵�
'    FiltrateStr():      �����ִ�
'�޸ļ�¼��
'    2005.07.08    �ƽ�
'-------------------------------------------------------


'��ͼ����ȡ��������
Function BRName(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H10, &H10)) = False Then
        BRName = "������" & Image.Attributes(&H10, &H10)
    End If
End Function

'��ͼ����ȡͼ������
Function IMGModality(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H8, &H60)) = False Then
        IMGModality = "Ӱ�����" & Image.Attributes(&H8, &H60)
    End If
End Function

'��ͼ����ȡ����ID
Function BRUID(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H10, &H20)) = False Then
        BRUID = Image.Attributes(&H10, &H20)
    End If
End Function

'��ͼ����ȡ���UID
Function CKUID(Image As DicomImage) As String
    If IsNull(Image.Attributes(&H20, &HD)) = False Then
        CKUID = Image.Attributes(&H20, &HD)
    End If
End Function

'��ͼ����ȡ����
Function CheckDate(Image As DicomImage) As String
    With Image
        If IsNull(.Attributes(&H8, &H22)) = False Then
            CheckDate = "���ڣ�" & .Attributes(&H8, &H22)
        End If
        If Len(Trim(CheckDate)) < 1 Then
            If IsNull(.Attributes(&H8, &H23)) = False Then
                CheckDate = "���ڣ�" & .Attributes(&H8, &H23)
            End If
        End If
    End With
End Function

'��ͼ����ȡʱ��
Function CheckTime(Image As DicomImage) As String
    On Error Resume Next
    With Image
        If IsNull(.Attributes(&H8, &H30)) = False Then
            CheckTime = "ʱ�䣺" & .Attributes(&H8, &H30)
        End If
        If Len(Trim(CheckTime)) < 1 Then
            If IsNull(.Attributes(&H8, &H32)) = False Then
                CheckTime = "ʱ�䣺" & .Attributes(&H8, &H32)
            End If
        End If
        If Len(Trim(CheckTime)) < 1 Then
            If IsNull(.Attributes(&H8, &H33)) = False Then
                CheckTime = "ʱ�䣺" & .Attributes(&H8, &H33)
            End If
        End If
    End With
End Function

'��ͼ����ȡ���к�
Function SeriesNum(Image As DicomImage) As String
    With Image
        If IsNull(.Attributes(&H20, &H11)) = False Then
            SeriesNum = "���кţ�" & .Attributes(&H20, &H11)
        End If
    End With
End Function

'��ͼ����ȡ��λ
Function CheckPart(Image As DicomImage) As String
    With Image
        If IsNull(Image.Attributes(&H18, &H15)) = False Then
            CheckPart = "��λ��" & Image.Attributes(&H18, &H15)
        End If
    End With
End Function

'������߲˵����Ƿ��ظ�
Function CheckMenuClass(CheckStr As Variant, MenuName As String, intLevel As Integer, intOneLevel As Integer, intTwoLevel As Integer) As Boolean
    '---------------------------------------------------------------------------------
    '���ܣ�                                  ����һ�α����Ƿ����ظ����ִ�
    '������
    '           CheckStr                     �����ִ�
    '           MenuName                     Ҫ����Ƿ��ظ����ִ�
    '           intLevel                     ��ǰ����
    '           intOneLevel                  ��һ���е�N��
    '           intTwoLevel                  �ڶ����е�N��
    '���أ�                                  =True��ʾ���ظ�  =Flase��ʾû���ظ�
    '�ϼ���������̣�                        ��
    '�¼���������̣�                        ��
    '���õ��ⲿ������                        ��
    '�����ˣ�                                ���� 2005-7-7
    '----------------------------------------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim z As Integer
    Select Case intLevel
        Case 1
            For i = 0 To 20
                If CheckStr(i, 0, 0) = MenuName Then
                    CheckMenuClass = True
                    Exit Function
                End If
            Next
        Case 2
            For j = 0 To 40
                If CheckStr(intOneLevel - 1, j, 0) = MenuName Then
                    CheckMenuClass = True
                    Exit Function
                End If
            Next
        Case 3
            For z = 0 To 60
                If CheckStr(intOneLevel - 1, intTwoLevel, z) = MenuName Then
                    CheckMenuClass = True
                    Exit Function
                End If
            Next
    End Select
    CheckMenuClass = False
End Function

Public Sub PopMenu(f As frmViewer, imgs As DicomImages)
'------------------------------------------------
'���ܣ������Ҽ������Ĵ�ͼ��˵�
'������f--��ʾ�����˵��Ĵ��壻imgs--���ɵ����˵���ͼ����Щͼ����ÿ�����еĵ�һ��ͼ��
'���أ���
'�ϼ���������̣�frmViewer.picViewer_MouseUp
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ�
'------------------------------------------------
    '�Ҽ��˵�
    '����˵�����
    Dim MenuClass(20, 40, 60) As String
    Dim MenuTag(20, 40, 60) As Integer
    Dim UserName As String
    Dim CheckName As String
    Dim CKTime As String
    Dim CheckCKTime As String
    Dim CKPart As String
    '����ѭ������
    Dim i As Integer
    Dim j As Integer
    Dim z As Integer
    Dim k As Integer
    Dim l As Integer
    Dim OneClass As Integer
    Dim TwoClass As Integer
    
    Dim PopupBar As CommandBar
    Dim ControlClass1 As CommandBarPopup
    Dim ControlClass2 As CommandBarPopup
    Dim ControlClass3 As CommandBarControl
    
    If imgs.Count < 1 Then
        Exit Sub
    End If
    
    '********************����**************************
    For i = 1 To imgs.Count
        UserName = BRUID(imgs(i)) & "," & BRName(imgs(i)) & "," & IMGModality(imgs(i))
        If Len(Trim(UserName)) > 1 Then
            If CheckMenuClass(MenuClass, UserName, 1, 1, 1) = False Then
                MenuClass(OneClass, 0, 0) = UserName
                '��¼Viewer
                MenuTag(OneClass, 0, 0) = i
                OneClass = OneClass + 1
            End If
        End If
    Next
    '***************************************************
    '*******************���ʱ��************************
    For i = 1 To OneClass
        k = 1
        UserName = MenuClass(i - 1, 0, 0)
        For j = 1 To imgs.Count
            CheckName = BRUID(imgs(j)) & "," & BRName(imgs(j)) & "," & IMGModality(imgs(j))
            If UserName = CheckName Then
                CKTime = CheckDate(imgs(j))
                CKTime = CKTime & " " & CheckTime(imgs(j))
                CKTime = CKUID(imgs(j)) & "," & CKTime
                CKTime = Trim(CKTime)
                If Len(CKTime) > 1 And CheckMenuClass(MenuClass, CKTime, 2, i, 1) = False Then
                    MenuClass(TwoClass, k, 0) = CKTime
                    '��¼Viewer
                    MenuTag(TwoClass, k, 0) = j
                    k = k + 1
                    TwoClass = TwoClass + 1
                End If
            End If
        Next
    Next
    '********************����+��λ+��������************************
    For i = 1 To OneClass
        UserName = MenuClass(i - 1, 0, 0)
        For j = 1 To TwoClass
            CKTime = MenuClass(i - 1, j, 0)
            k = 1
            For z = 1 To imgs.Count
                CheckName = BRUID(imgs(z)) & "," & BRName(imgs(z)) & "," & IMGModality(imgs(z))
                CheckCKTime = CheckDate(imgs(z))
                CheckCKTime = CheckCKTime & " " & CheckTime(imgs(z))
                CheckCKTime = CKUID(imgs(z)) & "," & CheckCKTime
                If UserName = CheckName And CKTime = CheckCKTime Then
                    CKPart = "," & SeriesNum(imgs(z))
                    CKPart = CKPart & "," & CheckPart(imgs(z))
                    CKPart = CKPart & ",����������" & imgs(z).SeriesDescription
                    If Len(Trim(CKPart)) > 0 And CheckMenuClass(MenuClass, CKPart, 3, i, j) = False Then
                        MenuClass(i - 1, j, k) = CKPart
                        MenuTag(i - 1, j, k) = z
                        k = k + 1
                        '��¼Viewer
                        
                    End If
                End If
            Next
        Next
    Next
    
    '**********************************************************
    '���������˵�
        Set PopupBar = f.ComToolBar.Add("�����˵�", xtpBarPopup)
    '���˶��ڵ��ַ�
    FiltrateStr MenuClass
    
    '���ɲ˵�
    k = 499
    For i = 0 To OneClass - 1
        If Len(MenuClass(i, 0, 0)) > 0 And OneClass > 1 Then
            With PopupBar
                k = k + 1
                Set ControlClass1 = PopupBar.Controls.Add(xtpControlButtonPopup, k, MenuClass(i, 0, 0))
            End With
        End If
        For j = 1 To TwoClass
            If Len(MenuClass(i, j, 0)) > 0 Then
                k = k + 1
                If TwoClass > 1 Then
                    If OneClass > 1 Then
                        Set ControlClass2 = ControlClass1.CommandBar.Controls.Add(xtpControlButtonPopup, k, MenuClass(i, j, 0))
                    Else
                        Set ControlClass2 = PopupBar.Controls.Add(xtpControlButtonPopup, k, MenuClass(i, j, 0))
                    End If
                End If
                For z = 1 To 60
                    If Len(MenuClass(i, j, z)) > 0 Then
                        k = k + 1
                        If TwoClass > 1 Then
                            Set ControlClass3 = ControlClass2.CommandBar.Controls.Add(xtpControlButton, k, MenuClass(i, j, z))
                        Else
                            Set ControlClass3 = PopupBar.Controls.Add(xtpControlButton, k, MenuClass(i, j, z))
                        End If
                        ControlClass3.Category = MenuTag(i, j, z)
                    End If
                Next
            End If
        Next
    Next
    PopupBar.ShowPopup
End Sub

Private Sub FiltrateStr(MenuStr As Variant)
'------------------------------------------------
'���ܣ������ִ�
'������MenuStr--
'���أ���
'�ϼ���������̣�mdlPopupMenu.PopMenu
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ�����
'------------------------------------------------
    Dim i, j, z As Integer
    Dim StrLong As Integer
    For i = 0 To 20
        For j = 0 To 40
            For z = 0 To 60
                StrLong = InStr(MenuStr(i, j, z), ",")
                MenuStr(i, j, z) = Mid$(MenuStr(i, j, z), StrLong + 1)
            Next
        Next
    Next
End Sub

Public Sub ShowFrameSelectImagePopup(f As frmViewer) ', img As DicomImage, lblFrame As DicomLabel)
'------------------------------------------------
'���ܣ�������ѡͼ���ʱ�� ������Ҽ��ĵ����˵�
'������f--��ʾ�����˵��Ĵ��壻 img�����˵���Ӧ��viewer�е�ͼ��lblFrameͼ��ѡ���
'���أ���
'------------------------------------------------

Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '����Ҽ������˵�
    Set cbrToolBar = f.ComToolBar.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ID_ACtive_SaveInReport, "���汨��ͼ")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Public Sub ShowPopup(f As frmViewer, img As DicomImage)
'------------------------------------------------
'���ܣ���������Ҽ������˵�
'������f--��ʾ�����˵��Ĵ��壻 img�����˵���Ӧ��viewer�е�ͼ������ȷ��Ӱ�����
'���أ���
'�����ˣ��ƽ�
'ʱ�䣺2008-4-18
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
Dim cbrToolPopup2 As CommandBarPopup
    
    
    '����Ҽ������˵�
    Set cbrToolBar = f.ComToolBar.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ID_View_UpSeries, "��һ����")
        Set cbrControl = .Add(xtpControlButton, ID_View_DownSeries, "��һ����")
        Set cbrControl = .Add(xtpControlButton, ID_Active_Cruise, "����")
        Set cbrControl = .Add(xtpControlButton, ID_Active_Zoom, "����")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, "�ֶ�����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_AdjustWindow_HandAdjustWindow, "�ֶ�����")
        subSetWidthLevelF img, f, cbrToolPopup

        Set cbrControl = .Add(xtpControlButton, ID_Tool_Magnifier, "�Ŵ�")
        Set cbrControl = .Add(xtpControlButton, ID_ACtive_FrameSelectImage, "��ѡͼ��")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, ID_Active_Lable, "��ע����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_ACtive_Mouse_Value, "��ʾCTֵ")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Rect, "����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Ellipse, "��Բ")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Area, "������")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Angle, "�Ƕ�")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Arrowhead, "��ͷ")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Text, "����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_BeeLine, "ֱ��")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_Curve, "����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Lable_CadioThoracicRatio, "���ر�")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, 0, "ͼ�����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_LeftRight, "ˮƽ����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_TopButton, "��ֱ����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_Left90, "��ת90��")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Eddy_Right90, "��ת90��")
        
        Set cbrControl = .Add(xtpControlButton, ID_Tool_Movie, "��Ӱ����")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, ID_Active_SieveLens, "ͼ����ǿ")
        
        Set cbrToolPopup2 = cbrToolPopup.CommandBar.Controls.Add(xtpControlButtonPopup, ID_Active_SieveLens_Model, "�˾�ģ��")
        Call subSetFilterF(img, f, cbrToolPopup2)
        
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetAdd, "��Ե��ǿǿ������")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_LancetMinus, "��Ե��ǿǿ�ȼ���")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveAdd, "��Ե��ǿ��������")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Sievelens_LeftMoveMinus, "��Ե��ǿ���ȼ���")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessAdd, "ƽ������")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_SieveLens_FlatnessMinus, "ƽ������")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_Sievelens_PhotoReset, "ͼ��ԭ")
        
        Set cbrToolPopup = .Add(xtpControlButtonPopup, 0, "�߼�ͼ����")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_ALL, "���ж�λ��")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_FirstLast, "��β��λ��")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_Now, "��ǰ��λ��")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Active_PointingLine_3DLine, "��ά���")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Tool_ArrowyCoronaryReset, "MPR")
        Set cbrControl = cbrToolPopup.CommandBar.Controls.Add(xtpControlButton, ID_Tool_SlopeReconstruction, "б���ؽ�")
        
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub
