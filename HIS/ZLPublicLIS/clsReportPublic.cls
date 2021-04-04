VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReportPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'�ӿ����ࣺ �����ӿ�
'ʵ�ֹ��ܣ� ֽ�š���ӡ����ҳü��ҳ�š��ı�����ȴ���
'������Ա��
'����ʱ�䣺
'�ӿ�������
'�ӿڷ����� 1.InitElement
'           2.OpenSetupDialog
'           3.FormatHeadData
'           4.FormatFootData
'           5.FormatParagraph
'######################################################################################################################

Private mrsParameter        As ADODB.Recordset
Private mobjRect            As USERRECT
Private mobjFont            As USERFONT
Private mintCount           As Integer

'�ӿ�
'######################################################################################################################

Public Function SetElementParameter(ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ������������
    '������rsProperty    ���Լ�¼��
    '���أ����óɹ�����True,���򷵻�False
    '******************************************************************************************************************
    Set mrsParameter = rsParameter
    SetElementParameter = True
    
End Function

Public Function InitElement(ByRef objPaper As USERPAPER, ByRef strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '���ܣ������ӿڣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim objclsLisPrintSet As New clsLisPrintSet
        
    InitElement = objclsLisPrintSet.GetPaper(objPaper, strRegisterPath)
    
    '�̶�ֽ�ŵ���ز���
    '------------------------------------------------------------------------------------------------------------------
    With objPaper
        .PaperSize = 9
        .Orientation = 1
        .Width = gclsLisReportLib.GetTwipsX(21)                                         'ֽ�ſ���
        .Height = gclsLisReportLib.GetTwipsY(29.7)                                      'ֽ�Ÿ߶�
        .BorderLeft = gclsLisReportLib.GetTwipsX(0.5)                                   '��߽�
        .BorderRight = gclsLisReportLib.GetTwipsX(0.5)                                  '�ұ߽�
        .BorderTop = 0 'gclsLisReportLib.GetTwipsY(1)                                   '�ϱ߽�
        .BorderBottom = gclsLisReportLib.GetTwipsY(2)                                   '�±߽�
        .PageHead = gclsLisReportLib.GetTwipsY(1)                                       'ҳü�߶�
        .PageFoot = gclsLisReportLib.GetTwipsY(1)                                       'ҳ�Ÿ߶�
        .SpaceTop = gclsLisReportLib.GetTwipsY(0)                                       '
        .SpaceBottom = gclsLisReportLib.GetTwipsY(0)
    End With
    
End Function

Public Function OpenSetupDialog(frmMain As Object, ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '���ܣ������ӿڣ�����/ѡ������
    '������
    '˵�������ô˷���ǰ�������ȵ���InitData����
    '���أ�
    '******************************************************************************************************************
    Dim objPaper As USERPAPER
    
    strRegisterPath = strRegisterPath
    
    '��mmΪ��λ���浽ע�����
    With objPaper
        .PaperSize = 9
        .Orientation = 1
        .Width = 210
        .Height = 297
        .BorderLeft = 5
        .BorderRight = 5
        .BorderTop = 10
        .BorderBottom = 10
        .PageHead = 23                               'ҳü�߶�
        .PageFoot = 4
        .SpaceTop = 0
        .SpaceBottom = 0
    End With
    
    SaveSetting "ZLSOFT", strRegisterPath, "ֽ��", objPaper.PaperSize
    SaveSetting "ZLSOFT", strRegisterPath, "����", objPaper.Width
    SaveSetting "ZLSOFT", strRegisterPath, "�߶�", objPaper.Height
    SaveSetting "ZLSOFT", strRegisterPath, "ֽ��", objPaper.Orientation
    SaveSetting "ZLSOFT", strRegisterPath, "��߾�", objPaper.BorderLeft
    SaveSetting "ZLSOFT", strRegisterPath, "�ұ߾�", objPaper.BorderRight
    SaveSetting "ZLSOFT", strRegisterPath, "�ϱ߾�", objPaper.BorderTop
    SaveSetting "ZLSOFT", strRegisterPath, "�±߾�", objPaper.BorderBottom
    
    OpenSetupDialog = frmListReportSet.ShowDialog(frmMain, 9, False, strRegisterPath)
    
End Function

Public Function GetPageHeadData() As Boolean
    '******************************************************************************************************************
    '���ܣ������ӿڣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTmp                      As String
    Dim objSvrFont                  As USERFONT
    Dim varTmp                      As Variant
    Dim objRect                     As USERRECT
    Dim objFont                     As USERFONT
    Dim objclsGrID                  As clsGrid
    Dim objclsLisReportHead         As New clsLisReportHead

    On Error GoTo errHand

    '1.��ȡ��ǰҪ�������ʼλ��(����Ĵ��벻��ɾ��)
    '------------------------------------------------------------------------------------------------------------------
    objRect = CurrentRect

    '2.��ȡ��ǰ����������(����Ĵ��벻��ɾ��)
    '------------------------------------------------------------------------------------------------------------------
    objFont = CurrentFont

    '------------------------------------------------------------------------------------------------------------------
    CurrentRect = objRect

    GetPageHeadData = True

    Exit Function

    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    gclsLisReportLib.Font = objSvrFont
End Function

Public Function GetPageFootData() As Boolean
    '******************************************************************************************************************
    '���ܣ������ӿڣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objSvrFont              As USERFONT
    Dim objRect                 As USERRECT
    Dim objFont                 As USERFONT
    Dim objclsLisReportFoot     As New clsLisReportFoot
    
    On Error GoTo errHand
    
    '1.��ȡ��ǰҪ�������ʼλ��(����Ĵ��벻��ɾ��)
    '------------------------------------------------------------------------------------------------------------------
    objRect = CurrentRect
    
    '2.��ȡ��ǰ����������(����Ĵ��벻��ɾ��)
    '------------------------------------------------------------------------------------------------------------------
    objFont = CurrentFont


    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - 200  '- gclsLisReportLib.Paper.PageFoot '- 200 '+ gclsLisReportLib.GetTwipsY(0.2)
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    objRect.Y1 = objRect.Y0
    
    Call objclsLisReportFoot.InsertLine(objRect, objFont)
    
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - 150 '- gclsLisReportLib.Paper.PageFoot '- 150 '+ gclsLisReportLib.GetTwipsY(0.3)
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    objRect.Y1 = 0

    objFont.HAlignment = 2
    Call objclsLisReportFoot.InsertPageNo("�� n ҳ / �� m ҳ", objRect, objFont)
    
    '------------------------------------------------------------------------------------------------------------------
    CurrentRect = objRect

    GetPageFootData = True
    
    Exit Function

    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    gclsLisReportLib.Font = objSvrFont
End Function

Public Property Get CurrentRect() As USERRECT
    CurrentRect = mobjRect
End Property

Public Property Let CurrentRect(vData As USERRECT)
    mobjRect = vData
End Property

Public Property Get CurrentFont() As USERFONT
    CurrentFont = mobjFont
End Property

Public Property Let CurrentFont(vData As USERFONT)
    mobjFont = vData
End Property

Public Function FormatSetup(frmMain As Object, ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '���ܣ������ӿڣ�����/ѡ������
    '������
    '˵�������ô˷���ǰ�������ȵ���InitData����
    '���أ�
    '******************************************************************************************************************
    Dim objPaper As USERPAPER
    
    strRegisterPath = strRegisterPath
    
    '��mmΪ��λ���浽ע�����
    With objPaper
        .PaperSize = 9
        .Orientation = 1
        .Width = 210
        .Height = 297
        .BorderLeft = 5
        .BorderRight = 5
        .BorderTop = 10
        .BorderBottom = 10
        .PageHead = 23                               'ҳü�߶�
        .PageFoot = 4
        .SpaceTop = 0
        .SpaceBottom = 0
    End With
    
    SaveSetting "ZLSOFT", strRegisterPath, "ֽ��", objPaper.PaperSize
    SaveSetting "ZLSOFT", strRegisterPath, "����", objPaper.Width
    SaveSetting "ZLSOFT", strRegisterPath, "�߶�", objPaper.Height
    SaveSetting "ZLSOFT", strRegisterPath, "ֽ��", objPaper.Orientation
    SaveSetting "ZLSOFT", strRegisterPath, "��߾�", objPaper.BorderLeft
    SaveSetting "ZLSOFT", strRegisterPath, "�ұ߾�", objPaper.BorderRight
    SaveSetting "ZLSOFT", strRegisterPath, "�ϱ߾�", objPaper.BorderTop
    SaveSetting "ZLSOFT", strRegisterPath, "�±߾�", objPaper.BorderBottom
    
    FormatSetup = frmListReportSet.ShowDialog(frmMain, 9, False, strRegisterPath)
    
End Function

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long

    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)

    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Private Sub Class_Initialize()
    Set gclsLisReportLib = New clsLisReportLib
End Sub

Private Sub Class_Terminate()
    Set mrsParameter = Nothing
End Sub

