Attribute VB_Name = "mdlBlackList"
Option Explicit

Public Const G_AlternateColor As Long = 16772055   '�н���ɫ
Public Const G_LostFocusColor As Long = &HE0E0E0   'ʧȥ����ʱ�����񱳾�ɫ
Public Enum gEM_BlackListFun
    Em_Pane_FunFace = 1
    Em_Pane_Face = 2
    Em_Pane_Type = 11 '������¼����
    Em_Pane_Reason = 12 '���õĲ�����¼ԭ��
    Em_Pane_Record = 13 '������¼����
End Enum

Public Function zlGetFirstCommandBar(ByRef objControls As CommandBarControls) As Long
    '���ܣ���ȡ��������ӡԤ����ť��ĵ�һ����ť��index
    Dim objControl As CommandBarControl, idx As Long
    
    For Each objControl In objControls
        If objControl.ID = conMenu_File_Preview Then
            idx = objControl.Index + 1
        End If
    Next
    zlGetFirstCommandBar = idx
End Function

Public Function zlGetPopupCommandBar(frmMain As Form, cbsMain As CommandBars, _
    Optional ByVal lngControlPopupID As Long = conMenu_EditPopup) As CommandBar
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������˵�
    '����:���ص����˵�����
    '����:���˺�
    '����:2018-11-08 11:21:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup, cbCommandBar As CommandBar
    Dim cbrControl As CommandBarControl, cbrControlNew As CommandBarControl
    Dim i As Integer
    
    On Error GoTo errHandle
      
    Set objPopup = cbsMain.FindControl(xtpControlPopup, lngControlPopupID, , True)
    If objPopup Is Nothing Then Exit Function
    Set cbCommandBar = cbsMain.Add("Popup", xtpBarPopup) '�����˵�
    If cbCommandBar Is Nothing Then Exit Function
    
    For i = 1 To objPopup.CommandBar.Controls.Count
        Set cbrControl = objPopup.CommandBar.Controls(i)
        Call frmMain.zlUpdateCommandBars(cbrControl) '�ж��Ƿ�ɼ�����Ϊ��һ��ʱ�˵���û��ִ��Update
        If cbrControl.Visible Then
            Set cbrControlNew = cbCommandBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
            cbrControlNew.BeginGroup = cbrControl.BeginGroup
            cbrControlNew.Enabled = cbrControl.Enabled
        End If
    Next
    Set zlGetPopupCommandBar = cbCommandBar
    
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetCommbarFromName(ByVal objThis As CommandBars, ByVal strName As String, Optional intIndex_Out As Integer) As CommandBar
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ˵����ƣ���ȡָ���Ĳ˵�����
    '���:strName-����
    '����:intIndex_Out-���ص�����
    '����:�ɹ�����CommandBar,���򷵻�False
    '����:���˺�
    '����:2018-11-15 15:13:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    
    For i = 1 To objThis.Count
        If objThis(i).Title = strName Then
        
            Set GetCommbarFromName = objThis(i)
            intIndex_Out = i: Exit Function
        End If
    Next
    intIndex_Out = 0
    Set GetCommbarFromName = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Public Sub SetReportControlBackColorAlternate(rptData As ReportControl, Optional CustomColor As OLE_COLOR = -1)
    '����ReportControl���н���ɫ
    Dim i As Long, objItem As ReportRecordItem
    Dim lngRowCount As Long '�����к�
    
    On Error Resume Next
    For i = 0 To rptData.Rows.Count - 1
        If rptData.Rows(i).GroupRow Then
            lngRowCount = 0
        Else
            For Each objItem In rptData.Rows(i).Record
                If lngRowCount Mod 2 = 0 Then
                    objItem.BackColor = rptData.PaintManager.BackColor
                Else
                    objItem.BackColor = IIf(CustomColor = -1, G_AlternateColor, CustomColor)
                End If
            Next
            lngRowCount = lngRowCount + 1
        End If
    Next
End Sub
Public Function zlGetVsfGrid(rptData As ReportControl, ByRef vsGrid As VSFlexGrid, Optional ByVal strHiddenCols As String) As Boolean
    '����:��ReportControlת��ΪVSFlexGrid
    '���:
    '   strHiddenCols ����������(������0��ʼ)����ʽ����1,��2,��3,...
    
    Dim i As Long, j As Long, lngRowIndex As Long
    Dim varData As Variant
    
    Err = 0: On Error GoTo ErrHandler
    With vsGrid
        .Clear
        .Cols = rptData.Columns.Count
        .Rows = rptData.Records.Count + 1
        .FixedAlignment(-1) = flexAlignCenterCenter
        
        '������
        For i = 0 To rptData.Columns.Count - 1
            .TextMatrix(0, i) = rptData.Columns(i).Caption
            .ColWidth(i) = rptData.Columns(i).Width * 16
            .ColAlignment(i) = Choose(rptData.Columns(i).Alignment + 1, 1, 4, 7)
        Next
        '������
        If strHiddenCols <> "" Then
            varData = Split(strHiddenCols, ",")
            For i = 0 To UBound(varData)
                .ColWidth(Val(varData(i))) = 0
            Next
        End If
        
        '������
        lngRowIndex = 1
        For i = 0 To rptData.Rows.Count - 1
            If rptData.Rows(i).GroupRow = False Then
                For j = 0 To rptData.Columns.Count - 1
                    .TextMatrix(lngRowIndex, j) = rptData.Rows(i).Record(j).Value
                Next
                lngRowIndex = lngRowIndex + 1
            End If
        Next
    End With
    zlGetVsfGrid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then Resume
End Function

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional bln������� As Boolean = True, Optional bln���� As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     bln�������     �Ƿ���и������
    '     bln����         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    zlDblIsValid = zlCommFun.DblIsValid(strInput, intMax, bln�������, bln����, hWnd, str��Ŀ)
End Function

