Attribute VB_Name = "mdlCISAudit"
'#########################################################################
'##ģ �� ����mdlCISAudit.bas
'##�� �� �ˣ�ף��
'##��    �ڣ�2012��9��26��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ������������������
'##��    ����
'#########################################################################

Option Explicit


'######################################################################################################################
'��������

Public Const strSplitCmb = "��"

Public Enum COLOR
    ��ɫ = &H80000005
    ��ɫ = &HFF&
    ��ɫ = &HFF0000
    ��ɫ = 0
    �ǽ��� = &HFFEBD7
    ���� = &HFFCC99
    ǳ��ɫ = &HE0E0E0
    ���ɫ = &H8000000C
    ��ɫ = &H8000000F
    ǳ��ɫ = &H80000018
    ��ɫ = &HF5F5F5
    ����ɫ = 0
    ͣ��ɫ = 255
    �϶�ɫ = &HFFE0D9
    ����ģ��ɫ = &HC00000
    
    ��������ɫ = &H40C0&
    ����ǰ��ɫ = &H8000000E
    ���걳��ɫ = &H80C0FF
    �ͱ걳��ɫ = &H80FFFF
    ����ǰ��ɫ = &H80000012
    Ĭ��ǰ��ɫ = &H80000008
    
End Enum

'ö��
Public Enum COLOR_NativeXpPlain
    BackgroundDark = 14054755
    BackgroundLight = 15180411
    HighlightBorderBottomRight = 8388608
    HighlightBorderTopLeft = 8388608
    HighlightHot = 12775167
    HighlightPressed = 4096254
    HighlightSelected = 7323903
    NormalGroupCaptionDark = 14215660
    NormalGroupCaptionLight = 14215660
    NormalGroupCaptionTextHot = 0
    NormalGroupCaptionTextNormal = 0
    NormalGroupClient = 16244694
    NormalGroupClientBorder = 16777215
    NormalGroupClientLink = 12999969
    NormalGroupClientLinkHot = 16748098
    NormalGroupClientText = 0
    SpecialGroupCaptionDark = 14215660
    SpecialGroupCaptionLight = 14215660
    SpecialGroupCaptionTextHot = 0
    SpecialGroupCaptionTextSpecial = 0
    SpecialGroupClient = 16244694
    SpecialGroupClientBorder = 16777215
    SpecialGroupClientLink = 12999969
    SpecialGroupClientLinkHot = 16748098
    SpecialGroupClientText = 0
End Enum

Public Enum REGISTER
    ע����Ϣ
    ˽��ģ��
    ˽��ȫ��
    ����ģ��
    ����ȫ��
End Enum

Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Private mstrSQL As String
Private mstrTitle As String

'==================================================================================================
'=����:ȥ���ַ����еĵ�����("'")(ConvertString)
'=��ڲ���:
'=1).sStr          ����:String
'=���ڲ���:��
'=����:ȥ���ַ���(sStr)�еĵ�����
'=����:2004-12-11
'=���:ŷ��
'=˵��:��SQL����в��ܴ�������
'==================================================================================================
Function ConvertString(ByVal sStr As String) As String
On Error GoTo ErrH
    ConvertString = Replace(sStr, "'", "")
    ConvertString = Replace(ConvertString, "��", "")
    ConvertString = Replace(ConvertString, "&", "")
    Exit Function
ErrH:
    Err.Clear
    ConvertString = ""
End Function

'/******************************************/
'=����:BigNote
'=����:�Ŵ�ע�ֶα༭��,�����ر༭����ı�
' ����:mStr   �Ѿ������ַ���
'      mTitle �༭���ڵı���
'=���:����
'=����:2002-04-03
'/******************************************/
Function Big_Note(mStr As String, mTitle As String, Optional bReadOnly As Boolean, Optional bSqlCheck As Boolean = False) As String
On Error GoTo ErrH
    With FrmNoteBox
        .SqlCheck = bSqlCheck
        .StrText = mStr
        .StrTile = mTitle
        .ReadOnly = bReadOnly
        .Show vbModal
        DoEvents
        Big_Note = IIf(bSqlCheck, .StrText, ConvertString(.StrText))
    End With
    Set FrmNoteBox = Nothing
    Exit Function
ErrH:
    Err.Clear
End Function

Public Function CheckAuditSql_IN(strSQL As String, Optional blnMsg As Boolean = False, Optional intSource As Integer = 0) As Boolean

'=���ܣ� ���SQL���书�� ͨ�� ���� �ж�����
'intSource ����Դ =0 �������ڱ�׼����ִ�У�=1������EMR��ִ��
'���أ�����ͨ������True
    Dim rsTemp          As ADODB.Recordset
    Dim zlCheck         As New clsCheck
    Dim strReturn       As String, strEMRSQL As String, strParm As String
    On Error GoTo ErrH
    If strSQL = "" Then CheckAuditSql_IN = True: Exit Function
    If strSQL = "" Then Exit Function
    strSQL = UCase(strSQL)
    If intSource = 1 Then
        'EMR��Ĳ�ѯ���ܱ�ת�ɴ�дִ�У���Щ���������Ҫ���ִ�Сд,ת�ɴ�дֻ���ڼ���Ƿ���ڸ���ɾ�����
        strEMRSQL = Replace(strSQL, "[MID]", "Hextoraw(:mid)")
        strEMRSQL = Replace(strEMRSQL, "[ALIDIN]", "Hextoraw(:alidin)")
    Else
        strSQL = Replace(strSQL, "[����ID]", "-1")
        strSQL = Replace(strSQL, "[��ҳID]", "-1")
        strSQL = "Select * From (" & strSQL & ") "
    End If
    
    If InStr(1, strSQL, "INSERT") = 1 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ�����д�����ݣ�", vbCritical

        Exit Function
    ElseIf InStr(1, strSQL, "UPDATE") = 1 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ����ܸ������ݣ�", vbCritical

        Exit Function
    ElseIf InStr(1, strSQL, "DELETE") = 1 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ�����ɾ�����ݣ�", vbCritical

        Exit Function
    ElseIf InStr(1, strSQL, ";") > 0 Then
        zlCheck.Msg_OK "�﷨���ʧ�ܣ�����ʹ�á�;����", vbCritical

        Exit Function
    End If
    
    If intSource = 1 Then
        strParm = IIf(InStr(strEMRSQL, ":mid") = 0, "", "A^16^mid")
        If InStr(strEMRSQL, ":alidin") > 0 Then
            If InStr(strEMRSQL, ":mid") > 0 Then
                strParm = strParm & "|"
            End If
            strParm = strParm & "A^16^mid"
        End If
        strReturn = gobjEmr.OpenSQLRecordset(strEMRSQL, strParm, rsTemp)
        If (strReturn <> "" And InStr(strReturn, "ORA-01403") = 0) Or rsTemp Is Nothing Then
            zlCheck.Msg_OK "������ݼ��ʧ��:" & vbCrLf & "��" & strReturn & "��" & vbCrLf & "������¼��������ݼ����䣡", vbExclamation
            Exit Function
        End If
    Else
        strSQL = "Select * From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecord("select ZL_FUN_ExecSql('" & Replace(strSQL, "'", "''") & "') from dual", "mdlCISAudit")
        If rsTemp Is Nothing Then
            zlCheck.Msg_OK "�﷨���ʧ�ܣ�", vbCritical
            Exit Function
        Else
            If InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]:ORA-01403") > 0 Then
                'û�ҵ��κ�����
            ElseIf InStr(1, rsTemp.Fields(0), "[zlsoft]Error[zlsoft]") > 0 Then
                zlCheck.Msg_OK "������ݼ��ʧ��:" & vbCrLf & "��" & Mid(rsTemp.Fields(0), 23) & "��" & vbCrLf & "������¼��������ݼ����䣡", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    If blnMsg Then zlCheck.Msg_OK "�﷨�����ɹ���"
    CheckAuditSql_IN = True
    Set zlCheck = Nothing
    Exit Function
ErrH:
    zlCheck.Msg_OK "�﷨�����ʧ�ܣ�" & vbCrLf & Err.Description, vbCritical
    '�ж�״̬ ���������������
    Err.Clear

    Set zlCheck = Nothing
End Function

'==============================================================================
'=���ܣ� ���SQL���书��
'==============================================================================
Public Function CheckAuditSql_OUT(strSQL As String, Optional lng����ID As Long = -1, Optional lng��ҳID As Long = -1) As String
    On Error GoTo ErrH
    strSQL = UCase(strSQL)
    strSQL = Replace(strSQL, "[����ID]", CStr(lng����ID))
    strSQL = Replace(strSQL, "[��ҳID]", CStr(lng��ҳID))
    
    CheckAuditSql_OUT = strSQL
        
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetEMR_MID_ALIDIN(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByRef strMid As String, ByRef strAlidin As String) As Boolean
    On Error GoTo ErrHandle
    Dim strReturn As String, strExtend_Tag As String, rsTemp As New ADODB.Recordset
    strExtend_Tag = GetEMRIn_Tag(lngPatiID, lngPageId)
    If strExtend_Tag = "" Then Exit Function
    gstrSQL = "Select Rawtohex(ID) As ID, Rawtohex(Master_Id) As Master_Id From Bz_Act_Log Where Extend_Tag = :extendtag"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtend_Tag & "^16^extendtag", rsTemp)
    If strReturn <> "" Then Exit Function
    strMid = rsTemp!Master_id
    strAlidin = rsTemp!ID
    
    GetEMR_MID_ALIDIN = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'==============================================================================
'=���ܣ� �����Ŀ ���ö��� �� �����������
'==============================================================================
Public Function AuditFileTran(strUsed As String, intSource As Integer) As String
    Dim strType     As String
    On Error GoTo ErrH
    Select Case strUsed
        Case "2" 'סԺ����
            strType = IIf(intSource = 0, "2", "02")
        Case "3" '������
            strType = IIf(intSource = 0, "4", "03")
        Case "4" '�����¼
            strType = "3"
        Case "6" 'ҽ������
            strType = "7"
        Case "7" '����֤��
            strType = IIf(intSource = 0, "5", "04")
        Case "8" '֪���ļ�
            strType = IIf(intSource = 0, "6", "05")
    End Select
    AuditFileTran = strType
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Function GetPPFS() As String
    On Error GoTo ErrH
    '����ƥ��
    GetPPFS = ""
    If Val(zlDatabase.GetPara("����ƥ��")) = 0 Then
        GetPPFS = "%"
    End If
    Exit Function
ErrH:
    Err.Clear
    Exit Function
End Function

Public Function GetDateTime(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '����:��ȡ����ʱ��
    '����:
    '����:
    '******************************************************************************************************************
    Dim intDay As Integer
    
    Select Case strMode
    Case "��  ʱ"      '��ʱ
        GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����,bytFlag=1,���ܿ�ʼʱ��,=2,���ܽ���ʱ��
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"      '������
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "������"      '������
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "��  ��"   'ȫ��
        If bytFlag = 1 Then
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "��  ��"       '����
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰһ��"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "ǰ����"
        If bytFlag = 1 Then
            GetDateTime = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetDateTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function


Public Function DockPannelCreate(ByRef dkpMain As DockingPane, ByVal intIndex As Integer, _
                                    ByVal lngCX As Long, ByVal lngCY As Long, _
                                    ByVal bytDirection As DockingDirection, _
                                    Optional ByVal objNeighbour As Pane = Nothing, _
                                    Optional ByVal strTitle As String, _
                                    Optional ByVal bytOptions As PaneOptions) As Pane
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Set DockPannelCreate = dkpMain.CreatePane(intIndex, lngCX, lngCY, bytDirection, objNeighbour)
    DockPannelCreate.Title = strTitle
    DockPannelCreate.Options = PaneNoCaption
    
End Function

Public Function TabControlInit(ByRef tbc As TabControl, _
                                Optional ByVal bytAppearance As XTPTabAppearanceStyle = xtpTabAppearancePropertyPage2003) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With tbc
        
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
'            .Position = bytPosition
        End With
        
        Set .Icons = frmPubResource.imgPublic.Icons
        

        
    End With

    TabControlInit = True
    
End Function

Public Function CopyMenu(ByVal cbsMain As Object, Optional ByVal intNo As Integer = 2) As CommandBar
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrPopupItem2 As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrControl2 As CommandBarControl
    
    '�����˵�����
    
    On Error GoTo errHand
    
    If cbsMain.ActiveMenuBar.Controls(intNo).Visible = False Then Exit Function

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(intNo)
    Set cbrPopupBar = cbsMain.Add("�����˵�", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        
        Set cbrPopupItem = cbrPopupBar.Controls.Add(cbrControl.Type, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.Parameter = cbrControl.Parameter
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
        
        If cbrControl.Type = xtpControlButtonPopup Then
            For Each cbrControl2 In cbrControl.CommandBar.Controls
                Set cbrPopupItem2 = cbrPopupItem.CommandBar.Controls.Add(xtpControlButton, cbrControl2.ID, cbrControl2.Caption)
                cbrPopupItem2.Parameter = cbrControl2.Parameter
            Next
        End If
        
    Next
    
    Set CopyMenu = cbrPopupBar
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SendLMouseButton(ByVal lngHwnd As Long, ByVal X As Single, ByVal Y As Single)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngX As Long
    Dim lngY As Long
    Dim lngLoop As Long
    Dim lngXY As Long
            
    lngX = X / 15
    lngY = Y / 15
        
    lngXY = 2
    For lngLoop = 1 To 15
        lngXY = lngXY * 2
    Next
    
    lngXY = lngXY * lngY + lngX
    
    SendMessage lngHwnd, WM_LBUTTONDOWN, 0, ByVal lngXY
    SendMessage lngHwnd, WM_LBUTTONUP, 0, ByVal lngXY

End Sub



Public Function IncStr(ByVal strVal As String) As String
    '******************************************************************************************************************
    '���ܣ���һ���ַ����Զ���1��
    '˵����ÿһλ��λʱ,���������,��ʮ���ƴ���,����26���ƴ���
    '******************************************************************************************************************
    Dim i As Long, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function

Public Function RestoreTaskPanelPaterrn(ByVal objTpl As Object)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With objTpl
        
        .ColorSet.BackgroundDark = COLOR_NativeXpPlain.BackgroundDark
        .ColorSet.BackgroundLight = COLOR_NativeXpPlain.BackgroundLight
        .ColorSet.HighlightBorderBottomRight = COLOR_NativeXpPlain.HighlightBorderBottomRight
        .ColorSet.HighlightBorderTopLeft = COLOR_NativeXpPlain.HighlightBorderTopLeft
        .ColorSet.HighlightHot = COLOR_NativeXpPlain.HighlightHot
        .ColorSet.HighlightPressed = COLOR_NativeXpPlain.HighlightPressed
        .ColorSet.HighlightSelected = COLOR_NativeXpPlain.HighlightSelected
        
        .ColorSet.NormalGroupCaptionDark = COLOR_NativeXpPlain.NormalGroupCaptionDark
        .ColorSet.NormalGroupCaptionLight = COLOR_NativeXpPlain.NormalGroupCaptionLight
        .ColorSet.NormalGroupCaptionTextHot = COLOR_NativeXpPlain.NormalGroupCaptionTextHot
        .ColorSet.NormalGroupCaptionTextNormal = COLOR_NativeXpPlain.NormalGroupCaptionTextNormal
        .ColorSet.NormalGroupClient = COLOR_NativeXpPlain.NormalGroupClient
        .ColorSet.NormalGroupClientBorder = COLOR_NativeXpPlain.NormalGroupClientBorder
        .ColorSet.NormalGroupClientLink = COLOR_NativeXpPlain.NormalGroupClientLink
        
        .ColorSet.NormalGroupClientLinkHot = COLOR_NativeXpPlain.NormalGroupClientLinkHot
        .ColorSet.NormalGroupClientText = COLOR_NativeXpPlain.NormalGroupClientText
        .ColorSet.SpecialGroupCaptionDark = COLOR_NativeXpPlain.SpecialGroupCaptionDark
        .ColorSet.SpecialGroupCaptionLight = COLOR_NativeXpPlain.SpecialGroupCaptionLight
        .ColorSet.SpecialGroupCaptionTextHot = COLOR_NativeXpPlain.SpecialGroupCaptionTextHot
        .ColorSet.SpecialGroupCaptionTextSpecial = COLOR_NativeXpPlain.SpecialGroupCaptionTextSpecial
        .ColorSet.SpecialGroupClient = COLOR_NativeXpPlain.SpecialGroupClient
        .ColorSet.SpecialGroupClientBorder = COLOR_NativeXpPlain.SpecialGroupClientBorder
        .ColorSet.SpecialGroupClientLink = COLOR_NativeXpPlain.SpecialGroupClientLink
        .ColorSet.SpecialGroupClientLinkHot = COLOR_NativeXpPlain.SpecialGroupClientLinkHot
        .ColorSet.SpecialGroupClientText = COLOR_NativeXpPlain.SpecialGroupClientText
    End With
End Function

Public Function RestoreDockPanelPaterrn(ByVal objDkp As Object)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With objDkp
        
        .ColorSet.BackgroundDark = COLOR_NativeXpPlain.BackgroundDark
        .ColorSet.BackgroundLight = COLOR_NativeXpPlain.BackgroundLight
        .ColorSet.HighlightBorderBottomRight = COLOR_NativeXpPlain.HighlightBorderBottomRight
        .ColorSet.HighlightBorderTopLeft = COLOR_NativeXpPlain.HighlightBorderTopLeft
        .ColorSet.HighlightHot = COLOR_NativeXpPlain.HighlightHot
        .ColorSet.HighlightPressed = COLOR_NativeXpPlain.HighlightPressed
        .ColorSet.HighlightSelected = COLOR_NativeXpPlain.HighlightSelected
        
        .ColorSet.NormalGroupCaptionDark = COLOR_NativeXpPlain.NormalGroupCaptionDark
        .ColorSet.NormalGroupCaptionLight = COLOR_NativeXpPlain.NormalGroupCaptionLight
        .ColorSet.NormalGroupCaptionTextHot = COLOR_NativeXpPlain.NormalGroupCaptionTextHot
        .ColorSet.NormalGroupCaptionTextNormal = COLOR_NativeXpPlain.NormalGroupCaptionTextNormal
        .ColorSet.NormalGroupClient = COLOR_NativeXpPlain.NormalGroupClient
        .ColorSet.NormalGroupClientBorder = COLOR_NativeXpPlain.NormalGroupClientBorder
        .ColorSet.NormalGroupClientLink = COLOR_NativeXpPlain.NormalGroupClientLink
        
        .ColorSet.NormalGroupClientLinkHot = COLOR_NativeXpPlain.NormalGroupClientLinkHot
        .ColorSet.NormalGroupClientText = COLOR_NativeXpPlain.NormalGroupClientText
        .ColorSet.SpecialGroupCaptionDark = COLOR_NativeXpPlain.SpecialGroupCaptionDark
        .ColorSet.SpecialGroupCaptionLight = COLOR_NativeXpPlain.SpecialGroupCaptionLight
        .ColorSet.SpecialGroupCaptionTextHot = COLOR_NativeXpPlain.SpecialGroupCaptionTextHot
        .ColorSet.SpecialGroupCaptionTextSpecial = COLOR_NativeXpPlain.SpecialGroupCaptionTextSpecial
        .ColorSet.SpecialGroupClient = COLOR_NativeXpPlain.SpecialGroupClient
        .ColorSet.SpecialGroupClientBorder = COLOR_NativeXpPlain.SpecialGroupClientBorder
        .ColorSet.SpecialGroupClientLink = COLOR_NativeXpPlain.SpecialGroupClientLink
        .ColorSet.SpecialGroupClientLinkHot = COLOR_NativeXpPlain.SpecialGroupClientLinkHot
        .ColorSet.SpecialGroupClientText = COLOR_NativeXpPlain.SpecialGroupClientText
    End With
End Function

Public Function SetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strKeyValue As String) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ָ������Ϣ������ע�����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strKeyValue-��ֵ
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        Call SaveSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue)
        
    Case ˽��ģ��

        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ˽��ȫ��

        Call SaveSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue)
        
    Case ����ģ��

        Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strKeyValue)
        
    Case ����ȫ��
        
        Call SaveSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue)
        
    End Select
    
    SetRegister = True
    
errHand:
    
End Function

Public Function GetRegister(ByVal enmRegister As REGISTER, ByVal strSection As String, ByVal strKey As String, ByVal strDefKeyValue As String) As String
    '******************************************************************************************************************
    '���ܣ� ��ָ����ע����Ϣ��ȡ����
    '������ enmRegister-ע������
    '       strSection-ע���Ŀ¼
    '       strKey-����
    '       strDefKeyValue-ȱʡ��ֵ
    '���أ� strKeyValue-��ֵ
    '******************************************************************************************************************

    Dim strValue As String
    
    On Error GoTo errHand
    
    Select Case enmRegister
    Case ע����Ϣ
        
        strValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ģ��

        strValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ˽��ȫ��

        strValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ģ��

        strValue = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & strSection, strKey, strDefKeyValue)
        
    Case ����ȫ��
        
        strValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, strDefKeyValue)
        
    End Select
    
    GetRegister = strValue
    
errHand:
End Function

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, Optional ByVal blnNotCache As Boolean) As String
    '******************************************************************************************************************
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************
    
    On Error GoTo errHand
    
    GetPara = zlDatabase.GetPara(varPara, glngSys, lngModual, strDefault, blnNotCache)

errHand:

End Function

Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�����ָ���Ĳ���ֵ
    '������varPara=�����Ż�������������ֻ��ַ����ʹ�������
    '      strValue=Ҫ���õĲ���ֵ
    '      lngModual=ʹ�øò�����ģ��ţ���1230
    '      blnPrivate=�ò����Ƿ��û�˽�в���
    '���أ������Ƿ�ɹ�
    '******************************************************************************************************************

    On Error GoTo ErrH
        
    SetPara = zlDatabase.SetPara(varPara, strValue, glngSys, lngModual, blnSetup)

    Exit Function
    
ErrH:

End Function


Public Function ParamCreate(ByRef rs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    Set rs = New ADODB.Recordset
    
    With rs
        
        .Fields.Append "������", adVarChar, 50
        .Fields.Append "����ֵ", adVarChar, 50
        
        .Open
    End With
    
    ParamCreate = True
    
    Exit Function
    
errHand:
    
End Function

Public Function ParamAdd(ByRef rs As ADODB.Recordset, ByVal strParamName As String, Optional ByVal strParamValue As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.AddNew
    
    rs("������").Value = strParamName
    rs("����ֵ").Value = strParamValue
    
    ParamAdd = True
    
    Exit Function
    
errHand:
End Function

Public Function ParamRead(ByRef rs As ADODB.Recordset, ByVal strParamName As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "������='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        ParamRead = rs("����ֵ").Value
    End If
    
    Exit Function
    
errHand:
End Function

Public Function ParamWrite(ByRef rs As ADODB.Recordset, ByVal strParamName As String, ByVal strParamValue As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error GoTo errHand
    
    rs.Filter = ""
    rs.Filter = "������='" & strParamName & "'"
    If rs.RecordCount > 0 Then
        rs("����ֵ").Value = strParamValue
    End If
    
    Exit Function
    
errHand:
End Function


Public Function ShowPubSelect(ByVal frmParent As Object, _
                                ByVal obj As Object, _
                                ByVal bytStyle As Byte, _
                                ByVal strLvw As String, _
                                ByVal strSavePath As String, _
                                ByVal strDescrible As String, _
                                ByVal rsData As ADODB.Recordset, _
                                ByRef rsResult As ADODB.Recordset, _
                                Optional ByVal lngCX As Long = 9000, _
                                Optional ByVal lngCY As Long = 4500, _
                                Optional ByVal blnMuliSel As Boolean = False, _
                                Optional ByVal strInitKey As String = "", _
                                Optional ByVal strFilterControl As String = "", _
                                Optional ByVal blnOneReturn As Boolean = False) As Byte
    '******************************************************************************************************************
    '���ܣ�������+�б�ṹ,Ӧ���ڱ��ؼ�
    '������
    '      bytStyle:1-TreeView;2-ListView;3-TreeView+ListView
    '���أ�0:ȡ��ѡ��;1:ѡ��;2:�����ݷ���
    '******************************************************************************************************************
    
    Dim lngX As Long
    Dim lngY As Long
    Dim lngObjHeight As Long
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI

    On Error GoTo errHand
    
    If rsData.BOF Then
        ShowPubSelect = 2
        Exit Function
    End If
    
    If blnOneReturn Then
        If rsData.RecordCount = 1 Then
            Set rsResult = rsData
            ShowPubSelect = 1
            Exit Function
        End If
    End If
    
    Call ClientToScreen(obj.hWnd, objPoint)
    
    Select Case TypeName(obj)
    Case "TextBox", "CommandButton"
    
        lngX = objPoint.X * Screen.TwipsPerPixelX - Screen.TwipsPerPixelX
        lngY = obj.Height + objPoint.Y * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        lngObjHeight = obj.Height
        
    Case Else
        lngX = objPoint.X * Screen.TwipsPerPixelX + obj.CellLeft
        lngY = objPoint.Y * Screen.TwipsPerPixelY + obj.CellTop + obj.CellHeight
        lngObjHeight = obj.CellHeight
    End Select
    
    ShowPubSelect = frmPubSelDialog.ShowDialog(frmParent, bytStyle, rsData, strLvw, strDescrible, lngX, lngY, lngCX, lngCY, lngObjHeight, strInitKey, strSavePath, , False, blnMuliSel, strFilterControl)
                                
    If ShowPubSelect = 1 Then
        Set rsResult = rsData
        
        If rsResult.BOF Then
            ShowPubSelect = 0
        End If
        
    End If

    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Public Function GetQuestion(rsCondition As ADODB.Recordset, ByVal strDept As String, ByVal bytApplyMode As Byte, Optional ByVal lngKey As Long, Optional ByVal strStart As String, Optional ByVal strEnd As String, Optional ByVal lng����ID As Long = -1, Optional ByVal lng��ҳID As Long, Optional ByVal lngCur���� As Long, Optional ByVal str��ʼʱ�� As String, Optional ByVal str����ʱ�� As String, Optional ByVal str������ As String) As ADODB.Recordset
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    Dim strTmp As String
    Dim strSQL As String
    Dim int�������� As Integer
    Dim strFilter As String
    
    Dim varState As Variant
    mstrTitle = "���ݰ���"
    
    '1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-ҽ������;7-����֤��;8-֪���ļ�
    
    '�γɲ鲡�˵��ӣӣѣ����

    If bytApplyMode > 1 Then
        '1.���״̬
        '------------------------------------------------------------------------------------------------------------------
        
        strSQL = _
            "Select B.����id, B.��ҳid, B.����,b.��Ժ����id" & vbNewLine & _
            "From (Select B.����id, B.��ҳid, B.����,b.��Ժ����id " & vbNewLine & _
            "       From �����ύ��¼ C, ������ҳ B" & vbNewLine & _
            "       Where C.�ύʱ�� Between [6] And [7] And C.����id = B.����id And C.��ҳid = B.��ҳid And" & vbNewLine & _
            "             B.����״̬ In ([12], [13], [14], [15])" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select B.����id, B.��ҳid, B.����,b.��Ժ����id " & vbNewLine & _
            "       From �����ύ��¼ C, ������ҳ B" & vbNewLine & _
            "       Where C.�ύʱ�� Between [8] And [9] And C.����id = B.����id And C.��ҳid = B.��ҳid And B.����״̬ = 5) B"
        
        If ParamRead(rsCondition, "��Ժ���") <> "" Then
            strSQL = strSQL & " Where Exists (Select 1 From ������ϼ�¼ x Where x.����id=b.����id And x.��ҳid=b.��ҳid And x.������� In (3,13) And x.��Ժ���=[16])"
        End If
        
        '2.��Ժ,��Ժ
        '------------------------------------------------------------------------------------------------------------------
        
        If ParamRead(rsCondition, "��Ժ���") = "" Then
            
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.����id,b.��ҳid,b.����,b.��Ժ����id From ������Ϣ a,������ҳ b Where a.����id=b.����id And Nvl(b.��ҳID,0)<>0 And Nvl(b.״̬,0)<>1 And b.��Ժ���� Is Null "
                    
            If ParamRead(rsCondition, "��ǰ����") <> "" Then strSQL = strSQL & " And b.��ǰ����=[17] "
                        
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.����id,b.��ҳid,b.����,b.��Ժ����id From ������Ϣ a,������ҳ b Where a.����id=b.����id And Nvl(b.��ҳID,0)<>0 And Nvl(b.״̬,0)<>1 And b.��Ժ���� Between [10] And [11]"
                    
        Else
            strSQL = strSQL & " Union All " & vbNewLine & _
                    "Select b.����id,b.��ҳid,b.����,b.��Ժ����id From ������Ϣ a,������ҳ b Where a.����id=b.����id And Nvl(b.��ҳID,0)<>0 And Nvl(b.״̬,0)<>1 And b.��Ժ���� Between [10] And [11] " & vbNewLine & _
                        " And Exists (Select 1 From ������ϼ�¼ x Where x.����id=b.����id And x.��ҳid=b.��ҳid And x.������� In (3,13) And x.��Ժ���=[16])"
        End If
        
        strSQL = "Select b.����id,b.��ҳid From (" & strSQL & ") b,Table (Cast(f_Num2List([18]) As zlTools.t_NumList)) f Where b.��Ժ����id=f.Column_Value "
        Select Case Val(ParamRead(rsCondition, "��������"))
        Case 1          '��ҽ������
            strSQL = strSQL & " And b.���� Is Null "
        Case 2          'ҽ������
            strSQL = strSQL & " And b.���� Is Not Null "
            If ParamRead(rsCondition, "ҽ������") <> "" Then
                strSQL = strSQL & " And b.���� In (" & ParamRead(rsCondition, "ҽ������") & ") "
            End If
        End Select
        
        strTmp = Val(ParamRead(rsCondition, "�ȴ�����")) & ";" & Val(ParamRead(rsCondition, "�ܾ�����")) & ";" & Val(ParamRead(rsCondition, "�������")) & ";" & Val(ParamRead(rsCondition, "��鷴��"))
        varState = Split(strTmp, ";")
        If Val(varState(0)) = 1 Then varState(0) = 1
        If Val(varState(1)) = 1 Then varState(1) = 2
        If Val(varState(2)) = 1 Then varState(2) = 3
        If Val(varState(3)) = 1 Then varState(3) = 4
    
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    strTmp = "Decode(a.��������,1,'סԺҽ��',2,'סԺ����',3,'������',4,'�����¼',5,'��ҳ��¼',6,'ҽ������',7,'����֤��',8,'֪���ļ�',9,'�ٴ�·��') As ��������"
    
    '������������Ժ���zq
    If lngCur���� = 0 Then
        If str������ = "" Then
            strFilter = "A.����ʱ�� BetWeen [20] And [21] And "
        Else
            str������ = GetFeedback(str������)
            strFilter = "A.����ʱ�� BetWeen [20] And [21] And " & str������
        End If
    ElseIf lngCur���� = 1 Then
         If str������ = "" Then
            strFilter = " (A.�������� is null or A.�������� =[19]) And A.����ʱ�� BetWeen [20] And [21] And "
        Else
            str������ = GetFeedback(str������)
            strFilter = " (A.�������� is null or A.�������� =[19]) And A.����ʱ�� BetWeen [20] And [21] And " & str������
        End If
    Else
        If str������ = "" Then
            strFilter = " A.�������� =[19] And A.����ʱ�� BetWeen [20] And [21] And "
        Else
            str������ = GetFeedback(str������)
            strFilter = " A.�������� =[19] And A.����ʱ�� BetWeen [20] And [21] And " & str������
        End If
    End If
    
    Select Case bytApplyMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1, 0                     'ָ��

        mstrSQL = _
            "Select Decode(a.��������,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As ͼ��,a.ID, a.���id, a.�ύid, a.����id, a.��ҳid,a.�������� As ��������id," & strTmp & ", a.��¼����, a.��¼״̬, a.�������, a.������Ŀid, a.������, a.����ʱ��, a.��������,a.���ּ���," & vbNewLine & _
            "       a.����˵��, a.������, a.����ʱ��,Decode(a.��������,4,b.����,c.��������) As �ļ�����,a.�ļ�id,a.ҽ��id,a.����id,a.����,a.��ֵ,a.����˵��,a.��������,a.������¼,e.����,f.���� As ���� " & vbNewLine & _
            "From ����������¼ A,�����ļ��б� b,���Ӳ�����¼ c,������ҳ d,������Ϣ e,���ű� f " & vbNewLine & _
            "Where A.ID = [1] And a.�ļ�id=b.ID(+) And a.�ļ�id=c.ID(+) And a.����id=d.����id And a.��ҳid=d.��ҳid And e.����id=d.����id And f.ID=d.��Ժ����id"
    '------------------------------------------------------------------------------------------------------------------
    Case 2                      '�ȴ��޸�
        
        mstrSQL = _
            "Select Decode(a.��������,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As ͼ��,A.ID, A.���id, A.�ύid, A.����id, A.��ҳid,a.�������� As ��������id,a.�ļ�id,a.ҽ��id,a.����id," & strTmp & ", A.�������,a.����,a.��ֵ,a.����˵��,a.������¼,a.��������, C.����,D.���� As ����" & vbNewLine & _
            "From ����������¼ A, ������ҳ B, ������Ϣ C,���ű� D" & vbNewLine & _
            "Where " & strFilter & "A.��¼״̬ = 1 And A.����id = B.����id And A.��ҳid = B.��ҳid And C.����id = A.����id And d.ID=b.��Ժ����id"
'        strFilter
        
        
        If lng����ID > -1 Then
            mstrSQL = mstrSQL & " And a.����id=[4] And a.��ҳid=[5]"
        Else
            mstrSQL = mstrSQL & " And (a.����id,a.��ҳid) In (" & strSQL & ")"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 3                      '�ȴ�����

        mstrSQL = _
            "Select Decode(a.��������,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As ͼ��,A.ID, A.���id, A.�ύid, A.����id, A.��ҳid,a.�������� As ��������id,a.�ļ�id,a.ҽ��id,a.����id," & strTmp & ", A.�������,a.����,a.��ֵ,a.����˵��,a.������¼,a.��������, C.����,D.���� As ����" & vbNewLine & _
            "From ����������¼ A, ������ҳ B, ������Ϣ C,���ű� D" & vbNewLine & _
            "Where A.��¼״̬ = 2 And A.����id = B.����id And A.��ҳid = B.��ҳid And C.����id = A.����id And D.ID=B.��Ժ����id"

        If lng����ID > -1 Then
            mstrSQL = mstrSQL & " And a.����id=[4] And a.��ҳid=[5]"
        Else
            mstrSQL = mstrSQL & " And (a.����id,a.��ҳid) In (" & strSQL & ")"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 4                      '��������
        
        If lng����ID > -1 Then
            mstrSQL = _
                "Select Decode(a.��������,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As ͼ��,A.ID, A.���id, A.�ύid, A.����id, A.��ҳid,a.�������� As ��������id,a.�ļ�id,a.ҽ��id,a.����id," & strTmp & ", A.�������,a.����,a.��ֵ,a.����˵��,a.��������,a.������¼, C.����,D.���� As ����" & vbNewLine & _
                "From (Select A.ID, A.���id, A.�ύid, A.����id, A.��ҳid, A.��������,a.�ļ�id,a.ҽ��id,a.����id, A.�������,a.����,a.��ֵ,a.����˵��,a.��������,a.������¼" & vbNewLine & _
                "       From ����������¼ A" & vbNewLine & _
                "       Where a.����ʱ�� Between [2] And [3] And a.��¼״̬ = 3" & IIf(lng����ID > -1, " And a.����id=[4] And a.��ҳid=[5] ", "") & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select A.ID, A.���id, A.�ύid, A.����id, A.��ҳid, A.��������,a.�ļ�id,a.ҽ��id,a.����id, A.�������,a.����,a.��ֵ,a.����˵��,a.��������,a.������¼" & vbNewLine & _
                "       From ����������ʷ A" & vbNewLine & _
                "       Where a.����ʱ�� Between [2] And [3] " & IIf(lng����ID > -1, " And a.����id=[4] And a.��ҳid=[5] ", "") & ") A, ������ҳ B, ������Ϣ C,���ű� D" & vbNewLine & _
                "Where A.����id = B.����id And A.��ҳid = B.��ҳid And C.����id = A.����id And D.ID=B.��Ժ����id"
        Else
            mstrSQL = _
                "Select Decode(a.��������,1,'object_advice',2,'object_case',3,'object_case',4,'object_tend',5,'object_first',6,'object_report','object_file') As ͼ��,A.ID, A.���id, A.�ύid, A.����id, A.��ҳid,a.�������� As ��������id,a.�ļ�id,a.ҽ��id,a.����id," & strTmp & ", A.�������,a.����,a.��ֵ,a.����˵��,a.��������,a.������¼, C.����,D.���� As ����" & vbNewLine & _
                "From (Select A.ID, A.���id, A.�ύid, A.����id, A.��ҳid, A.��������,a.�ļ�id,a.ҽ��id,a.����id, A.�������,a.����,a.��ֵ,a.����˵��,a.��������,a.������¼" & vbNewLine & _
                "       From ����������¼ A" & vbNewLine & _
                "       Where a.����ʱ�� Between [2] And [3] And a.��¼״̬ = 3" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select A.ID, A.���id, A.�ύid, A.����id, A.��ҳid, A.��������,a.�ļ�id,a.ҽ��id,a.����id, A.�������,a.����,a.��ֵ,a.����˵��,a.��������,a.������¼" & vbNewLine & _
                "       From ����������ʷ A" & vbNewLine & _
                "       Where a.����ʱ�� Between [2] And [3]) A, ������ҳ B, ������Ϣ C,���ű� D" & vbNewLine & _
                "Where A.����id = B.����id And A.��ҳid = B.��ҳid And C.����id = A.����id And D.ID=B.��Ժ����id"
            
            mstrSQL = "Select * From (" & mstrSQL & ") A Where (a.����id,a.��ҳid) In (" & strSQL & ")"
        End If
        
    End Select
    
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    
    If bytApplyMode > 1 Then
        If strStart = "" Then
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(Now), CDate(Now), lng����ID, lng��ҳID, CDate(ParamRead(rsCondition, "��鿪ʼʱ��")), CDate(ParamRead(rsCondition, "������ʱ��")), CDate(ParamRead(rsCondition, "�鵵��ʼʱ��")), CDate(ParamRead(rsCondition, "�鵵����ʱ��")), CDate(ParamRead(rsCondition, "��Ժ��ʼʱ��")), CDate(ParamRead(rsCondition, "��Ժ����ʱ��")), Val(varState(0)), Val(varState(1)), Val(varState(2)), Val(varState(3)), ParamRead(rsCondition, "��Ժ���"), ParamRead(rsCondition, "��ǰ����"), strDept, lngCur����, CDate(str��ʼʱ��), CDate(str����ʱ��)) ', UCase(str������)
        Else
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(strStart), CDate(strEnd), lng����ID, lng��ҳID, CDate(ParamRead(rsCondition, "��鿪ʼʱ��")), CDate(ParamRead(rsCondition, "������ʱ��")), CDate(ParamRead(rsCondition, "�鵵��ʼʱ��")), CDate(ParamRead(rsCondition, "�鵵����ʱ��")), CDate(ParamRead(rsCondition, "��Ժ��ʼʱ��")), CDate(ParamRead(rsCondition, "��Ժ����ʱ��")), Val(varState(0)), Val(varState(1)), Val(varState(2)), Val(varState(3)), ParamRead(rsCondition, "��Ժ���"), ParamRead(rsCondition, "��ǰ����"), strDept, lngCur����, CDate(str��ʼʱ��), CDate(str����ʱ��))
        End If
    Else
        If strStart = "" Then
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(Now), CDate(Now), lng����ID, lng��ҳID)
        Else
            Set GetQuestion = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, CDate(strStart), CDate(strEnd), lng����ID, lng��ҳID)
        End If
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetProjectUse(ByVal lng����ID As Long) As ADODB.Recordset
'����:���÷����Ƿ��ڲ���������¼���Ѿ���ʹ�ù�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select Count(A.ID) as ���� From �������Ŀ¼ A,���������� B,������鷽�� C,����������¼ D" & vbNewLine & _
            "Where A.����ID= B.ID And B.����ID = C.ID And A.id =D.������ĿID And C.ID=[1] And Rownum >0"
    Set GetProjectUse = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetItemUse(ByVal lng��ĿID As Long) As ADODB.Recordset
'����:������Ŀ�Ƿ��ڲ���������¼���Ѿ���ʹ�ù�
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select Count(*) as ���� From ����������¼ where ������ĿID=[1]"
    Set GetItemUse = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng��ĿID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetRelevanceID(ByVal lng����ID As Long) As ADODB.Recordset
'����:���÷�����¼���Ƿ����������ID,��ȡ����Ҫ������¼��
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select ���ID From ����������¼ Where ID =[1]"
    Set GetRelevanceID = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID)
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetFeedback(ByVal str������ As String) As String
'����:�����˲�ѯ��
    Dim strTemp() As String
    Dim lngRow As Long
    Dim strFeedback As String
    Dim strTempSql As String
    If InStrRev(str������, ",", -1) Then
        strTemp = Split(str������, ",")
        strFeedback = ""
        For lngRow = 0 To UBound(strTemp)
            strFeedback = strFeedback & "'" & strTemp(lngRow) & "'" & ","
        Next
    End If
    
    If strFeedback <> "" Then
        If Right(strFeedback, 1) = "," Then
              strTempSql = " A.������ in (" & Left(strFeedback, Len(strFeedback) - 1) & ") And "
              GetFeedback = strTempSql
        End If
    Else
        If str������ <> "" Then
            strTempSql = " A.������ = '" & str������ & "' And "
            GetFeedback = strTempSql
        End If
    End If
End Function

Public Function GetExamineStartUse() As Boolean
'����:����Ƿ��������˵���鷽��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select ID From ������鷽�� Where ����ʱ�� is not Null"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic")
    If rsTmp.RecordCount > 0 Then
        GetExamineStartUse = True
    Else
        GetExamineStartUse = False
    End If
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetHavePath(ByVal lng����ID As Long) As Boolean
'���ܣ����ָ�����һ����Ƿ��п��õ��ٴ�·��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrH
    
    strSQL = "Select a.Id" & vbNewLine & _
            "From �ٴ�·��Ŀ¼ A, �ٴ�·���汾 B, �ٴ�·������ C," & vbNewLine & _
            "     (Select ����id From �������Ҷ�Ӧ Where ����id = [1]" & vbNewLine & _
            "       Union" & vbNewLine & _
            "       Select ID From ���ű� Where ID = [1]) D" & vbNewLine & _
            "Where a.Id = b.·��id And a.���°汾 = b.�汾�� And a.Id = c.·��id(+) And (c.����id = d.����id or c.����id is null) And Rownum < 2"
    On Error GoTo ErrH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng����ID)
    GetHavePath = Not rsTmp.EOF
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetEPRFile(ByVal lngKey As Long, Optional ByVal lngҽ��id As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    If lngҽ��id > 0 Then
        mstrSQL = "Select  '<'||c.ҽ������||'>' || a.��������  As ���� From ���Ӳ�����¼ a,����ҽ������ b,����ҽ����¼ c Where a.ID=[1] And a.ID=b.����ID And b.ҽ��id=[2] And b.ҽ��id=c.ID"
    Else
        mstrSQL = "Select �������� As ���� From ���Ӳ�����¼ a Where a.ID=[1]"
    End If
    
    Set GetEPRFile = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey, lngҽ��id)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEPRFileStruct(ByVal lngKey As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '����:
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    mstrSQL = "Select ���� From �����ļ��б� a Where a.ID=[1]"

    Set GetEPRFileStruct = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lngKey)
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetSubmitID(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Long
    '******************************************************************************************************************
    '����:�����ύID
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    Dim rs As ADODB.Recordset
    
    mstrSQL = "Select ID From  �����ύ��¼ where ����ID =[1] and ��ҳID =[2] and ����ʱ�� is not Null"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng����ID, lng��ҳID)
    If rs.RecordCount = 1 Then
        GetSubmitID = NVL(rs!ID, 0)
        Exit Function
    End If
    
    GetSubmitID = 0
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetAuditInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
    '******************************************************************************************************************
    '����:�����ύID
    '����:
    '����:
    '******************************************************************************************************************
    On Error GoTo errHand
    
    
    mstrSQL = "Select ����ת��,��Ժ����ID From ������ҳ Where ����ID =[1] And ��ҳID = [2]"
    Set GetAuditInfo = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng����ID, lng��ҳID)
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetCISStruct(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal blnDataMove As Boolean) As ADODB.Recordset
    '******************************************************************************************************************
    '���ܣ���ȡ���Ӳ����Ľṹ
    '������
    '���أ����ؼ�¼��
    '******************************************************************************************************************
    Dim arySerial As Variant
    Dim strTmp As String
    Dim strSerial(1 To 9) As String
    Dim intCount As Integer
    Dim strSQL As String
    Dim strSQL1 As String
    Dim rs As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim blnPath As Boolean '�Ƿ���Ȩ����ʾ�ٴ�·��
    
    '1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-ҽ������;7-����֤��;8-֪���ļ�
    On Error GoTo errHand
    
    strTmp = Trim(zlDatabase.GetPara("��������˳��", glngSys, 1560, "5;1;6;2;3;4;8;7;9"))
    If strTmp = "" Then strTmp = "5;1;6;2;3;4;8;7;9"
    arySerial = Split(strTmp, ";")
    For intCount = 0 To UBound(arySerial)
        strSerial(Val(arySerial(intCount))) = intCount
    Next
    
    '���˿��Ҵ��ڿ��õ��ٴ�·��ʱ����ʾ�ٴ�·����¼
    '���ж��Ƿ���"�ٴ�·��Ӧ��" ���=1256
    If GetPrivFunc(glngSys, 1256) <> "" Then
        blnPath = GetHavePath(lng����ID) 'mlng����ID
    End If
    
    mstrSQL = _
        "Select *" & vbNewLine & _
        "From (Select 'R5' As ID, '' As �ϼ�id, '��ҳ��¼' As ����, '' As ����,1 As ĩ��,'object_first' As ͼ��,[7] As ���� " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R2' As ID, '' As �ϼ�id, 'סԺ����' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,[4] As ���� " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R3' As ID, '' As �ϼ�id, '������' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,[5] As ���� " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R4' As ID, '' As �ϼ�id, '�����¼' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,[6] As ���� " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R1' As ID, '' As �ϼ�id, 'סԺҽ��' As ����, '' As ����,1 As ĩ��,'object_advice' As ͼ��,[3] As ���� " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R6' As ID, '' As �ϼ�id, 'ҽ������' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,[8] As ���� " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R7' As ID, '' As �ϼ�id, '����֤��' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,[9] As ���� " & vbNewLine & _
        "       From Dual Union All" & vbNewLine & _
        "       Select 'R8' As ID, '' As �ϼ�id, '֪���ļ�' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,[10] As ���� " & vbNewLine & _
        "       From Dual " & vbNewLine & _
        IIf(blnPath, " Union All Select 'R9' As ID, '' As �ϼ�id, '�ٴ�·��' As ����, '' As ����,0 As ĩ��,'Folder' As ͼ��,[11] As ���� From Dual", "")
        
    mstrSQL = mstrSQL & " Union All" & vbNewLine & _
        "Select a.�ϼ�id || 'K' || Trim(To_Char(a.ID))|| ','||Trim(To_Char(Nvl(a.ҽ��id,0)))||',0' As ID, �ϼ�id, Decode(a.ҽ��id, Null, a.����, '<'||b.ҽ������||'>' || a.����)||'������:'||To_Char(a.����ʱ��,'yyyy-mm-dd hh24:mi')|| Decode(a.���汾, 1, '��д��', '�޶���') || a.������ || '��' || To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi') || Decode(Nvl(a.ǩ������, 0), 0, '����(δ���)', 1, '���', '��ǩ') ||Decode(a.���ʱ��,Null,'��',',���:'||To_Char(a.���ʱ��,'yyyy-mm-dd hh24:mi')||'��') As ����, Trim(To_Char(a.ID))||';'||Decode(a.ҽ��id,Null,'0',Trim(To_Char(a.ҽ��id))) As ����,1 As ĩ��,Decode(��������,2,'object_case',3,'object_case',7,'object_report','object_file') As ͼ��,���� " & vbNewLine & _
        "From (Select ID, Decode(��������, 2, 'R2', 4, 'R3', 7, 'R6',6,'R8',5,'R7') As �ϼ�id,ǩ������,����ʱ��,������,���汾, a.�������� As ����,c.ҽ��id,a.��������,a.����ʱ��,a.���ʱ��,To_Char(a.����ʱ��,'yyyy-mm-dd hh24:mi:ss') As ���� " & vbNewLine & _
        "       From ���Ӳ�����¼ a,����ҽ������ c " & vbNewLine & _
        "       Where a.����id = [1] And a.��ҳid = [2] And c.����id(+)=a.ID And �������� In (2, 3, 4, 5, 6, 7)) a,����ҽ����¼ b Where a.ҽ��id=b.Id(+) "

    
    strSQL1 = "Select 1 From ���˻����¼ A Where a.����id = [1] And a.��ҳid = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL1, "����Ƿ�����ϰ�����", lng����ID, lng��ҳID)
    If rsTemp.RecordCount > 0 Then
        mstrSQL = mstrSQL & " Union All" & _
            "       Select 'R4K' || Trim(To_Char(A.Id))||',0,'|| Trim(To_Char(A.����Id)) As ID, 'R4' As �ϼ�id," & vbNewLine & _
            "              A.���� || '(' || B.���� || '��' || To_Char(A.��ʼ, 'yyyy-mm-dd hh24:mi') || ' �� ' ||" & vbNewLine & _
            "               To_Char(A.��ֹ, 'yyyy-mm-dd hh24:mi') || ')' As ����, Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'yyyy-mm-dd hh24:mi')||' �� '||To_Char(A.��ֹ, 'yyyy-mm-dd hh24:mi')||';'||Trim(To_Char(A.ID)) As ����,1 As ĩ��,'object_tend' As ͼ��,To_Char(a.��ʼ,'yyyy-mm-dd hh24:mi:ss') As ����" & vbNewLine & _
            "       From (Select F.ID, F.���, F.����, R.��ʼ, R.��ֹ, R.����id, ����" & vbNewLine & _
            "              From (Select ID, ���, ����, 3 As ������, ͨ��, 0 As ����id, ����" & vbNewLine & _
            "                     From �����ļ��б�" & vbNewLine & _
            "                     Where ���� = 3 And ���� < 0" & vbNewLine & _
            "                     Union All" & vbNewLine & _
            "                     Select L.ID, L.���, L.����, F.���� As ������, L.ͨ��, A.����id, L.����" & vbNewLine & _
            "                     From �����ļ��б� L, ����ҳ���ʽ F, ����Ӧ�ÿ��� A" & vbNewLine & _
            "                     Where L.���� = 3 And L.���� = 0 And L.���� = F.���� And L.��� = F.��� And L.ID = A.�ļ�id(+)) F," & vbNewLine & _
            "                   (Select R.����id, Nvl(Min(R.������), 3) As ������, Min(R.����ʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ" & vbNewLine & _
            "                     From ���˻����¼ R" & vbNewLine & _
            "                     Where R.������Դ = 2 And R.����id = [1] And Nvl(R.��ҳid, 0) = [2] And Nvl(R.Ӥ��, 0) = 0" & vbNewLine & _
            "                     Group By R.����id) R" & vbNewLine & _
            "              Where (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = R.����id) And F.������ >= R.������) A, ���ű� B" & vbNewLine & _
            "       Where A.����id = B.ID)" & vbNewLine & _
            "Order By Decode(�ϼ�id,Null,' ',�ϼ�id),����"
    Else
        mstrSQL = mstrSQL & " Union All" & _
                   " Select 'R4K'||Trim(To_Char(A.ID))||',0,'||Trim(To_Char(A.����Id)) As ID,'R4' As �ϼ�id," & vbNewLine & _
                   "     A.����||'('||B.����||'��'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI') || '��' ||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI') || ')' As ����," & vbNewLine & _
                   "      Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'YYYY-MM-DD HH24:MI')||'��'||To_Char(A.��ֹ, 'YYYY-MM-DD HH24:MI')||';'||Trim(To_Char(A.ID))||';'||Trim(To_Char(A.Ӥ��)) As ����," & vbNewLine & _
                   "       1 As ĩ��,'object_tend' As ͼ��,To_Char(a.��ʼ,'YYYY-MM-DD HH24:MI:SS') As ����" & vbNewLine & _
                   " From (" & vbNewLine & _
                   "   Select R.ID, F.���, R.����,R.Ӥ��, R.��ʼ, NVL(R.��ֹ,nvl(R.ʱ��,R.��ʼ)) ��ֹ, R.����id, ����" & vbNewLine & _
                   "   From (" & vbNewLine & _
                   "       Select L.ID, L.���, L.����, F.���� As ������, L.ͨ��, L.����" & vbNewLine & _
                   "          From ����ҳ���ʽ F, �����ļ��б� L" & vbNewLine & _
                   "          Where L.���� = 3 And L.���� = F.���� And L.��� = F.��� And (L.ͨ��=1 OR L.ͨ��=2)" & vbNewLine & _
                   "" & vbNewLine & _
                   "       ) F,(" & vbNewLine & _
                   "       Select R.ID,R.����id,R.�ļ����� ����,R.��ʽID,nvl(R.Ӥ��,0) Ӥ��,Min(R.��ʼʱ��) As ��ʼ, Max(R.����ʱ��) As ��ֹ,MAX(T.����ʱ��) ʱ��" & vbNewLine & _
                   "          From ���˻����ļ� R,���˻������� T" & vbNewLine & _
                   "          Where R.ID=T.�ļ�ID(+) And R.����id = [1] And Nvl(R.��ҳid, 0) = [2]" & vbNewLine & _
                   "          Group By R.ID,R.�ļ�����,R.����id,R.��ʽID,R.Ӥ��" & vbNewLine & _
                   "       ) R" & vbNewLine & _
                   "       Where F.ID=R.��ʽID" & vbNewLine & _
                   "   ) A, ���ű� B Where A.����id = B.ID And DECODE(A.����,-1,0,A.Ӥ��)=A.Ӥ��)" & vbNewLine & _
                   " Order By Decode(�ϼ�id,Null,' ',�ϼ�id),����"
    
    End If
        
        
    If lng����ID > 0 Then
        'ֻ����סԺ����
        If blnDataMove Then
            mstrSQL = Replace(mstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
            mstrSQL = Replace(mstrSQL, "����ҽ����¼", "H����ҽ����¼")
            mstrSQL = Replace(mstrSQL, "����ҽ������", "H����ҽ������")
            mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
        End If
    End If
    
    Set GetCISStruct = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, lng����ID, lng��ҳID, strSerial(1), strSerial(2), strSerial(3), strSerial(4), strSerial(5), strSerial(6), strSerial(7), strSerial(8), strSerial(9))
    
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetEMRIn_Tag(ByVal lngPatiID As Long, ByVal lngPageId As Long) As String
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "Select Nvl(a.Id, b.Id) ID" & vbNewLine & _
                "From (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [1] And ��ҳid = [2] And ��ʼԭ�� = 2 And Nvl(���Ӵ�λ, 0) = 0) A," & vbNewLine & _
                "     (Select Max(ID) ID From ���˱䶯��¼ Where ����id = [2] And ��ҳid = [2] And ��ʼԭ�� = 1 And Nvl(���Ӵ�λ, 0) = 0) B"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ԺID", lngPatiID, lngPageId)
    If rsTemp Is Nothing Then Exit Function
    If NVL(rsTemp!ID) = "" Then Exit Function
    GetEMRIn_Tag = "BD_" & rsTemp!ID
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetEmrCISStruct(ByVal lngPatiID As Long, ByVal lngPageId As Long) As ADODB.Recordset
Dim rsTemp As New ADODB.Recordset, strExtendTag As String, strReturn As String
    If gobjEmr Is Nothing Then Set GetEmrCISStruct = Nothing: Exit Function
    strExtendTag = GetEMRIn_Tag(lngPatiID, lngPageId)
    If strExtendTag = "" Then Set GetEmrCISStruct = Nothing: Exit Function
    
    '�ϼ�ID��ID�����ƣ�������ͼ��
    gstrSQL = "Select Decode(e.Kind, '02', 'R2', '03', 'R3', '04', 'R7', '05', 'R8', 'R2') �ϼ�id, Nvl(d.Subdoc_Id, Rawtohex(b.Id)) As ID," & vbNewLine & _
                "       d.Subdoc_Id As ���ĵ�id," & vbNewLine & _
                "       Nvl(d.Subdoc_Title, b.Title) || Decode(d.Completor, Null, '', '�� ' || d.Completor || ' ��' || To_Char(d.Complete_Time, 'yyyy-mm-dd hh24:mi') || 'ǩ����') As ����" & vbNewLine & _
                "       , Rawtohex(b.Id) || Decode(d.Subdoc_Id, Null, Null, '|' || d.Subdoc_Id) As ����, 'object_case' As ͼ��" & vbNewLine & _
                "From Bz_Doc_Log B," & vbNewLine & _
                "     (Select Distinct a.Id, c.Antetype_Id, c.Subdoc_Id, c.Subdoc_Title, c.Real_Doc_Id, c.Complete_Time, c.Completor" & vbNewLine & _
                "       From Bz_Act_Log A, Bz_Doc_Tasks C" & vbNewLine & _
                "       Where a.Extend_Tag = :etag And a.Id = c.Actlog_Id And c.Real_Doc_Id Is Not Null) D, Antetype_List E" & vbNewLine & _
                "Where b.Actlog_Id = d.Id And d.Real_Doc_Id = b.Id And d.Antetype_Id = e.Id And Decode(d.Subdoc_Id, Null, d.Antetype_Id, b.Antetype_Id) = b.Antetype_Id " & vbNewLine & _
                "Order By e.Code, b.Creat_Time,d.Complete_Time"
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strExtendTag & "^16^etag", rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEmrCISStruct = Nothing: Exit Function
    End If
    
    Set GetEmrCISStruct = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function GetEMRFile(ByVal strKey As String) As ADODB.Recordset
Dim strDocid As String, strSubDocid As String, strReturn As String, rsTemp As ADODB.Recordset
    On Error GoTo errHand
    If gobjEmr Is Nothing Then Exit Function
    If InStr(strKey, "|") = 0 Then
        strDocid = strKey
    Else
        strDocid = Split(strKey, "|")(0)
        strSubDocid = Split(strKey, "|")(1)
    End If
    
    gstrSQL = "Select Nvl(b.Subdoc_Title, a.Title) ����" & vbNewLine & _
                "From Bz_Doc_Log A, Bz_Doc_Tasks B" & vbNewLine & _
                "Where a.Id = Hextoraw(:docid) And a.Id = b.Real_Doc_Id" & IIf(strSubDocid = "", "", " And b.Subdoc_Id =:subdocid")
    strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, strDocid & "^16^docid" & IIf(strSubDocid = "", "", "|" & strSubDocid & "^16^subdocid"), rsTemp)
    If strReturn <> "" Then
        MsgBox strReturn, vbCritical, gstrSysName
        Set GetEMRFile = Nothing: Exit Function
    End If
    
    Set GetEMRFile = rsTemp
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

