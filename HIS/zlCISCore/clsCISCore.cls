VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISCore"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Sub Class_Initialize()
    Set gobjCISCore = Me
End Sub

Public Sub InitCISCore(cnMain As ADODB.Connection)
'------------------------------------------------
'���ܣ� �����ĳ�ʼ������������ı�Ҫ���ݴ�������
'������
'   cnMain:����������ݿ�����
'���أ�
'------------------------------------------------
    Set gcnOracle = cnMain
    
    gstrAviPath = GetSetting("ZLSOFT", "ע����Ϣ", "gstrAviPath", "")
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrHelpPath = gstrAviPath & "\help"
    glngSys = 100
    Call GetUserInfo
    gstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
End Sub

Public Function ShowSpecChar(frmParent As Object) As String
'���ܣ���ģ̬�������������ַ�����
'������frmParent=���ø�����
'���أ�ѡ��������ַ�����ȡ���������ؿ�
    Dim frmNew As frmSpecChar
    Set frmNew = New frmSpecChar
    frmNew.Show 1, frmParent
    If gblnOK Then ShowSpecChar = frmNew.mstrChar
End Function

Public Function EditMap(frmParent As Object, varԪ�� As Variant, Optional lng����ID As Long, Optional blnViewOnly As Boolean) As Variant
'���ܣ��ڵ�����ģ̬�����б༭��鿴ָ���Ĳ������ͼ
'������frmParnet=���ø�����
'      varԪ��=���ͼԪ�صı���(�ַ���)��ID(������)
'      lng����ID=Ҫ�޸ĵ�"���˲�������"�б��ͼԪ�ض�Ӧ��ID��������������ʾ������ע
'      blnViewOnly=�Ƿ�ֻ�鿴�����ܱ༭
'���أ������ַ������飬����������ϵ�SQL����У�δ���κα��ʱ,������û��Ԫ��(Ubound=-1)
'      ÿ���ַ����ĸ�ʽ�磺"����,'����','����','�㼯',X1,Y1,X2,Y2,���ɫ,��䷽ʽ,����ɫ,����,�߿�"
'      ȡ���༭��鿴ģʽ����Empty(Not isArray)��
    Dim frmNew As frmMapEdit
    Dim rsTmp As New ADODB.Recordset
    Dim arrSQL() As Variant, strSQL As String
    
    Dim objCaseMap As StdPicture, i As Long
    Dim objMapItems As New MapItems, objMapItem As MapItem
    Dim lngMapID As Long, strMapName As String
    
    On Error GoTo errH
        
    '��ȡ���ͼԪ�ص�����
    If TypeName(varԪ��) = "String" Then
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ����=[1]"
    Else
        strSQL = "Select * From ����Ԫ��Ŀ¼ Where ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", CStr(varԪ��))
    If rsTmp.EOF Then Exit Function '����Ҫ��ͼ�α���
    
    lngMapID = rsTmp!ID
    strMapName = rsTmp!���� & IIf(IsNull(rsTmp!˵��), "", "(" & rsTmp!˵�� & ")")
    
    Set objCaseMap = ReadCaseMap(rsTmp!ID)
    If objCaseMap Is Nothing Then Exit Function '����Ҫ��ͼ�α���
    
    '��ȡ���ͼ�ı�ע����
    If lng����ID <> 0 Then
        strSQL = "Select * From ���˲������ͼ Where ����ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "clsCISCore", lng����ID)
        Do While Not rsTmp.EOF
            With rsTmp
                objMapItems.Add !����, zlCommFun.NVL(!����), _
                    IIf(IsNull(!����), IIf(!���� = 0, "����,9,0,0000", ""), !����), _
                    zlCommFun.NVL(!�㼯), zlCommFun.NVL(!X1, 0), _
                    zlCommFun.NVL(!Y1, 0), zlCommFun.NVL(!X2, 0), _
                    zlCommFun.NVL(!Y2, 0), zlCommFun.NVL(!���ɫ, &HFFFFFF), _
                    zlCommFun.NVL(!��䷽ʽ, -1), zlCommFun.NVL(!����ɫ, 0), _
                    zlCommFun.NVL(!����, 0), zlCommFun.NVL(!�߿�, 1), "_" & objMapItems.Count + 1
            End With
            rsTmp.MoveNext
        Loop
    End If
    
    On Error GoTo 0
    
    Set frmNew = New frmMapEdit
    frmNew.mblnModi = Not blnViewOnly
    frmNew.mlngMapID = lngMapID
    frmNew.mstrMapName = strMapName
    Set frmNew.mobjCaseMap = objCaseMap
    Set frmNew.mobjMapItems = objMapItems
    frmNew.Show 1, frmParent
    
    If gblnOK Then
        arrSQL = Array()
        If frmNew.mobjMapItems.Count > 0 Then
            ReDim arrSQL(frmNew.mobjMapItems.Count - 1)
            For i = 1 To frmNew.mobjMapItems.Count
                Set objMapItem = frmNew.mobjMapItems(i)
                '"����,'����','����','�㼯',X1,Y1,X2,Y2,���ɫ,��䷽ʽ,����ɫ,����,�߿�"
                With objMapItem
                    arrSQL(i - 1) = .���� & ",'" & .���� & "','" & .���� & "','" & .�㼯 & "'," & _
                        .X1 & "," & .Y1 & "," & .X2 & "," & .Y2 & "," & .���ɫ & "," & .��䷽ʽ & "," & _
                        .����ɫ & "," & .���� & "," & .�߿�
                End With
            Next
        End If
        EditMap = arrSQL
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowMapInOjbect(objDraw As Object, varԪ�� As Variant, Optional lng����ID As Long, Optional x As Long, Optional y As Long, Optional W As Long, Optional H As Long, Optional blnMoved As Boolean = False)
'���ܣ���ָ���Ķ���(PictureBox��Form)����ʾ���ͼ
'������objDraw=PictureBox�������,����ScaleMode����ΪPixel
'      varԪ��=���ͼԪ�صı���(�ַ���)��ID(������)
'      lng����ID="���˲�������"�б��ͼԪ�ض�Ӧ��ID,�������,����ʾ���ͼ����
'      X,Y,W,H=��ʾ��Ŀ��ͻ��˷�Χ,���Բ�ָ��,��λΪPixel
'˵�����������øú������д�ӡ���(��Ϊ��API��ͼ,��˲���ֱ�ӽ�objDrawָ��Ϊ��ӡ��,������PictureBox�ϰ�һ�����������,ȡPictureBox.Image�������ӡ��)
    Call ShowMapInOjbect_1(objDraw, varԪ��, lng����ID, x, y, W, H, blnMoved)
End Sub

Public Function EditPatientFile(ByVal FileID As String, Optional ByVal sPatientID As String = "", _
    Optional ByVal sPageID As String = "", Optional ByVal iPatientType As Integer = 0, _
    Optional ByVal sTemplateID As String = "", Optional ByVal bSample As Boolean = False, _
    Optional ParentForm As Object, Optional frmNewEdit As Object, Optional AllowEdit As Boolean = True, Optional FileType As Integer = 0, _
    Optional ByVal btModal As Byte = 0, Optional ByVal lngAdviceID As Long = 0) As Long
'���ܣ��༭�����ļ�����ʾ��
'������FileID ������¼ID����ʾ��ID
'      sPatientID ����ID
'      sPageID ��ҳID��Һŵ�ID
'      iPatientType 0=���ﲡ�� 1=סԺ����
'      sTemplateID ����ģ���ļ�ID
'      bSample �Ƿ���ʾ��
'      FileType �������ࣺ1-���ﲡ��;2-סԺ����;3-������¼;4-�������;5-���Ƶ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If Len(FileID) = 0 Then
        strSQL = "Select Count(*)" + _
            " From �����ļ���� Where �����ļ�ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", sTemplateID)
        If rsTmp(0) = 0 Then
            MsgBox "δ���岡���ļ���ɣ����ܱ༭", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        strSQL = "Select Count(*)" + _
            " From ���˲������� Where ������¼ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", FileID)
        If rsTmp(0) = 0 Then
            If Len(sTemplateID) > 0 Then
                strSQL = "Select Count(*)" + _
                    " From �����ļ���� Where �����ļ�ID=[1]"
                Set rsTmp = OpenSQLRecord(strSQL, "zl9CISCore", sTemplateID)
                If rsTmp(0) = 0 Then
                    MsgBox "δ���岡���ļ���ɣ����ܱ༭", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                MsgBox "û�в������ݣ����ܱ༭", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    Set frmNewEdit = New frmFileEdit
    With frmNewEdit
        EditPatientFile = .ShowMe(FileID, sPatientID, sPageID, iPatientType, sTemplateID, bSample, ParentForm, AllowEdit, FileType, btModal, lngAdviceID)
    End With
End Function

Public Sub PreviewPatientFile(objParent As Object, ByVal lng�������� As Long, ByVal blnCurCase As Boolean, ByVal lngCurCase As Long, ByVal lng����ID As Long, _
                        ByVal var��ҳ�򵥾� As Variant, ByVal blnPatiInfo As Boolean, ByVal lngY As Long, Optional ByVal lngҳ�� As Long = 0, Optional ByVal lng��ʼҳ As Long = 0, Optional ByVal lng����ҳ As Long = 0)
'���ܣ���ӡ���в���
'������ObjParent        �����߶���
'       blnCurCase      �Ƿ�Ϊֻ��ӡ�����ǰ��ҳ
'       lngCurCase      ָ����ǰ��ӡ������Ƿݲ�������ӡ���ʱ�ʹ��Ƿ������ӡ���
'                       ����ʱ��ʾ������¼ID
'       lng����id
'       var��ҳ�򵥾�   �����סԺ���˾ͼ�¼��ҳID����������ﲡ�˾ͼ�¼�Һŵ���ͨ�����������ж���סԺ��������
'       blnPatiInfo     �Ƿ��ӡ������Ϣ
'       lngY            ��ӡ��ʼ��Y����
'       lngҳ��         ������ʼ��ҳ��,Ϊ0ʱ��ʾ����ӡҳ��
    Dim frmPreview As New frmCasePrint
    
    PrintOutCase objParent, frmPreview, lng��������, blnCurCase, lngCurCase, lng����ID, var��ҳ�򵥾�, blnPatiInfo, lngY, lngҳ��, lng��ʼҳ, lng����ҳ
    frmPreview.Preview objParent, lng��������, blnCurCase, lngCurCase, lng����ID, var��ҳ�򵥾�, blnPatiInfo, lngY, lngҳ��, lng��ʼҳ, lng����ҳ
End Sub

Public Sub PrintPatientFile(objParent As Object, ByVal lng�������� As Long, ByVal blnCurCase As Boolean, ByVal lngCurCase As Long, ByVal lng����ID As Long, _
                        ByVal var��ҳ�򵥾� As Variant, ByVal blnPatiInfo As Boolean, ByVal lngY As Long, Optional ByVal lngҳ�� As Long = 0, Optional ByVal lng��ʼҳ As Long = 0, Optional ByVal lng����ҳ As Long = 0)
'���ܣ���ӡ���в���
'������ObjParent        �����߶���
'       blnCurCase      �Ƿ�Ϊֻ��ӡ�����ǰ��ҳ
'       lngCurCase      ָ����ǰ��ӡ������Ƿݲ�������ӡ���ʱ�ʹ��Ƿ������ӡ���
'                       ����ʱ��ʾ������¼ID
'       lng����id
'       var��ҳ�򵥾�   �����סԺ���˾ͼ�¼��ҳID����������ﲡ�˾ͼ�¼�Һŵ���ͨ�����������ж���סԺ��������
'       blnPatiInfo     �Ƿ��ӡ������Ϣ
'       lngY            ��ӡ��ʼ��Y����
'       lngҳ��         ������ʼ��ҳ��,Ϊ0ʱ��ʾ����ӡҳ��
    Dim intPage As Integer
    
    intPage = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\��ӡ����", "ֽ��", Printer.PaperSize)
    If IsWindowsNT And intPage = 256 Then DelCustomPaper
    
    If Not InitPrint(objParent) Then
        MsgBox "��ӡ����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    PrintOutCase objParent, Printer, lng��������, blnCurCase, lngCurCase, lng����ID, var��ҳ�򵥾�, blnPatiInfo, lngY, lngҳ��, lng��ʼҳ, lng����ҳ
    'WinNT�Զ���ֽ�Ŵ���
    If IsWindowsNT And intPage = 256 Then DelCustomPaper

    Call InitPrint(objParent)
End Sub

Public Sub EditRequest(frmParent As Object, ByVal lngҽ��ID As Long, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal strҽ������ As String, Optional ByVal ReadOnly As Boolean = False, Optional ByVal ModalWindow As Boolean = True, _
    Optional frmEditForm As Object, Optional ByVal DataMoved As Boolean = False)
'���ܣ��༭���뵥��
'������ModalWindow���༭�����Ƿ�Ϊģʽ����
'      frmEditForm�����ر༭���ڶ���
    On Error Resume Next
    Set frmEditForm = New frmBillEdit
    With frmEditForm
        .ShowMe frmParent, lngҽ��ID, lng����ID, lng����ID, strҽ������, ReadOnly, ModalWindow, DataMoved
    End With
End Sub

Public Function getRequestForm() As Object
'���ܣ�������ʾ���뵥�ݵĴ���
    Set getRequestForm = New frmBillView
End Function

Public Sub AddRequest(frmParent As Object, ByVal lng����ID As Long, ByVal var��ҳ��Һ� As Variant, ByVal lng����ID As Long, _
    ByVal bln��ʿվ As Boolean, Optional ByVal ModalWindow As Boolean = True, Optional frmEditForm As Object, Optional ByVal lngǰ��ID As Long = 0)
'���ܣ���������ҽ�����༭���뵥��
'������ModalWindow���༭�����Ƿ�Ϊģʽ����
'      frmEditForm�����ر༭���ڶ���
    On Error Resume Next
    Set frmEditForm = New frmBillEdit
    With frmEditForm
        .ShowMe_Request frmParent, lng����ID, var��ҳ��Һ�, lng����ID, bln��ʿվ, ModalWindow, lngǰ��ID
    End With
End Sub

Public Sub EditReport(frmParent As Object, _
                        ByVal strNO As String, _
                        ByVal int��¼���� As Integer, _
                        ByVal lng����ID As Long, _
                        ByVal lng����ID As Long, _
                        ByVal strҽ������ As String, _
                        Optional ByVal ReadOnly As Boolean = False, _
                        Optional ByVal ModalWindow As Boolean = True, _
                        Optional frmEditForm As Object, _
                        Optional ByVal lngǰ��ID As Long = 0, _
                        Optional ByVal Ifҽ��ִ�� As Boolean = False, _
                        Optional ByVal blnShow���� As Boolean = True, _
                        Optional ByVal lngҽ��ID As Long = 0, _
                        Optional ByVal DataMoved As Boolean = False, _
                        Optional ByVal strPrivs As String = "000")
'���ܣ��༭����
'������ModalWindow���༭�����Ƿ�Ϊģʽ����
'      frmEditForm�����ر༭���ڶ���
    On Error Resume Next
    Set frmEditForm = New frmBillEdit
    With frmEditForm
        .ShowMe_Report frmParent, strNO, int��¼����, lng����ID, lng����ID, strҽ������, ReadOnly, ModalWindow, lngǰ��ID, Ifҽ��ִ��, blnShow����, lngҽ��ID, DataMoved, strPrivs
    End With
End Sub

Public Sub PrintSetup_Patient(objParent As Object)
'���ܣ�������ӡ����
    frmPrintSet.Show vbModal, objParent
End Sub

Public Function PrintOptionSetup_Patient(objParent As Object, ByVal blnFirst As Boolean, ByRef blnCurCase As Boolean, _
    ByRef blnPatiInfo As Boolean, ByRef lngBeginY As Long, ByRef intBeginPage As Integer, _
    Optional ByVal lng������¼ID As Long, Optional ByRef lng��ʼҳ As Long = 0, Optional ByRef lng����ҳ As Long = 0, Optional ByRef lngPatientID As Long, _
    Optional ByRef lngPageID As Long) As Byte
'���ܣ�������ӡѡ��
    PrintOptionSetup_Patient = frmCasePrintSet.PrintSet(objParent, blnFirst, blnCurCase, _
        blnPatiInfo, lngBeginY, intBeginPage, lng������¼ID, lng��ʼҳ, lng����ҳ, lngPatientID, lngPageID)
End Function

'������ʷ������ѯ
Public Function QryPatientFile(frmParent As Object, Optional ByVal ModalWindow As Boolean = True) As Object
    Set QryPatientFile = frmPatiFileQry
    Call frmPatiFileQry.ShowMe(frmParent, ModalWindow)
End Function

Public Sub ShowDiagHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long)
    '---------------------------------------------
    '���ܣ������ϼ�����Ҫ����ģ̬���ģ̬��ʾ�����ο�
    '��Σ�frmParent-�����壻
    '      blnModal-�Ƿ�ģ̬��ʾ��ͨ�����ϼ�����һ�£���
    '      lngItemId-Ҫ��ʾ�ļ���ID����Ϊ0ʱ��ȱʡ����ʾĿ¼����
    '---------------------------------------------
    Call frmDiagHelp.ShowMe(bytModal, frmParent, lngItemID)
End Sub

Public Sub ShowClinicHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long)
    '---------------------------------------------
    '���ܣ������ϼ�����Ҫ����ģ̬���ģ̬��ʾ���Ʋο�
    '��Σ�frmParent-�����壻
    '      blnModal-�Ƿ�ģ̬��ʾ��ͨ�����ϼ�����һ�£���
    '      lngItemId-Ҫ��ʾ��������ĿID����Ϊ0ʱ��ȱʡ����ʾĿ¼����
    '---------------------------------------------
    Call frmClinicHelp.ShowMe(bytModal, frmParent, lngItemID)
End Sub

Public Sub ShowRcdAnalyse(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngPatiId As Long)
    '---------------------------------------------
    '���ܣ������ϼ�����Ҫ����ģ̬���ģ̬��ʾ���˲�ʷ����
    '��Σ�frmParent-�����壻
    '      blnModal-�Ƿ�ģ̬��ʾ��ͨ�����ϼ�����һ�£���
    '      lngItemId-Ҫ��ʾ�Ĳ���ID
    '---------------------------------------------
    Call frmRcdAnalyse.ShowMe(bytModal, frmParent, lngPatiId)
End Sub

Public Sub ShowDiagnotor(ByVal bytModal As Byte, ByVal frmParent As Object, _
    ByVal lngPatiId As Long, ByVal blnInpati As Boolean, _
    Optional ByVal lngPageID As Long = 1, Optional ByVal strRegistNo As String)
    '---------------------------------------------
    '���ܣ������ϼ�����Ҫ�󣬵��ò��˼���ɸ����򣬹�ҽ���ο�
    '��Σ�frmParent-�����壻
    '      blnModal-�Ƿ�ģ̬��ʾ��ͨ�����ϼ�����һ�£���
    '      lngPatiId-Ҫ��ʾ�Ĳ���ID��
    '      blnInpati-�Ƿ�סԺ���ˣ�����Ϊ���ﲡ�ˣ�
    '      lngPageId-Ҫ��ʾ��סԺ������ҳID��
    '      strRegist-Ҫ��ʾ�����ﲡ�˹Һŵ��ţ�
    '---------------------------------------------
    Call frmDiagnotor.ShowMe(bytModal, frmParent, lngPatiId, blnInpati, lngPageID, strRegistNo)
End Sub

Public Sub EditElementDemo(ByVal frmParent As Object, ByVal blnAdd As Boolean, ByVal lngԪ��Id As Long, Optional ByVal lngDemoID As Long = 0)
    '---------------------------------------------
    '���ܣ��༭Ԫ��ʾ��
    '��Σ�frmParent-�����壻
    '      blnAdd-�Ƿ����ӣ�
    '      lngԪ��Id-Ԫ��ID��
    '      lngDemoID-ʾ����Ӧ�Ĳ�������ID
    '---------------------------------------------
    frmElementDemo.ShowMe frmParent, blnAdd, lngԪ��Id, lngDemoID
End Sub

Public Function RequestRegister(objParent As Object, ByVal lngDeptID As Long, Optional ByVal iItemType As Integer = 1) As Boolean
'���ܣ�ԤԼ����
'���أ�ȷ��/ȡ��
    RequestRegister = frmReqEdit.ShowMe_Request(objParent, lngDeptID, iItemType)
End Function

Public Function ShowFileObject(ByVal frmMain As Object, _
                                    ByVal objContainer As Object, _
                                    ByVal lngҽ��ID As Long, _
                                    ByVal lng����id As Long, _
                                    ByVal cnMain As ADODB.Connection, _
                                    ByVal strDBUser As String, _
                                    ByVal lngSys As Long, _
                                    Optional ByVal strPrivs As String = "", _
                                    Optional ByVal SysName As String = "����") As Object
    Dim lngStyle As Long

    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")

    gstrDBUser = strDBUser

    If Not GetUserInfo Then
        MsgBox "��ǰ�û�δ���ö�Ӧ����Ա��Ϣ,����ϵͳ����Ա��ϵ,�ȵ��û���Ȩ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If

    glngSys = lngSys
'    Set gfrmMain = frmMain
'    gstrUnitName = GetUnitName
    
    Call InitCISCore(gcnOracle)
        
    Set ShowFileObject = New frmFileRequest
    
    Load ShowFileObject
    
    lngStyle = GetWindowLong(ShowFileObject.hWnd, GWL_STYLE)
    
    Call SetWindowLong(ShowFileObject.hWnd, GWL_STYLE, lngStyle Or WS_CHILD)
    Call SetParent(ShowFileObject.hWnd, objContainer.hWnd)
    Call MoveWindow(ShowFileObject.hWnd, 0, 0, objContainer.ScaleWidth / Screen.TwipsPerPixelX, objContainer.ScaleHeight / Screen.TwipsPerPixelY, 1)
            
    '���ù��ܴ�������ˢ��
    
'    Call ShowFileObject.zlRefresh(frmMain, lngҽ��ID, lng����id, strPrivs, SysName)
    
    ShowFileObject.Show
End Function