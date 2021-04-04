VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmDockInTend_File 
   BorderStyle     =   0  'None
   Caption         =   "�ļ�����"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeDockingPane.DockingPane DkpMain 
      Left            =   150
      Top             =   90
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDockInTend_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'######################################################################################################################

Private mintSel As Integer          '��ǰѡ��״̬
Private mfrmTendBody As Object
Private mfrmPartogram As Object
Private WithEvents mfrmTendFile As frmTendFileEditor
Attribute mfrmTendFile.VB_VarHelpID = -1
Public WithEvents zlEvent_Print As zlTFPrintMethod
Attribute zlEvent_Print.VB_VarHelpID = -1

Private mobjParent As Object
Private mblnFirst As Boolean
Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mlngPatiID As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mintBaby As Integer
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean

Private rsTemp As New ADODB.Recordset
Private mfrmMain As Object
Private mblnTendArchive As Boolean
Private mbytFontSize As Byte

Private Enum enuSEL
    ���µ�
    ��¼��
    ����ͼ
End Enum

Public Event Activate()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
Public Event ISChartArchive(ByVal blnArchive As Boolean)
Public Event StartTimer(ByVal blnStart As Boolean)

Public Sub ReSetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С(����ģ���Ѿ����ص���)
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CtlFont As StdFont
    Dim objCtrl As Control
    Dim bytFontSize As Byte
    
    bytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))

    Me.FontSize = bytFontSize
   
    Set CtlFont = DkpMain.PaintManager.CaptionFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set DkpMain.PaintManager.CaptionFont = CtlFont
End Sub

Private Sub Form_Activate()
    If mblnFirst Then
'        mfrmTendBody.Show
'        mfrmTendFile.Show
        mblnFirst = False
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    Dim objPane As Pane
    
    With DkpMain
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.CloseGroupOnButtonClick = True
        .Options.HideClient = True
        
        Set objPane = .CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "�ļ�����": objPane.Options = PaneNoCaption
    End With
    
    '�������²����뻤���¼�������̲���
    If Not CreateBodyEditor Then Exit Sub
    Set mfrmTendBody = gobjBodyEditor.GetTendBody
    Set mfrmTendFile = New frmTendFileEditor
    '�����²�������Ϊ�ޱ��������Ӵ���
    Call FormSetCaption(mfrmTendBody, False, False)
    Call mfrmTendBody.zlInit
    Load mfrmTendBody
    Load mfrmTendFile
    Set mfrmPartogram = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmTendBody Is Nothing Then Unload mfrmTendBody
    If Not mfrmPartogram Is Nothing Then Unload mfrmPartogram
    Unload mfrmTendFile
    Set mfrmPartogram = Nothing
    Set mfrmTendBody = Nothing
    Set mfrmTendFile = Nothing
    Set zlEvent_Print = Nothing
End Sub

Public Sub InitData(ByVal objParent As Object, ByVal strPrivs As String)
    mstrPrivs = strPrivs
    Set mobjParent = objParent
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal blnEdit As Boolean, _
    Optional ByVal blnDoctorStation As Boolean, Optional ByVal intSEL As Integer, Optional ByVal lngKey As Long, Optional ByVal intCurveReSize As Integer = 0) As Long
    Dim blnArchive As Boolean

    'DkpMain.Panes(1).Handle = IIf(intSel = ���µ�, mfrmTendBody.Hwnd, IIf(intSel = ����ͼ, mfrmPartogram.Hwnd, mfrmTendFile.Hwnd))
    mblnDoctorStation = blnDoctorStation
    mblnEdit = blnEdit
    
    If intSEL = 0 Or intSEL = 2 Then
        blnArchive = ArchiveChart(lngKey)
        RaiseEvent ISChartArchive(blnArchive)
        Call mfrmTendFile.SetArchiveData(lngPatiID, lngPageId, intBaby)
        If (InStr(1, ";" & mstrPrivs & ";", IIf(intSEL = 0, ";���µ���ͼ;", ";����ͼ��ͼ;")) > 0 And mblnDoctorStation = False) Then
            If (blnEdit And lngPatiID > 0 And blnArchive = False) Then
                mblnEdit = True
            End If
        End If
    End If
    Select Case intSEL
    Case 0
        If mfrmTendBody Is Nothing Then Exit Function
        mfrmTendBody.Visible = (intSEL = ���µ�)
        DkpMain.Panes(1).Handle = mfrmTendBody.hWnd
        Call mfrmTendBody.zlRefresh(Me, lngPatiID & ";" & lngPageId & ";" & lngDeptID & ";" & lngKey & ";" & IIf(gblnOut = True, 1, 0) & ";" & IIf(mblnEdit = True, 1, 0) & ";" & intBaby & ";" & intCurveReSize, mstrPrivs)
    Case 1
        If mfrmTendFile Is Nothing Then Exit Function
        mfrmTendFile.Visible = (intSEL = ��¼��)
        DkpMain.Panes(1).Handle = mfrmTendFile.hWnd
        Call mfrmTendFile.ShowMe(Nothing, lngKey, lngPatiID, lngPageId, lngDeptID, intBaby, True, mstrPrivs, blnEdit)
    Case 2
        If Not CreatePartogram Then Exit Function
        If mfrmPartogram Is Nothing Then
            Set mfrmPartogram = gobjPartogram.GetPartogram
            Load mfrmPartogram
        End If
        mfrmPartogram.Visible = (intSEL = ����ͼ)
        DkpMain.Panes(1).Handle = mfrmPartogram.hWnd
        Call mfrmPartogram.zlRefresh(Nothing, lngKey & ";" & lngPatiID & ";" & lngPageId & ";" & lngDeptID & ";" & IIf(mblnEdit = True, 1, 0), mstrPrivs)
    End Select
    RaiseEvent StartTimer(intSEL = 0)
End Function

Public Sub ViewReSetFontSize(ByVal intSEL As Integer, ByVal bytSize As Byte)
    Select Case intSEL
        Case 0 '���µ�
            If mfrmTendBody Is Nothing Then Exit Sub
            mfrmTendBody.Visible = (intSEL = ���µ�)
            DkpMain.Panes(1).Handle = mfrmTendBody.hWnd
            'Ŀǰֻ�б�׼�����ɴ˹���
            On Error Resume Next
            Call mfrmTendBody.SetFontSize(bytSize)
            If Err <> 0 Then Err.Clear
        Case 1 '��¼��
            If mfrmTendFile Is Nothing Then Exit Sub
            mfrmTendFile.Visible = (intSEL = ��¼��)
            DkpMain.Panes(1).Handle = mfrmTendFile.hWnd
            Call mfrmTendFile.SetFontSize(bytSize)
        Case 2 '����ͼ
            If Not CreatePartogram Then Exit Sub
            If mfrmPartogram Is Nothing Then
                Set mfrmPartogram = gobjPartogram.GetPartogram
                Load mfrmPartogram
            End If
            mfrmPartogram.Visible = (intSEL = ����ͼ)
            DkpMain.Panes(1).Handle = mfrmPartogram.hWnd
            Call mfrmPartogram.SetFontSize(bytSize)
    End Select
End Sub

Private Sub mfrmTendFile_AfterDataChanged(ByVal blnChange As Boolean)
    RaiseEvent AfterDataChanged(blnChange)
End Sub

Private Sub mfrmTendFile_AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
    RaiseEvent AfterRowColChange(strInfo, blnImportant, blnSign, blnArchive)
End Sub

Public Function zlViewAnimalHeat(ByVal strPara As String, ByVal bytMode As Byte, ByVal strPrivs As String, ByVal bytSize As Byte) As Boolean
    Dim objTendBody As Object
    Dim blnOK As Boolean
    If Not CreateBodyEditor Then Exit Function
    Set objTendBody = gobjBodyEditor.GetNewTendBody
    On Error Resume Next
    objTendBody.Resize = bytSize
    objTendBody.DoctorStation = mblnDoctorStation
    If Err <> 0 Then Err.Clear
    On Error GoTo ErrHand
    blnOK = objTendBody.ShowEdit(Me, strPara, bytMode, strPrivs)
    zlViewAnimalHeat = blnOK
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub zlViewCaveData(ByVal intDataEditor As Integer)
    Call mfrmTendBody.BodyEditCur(intDataEditor)
End Sub

Public Function zlViewFile(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal blnChildForm As Boolean, ByVal strPrivs As String, ByVal blnEdit As Boolean, ByVal bytSize As Byte)
    Dim objTendEditor As New frmTendFileEditor
    Dim blnOK As Boolean
    
    objTendEditor.mstrPrivs = mstrPrivs
    objTendEditor.mblnDoctorStation = mblnDoctorStation
    '51589,������,2013-02-28,ͬ������򿪼�¼���༭���ݹرմ��ڣ�ˢ���Ӵ�������
    blnOK = objTendEditor.ShowMe(Me, lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, blnChildForm, strPrivs, blnEdit, bytSize)
    zlViewFile = blnOK
End Function

Public Sub zlViewpartogram(ByVal strPara As String, ByVal bytMode As Byte, ByVal strPrivs As String, ByVal bytSize As Byte)
    Dim blnOK As Boolean
    Dim objPartogram As Object
    If Not CreatePartogram Then Exit Sub
    If gobjPartogram Is Nothing Then Exit Sub
    If mfrmPartogram Is Nothing Then Exit Sub
    Set objPartogram = gobjPartogram.GetNewPartogram
    blnOK = objPartogram.ShowEdit(Me, strPara & ";" & mfrmPartogram.FileNumIndex, bytMode, strPrivs, bytSize)
    If blnOK = True Then
        Call mfrmPartogram.zlRefresh(Nothing, objPartogram.PartogramParam, strPrivs)
        Call mfrmPartogram.SetFontSize(bytSize)
    End If
End Sub

Public Sub zlViewpartogramEditor(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, _
    ByVal intBaby As Integer, ByVal strPrivs As String, ByVal bytSize As Byte)
    Dim blnChange As Boolean
    '���봫��mfrmPartogram����
    If Not CreatePartogram Then Exit Sub
    If gobjPartogram Is Nothing Then Exit Sub
    If mfrmPartogram Is Nothing Then Exit Sub
    blnChange = gobjPartogram.zlPartogramEditor(mfrmPartogram, lngFileID, lngPatiID, lngPageId, lngDeptID, intBaby, strPrivs, mblnEdit, bytSize)
    If blnChange = True Then
         Call mfrmPartogram.zlRefresh(Nothing, mfrmPartogram.PartogramParam, strPrivs)
         Call mfrmPartogram.SetFontSize(bytSize)
    End If
End Sub

Public Function zlPrintTendFile(ByVal bytKind As Byte, Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String = "") As Long
    '����:  ��ӡ;        'bytMode 1 ��ӡ 2 Ԥ�� 3���Excel
    Dim strSQL As String
    
    If bytKind = 1 Then
        '���µ�(����ֵ:1-�ɹ�;2-��ӡ)
        Call mfrmTendBody.zlPrintBody(bytMode, strPrintDevice)
    ElseIf bytKind = 3 Then
        '����ͼ
        If Not CreatePartogram Then Exit Function
        If gobjPartogram Is Nothing Then Exit Function
        If mfrmPartogram Is Nothing Then Exit Function
        Call gobjPartogram.zlPrintData(mfrmPartogram, bytMode, strPrintDevice)
    Else
        '�����¼��
        Call mfrmTendFile.zlPrintTend(bytMode, strPrintDevice)
    End If
End Function

Public Sub zlPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer, ByVal lngDeptID As Long, ByVal lngFileID As Long, ByVal bytKind As Byte, ByVal strPrintDeviceName As String, ByVal blnPrinter As Boolean)
    '����:��ӡ�����¼���ļ�
    Dim strPrintName As String
    '��ʼ�������
    If bytKind = 1 Then '���µ�
        If Not CreateBodyEditor Then Exit Sub
        Call gobjBodyEditor.zlCurvePrint(2, blnPrinter, 0, -1, strPrintDeviceName, "", lngFileID & ";" & lngPatiID & ";" & lngPageId & ";" & lngDeptID & ";" & intBaby)
    ElseIf bytKind = 3 Then '����ͼ
        If Not CreatePartogram Then Exit Sub
        Call gobjPartogram.PrintPartogram(Me, lngFileID, lngPatiID, lngPageId, lngDeptID, -1, -1, blnPrinter, strPrintDeviceName)
    Else '��¼��
        gstrPrivs = GetPrivFunc(0, 16)
        strPrintName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName)
        If zlEvent_Print Is Nothing Then
            Set zlEvent_Print = New zlTFPrintMethod
        End If
        Call zlEvent_Print.InitPrint(gcnOracle, gstrDBUser)
        Call zlEvent_Print.zlPrintAsk(lngPatiID, lngPageId, intBaby, lngFileID, (strPrintDeviceName <> ""))
        '��ӡ֮ǰ�������ӡ
        If blnPrinter = True Then
            Call zlDatabase.ExecuteProcedure("ZL_���˻����ӡ_CLEAR(0,0,0," & lngFileID & "," & 1 & ")", "�����ӡ����")
        End If
        If strPrintDeviceName <> "" Then Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", strPrintDeviceName)
        zlEvent_Print.zlPrintOrViewTends True, IIf(blnPrinter = True, 1, 2)
        If strPrintDeviceName <> "" Then Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", strPrintName)
    End If
End Sub

Public Sub BulkPrintDocument(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptID As Long, ByVal intBaby As Integer)
    '���������ӡ
    Dim arrFile() As Variant
    Dim i As Integer
    Dim lngFileID As Long, int��� As Integer, intKind As Integer
    Dim strPrintDeviceName As String

    If Printers.Count = 0 Then
        MsgBox "ע�⣺" & Chr(13) _
            & "    δ��װ��ӡ������ͨ��ϵͳ���õĴ�ӡ��" & Chr(13) _
            & "������Ӱ�װ��ӡ����", vbCritical + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    gstrPrivs = GetPrivFunc(0, 16)
    If InStr(1, ";" & gstrPrivs & ";", ";��ӡ;") = 0 Then
        MsgBox "��û�д�ӡȨ�ޣ�����ʹ�øù��ܡ�", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    strPrintDeviceName = Printer.DeviceName
    
    arrFile = Array()
    arrFile = frmNurseFileSelect.ShowMe(lngPatiID, lngPageId, intBaby)
    If UBound(arrFile) = -1 Then Exit Sub
    For i = 0 To UBound(arrFile)
        lngFileID = Val(Split(CStr(arrFile(i)), "_")(0))
        int��� = Val(Split(CStr(arrFile(i)), "_")(1))
        intKind = Val(Split(CStr(arrFile(i)), "_")(2)) + 1
        Select Case intKind
            Case 0  '���µ�
                strPrintDeviceName = zlDatabase.GetPara("���µ���ӡ��", glngSys, 1255, Printer.DeviceName, , True)
                If Not CreateBodyEditor Then Exit Sub
                Call gobjBodyEditor.zlCurvePrint(2, True, 0, 1, strPrintDeviceName, "", lngFileID & ";" & lngPatiID & ";" & lngPageId & ";" & lngDeptID & ";" & int���)
            Case 1 '��¼��
                strPrintDeviceName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName)
                If zlEvent_Print Is Nothing Then
                    Set zlEvent_Print = New zlTFPrintMethod
                End If
                Call zlEvent_Print.zlPrintAsk(lngPatiID, lngPageId, int���, lngFileID, True)
                '��ӡ֮ǰ�������ӡ
                Call zlDatabase.ExecuteProcedure("ZL_���˻����ӡ_CLEAR(0,0,0," & lngFileID & "," & 1 & ")", "�����ӡ����")
                zlEvent_Print.zlPrintOrViewTends (strPrintDeviceName <> ""), 1
            Case 2 '����ͼ
                strPrintDeviceName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\Default", "DeviceName", Printer.DeviceName)
                If Not CreatePartogram Then Exit Sub
                Call gobjPartogram.PrintPartogram(Me, lngFileID, lngPatiID, lngPageId, lngDeptID, -1, -1, True, strPrintDeviceName)
        End Select
    Next i
End Sub

Public Sub SaveData(blnSave As Boolean)
    If blnSave Then
        blnSave = mfrmTendFile.SaveData
    Else
        blnSave = mfrmTendFile.CancelData
    End If
End Sub

Public Sub SignData(blnOK As Boolean, blnVerify As Boolean, blnExchange As Boolean)
    If blnOK Then
        Call mfrmTendFile.SignData(blnVerify, blnExchange)
    Else
        Call mfrmTendFile.UnSignData(blnVerify)
    End If
End Sub

Public Sub ArchiveData(blnOK As Boolean)
    If blnOK Then
        Call mfrmTendFile.ArchiveData
    Else
        Call mfrmTendFile.UnArchiveData
    End If
End Sub

Public Sub SignMarker()
    Call mfrmTendFile.SignMarker
End Sub
