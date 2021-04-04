VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockReport 
   BorderStyle     =   0  'None
   Caption         =   "���Ʊ������"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picRichEdit 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   150
      ScaleHeight     =   2805
      ScaleWidth      =   5055
      TabIndex        =   2
      Top             =   435
      Width           =   5055
      Begin zlRichEditor.Editor edtThis 
         Height          =   1890
         Left            =   30
         TabIndex        =   3
         Top             =   15
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3334
         WithViewButtonas=   0   'False
         ShowRuler       =   0   'False
      End
   End
   Begin VB.PictureBox picNote 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   135
      ScaleHeight     =   375
      ScaleWidth      =   6330
      TabIndex        =   0
      Top             =   45
      Width           =   6330
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   135
         TabIndex        =   1
         Top             =   90
         Width           =   180
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   15
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   135
      Top             =   4710
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'����
'-----------------------------------------------------
Const conPane_Note = 1
Const conPane_Content = 2
Const conPane_Table = 3
Const conPane_Annex = 4

'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event Activate()
Public Event AfterSaved(ByVal lngOrderId As Long, ByVal lngSaveType As Long)
Public Event AfterOpen(ByVal intEditType As EditTypeEnum)
Public Event AfterClosed(ByVal lngOrderId As Long)
Public Event AfterPrinted(ByVal lngOrderId As Long)
Public Event AfterDeleted(ByVal lngOrderId As Long)
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ���߶Ա�����(1258)��Ȩ�޴�
Private mblnSearch As Boolean   '��ǰʹ�����Ƿ�߱���������(1273)Ȩ

Private mlngOrderId     As Long         'ҽ��id
Private mblnMoved       As Boolean      '�Ƿ�ת��
Private mblnCanPrint    As Boolean      '�ɷ��ӡ
Private mintPati��Դ    As Long
Private mlngPati����ID  As Long
Private mlngPati��ҳID  As Long
Private mlngPatiӤ��    As Long
Private mlngEPR����ID   As Long
Private mlngEPR����ID   As Long
Private mstrEPR�������� As String
Private mstrEPR������   As String
Private mstrEPR������   As String
Private mstrEPR�鵵��   As String
Private mstrEPR���ʱ�� As String
Private mintEPRǩ������ As Integer
Private mintEPRǩ���汾 As Integer
Private mintEPR���汾 As Integer
Private mlngEPR����ID   As Long
Private mbyeEPR�༭��ʽ As Byte
Private mlngSingCount   As Long

Private mlngDeptId As Long          '��ǰ��������id
Private mblnEdit As Boolean         '�Ƿ��������
Private mlngModule As Long

Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR
Private mObjTabEprView As cTableEPR
Private mfrmAnnex As frmDockAnnex    '������������
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1

Private mstrPrinterDeviceName As String
Private mlngPrintCopies As Long

Dim mcbsThis As Object          'CommandBar�ؼ�


'------------------------------------------------------------
'����Ϊ��������
'------------------------------------------------------------
Public Sub zlDefCommandBars(ByVal cbsThis As Object)
    '-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
    Set mcbsThis = cbsThis
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    
    '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If cbrMenuBar Is Nothing Then
        Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����(&E)", cbrMenuBar.Index + 1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "��д(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�޶�(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "����(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����XML(&L)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "�������(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
    End With
    
    '����������
    '-----------------------------------------------------
    Set cbrToolBar = cbsThis(2)
    For Each cbrControl In cbrToolBar.Controls '�����ǰ������һ��Control
        If Val(Left(cbrControl.ID, 1)) <> conMenu_FilePopup And Val(Left(cbrControl.ID, 1)) <> conMenu_ManagePopup Then
            Set cbrControl = cbrToolBar.Controls(cbrControl.Index - 1): Exit For
        End If
    Next
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "��д", cbrControl.Index + 1): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�޶�", cbrControl.Index + 1)
    End With

    '����Ŀ����
    '-----------------------------------------------------
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("U"), conMenu_Edit_Audit
'        .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
    End With
    
    '���ò���������
    '-----------------------------------------------------
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_ExportToXML
        .AddHiddenCommand conMenu_Tool_Search
    End With
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
    If mblnMoved And (Control.ID = conMenu_File_Open Or Control.ID = conMenu_File_ExportToXML Or _
                    Control.ID = conMenu_Edit_Modify Or Control.ID = conMenu_Edit_Delete Or Control.ID = conMenu_Edit_Audit) Then
        MsgBox "�ò��˵ı��������Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                    "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_Open
        '�����Ķ�
        If mbyeEPR�༭��ʽ = 0 Then
            Dim fViewDoc As New frmEPRView
            fViewDoc.ShowMe Me, mlngEPR����ID
        Else
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_�������༭, mlngEPR����ID, True, 0, mintPati��Դ, mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved, , Val(gstrESign))
        End If
    Case conMenu_File_PrintSet
        Call zlPrintSet
    Case conMenu_File_Preview
        If mblnCanPrint Then
            Call zlEPRPrint(True)
        Else
            MsgBox "��ǰ����δ��ˣ����ܴ�ӡ�����飡", vbInformation, gstrSysName
        End If
    Case conMenu_File_Print
        If mblnCanPrint Then
            Call zlEPRPrint(False)
        Else
            MsgBox "��ǰ����δ��ˣ����ܴ�ӡ�����飡", vbInformation, gstrSysName
        End If
    Case conMenu_File_BatPrint
        If mblnCanPrint Then
            Call zlEPRPrint(False, True)
        Else
            MsgBox "��ǰ����δ��ˣ����ܴ�ӡ�����飡", vbInformation, gstrSysName
        End If
    Case conMenu_File_ExportToXML:
        '������XML�ļ�
        Dim strF As String
        dlgThis.Filename = "����_" & mstrEPR�������� & "(" & mlngEPR����ID & "," & mintEPR���汾 & ").xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        strF = dlgThis.Filename
        If Err.Number <> 0 Then Err.Clear: Exit Sub
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If mbyeEPR�༭��ʽ = 0 Then
            Dim DocXML As New cEPRDocument
            '��ͨסԺ����
            DocXML.InitAndOpenEPR mlngEPR����ID, mintEPR���汾, , True
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else    '���ʽ����
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_�������༭, mlngEPR����ID, False, 0, mintPati��Դ, mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved)
            If mObjTabEprView.zlExportXML(strF) Then
                MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        End If
    Case conMenu_Edit_Modify
        If CheckCommitCheckup = False Then Exit Sub '��Ժ���˲����ύ�����ֹ�޸�
        Dim frmThis As Form, bFinded As Boolean
        If mbyeEPR�༭��ʽ = 1 Then '���ʽ����
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(mlngEPR����ID, mlngPati����ID, mlngPati��ҳID, mintPati��Դ, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.EPRFileInfo.lngModule = mlngModule
                mObjTabEpr.InitOpenEPR Me, IIf(mlngEPR����ID = 0, cprEM_����, cprEM_�޸�), cprET_�������༭, IIf(mlngEPR����ID = 0, mlngEPR����ID, mlngEPR����ID), True, 0, mintPati��Դ, _
                    mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved, , Val(gstrESign)
                RaiseEvent AfterOpen(cprET_�������༭)
            End If
            mlngSingCount = mObjTabEpr.Signs.Count
        Else
            For Each frmThis In Forms
                If frmThis.Name = "frmMain" Then
                    If Not frmThis.Document Is Nothing Then
                        If frmThis.Document.EPRPatiRecInfo.ҽ��id = mlngOrderId And frmThis.ChildMode = False Then
                            frmThis.Show
                            bFinded = True
                        End If
                    Else
                        Unload frmThis
                    End If
                End If
            Next
            If bFinded = False Then
                strInfo = Clipboard.GetText '�ݴ�
                Set mobjDoc = New cEPRDocument
                If mlngEPR����ID = 0 Then
                    mobjDoc.InitEPRDoc cprEM_����, cprET_�������༭, mlngEPR����ID, _
                        mintPati��Դ, mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mblnMoved
                Else
                    mobjDoc.InitEPRDoc cprEM_�޸�, cprET_�������༭, mlngEPR����ID, _
                        mintPati��Դ, mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mblnMoved
                End If
                
                mobjDoc.EPRFileInfo.lngModule = mlngModule
                
                mobjDoc.ShowEPREditor Me, mblnCanPrint
                If Trim(strInfo) <> "" Then '�ָ�ճ��������
                    DoEvents
                    Clipboard.SetText strInfo
                End If
                RaiseEvent AfterOpen(cprET_�������༭)
            End If
            mlngSingCount = mobjDoc.Signs.Count
        End If
        
    Case conMenu_Edit_Delete
        strInfo = "���ɾ����ݡ�" & mstrEPR�������� & "����"
        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_���Ӳ�����¼_Delete(" & mlngEPR����ID & ")"
        Err = 0: On Error GoTo errHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Err = 0: On Error GoTo 0
        RaiseEvent AfterDeleted(mlngOrderId)
        Call Me.zlRefresh(mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule)
    
    Case conMenu_Edit_Audit
        If CheckCommitCheckup = False Then Exit Sub '��Ժ���˲����ύ�����ֹ�޸�
        If mbyeEPR�༭��ʽ = 1 Then '���ʽ����
            If Not mObjTabEpr Is Nothing Then
                bFinded = mObjTabEpr.Showfrm(mlngEPR����ID, mlngPati����ID, mlngPati��ҳID, mintPati��Դ, mlngDeptId)
            End If
            If bFinded = False Then
                Set mObjTabEpr = New cTableEPR
                mObjTabEpr.EPRFileInfo.lngModule = mlngModule
                mObjTabEpr.InitOpenEPR Me, cprEM_�޸�, cprET_���������, mlngEPR����ID, True, 0, mintPati��Դ, _
                    mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved, , Val(gstrESign)
                RaiseEvent AfterOpen(cprET_���������)
            End If
            mlngSingCount = mObjTabEpr.Signs.Count
        Else
            Dim frmAudit As Form, bFindedAudit As Boolean
            For Each frmAudit In Forms
                If frmAudit.Name = "frmMain" Then
                    If Not frmAudit.Document Then
                        If frmAudit.Document.EPRPatiRecInfo.ҽ��id = mlngOrderId And frmAudit.ChildMode = False Then
                            frmAudit.Show
                            bFindedAudit = True
                        End If
                    Else
                        Unload frmAudit
                    End If
                End If
            Next
            If bFindedAudit = False Then
                Set mobjDoc = New cEPRDocument
                mobjDoc.InitEPRDoc cprEM_�޸�, cprET_���������, mlngEPR����ID, _
                    mintPati��Դ, mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId
                    
                mobjDoc.EPRFileInfo.lngModule = mlngModule
                
                mobjDoc.ShowEPREditor Me
                RaiseEvent AfterOpen(cprET_���������)
            End If
            mlngSingCount = mobjDoc.Signs.Count
        End If
        
    Case conMenu_Edit_Copy
        Call edtThis.Copy
    Case conMenu_Tool_Search: frmEPRSearchMan.ShowSearchReport Me, mlngDeptId
    Case conMenu_View_Refresh:  Call Me.zlRefresh(mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule)
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Tool_SignVerify
        Call VerifySignature(Me, mlngEPR����ID, mblnMoved)
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    Select Case Control.ID
    Case conMenu_File_Open
        Control.Enabled = (Val(mlngEPR����ID) <> 0)
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_ExportToXML
        Control.Enabled = (Val(mlngEPR����ID) <> 0 And InStr(1, mstrPrivs, "�����ӡ") > 0)
    Case conMenu_Edit_Modify
        Control.Enabled = (mblnEdit And mlngOrderId > 0 And InStr(1, mstrPrivs, "������д") > 0)
        If Control.Enabled And mlngEPR����ID > 0 Then
            If Control.Enabled Then Control.Enabled = (mlngDeptId = mlngEPR����ID)   '���Ʋ����ſ��Ը�
            If mstrEPR���ʱ�� = "" Then
                Control.Enabled = (InStr(1, mstrPrivs, "���˱���") > 0 Or mstrEPR������ = Trim(gstrUserName))
            ElseIf mstrEPR�鵵�� = "" And mintEPR���汾 <= 1 And InStr(1, ",1,2,4,", mintEPRǩ������) > 0 Then
                Control.Enabled = (InStr(1, mstrPrivs, "���˱���") > 0 Or InStr(1, mstrEPR������, Trim(gstrUserName)) > 0)
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_Delete
        Control.Enabled = (mblnEdit And mlngEPR����ID <> 0) And (InStr(1, mstrPrivs, "������д") > 0 Or InStr(1, mstrPrivs, "ǿ��ɾ��") > 0)
        If Control.Enabled And InStr(1, mstrPrivs, "ǿ��ɾ��") > 0 Then Exit Sub                '�߱�ǿ��ɾ��Ȩ�ޣ��򲻽��к������ж�
        If Control.Enabled Then Control.Enabled = (mlngDeptId = mlngEPR����ID)     '���Ʋ����ſ���ɾ
        If Control.Enabled Then Control.Enabled = (mstrEPR���ʱ�� = "")                'δ��ɲ�������ɾ
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "���˱���") > 0 Or mstrEPR������ = Trim(gstrUserName))
    
    Case conMenu_Edit_Audit
        Control.Enabled = (mblnEdit And mlngPati����ID > 0 And InStr(1, mstrPrivs, "�����޶�") > 0)
        If Control.Enabled Then Control.Enabled = (mlngDeptId = mlngEPR����ID)      '���Ʋ����ſ������
        If Control.Enabled Then Control.Enabled = (mstrEPR���ʱ�� <> "")           '��ɲ����ſ�����
        If Control.Enabled Then Control.Enabled = (mstrEPR�鵵�� = "")              'δ�鵵����������
    Case conMenu_Tool_Search
        Control.Enabled = mblnSearch
    Case conMenu_Edit_Copy
        Control.Visible = mbyeEPR�༭��ʽ = 0
        If Control.Visible Then Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
    End Select

End Sub
Public Sub RefreshList()
    zlRefresh mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule
End Sub
Public Function zlRefresh(ByVal lngOrderId As Long, ByVal lngDeptId As Long, ByVal blnEdit As Boolean, _
                            Optional ByVal blnForce As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal blnCanPrint As Boolean = True, Optional ByVal lngModule As Long) As Long
'�쳣����0,���򷵻�1
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    strTemp = ""
    
    If mlngDeptId <> lngDeptId Or gstrESign = "" Then '��ȡ�Ƿ񱾲������õ���ǩ��,���ұ����ûȡ��ʱ��ȡ
        gstrESign = getPassESign(7, lngDeptId)
    End If
    
    mlngDeptId = lngDeptId: mblnEdit = blnEdit
    If mlngOrderId = lngOrderId And blnForce = False Then Exit Function
    mlngOrderId = lngOrderId: mblnMoved = blnMoved: mblnCanPrint = blnCanPrint: mlngModule = lngModule
    
    
    Err = 0: On Error GoTo errHand
    
    mintPati��Դ = 0: mlngPati����ID = 0: mlngPati��ҳID = 0: mlngPatiӤ�� = 0
    gstrSQL = "Select l.������Դ, l.����id, l.�Һŵ�, l.��ҳid, l.Ӥ��, a.�����ļ�id" & vbNewLine & _
            "From ����ҽ����¼ l, ��������Ӧ�� a" & vbNewLine & _
            "Where l.������Ŀid = a.������Ŀid(+) And a.Ӧ�ó���(+) = Decode(l.������Դ, 2, 2, 4, 4, 1) And l.Id = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOrderId)
    With rsTemp
        If .RecordCount > 0 Then
            mintPati��Դ = Val("" & !������Դ)
            mlngPati����ID = Val("" & !����ID)
            If mintPati��Դ <> 1 Then
                mlngPati��ҳID = Val("" & !��ҳID)
            Else
                strTemp = "" & !�Һŵ�
            End If
            mlngPatiӤ�� = Val("" & !Ӥ��)
            mlngEPR����ID = Val("" & !�����ļ�id)
        End If
    End With
    
    If mlngEPR����ID <> 0 Then
        gstrSQL = "Select ���� From �����ļ��б� where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngEPR����ID)
        mbyeEPR�༭��ʽ = IIf(NVL(rsTemp!����, 0) = 2, 1, 0)
    Else
        mbyeEPR�༭��ʽ = 0
    End If
    
    If mintPati��Դ = 1 Then
        gstrSQL = "Select ID From ���˹Һż�¼ Where NO = [1] and ��¼����=1  and ��¼״̬=1"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "���˹Һż�¼", "H���˹Һż�¼")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
        If rsTemp.RecordCount > 0 Then
            mlngPati��ҳID = rsTemp!ID
        End If
    End If

    mlngEPR����ID = 0: mstrEPR�������� = "": mstrEPR������ = "": mstrEPR������ = "": mstrEPR�鵵�� = ""
    mlngEPR����ID = 0: mstrEPR���ʱ�� = "": mintEPRǩ������ = 0: mintEPRǩ���汾 = 0: mintEPR���汾 = 1
    Me.lblNote.Caption = "��ʾ�� ��δ��д���棡"
    gstrSQL = "Select l.Id, l.��������, l.������, l.������, l.���ʱ��, l.���汾, l.ǩ������, l.�鵵��, l.����id, l.������," & vbNewLine & _
            "       Nvl(Max(c.��ʼ��), 0) As ǩ���汾,l.�༭��ʽ" & vbNewLine & _
            "From ���Ӳ�����¼ l, ���Ӳ������� c, ����ҽ������ r" & vbNewLine & _
            "Where l.Id = c.�ļ�id(+) And l.Id = r.����id And l.�������� = 7 And c.��������(+) = 8 And r.ҽ��id = [1]" & vbNewLine & _
            "Group By l.Id, l.��������, l.������, l.������, l.���ʱ��, l.���汾, l.ǩ������, l.�鵵��, l.����id, l.������,�༭��ʽ"
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
        gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOrderId)
    With rsTemp
        If Not .EOF Then
            mlngEPR����ID = !ID
            mstrEPR�������� = "" & !��������
            mstrEPR������ = "" & !������
            mstrEPR������ = "" & !������
            mstrEPR���ʱ�� = Trim("" & Format(!���ʱ��, "yyyy��MM��dd��hhʱmm��"))
            mintEPR���汾 = Val("" & !���汾)
            mintEPRǩ������ = Val("" & !ǩ������)
            mstrEPR�鵵�� = "" & !�鵵��
            mintEPRǩ���汾 = Val("" & !ǩ���汾)
            mlngEPR����ID = Val("" & !����ID)
            
            If mstrEPR���ʱ�� = "" Then
                Me.lblNote.Caption = "��ʾ����ǰ����������" & mstrEPR������ & "��д����δ��ɡ�"
            ElseIf mintEPRǩ������ = mintEPR���汾 Then
                Me.lblNote.Caption = "��ʾ�����������" & mstrEPR���ʱ�� & "��" & !������ & IIf(mintEPR���汾 = 1, "��д", "����޶���")
            Else
                Me.lblNote.Caption = "��ʾ�����������" & mstrEPR���ʱ�� & "��" & !������ & IIf(mintEPR���汾 = 1, "��д", "�����޶���")
            End If
            
            mbyeEPR�༭��ʽ = !�༭��ʽ '�б���ʱ�Ա���ʵ�ʱ༭��ʽΪ׼
        End If
        '������ʾ�ĵ�
        If mbyeEPR�༭��ʽ = 1 And mlngEPR����ID <> 0 Then '���ʽ���������Ѿ�д������
            With edtThis
                .Text = vbCrLf & Space(4) & "���ļ�Ϊ���ʽ���������ڼ�����..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "����": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
            End With
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_�޸�, cprET_���������, mlngEPR����ID, False, 0, mintPati��Դ, mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved)
            Call mObjTabEprView.zlRefreshDockfrm
            
            dkpMan.FindPane(conPane_Content).Close
            dkpMan.ShowPane conPane_Table
            dkpMan.RedrawPanes
        Else
            Call zlRefDocment(mlngEPR����ID)
            
            dkpMan.FindPane(conPane_Table).Close
            dkpMan.ShowPane conPane_Content
            dkpMan.RedrawPanes
        End If
        Call mfrmAnnex.zlRefresh(mlngEPR����ID, IIf(mblnEdit, mstrPrivs, ""))
    End With
    If mlngEPR����ID <> 0 Then zlRefresh = 1
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlRefDocment(ByVal lngEPRid As Long)
    '���ܣ�ˢ�²�����ʾ���ݣ�
    '������lngEPRId-���Ӳ�����¼ID
    Dim mstrPrivs As String, blnPrivacy As Boolean, Elements As New cEPRElements
    Dim rs As New ADODB.Recordset, lngKey As Long
    
    Dim strTemp As String, strZipFile As String
    
    Me.edtThis.Freeze
    Me.edtThis.ReadOnly = False
    Me.edtThis.NewDoc
    strZipFile = zlBlobRead(5, lngEPRid, , mblnMoved)
    If gobjFSO.FileExists(strZipFile) Then
        strTemp = zlFileUnzip(strZipFile)
        If gobjFSO.FileExists(strTemp) Then
            '���ļ�
            Me.edtThis.OpenDoc strTemp
            gobjFSO.DeleteFile strTemp, True
        End If
        gobjFSO.DeleteFile strZipFile, True
        Me.edtThis.SelStart = 0
    End If
    If lngEPRid > 0 Then
        '����ҳ���ʽ
        Dim mEPRFileInfo As New cEPRFileDefineInfo
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select c.ID, a.��ʽ From   ����ҳ���ʽ a, �����ļ��б� b, ���Ӳ�����¼ c " & _
                " Where  c.�ļ�id = b.id And a.���� = b.���� And a.��� = b.ҳ�� And c.ID = [1]"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
        If Not rs.EOF Then
            mEPRFileInfo.��ʽ = zlCommFun.NVL(rs("��ʽ").Value)
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
            Me.edtThis.ResetWYSIWYG
        End If
        Set mEPRFileInfo = Nothing
    End If
    Me.edtThis.UnFreeze
    edtThis.RefreshTargetDC
    Me.edtThis.ReadOnly = True
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub ConfigPrint(ByVal strPrintDevice As String, ByVal lngCopies As Long)
'���ô�ӡ��
    mstrPrinterDeviceName = strPrintDevice
    mlngPrintCopies = lngCopies
End Sub


Private Sub zlEPRPrint(blnPreview As Boolean, Optional blnStilly As Boolean)
    '-------------------------------------------------
    '����: ��ӡ��ǰ�ĵ�
    '����:  blnPreview Ԥ��
    '       blnStilly  ǿ�ƾ�Ĭ��ӡ
    '-------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intOutMode As Integer, strBillNo As String, blnNoAsk As Boolean
    blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")
    If blnStilly Then blnNoAsk = True
    
    If Trim(mstrPrinterDeviceName) = "" Then
        mstrPrinterDeviceName = Printer.DeviceName
        mlngPrintCopies = Printer.Copies
    End If
    
    intOutMode = 0: strBillNo = ""
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select f.ͨ��, f.��� From ���Ӳ�����¼ l, �����ļ��б� f Where l.�ļ�id = f.Id And l.Id = [1]"
    If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�����¼", "H���Ӳ�����¼")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngEPR����ID)
    If rsTemp.RecordCount > 0 Then
        intOutMode = Val("" & rsTemp!ͨ��)
        strBillNo = "ZLCISBILL" & Format(rsTemp!���, "00000") & "-2"
    End If
    
    If intOutMode <> 2 Then
        If mbyeEPR�༭��ʽ = 0 Then 'RichEpr�༭
            'ֱ�Ӵ�ӡ
            Set mfrmPrintPreview = New frmPrintPreview
            Call mfrmPrintPreview.DoMultiDocPreview(Me, cpr���Ʊ���, , , cpr���Ʊ���, , mlngEPR����ID, Not blnPreview, False, blnNoAsk, mblnMoved, , mstrPrinterDeviceName, mlngPrintCopies)
            Unload mfrmPrintPreview: Set mfrmPrintPreview = Nothing
        ElseIf mlngEPR����ID <> 0 Then '���ʽ����
            mObjTabEprView.InitOpenEPR Me, cprEM_�޸�, cprET_�������༭, mlngEPR����ID, False, 0, mintPati��Դ, mlngPati����ID, mlngPati��ҳID, mlngPatiӤ��, mlngDeptId, mlngOrderId, mstrPrivs, mblnMoved
            mObjTabEprView.zlPrintDoc Me, blnPreview
        End If
    Else
        '�Զ��屨���ӡ
        Dim strExseNo As String, intExseKind As Integer
        Dim objFile As New Scripting.FileSystemObject
        Dim strPicPath As String, strPicFile As String
        Dim cTable As cEPRTable, oPicture As StdPicture
        Dim aryPara(19) As String, intPCount As Integer
        Dim aryFlagPara(1) As String
        Dim intRows As Integer, intCols As Integer
        Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
        Dim i As Integer
        
        gstrSQL = "Select ��¼����, No From ����ҽ������ Where ҽ��id = [1]"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngOrderId)
        If rsTemp.RecordCount = 0 Then Exit Sub
        strExseNo = "" & rsTemp!NO
        intExseKind = Val("" & rsTemp!��¼����)
        If mobjReport Is Nothing Then Set mobjReport = New clsReport
        If Not blnNoAsk Then
            If mobjReport.ReportPrintSet(gcnOracle, glngSys, strBillNo, Me) = False Then Exit Sub
        End If
        
        '��ȡͼ��
        strPicPath = App.Path & "\TmpImage\"
        If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
        
        '��ȡ����ͼ��(�������ͼ)���ɱ����ļ�
        'һ���������п������ж������ͼ
        intPCount = 0
        gstrSQL = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
        If mblnMoved Then gstrSQL = Replace(gstrSQL, "���Ӳ�������", "H���Ӳ�������")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngEPR����ID)
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If cTable.GetTableFromDB(cprET_���������, mlngEPR����ID, Val("" & rsTemp!���Id), , IIf(mblnMoved, "H���Ӳ�������", "���Ӳ�������")) Then
                For i = 1 To cTable.Pictures.Count
                    strPicFile = strPicPath & "PACSPic" & i & ".JPG"
                    If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        Set oPicture = cTable.Pictures(i).DrawFinalPic
                    Else
                        Set oPicture = cTable.Pictures(i).OrigPic
                    End If
                    SavePicture oPicture, strPicFile
                    If objFile.FileExists(strPicFile) Then
                        '������ͼ��ͼ���·��
                        If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                            aryFlagPara(0) = strPicFile
                        Else
                            aryPara(intPCount) = strPicFile
                            dcmImages.AddNew
                            dcmImages(dcmImages.Count).FileImport strPicFile, "BMP"
                            intPCount = intPCount + 1
                            If intPCount > UBound(aryPara) Then Exit Do
                        End If
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
        
        '�ж��Ƿ���Ҫ�Զ����ͼ���Զ��屨����ֻ������һ��ͼ������Զ����ͼ��
        '���²�һ�����ݿ�
        gstrSQL = "Select b.����,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = 1 And b.���� not like '���%'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strBillNo)
        If rsTemp.RecordCount = 1 And intPCount >= 1 Then
            '���ͼ��
            ResizeRegion intPCount, rsTemp("W"), rsTemp("H"), intRows, intCols
            Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
            dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
        End If
        
        '��ȡ�Զ��屨���е�ͼ����
        intPCount = 0
        gstrSQL = "Select b.���� From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = 1" & vbNewLine & _
        "       Order By b.����" 'Trunc(b.y/567),Trunc(b.x/567)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strBillNo)
        Do While Not rsTemp.EOF
            If aryPara(intPCount) = "" Then Exit Do '�����е�ͼ�αȱ����ж�
            '�ֱ�װ�ر��ͼ�ͱ���ͼ��
            If InStr(rsTemp!����, "���") <> 0 Then
                If aryFlagPara(0) <> "" Then aryFlagPara(0) = rsTemp!���� & "=" & aryFlagPara(0)
            Else
                aryPara(intPCount) = rsTemp!���� & "=" & aryPara(intPCount)
                intPCount = intPCount + 1
                If intPCount > UBound(aryPara) Then Exit Do
            End If
            rsTemp.MoveNext
        Loop
        For i = intPCount To UBound(aryPara) '�����е�ͼ�αȱ�������
            If aryPara(i) Like "*=*" Then aryPara(i) = ""
        Next
        
        '���ñ���
       Call mobjReport.ReportOpen(gcnOracle, glngSys, strBillNo, Nothing, _
            "NO=" & strExseNo, "����=" & intExseKind, "ҽ��ID=" & mlngOrderId, aryFlagPara(0), _
            aryPara(0), aryPara(1), aryPara(2), aryPara(3), aryPara(4), aryPara(5), _
            aryPara(6), aryPara(7), aryPara(8), aryPara(9), aryPara(10), aryPara(11), _
            aryPara(12), aryPara(13), aryPara(14), aryPara(15), aryPara(16), aryPara(17), _
            aryPara(18), aryPara(19), IIf(blnPreview, 1, 2))
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Sub RefPacsPic()
'����: ˢ�����ڱ༭�����PACSͼƬ
    If mbyeEPR�༭��ʽ = 0 Then
        Dim frmThis As Form
        For Each frmThis In Forms
            If frmThis.Name = "frmMain" Then
                If Not frmThis.Document Then
                    With frmThis.Document
                        If .EPRPatiRecInfo.ҽ��id = mlngOrderId Then
                            Call frmThis.RefPacsPic
                        End If
                    End With
                End If
                
                Exit Sub
            End If
        Next
    Else
        mObjTabEpr.zlRefreshPacsPic mlngOrderId
    End If
End Sub

'------------------------------------------------------------
'����Ϊ�����¼���Ӧ
'------------------------------------------------------------
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Note
        Item.Handle = picNote.hWnd
    Case conPane_Content
        Item.Handle = picRichEdit.hWnd
    Case conPane_Table
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Annex
        Item.Handle = mfrmAnnex.hWnd
    End Select
End Sub

Private Sub edtThis_KeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = vbKeyC Then
        Call edtThis.Copy
    End If
End Sub

Private Sub Form_Load()
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, Pane4 As Pane
    Set mfrmAnnex = New frmDockAnnex
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    mblnSearch = (InStr(1, GetPrivFunc(glngSys, 1273), "����") > 0)
    mstrPrivs = GetPrivFunc(glngSys, 1258)
    
    Set Pane1 = dkpMan.CreatePane(conPane_Note, 200, 15, DockTopOf, Nothing)
    Pane1.Title = "��ʾ": Pane1.MinTrackSize.Height = 360 / Screen.TwipsPerPixelY: Pane1.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_Content, 1200, 200, DockBottomOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane3 = dkpMan.CreatePane(conPane_Table, 1200, 200, DockBottomOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    Set Pane4 = dkpMan.CreatePane(conPane_Annex, 200, 15, DockBottomOf, Nothing)
    Pane4.Title = "����": Pane4.MinTrackSize.Height = 360 / Screen.TwipsPerPixelY: Pane4.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
    Pane4.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With dkpMan
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mObjTabEprView.zlGetForm
    Unload mfrmAnnex
    Unload mfrmPrintPreview
    
    Set mfrmAnnex = Nothing
    Set mobjReport = Nothing
    Set mfrmPrintPreview = Nothing
    Set mobjDoc = Nothing
    Set mObjTabEpr = Nothing
    Set mObjTabEprView = Nothing
    Set mcbsThis = Nothing
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    If mcbsThis Is Nothing Then Exit Sub
Dim Popup As CommandBar
Dim cbrControl As CommandBarControl
    
    Set Popup = mcbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "��д(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�޶�(&U)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "����(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "����XML(&L)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "�������(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
        Popup.ShowPopup
    End With
End Sub
Public Sub EditorClosed(lngOrderId As Long)
    RaiseEvent AfterClosed(lngOrderId)
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
    Call Event_AfterPrinted(lngRecordId)
End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)
    Dim rsTemp As New ADODB.Recordset, lngҽ��id As Long
    Dim lngSaveType As Long
        
    '����ǵ�ǰ���棬��ˢ�µ�ǰ��ʾ����
    If lngRecordId = mlngEPR����ID Then
        Call Me.zlRefresh(mlngOrderId, mlngDeptId, mblnEdit, True, mblnMoved, mblnCanPrint, mlngModule)
    End If
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ҽ��id,����״̬ From ����ҽ������ Where ����Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.RecordCount > 0 Then
        lngҽ��id = Val("" & rsTemp!ҽ��id)
        'ȡ�������ע�ͣ����ݵ��õ�ģ����д�ù��̵�Ȩ�޽ű�
'        If Val("" & rsTemp!����״̬) = 1 Then
'            gstrSQL = "Zl_������ļ�¼_Cancel(" & lngҽ��ID & "," & lngRecordId & ",Null)"
'            Call zldatabase.ExecuteProcedure(gstrSQL, "���²���״̬")
'        End If

        '�����Ǳ�����༭������ȫ�Ĳ����༭��
        If mbyeEPR�༭��ʽ = 1 Then '���ʽ����
            If mlngSingCount = mObjTabEpr.Signs.Count Then
                lngSaveType = 0 '��ͨ����
            Else
                If mObjTabEpr.ET <> TabET_��������� Then
                    lngSaveType = 1 '���ǩ��
                Else
                    lngSaveType = 2 '���ǩ��
                End If
            End If
        Else    'ȫ�Ĳ����༭��
            If mlngSingCount = mobjDoc.Signs.Count Then
                lngSaveType = 0 '��ͨ����
            ElseIf mlngSingCount < mobjDoc.Signs.Count Then
                If mobjDoc.Signs(mobjDoc.Signs.Count).ǩ������ > cprSL_���� Then
                    lngSaveType = 2 '���ǩ��
                Else
                    lngSaveType = 1 '���ǩ��
                End If
            End If
            
            mlngSingCount = mobjDoc.Signs.Count
        End If
        
        RaiseEvent AfterSaved(lngҽ��id, lngSaveType)
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Public Sub Event_Saved(lngRecordId As Long)
    mobjDoc_AfterSaved lngRecordId
End Sub
Public Sub Event_AfterPrinted(lngRecordId As Long)
    Dim rsTemp As New ADODB.Recordset
    
    Err.Clear
    If mblnMoved Then Exit Sub 'ת�����Ĳ���,��������ӡ�¼�,Ŀǰ�Ĵ�ӡ�¼�ֻ��ΪӰ�����¼��Ǳ�־
    On Error GoTo errHand
    gstrSQL = "Select ҽ��id From ����ҽ������ Where ����Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    If rsTemp.RecordCount > 0 Then
        RaiseEvent AfterPrinted(NVL(rsTemp!ҽ��id, 0))
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    RaiseEvent AfterPrinted(mlngOrderId)
End Sub
Private Function CheckCommitCheckup() As Boolean
'���ܣ���Ժ���˲����ύ����,���ؼ٣�����Ϊ��
Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    '��ҽ��Ժ��д��鱨��ʱ�������鵵��������д
    CheckCommitCheckup = False
    
    If mintPati��Դ = 2 Then
        gstrSQL = "Select count(����ID) ��¼ From ������ҳ Where ����id=[1] And ��ҳid =[2] And ��Ժ���� Is Not Null And Nvl(����״̬, 0) = 5"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����״̬", mlngPati����ID, mlngPati��ҳID)
        If rsTemp!��¼ >= 1 Then Exit Function '
    End If
    CheckCommitCheckup = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub picRichEdit_Resize()
    With edtThis
        .Top = 0: .Left = 0
        .Width = picRichEdit.ScaleWidth: .Height = picRichEdit.ScaleHeight
    End With
End Sub


