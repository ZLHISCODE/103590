VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRView 
   Caption         =   "����������"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "frmEPRView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   9705
   StartUpPosition =   3  '����ȱʡ
   Begin zlRichEditor.Editor edtOrig 
      Height          =   2310
      Left            =   225
      TabIndex        =   0
      Top             =   1305
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4075
      Title           =   ""
      ShowRuler       =   0   'False
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7425
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRView.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9763
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2716
            MinWidth        =   2716
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1658
            MinWidth        =   1658
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin zlRichEditor.Editor edtClear 
      Height          =   2625
      Left            =   225
      TabIndex        =   2
      Top             =   3735
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4630
      Title           =   ""
      ShowRuler       =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   4770
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�ļ� "File"
Private Const ID_File_SaveCopy = 302    '���渱��(A)...
Private Const ID_File_SaveTxt = 303     '����Ϊ�ı�(V)...
Private Const ID_FILE_PRINT = 304       '��ӡ(P)...
Private Const ID_FILE_Copy = 305        '���Ƶ�������(C)
Private Const ID_FILE_EXIT = 306        '�˳�(X)

'��ͼ "View"
Private Const ID_View_Mode = 311        '��ʾ״̬(&S)
Private Const ID_View_Mode_Orig = 312   'ԭʼ״̬(&O)
Private Const ID_View_Mode_Clear = 313  '���״̬(&C)
Private Const ID_View_StatusBar = 314   '״̬��(S)

'���� "Help"
Private Const ID_HELP_CONTENT = 500     '��������
Private Const ID_HELP_CONTACT = 502     '���ͷ���
Private Const ID_HELP_ONLINE = 503      '����ҽҵ
Private Const ID_HELP_ABOUT = 504       '����...

Private mlng��¼ID As Long              '��¼ID
Private mlngPatiId As Long, mlngPageId As Long '��ҳID
Private mlngFileType  As Integer           '��������
Private mlngMode As Long                '��ʾģʽ:0~Orig; 1~Clear
Private blnPrivacyProtect As Boolean    '�Ƿ�������˽����

Public Tables As cEPRTables             '��񼯺�
Public Pictures As cEPRPictures         'ͼƬ����
Public Compends As cEPRCompends         '��ټ���
Public Elements As cEPRElements         '����Ҫ�ؼ���
Public Signs As cEPRSigns               'ǩ���鼯��

Private mblnChildMode As Boolean        '�Ƿ���Ƕ��༭���Ӵ���
Private mblnCanPrint As Boolean         '�Ƿ���Դ�ӡ
Private mlngAdviceID As Long            'ҽ��ID
Private mfrmParent As Object            '���ô���

Public Property Get ChildMode() As Boolean
    ChildMode = mblnChildMode
End Property

Public Property Let ChildMode(vData As Boolean)
    mblnChildMode = vData
    If mblnChildMode Then
        Me.BorderStyle = 0
        SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Xor WS_BORDER Xor WS_THICKFRAME Xor WS_DLGFRAME
    Else
        Me.BorderStyle = 2
    End If
End Property

Public Property Get CanPrint() As Boolean
    CanPrint = mblnCanPrint
End Property

Public Property Let CanPrint(vData As Boolean)
    mblnCanPrint = vData
End Property

'################################################################################################################
'## ���ܣ�  ��ʾ�����ļ����Ĵ���
'##
'## ������  frmParent       ��������
'##         lng��¼ID       ����¼ID
'##         blnPrivacyOn    ���Ƿ�������˽����
'##         blnCanPrint     ���Ƿ������ӡ
'##         blnChildMode    ���Ƿ���Ƕ�뷽ʽ
'################################################################################################################
Public Sub ShowMe(ByRef frmParent As Object, ByVal lng��¼ID As Long, _
    Optional blnPrivacyOn As Boolean = False, _
    Optional blnCanPrint As Boolean = True, _
    Optional blnChildMode As Boolean = False, _
    Optional lngAdviceID As Long)
    
    Dim objControl As CommandBarControl
    Set mfrmParent = frmParent
    mlngFileType = 0
    blnPrivacyProtect = blnPrivacyOn
    mblnCanPrint = blnCanPrint
    mlngAdviceID = lngAdviceID
    Me.ChildMode = blnChildMode
    
    Call InitCommandBars    '��������ʼ��
    
    zlCommFun.ShowFlash "���Ժ�..."
    Screen.MousePointer = vbHourglass
    mlng��¼ID = lng��¼ID      '��¼ID
    mlngMode = 1                '���ģʽ
    
    Call OpenSignleEPR
    Call ShowEPRFile
    
    zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    Me.Show IIf(frmParent.BorderStyle = vbSizable Or frmParent.BorderStyle = vbBSNone, vbModeless, vbModal), frmParent
    Exit Sub
LL:
    Unload Me
    MsgBox "�޷��򿪸��ļ�", vbOKOnly + vbInformation, gstrSysName
End Sub

Private Sub ShowEPRFile()
    edtOrig.Visible = (mlngMode = 0)
    edtClear.Visible = (mlngMode = 1)
    Call cbsThis_Resize
End Sub

Public Sub OpenSignleEPR()
    zlCommFun.ShowFlash "���Ժ�..."
    Screen.MousePointer = vbHourglass
'    DoEvents
    LockWindowUpdate Me.hWnd
    '=================================================================================================
    Dim i As Long, strPath As String, strF As String
    Dim rs As New ADODB.Recordset
    Dim Doc As New cEPRDocument, Elements As New cEPRElements
    Dim lngStart As Long, lngLen  As Long
    Dim lng����ID As Long, lng��ҳID As Long, �������� As EPRDocTypeEnum
    Dim lngKey As Long, blnPrivacy As Boolean
    If blnPrivacyProtect = True Then
        blnPrivacy = InStr(gstrPrivsEpr, ";������˽����;") = 0     '������˽��Ŀ
    End If
    
    gstrSQL = "select ����ID,��ҳID,�������� from ���Ӳ�����¼ where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��¼ID)
    If Not rs.EOF Then
        mlngPatiId = NVL(rs("����ID"), 0)
        mlngPageId = NVL(rs("��ҳID"), 0)
        mlngFileType = NVL(rs("��������"), 1)
    End If
    rs.Close
    
    edtOrig.ForceEdit = True
    edtClear.ForceEdit = True
    '������ʱ�ļ�
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    strF = strPath & "\" & App.hInstance & CLng(Timer) & ".TMP"
    Doc.InitEPRDoc cprEM_�޸�, cprET_���������, mlng��¼ID, IIf(mlngFileType = 2, 2, 1), lng����ID, CStr(lng��ҳID), 0, glngDeptId, mlngAdviceID
    Doc.OpenEPRDoc Doc.frmEditor.Editor1         '�򿪸��ļ�
    '�����滻��Ŀ
    If blnPrivacy Then
        '��ȡ���е�Ҫ��
        gstrSQL = "Select A.ID,A.������ From ���Ӳ������� A, ��˽������Ŀ B,����������Ŀ C " & _
            "Where A.�������� = 4 And A.�滻�� = 1 And A.�ļ�id = [1] And A.������� > 0 and B.��Ŀid = C.ID And A.Ҫ������ =C.������ And C.�滻�� = 1 "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��¼ID)
        If Not rs.EOF Then
            Do While Not rs.EOF
                lngKey = Elements.Add(NVL(rs("������"), 0))
                Elements("K" & lngKey).GetElementFromDB cprET_�������༭, rs("ID"), True, "���Ӳ�������"
                '�滻Ҫ������
                Elements("K" & lngKey).�����ı� = String(Len(Elements("K" & lngKey).�����ı�), "*")
                Elements("K" & lngKey).Refresh Doc.frmEditor.Editor1
                rs.MoveNext
            Loop
        End If
        rs.Close
    End If
    Doc.frmEditor.SaveDocToFile strF, False     '�洢�������ʱ�ļ�
    
    With edtOrig
        .NewDoc
        .ForceEdit = True
        .ViewMode = cprNormal
        .OpenDoc strF
        
        '����ҳüҳ��
        Set .Picture = Doc.frmEditor.Editor1.Picture
        .HeadFileTextRTF = Doc.frmEditor.Editor1.HeadFileTextRTF
        .FootFileTextRTF = Doc.frmEditor.Editor1.FootFileTextRTF
        
        Call Doc.GetReplacedHeadFootString(edtOrig)
        '����ҳ���ʽ
        Doc.EPRFileInfo.SetFormat edtOrig, Doc.EPRFileInfo.��ʽ
        edtOrig.ResetWYSIWYG    'ˢ�����������ã�WYSIWYG����ʾ
        
        '��ҳ
        .ViewMode = cprNormal
        .AuditMode = True
        .Range(0, 0).Selected
        .ForceEdit = False
        .ReadOnly = True
    End With

    With edtClear
        .NewDoc
        .ForceEdit = True
        .ViewMode = cprNormal
        .OpenDoc strF
        
        '����ҳüҳ��
        Set .Picture = Doc.frmEditor.Editor1.Picture
        .HeadFileTextRTF = Doc.frmEditor.Editor1.HeadFileTextRTF
        .FootFileTextRTF = Doc.frmEditor.Editor1.FootFileTextRTF
        
        Call Doc.GetReplacedHeadFootString(edtClear)
        '����ҳ���ʽ
        Doc.EPRFileInfo.SetFormat edtClear, Doc.EPRFileInfo.��ʽ
        edtClear.ResetWYSIWYG    'ˢ�����������ã�WYSIWYG����ʾ
        
        '��ҳ
        .SelectAll
        .AuditMode = True
        .AcceptAuditText
        .ViewMode = cprNormal
        .Range(0, 0).Selected
        .ForceEdit = False
        .ReadOnly = True
    End With
    If gobjFSO.FileExists(strF) Then gobjFSO.DeleteFile strF    'ɾ����ʱ�ļ�
 
    Doc.frmEditor.Editor1.Modified = False
    
    Set rs = Nothing
    
    '=================================================================================================
    LockWindowUpdate 0
    zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
End Sub

'################################################################################################################
'## ���ܣ�  ���ΪRTF�ļ�
'################################################################################################################
Private Function SaveAsRTFFile() As Boolean
    On Error GoTo LL
    Dim strF As String
    dlgThis.Filename = ""
    dlgThis.Filter = "*.rtf|*.rtf|*.*|*.*"
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If strF <> "" Then
        If gobjFSO.FileExists(strF) Then
            If MsgBox("�ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbOK Then
                gobjFSO.DeleteFile strF, True
            Else
                Exit Function
            End If
        End If
        '���浽�ļ�
        If Me.edtOrig.Visible Then
            Me.edtOrig.SaveDoc strF
        Else
            Me.edtClear.SaveDoc strF
        End If
        MsgBox "����ɹ����ļ���:" & vbCrLf & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    SaveAsRTFFile = True
    Exit Function
LL:
    MsgBox "����ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
    SaveAsRTFFile = False
End Function

'################################################################################################################
'## ���ܣ�  ���ΪTXT�ļ�
'################################################################################################################
Private Function SaveAsTxtFile() As Boolean
    On Error GoTo LL
    Dim strF As String
    dlgThis.Filename = ""
    dlgThis.Filter = "*.txt|*.txt|*.*|*.*"
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If strF <> "" Then
        '���浽�ļ�
        If gobjFSO.FileExists(strF) Then
            If MsgBox("�ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbOK Then
                gobjFSO.DeleteFile strF, True
            Else
                Exit Function
            End If
        End If
        Const ForReading = 1, ForWriting = 2, ForAppending = 3
        Dim fs As FileSystemObject, f As TextStream
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.OpenTextFile(strF, ForWriting, TristateUseDefault)
        If Me.edtOrig.Visible Then
            f.Write Me.edtOrig.Text
        Else
            f.Write Me.edtClear.Text
        End If
        
        f.Close
        MsgBox "����ɹ����ļ���:" & vbCrLf & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    SaveAsTxtFile = True
    Exit Function
LL:
    MsgBox "����ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
    SaveAsTxtFile = False
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_File_SaveCopy
        '���渱��(A)...
        Call SaveAsRTFFile
    Case ID_File_SaveTxt
        '����Ϊ�ı�(V)...
        Call SaveAsTxtFile
    Case ID_FILE_Copy
        If Control.Enabled And Control.Visible Then '��ݼ�ִ��ʱ��Ҫ�ж�
            gstrCopyPID = CStr(mlngPatiId)
            If Me.edtOrig.Visible Then
                edtOrig.Copy
            Else
                edtClear.Copy
            End If
        End If
    Case ID_FILE_PRINT
        '��ӡ(P)...
        If Me.edtOrig.Visible Then
            If edtOrig.PrintDoc(False, 0, 0, "", 1) = False Then Exit Sub
        Else
            If edtClear.PrintDoc(False, 0, 0, "", 1) = False Then Exit Sub
        End If
        If mfrmParent Is Nothing Then Exit Sub
        If InStr(mfrmParent.Caption, "���Ʊ������") > 0 Or InStr(mfrmParent.Caption, "���Ʊ������") > 0 And mlngFileType = cpr���Ʊ��� Then '���򿪶������ʱ�������¼���ֻ���Ǹ�����ķ���
            Call mfrmParent.Event_AfterPrinted(mlng��¼ID)
        End If
        Call PrintTag(mlng��¼ID, mlngFileType, mlngPatiId, mlngPageId)
        On Error Resume Next
        mfrmParent.RefreshList: Err.Clear
        Unload Me
    Case ID_FILE_EXIT
        '�˳�(X)
        Unload Me
    Case ID_View_Mode_Orig
        'ԭʼ״̬
        mlngMode = 0
        Call ShowEPRFile
    Case ID_View_Mode_Clear
        '����״̬
        mlngMode = 1
        Call ShowEPRFile
    Case ID_View_StatusBar
        '״̬��(S)
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case ID_HELP_CONTENT
        '��������
        ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
    Case ID_HELP_CONTACT
        '���ͷ���
        Call zlMailTo(Me.hWnd)
    Case ID_HELP_ONLINE
        '����ҽҵ
        Call zlHomePage(Me.hWnd)
    Case ID_HELP_ABOUT
        '����...
        ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
    End Select
    
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height / Screen.TwipsPerPixelY
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.cbsThis.GetClientRect Left, Top, Right, Bottom
    edtOrig.Width = 0: edtOrig.Height = 0
    edtOrig.Move Left * Screen.TwipsPerPixelX, Top * Screen.TwipsPerPixelY, _
        (Right - Left) * Screen.TwipsPerPixelX, (Bottom - Top) * Screen.TwipsPerPixelY
    edtClear.Width = 0: edtClear.Height = 0
    edtClear.Move Left * Screen.TwipsPerPixelX, Top * Screen.TwipsPerPixelY, _
        (Right - Left) * Screen.TwipsPerPixelX, (Bottom - Top) * Screen.TwipsPerPixelY
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_File_SaveCopy
        '���渱��(A)
        Control.Enabled = mblnCanPrint
    Case ID_FILE_Copy
        If Me.edtOrig.Visible Then
            Control.Enabled = (Trim(Me.edtOrig.SelText) <> "" And Me.edtOrig.ViewMode <> cprPaper)
        Else
            Control.Enabled = (Trim(Me.edtClear.SelText) <> "" And Me.edtClear.ViewMode <> cprPaper)
        End If
        Control.Visible = InStr(gstrPrivsEpr, "���ݸ���") > 0
    Case ID_File_SaveTxt
        '����Ϊ�ı�(V)...
        Control.Enabled = mblnCanPrint
    Case ID_FILE_PRINT
        '��ӡ(P)...
        Control.Enabled = mblnCanPrint
    Case ID_FILE_EXIT
        '�˳�(X)
    Case ID_View_StatusBar
        '״̬��(S)
        Control.Checked = stbThis.Visible
    Case ID_View_Mode_Orig
        'ԭʼ״̬
        Control.Checked = (mlngMode = 0)
    Case ID_View_Mode_Clear
        '����״̬
        Control.Checked = (mlngMode = 1)
    Case ID_HELP_CONTENT
        '��������
    Case ID_HELP_CONTACT
        '���ͷ���
    Case ID_HELP_ONLINE
        '����ҽҵ
    Case ID_HELP_ABOUT
        '����...
    End Select
End Sub

Private Sub InitCommandBars()
Dim BarMain As CommandBar
Dim cbp�ļ� As CommandBarPopup      '�ļ��˵�
Dim cbp��ͼ As CommandBarPopup      '��ͼ�˵�
Dim cbp���� As CommandBarPopup      '�����˵�
    '����λ�ûָ�
    Call RestoreWinState(Me, App.ProductName)
    '## �˵���ʼ��
    Dim cbpPopup As CommandBarPopup                     '��ʱ����
    Dim cbpPopupSub As CommandBarPopup                  '��ʱ����
    Dim objControl As CommandBarControl                 '�������ؼ�
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True         '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    cbsThis.ActiveMenuBar.Title = "�˵���"
    Set cbp�ļ� = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ļ�(&F)")
    With cbp�ļ�.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_File_SaveCopy, "���渱��(&A)..."): objControl.IconId = 104
        .Add xtpControlButton, ID_File_SaveTxt, "���Ϊ�ı�(&T)..."
        Set objControl = .Add(xtpControlButton, ID_FILE_Copy, "�����ı�(&C)")
        objControl.Visible = False
        
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "��ӡ(&P)..."): objControl.IconId = 103
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�˳�(&X)"): objControl.IconId = 191
        objControl.BeginGroup = True
    End With
    
    Set cbp��ͼ = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "��ͼ(&V)")
    With cbp��ͼ.CommandBar.Controls
        Set cbpPopup = .Add(xtpControlPopup, 0, "������(&T)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, "�������б�"
        .Add xtpControlButton, ID_View_StatusBar, "״̬��(&S)"
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Orig, "ԭʼ״̬(&O)"): objControl.BeginGroup = True: objControl.STYLE = xtpButtonCaption
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Clear, "����״̬(&C)"): objControl.STYLE = xtpButtonCaption
    End With
    
    Set cbp���� = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&H)")
    With cbp����.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "��������(&H)")
        objControl.BeginGroup = True
        Set cbpPopupSub = .Add(xtpControlPopup, 0, "&Web�ϵ�" & gstrProductName)
        objControl.BeginGroup = True
        Set objControl = cbpPopupSub.CommandBar.Controls.Add(xtpControlButton, ID_HELP_ONLINE, gstrProductName & "����(&H)"): objControl.IconId = conMenu_Help_Web_Forum
        Set objControl = cbpPopupSub.CommandBar.Controls.Add(xtpControlButton, ID_HELP_CONTACT, "���ͷ���(&M)"): objControl.IconId = conMenu_Help_Web_Mail
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "����(&A)..."): objControl.IconId = conMenu_Help_About
        objControl.BeginGroup = True
    End With
    
    Set BarMain = cbsThis.Add("������", xtpBarTop)
    With BarMain.Controls
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Orig, "ԭʼ״̬(F5)"): objControl.BeginGroup = True: objControl.STYLE = xtpButtonCaption
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Clear, "����״̬(F6)"): objControl.STYLE = xtpButtonCaption
        
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "����"): objControl.IconId = conMenu_Help_Help
        objControl.BeginGroup = True
        objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�ر�")
        objControl.BeginGroup = True
        objControl.STYLE = xtpButtonIconAndCaption
    End With
    
    '�ȼ���
    cbsThis.KeyBindings.Add FCONTROL, Asc("S"), ID_File_SaveCopy
    cbsThis.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_FILE_Copy
    cbsThis.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT
    
    cbsThis.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT
    cbsThis.KeyBindings.Add 0, VK_F5, ID_View_Mode_Orig
    cbsThis.KeyBindings.Add 0, VK_F6, ID_View_Mode_Clear
End Sub

Private Sub edtClear_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    'û�����ݸ���Ȩ�޲�������
    If InStr(gstrPrivsEpr, "���ݸ���") = 0 Then Exit Sub
    
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_FILE_Copy, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub edtOrig_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    'û�����ݸ���Ȩ�޲�������
    If InStr(gstrPrivsEpr, "���ݸ���") = 0 Then Exit Sub
    
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_FILE_Copy, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub Form_Load()
    Set Tables = New cEPRTables
    Set Pictures = New cEPRPictures
    Set Compends = New cEPRCompends
    Set Elements = New cEPRElements
    Set Signs = New cEPRSigns
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Set Tables = Nothing
    Set Pictures = Nothing
    Set Compends = Nothing
    Set Elements = Nothing
    Set Signs = Nothing
    Set mfrmParent = Nothing
End Sub
Private Sub PrintTag(ByVal lngId As Long, ByVal lngFileType As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long)
On Error GoTo errHand
    gstrSQL = "Zl_���Ӳ�����ӡ_Insert(" & lngId & "," & lngFileType & "," & lngPatiID & "," & lngPageId & ",'" & gstrUserName & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
