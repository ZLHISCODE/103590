VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#5.3#0"; "zl9PacsControl.ocx"
Begin VB.Form frmReport 
   Caption         =   "PACS ����༭"
   ClientHeight    =   8250
   ClientLeft      =   15
   ClientTop       =   300
   ClientWidth     =   10950
   ClipControls    =   0   'False
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   10950
   Begin VB.PictureBox picReportWordContainer 
      BorderStyle     =   0  'None
      Height          =   2912
      Left            =   882
      ScaleHeight     =   2910
      ScaleWidth      =   2910
      TabIndex        =   11
      Top             =   4032
      Visible         =   0   'False
      Width           =   2912
   End
   Begin VB.PictureBox picReportViewContainer 
      BorderStyle     =   0  'None
      Height          =   3164
      Left            =   378
      ScaleHeight     =   3165
      ScaleWidth      =   2790
      TabIndex        =   10
      Top             =   2772
      Width           =   2786
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2016
      Top             =   126
   End
   Begin VB.Timer tmrCheckingReportState 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2772
      Tag             =   "0"
      Top             =   126
   End
   Begin VB.PictureBox picReportHistoryList 
      Height          =   5895
      Left            =   5400
      ScaleHeight     =   5835
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      Begin zl9PacsControl.ucSplitter ucSplitterH 
         Height          =   135
         Left            =   120
         TabIndex        =   9
         Top             =   2490
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   238
         MousePointer    =   7
         SplitType       =   0
         Con1MinSize     =   200
         Con2MinSize     =   650
         Control1Name    =   "lvHistoryList"
         Control2Name    =   "picReportDetail"
      End
      Begin VB.PictureBox picReportDetail 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   3510
         TabIndex        =   5
         Top             =   2625
         Width           =   3540
         Begin VB.CommandButton cmdSelectWord 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            Picture         =   "frmReport.frx":0CCA
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "����ǰѡ�е��ı�д�뱨��"
            Top             =   0
            Width           =   1200
         End
         Begin VB.CommandButton cmdViewImage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            Picture         =   "frmReport.frx":1EC4
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "�鿴���߱��μ���Ӱ��"
            Top             =   0
            Width           =   1200
         End
         Begin RichTextLib.RichTextBox rtxtReport 
            Height          =   3495
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   3135
            _ExtentX        =   5556
            _ExtentY        =   6191
            _Version        =   393217
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmReport.frx":3166
         End
      End
      Begin VB.CheckBox chkOtherDeptReport 
         Caption         =   "�鿴��������ʷ����"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvHistoryList 
         Height          =   2010
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3545
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "dfd"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "dsd"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin zlRichEditor.Editor edtEditor 
      Height          =   1092
      Left            =   1386
      TabIndex        =   1
      Top             =   756
      Visible         =   0   'False
      Width           =   1694
      _ExtentX        =   3016
      _ExtentY        =   1931
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   3402
      Top             =   126
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "����"
   End
   Begin RichTextLib.RichTextBox rtxtSaveElement 
      Height          =   742
      Left            =   0
      TabIndex        =   0
      Top             =   756
      Visible         =   0   'False
      Width           =   1092
      _ExtentX        =   1931
      _ExtentY        =   1296
      _Version        =   393217
      TextRTF         =   $"frmReport.frx":3203
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   1078
      Top             =   126
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmReport.frx":3292
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenu

Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ�������ѡ��"
Private Const M_STR_MODULE_MENU_TAG As String = "����"
Private Const M_STR_LISTVIEWKEY_DESCRIBE As String = "describe"
Private Const M_STR_LISTVIEWKET_PROCESS As String = "process"

Private mlngModule As Long
Private mstrPrivs As String         'Ȩ���ַ���
Private mlngDeptID As Long          '��ǰ����ID
Private mobjOwner As Object

Private mlngTmpAdviceId As Long
Private mlngTmpSendNo As Long

Private mlngAdviceID As Long        'ҽ��ID
Private mlngSendNo As Long          '���ͺ�
Private mblnMoved As Boolean        '�Ƿ�ת��
Private mlngStudyState As Long

Private WithEvents mfrmReportView As frmReportView
Attribute mfrmReportView.VB_VarHelpID = -1
Private WithEvents mfrmReportImage As frmReportImage
Attribute mfrmReportImage.VB_VarHelpID = -1
Private mfrmReportSpecial As Object
Private WithEvents pobjPacsCore As zl9PacsCore.clsViewer     '��Ƭվ����
Attribute pobjPacsCore.VB_VarHelpID = -1
Private WithEvents mfrmReportWord As frmReportWord          '�ʾ�ʾ������
Attribute mfrmReportWord.VB_VarHelpID = -1
Private WithEvents mobjCustomReport As clsReport                  '�Զ��屨�����
Attribute mobjCustomReport.VB_VarHelpID = -1
Private WithEvents mobjReport As zlRichEPR.cDockReport      '�������
Attribute mobjReport.VB_VarHelpID = -1
Private mobjWork_ImageCap As Object ' zl9PacsCapture.clsPacsCapture  '��Ƶ�ɼ�ģ��
Attribute mobjWork_ImageCap.VB_VarHelpID = -1

Private mblnSingleWindow As Boolean     '�Ƿ�ʹ�ö���������ʾ����༭����True-����������ʾ��False-Ƕ��ʽ��ʾ
Private mlngEPRDeptID As Long   '��ǰ�����С����Ӳ�����¼������¼�Ŀ���ID
Private mstrEPR������ As String '��ǰ�����еġ����Ӳ�����¼������¼�Ĵ�����
Private mstrEPR������ As String '��ǰ�����еġ����Ӳ�����¼������¼�ı�����
Private mlngEPRǩ������ As Long '��ǰ�����еġ����Ӳ�����¼������¼��ǩ������
Private mdtReportTime As Date   '���汣��ʱ��
Private mlngPassType As Long                 '������֤����ϵͳ������ 0-���룻1�����֣�2�����߽Կ�

Private mFileID As Long         '�����ļ�ID,������ʽ�ļ�
Private mReportID As Long       '���������ļ�ID
Private mFormatID As Long     '��������ID
Private mModelName As String     '��������
Private mintEditType As Integer '����״̬ 0 ������1��д��2 �޶�
Private mintReportViewType As Integer ' 0-�������CheckView��1-������Result��2-����Advice
Private miES As Integer
Private miEE As Integer

Private mstrCurReportViewType As String

Private mHasChangeFormat As Boolean     '��¼�Ƿ�����˸�ʽ


Private mblnModified As Boolean              '���������Ƿ�ı�
Private mblnReadOnly As Boolean         '�Ƿ�ֻ��״̬�����������޸ı���

Public mblnEditable As Boolean         '�Ƿ���Ա༭����

Private mstrModifyEdit As String        '��ǰ�����Ƿ����޶�״̬���������޶������û��ǩ������¼�����˵��������ձ�ʾ�����������
Private mblnCanUntread As Boolean       '�Ƿ�������ˡ��������Ѿ�����ӡ�����ұ���˺󣬲��������

Private mSigns As New cEPRSigns         '��ǰ�ĵ��е�ǩ��
Private m���汾 As Integer         '���汾
Private mĿ��汾 As Integer        'Ŀ��汾
Private mǩ������ As EPRSignLevelEnum        '1-��д;2-����ҽʦ����;3-����ҽʦ���ġ�סԺ��������Ĳ���ֻ����д������״̬
Private mModified As Boolean
Public mblnShowImage As Boolean            '�Ƿ���ʾ����ͼ��
Private mblnShowSpecial As Boolean         '�Ƿ���ʾר�Ʊ���
Public mblnShowVideoCapture As Boolean     '�Ƿ���ʾͼ��ɼ�


Private mstrPatholMaterialInfo As String    '����ȡ����ʾ����
Private mstrSpecialForm As String           'ר�Ʊ��洰������
Private mlngShowBigImg As Long              '��������ʾ��ͼ
Private mintMinImageCount As Integer        '��������ͼ��ʾ����
Private mblnExitAfterPrint As Boolean       '�����ӡ��رմ���
Private mintImageDblClick As Integer        '����ͼ˫����Ĳ�����0--ֱ��д�뱨�棻1--��ͼ��༭����

Private mblnIgnoreResult As Boolean         '���Խ��������
'Private mintCriticalValues As Integer                       'Σ��ֵ
Private mintConformDetermine As Integer                     '�������
Private mstrImageLevel As String                            'Ӱ�������ȼ���
Private mstrReportLevel As String                           '���������ȼ���
Private mintImageLevel As Integer                           'Ӱ�������ж�
Private mintReportLevel As Integer                          '���������ж�

Private mlngHintType As Long

Private mblnReportWithResult As Boolean      '��Ӱ�����Ϊ����

Private mblnShowWord As Boolean             '�Ƿ�̬��ʾ�ʾ�ʾ����True-һֱ��ʾ�ʾ�ʾ����False--˫���������ʾ�ʾ�ʾ������
Private mintWordDblClick As Integer         '�ʾ�˫����Ĳ�����0--ֱ��д�뱨�棻1--�򿪴ʾ�༭����
Private mblnRptImg2CapImg As Boolean
Private mstrFormatInfo As String
Private mblnCheckPrintPara As Boolean         'ƽ����Ҫ��˲��ܴ�ӡ =true����������
Private mblnCanPrint As Boolean             '�ò��˵ı��棬�Ƿ������ӡ
Private mblnCheckOtherDeptReport As Boolean     '�Ƿ�ͨ����ʷ���湦�ܲ鿴�����Ƶ���ʷ����
Private mblnUntreadPrinted As Boolean           '��˴�ӡ���Ƿ�������ˣ�True--���Ի��ˣ�False--�����Ի��ˡ�
Private mblnPrintView As Boolean            '������δ�ҵ���Ӧ�����ļ�������� ����ӡ����Ԥ������ť�Ľ���״̬��true Ϊ����  false Ϊ������
Private mblnIsReportDelete As Boolean      '�Ƿ���ɾ�����浥��
Private mblnTechReptSame As Boolean        'ֻ����д�Լ����ı���
Private mlngPrintFormat As Long            '�����ӡ��ʽ
Private mblnIsPetitionScan As Boolean      '�Ƿ��������뵥ɨ��
Private mblnSetFocusWithReport As Boolean '����л�ʱ��λ����༭
Private mblnAllowLocate As Boolean
Private mblnIsPrint As Boolean             '�����ֱ�Ӵ�ӡ

Private mobjFSO As New Scripting.FileSystemObject    'FSO����
Private mclsUnzip As New cUnzip
Private mclsZip As New cZip

Private mlngCY21 As Long                 '�ı�����ĸ߶�
Private mlngCY22 As Long                 'ר�Ʊ���ĸ߶�
Private mlngCX1 As Long                  'ģ��Ŀ��
Private mlngCX2 As Long                  '�ı�����Ŀ��
Private mlngCX3 As Long                  'ͼ������Ŀ��
Private mlngCY3 As Long                  'ͼ������ĸ߶�
Private mlngCX4 As Long                  '��Ƶ�ɼ�����Ŀ��
Private mlngCY4 As Long                  '��Ƶ�ɼ�����ĸ߶�
Private mlngPicHistoryY As Long          '������ʷ����ĸ߶�
Private mlngPicHistoryX As Long          '������ʷ����Ŀ��
Private mlngPrivateWordY As Long         '˽�˳��ôʾ�����ĸ߶�
Private mblnExitAfterSign As Boolean     'ǩ�����˳�
Private mintPaneID As Integer             '��ǰѡ�е�Pane ID

Private mblnPrintOK As Boolean           '��ӡ���

Private mblnMenuDownState As Boolean    '����˫����������������
Private mblnIsSignSave As Boolean

Private mblnCompareSize As Boolean

'������Ϣ
Private mstrҽ������ As String             'ҽ������
Private mstrҽ������ As String        'ҽ������

Private Type rptFormat
    ID As Long          '�����ʽID
    strName As String   '�����ʽ����
End Type
Private rptFormats() As rptFormat

'�����Զ��屨��Ĵ�ӡ��ʽ
Private mblnʹ���Զ��屨�� As Boolean           '��ӡ��ʽ�Ƿ����Զ��屨��
Private mstr������ As String                  '�Զ��屨��ı��
Private mblnRefreshRptFormat As Boolean         '��ӡ��ʽ��Ҫˢ��
Private mstrѡ�б����ʽ As String              '��ѡ�е��Զ��屨���ʽ
Private mblOneReportFormat As Boolean           '�Ƿ�ֻ��ѡ��һ�ִ�ӡ��ʽ

'�ʾ�ʾ������Ӻ��޸�
Private mintWordPower As Integer        '�ʾ����Ȩ��Χ
'    mintWordPower=-1�����߱��ʾ����Ȩ;
'    mintWordPower=0��ȫԺ����ʱ��ʾ���е�ʾ����Ҳ���Ը���;
'    mintWordPower=1�����ң���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ��ҹ��л�������Ա˽�е�ʾ���������ܸ���ȫԺͨ��ʾ��;
'    mintWordPower=2�����ˣ���ʱ��ʾȫԺͨ��ʾ��(����id is null)�����ڿ���ͨ��ʾ��(��Աid is null)�͸���ʾ����������ʾ���ɸ���

Private Const Report_Element_����ǩ�� = "����ǩ��"

Private mObjActiveMenuBar As CommandBars
Private mblnRefreshState As Boolean

Public mblnClosed As Boolean        '�жϸñ���༭���Ƿ��Ѿ����ر�

'��������¼�
Public Event AfterOpen()
Public Event BeforeEdit(ByVal lngOrderID As Long)

'frmOwnerForm��Ҫ�����ڸ��¼�ִ��ʱ,����ڸ��¼�����ģ̬������ʾ������Ҫ�ƶ�ownerӵ���ߣ��ұ���ʹ�øò����������
'��ɳ����ڼ���״̬�����ڵ��Ӳ���ǩ����ѡ�������ʣ������ʴ�����δ���øò���������£���������������
Public Event OnImageCountChanged(ByVal intType As Integer, ByVal isNeedRefreshTitle As Boolean)
Public Event AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Form, ByVal lngSaveType As Long, ByVal isRefreshFace As Boolean)
Public Event AfterClosed(ByVal lngOrderID As Long)
Public Event AfterPrinted(ByVal lngOrderID As Long)
Public Event AfterDeleted(ByVal lngOrderID As Long)
Public Event AfterReleationImage(ByVal lngOrderID As Long, ByVal lngSendNO As Long, ByVal intStep As Integer, ByVal lngReleationType As Long)

'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenu
    Set zlMenu = Me
End Property


'��ȡ��ǰ�����ҽ��ID
Property Get AdviceId()
    AdviceId = mlngAdviceID
End Property


'���ñ����ͼ�������
Property Get PacsCore() As zl9PacsCore.clsViewer
    Set PacsCore = pobjPacsCore
End Property

Property Set PacsCore(objPacsCore As zl9PacsCore.clsViewer)
    Set pobjPacsCore = objPacsCore
End Property

Property Get ReportViewForm() As frmReportView
    Set ReportViewForm = mfrmReportView
End Property

Property Get CurReportViewType() As String
    CurReportViewType = mstrCurReportViewType
End Property

Public Sub NotificationRefresh()
'֪ͨˢ��
    mblnRefreshState = False
End Sub


Public Sub VideoCallBack(EventType As Long, lngAdviceID As Long, _
    Optional strStudyUID As String, Optional strPatientName As String, Optional blnIsLock As Boolean)
    '���ڱ���ͨ��......
End Sub

Public Sub UpdateImageVideoState(ByVal lngEventType As TVideoEventType, ByVal lngAdviceID As Long, _
    ByVal other As Variant)
    
    Dim i As Integer
    Dim strInstanceUID As String
    
    '�������༭������ʾ�˱���ͼ���¼�����Ϊ����ͼ����ִ�и��±���ͼ�����
    If mblnShowImage Then
        If lngEventType = TVideoEventType.vetAfterUpdateImg Or lngEventType = vetExportImage Or lngEventType = vetImportImage Or _
        ((lngEventType = TVideoEventType.vetUpdateImg Or lngEventType = TVideoEventType.vetCaptureFirstImg Or lngEventType = TVideoEventType.vetDelAllImg Or lngEventType = TVideoEventType.vetImgDeled) And lngAdviceID = mlngAdviceID) Then
            Call RefPacsPic(lngEventType)
            Exit Sub
        ElseIf lngEventType = TVideoEventType.vetAddReportImg And mfrmReportImage Is Nothing = False Then
            strInstanceUID = other
            Call mfrmReportImage.ReportImageAdd(strInstanceUID)
        End If
    End If
    
    If Not mblnShowVideoCapture Then Exit Sub
    
    Select Case lngEventType
        Case TVideoEventType.vetLockStudy
            For i = 1 To dkpMain.PanesCount
                If dkpMain.Panes(i).Title Like "*��Ƶ�ɼ�*" Then
                    dkpMain.Panes(i).Title = "��" & other & "����Ƶ�ɼ�"
                    Exit For
                End If
            Next i
        Case TVideoEventType.vetUnLockStudy
            For i = 1 To dkpMain.PanesCount
                If dkpMain.Panes(i).Title Like "*��Ƶ�ɼ�*" Then
                    dkpMain.Panes(i).Title = "��Ƶ�ɼ�"
                    Exit For
                End If
            Next i
'        Case TVideoEventType.vetUpdateImg
'            '����ͼ��
'            If lngAdviceID = mlngAdviceID Then
'                Call RefPacsPic
'            End If
    End Select
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnMenuDownState Then
         If MsgBoxD(Me, "��ǰ������δ��ɣ�ǿ���˳�������ɳ����쳣���Ƿ������", vbYesNo, "����") = vbNo Then Cancel = True
    End If
End Sub

'�ӿ�ʵ�ֲ���*********************************************************************************

Public Function IWorkMenu_zlIsModuleMenu(ByVal objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'�жϲ˵��Ƿ����ڸ�ģ��˵�
    IWorkMenu_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenu_zlCreateMenu(objMenuBar As Object)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    Set mObjActiveMenuBar = objMenuBar
    
'    If Not HasMenu(objMenuBar, conMenu_EditPopup) Then
        '�༭�˵�:���ڹ���˵�(���������û��)���ļ��˵�����
        '-----------------------------------------------------

        Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����", 3, False)
        cbrMenuBar.ID = conMenu_EditPopup
        cbrMenuBar.Category = ""
        
        With cbrMenuBar.CommandBar
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_NoAskPrint, "ʹ�þ�Ĭ��ӡ", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Preview, "Ԥ��", "", 102, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Print, "��ӡ", "", 103, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_BatPrint, "������ӡ", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_Open, "��д", "", 3002, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_ClearWritingState, "���״̬", "", 21903, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Edit_Delete, "ɾ��", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Open, "����", "", 0, True)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_ExportToXML, "����XML��", "", 0, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Search, "���������", "", 0, False)
        End With
'    End If
End Sub


Public Sub IWorkMenu_zlCreateToolBar(objToolBar As Object)
'����������
    Dim cbrControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long
    
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue, , True)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index

    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Preview, "Ԥ��", "����Ԥ��", 102, True, lngIndex + 1)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Print, "��ӡ", "�����ӡ", 103, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_PacsReport_Open, "��д", "", 2607, False, lngIndex + 3) 'IconId=3002
End Sub


Public Sub IWorkMenu_zlClearMenu()
'����������Ĳ˵�
    Exit Sub
End Sub


Public Sub IWorkMenu_zlClearToolBar()
'��������Ĺ�����
    Exit Sub
End Sub



Public Sub IWorkMenu_zlUpdateMenu(ByVal control As XtremeCommandBars.ICommandBarControl)
    Call cbrMain_Update(control)
End Sub

Public Sub IWorkMenu_zlExecuteMenu(ByVal lngMenuId As Long)
    Dim objControl As XtremeCommandBars.ICommandBarControl
    
    Set objControl = mObjActiveMenuBar.FindControl(, lngMenuId, , True)
    If objControl Is Nothing Then Exit Sub
    
    Call cbrMain_Execute(objControl)
End Sub


Public Sub IWorkMenu_zlPopupMenu(objPopup As XtremeCommandBars.ICommandBar)
'�����Ҽ��˵�
    Exit Sub
End Sub

Public Sub IWorkMenu_zlRefreshSubMenu(objMenuBar As Object)
'ˢ�µ������Ӳ˵�
    Exit Sub
End Sub

'*************************************************************************************************


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function


Public Sub zlInitModule(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDepartId As Long, Optional owner As Object = Nothing, Optional blnSingleWindow As Boolean = False)
'��ʼ������ģ��
    Dim blnRestoreWindow As Boolean
    
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngDeptID = lngDepartId
    
    If Not owner Is Nothing Then Set mobjOwner = owner

    mblnSetFocusWithReport = Val(GetDeptPara(lngDepartId, "����л�ʱ��λ����༭", "1")) = 1
    mblnSingleWindow = blnSingleWindow
    
    blnRestoreWindow = IIf(mlngDeptID = 0, True, False)
    
    '��ʼ�Ӵ���
    If mfrmReportView Is Nothing Then Set mfrmReportView = New frmReportView      '��������
    
    If mfrmReportWord Is Nothing Then Set mfrmReportWord = New frmReportWord      '�ʾ�ʾ��
    If mobjReport Is Nothing Then Set mobjReport = New zlRichEPR.cDockReport      '���Ӳ�������
    
    Call InitLoaclParas(mlngDeptID, mlngModule, mstrPrivs, mlngModule = G_LNG_PACSSTATION_MODULE)

    Call InitFaceScheme  '��ʼ���沼��,���������
    
    Call subShowHistoryList
    
    'ʹRestoreWinState�����ڸö������������ִֻ��һ�Σ��������Ƕ�׵ı���༭��λ�ô�λ��
    If blnRestoreWindow Then Call RestoreWinState(Me, App.ProductName)
    
    '��ȡ�û��Ĵʾ�ʾ��Ȩ�ޣ��������޹�
    mintWordPower = zlGetWordPower
    
    '���������Ƶ�ɼ����ڣ������������Ƶ�ɼ����ڳ�ʼ��
    If mblnShowVideoCapture Then
        Call InitActiveVideoModuleObj
    End If
    
    mstrCurReportViewType = ""
End Sub


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngStudyState As Long, ByVal blnMoved As Boolean)
'ͬ�����ҽ����Ϣ
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mblnMoved = blnMoved
    mlngStudyState = lngStudyState
    mblnRefreshState = True
    
    If Not mobjWork_ImageCap Is Nothing Then Call mobjWork_ImageCap.zlUpdateStudyInf(lngAdviceID, lngSendNO, lngStudyState, blnMoved, mReportID <> 0)
End Sub


Public Function zlRefreshFace(Optional blnForceRefresh As Boolean = False, Optional blnIsDockActive As Boolean = False) As Boolean
On Error GoTo errHandle
    Dim lngNewAdviceId As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngOldFileID As Long            '��¼��ǰ���ļ�ID�������Ƚ����Ƶ����Ƿ����ı�
    Dim blnPrinted As Boolean           '�����Ƿ��Ѿ�����ӡ
    Dim lngStudyState As Long           '����״̬��������ҽ������.ִ�й��̡�
    Dim str����� As String             '��������Ѿ���ˣ������������ǩ����
    Dim str�������  As String          '��¼�������������
    Dim thisUserSignLevel As EPRSignLevelEnum   '��ǰ�û���ǩ������
    Dim arrReportFormat() As String
    Dim strRegPath As String
    Dim str��鱨��ID As String         '�����ĵ��༭����Ӧ����ID
    
    If (mlngTmpAdviceId = mlngAdviceID And mlngTmpSendNo = mlngSendNo And mblnRefreshState) And Not blnForceRefresh Then
        If blnIsDockActive = False Then tmrFocus.Enabled = True
        Exit Function
    End If
    
    lngNewAdviceId = mlngAdviceID
    mlngAdviceID = mlngTmpAdviceId
    
    '�ж���һ�εĸĶ��Ƿ񱣴�
    Call PromptModify
    
    '�ָ�ҽ��ID
    mlngAdviceID = lngNewAdviceId
    
    mlngTmpAdviceId = lngNewAdviceId
    mlngTmpSendNo = mlngSendNo
    mblnRefreshState = True
    mstrCurReportViewType = ""
    
    With Me.cbrMain.Options
        If mblnSingleWindow = True Then
            .SetIconSize True, 24, 24
        Else
            .SetIconSize True, 16, 16
        End If
    End With
    
    lngOldFileID = mFileID
    mReportID = 0
    mFileID = 0
    mintReportViewType = -1
    mblnIsReportDelete = False
    mblnModified = False
    mblnReadOnly = True
    mstrҽ������ = ""
    blnPrinted = False
    mblnCanUntread = True
    mblnPrintView = False
    
    If mlngAdviceID <> 0 Then
        If mblnMoved = True Then    '��ת�����򱨸�Ϊֻ��
            mblnReadOnly = True
        Else
            '��ѯҽ��ִ��״̬,���Ƿ��Ժ�鵵
            strSql = "Select a.ִ�й���,c.��Ժ����,c.����״̬,c.���ʱ�� From ����ҽ������ a,����ҽ����¼ b,������ҳ c Where " _
                & " a.ҽ��ID = b.Id And  b.����ID = c.����ID(+) And b.��ҳID = c.��ҳID(+) " _
                & " And a.ҽ��ID= [1] "
                
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            
            If rsTemp.EOF = False Then
                lngStudyState = Nvl(rsTemp!ִ�й���, 0)
                '����ɵı��棬Ϊֻ��״̬
                mblnReadOnly = IIf(lngStudyState = 6 Or lngStudyState = 0, True, False)
                '��Ժ�ҹ鵵�󣬱��治�ɲ���,����״̬Ϊ5��ʾ���鵵
                If mblnReadOnly = False Then mblnReadOnly = IIf(Nvl(rsTemp!��Ժ����) <> "" And (Nvl(rsTemp!����״̬, 0) = 5 Or Nvl(rsTemp!���ʱ��, "") <> ""), True, False)
            End If
            
            '�������ֻ��״̬���ٲ�ѯ���沢��״̬
            If mblnReadOnly = False Then
                If CheckConcurrentReport(Me, mlngAdviceID, True) = False Then
                    mblnReadOnly = True
                End If
            End If
        End If
        
        '��ѯ�����ļ�ID
        strSql = "Select ����ID,RawToHex(��鱨��ID) ��鱨��ID From ����ҽ������ Where ҽ��ID= [1]"
        If mblnMoved = True Then
            strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        
        If rsTemp.EOF = True Then
            '���û�в鵽��¼��˵�����˻�û�б��棬��Ҫ����������Ŀ��������
            strSql = "Select l.������Դ, a.�����ļ�id" & vbNewLine & _
                "From ����ҽ����¼ l, ��������Ӧ�� a" & vbNewLine & _
                "Where l.������Ŀid = a.������Ŀid(+) And a.Ӧ�ó���(+) = Decode(l.������Դ, 2, 2, 4 ,4, 1) And l.Id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            If rsTemp.EOF = True Then
                mFileID = 0
            Else
                mFileID = Nvl(rsTemp!�����ļ�id, 0)
                mintEditType = 0    '��������
            End If
            
            mReportID = 0
            mlngEPRDeptID = 0
            mstrEPR������ = UserInfo.����
            mstrEPR������ = UserInfo.����
            mlngEPRǩ������ = 0
            mdtReportTime = zlDatabase.Currentdate
        Else
            str��鱨��ID = Nvl(rsTemp!��鱨��ID)
            If str��鱨��ID <> "" Then
                 MsgBoxD Me, "�˼����ʹ��PACS���ܱ���༭�����д򿪼���ز�����", vbExclamation, gstrSysName
            Else
                mReportID = Nvl(rsTemp!����Id, 0)
                mintEditType = 1    '��д���棬�������޶������أ�
                '������Ƹ�ʽ���ļ�ID
                strSql = "Select �ļ�ID,����ID,������,������,ǩ������,����ʱ�� From ���Ӳ�����¼  Where Id =[1]"
                If mblnMoved = True Then
                    strSql = Replace(strSql, "���Ӳ�����¼", "H���Ӳ�����¼")
                End If
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
                
                mFileID = rsTemp!�ļ�ID
                mlngEPRDeptID = rsTemp!����ID
                mstrEPR������ = Nvl(rsTemp!������)
                mstrEPR������ = Nvl(rsTemp!������)
                mlngEPRǩ������ = Nvl(rsTemp!ǩ������, 0)
                mdtReportTime = Nvl(rsTemp!����ʱ��, zlDatabase.Currentdate)
            End If
        End If
        
        '��������ļ�ID�Ҳ�������ʾ����������Ŀ��Ӧ�Ĳ����ļ�
        If mFileID = 0 Then
            mlngAdviceID = 0
            '���û���ҵ������ļ�����������ع��ܣ����û�в����ļ�mlngAdviceIDֵ�ͱ��޸�Ϊ0
            mblnReadOnly = True
            mblnPrintView = True
            If str��鱨��ID = "" Then Call MsgBoxD(Me, "δ�ҵ���������Ŀ��Ӧ�Ĳ����ļ����뵽������Ŀ����������")
        ElseIf mFileID <> lngOldFileID Then '���Ƶ��ݷ����ı䣬��Ҫ�ı����Ƶ��ݶ�Ӧ�Ĵ�ӡ���ò˵�
            mblnRefreshRptFormat = False
            mblnʹ���Զ��屨�� = False
            
            '�����Ƶ��ݸı�ʱ�������ѡ�еı����ʽ����һ�δ򿪣�ʹ��ԭ�����õ�Ĭ�ϸ�ʽ
            If lngOldFileID <> 0 Then
                mstrѡ�б����ʽ = ""
            End If
    
            '���ж��Ƿ�ʹ���Զ��屨��
            strSql = "Select ͨ��,��� From �����ļ��б�  Where Id =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�����ӡ��ʽ", mFileID)
            If rsTemp.EOF = False Then
                If Nvl(rsTemp!ͨ��) = 2 Then
                    mblnʹ���Զ��屨�� = True     'ʹ���Զ��屨���ʽ��ӡ
                    
                    If mstr������ <> "ZLCISBILL" & Format(Nvl(rsTemp!���), "00000") & "-2" Then
                        '��ע����еı������뵱ǰ�ı����Ų���ͬʱ������ע����б�����
                        '�Ա������������ӡʱ��������ǰһ�εı�����д�ӡ
                        If mblnSingleWindow = True Then
                            strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
                        Else
                            strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
                        End If
                        
                        mstrѡ�б����ʽ = ""
                        SaveSetting "ZLSOFT", strRegPath, "������", "ZLCISBILL" & Format(Nvl(rsTemp!���), "00000") & "-2"
                    End If
                    
                    mstr������ = "ZLCISBILL" & Format(Nvl(rsTemp!���), "00000") & "-2"
                    mblnRefreshRptFormat = True
                Else
                    mblnʹ���Զ��屨�� = False    'ʹ�ñ༭��ʽ��ӡ
                    'ʹ�ñ༭��ʽ��ӡ������Զ��屨���ʽ������
                    mstrѡ�б����ʽ = ""
                    mstr������ = ""
                End If
            End If
        End If
        
        cbrMain.Item(2).Visible = True
        
        Call InitReportFormat        '��ʼ�������ʽ
        Call RefreshVersion(True)      'ˢ�±���汾
        Call RefreshSigns       'ˢ�±���ǩ��
        Call subShowHistoryList '��д������ʷ
    Else    'ҽ��IDΪ0
        cbrMain.Item(2).Visible = False
        'ֻ��״̬
        mblnReadOnly = True
    End If
    
    
    '�����ݿ��ѯ��������ʹ�ӡ״̬��Ϣ
    strSql = "Select �������,�����ӡ From Ӱ�����¼ Where ҽ��ID=[1] "
    If mblnMoved = True Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��ü�¼�ʹ�ӡ״̬", mlngAdviceID)
    If rsTemp.EOF = False Then
        str������� = Nvl(rsTemp!�������)
        blnPrinted = (Nvl(rsTemp!�����ӡ, 0) = 1)
    End If
    
    '���ݴ�ӡ�����״̬��ȷ�����α����Ƿ������д
    '1���������˴�ӡ��������ˡ�=True�����˿��Ի��ˣ������˿��Լ����޶�
    '2���������˴�ӡ��������ˡ�=False��������˲����Ѵ�ӡ�ı��棬ֻ�б��˿����޶�������ֻ���޶������ܻ��ˣ�������Ϊֻ��
    If blnPrinted And lngStudyState = 5 Then
        If mblnUntreadPrinted = False Then
            '��Ҫ�Ȳ��ұ��α��������ˣ���󱣴��ˣ���һ��������ˡ�
            '��Ϊ���������A��˺�BҲ��˱��棬Ȼ��B���ˣ���ʱ��������B�������������A��֮���ٴ�ӡ��
            
            strSql = "Select Ҫ�ر�ʾ As ǩ������,�����ı� as ǩ��,��ʼ��  From ���Ӳ������� Where �ļ�ID=[1] " _
                            & " And ��������= 8 order by ��ʼ�� "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ǩ����", mReportID)
                    
            If rsTemp.EOF = False Then
                str����� = Split(Nvl(rsTemp!ǩ��), ";")(0)
            End If
            
            If str����� <> UserInfo.���� Then
                mblnReadOnly = True
            Else
                '������ֻ��״̬�������д�����ǲ��ܻ���
                mblnCanUntread = False
            End If
        End If
    End If
    
    '�ͼ����ҽ�������޶��߼���ҽ���ı��棬�򿪱���󣬱���Ϊֻ���ġ�
    '�������ֻ���ڱ����Ѿ�ǩ������ȥ���ǣ�����ǩ������<>0���޸ĺ�δǩ���ģ��ں�����chkEditState�д���
    If mĿ��汾 > 1 And mlngEPRǩ������ <> 0 Then
        '�Լ���д�ı��棬Ӧ���ǿ��Ի��˵�
        '��ȡ��ǰ�û���ǩ������
        thisUserSignLevel = GetUserSignLevel(UserInfo.ID)
        If thisUserSignLevel < mlngEPRǩ������ Then
            mblnReadOnly = True
        End If
    End If
    
    '�жϱ����Ƿ���Ա༭
    Call chkEditState(False)
    
    '�жϱ����Ƿ���Դ�ӡ
    If mblnCheckPrintPara = True Then
        Call chkPrintState
    Else
        mblnCanPrint = True
    End If
    
    '---------------------��ʼ�����ִ���-----------------------------------------------------------
    
    Call ShowTitle(False)
    
    mblnExitAfterSign = IIf(Val(zlDatabase.GetPara("PACS����ǩ�����˳�", glngSys, mlngModule, True, "0")) = 0, False, True)
    
    '��ʼ����������
    '��ʼ���������ݱ༭����
    strSql = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
    If mblnMoved = True Then
        strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡҽ������", mlngAdviceID)
    Do Until rsTemp.EOF
        mstrҽ������ = mstrҽ������ & rsTemp!��Ŀ & ":" & Nvl(rsTemp!����) & vbCrLf
        rsTemp.MoveNext
    Loop
    
    '��ʼ���������ݴ���
    mfrmReportView.txtReview.Text = str�������
    mfrmReportView.txtReview.Enabled = CheckPopedom(mstrPrivs, "���")

    mfrmReportView.zlRefresh mReportID, mblnSingleWindow, mFileID, True, mblnEditable, mstrModifyEdit, mstrҽ������ & vbCrLf & vbCrLf & mstrҽ������, mblnShowWord, mstrFormatInfo, mblnMoved
    '��ʼ������ͼ�񴰿�
    If mblnShowImage = True And (Not mfrmReportImage Is Nothing) Then
        mfrmReportImage.zlRefresh mlngAdviceID, mFileID, mReportID, mblnSingleWindow, mlngShowBigImg, mintImageDblClick, mblnEditable, mblnMoved, _
                                    mintMinImageCount, GetReportImageSelected, mlngModule, mlngDeptID, mlngStudyState, mblnIsSignSave
    End If
    '��ʼ��ר�Ʊ��洰��
    If mblnShowSpecial = True And (Not mfrmReportSpecial Is Nothing) Then
        If mstrSpecialForm <> Report_Form_frmReportCustom Then
            mfrmReportSpecial.zlRefresh Me, mlngAdviceID, mReportID, mblnSingleWindow, mblnEditable, mblnMoved
        Else
            mfrmReportSpecial.Refresh mlngAdviceID, mReportID, mblnEditable, mblnMoved
        End If
    End If
    
    zlRefreshVideoData

    If blnIsDockActive = False Then tmrFocus.Enabled = True
'    If Not ConfigFocus Then
'        '���û�����ý���ʱ�Ĵʾ����......
'    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub AllowLocate(blnIsAllowLocate As Boolean)
    mblnAllowLocate = blnIsAllowLocate
End Sub

Function ConfigFocus() As Boolean
'���ý���
On Error GoTo errHandle
    ConfigFocus = Not mblnSetFocusWithReport
    
'    If GetActiveWindow = Me.hWnd Then Exit Function
    
    If mblnSetFocusWithReport Or mblnSingleWindow Or mblnAllowLocate Then
        If mstrCurReportViewType = "" Or mstrCurReportViewType = ReportViewType_������� Then
            mfrmReportView.rtxtCheckView.SetFocus
        End If
        
        If mstrCurReportViewType = ReportViewType_������ Then
            mfrmReportView.rtxtResult.SetFocus
        End If
        
        If mstrCurReportViewType = ReportViewType_���� Then
            mfrmReportView.rTxtAdvice.SetFocus
        End If
    End If
Exit Function
errHandle:
    err.Clear
End Function


Public Sub zlRefreshVideoData()
    If Not mobjWork_ImageCap Is Nothing Then Call mobjWork_ImageCap.zlRefreshData
End Sub

Public Sub zlEditReport()
    '��¼��ǰ�༭��ʽΪ�������ڷ�ʽ
    mblnSingleWindow = True
    
    'ʹ�ø÷���ʱ��˵����Ҫʹ�ö������ڴ򿪱���༭���������Ҫִ��RestoreWinState�����ָ�����λ��
    Call RestoreWinState(Me, App.ProductName)
    
    Call Me.Show(, mobjOwner)
    
    Call zlRefreshFace
    
    RaiseEvent AfterOpen
End Sub


Private Sub chkEditState(blnShowMessage As Boolean)
    'blnShowMessage---��Ƕ��ʽ��ģʽ�£��Ƿ���ʾ��ʾ��Ϣ
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    mstrModifyEdit = ""
    
    If mblnReadOnly = True Then
        mblnEditable = False
        Exit Sub
    End If
    
    If mĿ��汾 = 1 And CheckPopedom(mstrPrivs, "PACS������д") Then
        If mstrEPR������ = UserInfo.���� Then
            mblnEditable = True
        ElseIf (CheckPopedom(mstrPrivs, "PACS���˱���") And mlngEPRDeptID = mlngDeptID) Then  '�����˱���Ȩ�޵ģ�������д�����ҵı���
            mblnEditable = True
        Else
            mblnEditable = False
            If mblnSingleWindow = True Or blnShowMessage = True Then  '��������ģʽ������Ƕ��ʽ����Ҫ��ʾ��ֱ����ʾ
                MsgBoxD Me, "�������Ѿ���" & mstrEPR������ & "������д��������Ȩ���޸ġ�", vbOKOnly
            End If
        End If
    ElseIf mĿ��汾 > 1 And CheckPopedom(mstrPrivs, "PACS�����޶�") Then
        '�ڱ����޶���״̬�£��б����޶�Ȩ�޵��ˣ�������д�����ҵı��档
        If mstrEPR������ = UserInfo.���� Or mlngEPRǩ������ <> 0 Then   '����������Լ���󱣴�ģ�����ǰ����޸����Ѿ�ǩ��
            mblnEditable = True
        Else
            '�Ѿ��������޶��������,�޸��Ѿ����棬����û��ǩ�������治�ɱ༭����¼�޶�������
            mstrModifyEdit = mstrEPR������
            mblnEditable = False
        End If
    Else
        mblnEditable = False
    End If
    
    If mblnTechReptSame And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        strSql = " select ��鼼ʦ from Ӱ�����¼ where ҽ��id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        
        If rsTemp.RecordCount < 1 Then Exit Sub
        
        If Nvl(rsTemp!��鼼ʦ) <> "" And Nvl(rsTemp!��鼼ʦ) <> UserInfo.���� Then
            mblnEditable = False
        Else
            mblnEditable = True
        End If
    
    End If
End Sub

Public Function chkModified() As Boolean
    
    '�ı��ʽ
    If mHasChangeFormat = True Then
        chkModified = True
        Exit Function
    End If
    
    '�޸ı�������
    If Not mfrmReportView Is Nothing Then
        If mfrmReportView.pModified = True Then
            chkModified = True
            Exit Function
        End If
    End If
    
    '�޸ı���ͼ����ͼ
    If mblnShowImage = True And Not mfrmReportImage Is Nothing Then
        If mfrmReportImage.pMarkModified = True Or mfrmReportImage.pImageModified = True Then
            chkModified = True
            Exit Function
        End If
    End If
    
    '�޸�ר�Ʊ�����Ϣ
    If mblnShowSpecial = True And Not mfrmReportSpecial Is Nothing Then
        If mfrmReportSpecial.pModified = True Then
            chkModified = True
            Exit Function
        End If
    End If
End Function


Private Sub RefreshSigns()
'------------------------------------------------
'���ܣ�ˢ��ǩ������ɾ������ǩ���Ķ������´����ݿ��ȡ��ȷ��ǩ����������ݸ����ݿ��һ�£�ǩ������ˢ��֮����ñ�����
'������ ��
'���أ� ��
'-----------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim OneSign As cEPRSign
    Dim i As Integer
    Dim strSigns As String
    
    '���ԭ��ǩ��
    For i = 1 To mSigns.Count
        mSigns.Remove 1
    Next i
    mSigns.UpdateMaxKey
    
    strSql = "Select Id,������ From ���Ӳ������� Where �ļ�id= [1] And ��������=8 Order By ������"
    If mblnMoved = True Then
        strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
    While rsTemp.EOF = False
        Set OneSign = New cEPRSign
        If OneSign.GetSignFromDB(Val(rsTemp!ID)) = True Then
            OneSign.Key = Nvl(rsTemp!������, 0)
            mSigns.AddExistNode OneSign, IIf(OneSign.Key = 0, False, True)
            strSigns = strSigns & " " & OneSign.ǰ������ & OneSign.����
        End If
        rsTemp.MoveNext
    Wend
    
    '��дǩ���ı���
    mfrmReportView.txtSigns.Text = strSigns
End Sub

Private Sub RefreshVersion(blnIncVer As Boolean)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If mReportID = 0 Then
        '�������������£����汾=1��ǩ������=0
        m���汾 = 1
        mǩ������ = cprSL_�հ�
        mĿ��汾 = 1
    Else
        strSql = "Select ���汾,ǩ������ From ���Ӳ�����¼  Where Id =[1]"
        If mblnMoved = True Then
            strSql = Replace(strSql, "���Ӳ�����¼", "H���Ӳ�����¼")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        m���汾 = Nvl(rsTemp!���汾, 1)
        mǩ������ = Nvl(rsTemp!ǩ������, cprSL_�հ�)
        
        If blnIncVer Then
          mǩ������ = Nvl(rsTemp!ǩ������, cprSL_�հ�)
        Else
          mǩ������ = cprSL_�հ�
        End If
        
        mĿ��汾 = m���汾 + IIf(mǩ������ = cprSL_�հ�, 0, 1)
    End If
End Sub

Private Sub ShowTitle(blnChangeFormat As Boolean)
'blnChangeFormat �Ƿ��޸ĸ�ʽ

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strName As String
    Dim strSex As String
    Dim strAge As String
    Dim lngStyle  As Long
    Dim strDoctor As String
    Dim strCheckNo As String
    Dim strAdvice As String
        
    On Error GoTo errHandle
    
    If blnChangeFormat = True Then  '���ĸ�ʽ
        If mFormatID = 0 Then
            strSql = "Select nvl(b.����,a.����) ����,nvl(b.�Ա�,a.�Ա�) �Ա�,nvl(b.����,a.����) ����,b.ҽ������ From Ӱ�����¼ a ,����ҽ����¼ b  Where a.ҽ��ID= b.id and b.id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            mModelName = "��׼����"
        Else
            strSql = "Select nvl(c.����,a.����) ����,nvl(c.�Ա�,a.�Ա�) �Ա�,nvl(c.����,a.����) ����,a.����,b.����,c.ҽ������ From Ӱ�����¼ a,��������Ŀ¼ b,����ҽ����¼ c  Where a.ҽ��ID=c.id and c.id= [1] And b.Id =[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID, mFormatID)
            If rsTemp.EOF = False Then mModelName = rsTemp!����
        End If
    Else
        If mReportID = 0 Then
            strSql = "Select nvl(b.����,a.����) ����,nvl(b.�Ա�,a.�Ա�) �Ա�,nvl(b.����,a.����) ����,a.����,b.ҽ������ From Ӱ�����¼ a ,����ҽ����¼ b  Where a.ҽ��ID= b.id and b.id = [1]"
            If mblnMoved = True Then
                strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
                strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
            mModelName = "��׼����"
        Else
            strSql = "Select nvl(c.����,a.����) ����,nvl(c.�Ա�,a.�Ա�) �Ա�,nvl(c.����,a.����) ����,a.����,b.��������,b.������,b.���ʱ��,c.ҽ������ From Ӱ�����¼ a,���Ӳ�����¼ b,����ҽ����¼ c  Where a.ҽ��ID=c.id and c.id = [1] And b.Id =[2]"
            If mblnMoved = True Then
                strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
                strSql = Replace(strSql, "���Ӳ�����¼", "H���Ӳ�����¼")
                strSql = Replace(strSql, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID, mReportID)
            If rsTemp.EOF = False Then
                mModelName = rsTemp!��������
                If Nvl(rsTemp!���ʱ��) = "" Then
                    strDoctor = Nvl(rsTemp!������)
                End If
            End If
        End If
    End If

    If rsTemp.EOF = False Then
        strName = Nvl(rsTemp!����)
        strSex = Nvl(rsTemp!�Ա�)
        strAge = Nvl(rsTemp!����)
        strCheckNo = Nvl(rsTemp!����)
        strAdvice = Nvl(rsTemp!ҽ������)
        mstrҽ������ = strAdvice
    End If
    
    'û�����ر��⣬�Ÿ��±�����
    lngStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    If (lngStyle And WS_CAPTION) <> 0 Then
        Me.Caption = IIf(mĿ��汾 > 1, "[�����޶�]", "[������д]") & "   ��������" & strName & " �Ա�" & strSex & " ���䣺" & strAge & "��   ����ҽ����" & UserInfo.���� _
                     & " ���ţ�" & strCheckNo & "   ҽ����" & strAdvice
    End If
    mstrFormatInfo = IIf(mĿ��汾 > 1, "[�����޶�]", "[������д]") & "   " & mModelName
    If blnChangeFormat = False Then
        If mReportID = 0 Then
            mstrFormatInfo = mstrFormatInfo & " �±��棬��δ��ʼ��д"
        ElseIf strDoctor <> "" Then
            mstrFormatInfo = mstrFormatInfo & " " & strDoctor & " ������д����"
        End If
    End If
    
    If mstrѡ�б����ʽ <> "" Then
        If InStr(mstrFormatInfo, vbCrLf) <> 0 Then
            mstrFormatInfo = Left(mstrFormatInfo, InStr(mstrFormatInfo, vbCrLf) - 1)
        End If
        mstrFormatInfo = mstrFormatInfo & vbCrLf & "��ӡ��ʽ��" & mstrѡ�б����ʽ
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitFaceScheme()
'��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane, Pane5 As Pane
    Dim Pane6 As Pane
    Dim i As Integer
    Dim intPaneID As Integer
    
    '����Pane��ID˳�� 1-���������2-��ʷ���棻3-�ʾ�ʾ����4-����ͼ��5-��Ƶ�ɼ���6-ר�Ʊ��档
    
    If mlngDeptID = 0 Then Exit Sub
    
    On Error Resume Next
    If Not mfrmReportImage Is Nothing Then Call mfrmReportImage.InitParaForAfterImage(mlngDeptID, mlngModule)
        
    '����������ʾ����
    With Me.dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .PanelPaintManager.BoldSelected = True
        .TabPaintManager.Position = xtpTabPositionLeft  'TAB�ŵ������ʾ
'        .TabPaintManager.OneNoteColors = True           'һ��TABһ����ɫ��ʾ
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .TabPaintManager.BoldSelected = True
        dkpMain.Options.DefaultPaneOptions = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End With
    
    '�ȴ�ע����ȡԤ�����úõĴ��ڲ��֣�Ȼ�����������
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmReport" & IIf(mblnSingleWindow = True, "\SingleWindow\", "\") & mlngModule & "\" & TypeName(dkpMain), _
                dkpMain.Name & mlngDeptID, "")
    End If
    
    
    
    '��ʷ����
    intPaneID = PaneHasShow("��ʷ����")
    If intPaneID = 0 Then
        '������ʷ����ҳ��
        Set Pane1 = dkpMain.CreatePane(1, 300, 150, DockLeftOf)
        Pane1.Title = "��ʷ����" '��ʷ����
        'Pane1.Options = PaneNoCaption
    Else
        Set Pane1 = dkpMain.Panes(intPaneID)
    End If
    
    '�������
    intPaneID = PaneHasShow("�������")
    If intPaneID = 0 Then
        '���ؼ������ҳ��
        Set Pane2 = dkpMain.CreatePane(2, 600, 150, DockRightOf, Pane1)
        Pane2.Title = "�������"
        Pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
    End If
    
    '�ʾ�ʾ��
    intPaneID = PaneHasShow("�ʾ�ʾ��")
    If intPaneID = 0 And mblnShowWord = True Then
        '���شʾ�ʾ��ҳ��
        Set Pane3 = dkpMain.CreatePane(3, 300, 150, DockTopOf, Pane1)
        Pane3.Title = "�ʾ�ʾ��"
        'Pane3.Options = PaneNoCaption
        Pane3.AttachTo Pane1
        
    ElseIf intPaneID <> 0 And mblnShowWord = False Then
        '������ʾ�ʾ�ʾ��ҳ�棬ж�ظ�ҳ��
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    End If
    
    '����ͼ
    intPaneID = PaneHasShow("����ͼ")
    If intPaneID = 0 And mblnShowImage = True Then
        '���ر���ͼҳ��
        Set pane4 = dkpMain.CreatePane(4, 300, 150, DockTopOf, Pane1)
        pane4.Title = "����ͼ"
        'pane4.Options = PaneNoCaption
        pane4.AttachTo Pane1
    ElseIf intPaneID <> 0 And mblnShowImage = False Then
        '������ʾ����ͼҳ�棬ж�ظ�ҳ��
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    End If
    
    '��Ƶ�ɼ�
    intPaneID = PaneHasShow("��Ƶ�ɼ�")
    If intPaneID = 0 And mblnShowVideoCapture = True Then
        '������Ƶ�ɼ�ҳ��
        Set Pane5 = dkpMain.CreatePane(5, 300, 150, DockTopOf, Pane1)
        Pane5.Title = "��Ƶ�ɼ�"
        'Pane5.Options = PaneNoCaption
        Pane5.AttachTo Pane1
    ElseIf intPaneID <> 0 And mblnShowVideoCapture = False Then
        '������ʾ��Ƶ�ɼ�ҳ�棬ж�ظ�ҳ��
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    ElseIf intPaneID <> 0 And mblnShowVideoCapture Then
        '�����������˳�����վ���˳�ʱδ������������£�����ע����б���dkpMain������ֵ�����а�����Ƶ�ɼ�ҳ��ı��⣬
        '����Ϊ"�����塿��Ƶ�ɼ�"������һ������ʱ����ע����л�ȡdkpMain������ֵ����ʱû��������������£���Ƶ�ɼ�ҳ��
        '�ı��⻹��Ϊ"�����塿��Ƶ�ɼ�"���������������
        If dkpMain.Panes(intPaneID).Title <> "��Ƶ�ɼ�" Then dkpMain.Panes(intPaneID).Title = "��Ƶ�ɼ�"
    End If
    
    'ר�Ʊ���
    intPaneID = PaneHasShow("ר�Ʊ���")
    If intPaneID = 0 And mblnShowSpecial = True Then
        '����ר�Ʊ���ҳ��
        Set Pane6 = dkpMain.CreatePane(6, 300, 150, DockTopOf, Pane1)
        Pane6.Title = "ר�Ʊ���"
        'Pane6.Options = PaneNoCaption
        Pane6.AttachTo Pane1
    ElseIf intPaneID <> 0 And mblnShowSpecial = False Then
        '������ʾר�Ʊ���ҳ�棬ж�ظ�ҳ��
        Call dkpMain.DestroyPane(dkpMain.Panes(intPaneID))
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'ȫ�����غ���ʾ��֮�������ñ�ѡ�е�TAB
    For i = 1 To dkpMain.PanesCount
        Call DkpMain_AttachPane(dkpMain.Panes(i))
        
        If dkpMain.Panes(i).Title = "�ʾ�ʾ��" _
            And (mintPaneID <= 0 Or mintPaneID > dkpMain.PanesCount) Then
            mintPaneID = i
        End If
    Next i

    If mintPaneID <= dkpMain.PanesCount Then
        Call dkpMain.Panes(mintPaneID).Select
    End If
End Sub

Private Function PaneHasShow(strTitle As String) As Integer
'------------------------------------------------
'���ܣ���ѯDockingPane�е�Pane�Ƿ��Ѿ���ʾ
'������ strTitle --- Pane��Title
'���أ�����ҵ�Pane������Pane��ID������Ҳ���������0
'------------------------------------------------
    Dim i As Integer
    
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).Title Like "*" & strTitle & "*" Then
            PaneHasShow = i
            Exit Function
        End If
    Next i

    PaneHasShow = 0
End Function

Private Sub PrintReport(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    '�жϱ����Ƿ���Դ�ӡ
    If mblnCheckPrintPara = True Then
        Call chkPrintState
    End If
    
    If PromptModify = False And mReportID = 0 Then
            MsgBoxD Me, "�´����ı��棬û�б����޷���ӡ��Ԥ�������ȱ��汨�档"
            mblnModified = True
    ElseIf mblnCanPrint = False Then
        MsgBoxD Me, "��ǰ����δ��ˣ����ܴ�ӡ�����飡", vbInformation, gstrSysName
    Else    '���Դ�ӡ
        '��ӡǰ�ж��Ƿ���Ҫ��ʾ�����Ժ�Ӱ������
        If control.ID = conMenu_File_Print And mlngHintType = 2 Then 'mlngHintType = 2��ʾ��ӡǰ����
            Dim strResultInput As String
            
            strResultInput = ""
            If mblnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
                gstrSQL = "ZL_Ӱ����_���(" & mlngAdviceID & ",0)"
                zlDatabase.ExecuteProcedure gstrSQL, "���������"
            End If
                
            strSql = "Select B.Σ��״̬, A.�������, B.Ӱ������, B.��������, B.������� " & _
                     "From ����ҽ������ A, Ӱ�����¼ B " & _
                     "Where A.ҽ��id = B.ҽ��id and B.ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�������", mlngAdviceID)

'                    If IsNull(rsTemp!Σ��״̬) And mintCriticalValues <> 0 Then strResultInput = "Σ��״̬|"
            If IsNull(rsTemp!�������) And Not mblnIgnoreResult Then strResultInput = strResultInput & "�������|"
            If IsNull(rsTemp!Ӱ������) And mstrImageLevel <> "" And mintImageLevel <> 0 And CheckPopedom(mstrPrivs, "Ӱ���ʿ�") Then strResultInput = strResultInput & "Ӱ������|"
            If IsNull(rsTemp!��������) And mstrReportLevel <> "" And mintReportLevel <> 0 And CheckPopedom(mstrPrivs, "�����ʿ�") Then strResultInput = strResultInput & "��������|"
            If IsNull(rsTemp!�������) And mintConformDetermine <> 0 Then strResultInput = strResultInput & "�������|"
                
            If strResultInput <> "" Then Call PromptResult(mlngAdviceID, mlngModule, Me, mlngDeptID, strResultInput)
        End If
            
        '��ӡ�������Ԥ������
        If mblnʹ���Զ��屨�� = True Then
            mblnPrintOK = False
            Call subPrintReport(IIf(control.ID = conMenu_File_Preview, False, True), control.ID = conMenu_File_BatPrint)
        Else        'ʹ�ñ༭ģʽ��ӡ�����ò����Ĵ�ӡ����
            mobjReport.zlRefresh 0, 0, , , , mlngModule
            mobjReport.zlRefresh mlngAdviceID, UserInfo.����ID, , , mblnCanPrint, mlngModule
            mblnPrintOK = False     '��Ǵ�ӡ�Ƿ���ɣ���AfterPrinted�¼��б����ó�True
            mobjReport.zlExecuteCommandBars control
        End If
        
        '��ӡ���˳�
        If mblnExitAfterPrint = True And control.ID = conMenu_File_Print _
            And mblnSingleWindow = True And mblnPrintOK = True Then
            Call PromptModify
            Call SetMenuDownState(False)
            Unload Me
        Else
            'ˢ�½��沼��
            'dkpMain.RecalcLayout
'                    Me.Refresh
        End If
    End If
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim NewControl As XtremeCommandBars.CommandBarControl
    Dim strInfo As String
    
    If mblnMenuDownState Then Exit Sub
    
    mblnMenuDownState = True
    
    Select Case control.ID
        Case conMenu_PacsReport_Save        '���汨��
            Call SaveReport(True)
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_BatPrint       '��ӡ����,Ԥ������,������ӡ
            Call PrintReport(control)
        Case conMenu_Edit_Modify       '�ò����༭���򿪱���
                mobjReport.zlRefresh 0, 0, , , , mlngModule
                If mĿ��汾 > 1 Then       '�޶�ģʽ
                    Set NewControl = cbrMain.FindControl(, conMenu_Edit_Audit, False)
                    mobjReport.zlRefresh mlngAdviceID, UserInfo.����ID, , , , mlngModule
                    mobjReport.zlExecuteCommandBars NewControl
                Else
                    mobjReport.zlRefresh mlngAdviceID, UserInfo.����ID, , , , mlngModule
                    mobjReport.zlExecuteCommandBars control
                End If
        Case conMenu_File_Open, conMenu_File_ExportToXML, conMenu_Tool_Search      '���ı���,����XML���������
            mobjReport.zlRefresh 0, 0, , , , mlngModule
            mobjReport.zlRefresh mlngAdviceID, UserInfo.����ID, , , , mlngModule
            mobjReport.zlExecuteCommandBars control
        Case conMenu_PacsReport_Sign                        'ǩ��
            Call AddSign
        Case conMenu_PacsReport_DelSign                     '����
            Call DoUntread
        Case conMenu_PacsReport_Reject                      '����
            Call RejectReport
        Case conMenu_PacsReport_RejectHistory               '������ʷ
            Call ShowRejectHistory
        Case conMenu_PacsReport_VerifySign_Item             'ǩ����֤
            Call FuncAdviceSignVerify(Val(control.Parameter), mblnMoved)
        Case conMenu_PacsReport_SelFormat_Item              'ѡ���ʽ
            Call ChangeFormat(Val(control.Parameter))
        Case conMenu_PacsReport_RepFormat_Item              'ѡ���ӡ��ʽ
            Call subChangeRptFormat(control.Index)
        Case conMenu_PacsReport_FontSetDefault To conMenu_PacsReport_FontSetUser            '�����ı�������
            Dim cbrEdit As CommandBarEdit
                
            Set cbrEdit = cbrMain.FindControl(xtpControlEdit, conMenu_PacsReport_FontSetUser, True, True)
        
            If control.ID = conMenu_PacsReport_FontSetUser Then
            '������Զ����ֺţ��ж��Ƿ���Ϲ���
                
                If Not CheckUserFontValidate(cbrEdit.Text) Then
                '�����Ϲ����൱������ʧ��
                    cbrEdit.Text = ""
                    mblnMenuDownState = False
                    Exit Sub
                End If
                Call SetMeneFontSize(Abs(Val(cbrEdit.Text)))
                Call mobjOwner.DoFontSize(mblnSingleWindow, (Abs(Val(cbrEdit.Text))))
                Call zlDatabase.SetPara("������ʾ�ֺ�", (Abs(Val(cbrEdit.Text))), glngSys, glngModul)
            Else
            '�����Զ����ֺţ�ǰ��򹴣��Զ���textΪ�ձ�ʾδѡ���Զ����ֺ�
                cbrEdit.Text = ""
                control.Checked = True
                Call SetMeneFontSize(Val(control.Caption))
                Call mobjOwner.DoFontSize(mblnSingleWindow, Val(control.Caption))
                Call zlDatabase.SetPara("������ʾ�ֺ�", Val(control.Caption), glngSys, glngModul)
            End If
        Case conMenu_File_PrintSet                          '��ӡ����
            Call zlPrintSet
        Case conMenu_Edit_Delete                            'ɾ������
            If mReportID = 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            strInfo = "���ɾ����ݡ�" & mModelName & "����"
            If MsgBoxD(Me, strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            strSql = "Zl_���Ӳ�����¼_Delete(" & mReportID & ")"
            
            mblnIsReportDelete = True
            
            
            zlDatabase.ExecuteProcedure strSql, Me.Caption
            
            err = 0: On Error GoTo 0
            RaiseEvent AfterDeleted(mlngAdviceID)
            
            Call Me.zlRefreshFace(True)
        Case conMenu_PacsReport_ClearWritingState           '������桰�����С����
            strSql = "select ������� from Ӱ�����¼ where ҽ��id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���������", mlngAdviceID)
            
            If rsTemp.RecordCount <= 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            If Trim(Nvl(rsTemp!�������)) = "" Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            strInfo = "�������״̬���������ڴ����У�ȷ��Ҫ�����ݱ����״̬��"
            If MsgBoxD(Me, strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            Call UpdateReporter(mlngAdviceID, "")
            
        Case conMenu_PacsReport_History                     '��ʾ�����޶���ʷ
            Call frmReportHistory.zlShowMe(Me, mlngAdviceID, mReportID)
        Case conMenu_PacsReport_SaveWord                    '����ʾ�ʾ��
            Call subSaveWord(0)
        Case conMenu_PacsReport_AddNumber                   '���ı���������
            Call AddNumber
        Case conMenu_PacsReport_PrivOrder                   '��һ��ҽ��
            Call ChangeOrder(1)
        Case conMenu_PacsReport_NextOrder                   '��һ��ҽ��
            Call ChangeOrder(2)
        
'        Case comMenu_Petition_Capture                       '�鿴ɨ�赥
'            Call comMenu_Petition_ɨ�����뵥
            
        Case conMenu_PacsReport_Default                     '����Ĭ�Ͻ���
            Call ReStoreFace
            
        Case conMenu_File_Exit                              '   �˳�
            Call PromptModify
            
            mblnMenuDownState = False
            
            Unload Me
        Case Else
        
    End Select
    
    mblnMenuDownState = False
    Exit Sub
errHandle:
    mblnMenuDownState = False
    If ErrCenter = 1 Then Resume
End Sub


Private Sub ReStoreFace()
'ɾ��ע�������Ĭ�Ͻ���
 On Error GoTo errHandle
 
    '�رչ���վ �������ý��沼��
    If MsgBoxD(Me, "�ָ�����Ĭ�ϲ�����Ҫ�رչ���վ���Ƿ������", vbYesNo, gstrSysName) = vbYes Then
        Unload mobjOwner
    Else
        Exit Sub
    End If
    

    Call ClearFaceConfig
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ClearFaceConfig()
On Error Resume Next
 Dim strReportRegPath As String
 Dim strImageRegPath As String
 Dim strViewRegPath As String
 Dim strWordRegPath As String
 
 Dim strIndividualPath As String
 
    If mblnSingleWindow = True Then
        strReportRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
        strImageRegPath = "����ģ��\" & App.ProductName & "\frmReportImage\SingleWindow"
        strViewRegPath = "����ģ��\" & App.ProductName & "\frmReportView\SingleWindow"
        strWordRegPath = "����ģ��\" & App.ProductName & "\frmReportWord\SingleWindow"
    Else
        strReportRegPath = "����ģ��\" & App.ProductName & "\frmReport"
        strImageRegPath = "����ģ��\" & App.ProductName & "\frmReportImage"
        strViewRegPath = "����ģ��\" & App.ProductName & "\frmReportView"
        strWordRegPath = "����ģ��\" & App.ProductName & "\frmReportWord"
    End If
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        strIndividualPath = "����ģ��\" & App.ProductName & "\frmReport\" & mlngModule & "\Dockingpane"
        
        Call DeleteSetting("ZLSOFT", strIndividualPath, "dkpMain" & mlngDeptID)
    End If
   
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CX1")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CX2")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CX3")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CY21")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "CY3")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "PicHistoryX")
    Call DeleteSetting("ZLSOFT", strReportRegPath, "PicHistoryY")

    Call DeleteSetting("ZLSOFT", strImageRegPath, "CY1")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "CY2")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "CY3")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "MarkW")
    Call DeleteSetting("ZLSOFT", strImageRegPath, "RptImgW")
    
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY1")
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY2")
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY3")
    Call DeleteSetting("ZLSOFT", strViewRegPath, "CY4")
    
    
    Call DeleteSetting("ZLSOFT", strWordRegPath, "PrivateWordH")
    Call DeleteSetting("ZLSOFT", strWordRegPath, "WordShowH")
    Call DeleteSetting("ZLSOFT", strWordRegPath, "WordTreeH")
    
err.Clear
End Sub

'Private Sub comMenu_Petition_ɨ�����뵥()
''ɨ�����뵥
'On Error GoTo errFree
'    Dim frmPetitionCap As New frmPetitionCapture
'    Dim rsTemp As ADODB.Recordset
'    Dim strSQL As String
'
'
'    strSQL = "select a.����,a.����,a.�Ա�,a.ҽ������,b.�����,b.סԺ��,c.���� from ����ҽ����¼ a,������Ϣ b,���ű� c where a.����id = b.����id and a.���˿���id = c.id and a.id = [1]"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�õ�������Ϣ", mlngAdviceID)
'
'    If rsTemp.RecordCount = 0 Then
'         MsgBoxD Me, "û���ҵ��ò�����ؼ�¼", vbInformation, gstrSysName
'         Exit Sub
'    End If
'
'    '��ɨ�����뵥����
'    Call frmPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
'                                                mlngDeptId, _
'                                                Nvl(rsTemp!����), _
'                                                Nvl(rsTemp!����), _
'                                                Nvl(rsTemp!����), _
'                                                Nvl(rsTemp!�Ա�), _
'                                                Nvl(Mid(rsTemp!ҽ������, 1, InStr(rsTemp!ҽ������, ":") - 1)), _
'                                                Nvl(Mid(rsTemp!ҽ������, InStr(rsTemp!ҽ������, ":") + 1, Len(rsTemp!ҽ������))), True, False, _
'                                                mlngAdviceID)
'
'errFree:
'    Call Unload(frmPetitionCap)
'    Set frmPetitionCap = Nothing
'End Sub

Private Sub RefreshViewTag(rText As RichTextBox)
    Dim strItem() As String
    Dim i As Integer
    Dim intCnt As Integer
    
    '�޸ĸ��ı����TAG,���TAGΪ�գ�����ʱ����¼
    If rText.tag <> "" Then
        strItem = Split(rText.tag, "|")
        rText.tag = ""
        strItem(15) = Nvl(rText.SelFontName, "����")     'FontName

        strItem(17) = Nvl(rText.SelBold, "False")    'FontBold
        strItem(18) = Nvl(rText.SelItalic, "False")    'FontItalic
        
        For i = 0 To UBound(strItem()) - 1
            rText.tag = rText.tag & strItem(i) & "|"
        Next i
                
    End If
End Sub


Private Sub DoUntread()
'���ˣ�����ǩ�����޶�
    Dim lngVersion As Long
    Dim lngSignKey As Long
    Dim strSql As String
    Dim arrSQL() As String
    Dim blIsUntread As Boolean
    Dim intRobackType As Integer '����ǩ������
    Dim i As Long
    
    If mSigns.Count = 1 Then  'ֻ��һ��ǩ������ʾ��ǰ����дģʽ�µĻ���
        If frmEPRUntread.ShowMe(mReportID, cprET_�������༭, lngVersion, lngSignKey, Me) = False Then Exit Sub
    Else
        If frmEPRUntread.ShowMe(mReportID, cprET_���������, lngVersion, lngSignKey, Me) = False Then Exit Sub
    End If
    If lngSignKey > 0 Or lngVersion > 0 Then
        If MsgBoxD(Me, "ע�⣺���˲��������ɻָ����Ƿ������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    End If
    
    '�������ֻ��˷�ʽ
    If lngSignKey > 0 Then
        '�ȴ�������ǩ����Ȼ�������ǩ��
        If mSigns("K" & lngSignKey).ǩ����ʽ = 2 Then
            '����ǩ����֤
            err.Clear: On Error Resume Next
            If gobjESign Is Nothing Then
                Set gobjESign = Interaction.GetObject(, "zl9ESign.clsESign")
                If gobjESign Is Nothing Then Set gobjESign = CreateObject("zl9ESign.clsESign")
                If err <> 0 Then err = 0
                
                If Not gobjESign Is Nothing Then
                    If Not gobjESign.Initialize(gcnOracle, glngSys) Then
                        MsgBoxD Me, "����֤���ʼ��ʧ�ܣ���ʹ����ȷ������֤�顣", vbInformation + vbOKOnly, "��дǩ��"
                        Exit Sub
                    End If
                Else
                    MsgBoxD Me, "����ǩ��������ʼ��ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            If Not gobjESign.CheckCertificate(gstrDBUser) Then
                '����֤��ͣ��ʱ�����Լ�������ǩ������
                If Not gobjESign.CertificateStoped(UserInfo.����) Then
                    Exit Sub
                End If
            End If
        End If
        
        ReDim arrSQL(1)
        
        '���ǩ��,�������ʽ
        SaveReportFormat mSigns("K" & lngSignKey), False, arrSQL
                
        intRobackType = CheckSignRollbackType(mSigns("K" & lngSignKey).ID, mReportID)
        If mSigns.Count = 1 And (intRobackType = 2 Or intRobackType = 3) Then intRobackType = 4
        
        For i = 0 To UBound(arrSQL)
            If Trim(arrSQL(i)) <> "" Then
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ǩ��")
            End If
        Next i
        
        blIsUntread = True
        mSigns.Remove "K" & lngSignKey
        
    ElseIf lngVersion > 1 Then  '�����޶�
        'ֱ���޸����ݿ����ݾͿ�����  '�ѻ����޶����浽���ݿ�
        strSql = "ZL_Ӱ�񱨸����(0," & mReportID & "," & lngVersion & ")"
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        blIsUntread = False
        
        Call chkEditState(False)
        'ˢ�¸����ı����壬ͼ���岻��Ҫˢ��
        mfrmReportView.zlRefresh mReportID, mblnSingleWindow, mFileID, False, mblnEditable, mstrModifyEdit, mstrҽ������ & vbCrLf & vbCrLf & mstrҽ������, mblnShowWord, mstrFormatInfo, mblnMoved
        If mblnShowSpecial = True Then
            If mstrSpecialForm <> Report_Form_frmReportCustom Then
                mfrmReportSpecial.zlRefresh Me, mlngAdviceID, mReportID, mblnSingleWindow, mblnEditable, mblnMoved
            Else
                mfrmReportSpecial.Refresh mlngAdviceID, mReportID, mblnEditable, mblnMoved
            End If
        End If
    End If

    Call RefreshVersion(True)
    Call RefreshSigns
    '���������� ������
    Call ShowTitle(False)
    '�ָ��޸ı��
    Call subSetModifyFlag(False)
    
    If blIsUntread = True Then
    '�����ж��ǻ���ǩ�����ǻ����޶�
    '����ǩ�����������
        If intRobackType = 1 Then
        '����ǩ��
            Call AfterReportSaved(mlngAdviceID, 4)
        ElseIf intRobackType = 2 Or intRobackType = 3 Then
        '���ǩ��
            Call AfterReportSaved(mlngAdviceID, 5)
        ElseIf intRobackType = 4 Then
        '����ֱ����˵����
            Call AfterReportSaved(mlngAdviceID, 7)
        End If
    Else
    '�����޶�
        Call AfterReportSaved(mlngAdviceID, 3)
    End If
    
    If mblnSingleWindow = True Then
        '���ڵ������ڣ�ˢ�´�������
        Call zlRefreshFace(True)
    End If
End Sub

Private Function SaveReport(blnRaiseEvent As Boolean, Optional blnIncVer As Boolean = True) As Boolean
'------------------------------------------------
'���ܣ����汨�棬���汨���ʽ�����ݣ����ǲ�����ǩ��
'������     blnRaiseEvent -- �Ƿ񴥷����汣����ɵ��¼���True-�����¼���False-�������¼�����ǩ��֮ǰ����ʱ��Ӧ�ò������¼�����ǩ���Ĺ����е�������
'���أ� ��
'-----------------------------------------------
    Dim lngSaveAdviceID As Long '��¼��ǰ��ҽ��ID���ڱ��汣��Ĺ����У�����ҽ��ID�ᱻ���ⲿ�ı�
    Dim strOldReportViewType As String
    Dim blnIsSignStart As Boolean
    Dim blnSaveItemOk As Boolean
    
    On Error GoTo err
    
    'OutputDebugString "ZLPACS>>SaveReport:1 ��ʼִ�� ҽ��IDΪ [" & mlngAdviceID & "] �ı��汣��..."
    
    SaveReport = False
    
    If mblnIsSignSave Then blnIsSignStart = True
    mblnIsSignSave = True
    
    '�жϱ����ı��γ����Ƿ񳬹�2000���ַ����������������ʾ�����˳�
    If Len(mfrmReportView.rtxtCheckView.Text) > 2000 Or Len(mfrmReportView.rtxtResult.Text) > 2000 _
        Or Len(mfrmReportView.rTxtAdvice.Text) > 2000 Then
        If Not blnIsSignStart Then mblnIsSignSave = False
        
        MsgBoxD Me, "�����м�����������������߽������������2000����ɾ���������ֺ��ٱ��档", vbInformation, gstrSysName
        Exit Function
    End If
    
    lngSaveAdviceID = mlngAdviceID
    
    'OutputDebugString "ZLPACS>>SaveReport:2 ���ñ�����Ŀ���淽��."
    
    If mHasChangeFormat = True Then     '�����˸�ʽ��Ҫ���ݸ�ʽID�����´�������
        If mFormatID = 0 Then
            blnSaveItemOk = SaveReportItems(True, 0)
        Else
            blnSaveItemOk = SaveReportItems(True, 1)
        End If
        mHasChangeFormat = False
    Else
        If mReportID = 0 Then    '��������
            blnSaveItemOk = SaveReportItems(True, 0)
        Else
            blnSaveItemOk = SaveReportItems(False, 0)
        End If
    End If
    
    'OutputDebugString "ZLPACS>>SaveReport:3 ������Ŀ���淽���������."
    
    '���汣��ʧ��
    If blnSaveItemOk = False Then
        'OutputDebugString "ZLPACS>>SaveReport:4 ������Ŀ���淽������ʧ��."
        
        subSetModifyFlag True
        
        If Not blnIsSignStart Then mblnIsSignSave = False
        Exit Function
    End If
    
    mModified = False
    
    'OutputDebugString "ZLPACS>>SaveReport:5 ��ʼˢ�°汾��Ϣ."
    
    
    Call RefreshVersion(blnIncVer)
    
    'OutputDebugString "ZLPACS>>SaveReport:6 ��ʾ�������."
'    '���������� ������
    Call ShowTitle(False)
    '�ָ��޸ı��
    Call subSetModifyFlag(False)
    
    'OutputDebugString "ZLPACS>>SaveReport:7 ��ձ��������."
    '��ձ��������
    Call UpdateReporter(lngSaveAdviceID, "")
    
    mdtReportTime = GetReportLastSaveTime(lngSaveAdviceID)
    
    'OutputDebugString "ZLPACS>>SaveReport:8 �������汣������¼�."
    '�������汣����ɵ��¼�
    If blnRaiseEvent Then Call AfterReportSaved(lngSaveAdviceID, 0)
    
    'OutputDebugString "ZLPACS>>SaveReport:9 ���汣������¼����ý���."
    
    If mblnSingleWindow = True And blnRaiseEvent Then
        '���ڵ������ڣ�������������¼�����ˢ�´�������
        If mblnExitAfterSign = False Then
            strOldReportViewType = mstrCurReportViewType
            Call zlRefreshFace(True)
            mstrCurReportViewType = strOldReportViewType
        End If
    End If
        
    SaveReport = True
    If Not blnIsSignStart Then mblnIsSignSave = False
    
    'OutputDebugString "ZLPACS>>SaveReport:10 ҽ��IDΪ[" & mlngAdviceID & "]�ı��汣�����."
    
    Exit Function
err:
    SaveReport = False
    If Not blnIsSignStart Then mblnIsSignSave = False
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddSign()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngPatientID As Long
    Dim lngPageID As Long
    Dim strZipFile As String
    Dim strTemp As String
    Dim OneSign As cEPRSign
    Dim lngKey As Long
    Dim lngMaxSignLevel As Long
    Dim int��ʼ��  As Integer   '���α���ǩ���Ŀ�ʼ��
    Dim lngSaveType As Long
    Dim arrSQL() As String
    Dim i As Long
    Dim blֱ�����ǩ�� As Boolean
        
    blֱ�����ǩ�� = False
    
    If CheckConcurrentReport(mobjOwner, mlngAdviceID) = False Then Exit Sub
    
'    OutputDebugString "ZLPACS>>AddSign:1 ����ҽ��IDΪ[" & mlngAdviceID & "]��ǩ������..."
    
    mblnIsSignSave = True
    On Error GoTo errHandle
        '�ȱ��汨��,���ǲ��������汣����ɵ��¼�,Ȼ���ٴ���ǩ����ǩ��֮�󴥷����汣����ɵ��¼�
        
'        OutputDebugString "ZLPACS>>AddSign:2 ��ʼ���ñ��汣��."
        
        If SaveReport(False, False) = False Then
'            OutputDebugString "ZLPACS>>AddSign:3 ���汣�����ʧ��."
            mblnIsSignSave = False
            Exit Sub
        End If
            
            
'        OutputDebugString "ZLPACS>>AddSign:4 ��ѯǩ����Ϣ."
        
        '��ѯ����ID����ҳid
        strSql = "Select ����id,��ҳid From  ����ҽ����¼ Where id= [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        lngPatientID = Nvl(rsTemp!����ID, 0)
        lngPageID = Nvl(rsTemp!��ҳID, 0)
        
        '��ȡ���ǩ������
        strSql = "Select Ҫ�ر�ʾ As ǩ������,�����ı� as ǩ��,��ʼ��  From ���Ӳ������� Where �ļ�ID=[1] " _
                            & " And ��������= 8 order by ǩ������ desc "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ǩ������", mReportID)
        If rsTemp.EOF = False Then
            lngMaxSignLevel = Nvl(rsTemp!ǩ������, 0)
        End If
        
        '���㱾��ǩ���Ŀ�ʼ�棬����ǩ���汾
        If (mModified Or (mintEditType = 2 And mǩ������ = cprSL_�հ�)) Or (mintEditType = 1 Or mintEditType = 0) Then
            int��ʼ�� = mĿ��汾
        Else
            int��ʼ�� = mĿ��汾 - 1
        End If
        
        If int��ʼ�� > 16 Then
            mblnIsSignSave = False
            MsgBoxD Me, "Ŀǰϵͳ֧�ֵ����汾��Ϊ16������˻�����������", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
        
'        OutputDebugString "ZLPACS>>AddSign:5 ����ǩ������."
        '����ǩ������
        Set OneSign = frmEPRSign.ShowMe(Me, mlngPassType, mReportID, lngPatientID, lngPageID, mstrPrivs, lngMaxSignLevel, int��ʼ��)
        
        lngSaveType = 0
        If Not OneSign Is Nothing Then
'            OutputDebugString "ZLPACS>>AddSign:6 ���ǩ���ж�."
            
            'ǩ���ˣ����ж�һ�£��Ƿ�ڶ������ǩ�������������ʾ�Ƿ�ȷ��Ҫǩ��
            If OneSign.ǩ������ = cprSL_���� Then
                If mSigns.Count >= 1 Then
                    If MsgBoxD(Me, "���α����Ѿ���ǩ���ˣ��Ƿ�Ҫ�ٴ�ǩ����", vbOKCancel, "���ǩ���ظ�") = vbCancel Then
                        mblnIsSignSave = False
                        Exit Sub
                    End If
                End If
            End If
            
            'ǩ���ˣ����ж�һ�£��Ƿ�ڶ������ǩ�������������ʾ�Ƿ�ȷ��Ҫǩ��
            If OneSign.ǩ������ = cprSL_���� Then
                If lngMaxSignLevel >= 3 Then
                    '�ٴ����ǩ��
                    If MsgBoxD(Me, "���α����Ѿ������ǩ���ˣ��Ƿ�Ҫ�ٴ����ǩ����", vbOKCancel, "�����ǩ���ظ�") = vbCancel Then
                        mblnIsSignSave = False
                        Exit Sub
                    End If
                End If
            End If
                        
            If mSigns.Count = 0 And (OneSign.ǩ������ = cprSL_���� Or OneSign.ǩ������ = cprSL_����) Then blֱ�����ǩ�� = True
            
            'ǩ���ˣ����汨�����ݺ�ǩ��
            lngKey = mSigns.AddExistNode(OneSign)
            
            ReDim arrSQL(1)
            
'            OutputDebugString "ZLPACS>>AddSign:7 ����ǩ������."
            
            'ǩ��ֱ�ӵ���SaveReportFormat�Ϳ�����
            Call SaveReportFormat(mSigns("K" & lngKey), True, arrSQL)
            
            For i = 0 To UBound(arrSQL)
                If Trim(arrSQL(i)) <> "" Then
                    Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����ǩ��")
                End If
            Next i
            
'            OutputDebugString "ZLPACS>>AddSign:8 ǩ��������ɣ�ˢ��ǩ����Ϣ."
            
            'ˢ��ǩ������ȷ�������ݿ��һ��
            Call RefreshSigns
            
'            OutputDebugString "ZLPACS>>AddSign:9 ���±�����."
            
            Call UpdateReporter(mlngAdviceID, "")
            
            lngSaveType = IIf(OneSign.ǩ������ < cprSL_����, 1, 2)
            
            
        End If
        
'        OutputDebugString "ZLPACS>>AddSign:10 �������汣���¼�."
        
        '�����Ƿ�ȷ�Ͻ���ǩ������Ҫ�������汣����ɵ��¼�
        '�������汣����ɵ��¼�
        If blֱ�����ǩ�� Then lngSaveType = 6
                
        Call AfterReportSaved(mlngAdviceID, lngSaveType)
        
'        OutputDebugString "ZLPACS>>AddSign:11 ���汣���¼��������."
        
        '���ǩ���ɹ�������������ǩ�����˳�����ж�ر��洰��
        If Not OneSign Is Nothing And mblnExitAfterSign = True And mblnSingleWindow = True Then
            Call SetMenuDownState(False)
            Unload Me
        ElseIf mblnExitAfterSign = False Then
            Call zlRefreshFace(True)
        End If
        
'        OutputDebugString "ZLPACS>>AddSign:12 ҽ��IDΪ[" & mlngAdviceID & "]��ǩ���������."
        
        mblnIsSignSave = False
    Exit Sub
errHandle:
    mblnIsSignSave = False
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Private Sub RejectReport()
'���ر���
Dim frmRj As frmReject
Dim i As Long
Dim lngAdviceColIndex As Long
Dim lngProcedureColIndex As Long
Dim lngRowIndex As Long
    
On Error GoTo errFree
    If mReportID <= 0 Then
        MsgBoxD Me, "��ǰ���û�б��棬���ܽ��в��ز�����", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Set frmRj = New frmReject
    
    Call frmRj.ShowRejectWindow(mlngAdviceID, mReportID, Me)
    
    If frmRj.IsOk Then
        Call SendMsgToMainWindow(Me, wetRejectReport, mlngAdviceID)
    End If
errFree:
    Unload frmRj
    Set frmRj = Nothing
End Sub


Private Sub ShowRejectHistory()
'��ʾ������ʷ
Dim frmRj As frmReject
    
On Error GoTo errFree
    If mReportID <= 0 Then
        MsgBoxD Me, "��ǰ���û�б��棬�����ڲ�����ʷ��¼��", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Set frmRj = New frmReject
    
    Call frmRj.ShowRejectHistory(mlngAdviceID, mReportID, Me)
errFree:
    Unload frmRj
    Set frmRj = Nothing
End Sub


Private Sub ChangeFormat(lngFormatId As Long)
    mFormatID = lngFormatId
    mHasChangeFormat = True
    '���½�����ʾ
    If mblnShowImage = True Then
        mfrmReportImage.zlChangeFormat lngFormatId
    End If
    '���������� ������
    Call ShowTitle(True)
End Sub

Private Function HasReportImage(ByVal lngFileId As Long) As Boolean
'��ѯ�Ƿ��б���ͼ��
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    HasReportImage = False
    strSql = "select count(1) from �����ļ��ṹ where ��������=3 and substr(��������, instr(��������,';',1,18)+1, 1) = '2' And �ļ�ID=[1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "����ͼ���ж�", lngFileId)
    
    If rsData.RecordCount > 0 Then HasReportImage = True
End Function

Private Function SaveReportItems(blnCreate As Boolean, iAction As Integer) As Boolean
 ' iAction = 0   '�Ӳ����ļ��б�������, iType = 1    '�Ӳ�������Ŀ¼��������
    Dim arySql() As String
    Dim i As Long
    Dim blnInTrans As Boolean
    Dim blnImageSaveOk As Boolean
    
On Error GoTo errHandle

'    OutputDebugString "ZLPACS>>SaveReportItems:1 ��ʼִ�б�����Ŀ����..."
    
    SaveReportItems = False
    
    If blnCreate = True Then        '��������
        If CreateReport(iAction) = False Then
            Exit Function
        End If
    End If
    
    'TODO �������ﴦ��
    ReDim arySql(1)
    
'    OutputDebugString "ZLPACS>>SaveReportItems:2 ��ȡ��������ִ�����."
    '���汨������
    Call SaveReportView(arySql)
    
     '�����±��棬δ�������ͼ�������������⣬��ˢ��һ�Σ�ȷ������ͼ�ǵ�ǰ���ߵ�
    If mblnShowImage = True Then
        If mfrmReportImage.pImageModified = False And mfrmReportImage.pMarkModified = False And blnCreate = True Then
            mfrmReportImage.zlRefresh mlngAdviceID, mFileID, mReportID, mblnSingleWindow, _
                    mlngShowBigImg, mintImageDblClick, mblnEditable, mblnMoved, mintMinImageCount, _
                    True, mlngModule, mlngDeptID, mlngStudyState, IIf(blnCreate, False, mblnIsSignSave)
        End If
    End If
        
'    OutputDebugString "ZLPACS>>SaveReportItems:3 ��ȡ���ͼ����ִ�����."
    '������ͼ���
    Call SavePicMarks(blnCreate, arySql)
    
    blnImageSaveOk = True
    
    '���汨��ͼ
    If mblnShowImage = True Then
       If HasReportImage(mFileID) Then
'            OutputDebugString "ZLPACS>>SaveReportItems:4 ���ñ���ͼ���淽��."

            If mfrmReportImage.pImageModified = True Or blnCreate = True Then
                blnImageSaveOk = SaveReportImages(blnCreate, arySql)
            End If
            
'            OutputDebugString "ZLPACS>>SaveReportItems:5 ����ͼ���淽���������."
        End If
    End If
    
'    OutputDebugString "ZLPACS>>SaveReportItems:6 ���汨���ʽ."
    
    '���汨���ʽ,ǩ�����󴫿գ���ʾֻ�����ʽ��������ǩ��
    Call SaveReportFormat(Nothing, True, arySql)
    
'    OutputDebugString "ZLPACS>>SaveReportItems:7 �����ʽ�������."
    
    If blnImageSaveOk = False Then
        Call MsgBox("����ͼ���ϴ����ִ��󣬽��ݶԱ������ݽ��б��棬���Ժ�����ͼ����Ӳ����档", vbOKOnly, "����")
    End If
    
'    OutputDebugString "ZLPACS>>SaveReportItems:8 �ύ���ݵ�������."
    
    gcnOracle.BeginTrans        '----------���汨������
    blnInTrans = True
    For i = 0 To UBound(arySql)
        If Trim(arySql(i)) <> "" Then
            Call zlDatabase.ExecuteProcedure(CStr(arySql(i)), "���汨������")
        End If
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False

'    OutputDebugString "ZLPACS>>SaveReportItems:9 ������Ŀ�������."
    
    SaveReportItems = True
Exit Function
errHandle:
    SaveReportItems = False
    If blnInTrans Then gcnOracle.RollbackTrans
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Function

Private Function CreateReport(iType As Integer) As Boolean
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    CreateReport = False
    
    ' iType = 0   '�Ӳ����ļ��б�������, iType = 1    '�Ӳ�������Ŀ¼��������

    '�������Ӳ�������
    strSql = "ZL_Ӱ�񱨸�����_����(" & mlngAdviceID & "," & mFileID & "," & mFormatID & "," & iType & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    '�´����ı��棬�����ݿ��ж�ȡ��������ID
    strSql = "Select ����ID From ����ҽ������ Where ҽ��ID= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
    If rsTemp.EOF = True Then
        MsgBoxD Me, "������������ȷ���޷����ҵ���������ID"
        Exit Function
    Else
        mReportID = rsTemp!����Id
    End If
    CreateReport = True
    Exit Function
err:
    If ErrCenter() = 1 Then Resume Next
End Function

Private Function SaveReportImages(blnCreate As Boolean, ByRef arrSQL() As String) As Boolean
    Dim dblImgTableID  As Double
    Dim strTabIds() As String
    Dim iImgCount As Integer
    Dim strSql  As String
    Dim rsTemp As ADODB.Recordset
    Dim strTempFile As String
    Dim i As Integer
    Dim j As Integer
    Dim strPicAttrs As String
    Dim cFTP As New clsFtp
    Dim strFTPUser As String
    Dim strFTPPwd As String
    Dim strFtpIp As String
    Dim strFTPDirUrl As String
    Dim strSaveDeviceID As String
    Dim strBufferDir As String
    Dim strLocalDir As String
    Dim strBurFile As String
    Dim lngCheckResult As Long
    Dim lngResult As Long
    Dim strTabIdExs As String
    
    On Error GoTo errHandle
    
    SaveReportImages = True
    
'    OutputDebugString "ZLPACS>>SaveReportImages:1 ��ʼִ�б���ͼ����..."
    
    If mfrmReportImage.dcmReportImage.Count <= 1 Then Exit Function
    
    strTempFile = App.Path & "\Temp.jpg"
    
'    OutputDebugString "ZLPACS>>SaveReportImages:2 ��ʱ�ļ�����Ϊ:" & strTempFile
    
    '��ȡ���ID��
    
'    OutputDebugString "ZLPACS>>SaveReportImages:3 ��ȡ����ID��..."
    
    If blnCreate = True Then
        strSql = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
            " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By �������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        If rsTemp.RecordCount > 0 Then
            ReDim strTabIds(rsTemp.RecordCount - 1) As String
            For i = 0 To rsTemp.RecordCount - 1
                strTabIds(i) = rsTemp!���ID
                strTabIdExs = strTabIdExs & ";" & strTabIds(i)
                
                If i = 0 Then
                    mfrmReportImage.pTableID = rsTemp!���ID
                Else
                    mfrmReportImage.pTableID = mfrmReportImage.pTableID & ";" & rsTemp!���ID
                End If
                rsTemp.MoveNext
            Next i
        End If
    Else
        strTabIds = Split(mfrmReportImage.pTableID, ";")
    End If
    
'    OutputDebugString "ZLPACS>>SaveReportImages:4 TabID����ȡ��ɣ�����Ϊ��" & strTabIdExs
    
    '���ж������Ƿ�Ϊ��
    If SafeArrayGetDim(strTabIds) <> 0 Then
        '��ȡ���汨��ͼ��FTP��Ϣ
        strBufferDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
        
'        OutputDebugString "ZLPACS>>SaveReportImages:5 ������ʱ����Ŀ¼Ϊ��" & strBufferDir
        
'        OutputDebugString "ZLPACS>>SaveReportImages:6 ��ȡFTP��Ϣ."
        
        strSql = "Select λ��һ,λ�ö�,���UID,�������� From Ӱ�����¼ Where ���UID is not null And ҽ��ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngAdviceID)
        If rsTemp.RecordCount <> 0 Then
            strSaveDeviceID = Nvl(rsTemp!λ��һ)
            If strSaveDeviceID = "" Then
                strSaveDeviceID = Nvl(rsTemp!λ�ö�)
            End If
            strLocalDir = Format(Nvl(rsTemp!��������), "yyyyMMdd") & "/" & Nvl(rsTemp!���uid)
            
'            OutputDebugString "ZLPACS>>SaveReportImages:7 ����Ŀ¼Ϊ��" & strLocalDir
            
            Call funGetStorageDevice(Me, strSaveDeviceID, strFTPDirUrl, strFtpIp, strFTPUser, strFTPPwd)
            
'            OutputDebugString "ZLPACS>>SaveReportImages:8 ����FTP����."
            
            lngResult = cFTP.FuncFtpConnect(strFtpIp, strFTPUser, strFTPPwd)
            
'            OutputDebugString "ZLPACS>>SaveReportImages:9 FTP���Ӵ������,����ֵ��" & lngResult
        End If
        
        
'        OutputDebugString "ZLPACS>>SaveReportImages:10 ��ʼ�ϴ�����ͼ."
        
        '�ж�Ŀ¼�Ƿ����
        Call MkLocalDir(strBufferDir & "" & strLocalDir & "\")
        
        '�����ͱ���ÿһ��ͼ����
        For i = 0 To UBound(strTabIds)
            dblImgTableID = Val(strTabIds(i))
            iImgCount = mfrmReportImage.dcmReportImage(i + 1).Images.Count
            strPicAttrs = ""
            
            For j = 1 To iImgCount
                strPicAttrs = strPicAttrs & ";" & mfrmReportImage.dcmReportImage(i + 1).Images(j).tag & "," & mlngAdviceID
            Next j
            
'            OutputDebugString "ZLPACS>>SaveReportImages:11 dblImgTableID:" & lngImgTableID & " ����ͼ��������:" & strPicAttrs
            
            strSql = "ZL_Ӱ�񱨸�ͼ��_����(" & dblImgTableID & ",'" & strPicAttrs & "')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSql
    
'            zlDatabase.ExecuteProcedure strSql, Me.Caption
            
'            OutputDebugString "ZLPACS>>SaveReportImages:12 ��ʼ��ѵ����ͼ���ϴ�."
            
            '���汨��ͼ�ļ���FTPĿ¼��
            For j = 1 To mfrmReportImage.dcmReportImage(i + 1).Images.Count
                strBurFile = strBufferDir & "" & strLocalDir & "\" & mfrmReportImage.dcmReportImage(i + 1).Images(j).tag
                strBurFile = Replace(strBurFile, "/", "\")
                
'                OutputDebugString "ZLPACS>>SaveReportImages:13 ����ͼ��ǰ���:" & j & " ����λ��:" & strBurFile
'                OutputDebugString "ZLPACS>>SaveReportImages:14 ��ʼ����JPG�ļ�."
                
                mfrmReportImage.dcmReportImage(i + 1).Images(j).FileExport strBurFile, "JPG"
                
                If FileExists(strBurFile) = True Then
'                    OutputDebugString "ZLPACS>>SaveReportImages:15 �ļ���������,�ļ�Ϊ��" & strBurFile
                Else
'                    OutputDebugString "ZLPACS>>SaveReportImages:16 �ļ�����ʧ��."
                End If
reLoad:
'                OutputDebugString "ZLPACS>>SaveReportImages:17 ��ʼ�ϴ��ļ�."
                
                lngResult = cFTP.FuncUploadFile(strFTPDirUrl & strLocalDir & "/", _
                                        strBurFile, _
                                        mfrmReportImage.dcmReportImage(i + 1).Images(j).tag)
                                        
'                OutputDebugString "ZLPACS>>SaveReportImages:18 �ļ��ϴ���ɣ�����ֵ:" & lngResult
'                OutputDebugString "ZLPACS>>SaveReportImages:19 ��ʼ�ļ�һ���Լ��."
                If mblnCompareSize Then
                    lngCheckResult = ChechReportImgAndReload(cFTP, strBurFile, strFTPDirUrl & strLocalDir & "/", mfrmReportImage.dcmReportImage(i + 1).Images(j).tag)
                    
                    If lngCheckResult = 2 Then  'Ϊ2��ʾ����
                        GoTo reLoad
                    ElseIf lngCheckResult = 1 Then  'Ϊ1��ʾͼ��ʧ�ܲ�����
                        SaveReportImages = False
                    End If
                End If
                
'                OutputDebugString "ZLPACS>>SaveReportImages:20 �ļ�һ���Լ�����."
            Next j
            
        Next i
    End If
    
'    OutputDebugString "ZLPACS>>SaveReportImages:21 �Ͽ�Ftp����."
    
    cFTP.FuncFtpDisConnect
    
    mfrmReportImage.pImageModified = False
    
'    OutputDebugString "ZLPACS>>SaveReportImages:22 ����ͼ�������."
    Exit Function
errHandle:
    cFTP.FuncFtpDisConnect
    
    If ErrCenter() = 1 Then Resume Next
End Function

Private Function ChechReportImgAndReload(cFTP As clsFtp, ByVal strSrcFile As String, strFtpFilePath As String, ByVal strFileName As String) As Long
'����ϴ����ļ��ͱ����ļ���С�Ƿ�һ�£���һ���򷵻�true
    Dim blnResult As Boolean, blnReUpload As Boolean
    Dim lngFtpFileSzie As Long, lngDestFileSize As Long
    Dim strMessage As String
    Dim objFileSystem As New FileSystemObject
    
On Error GoTo errHandle

    ChechReportImgAndReload = 0
    
    '�ϴ���Ա�һ�´�С���ж��Ƿ������ϴ�
    lngDestFileSize = objFileSystem.GetFile(strSrcFile).Size
    lngFtpFileSzie = cFTP.FuncFtpGetFileSize(strFtpFilePath, strFileName)

    If lngFtpFileSzie < lngDestFileSize Then
        strMessage = "�ϴ�����ļ���С[" & lngFtpFileSzie & "]��ԭ�ļ���С[" & lngDestFileSize & "]��һ��" & vbCrLf & _
                     "ԭ�ļ���" & strSrcFile & vbCrLf & _
                     "FTP�ļ���" & strFtpFilePath & strFileName & vbCrLf & _
                     "�Ƿ���Ҫ�����ϴ���"
        
        If MsgBox(strMessage, vbQuestion + vbYesNo, "��ʾ") = vbYes Then
            ChechReportImgAndReload = 2 'Ϊ2��ʾ����
        Else
            ChechReportImgAndReload = 1 'Ϊ1��ʾʧ�ܲ�����
        End If
    End If
Exit Function
errHandle:
    If MsgBox("ͼ���ļ�[����:" & strSrcFile & "  FTP:" & strFtpFilePath & "/" & strFileName & "]һ���Լ�����,����ԭ��:" & err.Description & "��" & vbCrLf & "�Ƿ����ԣ�", vbQuestion + vbYesNo, "��ʾ") = vbYes Then
        ChechReportImgAndReload = 2
    Else
        ChechReportImgAndReload = 1
    End If
End Function

Private Sub SavePicMarks(blnCreate As Boolean, ByRef arrSQL() As String)
    Dim strMarks As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim dblMarkImageID As Double
    Dim i As Integer
    
    If mfrmReportImage Is Nothing Then Exit Sub
    
    If mfrmReportImage.pobjMarks Is Nothing Then
        mfrmReportImage.pMarkModified = False
        Exit Sub
    End If
    '��������ı�
    For i = 1 To mfrmReportImage.pobjMarks.Count
        If i = 1 Then
            strMarks = mfrmReportImage.pobjMarks(i).��������
        Else
            strMarks = strMarks & "||" & mfrmReportImage.pobjMarks(i).��������
        End If
    Next i

    If blnCreate = True Then
        '�´����ı��棬�ӵ��Ӳ��������ж�ȡ���ͼID
        strSql = "Select Id From ���Ӳ������� Where �ļ�ID=[1] And  ��������= 5 And substr(��������,1,1)='1' "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        If rsTemp.EOF = False Then  '�б��ͼ
            dblMarkImageID = Val(rsTemp!ID)
        Else    'û�б��ͼ
            dblMarkImageID = 0
        End If
        
        mfrmReportImage.pMarkImageID = dblMarkImageID
    Else
        dblMarkImageID = mfrmReportImage.pMarkImageID
    End If
    
    strSql = "ZL_Ӱ�񱨸��ע_����(" & dblMarkImageID & ",'" & strMarks & "')"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSql
        
'    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    mfrmReportImage.pMarkModified = False
End Sub

Private Sub SaveReportFormat(OneSign As cEPRSign, blnAddSign As Boolean, ByRef arrSQL() As String)
'------------------------------------------------
'���ܣ����汨���ʽRTF�ļ����Ա������ǩ�����߻���
'������     OneSign -- ��Ϊ�գ����ʾ����ǩ�����߻��ˣ�Ϊ�գ���ʾֻ�Ǳ����ʽ��������ǩ��
'           blnAddSign ���ӻ��߻���ǩ����True--����ǩ��,OneSignΪ�ձ�ʾ���汨���ʽ��False--����ǩ��
'���أ� �ޣ�ֱ�ӱ���RTF�����ʽ�ĵ����Ա���ǩ�����߻���
'-----------------------------------------------
    Dim strZipFile As String
    Dim strTemp As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strReport As String
    Dim lngSignPos As Long
    Dim strReportFormatFile As String
    Dim strErrCount As String
    
    strErrCount = ""
    
reLoad:
    strReportFormatFile = App.Path & "\ReportTemp" & strErrCount
    
    '�ȸ��Ʊ����ʽ
    If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
    
    '�����ݿ��ȡRTF�����ʽ�ĵ�
    strZipFile = zlBlobRead(5, mReportID, strReportFormatFile)
    
    '��ѹ���ļ�
    strTemp = zlFileUnzip(strZipFile)
    
    If strTemp <> "" Then
        If blnAddSign = True Then
            '�����ļ������ݱ������ݣ��޸�����Ҫ������
            '��ȡRTF�ļ�����
            rtxtSaveElement.Filename = strTemp
            strReport = rtxtSaveElement.TextRTF
            
            '��ȡ���ݿ��е�Ҫ�أ��Ѹ���Ҫ��������д����ʽ��
            strSql = "Select ������,�����ı�,Ҫ������ From ���Ӳ������� Where �ļ�ID= [1] And �������� = 4 And ��ֹ��=0 and �������� =0 order by ������ "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
            While (rsTemp.EOF = False)
                ReplaceElement strReport, "E", rsTemp!������, Nvl(rsTemp!�����ı�, " ")
                rsTemp.MoveNext
            Wend
            
            '����RTF�ļ�
            rtxtSaveElement.TextRTF = strReport
            rtxtSaveElement.SaveFile strTemp
        End If
        
        '�����ǩ�����򱣴�ǩ��
        If Not OneSign Is Nothing Then
            edtEditor.OpenDoc strTemp
            If blnAddSign = True Then   '����ǩ��
                '����д��ǩ����λ��
                strSql = "Select ������ From ���Ӳ������� Where �ļ�ID= [1] And �������� = 4 And Ҫ������ ='����ǩ��' "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
                lngSignPos = -1
                If rsTemp.EOF = False Then
                    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    
                    bFinded = FindKey(edtEditor, "E", Nvl(rsTemp!������, 0), lKSS, lKSE, lKES, lKEE, bNeeded)
                    If bFinded = True Then lngSignPos = lKEE
                End If
                
                '��ָ��λ��д��ǩ��
                OneSign.InsertIntoEditor edtEditor, lngSignPos
                
                '��ǩ�����浽���ݿ�
                strSql = "ZL_Ӱ�񱨸�ǩ��_����(" & mReportID & "," & OneSign.��ʼ�� & "," & OneSign.��ֹ�� & " ,'" & OneSign.�������� & "','" & OneSign.���� & _
                        "','" & OneSign.ǰ������ & "','" & OneSign.ʱ��� & "'," & OneSign.ǩ������ & ",'" & OneSign.ǩ����Ϣ & "')"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSql
                
                'zlDatabase.ExecuteProcedure strSql, Me.Caption
            Else    '����ǩ��
                OneSign.DeleteFromEditor edtEditor
                
                '�ѻ���ǩ�����浽���ݿ�
                strSql = "ZL_Ӱ�񱨸����(" & OneSign.ID & "," & mReportID & ",0)"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSql
                
                'zlDatabase.ExecuteProcedure strSql, Me.Caption
            End If
            
            '�������ʱ�ļ�
            edtEditor.SaveDoc strTemp
        End If
        
        'ѹ���ļ�
        strZipFile = zlFileZip(strTemp)
        
        '�����ʽ
        zlBlobSave 5, mReportID, strZipFile, arrSQL
    
        'ɾ����ʱzip�ļ�
        Kill strZipFile
    Else
        If MsgBoxD(Me, "�޷���ȡ���߽�ѹ�����ʽ" & strReportFormatFile & vbCrLf & "��ʹ�á������༭���ķ������༭�˱�������Զ�ȡ���Ƿ����ԣ�", vbYesNo) = vbYes Then
            If Dir(strReportFormatFile) <> "" Then Kill strReportFormatFile
            
            strErrCount = CStr(Val(strErrCount) + 1)
            GoTo reLoad
        End If
    End If
End Sub

Private Function CheckSignRollbackType(ByVal dblID As Double, lngReportID As Long) As Integer
'������ǩ������ lngID�����Ӳ�������.ID  ��lngReportID�����Ӳ�������.�ļ�ID
'���� 0������  1���ǩ��  2/3���ǩ��
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH

    CheckSignRollbackType = 0
    strSql = "Select Ҫ�ر�ʾ From ���Ӳ�������  where ID=[1] and �ļ�ID =[2]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dblID, lngReportID)
    If rsTemp.RecordCount = 1 Then
        CheckSignRollbackType = Nvl(rsTemp!Ҫ�ر�ʾ)
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReplaceElement(strReport As String, strKeyType As String, lngKey As Long, strElement As String) As Boolean
    Dim sTMP As String
    Dim i As Long
    Dim j As Long
    Dim lLength As Long
    Dim strNewReport As String
    Dim lngES As Long
    Dim lngEE As Long
    Dim strChar As String
    Dim lulWave As Long
    Dim lulNone As Long
    
    sTMP = strKeyType & "S(" & Format(lngKey, "00000000")
    i = 1
LL1:
    i = InStr(i, strReport, sTMP)
    If i <> 0 Then
        '���Ƿ�ؼ��֣���Ϊ�ؼ��֣��������������ܱ����ġ�
        If ProtectAndHide(strReport, i - 1, i) = False Then
            i = i + 1
            GoTo LL1
        End If
        '�Ѿ��ҵ���ʼ�ؼ��֣���������ַ������滻��Щ�ַ�
        j = i + 16
        lngES = j
        '���ҽ����ؼ���
        sTMP = strKeyType & "E(" & Format(lngKey, "00000000")
LL2:
        j = InStr(j, strReport, sTMP)
        If j <> 0 Then
            '���Ƿ�ؼ��֣���Ϊ�ؼ��֣��������������ܱ����ġ�
            If ProtectAndHide(strReport, j - 1, j) = False Then
                j = j + 1
                GoTo LL2
            End If
            lngEE = j - 1
            '�Ѿ��ҵ������ؼ��֣�˵���м������Ҫ�滻��Ҫ��
            
            '���˵����Ʒ��ţ�\cfN,\highlightN,\v0
            If getElementPos(strReport, lngES, lLength, lngEE, lulWave, lulNone) = True Then
                strNewReport = strReport
                '�ȴ����»����ˣ�ɾ���»����˵��������
                If lulWave <> 0 And lulNone <> 0 Then
                    strNewReport = Left(strNewReport, lulNone) & Right(strNewReport, Len(strNewReport) - lulNone - 7)
                End If
                '�ٴ���Ҫ�����ݣ��滻Ҫ������
                strChar = Mid(strElement, 1, 1)
                If (strChar >= "A" And strChar <= "Z") Or (strChar >= "a" And strChar <= "z") Or IsNumeric(strChar) Or strChar = " " Then
                    strNewReport = Left(strNewReport, lngES) & " " & StrToASC(strElement) & Right(strNewReport, Len(strNewReport) - lngES - lLength)
                Else
                    strNewReport = Left(strNewReport, lngES) & StrToASC(strElement) & Right(strNewReport, Len(strNewReport) - lngES - lLength)
                End If
                If lulWave <> 0 And lulNone <> 0 Then
                    strNewReport = Left(strNewReport, lulWave) & Right(strNewReport, Len(strNewReport) - lulWave - 7)
                End If
                strReport = strNewReport
                ReplaceElement = True
            End If
        End If
    End If
End Function

Private Function StrToASC(ByVal strIn As String) As String
    '�������ַ���ת��ΪASC��������Ӣ��һ��
    '�Ƚ������ַ�����ת�壺
    strIn = Replace(strIn, Chr(9), "\TAB ")
    strIn = Replace(strIn, Chr(13) + Chr(10), "\par ")
    Dim i As Long, s As String, lsChar As String, lsPart1 As String, lsPart2 As String
    Dim lsCharHex As String
    For i = 1 To Len(strIn)
        lsChar = Mid(strIn, i, 1)
        If lsChar = "?" Then
            lsCharHex = LCase(Hex(Asc(lsChar)))
            If Len(lsCharHex) = 4 Then
                lsCharHex = "\'" + Mid(lsCharHex, 1, 2) + "\'" + Mid(lsCharHex, 3, 2)
            Else
                lsCharHex = lsChar
            End If
            s = s + lsCharHex
        Else
            lsCharHex = LCase(Hex(Asc(lsChar)))
            If Len(lsCharHex) = 4 Then
                lsCharHex = "\'" + Mid(lsCharHex, 1, 2) + "\'" + Mid(lsCharHex, 3, 2)
            Else
                lsCharHex = lsChar
            End If
            s = s + lsCharHex
        End If
    Next
    StrToASC = s
End Function

Private Function getElementPos(ByVal strReport As String, ByRef lStart As Long, ByRef lLength As Long, _
    ByVal lEnd As Long, ByRef lulWave As Long, ByRef lulNone As Long) As Boolean
'    lulWave   '�»����˱��\ulwave�Ŀ�ʼλ��
'    lulNone    '�ر������»��߱��\ulnone�Ŀ�ʼλ��
    '���Ҵ�lStart��ʼ�ģ�Ԫ�������ı��Ŀ�ʼλ�úͳ���
    '���ҺͶ�λԪ���е��»����˱��\ulwave �� �ر������»��߱��\ulnone
    Dim lIndex As Long
    Dim lWordEnd As Long
    Dim blnSearch As Boolean
    Dim strChar As String
    Dim strNextChar As String
    Dim blnInWord As Boolean
    Dim strTemp As String
    
    lIndex = lStart
    blnSearch = True
    blnInWord = True
    
    While (blnSearch And lIndex < lEnd)
        strChar = Mid(strReport, lIndex, 1)
        If strChar = "\" Then       '��һ�������ַ���������һ�������ַ����������ı��Ŀ�ʼ
            strNextChar = Mid(strReport, lIndex + 1, 1)
            If strNextChar = "'" Or strNextChar = "{" Or strNextChar = "}" Or strNextChar = "\" Then     '�ı��Ŀ�ʼ
                '�����ҵ�һ�����Ʒ�
                blnInWord = True
                lStart = lIndex - 1
                While (blnInWord And lIndex <= lEnd)
                    lIndex = lIndex + 1
                    strChar = Mid(strReport, lIndex, 1)
                    If strChar = "\" Then
                        strNextChar = Mid(strReport, lIndex + 1, 1)
                        If strNextChar = "'" Or strNextChar = "{" Or strNextChar = "}" Or strNextChar = "\" Then
                            lIndex = lIndex + 1
                        Else
                            lWordEnd = lIndex - 1
                            blnInWord = False   '�˳�����ѭ��
                        End If
                    End If
                Wend
            Else    '�����ַ��Ŀ�ʼ
                '�����ȡһֱ�������ַ�����
                strTemp = Mid(strReport, lIndex, 1)
                lIndex = lIndex + 1
                While (Mid(strReport, lIndex, 1) <> "\" And Mid(strReport, lIndex, 1) <> " ")
                    strTemp = strTemp & Mid(strReport, lIndex, 1)
                    lIndex = lIndex + 1
                Wend
                If strTemp = "\ulwave" Then
                    lulWave = lIndex - 8
                ElseIf strTemp = "\ulnone" Then
                    lulNone = lIndex - 8
                    blnSearch = False   '�˳�����Ԫ�ص�ѭ��
                End If
            End If
        ElseIf strChar = " " Then   '���Ŀ�ʼ���������ĵ��ַ���Ӣ�ģ���������
            '�����ҵ�һ�����Ʒ�
            blnInWord = True
            lStart = lIndex - 1
            While (blnInWord And lIndex <= lEnd)
                lIndex = lIndex + 1
                strChar = Mid(strReport, lIndex, 1)
                If strChar = "\" Then
                    strNextChar = Mid(strReport, lIndex + 1, 1)
                    If strNextChar = "'" Or strNextChar = "{" Or strNextChar = "}" Or strNextChar = "\" Then
                        lIndex = lIndex + 1
                    Else
                        lWordEnd = lIndex - 1
                        blnInWord = False   '�˳�����ѭ��
                    End If
                End If
            Wend
            
        Else        '�ڲ�����ȷ��RTF�ļ������ز��Ҵ���
            getElementPos = False
            Exit Function
        End If
    Wend
    lLength = lWordEnd - lStart
    If lWordEnd = 0 Then  '˵���ǲ鵽Ҫ�ؽ����ˣ����˳��ģ�û�в��ҵ������ı�
        getElementPos = False
    Else
        getElementPos = True
    End If
End Function


Private Function ProtectAndHide(ByRef strReport As String, ByVal lStart As Long, ByVal lEnd As Long) As Boolean
    Dim lOnPos As Long
    Dim lOffPos As Long
    
    '��ǰ�����غͱ�����ʼ��ǣ�\v��\protect
    lOnPos = InStrRev(strReport, "\v", lStart, vbTextCompare)
    lOffPos = InStrRev(strReport, "\v0", lStart, vbTextCompare)
    If lOnPos > lOffPos And lOnPos <> 0 Then
        '���Һ�������ر��
        lOnPos = InStr(lEnd, strReport, "\v", vbTextCompare)
        lOffPos = InStr(lEnd, strReport, "\v0", vbTextCompare)
        If lOffPos <= lOnPos And lOffPos <> 0 Then
            '����ǰ��ı������
            lOnPos = InStrRev(strReport, "\protect", lStart, vbTextCompare)
            lOffPos = InStrRev(strReport, "\protect0", lStart, vbTextCompare)
            If lOnPos > lOffPos And lOnPos <> 0 Then
                '���Һ���ı������
                lOnPos = InStr(lEnd, strReport, "\protect", vbTextCompare)
                lOffPos = InStr(lEnd, strReport, "\protect0", vbTextCompare)
                If lOffPos <= lOnPos And lOffPos <> 0 Then
                    ProtectAndHide = True
                End If
            End If
        End If
    End If
End Function


Public Sub SaveReportView(ByRef arrSQL() As String)
    Dim strReport As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strElements As String
    'Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim intLevel As Integer 'ǩ������
    Dim strSQLLevel As String 'ǩ����ѯ
    Dim rsTempLevel As ADODB.Recordset 'ǩ����ѯ���
    Dim strUnitName As String
    
    
    On Error GoTo errHandle
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    
    
    '�޸ı���ǩ��Ҫ�أ����������滻Ϊ�� ��
    strElements = SPLITER_REPORT & Report_Element_����ǩ�� & SPLITER_ELEMENT & " "
    '��֯ר�Ʊ�������
    If mblnShowSpecial = True Then
        strElements = strElements & mfrmReportSpecial.getElementString
    End If
    '��֯���ı��εĶ�������,���TagΪ�գ�������ݿ��ȡĬ��ֵ
    If mfrmReportView.rtxtCheckView.tag = "" Or mfrmReportView.rtxtResult.tag = "" Or mfrmReportView.rTxtAdvice.tag = "" Then
        strSql = "Select a.�����ı� As ����, b.�������� From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 And b.��ֹ�� = 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mReportID)
        While rsTemp.EOF = False
            Select Case rsTemp!����
                Case "�������"
                    If mfrmReportView.rtxtCheckView.tag = "" Then
                        mfrmReportView.rtxtCheckView.tag = rsTemp!��������
                        RefreshViewTag mfrmReportView.rtxtCheckView
                    End If
                Case "������"
                    If mfrmReportView.rtxtResult.tag = "" Then
                        mfrmReportView.rtxtResult.tag = rsTemp!��������
                        RefreshViewTag mfrmReportView.rtxtResult
                    End If
                Case "����"
                    If mfrmReportView.rTxtAdvice.tag = "" Then
                        mfrmReportView.rTxtAdvice.tag = rsTemp!��������
                        RefreshViewTag mfrmReportView.rTxtAdvice
                    End If
            End Select
            rsTemp.MoveNext
        Wend
        Else
        RefreshViewTag mfrmReportView.rtxtCheckView
        RefreshViewTag mfrmReportView.rtxtResult
        RefreshViewTag mfrmReportView.rTxtAdvice
    End If
    
        
    '��󱣴���ı������ݣ���ʱ��������ݿ����ݣ��Զ����±����е�Ҫ��
    strReport = SPLITER_REPORT & "1" & mfrmReportView.rtxtCheckView.tag & SPLITER_ELEMENT & mfrmReportView.rtxtCheckView.Text & SPLITER_REPORT _
        & "2" & mfrmReportView.rtxtResult.tag & SPLITER_ELEMENT & mfrmReportView.rtxtResult.Text & SPLITER_REPORT _
        & "3" & mfrmReportView.rTxtAdvice.tag & SPLITER_ELEMENT & mfrmReportView.rTxtAdvice.Text
    
    '����ţ�80185
    'ʹ���������ǩ������
    '�������ݵ�ʱ�򣬱����ǩ������ʼ����0���������ǩ������ͨ��ǩ���Ĺ���������
    
    
    strSQLLevel = " Select a.ҽ��id,a.����id,b.ǩ������ " _
             & "  From ����ҽ������ a, ���Ӳ�����¼ b Where a.ҽ��id = [1] And a.����id = b.Id "
    Set rsTempLevel = zlDatabase.OpenSQLRecord(strSQLLevel, "��ȡ�Ƿ�ǩ��", CLng(mlngAdviceID))
    If rsTempLevel.EOF = True Then
        intLevel = 0
    Else
        intLevel = Nvl(rsTempLevel!ǩ������)
    End If
    
    strUnitName = zlRegInfo("��λ����")
    
    strSql = "ZL_Ӱ�񱨸�����_update(" & mlngAdviceID & "," & mReportID & ",'" & Replace(strReport, "'", "��") & " ','" & strElements & "'," & mĿ��汾 & "," & intLevel & ",'" & strUnitName & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSql
    
    '�����Ȩ�޵ģ����ܱ����������
    If CheckPopedom(mstrPrivs, "���") Then
        strSql = "Zl_Ӱ�����_Update(" & mlngAdviceID & ",'" & Replace(mfrmReportView.txtReview.Text, "'", "��") & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSql
    End If
    
'    gcnOracle.BeginTrans        '----------���汨������
'    blnInTrans = True
'    For i = 0 To UBound(arrSQL)
'        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "���汨������")
'    Next i
'    gcnOracle.CommitTrans
'    blnInTrans = False
    
    mfrmReportView.pModified = False
    If mblnShowSpecial = True Then
        mfrmReportSpecial.pModified = False
    End If
    
    Exit Sub
errHandle:
'    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
        Call SaveErrLog
End Sub

Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim cbrControlItem As CommandBarControl
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    '��Ӹ�ʽѡ�񵯳��˵�
    If CommandBar.Parent.ID = conMenu_PacsReport_SelFormat Then
        CommandBar.Controls.DeleteAll
        
        '����µĲ˵���
        For i = 1 To UBound(rptFormats)
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_SelFormat_Item, rptFormats(i).strName, i)
            cbrControlItem.Parameter = rptFormats(i).ID
        Next i
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_RepFormat Then
        If mblnRefreshRptFormat = True And mblnʹ���Զ��屨�� = True Then
            CommandBar.Controls.DeleteAll
        
            '����µĲ˵���
            strSql = "Select a.���,b.���,b.˵�� From zlreports a,zlrptfmts b Where a.Id=b.����ID And a.���=[1] Order By ���"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Զ��屨���ʽ", mstr������)
            
            While rsTemp.EOF = False
                Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_RepFormat_Item, rsTemp!��� & "-" & Nvl(rsTemp!˵��))
                cbrControlItem.Style = xtpButtonIconAndCaption
                cbrControlItem.Checked = (InStr(mstrѡ�б����ʽ, cbrControlItem.Caption) <> 0)
                cbrControlItem.Parameter = rsTemp!���
                cbrControlItem.CloseSubMenuOnClick = False
            
                rsTemp.MoveNext
            Wend
            
            '�ر�ˢ��
            mblnRefreshRptFormat = False
        End If
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_VerifySign Then
        'ǩ����֤�ĵ����˵����г�������֤��ǩ���汾
        CommandBar.Controls.DeleteAll
        
        '����µ�ǩ����֤�˵�
        strSql = "Select ��ʼ��,�����ı� as ǩ��ҽ�� From ���Ӳ������� Where �ļ�ID = [1] And �������� =8  Order By ��ʼ��"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ǩ���汾", mReportID)
        
        While rsTemp.EOF = False
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_VerifySign_Item, rsTemp!��ʼ�� & "-" & Nvl(rsTemp!ǩ��ҽ��))
            cbrControlItem.Style = xtpButtonIconAndCaption
            cbrControlItem.Checked = False
            cbrControlItem.Parameter = rsTemp!��ʼ��
            rsTemp.MoveNext
        Wend
    End If
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)

    Select Case control.ID
        Case conMenu_File_Print, conMenu_File_Preview      '��ӡ����,Ԥ������
            control.Visible = CheckPopedom(mstrPrivs, "PACS�����ӡ")

            '���δ�ҵ���Ӧ�Ĳ����ļ�����ô��ӡԤ����ť�ᱻ����
            If mblnPrintView = True Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
            
            If control.Enabled Then control.Enabled = mReportID <> 0
'        Case comMenu_Petition_Capture
'            '��ȡ�������뵥ɨ�����
'            mblnIsPetitionScan = IIf(Val(GetDeptPara(mlngDeptId, "�������뵥ɨ��", 1)) = 1, True, False)
'            If mblnIsPetitionScan Then
'                control.Visible = True
'            Else
'                control.Visible = False
'            End If
            
        Case conMenu_Edit_Modify        '����༭
            '�ɼ���Visible�����������һ����ֻ����û��״̬����Enable��ֻҪ�ɼ��Ϳ��Բ���
            '�ڱ�����д״̬�£��б�����дȨ�޵��ˣ�������д�Լ��ı��棬�����˱���Ȩ�޵��ˣ�������д�����ұ��˵ı���
            If mĿ��汾 = 1 And CheckPopedom(mstrPrivs, "PACS������д") Then
                If mstrEPR������ = UserInfo.���� Then
                    control.Visible = True
                ElseIf (CheckPopedom(mstrPrivs, "PACS���˱���") And mlngEPRDeptID = mlngDeptID) Then '�����˱���Ȩ�޵ģ�������д�����ҵı���
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            ElseIf mĿ��汾 > 1 And CheckPopedom(mstrPrivs, "PACS�����޶�") Then
                '�ڱ����޶���״̬�£��б����޶�Ȩ�޵��ˣ�������д�����ҵı��档
                control.Visible = True
            Else
                control.Visible = False
            End If
            
            control.Enabled = mblnEditable
            
        Case conMenu_PacsReport_Reject
            '�ж��Ƿ�߱����沵��Ȩ��
            '�жϵ�ǰ��������״̬�Ƿ�������
            If Not CheckPopedom(mstrPrivs, "���沵��") Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = mReportID <> 0 And Not mblnReadOnly
            End If
        Case conMenu_PacsReport_RejectHistory
            control.Visible = Not CheckPopedom(mstrPrivs, "���沵��")
            
        Case conMenu_PacsReport_Save    '����
            '�ڱ�����д״̬�£��б�����дȨ�޵��ˣ�������д�Լ��ı��棬�����˱���Ȩ�޵��ˣ�������д�����ұ��˵ı���
            If mĿ��汾 = 1 And CheckPopedom(mstrPrivs, "PACS������д") Then
                If mstrEPR������ = UserInfo.���� Then
                    control.Visible = True
                ElseIf (CheckPopedom(mstrPrivs, "PACS���˱���") And mlngEPRDeptID = mlngDeptID) Then  '�����˱���Ȩ�޵ģ�������д�����ҵı���
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            ElseIf mĿ��汾 > 1 And CheckPopedom(mstrPrivs, "PACS�����޶�") Then
                '�ڱ����޶���״̬�£��б����޶�Ȩ�޵��ˣ�������д�����ҵı��档
                control.Visible = True
            Else
                control.Visible = False
            End If
            
            '���ݱ����״̬��ȷ���Ƿ�����д��Enable
            If control.Visible = True Then
                If mblnReadOnly = True Then
                    control.Enabled = False
                ElseIf mblnModified = False Then
                    mblnModified = chkModified
                    
                    If mblnModified = True Then
                        If mdtReportTime <> GetReportLastSaveTime(mlngAdviceID) Then
                            mblnModified = False
                            
                            Call zlUpdateAdviceInf(mlngAdviceID, mlngSendNo, mlngStudyState, mblnMoved)
                            mblnIsSignSave = True
                            Call zlRefreshFace(True)
                            mblnIsSignSave = False
                        Else
                            control.Enabled = True
                                                
                            '�ӷǱ༭ģʽ������༭ģʽ����������༭�¼�
                            RaiseEvent BeforeEdit(mlngAdviceID)
                    
                            tmrCheckingReportState.Enabled = True
                        End If
                    Else
                        control.Enabled = False
                    End If
                Else
                    control.Enabled = True
                End If
            End If
            
        Case conMenu_PacsReport_Sign    'ǩ��
            
            '����дģʽ�£���û��ǩ���ģ�����ǩ��
            '���޶�ģʽ�£�ǩ������û�г���16�εģ�����ǩ����
            'ֻ��ģʽ�£�ʲô�����ܲ�����
            If mĿ��汾 = 1 And CheckPopedom(mstrPrivs, "PACS������д") Then     '��û��ǩ��,��������дȨ��
                If mstrEPR������ = UserInfo.���� Then '�Լ�д�ı��棬�Լ�ǩ��
                    control.Visible = True
                ElseIf (CheckPopedom(mstrPrivs, "PACS���˱���") And mlngEPRDeptID = mlngDeptID) Then     '�����˱���Ȩ�޵ģ����Ը������ҵı���ǩ��
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            ElseIf mĿ��汾 > 1 And CheckPopedom(mstrPrivs, "PACS�����޶�") Then     '�Ѿ���ǩ���ˣ��ٴ�ǩ��������Ҫ�޶���Ȩ��
                control.Visible = (mĿ��汾 <= 16)
            Else
                control.Visible = False
            End If
            If control.Visible = True Then control.Enabled = Not mblnReadOnly
            
        Case conMenu_PacsReport_VerifySign  'ǩ����֤
            'ֻ������������ǩ��������ʾǩ����֤��ť
            'ֻ�б�����д�������޶�Ȩ�޵��ˣ����ܶ�ǩ��������֤
            control.Visible = IIf(mlngPassType = 0, False, True)
            
            If control.Visible = True Then
                If mĿ��汾 > 1 And (CheckPopedom(mstrPrivs, "PACS�����޶�") Or CheckPopedom(mstrPrivs, "PACS������д")) Then
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            End If
        Case conMenu_PacsReport_DelSign '����
            
            'û��ǩ��֮ǰ�������Ի���,ֻ�ܻ����Լ���ǩ��������ͨ������������ǩ������Ȩ�ޣ����˱����������˵�ǩ��
            If mĿ��汾 > 1 And mSigns.Count > 0 Then  'ֻ��ǩ������ſ��Ի���
                If mSigns("K" & mSigns.GetMaxKey).���� = UserInfo.���� And mǩ������ <> cprSL_�հ� Then  '�����Լ���ǩ��
                    control.Visible = True
                ElseIf mstrEPR������ = UserInfo.���� And mǩ������ = cprSL_�հ� Then   '�����Լ����޶�
                    control.Visible = True
                ElseIf CheckPopedom(mstrPrivs, "PACS���˱���") And mlngEPRDeptID = mlngDeptID Then      '�����˱���Ȩ�޵�,���Ի��˱����ҵ�����ǩ��
                    control.Visible = True
                Else
                    control.Visible = False
                End If
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                control.Enabled = (Not mblnReadOnly) And mblnCanUntread
            End If
            
        Case conMenu_PacsReport_SelFormat  'ѡ���ʽ '�޶�ģʽ�£����������ø�ʽ
            If Not CheckPopedom(mstrPrivs, "PACS������д") Then
                control.Visible = False
            Else
                control.Visible = IIf(mĿ��汾 = 1, True, False)
            End If
        Case conMenu_PacsReport_RepFormat   'ѡ���ӡ��ʽ
            control.Visible = mblnʹ���Զ��屨��
        Case conMenu_PacsReport_RepFormat_Item  'ѡ������ӡ��ʽ
            control.Checked = InStr(mstrѡ�б����ʽ, control.Caption)
            control.IconId = IIf(control.Checked, 90002, 90001)
        Case conMenu_PacsReport_FontSet                     '�����ֺ�
            
        Case conMenu_PacsReport_FontSetDefault To conMenu_PacsReport_FontSetUser   '�����ֺ�
            control.Checked = False
            If Val(control.Caption) = getMenuFontSize Then control.Checked = True
            
        Case conMenu_PacsReport_SaveWord                    '����ʾ�ʾ��
            control.Visible = IIf(mintWordPower <> -1, True, False)
        Case conMenu_Edit_Delete                            'ɾ������
            control.Visible = (mReportID <> 0 And (CheckPopedom(mstrPrivs, "PACS������д") Or CheckPopedom(mstrPrivs, "PACS����ɾ��")))
            If control.Visible = True And CheckPopedom(mstrPrivs, "PACS����ɾ��") And mlngEPRDeptID = mlngDeptID Then Exit Sub      '����ǿ��ɾ�������ҵı���
            If control.Visible = True Then control.Visible = mlngEPRDeptID = mlngDeptID
            If control.Visible = True Then control.Visible = (mĿ��汾 = 1)
            If control.Visible = True Then control.Visible = (CheckPopedom(mstrPrivs, "PACS���˱���") Or mstrEPR������ = UserInfo.����)

            '�����ȶ�ɾ�������Enable���л������ã������湤��վ���棬������ݱ����״̬��������һ���Ƿ����ɾ��
            If control.Visible = True Then control.Enabled = Not mblnReadOnly
            
        Case conMenu_PacsReport_ClearWritingState       '������桰�����С���״̬,������������ҵı�����
            control.Visible = CheckPopedom(mstrPrivs, "PACS����ɾ��")
            If control.Visible = True And mlngAdviceID <> 0 Then
                '�����״̬���Ĳ˵�����ʱ�����Ҽ������˵�������������ʾ�˵�������һֱ��ˢ�µ�
                '����ʾ֮ǰ�Ȳ�ѯ���ݿ⣬�����ǰ�в����ˣ�����ʾ�˲˵�
                Dim rsTemp As ADODB.Recordset
                Dim strSql As String
                strSql = "Select ҽ��ID From Ӱ�����¼ Where ҽ��ID = [1] And ������� Is Not Null "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�жϱ����Ƿ��ڴ�����", mlngAdviceID)
                control.Visible = (rsTemp.RecordCount <> 0)
            End If
        
        Case conMenu_File_Exit      '�˳�,��������ģʽ�£���ʾ���˳�����ť
            control.Visible = IIf(mblnSingleWindow = True, True, False)
            
        Case conMenu_PacsReport_Default
    End Select
End Sub

Private Function GetReportLastSaveTime(ByVal lngAdviceID As Long) As Date
'��ȡ������󱣴��ʱ��
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetReportLastSaveTime = mdtReportTime
    
    strSql = "select ����ʱ�� from ����ҽ������ a, ���Ӳ�����¼ b where a.����ID=b.ID and a.ҽ��ID=[1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsData.RecordCount <= 0 Then
        If mReportID > 0 Then GetReportLastSaveTime = zlDatabase.Currentdate
        Exit Function
    End If
    
    GetReportLastSaveTime = Nvl(rsData!����ʱ��, mdtReportTime)
End Function

Private Sub chkOtherDeptReport_Click()
    mblnCheckOtherDeptReport = chkOtherDeptReport.value
    Call subShowHistoryList
End Sub

Private Sub cmdSelectWord_Click()
    Dim strReportVieweType As String
    
    On Error GoTo err
    
    ' mintReportViewType 0-�������CheckView��1-������Result��2-����Advice
    If mintReportViewType = 0 Then
        strReportVieweType = ReportViewType_�������
    ElseIf mintReportViewType = 1 Then
        strReportVieweType = ReportViewType_������
    Else
        strReportVieweType = ReportViewType_����
    End If
    
    If rtxtReport.SelText <> "" Then
        Call mfrmReportWord_WordSelected(rtxtReport.SelText, strReportVieweType, False, True)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub cmdViewImage_Click()
    '�򿪹�Ƭվ��ͼ��
    Dim lngViewAdviceID As Long
    Dim strTmp As String
    
    On Error GoTo err
    If lvHistoryList.SelectedItem Is Nothing Then Exit Sub
    
    strTmp = lvHistoryList.SelectedItem.Key
    If InStr(strTmp, M_STR_LISTVIEWKET_PROCESS) > 0 Or InStr(strTmp, M_STR_LISTVIEWKEY_DESCRIBE) > 0 Then
        Exit Sub
    Else
        lngViewAdviceID = Mid(strTmp, 2)
    End If
        
    Call OpenViewer(1, pobjPacsCore, lngViewAdviceID, True, mobjOwner)
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    '����Pane��ID˳�� 1-���������2-��ʷ���棻3-�ʾ�ʾ����4-����ͼ��5-��Ƶ�ɼ���6-ר�Ʊ��档
    Select Case Item.ID
        Case 1  '��ʷ����
            Item.Handle = picReportHistoryList.hWnd
            picReportHistoryList.Visible = True
        Case 2  '�������
            Item.Handle = picReportViewContainer.hWnd 'mfrmReportView.hWnd
            zlCommFun.ShowChildWindow mfrmReportView.hWnd, picReportViewContainer.hWnd
        Case 3  '�ʾ�ʾ��
            If Not mfrmReportWord Is Nothing Then
                Item.Handle = picReportWordContainer.hWnd ' mfrmReportWord.hWnd
                picReportWordContainer.Visible = True
                zlCommFun.ShowChildWindow mfrmReportWord.hWnd, picReportWordContainer.hWnd
            End If
        Case 4  '����ͼ
            If Not mfrmReportImage Is Nothing Then
                mfrmReportImage.mblnSingleWindow = mblnSingleWindow
                Item.Handle = mfrmReportImage.hWnd
                'picReportImageContainer.Visible = True
                'zlCommFun.ShowChildWindow mfrmReportImage.hWnd, picReportImageContainer.hWnd
            End If
        Case 5  '��Ƶ�ɼ�
            If Not mobjWork_ImageCap Is Nothing Then
                Item.Handle = mobjWork_ImageCap.ContainerHwnd
            End If
        Case 6  'ר�Ʊ���
            If Not mfrmReportSpecial Is Nothing Then Item.Handle = mfrmReportSpecial.hWnd
    End Select
End Sub


Private Sub Form_Activate()
    '��ʾǶ����Ƶ�ɼ�
    If Not mobjWork_ImageCap Is Nothing Then
        If mblnSingleWindow Then
            Call mobjWork_ImageCap.zlUpdateStudyInf(mlngAdviceID, mlngSendNo, mlngStudyState, mblnMoved, mReportID <> 0)
            Call mobjWork_ImageCap.zlRefreshData
        End If
        
        If mobjWork_ImageCap.HasVideo Then Exit Sub
        Call mobjWork_ImageCap.zlRefreshVideoWindow
    End If
    
'    If mblnSingleWindow Then ConfigFocus
End Sub


Private Sub InitActiveVideoModuleObj()
'��ʼ��ActivexExe��Ƶ�ɼ�ģ�����
    If mobjWork_ImageCap Is Nothing Then
        Set mobjWork_ImageCap = CreateObject("zl9PacsImageCap.clsPacsCapture") ' New zl9PacsCapture.clsPacsCapture
        mobjWork_ImageCap.ParentWindowKey = Me.Name & IIf(mblnSingleWindow = True, "Dock", "")
        mobjWork_ImageCap.IsReported = (mReportID <> 0)
        
        Call mobjWork_ImageCap.zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngDeptID, Me.hWnd, mobjOwner, True)
    End If
End Sub



Public Sub RefreshVideo()
    If Not mobjWork_ImageCap Is Nothing Then
        Call mobjWork_ImageCap.zlRefreshVideoWindow
    End If
End Sub


'���ͱ��棬��ר�Ʊ���������
Public Sub SendReport(ByVal strDescription As String, _
    ByVal strResult As String, ByVal strAdvice As String)
    
    mfrmReportView.rtxtCheckView.Text = strDescription
    mfrmReportView.rtxtResult.Text = strResult
    mfrmReportView.rTxtAdvice.Text = strAdvice
    
End Sub

'��ȡ���棬��ר�Ʊ���������
Public Sub GetReport(ByRef strDescription As String, _
    ByRef strResult As String, ByRef strAdvice As String)
    
    strDescription = mfrmReportView.rtxtCheckView.Text
    strResult = mfrmReportView.rtxtResult.Text
    strAdvice = mfrmReportView.rTxtAdvice.Text
    
End Sub

'������棬��ר�Ʊ���������
Public Sub ClearReport(Optional ByVal blnClearDescription As Boolean = True, _
    Optional ByVal blnClearResult As Boolean = True, _
    Optional ByVal blnClearAdvice As Boolean = True)
    
    If blnClearDescription Then mfrmReportView.rtxtCheckView.Text = ""
    If blnClearResult Then mfrmReportView.rtxtResult.Text = ""
    If blnClearAdvice Then mfrmReportView.rTxtAdvice.Text = ""
    
End Sub

Private Sub Form_Load()
    mblnClosed = False
    
    InitCommandBars '��ʼ���˵����������޹�
    
    mblnMenuDownState = False
End Sub


Private Function GetSignVerifyType() As Long
'��ȡǩ�����ͣ�Ĭ��Ϊ����ǩ��,1��ʾ����ǩ��
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    GetSignVerifyType = 0
    
    strSql = "select Zl_Fun_Getsignpar(7, [1]) as ǩ������ from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯǩ������", mlngDeptID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetSignVerifyType = Nvl(rsData!ǩ������, 0)
End Function

Private Sub InitLoaclParas(lngDeptID As Long, lngModuleId As Long, strPrivs As String, Optional blnIsPacsStation As Boolean = False)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strRegPath As String
    Dim blnInCreatProReport As Boolean
    
    If lngDeptID = 0 Then Exit Sub
    
    mlngDeptID = lngDeptID
    mstrPrivs = strPrivs
    mlngModule = lngModuleId
    
    '����Ĭ��ֵ
    mblnShowImage = False       'Ĭ�ϲ���ʾͼ������
    mblnShowSpecial = False     'Ĭ�ϲ���ʾר�Ʊ���
    mblnShowVideoCapture = False 'Ĭ�ϲ���ʾͼ��ɼ�����
    
    mstrSpecialForm = ""
    mblnExitAfterPrint = False  'Ĭ�ϴ�ӡ����󲻹رմ���
    mintWordDblClick = 0        'Ĭ�ϴʾ�˫����ֱ��д�뱨��
    mintImageDblClick = 0       'Ĭ������ͼ˫����ֱ��д�뱨��
    pReport_CheckViewName = "�������"  'Ĭ������
    pReport_ResultName = "������"     'Ĭ������
    pReport_AdviceName = "����"         'Ĭ������
'    mblnIgnoreResult = False            '���Խ��������
'    mintResultInput = 1                 '������ʾ��Ĭ����ǩ������ʾ
    mblnShowWord = True                 'Ĭ��һֱ��ʾ�ʾ�ʾ��
    mblnCheckPrintPara = False             'Ĭ�������ӡ
    mblnCheckOtherDeptReport = False    'Ĭ�ϲ��鿴�����Ƶ���ʷ����
    mblnUntreadPrinted = False          'Ĭ����˴�ӡ���������
    mintPaneID = 1                      'Ĭ��ѡ��PaneΪ��һ��Pane
    
    mblnTechReptSame = GetDeptPara(lngDeptID, "ֻ����д�Լ����ı���", 0) = "1"  'ֻ����д�Լ����ı���
    mstrPatholMaterialInfo = zlDatabase.GetPara("ȡ����������", glngSys, mlngModule, "1,1,1,1,1,1,1,1,1,1")
    
    '��ȡ��������������������򣬽������� ��ǩ������ĸ߶�
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    End If
    
    mlngCY21 = GetSetting("ZLSOFT", strRegPath, "CY21", 500)
    mlngCY22 = GetSetting("ZLSOFT", strRegPath, "CY22", 250)
    mlngCX1 = GetSetting("ZLSOFT", strRegPath, "CX1", 250)
    mlngCX2 = GetSetting("ZLSOFT", strRegPath, "CX2", 500)
    mlngCX3 = GetSetting("ZLSOFT", strRegPath, "CX3", 250)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 250)
    mlngCX4 = GetSetting("ZLSOFT", strRegPath, "CX4", 250)
    mlngCY4 = GetSetting("ZLSOFT", strRegPath, "CY4", 250)
    mlngPicHistoryX = GetSetting("ZLSOFT", strRegPath, "PicHistoryX", 250)
    mlngPicHistoryY = GetSetting("ZLSOFT", strRegPath, "PicHistoryY", 250)
    mlngPrivateWordY = GetSetting("ZLSOFT", strRegPath, "PrivateWordY", 250)
    
    mintPaneID = Val(GetSetting("ZLSOFT", strRegPath, "ѡ��PANE", 1))
    mstrѡ�б����ʽ = GetSetting("ZLSOFT", strRegPath, "ѡ�б����ʽ", "")
    mstr������ = GetSetting("ZLSOFT", strRegPath, "������", "")
     
    mblnCheckOtherDeptReport = (Val(zlDatabase.GetPara("�鿴������ʷ����", glngSys, mlngModule, 0)) = 1)
    
    mblnCompareSize = IIf(Val(GetSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", 1)) <> 0, True, False)
    Call SaveSetting("ZLSOFT", "����ģ��\Ftp", "����FTP�ļ���С�Ա�", IIf(mblnCompareSize, 1, 0))
    
'    '��ȡ��ǰǩ����ʽ��ϵͳ����26��,���Ʊ����Ǵ� 3��ʼ
'    mlngPassType = Val(Mid(zlDatabase.GetPara(26, glngSys), 3, 1))  '����,סԺ,ҽ��,���� (1111),Ϊ��Ĭ�ϲ�������ģʽ
    mlngPassType = GetSignVerifyType()
    
    On Error GoTo err
    strSql = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID)
    
    While Not rsTemp.EOF
        Select Case rsTemp!������
            Case "��ʾ����ͼ��"
                mblnShowImage = Nvl(rsTemp!����ֵ, 0)
            Case "��ʾ��Ƶ�ɼ�"
                mblnShowVideoCapture = Nvl(rsTemp!����ֵ, 0)
                
                If blnIsPacsStation Then
                  mblnShowVideoCapture = False
                End If
                
            Case "��ӡ���˳�"
                mblnExitAfterPrint = Nvl(rsTemp!����ֵ, 0)
            Case "����ͼԤ����ʽ"
                mlngShowBigImg = Nvl(rsTemp!����ֵ, 0)
            Case "��������ͼ����"
                mintMinImageCount = Val(Nvl(rsTemp!����ֵ, 8))
            Case "��ʾר�Ʊ���"
                mblnShowSpecial = Nvl(rsTemp!����ֵ, 0)
            Case "ר�Ʊ���ҳ"
                mstrSpecialForm = Nvl(rsTemp!����ֵ)
            Case "����ͼ˫������"
                mintImageDblClick = Val(Nvl(rsTemp!����ֵ, 0))
            Case "�����������"
                pReport_CheckViewName = Nvl(rsTemp!����ֵ, "�������")
            Case "����������"
                pReport_ResultName = Nvl(rsTemp!����ֵ, "������")
            Case "��������"
                pReport_AdviceName = Nvl(rsTemp!����ֵ, "����")
            Case "��ʾ�ʾ�ʾ��"
                mblnShowWord = IIf(Nvl(rsTemp!����ֵ, 0) = 0, True, False)
            Case "����ʾ�˫������"
                mintWordDblClick = Val(Nvl(rsTemp!����ֵ, 0))
            Case "ƽ������˲��ܴ򱨸�"
                mblnCheckPrintPara = Nvl(rsTemp!����ֵ, 0) = 1
            Case "��˴�ӡ���������"
                mblnUntreadPrinted = Nvl(rsTemp!����ֵ, 0) = 1
            Case "��ӡ��ʽѡ��ʽ"
                mlngPrintFormat = Nvl(rsTemp!����ֵ, 0)
            Case "��ѡ�����ʽ"
                mblOneReportFormat = IIf(Nvl(rsTemp!����ֵ, 0) = 0, False, True)
            Case "�����ֱ�Ӵ�ӡ"
                mblnIsPrint = IIf(Nvl(rsTemp!����ֵ, 0) = 0, False, True)
        End Select
        rsTemp.MoveNext
    Wend
    
    mstrImageLevel = Nvl(GetDeptPara(mlngDeptID, "Ӱ�������ȼ�", "��,��"))
    mstrReportLevel = Nvl(GetDeptPara(mlngDeptID, "���������ȼ�", "��,��"))
    mintImageLevel = Val(GetDeptPara(mlngDeptID, "Ӱ�������ж�", 0))               'Ӱ�������ж�
    mintReportLevel = Val(GetDeptPara(mlngDeptID, "���������ж�", 0))

'    mintCriticalValues = Val(GetDeptPara(mlngDeptID, "Σ������ж�", 0))           'Σ������ж�
    mblnIgnoreResult = GetDeptPara(mlngDeptID, "���Խ��������", 0) = "1" '        '���Խ��������
    mintConformDetermine = Val(GetDeptPara(mlngDeptID, "��������ж�", 0))         '��������ж�
    
    mlngHintType = Val(GetDeptPara(mlngDeptID, "��Ͻ����ʾ����", 0))
    
    
    mblnReportWithResult = GetDeptPara(mlngDeptID, "��Ӱ�����Ϊ����", 0) = "1" '  '��Ӱ�����Ϊ����
    
    '����ʾ�ʾ������
    If mblnShowWord = True And (Not mfrmReportWord Is Nothing) Then
        mfrmReportWord.mblnShowWord = mblnShowWord
        mfrmReportWord.mblnSingleWindow = mblnSingleWindow
        '���ֱ����ʾ�ʾ�ʾ������ȥ���ɼ�����Ŀ��ƿ�
        zlControl.FormSetCaption mfrmReportWord, False, False
    Else
        mfrmReportWord.mblnShowWord = mblnShowWord
        mfrmReportWord.mblnSingleWindow = mblnSingleWindow
        '���ֱ����ʾ�ʾ�ʾ������ȥ���ɼ�����Ŀ��ƿ�
        zlControl.FormSetCaption mfrmReportWord, True, True
    End If
                
'    'ж��ԭ�д���,ж�غ󣬻ᵼ�±����п�����������壬ר�Ʊ����Ժ�Ҫ�޸ĳ�һ��ͳһ�Ĵ��壬Ŀǰ��ʱ������
    If Not mfrmReportSpecial Is Nothing Then
'        If mstrSpecialForm <> Report_Form_frmReportCustom Then Unload mfrmReportSpecial
        If TypeName(mfrmReportSpecial) <> "clsZLPacsProReport" Then Unload mfrmReportSpecial
        
        Set mfrmReportSpecial = Nothing
    End If
'
'    If Not mfrmReportImage Is Nothing Then
'        Unload mfrmReportImage
'        Set mfrmReportImage = Nothing
'    End If
    
    
    'װ��ͼ����
    If mblnShowImage = True Then
        If mfrmReportImage Is Nothing Then Set mfrmReportImage = New frmReportImage
    End If
    
    '����ר�Ʊ��洰��
    If mblnShowSpecial = True Then
        
        Select Case mstrSpecialForm
            Case Report_Form_frmReportES
                Set mfrmReportSpecial = New frmReportES
            Case Report_Form_frmReportUS
                Set mfrmReportSpecial = New frmReportUS
            Case Report_Form_frmReportPathology
                Set mfrmReportSpecial = New frmReportPathology
            Case Report_Form_frmReportCustom
                blnInCreatProReport = True
                Set mfrmReportSpecial = CreateObject("ZLPacsProReport.clsZLPacsProReport")
                Call mfrmReportSpecial.InitPlugin(gcnOracle, Me)
                blnInCreatProReport = False
        End Select
    End If
    
    If mfrmReportSpecial Is Nothing Then    '���û���ҵ���Ӧ��ר�ƴ��壬������Ϊ��ʹ��ר�Ʊ���
        mblnShowSpecial = False
    End If
    
    Exit Sub
err:
    If blnInCreatProReport = True And (err.Number = 429 Or err.Number = -2147024770) Then
        MsgBoxD Me, "û���ҵ��Զ���ר�Ʊ��沿����ZLPacsProReport.dll������ע��˲��������ԡ�"
        Set mfrmReportSpecial = Nothing
        mblnShowSpecial = False
    Else
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End If
End Sub

Private Sub InitReportFormat()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i  As Integer
    
    ReDim rptFormats(1) As rptFormat
    rptFormats(1).ID = 0
    rptFormats(1).strName = "��׼��ʽ"
    
    If mFileID = 0 Then Exit Sub
    
    strSql = "Select Id,���� From ��������Ŀ¼ Where �ļ�ID = [1] And ����= 0 And (ͨ�ü�=0 Or (ͨ�ü�=1 And ����ID=[2]) " & _
            " Or (ͨ�ü�=2 And ��ԱID= [3])) "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mFileID, UserInfo.����ID, UserInfo.ID)
    If rsTemp.RecordCount <> 0 Then
        ReDim Preserve rptFormats(rsTemp.RecordCount + 1) As rptFormat
        For i = 1 To rsTemp.RecordCount
            rptFormats(i + 1).ID = rsTemp!ID
            rptFormats(i + 1).strName = rsTemp!����
            rsTemp.MoveNext
        Next i
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    Dim strRegPath As String
    
    '��ʾ�Ƿ񱣴汨��
    Call PromptModify

    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    End If
    
    '���汨����ʷ����Ŀ�Ⱥ͸߶�
    SaveSetting "ZLSOFT", strRegPath, "PicHistoryY", lvHistoryList.Height
    SaveSetting "ZLSOFT", strRegPath, "PicHistoryX", picReportHistoryList.Width
    
    
    '������ʷ������ʾ״̬
    zlDatabase.SetPara "�鿴������ʷ����", chkOtherDeptReport.value, glngSys, mlngModule
    
    '���汨���е�DockingPaneλ��
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", strRegPath & "\" & mlngModule & "\" & TypeName(dkpMain), dkpMain.Name & mlngDeptID, dkpMain.SaveStateToString)
    End If
    
    '�����һ����ѡ�е�PANE���
    mintPaneID = 1
    For i = 1 To dkpMain.PanesCount
        If dkpMain.Panes(i).Selected Then
            mintPaneID = i
            Exit For
        End If
    Next i
    SaveSetting "ZLSOFT", strRegPath, "ѡ��PANE", mintPaneID
    
    '������Զ��屨���ʽ��ӡ���򱣴�ѡ�б����ʽ
    If mblnʹ���Զ��屨�� Then
        SaveSetting "ZLSOFT", strRegPath, "ѡ�б����ʽ", mstrѡ�б����ʽ
        SaveSetting "ZLSOFT", strRegPath, "������", mstr������
    End If

    If mblnShowVideoCapture Then
        If Not mobjWork_ImageCap Is Nothing Then
            Set mobjWork_ImageCap = Nothing
        End If
    End If
    
    'ж���Ӵ���
    If Not mfrmReportView Is Nothing Then
        Unload mfrmReportView       '��������
        Set mfrmReportView = Nothing
    End If
    
    If Not mobjReport Is Nothing Then
        Unload mobjReport.zlGetForm        '���Ӳ�������
        Set mobjReport = Nothing
    End If
    
    If Not mfrmReportWord Is Nothing Then
        Unload mfrmReportWord       '�ʾ�ʾ��
        Set mfrmReportWord = Nothing
    End If
    
    If Not mfrmReportImage Is Nothing Then
        Unload mfrmReportImage   'ͼ��ѡ��
        Set mfrmReportImage = Nothing
    End If
    
    If Not mfrmReportSpecial Is Nothing Then
        If mstrSpecialForm <> Report_Form_frmReportCustom Then Unload mfrmReportSpecial
        Set mfrmReportSpecial = Nothing
    End If

    '��������ģʽ,��ģʽ�¼�¼����λ��,�����ر��¼�
    If mblnSingleWindow = True Then
        Call SaveWinState(Me, App.ProductName)
        
        RaiseEvent AfterClosed(mlngAdviceID)
        
'        If Not mobjOwner Is Nothing Then
'            mobjOwner.EditorClosed (mlngAdviceID)
'        End If
    End If

    mblnSingleWindow = False
    mblnClosed = True
End Sub


Private Sub lvHistoryList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    zlControl.LvwSortColumn lvHistoryList, ColumnHeader.Index
End Sub

Private Sub lvHistoryList_DblClick()
On Error GoTo errH
    Dim lngViewAdviceID As Long
    Dim lngViewReportID As Long
    Dim strTmp As String
    
    If lvHistoryList.SelectedItem Is Nothing Then Exit Sub
    strTmp = lvHistoryList.SelectedItem.Key
        
    If InStr(strTmp, M_STR_LISTVIEWKET_PROCESS) > 0 Or InStr(strTmp, M_STR_LISTVIEWKEY_DESCRIBE) > 0 Then
        Exit Sub
    Else
        lngViewAdviceID = Mid(strTmp, 2)
    End If
    
    lngViewReportID = lvHistoryList.SelectedItem.SubItems(5)
    Call frmReportHistory.zlShowMe(Me, lngViewAdviceID, lngViewReportID)
    Exit Sub
errH:
    Call MsgBox(err.Description, vbOKOnly, "��ʾ")
End Sub

Public Sub WordItemClick(strReportViewType As String, strReportViewTypeAlias As String, strContext As String)
    If mblnShowWord = True Then Exit Sub


    If mblnSingleWindow = True Then
        Call mfrmReportWord.zlShowMe(Me, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    Else
        Call mfrmReportWord.zlShowMe(mobjOwner, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    End If
End Sub

Private Sub lvHistoryList_ItemClick(ByVal Item As MSComctlLib.ListItem)
'��������lvHistoryList.ListItems�ؼ��ַ�Ϊ process�����̱��� ��describe���޼�������������������������������ݣ�ԭ��ʹ�õ�K��
On Error GoTo err
    Dim strSql As String
    Dim strText As String
    Dim strFormatContext As String
    Dim strSize As String
    Dim lngListKey As Long '�б���Ŀ�ؼ���ID
    Dim rsTemp As ADODB.Recordset
    
    rtxtReport.Text = ""
    strSize = IIf(Val(mfrmReportView.MenuFontSize) <> 0, Val(mfrmReportView.MenuFontSize), Val(mfrmReportView.rtxtCheckView.Font.Size))
    strSize = 2 * Round(Val(strSize))
    strFormatContext = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                       "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                       "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs24 "
        
    cmdViewImage.Enabled = False
    
    If InStr(Item.Key, M_STR_LISTVIEWKET_PROCESS) > 0 Then
        lngListKey = Val(Mid(Item.Key, 8, Len(Item.Key) - 7))
        Call LoadProcessReport(strFormatContext, strSize, lngListKey)
    ElseIf InStr(Item.Key, M_STR_LISTVIEWKEY_DESCRIBE) > 0 Then
        lngListKey = Val(Mid(Item.Key, 9, Len(Item.Key) - 8))
        Call LoadDescription(strFormatContext, strSize, lngListKey)
    Else
        Call LoadReportContent(Item, strFormatContext, strSize)
        '����Ƿ��б���ͼ��
    
        strSql = "Select ���UID from Ӱ�����¼ where ҽ��ID =[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Mid(Item.Key, 2))
        If rsTemp.EOF = False Then
            If Nvl(rsTemp!���uid) <> "" Then
                cmdViewImage.Enabled = True
            End If
        End If
    End If
    
    cmdSelectWord.Enabled = CheckPopedom(mstrPrivs, "PACS������д")

    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    Else
        Call MsgBox(err.Description, vbOKOnly, "��ʾ")
    End If
End Sub

Private Sub SetControlFocus(objControl As Object, ByVal strReportViewType As String)
On Error Resume Next
    If objControl.Visible Then
        mstrCurReportViewType = strReportViewType
        objControl.SetFocus
    End If
err.Clear
End Sub


Private Sub mfrmReportImage_AfterReleationImage(ByVal lngReleationType As Long)
    RaiseEvent AfterReleationImage(mlngAdviceID, mlngSendNo, mlngStudyState, lngReleationType)
End Sub

Private Sub mfrmReportImage_AfterShowBigImage()
On Error Resume Next
    If mfrmReportView Is Nothing Then Exit Sub
    If mfrmReportView.Visible = False Then Exit Sub
On Error Resume Next
    '����ƶ���ʾ��ͼ��λ����༭��
    Select Case CurReportViewType
        Case ReportViewType_�������
            If ReportViewForm.rtxtCheckView.Visible Then ReportViewForm.rtxtCheckView.SetFocus
        Case ReportViewType_������
            If ReportViewForm.rtxtResult.Visible Then ReportViewForm.rtxtResult.SetFocus
        Case ReportViewType_����
            If ReportViewForm.rTxtAdvice.Visible Then ReportViewForm.rTxtAdvice.SetFocus
    End Select
End Sub

Private Sub mfrmReportView_AdviceClick(ByVal strContext As String)
    If mstrCurReportViewType = ReportViewType_���� Then Exit Sub
    
    mintReportViewType = 2
    If mblnShowWord = True Then
        Call mfrmReportWord.zlRefresh(mFileID, ReportViewType_����, pReport_AdviceName, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mblnShowWord, mintWordDblClick, mintWordPower, mblnEditable)
        mstrCurReportViewType = ReportViewType_����
'        Call SetControlFocus(mfrmReportView, ReportViewType_����)
'        Call dkpMain.RedrawPanes
    End If
End Sub

Private Sub refreshWord(ByVal strContext As String)
    'ˢ�´ʾ���棬���������ý��� 100566
    If mstrCurReportViewType = ReportViewType_������� Then Exit Sub
    
    mstrCurReportViewType = ReportViewType_�������
        
    If mblnShowWord = True Then
        Call mfrmReportWord.zlRefresh(mFileID, ReportViewType_�������, pReport_CheckViewName, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mblnShowWord, mintWordDblClick, mintWordPower, mblnEditable)
    End If
End Sub

Private Sub mfrmReportView_CheckViewClick(ByVal strContext As String)
    If mstrCurReportViewType = ReportViewType_������� Then Exit Sub
    
    mintReportViewType = 0
    If mblnShowWord = True Then
        Call mfrmReportWord.zlRefresh(mFileID, ReportViewType_�������, pReport_CheckViewName, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mblnShowWord, mintWordDblClick, mintWordPower, mblnEditable)
        mstrCurReportViewType = ReportViewType_�������
'        Call SetControlFocus(mfrmReportView, ReportViewType_�������)
'        Call dkpMain.RedrawPanes
    End If
End Sub

Private Sub mfrmReportView_ResultClick(ByVal strContext As String)
    If mstrCurReportViewType = ReportViewType_������ Then Exit Sub
    
    mintReportViewType = 1
    If mblnShowWord = True Then
        Call mfrmReportWord.zlRefresh(mFileID, ReportViewType_������, pReport_ResultName, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mblnShowWord, mintWordDblClick, mintWordPower, mblnEditable)
        mstrCurReportViewType = ReportViewType_������
'        Call SetControlFocus(mfrmReportView, ReportViewType_������)
'        Call dkpMain.RedrawPanes
    End If
End Sub

Private Sub mfrmReportWord_AddSampleWord(ByVal blnIsAllWord As Boolean)
    Call subSaveWord(IIf(blnIsAllWord, 2, 0))
End Sub

Private Sub mfrmReportWord_ModifySampleWord()
    Call subSaveWord(1)
End Sub

Private Sub mfrmReportWord_WordSelected(strWord As String, strReportViewType As String, blnIsPopupWindInsert As Boolean, blnAddCrlf As Boolean)
    '�ж�����Ӧ�û�д�����
    
    '������治����༭���������޸ı���ʾ�
    If mblnReadOnly Then Exit Sub
    
    If blnAddCrlf = True Then
        strWord = strWord & vbCrLf
    End If
    
    Select Case strReportViewType
        Case ReportViewType_�������
            If blnIsPopupWindInsert Then mfrmReportView.rtxtCheckView.Text = ""
            Call mfrmReportView.zlWriteReport(strWord, 0)
        Case ReportViewType_������
            If blnIsPopupWindInsert Then mfrmReportView.rtxtResult.Text = ""
            Call mfrmReportView.zlWriteReport(strWord, 1)
        Case ReportViewType_����
            If blnIsPopupWindInsert Then mfrmReportView.rTxtAdvice.Text = ""
            Call mfrmReportView.zlWriteReport(strWord, 2)
        Case ReportViewType_�������
            If mfrmReportSpecial.Name = "frmReportES" Then
                If blnIsPopupWindInsert Then mfrmReportSpecial.txtPathologyDiag.Text = ""
                Call mfrmReportSpecial.zlWriteWord(strWord, strReportViewType)
            End If
        Case ReportViewType_��첿λ
            If mfrmReportSpecial.Name = "frmReportES" Then
                If blnIsPopupWindInsert Then mfrmReportSpecial.txt��첿λ.Text = ""
                Call mfrmReportSpecial.zlWriteWord(strWord, strReportViewType)
            End If
    End Select
End Sub

Private Sub mobjCustomReport_AfterPrint(ByVal ReportNum As String)
    '�����¼��ӡ���¼�
    If Not mobjOwner Is Nothing Then
        mobjOwner.AfterPrinted (mlngAdviceID)
    Else
        RaiseEvent AfterPrinted(mlngAdviceID)
    End If
    mblnPrintOK = True
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    
    '�����¼��ӡ���¼�
    If Not mobjOwner Is Nothing Then
        mobjOwner.AfterPrinted (lngOrderID)
    Else
        RaiseEvent AfterPrinted(lngOrderID)
    End If
    mblnPrintOK = True
End Sub

Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long, ByVal lngSaveType As Long)

    Call AfterReportSaved(lngOrderID, lngSaveType)
    '���±༭���е�����
    Call zlRefreshFace(True)
End Sub


Private Sub mfrmReportView_ShowWord(intReportViewType As Integer, strContext As String)
    Dim strReportViewType As String
    Dim strReportViewTypeAlias As String
    
    Select Case intReportViewType
        Case 0
            strReportViewType = ReportViewType_�������
            strReportViewTypeAlias = pReport_CheckViewName
        Case 1
            strReportViewType = ReportViewType_������
            strReportViewTypeAlias = pReport_ResultName
        Case 2
            strReportViewType = ReportViewType_����
            strReportViewTypeAlias = pReport_AdviceName
    End Select
    
    If mblnSingleWindow = True Then
        Call mfrmReportWord.zlShowMe(Me, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    Else
        Call mfrmReportWord.zlShowMe(mobjOwner, mFileID, strReportViewType, strReportViewTypeAlias, strContext, mlngAdviceID, mlngDeptID, mblnSingleWindow, mlngModule, mintWordPower, mblnEditable)
    End If
End Sub


Private Sub picReportDetail_Resize()
    On Error Resume Next
    
    rtxtReport.Left = 50
    rtxtReport.Top = cmdViewImage.Top + cmdViewImage.Height + 50
    rtxtReport.Width = Abs(picReportDetail.Width - 100)
    rtxtReport.Height = Abs(picReportDetail.Height - cmdViewImage.Height - 300)
    
    cmdViewImage.Left = 200
    cmdViewImage.Top = 200
    
    cmdSelectWord.Left = cmdViewImage.Left + cmdViewImage.Width + 200
    cmdSelectWord.Top = cmdViewImage.Top
End Sub

Private Sub picReportHistoryList_Resize()
    On Error Resume Next
    
    chkOtherDeptReport.Left = 0
    chkOtherDeptReport.Top = 0
    
    lvHistoryList.Left = 0
    lvHistoryList.Top = chkOtherDeptReport.Height + 10
    lvHistoryList.Width = picReportHistoryList.ScaleWidth
    lvHistoryList.Refresh
    
    picReportDetail.Left = 20
    picReportDetail.Width = Abs(picReportHistoryList.ScaleWidth - 20)
    picReportDetail.Height = Abs(picReportHistoryList.ScaleHeight - picReportDetail.Top - 50)
    
    Call ucSplitterH.RePaint
End Sub

Private Sub picReportViewContainer_Resize()
On Error Resume Next
    Call MoveWindow(mfrmReportView.hWnd, 0, 0, _
            picReportViewContainer.ScaleX(picReportViewContainer.Width, vbTwips, vbPixels), _
            picReportViewContainer.ScaleY(picReportViewContainer.Height, vbTwips, vbPixels), 1)
err.Clear
End Sub

Private Sub picReportWordContainer_Resize()
On Error Resume Next
    Call MoveWindow(mfrmReportWord.hWnd, 0, 0, _
            picReportWordContainer.ScaleX(picReportWordContainer.Width, vbTwips, vbPixels), _
            picReportWordContainer.ScaleY(picReportWordContainer.Height, vbTwips, vbPixels), 1)
err.Clear
End Sub

Private Sub pobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    Call RefPacsPic 'ˢ��ͼƬ
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrPopControl As CommandBarControl
    Dim intTMP As Integer
    Dim cbrEdit As CommandBarEdit
        
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '�ɼ�����������
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Open, "��д"): cbrControl.IconId = 3002: cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.IconId = 102: cbrControl.ToolTipText = "����Ԥ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.IconId = 103: cbrControl.ToolTipText = "�����ӡ"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Save, "����"): cbrControl.IconId = 3091: cbrControl.ToolTipText = "���汨��"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Sign, "ǩ��"): cbrControl.IconId = 3003: cbrControl.ToolTipText = "ǩ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Reject, "���沵��"): cbrControl.IconId = 229: cbrControl.ToolTipText = "���沵��"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_RejectHistory, "������ʷ"): cbrControl.IconId = 8341: cbrControl.ToolTipText = "������ʷ"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelSign, "����"): cbrControl.IconId = 3004: cbrControl.ToolTipText = "����ǩ��"
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_VerifySign, "ǩ����֤"): cbrControl.IconId = 8044: cbrControl.ToolTipText = "ǩ����֤"
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_RepFormat, "��ӡ��ʽ"): cbrControl.IconId = 3031: cbrControl.ToolTipText = "ѡ���Զ��屨���ʽ"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_AddNumber, "���"): cbrControl.IconId = 9023: cbrControl.ToolTipText = "����������������"
        'Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_FontSet, "����"): cbrControl.IconId = 509: cbrControl.ToolTipText = "��������"
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PacsReport_FontSet, "�ֺ�"): cbrControl.IconId = 509: cbrControl.ToolTipText = "��������"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSetDefault, "Ĭ��", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet14, "14", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet16, "16", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet22, "22", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet28, "28", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet36, "36", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet42, "42", "", 0, False)
                
                intTMP = Val(zlDatabase.GetPara("������ʾ�ֺ�", glngSys, glngModul))
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlEdit, conMenu_PacsReport_FontSetUser, "�Զ���", "", 0, False)
                
                If intTMP <> 0 And IsCostomFont(intTMP) Then
                    Set cbrEdit = cbrMain.FindControl(xtpControlEdit, conMenu_PacsReport_FontSetUser, True, True)
                    cbrEdit.Text = intTMP
                End If
            End With
        cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����"): cbrControl.IconId = 181: cbrControl.ToolTipText = "��ӡ����"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "�����޶�"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_History, "��ʷ"): cbrControl.IconId = 3564: cbrControl.ToolTipText = "�鿴��ǰ����ʷ����ĵ��޶����"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_SaveWord, "�ʾ�"): cbrControl.IconId = 741: cbrControl.ToolTipText = "���������ݱ���ɴʾ�ʾ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_PrivOrder, "��һ��"): cbrControl.IconId = 21802: cbrControl.ToolTipText = "��һ�����"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_NextOrder, "��һ��"): cbrControl.IconId = 21801: cbrControl.ToolTipText = "��һ�����"
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_PacsReport_SelFormat, "�����ʽ"): cbrControl.IconId = 227: cbrControl.ToolTipText = "ѡ��͸������浥��ʽ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�����༭"): cbrControl.IconId = 3002: cbrControl.ToolTipText = "�õ��Ӳ�����ʽ�༭����"
        'Set cbrControl = .Add(xtpControlButton, comMenu_Petition_Capture, "���뵥"): cbrControl.IconId = 3935: cbrControl.ToolTipText = "�鿴��ɨ������뵥ͼ��": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Default, "�ָ�����"): cbrControl.IconId = 3936: cbrControl.ToolTipText = "�ָ�Ĭ�Ͻ���": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"):: cbrControl.IconId = 191
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If (cbrControl.type = xtpControlButton) Or (cbrControl.type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "Main" '���ó�������˵�
    Next
    cbrToolBar.Position = xtpBarTop
End Sub



' �ӵ��Ӳ����и��ƹ�����һЩ����
'################################################################################################################
'## ���ܣ�  ��ָ����LOB�ֶθ���Ϊ��ʱ�ļ�
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  ������ݵ��ļ�����ʧ���򷵻��㳤��""
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobRead(ByVal Action As Long, ByVal KeyWord As String, Optional ByRef strFile As String) As String
    
    Const conChunkSize As Integer = 10240
    Dim lngFileNum As Long, lngCount As Long, lngBound As Long
    Dim aryChunk() As Byte, strText As String
    Dim rsLob As New ADODB.Recordset
    
    err = 0: On Error GoTo errHand
    
    lngFileNum = FreeFile
    If strFile = "" Then
        lngCount = 0
        Do While True
            strFile = App.Path & "\zlBlobFile" & CStr(lngCount) & ".tmp"
            If Len(Dir(strFile)) = 0 Then Exit Do
            lngCount = lngCount + 1
        Loop
    End If
    Open strFile For Binary As lngFileNum
    
    gstrSQL = "Select Zl_Lob_Read(" & Action & ",'" & KeyWord & "'," & "[1]) as Ƭ�� From Dual"
    lngCount = 0
    Do
        Set rsLob = zlDatabase.OpenSQLRecord(gstrSQL, "zlBlobRead", lngCount)
        If rsLob.EOF Then Exit Do
        If IsNull(rsLob.Fields(0).value) Then Exit Do
        strText = rsLob.Fields(0).value
        
        ReDim aryChunk(Len(strText) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
        Next
        
        Put lngFileNum, , aryChunk()
        lngCount = lngCount + 1
    Loop
    Close lngFileNum
    If lngCount = 0 Then Kill strFile: strFile = ""
    zlBlobRead = strFile
    Exit Function

errHand:
    Close lngFileNum
    Kill strFile: zlBlobRead = ""
End Function

'################################################################################################################
'## ���ܣ�  ��ָ�����ļ����浽ָ����¼��LOB�ֶ���
'##
'## ������  Action      :�������ͣ����������ǲ����ĸ���
'##         KeyWord     :ȷ�����ݼ�¼�Ĺؼ��֣����Ϲؼ����Զ��ŷָ�(��5-���Ӳ�����ʽΪ����)
'##         strFile     :�û�ָ����ŵ��ļ�������ָ��ʱ��ȡ��ǰ·�������ļ���
'##
'## ���أ�  �ɹ�����True��ʧ�ܷ���False
'##
'## ˵����  Actionȡֵ˵����
'##         0-�������ͼ�Σ�1-�����ļ���ʽ��2-�����ļ�ͼ�Σ�3-�������ĸ�ʽ��4-��������ͼ�Σ�5-���Ӳ�����ʽ��6-���Ӳ���ͼ�Σ�
'################################################################################################################
Public Function zlBlobSave(ByVal Action As Long, ByVal KeyWord As String, _
    ByVal strFile As String, ByRef arrSQL() As String) As Boolean
    Dim conChunkSize As Integer
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long
    Dim lngBlocks As Long, lngFileNum As Long
    Dim lngCount As Long, lngBound As Long
    Dim strSql As String
    Dim aryChunk() As Byte, aryHex() As String, strText As String
    
    lngFileNum = FreeFile
    Open strFile For Binary Access Read As lngFileNum
    lngFileSize = LOF(lngFileNum)
    
    err = 0: On Error GoTo errHand
    
    conChunkSize = 2000
    lngModSize = lngFileSize Mod conChunkSize
    lngBlocks = lngFileSize \ conChunkSize - IIf(lngModSize = 0, 1, 0)
    For lngCount = 0 To lngBlocks
        If lngCount = lngFileSize \ conChunkSize Then
            lngCurSize = lngModSize
        Else
            lngCurSize = conChunkSize
        End If
        
        ReDim aryChunk(lngCurSize - 1) As Byte
        ReDim aryHex(lngCurSize - 1) As String
        Get lngFileNum, , aryChunk()
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryHex(lngBound) = Hex(aryChunk(lngBound))
            If Len(aryHex(lngBound)) = 1 Then aryHex(lngBound) = "0" & aryHex(lngBound)
        Next
        
        strText = Join(aryHex, "")
        strSql = "Zl_Lob_Append(" & Action & ",'" & KeyWord & "','" & strText & "'," & IIf(lngCount = 0, 1, 0) & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSql
        
'        Call zlDatabase.ExecuteProcedure(strSql, "zlBlobSave")
    Next
    Close lngFileNum
    zlBlobSave = True
    Exit Function

errHand:
    Close lngFileNum
    zlBlobSave = False
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileUnzip(ByVal strZipFile As String) As String
    Dim strZipPath As String
    If strZipFile = "" Then Exit Function
    If Dir(strZipFile) = "" Then zlFileUnzip = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    If mobjFSO.FileExists(strZipPath & "TMP.RTF") Then mobjFSO.DeleteFile strZipPath & "TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPath
        .Unzip
    End With
    If Dir(strZipPath & "TMP.RTF") <> "" Then
        zlFileUnzip = strZipPath & "TMP.RTF"
    Else
        zlFileUnzip = ""
    End If
End Function

'################################################################################################################
'## ���ܣ�  ���ļ�ѹ��Ϊ���ļ��ŵ���ͬĿ¼��
'## ������  strFile     :ԭʼ�ļ�
'## ���أ�  ѹ���ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Public Function zlFileZip(ByVal strFile As String) As String
    Dim strZipFile As String, lngCount As Long
    If strFile = "" Then Exit Function
    If Dir(strFile) = "" Then zlFileZip = "": Exit Function
    
    lngCount = 0
    Do While True
        strZipFile = Left(strFile, Len(strFile) - Len(Dir(strFile))) & "ZLZIP" & lngCount & ".ZIP"
        If Dir(strZipFile) = "" Then Exit Do
        lngCount = lngCount + 1
    Loop
    
    With mclsZip
        .Encrypt = False: .AddComment = False
        .ZipFile = strZipFile
        .StoreFolderNames = False
        .RecurseSubDirs = False
        .ClearFileSpecs
        .AddFileSpec strFile
        .Zip
        If (.Success) Then
            zlFileZip = .ZipFile
        Else
            zlFileZip = ""
        End If
    End With
End Function




Public Sub RefPacsPic(Optional ByVal lngEventType As TVideoEventType = vetUpdateImg)
    'ˢ�¿�ѡ�ı���ͼ��
    If mblnShowImage = True Then
        If Not mfrmReportImage Is Nothing Then
            mfrmReportImage.RefPacsPic lngEventType
        End If
    End If
End Sub

Private Sub subSetModifyFlag(blnModifyFlag As Boolean)
    mblnModified = blnModifyFlag
    mfrmReportView.pModified = blnModifyFlag
    If mblnShowImage = True Then
        mfrmReportImage.pMarkModified = blnModifyFlag
        mfrmReportImage.pImageModified = blnModifyFlag
    End If
    If mblnShowSpecial = True Then
        mfrmReportSpecial.pModified = blnModifyFlag
    End If
End Sub

Public Function PromptModify(Optional blnCancelEdit As Boolean = False) As Boolean
    'blnCancelEdit =True ��ʾֱ��ȡ������
    If blnCancelEdit = True Then
        Call subSetModifyFlag(False)
        PromptModify = False
        Exit Function
    End If
    
    If mlngAdviceID <> 0 And mblnModified = True And (Not cbrMain.FindControl(, conMenu_PacsReport_Save, True, True) Is Nothing) And Not mblnIsReportDelete Then
        'ģ�ⰴ��ESC���̣�������������Ŀ��ٹ��˴��ڽ��й��˺󣬹��˲˵�����ʧ����Ȼ���Ե�������
        keybd_event VK_ESCAPE, 0, 0, 0
        keybd_event VK_ESCAPE, 0, 2, 0
        
        If MsgBoxD(Me, "���˵ı��������ı䣬�Ƿ񱣴棿", vbYesNo, gstrSysName) = vbYes Then
            If SaveReport(True) Then PromptModify = True
        Else
            '�����汨��ʱ����ձ����������
            mHasChangeFormat = False
            Call UpdateReporter(mlngAdviceID, "")
            
            Call subSetModifyFlag(False)
            PromptModify = False
            
            '����Ƕ��ʽ�ı��淽ʽ����ʱ�൱���ǹرմ���
            If mblnSingleWindow = False Then
                RaiseEvent AfterClosed(mlngAdviceID)
            End If
        End If
    End If
End Function

Private Sub subShowHistoryList()
    Dim strSql As String
    Dim strSQLBack As String
    Dim rsTemp As ADODB.Recordset
    Dim objItem As ListItem
    Dim strTime As String
    Dim iCount As Integer
    Dim strFilter As String
    
    
    
    '�ȼ��Ȩ�ޣ�ȷ���Ƿ���ʾ������ʷ����
    
    If CheckPopedom(mstrPrivs, "PACS�������Ʊ���") Then
        chkOtherDeptReport.value = IIf(mblnCheckOtherDeptReport = True, 1, 0)
        chkOtherDeptReport.Enabled = True
    Else
        chkOtherDeptReport.value = 0
        mblnCheckOtherDeptReport = False
        chkOtherDeptReport.Enabled = False
    End If
    
    If chkOtherDeptReport.value = 1 Then
        strFilter = ""
    Else
        strFilter = " And c.ִ�п���id+0 in(select  ����id  from ������Ա where ��Աid = [2] union all select to_Number([3]) from dual) "
    End If
                    
    strSql = "Select c.Id As ҽ��id, a.Ӱ�����, c.����ʱ��, c.ҽ������, b.����id ,a.�������� as ���ʱ�� " & _
            " From Ӱ�����¼ A, ����ҽ������ B, ����ҽ����¼ C, Ӱ�����¼ D, ����ҽ����¼ E " & _
            " Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And e.Id =[1] And b.ҽ��id = c.Id And " & _
            " (c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null And Nvl(RawToHex(��鱨��ID),' ') =' '"
            
    strSql = strSql & strFilter
    
    If mblnMoved = True Then
        strSQLBack = strSql
        strSQLBack = Replace(strSQLBack, "Ӱ�����¼", "HӰ�����¼")
        strSQLBack = Replace(strSQLBack, "����ҽ������", "H����ҽ������")
        strSQLBack = Replace(strSQLBack, "����ҽ����¼", "H����ҽ����¼")
        strSql = strSql & " UNION ALL  " & strSQLBack
    End If
    
    strSql = strSql & " Order By ���ʱ�� Asc "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ʾ������ʷ", mlngAdviceID, UserInfo.ID, mlngDeptID)
    
    lvHistoryList.ListItems.Clear
        
    zlControl.LvwSelectColumns lvHistoryList, "ҽ��ID,0,0,1;���,500,0,1;���,1000,0,1;���ʱ��,1100,0,1;ҽ������,2000,0,1;����ID,0,0,1", True
        
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
    'ֻ�в���վ�������±�����Ŀ
        iCount = loadPatholReportList(mlngAdviceID)
        If iCount = 0 Then iCount = 1
    Else
        iCount = 1
    End If
    
    With lvHistoryList
        Do While Not rsTemp.EOF
            Set objItem = .ListItems.Add(, "K" & rsTemp!ҽ��ID, rsTemp!ҽ��ID)
            '�������Ŀ
            objItem.SubItems(1) = iCount
            iCount = iCount + 1
            objItem.SubItems(2) = Nvl(rsTemp!Ӱ�����)
            strTime = Format(rsTemp!���ʱ��, "yyyy-mm-dd")
            objItem.SubItems(3) = strTime
            objItem.SubItems(4) = Nvl(rsTemp!ҽ������)
            objItem.SubItems(5) = Nvl(rsTemp!����Id)
            rsTemp.MoveNext
        Loop
    End With
        
    lvHistoryList.Height = mlngPicHistoryY
    
    If lvHistoryList.ListItems.Count > 0 Then
        lvHistoryList.ListItems(1).Selected = True
        Call lvHistoryList_ItemClick(lvHistoryList.ListItems(1))
    Else
        rtxtReport.Text = ""
    End If
    
    dkpMain.FindPane(1).Title = "��ʷ���棨" & lvHistoryList.ListItems.Count & "��"
End Sub

Private Sub ChangeOrder(intType As Integer)
    'intType �л����� 1 --��һ����2--��һ��
    Dim lngRowIndex As Long
    Dim lngNewOrderID As Long
    Dim lngNewSendNo As Long
    Dim blnMoved As Boolean
    
    On Error GoTo err
    
    If mobjOwner.ufgStudyList.DataGrid.Rows <= 1 Then Exit Sub

    lngRowIndex = mobjOwner.ufgStudyList.FindRowIndex(mlngAdviceID, "ҽ��ID", True)
    
    If lngRowIndex <= 0 Then Exit Sub
    
    '�л������֮ǰ�ļ��ı�����������
    If mblnSingleWindow Then Call UpdateReporter(mlngAdviceID, "")
    
    'ֻ���ڷǵǼ�״̬�½����л�
    Do While True
        '������һ������һ��ҽ��
        If intType = 1 Then     '��һ��ҽ��
            lngRowIndex = lngRowIndex - 1
            If lngRowIndex <= 0 Then lngRowIndex = mobjOwner.ufgStudyList.DataGrid.Rows - 1
        ElseIf intType = 2 Then         '��һ��ҽ��
            lngRowIndex = lngRowIndex + 1
            If lngRowIndex >= mobjOwner.ufgStudyList.DataGrid.Rows Then lngRowIndex = 1
        End If
        
        If mobjOwner.ufgStudyList.Text(lngRowIndex, "������") <> "�ѵǼ�" And mobjOwner.ufgStudyList.Text(lngRowIndex, "������") <> "�Ѿܾ�" Then Exit Do
    Loop
        
    
    Call zlUpdateAdviceInf(Val(mobjOwner.ufgStudyList.Text(lngRowIndex, "ҽ��ID")), _
                        Val(mobjOwner.ufgStudyList.Text(lngRowIndex, "���ͺ�")), _
                        Val(mobjOwner.ufgStudyList.Text(lngRowIndex, "���״̬")), _
                        IIf(mobjOwner.ufgStudyList.Text(lngRowIndex, "ת��") = 1, True, False))
        
    '��¼���������
    If mblnSingleWindow Then Call UpdateReporter(mlngAdviceID, UserInfo.����)
    
    If mblnSingleWindow = True Then      '�������ڣ�ֱ��ˢ�±�����
        Call zlRefreshFace(True)
    Else            'Ƕ��ʽ���ڣ�ͨ���ⲿ�¼�����ˢ��,ͬʱˢ�²�������������ҳ��
        mobjOwner.ufgStudyList.DataGrid.ShowCell lngRowIndex, 1
        mobjOwner.ufgStudyList.DataGrid.Row = lngRowIndex
    End If
            
    Exit Sub
err:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub AfterReportSaved(lngOrderID As Long, ByVal lngSaveType As Long)
'lngSaveType:0-��ͨ���棬1-���ǩ����2-���ǩ����3-�����޶� , 4-����ǩ��, 5-������ˣ�6-���������ǩ��ֱ�����ǩ��,7-���˲��������ǩ��ֱ�����ǩ��
    On Error GoTo err
    
    If (lngSaveType = 2 Or lngSaveType = 6) And mblnIsPrint Then
        Call PrintReport(cbrMain.FindControl(, conMenu_File_Print))
    End If
    
    If mblnSingleWindow = True Then
        '�Ե�������ִ�и������AfterReportSaved����
        If Not mobjOwner Is Nothing Then
            Call mobjOwner.AfterReportSaved(lngOrderID, Me, lngSaveType, True)
        End If
    Else

        '��Ƕ��ʽ���ڣ�����AfterSaved�¼�
        RaiseEvent AfterSaved(lngOrderID, Me, lngSaveType, False)
        '����Ƕ��ʽ�ı��淽ʽ����ʱ�൱���ǹرմ���,����AfterClosed�¼�
        RaiseEvent AfterClosed(lngOrderID)
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkPrintState()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
        
    strSql = "Select a.������,a.������,b.������־ ,b.Id From Ӱ�����¼ a ,����ҽ����¼ b Where a.ҽ��id = b.Id And b.Id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��֤�Ƿ���Դ�ӡ", mlngAdviceID)
    
    If rsTemp.EOF = False Then
        mblnCanPrint = IIf(Nvl(rsTemp!������־, 0) = 1, Nvl(rsTemp!������) <> "", Nvl(rsTemp!������) <> "")
    Else
        mblnCanPrint = False
    End If
End Sub

Private Sub AddNumber()
'���ı������ǰ�����������
'mintReportViewType 0-�������CheckView��1-������Result��2-����Advice

    Dim rText As RichTextBox
    Dim strText As String
    Dim iCount As Integer
    Dim iStart As Integer
    
    If mintReportViewType < 0 Or mintReportViewType > 2 Then Exit Sub
    
    '�ж����ĸ��ı��α�ѡ��,��ȡ�ı��εĶ�������
    If mintReportViewType = 0 Then
        Set rText = mfrmReportView.rtxtCheckView
    ElseIf mintReportViewType = 1 Then
        Set rText = mfrmReportView.rtxtResult
    ElseIf mintReportViewType = 2 Then
        Set rText = mfrmReportView.rTxtAdvice
    End If
    
    On Error GoTo err
    strText = rText.Text
    '���ж��ı����Ƿ�����
    If rText.Locked = True Then
        MsgBoxD Me, "�ı��α�����������˫����������������ֱ�š�", vbOKOnly, "��Ϣ��ʾ"
        Exit Sub
    End If
    '���жϸ��ı����е�һ���ַ��Ƿ�����1������ǣ�����ʾ�Ѿ������ֱ�ţ��Ƿ�Ҫ���
    If Left(strText, 1) = "1" Then
        If MsgBoxD(Me, "�����ı����Ѿ��������ֱ�ţ��Ƿ�Ҫ������ֱ�ţ�", vbOKCancel, "��Ϣ��ʾ") = vbCancel Then
            Exit Sub
        End If
    End If
    '��ʼ������ֱ��,ÿһ���س�֮��������ǿո񣬾�������
    iStart = 1
    '��һ��Ҳ��Ҫ�ж��Ƿ��������
    If Left(strText, 1) <> " " Then
        iCount = 1
        strText = iCount & ". " & strText
    Else
        iCount = 0
    End If
    iStart = InStr(iStart, strText, vbCrLf)
    While (iStart <> 0)
        If Mid(strText, iStart + 2, 1) <> " " And Mid(strText, iStart + 2, 2) <> vbCrLf And Mid(strText, iStart + 2, 1) <> "" Then
            iCount = iCount + 1
            strText = Left(strText, iStart + 1) & iCount & ". " & Right(strText, Len(strText) - iStart - 1)
        End If
        iStart = InStr(iStart + 1, strText, vbCrLf)
    Wend
    
    rText.Text = strText
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subChangeRptFormat(ByVal lngIndex As Long)
'���ı�ѡ�е��Զ��屨���ӡ��ʽ
    Dim cbrRptFormat As CommandBarControl
    Dim cbrRptFormatItem As CommandBarControl
    Dim i As Integer
    
    On Error GoTo err
    
    Set cbrRptFormat = cbrMain.FindControl(xtpControlButtonPopup, conMenu_PacsReport_RepFormat, True)
    
    mstrѡ�б����ʽ = ""
    
    If mblOneReportFormat Then
        For i = 1 To cbrRptFormat.CommandBar.Controls.Count
            Set cbrRptFormatItem = cbrRptFormat.CommandBar.Controls(i)
            If i = lngIndex Then
                cbrRptFormatItem.Checked = True
                mstrѡ�б����ʽ = cbrRptFormatItem.Caption
            Else
                cbrRptFormatItem.Checked = False
            End If
        Next i
    Else
        For i = 1 To cbrRptFormat.CommandBar.Controls.Count
            Set cbrRptFormatItem = cbrRptFormat.CommandBar.Controls(i)
            If cbrRptFormatItem.Index = lngIndex Then cbrRptFormatItem.Checked = Not cbrRptFormatItem.Checked
            If cbrRptFormatItem.Checked = True Then
                mstrѡ�б����ʽ = IIf(mstrѡ�б����ʽ = "", cbrRptFormatItem.Caption, mstrѡ�б����ʽ & "," & cbrRptFormatItem.Caption)
            End If
        Next i
    End If
    
    If InStr(mstrFormatInfo, vbCrLf) <> 0 Then
        mstrFormatInfo = Left(mstrFormatInfo, InStr(mstrFormatInfo, vbCrLf) - 1)
    End If
    mstrFormatInfo = mstrFormatInfo & vbCrLf & "��ӡ��ʽ��" & mstrѡ�б����ʽ
    Call mfrmReportView.zlRefreshLblFormat(mstrFormatInfo)
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subPrintReport(blnPrint As Boolean, blnSilent As Boolean)
'ʹ���Զ��屨���ӡ��Ԥ������
'������ blnPrint---True��ӡ��FalseԤ��
'       blnSilent ---ǿ�ƾ�Ĭ��ӡ��������ӡʱ��Ҫ
        
    Dim blnNoAsk As Boolean     '�Ƿ�Ĭ��ӡ
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strExseNo As String, intExseKind As Integer
    Dim strPicPath As String
    Dim objFile As New Scripting.FileSystemObject
    Dim intPCount As Integer
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim i As Integer, j As Integer, intParaCount As Integer
    Dim strPicFile As String
    Dim aryPara(19) As String, aryFlagPara(1) As String     '����ͼ�е�ͼ���¼
    Dim aryPrintPara(19) As String, strFlagString As String 'ʵ�ʴ����Զ��屨�������
    Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
    Dim arr�����ʽ() As String
    Dim int��ʽ�� As Integer
    Dim intRows As Integer, intCols As Integer
    Dim blnIsImageReport As Boolean
    
    On Error GoTo err
    
'    OutputDebugString "ZLPACS>>subPrintReport:1 ��ʼ�Զ��屨���ʽ��ӡ..."
    
    If mblnCanPrint = False Then
        MsgBoxD Me, "��ǰ����δ��ˣ����ܴ�ӡ�����飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ�Ĭ��ӡ
    blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")
    If blnSilent = True Then blnNoAsk = True
    
    '��ȡ����ļ�¼���ʺ�No
    strSql = "Select ��¼����, No From ����ҽ������ Where ҽ��id = [1]"
    If mblnMoved = True Then strSql = Replace(strSql, "����ҽ������", "H����ҽ������")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ǰ��¼���ʺ�No", mlngAdviceID)
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    strExseNo = "" & rsTemp!NO
    intExseKind = Val("" & rsTemp!��¼����)
    
    If mobjCustomReport Is Nothing Then Set mobjCustomReport = New clsReport
    
    If Not blnNoAsk Then
        If mobjCustomReport.ReportPrintSet(gcnOracle, glngSys, mstr������) = False Then
        '�˴�ˢ�»���ɽ������
            Exit Sub
        End If
    End If
    
    '��ȡͼ��
    strPicPath = App.Path & "\TmpImage\"
    If objFile.FolderExists(strPicPath) = False Then objFile.CreateFolder strPicPath
    
'    OutputDebugString "ZLPACS>>subPrintReport:2 ��ȡ��ӡͼ�񻺴�Ŀ¼Ϊ:" & strPicPath
    
    '��ȡ����ͼ�񣨰������ͼ�����ɱ����ļ�
    'һ���������п������ж������ͼ
    intPCount = 0
    strSql = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
    If mblnMoved = True Then strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ��", mReportID)
    
'    OutputDebugString "ZLPACS>>subPrintReport:3 ��ȡ����ͼ."
    Do While Not rsTemp.EOF
        Set cTable = New cEPRTable
        If cTable.GetTableFromDB(cprET_���������, mReportID, Val("" & rsTemp!���ID)) Then
        
'            OutputDebugString "ZLPACS>>subPrintReport:4 ҽ��idΪ" & mlngAdviceID & "�ı���ͼ����Ϊ:" & cTable.Pictures.Count
            For i = 1 To cTable.Pictures.Count
                strPicFile = strPicPath & "PACSPic" & i & ".JPG"
                
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                    Set oPicture = cTable.Pictures(i).DrawFinalPic
                Else
                    Set oPicture = cTable.Pictures(i).OrigPic
                End If
                
'                OutputDebugString "ZLPACS>>subPrintReport:5 �洢���Ϊ" & i & "�ı���ͼ��" & strPicFile
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
            Next i
        End If
        rsTemp.MoveNext
    Loop
    
    '����ѡ����Զ��屨���ʽ����֯ͼ��
    '���ֻѡ����һ�ָ�ʽ�������Ƿ�ֻ��һ��ͼ���,ֻ��һ��ͼ����ʱ���Զ����ͼ��
    '���ѡ����2�����ϵĸ�ʽ�����ֻ��һ��ͼ������������Զ����
    arr�����ʽ = Split(mstrѡ�б����ʽ, ",")
    
    '����û��ѡ���ʽ�����
    If UBound(arr�����ʽ) = -1 Then
        ReDim arr�����ʽ(0) As String
        arr�����ʽ(0) = "1-1"
    End If
    
'    OutputDebugString "ZLPACS>>subPrintReport:6 �жϱ����ʽ."
    
    If UBound(arr�����ʽ) = 0 Then     'ֻ��һ�ָ�ʽ
        int��ʽ�� = Split(arr�����ʽ(0), "-")(0)
        strSql = "Select b.����,b.W,b.H From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = [2] And b.���� not like '���%'"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯ�Ƿ���Ҫ���ͼ��", mstr������, int��ʽ��)
        
        
        If rsTemp.RecordCount = 1 And intPCount >= 1 Then
            '���ͼ��
'            OutputDebugString "ZLPACS>>subPrintReport:7 ��ʼ��ϱ���ͼ�񵽣�" & Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "="))
            
            ResizeRegion intPCount, rsTemp("W"), rsTemp("H"), intRows, intCols
            Set dcmResultImage = AssembleImage(dcmImages, intRows, intCols, rsTemp("H"), rsTemp("W"))
            dcmResultImage.FileExport Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "=")), "JPEG"
            
'            OutputDebugString "ZLPACS>>subPrintReport:8 ����ͼ��������,���ͼ��λ��Ϊ:" & Right(aryPara(0), Len(aryPara(0)) - InStr(aryPara(0), "="))
        End If
    End If
    
'    OutputDebugString "ZLPACS>>subPrintReport:9 ��ʼװ�ظ�ͼ��."
    
    '��ȡͼ�񣬵��ñ���
    
    blnIsImageReport = False
    intPCount = 0       '��¼ͼ�������
    For i = 0 To UBound(arr�����ʽ)
        int��ʽ�� = Split(arr�����ʽ(i), "-")(0)
        
        strSql = "Select b.���� From zlReports a, zlRptItems b" & vbNewLine & _
        "       Where a.Id = b.����id And a.��� = [1] And Nvl(b.����, 0) = 1 And b.���� = 11 And b.��ʽ�� = [2]" & vbNewLine & _
        "       Order By b.����" 'Trunc(b.y/567),Trunc(b.x/567)
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡͼ���", mstr������, int��ʽ��)
        
        'װ��ͼ������
        intParaCount = 0
        Do While Not rsTemp.EOF
            blnIsImageReport = True
            
            '�ֱ�װ�ڱ��ͼ�ͱ���ͼ
            If InStr(rsTemp!����, "���") <> 0 Then '���ͼ
                If aryFlagPara(0) <> "" Then strFlagString = rsTemp!���� & "=" & aryFlagPara(0)
            Else    '����ͼ
                If intPCount > UBound(aryPara) Then Exit Do     'ͼ���������������е�ͼ���˳�
                If aryPara(intPCount) = "" Then Exit Do         '�����е�ͼ���ȱ����еĶ࣬�˳�
                
                aryPrintPara(intParaCount) = rsTemp!���� & "=" & aryPara(intPCount)
                intPCount = intPCount + 1
                intParaCount = intParaCount + 1
            End If
            rsTemp.MoveNext
        Loop
        
        '��������ͼ�αȱ������ٵ����
        For j = intParaCount To UBound(aryPrintPara)
            If aryPrintPara(j) Like "*=*" Then aryPrintPara(j) = ""
        Next j
        
        If mlngModule = 1291 And blnIsImageReport Then
            If Trim(aryPrintPara(0)) = "" And Trim(aryPrintPara(1)) = "" And Trim(aryPrintPara(2)) = "" And Trim(aryPrintPara(3)) = "" Then
'                OutputDebugString "ZLPACS>>subPrintReport:10 �ޱ���ͼ����ʾ����."
                If MsgBox("δ���ִ���ӡ�ı���ͼ���Ƿ������ӡ��", vbYesNo, "��ʾ") = vbNo Then
'                    OutputDebugString "ZLPACS>>subPrintReport:11 �˳������ӡ."
                    Exit Sub
                End If
            End If
        End If
        
        '���ñ���
        Call mobjCustomReport.ReportOpen(gcnOracle, glngSys, mstr������, Nothing, _
            "NO=" & strExseNo, "����=" & intExseKind, "ҽ��ID=" & mlngAdviceID, strFlagString, _
            aryPrintPara(0), aryPrintPara(1), aryPrintPara(2), aryPrintPara(3), aryPrintPara(4), aryPrintPara(5), _
            aryPrintPara(6), aryPrintPara(7), aryPrintPara(8), aryPrintPara(9), aryPrintPara(10), aryPrintPara(11), _
            aryPrintPara(12), aryPrintPara(13), aryPrintPara(14), aryPrintPara(15), aryPrintPara(16), aryPrintPara(17), _
            aryPrintPara(18), aryPrintPara(19), "ReportFormat=" & int��ʽ��, IIf(blnPrint, 2, 1))
            
    Next i
    
    If mlngPrintFormat = 1 Then mstrѡ�б����ʽ = ""

    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subSaveWord(intType As Integer)
'���没���ʾ�ʾ��
'�ӱ���ʾ�����������ȡ�ʾ����ID����������,KEY="T-����ID",TEXT="��������"
'�ӱ���ʾ���Ҷ�ӽ����ȡ�ʾ�ID���ʾ����ƣ�KEY="L-ʾ��ID"��TEXT="ʾ������"
'������ intType ---  0 ������1 �޸�

    Dim strWordString As String
    Dim rText As RichTextBox
    Dim lngClassID As Long      '����ID
    Dim strClassName As String  '��������
    Dim objNode As Node
    Dim lngWordID As Long       '�ʾ�ʾ��ID
    Dim strWordName As String   '�ʾ�ʾ������
    
    
    If mfrmReportWord.trvWordTree.SelectedItem Is Nothing Then
        MsgBoxD Me, "��Ӵʾ�������ѡ����Ҫ����ʾ��λ�á�", vbOKOnly, gstrSysName
        Exit Sub
    End If
    Set objNode = mfrmReportWord.trvWordTree.SelectedItem
    
    If intType = 1 Then         '�޸ģ���Ҫ��ȡ�ʾ�ID�ͷ���ID��������ֻ��Ҫ��ȡ����ID
        '��ȡ�ʾ�ID
        '���жϵ�ǰѡ�еĽ���Ƿ����㻹��Ҷ�ӽ�㣬����Ƿ�����KEY=��T-...��,Ҷ�ӽ��KEY=��L-...��
        ''��Ҷ�ӽ�㣬��Ҫ�����ϼ����,�Ƿ����㣬ֱ����ȡ����ID������
        If Left(objNode.Key, 1) = "L" Then
            lngWordID = Right(objNode.Key, Len(objNode.Key) - 2)
            strWordName = objNode.Text
        Else
            MsgBoxD Me, "������ѡ����Ƿ��࣬��ѡ����Ҫ�޸ĵĴʾ䡣", vbOKOnly, gstrSysName
            Exit Sub
        End If
    ElseIf intType = 2 Then
        strWordString = ""
        
        If mfrmReportView.rtxtCheckView.Text <> "" Then
            strWordString = "<<����>>" & mfrmReportView.rtxtCheckView.Text
        End If
        
        If mfrmReportView.rtxtResult.Text <> "" Then
            strWordString = strWordString & vbCrLf & "<<���>>" & mfrmReportView.rtxtResult.Text
        End If
        
        If mfrmReportView.rTxtAdvice.Text <> "" Then
            strWordString = strWordString & vbCrLf & "<<����>>" & mfrmReportView.rTxtAdvice.Text
        End If
    Else
        '�ӱ����ж�ȡ�ʾ�����
        '��ȡ��ǰ��Ҫ����ɴʾ������
                'mintReportViewType= 0-�������CheckView��1-������Result��2-����Advice
                
        If mintReportViewType = 0 Then
            Set rText = mfrmReportView.rtxtCheckView
        ElseIf mintReportViewType = 1 Then
            Set rText = mfrmReportView.rtxtResult
        Else
            Set rText = mfrmReportView.rTxtAdvice
        End If
        
        If rText.SelLength = 0 Then
            strWordString = rText.Text
        Else
            strWordString = rText.SelText
        End If
    End If
    '��ȡ��ǰ�ʾ����ID
    If Left(objNode.Key, 1) = "L" Then  '��ǰ�����Ҷ�ӽ�㣬��ָ���ϼ�������
            Set objNode = objNode.Parent
    End If
    
    lngClassID = Right(objNode.Key, Len(objNode.Key) - 2)
    strClassName = objNode.Text
    
    Call frmReportWordList.zlShowMe(Me, strWordString, mintWordPower, lngClassID, strClassName, _
                                    mlngDeptID, lngWordID)
End Sub

Private Function GetReportImageSelected() As Boolean
'------------------------------------------------
'���ܣ���鱨��ͼҳ���Ƿ�ǰ�ҳ��
'������
'���أ�True�����ǻҳ�棬False�������ǻҳ��
'-----------------------------------------------
Dim i As Integer

On Error Resume Next

GetReportImageSelected = False

For i = 1 To dkpMain.PanesCount
    If dkpMain.Panes(i).Title = "����ͼ" Then
        GetReportImageSelected = dkpMain.Panes(i).Selected
        Exit For
    End If
Next i
End Function


Private Sub FuncAdviceSignVerify(intǩ���汾 As Integer, blnMoved As Boolean)
'------------------------------------------------
'���ܣ�У���鱨��ĵ���ǩ��(�ɶ���ת�Ƶ�����),У��汾Ϊintǩ���汾 ��ǩ��
'������ intǩ���汾 -- ������Ҫ��֤��ǩ���İ汾
'       blnMoved -- �����Ƿ�Ǩ��
'���أ�
'-----------------------------------------------
    Dim strSource As String
    Dim dblǩ��ID  As Double                  'ǩ�����ڵ��е�ID
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intRule As Integer                  '��¼ǩ������
    
    
    On Error GoTo err
    
    '���ݱ���ID��ǩ���汾����ǩ������
    strSql = "Select Id , ��ʼ�� From ���Ӳ������� Where �ļ�ID = [1] And �������� = 8 and ��ʼ�� =[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���ǩ���汾", mReportID, intǩ���汾)
    If rsTemp.RecordCount = 0 Then
        MsgBoxD Me, "���α���û�а汾Ϊ" & intǩ���汾 & "��ǩ�����޷�������ǩ������֤��"
        Exit Sub
    End If
    
    dblǩ��ID = Val(rsTemp!ID)
    
    '��ȡԴ��
    intRule = GetSignSourceString(2, mReportID, intǩ���汾, blnMoved, Nothing, strSource)
    '������صĹ���=0����ʾ��ȡԴ��ʧ��
    If intRule = 0 Then
        MsgBoxD Me, "���α���汾Ϊ" & intǩ���汾 & "��ǩ��Դ����ȡʧ�ܣ��޷�������ǩ������֤��"
        Exit Sub
    End If
    
    '����ǩ�����󣬶�Դ�Ľ���ǩ����֤
    err.Clear: On Error Resume Next
    If gobjESign Is Nothing Then
        Set gobjESign = Interaction.GetObject(, "zl9ESign.clsESign")
        If gobjESign Is Nothing Then Set gobjESign = CreateObject("zl9ESign.clsESign")
        If err <> 0 Then err = 0
        
        If Not gobjESign Is Nothing Then
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
    End If
        
    On Error GoTo err
        
    If Not gobjESign Is Nothing Then
        'ǩ����֤
        Call gobjESign.VerifySignature(strSource, dblǩ��ID, 2)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub tmrCheckingReportState_Timer()
'ѭ����鱨��༭״̬������ѱ��˱༭���������ʾ
On Error Resume Next
    If mblnReadOnly Then
        tmrCheckingReportState.Enabled = False
        Exit Sub
    End If
    
    If CheckConcurrentReport(Me, mlngAdviceID) Then
        mblnReadOnly = False
        
        tmrCheckingReportState.tag = Val(tmrCheckingReportState.tag) + 1
        
        '����10�����˳�
        If Val(tmrCheckingReportState.tag) > 5 Then tmrCheckingReportState.Enabled = False
    Else
        mblnReadOnly = True
    End If
    err.Clear
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
'����վ�˵����ı��ֺ�
    If Not mfrmReportView Is Nothing Then
        Call mfrmReportView.SetFontSize(bytFontSize)
                
        If lvHistoryList.ListItems.Count > 0 Then
            lvHistoryList.ListItems(1).Selected = True
            Call lvHistoryList_ItemClick(lvHistoryList.ListItems(1))
        End If
    End If
End Sub

Private Sub tmrFocus_Timer()
On Error Resume Next
    tmrFocus.Enabled = False
    
    Call ConfigFocus
err.Clear
End Sub

Private Sub mfrmReportImage_OnImageCountChanged(ByVal intType As Integer, ByVal isNeedRefreshTitle As Boolean)
    RaiseEvent OnImageCountChanged(intType, isNeedRefreshTitle)
End Sub

Public Sub RefreshAfterImage()
    If Not mfrmReportImage Is Nothing Then Call mfrmReportImage.RefreshAfterImage
End Sub

Public Sub UseAfterImgChanged(ByVal blUse As Boolean)
    If Not mfrmReportImage Is Nothing Then Call mfrmReportImage.UseAfterImgChanged(blUse)
End Sub

Public Sub SetMeneFontSize(ByVal intFontSize As Integer)
'�ı䱨��������ʾ�ֺ�

    If Not mfrmReportView Is Nothing Then
        mfrmReportView.MenuFontSize = intFontSize
        
        If lvHistoryList.ListItems.Count > 0 Then
            lvHistoryList.ListItems(1).Selected = True
            Call lvHistoryList_ItemClick(lvHistoryList.ListItems(1))
        End If
    End If
End Sub

Private Function getMenuFontSize() As Integer
    If Not mfrmReportView Is Nothing Then
        getMenuFontSize = mfrmReportView.MenuFontSize()
    End If
End Function

Private Function loadPatholReportList(ByVal lngAdviceID As Long) As Integer
'����ҽ��ID ���ز�����̱������ݵ���ʷ�����б���
'����  0 �쳣   ����ֵ: ��һ���������
'��������lvHistoryList.ListItems.Add��ӵĹؼ��ַ�Ϊprocess�����̱���  describe���޼�����
    Dim objItem As ListItem
    Dim intCount As Integer '�Ѿ��ù������
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    
    loadPatholReportList = 0
    
    '���ؾ޼������б���
    strSql = "select  a.����ҽ��ID,b.ȡ��ʱ�� " & _
                  "from ��������Ϣ a,����ȡ����Ϣ b " & _
                  "where a.����ҽ��id=b.����ҽ��id " & _
                  "and b.���= (select min(c.���) from ����ȡ����Ϣ c where c.����ҽ��id=a.����ҽ��id and a.ҽ��id=[1]) " & _
                  "and a.ҽ��id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ؾ޼������б���", mlngAdviceID)
    
    intCount = 1
    If rsTemp.RecordCount = 1 Then
        Set objItem = lvHistoryList.ListItems.Add(, M_STR_LISTVIEWKEY_DESCRIBE & rsTemp!����ҽ��id, rsTemp!����ҽ��id)
        objItem.SubItems(1) = intCount
        objItem.SubItems(2) = "�޼�����"
        objItem.SubItems(3) = getShortDate(Nvl(rsTemp!ȡ��ʱ��))
        objItem.SubItems(4) = ""
        objItem.SubItems(5) = rsTemp!����ҽ��id
        intCount = 2
    End If
    
    '���ع��̱����б���
    strSql = "select  b.�걾����,b.Id,b.��������,b.�������� " & _
                  "from ��������Ϣ a ,������̱��� b " & _
                  "where a.����ҽ��id=b.����ҽ��id and a.ҽ��id=[1] " & _
                  "order by b.�������� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���ع��̱����б���", mlngAdviceID)
    
    With lvHistoryList
        Do While Not rsTemp.EOF
            Set objItem = lvHistoryList.ListItems.Add(, M_STR_LISTVIEWKET_PROCESS & rsTemp!ID, rsTemp!ID)
            '�������Ŀ
            objItem.SubItems(1) = intCount
            intCount = intCount + 1
            objItem.SubItems(2) = getReportType(Val(Nvl(rsTemp!��������)))
            objItem.SubItems(3) = getShortDate(Nvl(rsTemp!��������))
            objItem.SubItems(4) = "�걾���ƣ�" & Nvl(rsTemp!�걾����)
            objItem.SubItems(5) = rsTemp!ID
            rsTemp.MoveNext
        Loop
    End With
    loadPatholReportList = intCount
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function
Private Function getShortDate(ByVal strDate) As String
    getShortDate = ""
    
    If IsDate(strDate) Then
        getShortDate = Format(strDate, "yyyy-mm-dd")
    End If
End Function

Private Function getReportType(ByVal intType As Integer) As String
'��þ��屨������  ���������ݿ��е�����
    getReportType = ""
    If intType < 0 Or intType > 3 Then Exit Function
    
    Select Case intType
        Case 0
            getReportType = "��������"
        Case 1
            getReportType = "���߱���"
        Case 2
            getReportType = "���ӱ���"
        Case 3
            getReportType = "��Ⱦ����"
    End Select
End Function

Private Sub LoadProcessReport(ByVal strFormatContextOld As String, ByVal strSize As String, ByVal lngListKey As Long)
'���벡����̱�������
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext  As String
    Dim strText As String
    Dim strTitle As String
    
    On Error GoTo errH
    strFormatContext = strFormatContextOld
    strSql = "select �����,������ from ������̱��� where id=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ʷ����鿴���̱���", lngListKey)
                
    If rsTemp.RecordCount <> 0 Then
        strTitle = "�����" & "��"
        strText = Nvl(rsTemp!�����) & vbCrLf
        strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strSize & " " & Replace(strText, vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
        
        strTitle = "������" & "��"
        strText = Nvl(rsTemp!������) & vbCrLf
        strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strSize & " " & Replace(strText, vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
            
        strFormatContext = strFormatContext & "}"
        rtxtReport.SelRTF = strFormatContext
        rtxtReport.SelStart = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub LoadDescription(ByVal strFormatContextOld As String, ByVal strSize As String, ByVal lngListKey As Long)
'����޼���������  mstrPatholMaterialInfo �걾����,ȡ��λ��,��״,������,��Ƭ��,��ȡҽʦ,ȡ��ʱ��,����,��ɫ,�걾��
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext  As String
    Dim strText As String
    Dim strTitle As String
    Dim str�޼����� As String
    Dim blnIsCell As Boolean '�Ŀ��Ƿ���ϸ������ ϸ��������2
    
    On Error GoTo errH
    
    strFormatContext = strFormatContextOld
    strSql = "select  a.�޼�����,a.�������,b.���,b.�걾����, b.��״,b.ȡ��λ��,b.������,b.��ȡҽʦ,b.����,b.��ɫ,b.�걾��,b.ȡ��ʱ��, b.�걾����, c.��Ƭ�� " & _
                      "from ��������Ϣ a ,����ȡ����Ϣ b ,������Ƭ��Ϣ c " & _
                      "where b.�Ŀ�id=c.�Ŀ�id and a.����ҽ��id=c.����ҽ��id and a.����ҽ��id=b.����ҽ��id and a.����ҽ��id=[1] order by b.��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ʷ����鿴�޼�����", lngListKey)
        
    If rsTemp.RecordCount <> 0 Then
        str�޼����� = Nvl(rsTemp!�޼�����)
        blnIsCell = (Val(Nvl(rsTemp!�������)) = 2)   'ϸ��������2
    End If
    
    If UBound(Split(mstrPatholMaterialInfo, ",")) <> 9 Then mstrPatholMaterialInfo = "1,1,1,1,1,1,1,1,1,1"
                
    While Not rsTemp.EOF
    
        strTitle = "����" & Nvl(rsTemp!���) & "��"
        strText = ""
        
        If Split(mstrPatholMaterialInfo, ",")(0) = 1 And Trim(Nvl(rsTemp!�걾����)) <> "" Then strText = "�걾���ƣ�" & Nvl(rsTemp!�걾����)
        If Split(mstrPatholMaterialInfo, ",")(1) = 1 And Trim(Nvl(rsTemp!ȡ��λ��)) <> "" Then strText = IIf(strText <> "", strText & "��" & "ȡ��λ�ã�" & Nvl(rsTemp!ȡ��λ��), "ȡ��λ�ã�" & Nvl(rsTemp!ȡ��λ��))
        If Split(mstrPatholMaterialInfo, ",")(2) = 1 And Trim(Nvl(rsTemp!��״)) <> "" Then strText = IIf(strText <> "", strText & "��" & "��״��" & Nvl(rsTemp!��״), "��״��" & Nvl(rsTemp!��״))
        
        If blnIsCell Then
            If Split(mstrPatholMaterialInfo, ",")(7) = 1 And Trim(Nvl(rsTemp!����)) <> "" Then strText = IIf(strText <> "", strText & "��" & "���ʣ�" & Nvl(rsTemp!����), "���ʣ�" & Nvl(rsTemp!����))
            If Split(mstrPatholMaterialInfo, ",")(8) = 1 And Trim(Nvl(rsTemp!��ɫ)) <> "" Then strText = IIf(strText <> "", strText & "��" & "��ɫ��" & Nvl(rsTemp!��ɫ), "��ɫ��" & Nvl(rsTemp!��ɫ))
            If Split(mstrPatholMaterialInfo, ",")(9) = 1 And Trim(Nvl(rsTemp!�걾��)) <> "" Then strText = IIf(strText <> "", strText & "��" & "�걾����" & Nvl(rsTemp!�걾��), "�걾����" & Nvl(rsTemp!�걾��))
            If Split(mstrPatholMaterialInfo, ",")(3) = 1 And Trim(Nvl(rsTemp!������)) <> "" Then strText = IIf(strText <> "", strText & "��" & "ϸ��������" & Nvl(rsTemp!������), "ϸ��������" & Nvl(rsTemp!������))
        Else
            If Split(mstrPatholMaterialInfo, ",")(3) = 1 And Trim(Nvl(rsTemp!������)) <> "" Then strText = IIf(strText <> "", strText & "��" & "�Ŀ�����" & Nvl(rsTemp!������), "�Ŀ�����" & Nvl(rsTemp!������))
        End If
        
        If Split(mstrPatholMaterialInfo, ",")(4) = 1 And Trim(Nvl(rsTemp!��Ƭ��)) <> "" Then strText = IIf(strText <> "", strText & "��" & "��Ƭ����" & Nvl(rsTemp!��Ƭ��), "��Ƭ����" & Nvl(rsTemp!��Ƭ��))
        If Split(mstrPatholMaterialInfo, ",")(5) = 1 And Trim(Nvl(rsTemp!��ȡҽʦ)) <> "" Then strText = IIf(strText <> "", strText & "��" & "��ȡҽʦ��" & Nvl(rsTemp!��ȡҽʦ), "��ȡҽʦ��" & Nvl(rsTemp!��ȡҽʦ))
        If Split(mstrPatholMaterialInfo, ",")(6) = 1 And Trim(Nvl(rsTemp!ȡ��ʱ��)) <> "" Then strText = IIf(strText <> "", strText & "��" & "ȡ��ʱ�䣺" & Nvl(rsTemp!ȡ��ʱ��), "ȡ��ʱ�䣺" & Nvl(rsTemp!ȡ��ʱ��))
        
        If strText <> "" Then
            strText = strText & vbCrLf
            strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strSize & " " & Replace(strText, vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
        End If
        
        rsTemp.MoveNext
    Wend
    
    If Trim(str�޼�����) <> "" Then strFormatContext = strFormatContext & "\b\cf2\fs24 " & "�޼�����:" & " \par\b0\cf0\fs" & strSize & " " & Replace(str�޼�����, vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
        
    strFormatContext = strFormatContext & "}"
    rtxtReport.SelRTF = strFormatContext
    rtxtReport.SelStart = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub LoadReportContent(ByVal Item As MSComctlLib.ListItem, ByVal strFormatContextOld As String, ByVal strSize As String)
'���뱨������
    Dim lngViewReportID As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim blnShow As Boolean
    Dim strFormatContext  As String
    Dim strText As String
    Dim strTitle As String
    
    On Error GoTo errH
    
    strFormatContext = strFormatContextOld
    lngViewReportID = Item.SubItems(5)
    '��ʾ��������
    
    '��ȡ���������
    strSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = [1] And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ֹ��=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngViewReportID)
                
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!����
            Case "�������"
                strTitle = pReport_CheckViewName
                strText = Nvl(rsTemp!����) & vbCrLf
                blnShow = True
            Case "������"
                strTitle = pReport_ResultName
                strText = Nvl(rsTemp!����) & vbCrLf
                blnShow = True
            Case "����"
                strTitle = pReport_AdviceName
                strText = Nvl(rsTemp!����) & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strSize & " " & Replace(strText, vbCrLf, " \par\cf0\fs" & strSize & " ") & "\par"
        rsTemp.MoveNext
    Wend
    
    strFormatContext = strFormatContext & "}"
    rtxtReport.SelRTF = strFormatContext
    rtxtReport.SelStart = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucSplitterH_OnMoveEnd()
    On Error Resume Next
    mlngPicHistoryY = lvHistoryList.Height
End Sub

Public Sub SetMenuDownState(ByVal blnValue As Boolean)
'���ܣ��޸�mblnMenuDownState��ֵ�����ڴ�������105988
    mblnMenuDownState = blnValue
End Sub

Private Function CheckUserFontValidate(ByVal strValue As String) As Boolean
'���򣺾���abs(val(?))����������֣�������֤��ͨ��������ʾ

    CheckUserFontValidate = True
    
    If Abs(Val(strValue)) = 0 Then
        Call MsgBoxD(Me, "��ע�⣬�Զ����ֺű�����һ������0�����֣�����������", vbOKOnly, gstrSysName)
        CheckUserFontValidate = False
        Exit Function
    End If
    
End Function

Private Function IsCostomFont(ByVal intFontSize As Integer) As Boolean
'���ܣ��ж��Ƿ�ʹ���Զ����ֺ�  ���� true-��
'���򣬲�����103523�����ظ�
    IsCostomFont = True
    
    If intFontSize = 0 Or intFontSize = 14 Or intFontSize = 16 Or intFontSize = 22 Or intFontSize = 28 Or intFontSize = 36 Or intFontSize = 42 Then
        IsCostomFont = False
    End If
    
End Function
