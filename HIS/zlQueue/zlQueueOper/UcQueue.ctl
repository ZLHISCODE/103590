VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl UcQueue 
   ClientHeight    =   6034
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10598
   ScaleHeight     =   6034
   ScaleWidth      =   10598
   ToolboxBitmap   =   "UcQueue.ctx":0000
   Begin VB.PictureBox picPlace 
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   315
      ScaleHeight     =   28
      ScaleWidth      =   42
      TabIndex        =   15
      Top             =   2085
      Width           =   45
   End
   Begin VB.Timer timerCard 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   900
      Top             =   330
   End
   Begin VB.PictureBox picCallFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   5535
      ScaleHeight     =   4452
      ScaleWidth      =   3738
      TabIndex        =   8
      Top             =   855
      Width           =   3735
      Begin XtremeReportControl.ReportControl rptCallList 
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   7223
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeSuiteControls.ShortcutCaption scCallInf 
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "�����б�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   8421504
      End
   End
   Begin VB.PictureBox picQueueFace 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   495
      ScaleHeight     =   4452
      ScaleWidth      =   4452
      TabIndex        =   1
      Top             =   915
      Width           =   4455
      Begin XtremeReportControl.ReportControl rptQueueList 
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   4260
         _Version        =   589884
         _ExtentX        =   7514
         _ExtentY        =   7858
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnSort =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   3450
         TabIndex        =   14
         Top             =   45
         Width           =   195
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   13
         Top             =   45
         Width           =   195
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1905
         TabIndex        =   12
         Top             =   45
         Width           =   195
      End
      Begin VB.CheckBox optOutQueue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   11
         Top             =   45
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "�Ŷ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   7
         Top             =   30
         Width           =   750
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   255
         Index           =   3
         Left            =   3645
         TabIndex        =   6
         Top             =   30
         Width           =   750
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   255
         Index           =   2
         Left            =   2865
         TabIndex        =   5
         Top             =   30
         Width           =   750
      End
      Begin VB.Label lblQueueFilter 
         BackStyle       =   0  'Transparent
         Caption         =   "����ͣ"
         Height          =   255
         Index           =   1
         Left            =   2130
         TabIndex        =   4
         Top             =   30
         Width           =   750
      End
      Begin XtremeSuiteControls.ShortcutCaption scQueueInf 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "�Ŷ��б�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16744576
         GradientColorDark=   16761024
      End
   End
   Begin VB.TextBox txtLocateValue 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   12.24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7965
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Timer tmrBroadCast 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   120
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   600
      Top             =   45
      _Version        =   589884
      _ExtentX        =   610
      _ExtentY        =   610
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkpMain 
      Bindings        =   "UcQueue.ctx":0312
      Left            =   3840
      Top             =   240
      _Version        =   589884
      _ExtentX        =   432
      _ExtentY        =   406
      _StockProps     =   0
   End
End
Attribute VB_Name = "UcQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Ĭ����������ʽ����
Private Const M_LNG_FORMAT_ORDER_LEN As Long = 20

Private Const M_LNG_ICON_QUEUEING As Long = 8264 '807         '�Ŷ���
Private Const M_LNG_ICON_DIAGNOSE As Long = 3009        '������
Private Const M_LNG_ICON_CALLING As Long = 745          '������
Private Const M_LNG_ICON_CALLED As Long = 732          '�Ѻ���

Private Const M_STR_ALL_POPEDOM As String = "[[#ALLPOPEDOM#]]"


'�Ŷ��б�ĵ�ǰѡ��״̬
Public Enum TQueueFromType
    qftWaitQueue = 0
    qftCalledQueue = 1
    qftFindQueue = 2
End Enum


Public Enum TQueueSelState
    qss�Ŷ��� = 0
    qss����ͣ = 1
    qss������ = 2
    qss����� = 3
End Enum


'�˵��ؼ��Ӷ���
Public Enum TMenuId
    mi��� = conMenu_Queue_PrintNumber
    mi˳�� = conMenu_Queue_CallNext
    miֱ�� = conMenu_Queue_CallThis
    mi�㲥 = conMenu_Queue_Broadcast
    
    mi��� = conMenu_Queue_InsertQueue

    mi���� = conMenu_Queue_RestartQueue
    
    mi��ͣ = conMenu_Queue_Pause
    mi���� = conMenu_Queue_Abandon
    mi�ָ� = conMenu_Queue_Restore
    
    mi���� = conMenu_Queue_RecDiagnose
    mi��� = conMenu_Queue_Finaled
    
    mi��λ = conMenu_Queue_Locate
    mi���� = conMenu_Queue_Find
    
    mi�޸� = conMenu_Queue_Update
    miˢ�� = conMenu_Queue_Refresh
    
    mi���� = conMenu_Queue_Setup
        
End Enum


Private WithEvents mobjMsg As clsQueueMsgCenter
Attribute mobjMsg.VB_VarHelpID = -1

Private mobjQueueManage As clsQueueOperation
Attribute mobjQueueManage.VB_VarHelpID = -1
Private mcnOracle As ADODB.Connection
Private mobjOwner As Object
Private mstrProTag As String


Private mblnIsSelectedCallingList As Boolean    '�Ƿ�ѡ���ѽкŶ���
Private mintWorkType As Integer                 'ҵ������
Private mstrPrivs As String                     'Ȩ���ַ���

Private mstrCustomOrderColName As String        '�Զ��������ֶ�   ������ReportControl�ؼ�
Private mblnIsShowBars As Boolean               '�Ƿ���ʾ������
Private mblnIsShowCalledQueue As Boolean        '�Ƿ���ʾ�Ѻ��ж���

Private mblnAutoComplete As Boolean             '�Զ�����ѽ������
Private mblnShowMySelfCalled As Boolean         '����ʾ�Լ����еĶ�������
Private mblnIsReleationQueueTag As Boolean      '�Ƿ�����Ŷӱ��,Ϊtrueʱ ��������ʾ���ŶӺ���Ϊ �Ŷӱ��+�ŶӺ���

Private mstrFindWay As String

Private mblnInitOk As Boolean                   '�Ƿ��ʼ�����
Private mstrLoginUserName As String             '��¼�û���

Private mstrLocateType As String                '��λ����
Private mlngLocateRowIndex As Long

Private mstrDataFields As String                '���ֶ����������ʾ�ֶ����У����Զ�����
Private mstrDisplayQueueFields As String        '�Ŷ��б�������
Private mstrDisplayCallFields As String         '�����б�������
Private mstrReason As String                    '���ԭ��

Private mstrQueryQueueNames As String           'Ҫ��ѯ��ʾ�Ķ�������
Private mstrGroupField As String                '������ ��Ϊ���򲻽��з���
Private mstrLastFixedQueue As String        '���ķ�������

Private mlngInterval As Long                    '��ѵʱ��

Private mrsVoiceContext As ADODB.Recordset  '�����ŵ��������ݼ�
Private mstrComputerName As String              '���ؼ��������
Private mdtLastVoiceDate As Date

'Private mlngQueueW1 As Long     '������ʾ���
'Private mlngQueueW2 As Long     '���ж�����ʾ���


Private mblnIsFindQueue As Boolean   '�Ƿ�Ϊ���Ҷ���

Private mlngMenuCaptionStyle As Long


'��ˢ����ر�������
Private mlngReadCount As Long
Private mlngStartTime As Long
Private mlngAvgTime As Long



'�����¼�
Public Event OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay, ByRef strCallContext As String, blnCancel As Boolean)
Public Event OnCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay)

Public Event OnPlayVoiceBefore(ByVal lngCallId As Long, ByVal lngQueueId As Long, ByVal strCallContext As String, ByRef blnCancel As Boolean)
Public Event OnPlayVoiceAfter(ByVal lngCallId As Long, ByVal lngQueueId As Long, ByVal strCallContext As String)

Public Event OnWorkBefore(ByVal lngListType As TQueueFromType, ByVal lngListRow As Long, ByVal lngQueueId As Long, ByVal lngOperationType As TOperationType, blnCancel As Boolean)
Public Event OnWorkAfter(ByVal lngQueueId As Long, ByVal strCurQueueName As String, ByVal lngOperationType As TOperationType)

Public Event OnReadBefore(rsDataRow As ADODB.Recordset, ByVal lngListType As TQueueFromType, blnCancel As Boolean)
Public Event OnReadAfter(rsDataRow As ADODB.Recordset, ByVal lngListType As TQueueFromType, objReportRecord As Object)

Public Event OnCreateQueueNo(ByVal lngQueueId As Long, ByVal strQueueName As String, ByRef strQueueNo As String)

'��ѯ�ŶӶ��������¼�
Public Event OnQueryQueueData(rsData As ADODB.Recordset, blnUseCustom As Boolean)

'���Ҷ�������ʱ�������¼�
Public Event OnFindData(ByVal strFindWay As String, ByVal strFindValue As String, txtFind As Object, rsData As ADODB.Recordset, ByRef blnUseCustom As Boolean)
    
'��λ��������ʱ�������¼�
Public Event OnLocateData(ByVal strLocateWay As String, ByVal strLocateValue As String, txtFind As Object, ByRef lngQueueId As Long, ByRef blnUseCustom As Boolean)
    
Public Event OnSelectionChanged(ByVal lngListType As TQueueFromType, ByVal lngQueueId As Long, objQueueList As Object, objReportRow As Object)
Public Event OnItemDblClick(ByVal lngListType As TQueueFromType, ByVal lngQueueId As Long, objReoprtRow As Object, objReportItem As Object)
Public Event OnQueueListChange(ByVal lngListType As TQueueFromType, objQueueList As Object)


Public Event OnCmdBarInit(objCommandBar As Object)
Public Event OnCmdBarUpdate(objComandBarControl As Object)
Public Event OnCmdBarExecute(objComandBarControl As Object, ByRef blnUseCustom As Boolean)

Public Event OnColumnInit(objQueueList As Object, objReportColumn As Object)

Public Event OnQueueListMouseDown(ByVal lngListType As TQueueFromType, Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnQueueListMouseUp(ByVal lngListType As TQueueFromType, Button As Integer, Shift As Integer, X As Long, Y As Long)

Public Event OnGroupHint(ByVal strHintContext As String)
Public Event OnFilter(rsData As ADODB.Recordset, ByRef blnCancel As Boolean, ByRef blnUseCustom As Boolean)
Public Event OnConfigEvent(ByRef blnUseCustom As Boolean)

Public Event OnModifyBefore(ByVal lngListType As TQueueFromType, ByVal lngQueueId As Long, ByRef objInput As Dictionary, ByRef blnCancel As Boolean, ByRef blnUseCustom As Boolean)
Public Event OnModifyAfter(ByVal lngQueueId As Long, objUpdateValue As Dictionary)

Public Event OnMsgRecevie(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset, ByRef blnUseCustom As Boolean)

Private Declare Function GetTickCount Lib "kernel32" () As Long


'***************************************************************
'ֻ�����Զ���


'��ǰʹ�õĶ�������
Property Get CurQueueType() As TQueueFromType
    If mblnIsFindQueue Then
        CurQueueType = qftFindQueue             '���Ҷ���
    Else
        If mblnIsSelectedCallingList = True Then
            CurQueueType = qftCalledQueue       '���ж���
        Else
            CurQueueType = qftWaitQueue         '�ŶӶ���
        End If
    End If
End Property


'��ȡҵ������
Property Get WorkType() As Long
    WorkType = mintWorkType
End Property


'�ŶӴ������
Property Get QueueOper() As clsQueueOperation
    Set QueueOper = mobjQueueManage
End Property

'CommandBar��������
Property Get CmdBar() As Object
    Set CmdBar = cbrMain
End Property


'�Ⱥ�����б����
Property Get WaitQueueList() As Object
    Set WaitQueueList = rptQueueList
End Property

'�Ѻ����б����
Property Get CallQueueList() As Object
    Set CallQueueList = rptCallList
End Property



'***************************************************************
'��д���Զ���


'�Ƿ�����Ŷӱ��
Property Get IsReleationQueueTag() As Boolean
    IsReleationQueueTag = mblnIsReleationQueueTag
End Property

Property Let IsReleationQueueTag(value As Boolean)
    mblnIsReleationQueueTag = value
End Property

'�Զ��������ֶ�
Property Get CustomOrderField() As String
    CustomOrderField = UCase(mstrCustomOrderColName)
End Property

Property Let CustomOrderField(value As String)
    If UCase(value) <> mstrCustomOrderColName Then
        mstrCustomOrderColName = UCase(value)
    End If

End Property

'������
Property Get ReportNum() As String
    ReportNum = mobjQueueManage.ReportNum
End Property

Property Let ReportNum(value As String)
    mobjQueueManage.ReportNum = value
End Property


'���������
Property Get LastFixedQueue() As String
    LastFixedQueue = mstrLastFixedQueue
End Property


Property Let LastFixedQueue(value As String)
    mstrLastFixedQueue = value
End Property


'���Ҵ������������Ĳ��ҷ�ʽ
Property Get FindWayEx() As String
    FindWayEx = mstrFindWay
End Property

Property Let FindWayEx(value As String)
    mstrFindWay = value
End Property


'�Ƿ���ʾ�ŶӽкŹ�������ť
Property Get IsShowBars() As Boolean
    IsShowBars = mblnIsShowBars 'cbrMain.ActiveMenuBar.Visible
End Property

Property Let IsShowBars(value As Boolean)
    mblnIsShowBars = value
    cbrMain.ActiveMenuBar.Visible = value
End Property


'�Ƿ���ʾ�Ѻ����ŶӶ���
Property Get IsShowCalledQueue() As Boolean
    IsShowCalledQueue = mblnIsShowCalledQueue
End Property

Property Let IsShowCalledQueue(value As Boolean)
    mblnIsShowCalledQueue = value
    
    If DkpMain.PanesCount <= 0 Then Exit Property
    
    If value Then
        DkpMain.Panes(2).Closed = False
    Else
        DkpMain.Panes(2).Closed = True
    End If
End Property

'�����ѯ�������ֶ�
Property Get DataFields() As String
    DataFields = UCase(mstrDataFields)
End Property

Property Let DataFields(value As String)
    mstrDataFields = UCase(value)
End Property


'���ú���ʱ��Ŀ�ĵ�
Property Get CalledTarget() As String
    CalledTarget = mobjQueueManage.CallTarget
End Property

Property Let CalledTarget(value As String)
    mobjQueueManage.CallTarget = value
End Property

'�Ŷ��б���ʾ�ֶ�����
Property Get DisplayQueueFields() As String
    DisplayQueueFields = UCase(mstrDisplayQueueFields)
End Property

Property Let DisplayQueueFields(value As String)
    If UCase(value) <> mstrDisplayQueueFields Then
        mstrDisplayQueueFields = UCase(value)
    End If
End Property

'�����б���ʾ�ֶ�����
Property Get DisplayCallFields() As String
    DisplayCallFields = mstrDisplayCallFields
End Property

Property Let DisplayCallFields(value As String)
    If UCase(value) <> mstrDisplayCallFields Then
        mstrDisplayCallFields = UCase(value)
    End If
End Property


'��ѵ���ʱ��(��λ������)
Property Get Interval() As Long
    Interval = mlngInterval
End Property

Property Let Interval(value As Long)
    mlngInterval = value
End Property



'������ʾ���ŶӶ�������  ע�����Ϊ�գ�����ʾ��ǰҵ�������µ����ж����е��ŶӺͺ�������
Property Get QueryQueueNames() As String
    QueryQueueNames = mstrQueryQueueNames
End Property

Property Let QueryQueueNames(value As String)
    mstrQueryQueueNames = value
End Property




'����(��������ݷ�����ʾ)
Property Get GroupField() As String
    GroupField = UCase(mstrGroupField)
End Property

Property Let GroupField(value As String)
    If UCase(value) <> mstrGroupField Then
        mstrGroupField = UCase(value)
    End If
End Property



'������Ч����
Property Get ValidDays() As Long
    ValidDays = mobjQueueManage.ValidDays
End Property

Property Let ValidDays(value As Long)
    mobjQueueManage.ValidDays = value
End Property


'��������
Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Property Set Font(value As StdFont)
    Call SetFont(value)
End Property


'�Ƿ��ͼ��
Property Get IsIconLarge() As Boolean
    IsIconLarge = cbrMain.Options.LargeIcons
End Property


Property Let IsIconLarge(value As Boolean)
    cbrMain.Options.LargeIcons = value
    
    Call cbrMain.RecalcLayout
End Property


'�Ƿ���ʾ��ť�ı�
Property Get IsShowToolText() As Boolean
    IsShowToolText = IIf(mlngMenuCaptionStyle = xtpButtonIcon, False, True)
End Property

Property Let IsShowToolText(value As Boolean)
    Dim i As Integer
    Dim cbrControl As CommandBarControl

    If value = False Then
        '����ʾ�ı�
        mlngMenuCaptionStyle = xtpButtonIcon
        cbrMain(1).ShowTextBelowIcons = False
    Else
        '��ʾ�ı�
        mlngMenuCaptionStyle = xtpButtonIconAndCaption
        cbrMain(1).ShowTextBelowIcons = True
    End If

    For Each cbrControl In cbrMain(1).Controls
        cbrControl.Style = mlngMenuCaptionStyle
    Next

    cbrMain.RecalcLayout
End Property

'���ԭ��
Property Get Reasons() As String
    Reasons = mstrReason
End Property

Property Let Reasons(value As String)
    mstrReason = value
End Property


Private Sub SetFont(ft As StdFont)
'����������ʾ
    Dim ftNew As StdFont
    
    If Font Is Nothing Then Exit Sub
    
    Set ftNew = New StdFont
    Call CopyFont(ft, ftNew)
    

    Set UserControl.Font = ftNew
    Set cbrMain.Options.Font = ftNew
    
    Set rptQueueList.PaintManager.CaptionFont = ftNew
    Set rptQueueList.PaintManager.TextFont = ftNew
    
    Set rptCallList.PaintManager.CaptionFont = ftNew
    Set rptCallList.PaintManager.TextFont = ftNew
    
    Set scQueueInf.Font = ftNew
    Set scCallInf.Font = ftNew

End Sub



Private Sub CopyFont(ftSource As StdFont, ByRef ftTarget As StdFont)
'������������
    
    ftTarget.Bold = ftSource.Bold
    ftTarget.Charset = ftSource.Charset
    ftTarget.Italic = ftSource.Italic
    ftTarget.Name = ftSource.Name
    ftTarget.Size = ftSource.Size
    ftTarget.Strikethrough = ftSource.Strikethrough
    ftTarget.Underline = ftSource.Underline
    ftTarget.Weight = ftSource.Weight
End Sub


Public Sub ApplyVoiceConfig()
'Ӧ����������
    Dim str����վ������ As String
    
    
   '��ȡ�кŷ�ʽ
    If Val(GetSetting("ZLSOFT", gstrRegPath, "���ŷ�ʽ", 1)) Then
         str����վ������ = GetSetting("ZLSOFT", gstrRegPath, "Զ�˺���վ��", "")
         
         If Trim(str����վ������) = "" Then str����վ������ = AnalyseComputer
    Else
        str����վ������ = AnalyseComputer
    End If
    
    mstrComputerName = AnalyseComputer
    
    mobjQueueManage.PlayStation = str����վ������
    mobjQueueManage.LocalStation = mstrComputerName

    mobjQueueManage.PlayTimeLength = Val(GetSetting("ZLSOFT", gstrRegPath, "��������ʱ��", 15))
    mobjQueueManage.PlayCount = Val(GetSetting("ZLSOFT", gstrRegPath, "�������Ŵ���", 2))
    mobjQueueManage.VoiceType = GetSetting("ZLSOFT", gstrRegPath, "��������", "")
    mobjQueueManage.IsPlayHintSound = Val(GetSetting("ZLSOFT", gstrRegPath, "��������ǰ������ʾ��", False))
    mobjQueueManage.PlaySpeed = Val(GetSetting("ZLSOFT", gstrRegPath, "������������", 0))
    mobjQueueManage.UseVbsPlay = IIf(Val(GetSetting("ZLSOFT", gstrRegPath, "����VBS�Զ������", 1)) = 0, False, True)
    mobjQueueManage.CusVoiceScript = GetSetting("ZLSOFT", gstrRegPath, "VBS�ű�", "")
    
    Interval = Val(GetSetting("ZLSOFT", gstrRegPath, "��ѯ���ʱ��", 30))
    
    mblnAutoComplete = Val(GetSetting("ZLSOFT", gstrRegPath, "�Զ�����ѽ������", 1))
    mblnShowMySelfCalled = Val(GetSetting("ZLSOFT", gstrRegPath, "ֻ��ʾ�Լ����еĶ���", 1))
    
    If Val(GetSetting("ZLSOFT", gstrRegPath, "������������", 1)) = 0 Then
        Call StopVoice
    Else
        Call StartVoice
    End If
 
End Sub

Public Function ShowVoiceConfig() As Boolean
'��ʾ��������
    ShowVoiceConfig = frmSetup.ShowMe(Me)
End Function


Public Sub UseMsgCenter(ByVal lngSys As Long, ByVal lngModule As Long, Optional ByVal strPrivs As String = "")
'������Ϣ����
    Call mobjQueueManage.UseMsgCenter(lngSys, lngModule, strPrivs)
    
    Set mobjMsg = gobjMsgCenter
End Sub


'��ʼ������
Public Sub InitQueue( _
    cnOracle As ADODB.Connection, _
    ByVal intWorkType As Integer, _
    ByVal objOwnerForm As Object, _
    ByVal strProTag As String, _
    Optional ByVal strLoginUser As String = "system", _
    Optional ByVal strPrivs As String = "[[#ALLPOPEDOM#]]")
      
    mblnInitOk = False
    
    '���ó�ʼ�����з���
    Call mobjQueueManage.InitQueue(cnOracle, intWorkType, strLoginUser)
    
    '����ȫ�ֱ���
    Set mcnOracle = cnOracle
    Set mobjOwner = objOwnerForm
    
    mstrProTag = strProTag
    mstrPrivs = strPrivs
    mintWorkType = intWorkType
    
    gstrRegPath = "����ģ��\" & mstrProTag & "\�Ŷӽк�"

    mblnIsSelectedCallingList = False
    
    If Trim(mstrDataFields) = "" Then
        mstrDataFields = mobjQueueManage.DefQueryCols
    End If
     
    '��ǰ��½���û���
    mstrLoginUserName = strLoginUser
    
    If DkpMain.PanesCount > 0 Then
        '�������ý���
        DkpMain.CloseAll
        DkpMain.DestroyAll

        Call InitFaceScheme
    End If
    
    Call InitLocalParas                                                 '��ʼ������
    Call SetCommandBarStyle
    Call InitCommandBars                                                '��ʼ��������������ť
    
    Call InitQueueList(rptQueueList, mstrGroupField, mstrCustomOrderColName, mstrDisplayQueueFields, UCase(mstrDataFields))          '��ʼ���ȴ����ж����б�
    Call InitQueueList(rptCallList, mstrGroupField, mstrCustomOrderColName, mstrDisplayCallFields, UCase(mstrDataFields))             '��ʼ���Ѻ����б�

    mblnInitOk = True
    
End Sub


Private Sub InitFaceScheme()
'��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane
    Dim dblQueueRate As Double
    
    On Error GoTo errHandle
    
    With DkpMain
        .SetCommandBars cbrMain
        
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    dblQueueRate = 0.666
    
    If gstrRegPath <> "" Then
        dblQueueRate = Val(GetSetting("ZLSOFT", gstrRegPath, "QueueListWidthRate", "0.6"))
    End If
    
    Set Pane1 = DkpMain.CreatePane(0, dblQueueRate * 100, 2000, DockLeftOf, Nothing)
                
    Pane1.Title = "�Ŷ��б�"
    Pane1.Tag = 0
    Pane1.Handle = picQueueFace.hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
       
    
    Set Pane2 = DkpMain.CreatePane(1, (1 - dblQueueRate) * 100, 2000, DockRightOf, Nothing)
    
    Pane2.Title = "�����б�"
    Pane2.Tag = 1
    Pane2.Handle = picCallFace.hwnd
    Pane2.Options = PaneNoFloatable Or PaneNoCaption

    If mblnIsShowCalledQueue Then
        DkpMain.Panes(2).Closed = False
    Else
        DkpMain.Panes(2).Closed = True
    End If

    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub lblQueueFilter_Click(Index As Integer)
On Error GoTo errHandle
    If optOutQueue(Index).value = 0 Then
        optOutQueue(Index).value = 1
    Else
        optOutQueue(Index).value = 0
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjMsg_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset)
'��Ϣ���մ���(����ֻ�ܽ��յ��Ŷӽк���ص���Ϣ)
    Dim blnUseCustom As Boolean
    Dim strValue As String
    
    '����������������Ϣ
    If strMsgItemIdentity = G_STR_MSG_QUEUE_004 Then Exit Sub
    
    '�ж���Ϣ�еĶ��������Ƿ���Ҫ���д���Ķ�������
    rsData.Filter = "node_name='queue_name'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "��Ϣ��Ч����⵽δ������Ч�Ķ������ƣ���ֹ��Ϣ����"
        Exit Sub
    End If
    
    strValue = Nvl(rsData!node_value)
    
    If InStr(mstrQueryQueueNames, strValue) <= 0 Then
        Debug.Print "����Ϣ�������в����ڵ�ǰҵ����Χ��������Ϣ����"
        Exit Sub
    End If
    
    blnUseCustom = False
    RaiseEvent OnMsgRecevie(strMsgItemIdentity, strXmlContext, rsData, blnUseCustom)
    
    '��������߽�������Ϣ���������ｫֱ���˳�
    If blnUseCustom Then Exit Sub
    
    
    
    Select Case strMsgItemIdentity
        Case G_STR_MSG_QUEUE_001    '�����Ϣ
            Call LineQueueMsgProcess(rsData)
            
        Case G_STR_MSG_QUEUE_002    '�����Ϣ
            '�Ӷ�����ɾ��������ʾ
            Call CompleteMsgProcess(rsData)
            
        Case G_STR_MSG_QUEUE_003    '״̬ͬ����Ϣ
            '����״̬�����б��е�����
            Call StateSyncMsgProcess(rsData)
    End Select
    
End Sub

Private Sub LineQueueMsgProcess(ByVal rsData As ADODB.Recordset)
'�Ŷ���Ϣ����
    Dim lngQueueId As Long
    Dim lngQueueRow As Long
    Dim lngRecordIndex As Long
    Dim objQueueList As ReportControl
    
    '�ж��Ѻ��ж������Ƿ���ڸ����ݣ�������ڣ�����Ҫ����ɾ��
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount > 0 Then lngQueueId = Nvl(rsData!node_value)
    
    Call LocateQueueRow(lngQueueId, objQueueList, lngQueueRow)
    
    If lngQueueRow > 0 Then
        lngRecordIndex = objQueueList.Rows(lngQueueRow).Record.Index
        objQueueList.Rows(lngQueueRow).Selected = False
        
        Call objQueueList.Records.RemoveAt(lngRecordIndex)
        Call objQueueList.Populate
        
        If objQueueList.Rows.Count > lngQueueRow Then
            objQueueList.Rows(lngQueueRow).Selected = True
        End If
    End If
    
    'ˢ���ŶӶ�������
    Call LoadWaitQueueData
End Sub

Private Sub CompleteMsgProcess(ByVal rsData As ADODB.Recordset)
'�����Ϣ�Ĵ������
    Dim lngQueueId As Long
    Dim lngQueueRow As Long
    Dim lngRecordIndex As Long
    Dim objQueueList As ReportControl
    
    '��ȡ��Ϣ�еĶ���ID
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "��Ϣ��Ч����⵽δ������Ч�Ķ���ID����ֹ��Ϣ����"
        Exit Sub
    End If
    
    lngQueueId = Val(Nvl(rsData!node_value))
    
    Call LocateQueueRow(lngQueueId, objQueueList, lngQueueRow)
    
    If lngQueueRow > 0 Then
        lngRecordIndex = objQueueList.Rows(lngQueueRow).Record.Index
        objQueueList.Rows(lngQueueRow).Selected = False
        
        Call objQueueList.Records.RemoveAt(lngRecordIndex)
        Call objQueueList.Populate
        
        If objQueueList.Rows.Count > lngQueueRow Then
            objQueueList.Rows(lngQueueRow).Selected = True
        End If
    End If
    
End Sub


Private Sub StateSyncMsgProcess(ByVal rsData As ADODB.Recordset)
'�Ŷ���Ϣ����
    Dim lngQueueId As Long
    Dim lngCurState As Long
    
    '��ȡ��Ϣ�еĶ���ID
    rsData.Filter = "node_name='queue_id'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "��Ϣ��Ч����⵽δ������Ч�Ķ���ID����ֹ��Ϣ����"
        Exit Sub
    End If
    
    lngQueueId = Val(Nvl(rsData!node_value))
    
    '��ȡ��Ϣ�еĶ���״̬
    rsData.Filter = "node_name='queue_state'"
    If rsData.RecordCount <= 0 Then
        Debug.Print "��Ϣ��Ч����⵽δ������Ч�Ķ���ID����ֹ��Ϣ����"
        Exit Sub
    End If
    
    lngCurState = Val(Nvl(rsData!node_value))
    
    'ˢ�½����е�״̬����
    Call RefreshQueueRowState(lngQueueId, lngCurState)
End Sub

Private Sub optOutQueue_Click(Index As Integer)
On Error GoTo errHandle
'    Dim i As Long
'
'    If Not mblnInitOk Then Exit Sub
'
'    If optOutQueue(Index).value <> 0 Then
'        optOutQueue(Index).Enabled = False
'        lblQueueFilter(Index).FontBold = True
'    End If
'
'    '����δ��ѡ�����ʾ��ʽ
'    For i = 0 To optOutQueue.Count - 1
'        If i <> Index Then
'            If optOutQueue(Index).value <> 0 Then
'                optOutQueue(i).value = 0
'                optOutQueue(i).Enabled = True
'
'                lblQueueFilter(i).FontBold = False
'            End If
'        End If
'    Next i
'
'    If optOutQueue(Index).value = 0 Then Exit Sub

    Call LoadWaitQueueData
    
    '���õ�ǰ�ŶӶ���Ϊ�������
    mblnIsSelectedCallingList = False
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadWaitQueueData()
'����ȴ���������
    Dim rsData As ADODB.Recordset
    
    Set rsData = QueryQueueData()
    
    If rsData Is Nothing Then
        rptQueueList.Records.DeleteAll
        rptQueueList.Populate
        
        Exit Sub
    End If
    
    Call LoadDataToList(rptQueueList, rsData)
    
    rptQueueList.Populate
End Sub

Private Sub LoadCallQueueData()
'������ж�������
    Dim rsData As ADODB.Recordset
    
    Set rsData = QueryQueueData()
    
    If rsData Is Nothing Then
        rptCallList.Records.DeleteAll
        rptCallList.Populate
        
        Exit Sub
    End If
    
    Call LoadDataToList(rptCallList, rsData)

End Sub

Private Sub picCallFace_Resize()
On Error GoTo errHandle
   
    scCallInf.Left = 0
    scCallInf.Top = 0
    scCallInf.Width = picCallFace.Width

    rptCallList.Left = 0
    rptCallList.Top = scCallInf.Height
    rptCallList.Width = picCallFace.ScaleWidth
    rptCallList.Height = picCallFace.ScaleHeight - scCallInf.Height
    
errHandle:
End Sub

Private Sub picQueueFace_Resize()
On Error GoTo errHandle
    
    scQueueInf.Left = 0
    scQueueInf.Top = 0
    scQueueInf.Width = picQueueFace.Width

    rptQueueList.Left = 0
    rptQueueList.Top = scQueueInf.Height
    rptQueueList.Width = picQueueFace.ScaleWidth
    rptQueueList.Height = picQueueFace.ScaleHeight - scQueueInf.Height
    
    
    optOutQueue(TQueueSelState.qss�Ŷ���).Left = scQueueInf.Width - 4000
    optOutQueue(TQueueSelState.qss�Ŷ���).Top = 65

    optOutQueue(TQueueSelState.qss����ͣ).Left = optOutQueue(0).Left + 1000
    optOutQueue(TQueueSelState.qss����ͣ).Top = 65


    optOutQueue(TQueueSelState.qss������).Left = optOutQueue(1).Left + 1000
    optOutQueue(TQueueSelState.qss������).Top = 65


    optOutQueue(TQueueSelState.qss�����).Left = optOutQueue(2).Left + 1000
    optOutQueue(TQueueSelState.qss�����).Top = 65
    
    
    lblQueueFilter(TQueueSelState.qss�Ŷ���).Left = optOutQueue(0).Left + optOutQueue(0).Width + 20
    lblQueueFilter(TQueueSelState.qss�Ŷ���).Top = 55
    
    lblQueueFilter(TQueueSelState.qss����ͣ).Left = optOutQueue(1).Left + optOutQueue(1).Width + 20
    lblQueueFilter(TQueueSelState.qss����ͣ).Top = 55
    
    lblQueueFilter(TQueueSelState.qss������).Left = optOutQueue(2).Left + optOutQueue(2).Width + 20
    lblQueueFilter(TQueueSelState.qss������).Top = 55
    
    lblQueueFilter(TQueueSelState.qss�����).Left = optOutQueue(3).Left + optOutQueue(3).Width + 20
    lblQueueFilter(TQueueSelState.qss�����).Top = 55
    
errHandle:
End Sub


Private Sub SetCommandBarStyle()
On Error GoTo errHandle
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = zlCommFun.GetPubIcons
    
    
    With cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        '.LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
Exit Sub
errHandle:
End Sub


Private Function CheckPopedom(ByVal strFuncName As String) As Boolean
'���Ȩ���Ƿ����
    CheckPopedom = False
    
    If mstrPrivs = M_STR_ALL_POPEDOM Then
        CheckPopedom = True
        Exit Function
    End If
    
    If InStr("," & mstrPrivs & ",", "," & strFuncName & ",") > 0 Then CheckPopedom = True
    If InStr(";" & mstrPrivs & ";", ";" & strFuncName & ";") > 0 Then CheckPopedom = True
    
End Function

Public Sub zlCreateMenuBars(cbrMenuBar As CommandBarPopup, Optional ByVal blnIsAllMenu As Boolean = False)
    Dim cbrControl As CommandBarControl
    
    If cbrMenuBar Is Nothing Then Exit Sub
    
    '�������ú��в˵�
    With cbrMenuBar.CommandBar.Controls
        If CheckPopedom("���") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_PrintNumber, "���"): cbrControl.IconId = 3571
        End If
        
        If CheckPopedom("˳��") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "˳��")
            cbrControl.BeginGroup = True
            cbrControl.IconId = 744
            cbrControl.ToolTipText = "��˳�������һ��"
        End If
        
        If CheckPopedom("ֱ��") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "ֱ��"): cbrControl.IconId = 732
        End If
        

        If CheckPopedom("�㲥") Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "�㲥"): cbrControl.IconId = 2600
        End If
        
        If Not blnIsAllMenu Then
            If CheckPopedom("����") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Filter, "����")
                cbrControl.BeginGroup = True
                cbrControl.IconId = 731
            End If
            
            If CheckPopedom("ˢ��") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "ˢ��"): cbrControl.IconId = 3003
            End If
        End If
    End With
    
    If blnIsAllMenu Then
    
        With cbrMenuBar.CommandBar.Controls
            If CheckPopedom("���") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_InsertQueue, "���")
                cbrControl.IconId = 2600
                cbrControl.BeginGroup = True
            End If
            
            If CheckPopedom("����") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RestartQueue, "����"): cbrControl.IconId = 2614
            End If
            
    
            If CheckPopedom("��ͣ") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Pause, "��ͣ"): cbrControl.IconId = 746: cbrControl.BeginGroup = True
            End If
            
            If CheckPopedom("����") Then    'Ȩ����ʹ�����ŵ�����
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Abandon, "����"): cbrControl.IconId = 8113
            End If
            
            If CheckPopedom("�ָ�") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Restore, "�ָ�"): cbrControl.IconId = 252
            End If
            
            If CheckPopedom("����") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RecDiagnose, "����"): cbrControl.IconId = 8264
            End If
            
            If CheckPopedom("���") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Finaled, "���"): cbrControl.IconId = 747
            End If
            
            If CheckPopedom("����") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Filter, "����"): cbrControl.IconId = 731
            End If
            
            If CheckPopedom("ˢ��") Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "ˢ��"): cbrControl.IconId = 3003
            End If
        End With
    End If
    
    For Each cbrControl In cbrMenuBar.Controls
        If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "QUEUE" '�����ŶӲ˵���ʶ
    Next
End Sub

'��ʼ�����ܰ�ť
Private Sub InitCommandBars()
    Dim cbrToolBar1 As CommandBar
    Dim cbrToolBar2 As CommandBar
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrCustom  As CommandBarControlCustom
    Dim i As Integer
    
    cbrMain(1).Visible = False
    
    For i = cbrMain.Count To 2 Step -1
        cbrMain(i).Controls.DeleteAll
    Next i
    
    Set cbrMain.Icons = zlCommFun.GetPubIcons
    
    
    '�ŶӺ��й���������
    If cbrMain.Count > 1 Then
        Set cbrToolBar1 = cbrMain(2)
    Else
        Set cbrToolBar1 = cbrMain.Add("������", XTPBarPosition.xtpBarTop)
    End If
    
    cbrToolBar1.Closeable = False
    '��CommandBar������ ���ó� ��ͼ�����ı�����ʽ
    cbrToolBar1.ShowTextBelowIcons = True

    With cbrToolBar1.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_PrintNumber, "���"): cbrControl.IconId = 103
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallNext, "˳��"): cbrControl.IconId = 744: cbrControl.ToolTipText = "��˳�������һ��": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_CallThis, "ֱ��"): cbrControl.IconId = 732
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Broadcast, "�غ�"): cbrControl.IconId = 745
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_InsertQueue, "���"): cbrControl.IconId = 2600: cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����Ϊ���Ⱥ���"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RestartQueue, "����"): cbrControl.IconId = 2614: cbrControl.ToolTipText = "���½�������Ŷ�"

        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Pause, "��ͣ"): cbrControl.IconId = 746: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Abandon, "����"): cbrControl.IconId = 8113: cbrControl.ToolTipText = "��������"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Restore, "�ָ�"): cbrControl.IconId = 252: cbrControl.ToolTipText = "�����ݻָ����Ŷ�״̬"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_RecDiagnose, "����"): cbrControl.IconId = 3009: cbrControl.ToolTipText = "�Ա������˽��н��ﴦ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Finaled, "���"): cbrControl.IconId = 747
                
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Filter, "����"): cbrControl.IconId = 731: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Refresh, "ˢ��"): cbrControl.IconId = 791
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Update, "�޸�"): cbrControl.IconId = 3003: cbrControl.ToolTipText = "�޸��Ŷ���Ϣ"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Setup, "����"): cbrControl.IconId = 181: cbrControl.ToolTipText = "��������": cbrControl.BeginGroup = True
        
    End With
    
    Call DoCmdBarInitEvent(cbrToolBar1)
    
    For Each cbrControl In cbrToolBar1.Controls
        If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "QUEUE" '�����ŶӲ˵���ʶ
    Next
        
    cbrToolBar1.Visible = mblnIsShowBars
    txtLocateValue.Visible = False
    
    If CheckPopedom("����") And CheckPopedom("��λ") Then
        
        If cbrMain.Count > 2 Then
            Set cbrToolBar2 = cbrMain(3)
        Else
            Set cbrToolBar2 = cbrMain.Add("������", XTPBarPosition.xtpBarTop)
        End If
        
        cbrToolBar2.Closeable = False
        cbrToolBar2.ShowTextBelowIcons = False
        
        With cbrToolBar2.Controls
            '������λ�Ȳ���
            Set cbrMenuBar = .Add(xtpControlPopup, conMenu_Queue_LocateType, "�ŶӺ�")
                cbrMenuBar.Id = conMenu_Queue_LocateType
                cbrMenuBar.Flags = xtpFlagRightAlign
    
            Set cbrCustom = .Add(xtpControlCustom, conMenu_Queue_LocateValue, "��λ����")
                cbrCustom.Handle = txtLocateValue.hwnd
                cbrCustom.Flags = xtpFlagRightAlign
                cbrCustom.Style = xtpButtonIconAndCaption
    
                txtLocateValue.Visible = True
                
            Set cbrCustom = .Add(xtpControlCustom, conMenu_Queue_LocateValue, "")
                cbrCustom.Handle = picPlace.hwnd
    
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Locate, "")
                cbrControl.Id = conMenu_Queue_Locate
                cbrControl.IconId = 8267
                cbrControl.Flags = xtpFlagRightAlign
                cbrControl.ToolTipText = "��λ�Ŷ�����"
                cbrControl.Checked = True
                
            Set cbrControl = .Add(xtpControlButton, conMenu_Queue_Find, "")
                cbrControl.Id = conMenu_Queue_Find
                cbrControl.IconId = 721
                cbrControl.Flags = xtpFlagRightAlign
                cbrControl.ToolTipText = "�����Ŷ�����"
        End With
        
        For Each cbrControl In cbrToolBar2.Controls
            If (cbrControl.Type = xtpControlButton) Or (cbrControl.Type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
            If cbrControl.Category = "" Then cbrControl.Category = "QUEUE" '�����ŶӲ˵���ʶ
        Next
        
        Call DockRightOfCommandBar(cbrToolBar2, cbrToolBar1)
        
        cbrToolBar2.Visible = mblnIsShowBars
        txtLocateValue.Visible = mblnIsShowBars
    End If
End Sub


Private Sub DockRightOfCommandBar(cbBarToDock As CommandBar, cbBarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbrMain.RecalcLayout
    
    cbBarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    'ʹcbBarToDock������ʼ����ʾ�����ұ�
    cbrMain.DockToolBar cbBarToDock, 300000, (Bottom + Top) / 2, cbBarOnLeft.Position

End Sub

Private Sub DoCmdBarInitEvent(cbrToolBar As Object)
    '������������ʼ���¼�
    RaiseEvent OnCmdBarInit(cbrToolBar)
End Sub


Private Function GetWaitQueueSelState() As TQueueSelState
'��ȡ�ŶӶ��е�������ʾ״̬���Ŷӣ���ͣ�����������
    Dim lngState As TQueueSelState
    Dim strState As String
    
     '�õ���ǰѡ�е��Ŷ�״ֵ̬
    lngState = -1
    
'    If optOutQueue(TQueueSelState.qss�Ŷ���).value Then lngState = TQueueSelState.qss�Ŷ���     '�Ŷ���״̬
'    If optOutQueue(TQueueSelState.qss����ͣ).value Then lngState = TQueueSelState.qss����ͣ     '����ͣ״̬
'    If optOutQueue(TQueueSelState.qss������).value Then lngState = TQueueSelState.qss������     '������״̬
'    If optOutQueue(TQueueSelState.qss�����).value Then lngState = TQueueSelState.qss�����     '�����״̬
    
    If rptQueueList.SelectedRows.Count > 0 Then
        If rptQueueList.SelectedRows(0).GroupRow = False Then
            strState = rptQueueList.SelectedRows(0).Record(GetColIndex("�Ŷ�״̬", rptQueueList)).value
            If strState = "�Ŷ���" Then
                lngState = TQueueSelState.qss�Ŷ���
            ElseIf strState = "����ͣ" Then
                lngState = TQueueSelState.qss����ͣ
            ElseIf strState = "������" Then
                lngState = TQueueSelState.qss������
            Else
                lngState = TQueueSelState.qss�����
            End If
            
            If lngState = -1 And strState <> "�����" Then lngState = TQueueSelState.qss�Ŷ���    '��δѡ���κ�״̬ʱ����ֻ�����Ŷ��е�����
        End If
    End If
    
    GetWaitQueueSelState = lngState
End Function


Private Function QueryQueueData() As ADODB.Recordset
'��ѯ�������ݵ����ݼ�
    Dim strSql As String
    Dim strTemp As String
    Dim blnUseCustom As Boolean
    Dim lngTimePoint As Long
    Dim strStartTime As String
    Dim strEndTime As String
    Dim blnHasQueueCol As Boolean
    Dim strSelectColumns As String
    Dim strOrderCondition As String
    Dim strCurQueryQueueNames As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim dtNow As Date
    
    Set QueryQueueData = Nothing
    
    blnUseCustom = False
    RaiseEvent OnQueryQueueData(rsData, blnUseCustom)
    
    '��δʹ���Զ����ѯʱ��blnUseCustomΪfalse������ʹ��ϵͳĬ�ϵĲ�ѯ����Դ
    If Not blnUseCustom Then
        lngTimePoint = Val(Format(Time, "h"))
        dtNow = zlDatabase.Currentdate
        If lngTimePoint <= 4 Then
            strStartTime = To_Date(Format(dtNow - 1, "yy-mm-dd 20:00:00"))
            strEndTime = To_Date(Format(dtNow, "yy-mm-dd 08:00:00"))
        Else
            strStartTime = To_Date(Format(dtNow, "yy-mm-dd 00:00:00"))
            strEndTime = To_Date(Format(dtNow, "yy-mm-dd 23:59:59"))
        End If
    
       '���õõ������ֶ��з���
        strSelectColumns = mobjQueueManage.DefQueryCols
        strOrderCondition = mobjQueueManage.CustomOrder
    
        '������������ӵ����ţ��Ա��ڲ�ѯSQL���
        strCurQueryQueueNames = Replace(mstrQueryQueueNames, ",", "','")
        strTemp = IIf(strCurQueryQueueNames = "", "", "and �������� in ('" & strCurQueryQueueNames & "') ")
        strSql = "select " & strSelectColumns & " from �ŶӽкŶ��� where �Ŷ�ʱ�� between " & strStartTime & " and " & strEndTime & " and ҵ������=" & mintWorkType & " " & strTemp & _
                IIf(strOrderCondition <> "", " order by " & strOrderCondition, "")
       
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ��������")
    End If

    If rsData Is Nothing Then
        MsgBox "���ص����ݼ�����Ϊ�գ�����ִ�����ݼ��ز�����", vbInformation, "�Ŷӽк�ϵͳ"
        Exit Function
    End If
    
    '�ж��Ƿ�����Ŷ�id��
    blnHasQueueCol = False
    'rsData����ʱ��������Ŷ�ID����
    For i = 0 To rsData.Fields.Count - 1
        If UCase(rsData.Fields.Item(i).Name) = "ID" Then
            blnHasQueueCol = True
            Exit For
        End If
    Next i


    '���ݼ��в������Ŷ�ID����ִ�м������ݲ��� ֱ���˳�����
    If Not blnHasQueueCol Then
        MsgBox "��ѯ���ݼ��в������Ŷ�ID������ִ�����ݼ��ز�����", vbInformation, "�Ŷӽк�ϵͳ"
        Exit Function
    End If
    
    Set QueryQueueData = rsData
    
End Function

Private Sub LoadDataToList(objCurQueueList As ReportControl, rsData As ADODB.Recordset, Optional ByVal blSetFocus As Boolean = True)
'�����������
'lngQueueType:0�ŶӶ��У�1���ж���
'blSetFocus �Ƿ��ֹ�����б��㣬Ĭ�� True
On Error GoTo errHandle
'�������ݵ��б�
    Dim rptRecord As ReportRecord
    Dim lngQueueLoadModle As TQueueSelState
    Dim i As Long
    Dim lngCurSelRow As Long
'    Dim lngQueryQueueState As Long
    Dim blnCancel As Boolean
    Dim lngOrdIndex As Long
    Dim strFilter As String
    Dim blnIsWaitQueue As Boolean
    Dim blnLoadData As Boolean
    Dim strQueueState As String
    
    objCurQueueList.Records.DeleteAll
    If rsData Is Nothing Then
        objCurQueueList.Populate
        Exit Sub
    End If
    
    blnIsWaitQueue = IIf(objCurQueueList.Name = rptQueueList.Name, True, False)
    
    If blnIsWaitQueue Then
        'lngQueueLoadModle = GetWaitQueueSelState()
'        lngQueryQueueState = Decode(lngQueueLoadModle, _
'                            TQueueSelState.qss�Ŷ���, TQueueState.qsQueueing, _
'                            TQueueSelState.qss������, TQueueState.qsAbstain, _
'                            TQueueSelState.qss����ͣ, TQueueState.qsPause, _
'                            TQueueState.qsComplete)
                  
        If optOutQueue(0).value = 1 Then
            strQueueState = "�Ŷ�״̬ = 0"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If optOutQueue(1).value = 1 Then
            strQueueState = "�Ŷ�״̬ = 3"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If optOutQueue(2).value = 1 Then
            strQueueState = "�Ŷ�״̬ = 2"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If optOutQueue(3).value = 1 Then
            strQueueState = "�Ŷ�״̬ = 4"
            
            If strFilter <> "" Then strFilter = strFilter & " Or "
            strFilter = strFilter & strQueueState
        End If
        
        If strFilter = "" Then strFilter = "�Ŷ�״̬=0 or �Ŷ�״̬=2 or �Ŷ�״̬=3 or �Ŷ�״̬=4"
    Else
        '���˳������У��Ѻ��У������У������е�����
        strFilter = "�Ŷ�״̬=1 or �Ŷ�״̬=7 or �Ŷ�״̬=8 or �Ŷ�״̬=9"
    End If
    
    rsData.Filter = strFilter
    
    '���û�����ݣ���ֱ���˳�
    If rsData.RecordCount <= 0 Then
        objCurQueueList.Populate
        Exit Sub
    End If
    
    '�õ���ǰ�����б�Ľ�������
    If objCurQueueList.SelectedRows.Count > 0 Then lngCurSelRow = objCurQueueList.SelectedRows(0).Index


    lngOrdIndex = GetColIndex("ORD", objCurQueueList)
    While Not rsData.EOF
    
        blnCancel = False
        RaiseEvent OnReadBefore(rsData, IIf(blnIsWaitQueue, TQueueFromType.qftWaitQueue, TQueueFromType.qftCalledQueue), blnCancel)
        
        If Not blnCancel Then
            blnLoadData = False
            
            '�����Ŷ�״̬Ϊ��ǰ��ѡ״̬������
'            If (Nvl(rsData("�Ŷ�״̬"), -1) = lngQueryQueueState And blnIsWaitQueue) _
'                Or ((Nvl(rsData("�Ŷ�״̬"), -1) = TQueueState.qsCalling _
'                Or Nvl(rsData("�Ŷ�״̬"), -1) = TQueueState.qsCalled _
'                Or Nvl(rsData("�Ŷ�״̬"), -1) = TQueueState.qsDiagnose _
'                Or Nvl(rsData("�Ŷ�״̬"), -1) = TQueueState.qsWaitCall) And Not blnIsWaitQueue) Then
                
                If mblnShowMySelfCalled Then
                    blnLoadData = IIf(UCase(Nvl(rsData!����ҽ��)) = UCase(mstrLoginUserName) Or Nvl(rsData!����ҽ��) = "", True, False)
                Else
                    blnLoadData = True
                End If
'            End If
            

            If blnLoadData Then
                Set rptRecord = objCurQueueList.Records.Add
                
                For i = 0 To objCurQueueList.Columns.Count - 1
                    rptRecord.AddItem ""
                Next
    
                Call SetReportRecordItem(objCurQueueList, rptRecord, rsData)
                
                '���ڵ�ʹ�����ݿ�Ĭ�ϵ�����ʱ���ܹ������������������������ݶ�������Ŷӽ���ؼ�����������
                If lngOrdIndex >= 0 Then
                    rptRecord.Item(lngOrdIndex).value = Format(rsData.AbsolutePosition, "00000000")
                End If
   
                '���ñ�����ɫ
                lngQueueLoadModle = Decode(Nvl(rsData("�Ŷ�״̬"), -1), _
                                    TQueueState.qsQueueing, TQueueSelState.qss�Ŷ���, _
                                    TQueueState.qsAbstain, TQueueSelState.qss������, _
                                    TQueueState.qsPause, TQueueSelState.qss����ͣ, _
                                    TQueueState.qsComplete, TQueueSelState.qss�����)
                                    
                Select Case lngQueueLoadModle
                    Case TQueueSelState.qss����ͣ
                        Call SetReportRecordColor(objCurQueueList, rptRecord, vbYellow)
                    Case TQueueSelState.qss������
                        Call SetReportRecordColor(objCurQueueList, rptRecord, &H8080FF)
                    Case TQueueSelState.qss�����
                        Call SetReportRecordColor(objCurQueueList, rptRecord, &HFF00&)
                    Case TQueueSelState.qss�Ŷ���
                        Call SetReportRecordColor(objCurQueueList, rptRecord, vbWhite)
                End Select

                RaiseEvent OnReadAfter(rsData, IIf(blnIsWaitQueue, TQueueFromType.qftWaitQueue, TQueueFromType.qftCalledQueue), rptRecord)
            End If
                        
        End If

        rsData.MoveNext
    Wend

    objCurQueueList.Populate
    
    '�ָ�ѡ����Ŷ�����
    If lngCurSelRow >= objCurQueueList.Rows.Count Then
        lngCurSelRow = IIf(objCurQueueList.Rows.Count <= 0, -1, rptQueueList.Rows.Count - 1)
    End If

    '103315 �������������������������ý����� ���²���ȡ������
    If  blSetFocus Then
        If lngCurSelRow > -1 Then
            objCurQueueList.Rows(lngCurSelRow).Selected = True
            Set objCurQueueList.FocusedRow = objCurQueueList.Rows(lngCurSelRow)
        End If
    End If
    
    '�ָ�����������ﲻ����������ֱ������˳���Լ��ָ��Ŷ�״̬�󣬶������ݿ��ܲ��ᰴ˳����ʾ
    objCurQueueList.SortOrder(objCurQueueList.SortOrder.Count - 1).SortAscending = True

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetColumnWidth(ByVal strColPros As String, strColName As String) As Long
'��ȡ�еĿ��
On Error GoTo errHandle
    Dim strColPro As String
    Dim lngColIndex As Long
    
    GetColumnWidth = 100
    
    If gstrRegPath = "" Then Exit Function
    
    lngColIndex = InStr(strColPros, strColName & ":")
    If lngColIndex <= 0 Then Exit Function
    
    strColPro = Mid(strColPros, lngColIndex, 255)
    strColPro = Replace(strColPro, strColName & ":", "")
    
    GetColumnWidth = Val(strColPro)
    
Exit Function
errHandle:
    GetColumnWidth = 100
End Function


Public Function GetValidCols(ByVal strCols As String, Optional ByVal strQueueTabPrefix As String) As String
'����߱��Ĳ�ѯ���ֶΣ�ID,��������,ҵ��ID,��������,�Ŷ�״̬,�Ŷ����,�ŶӺ���
    Dim strResult As String
    Dim strTabPrefix As String
    
    strResult = UCase(strCols)
    strTabPrefix = UCase(strQueueTabPrefix)
    
    If Trim(strResult) = "" Then
        GetValidCols = mobjQueueManage.GetAllQueueTabCols(strQueueTabPrefix)
        Exit Function
    End If
    
    strResult = ",," & strResult & ",,"
    
    strResult = Replace(strResult, ", ", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "ID,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "��������,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "ҵ��ID,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "��������,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "�Ŷ�״̬,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "�Ŷ����,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "�ŶӺ���,", ",")
    strResult = Replace(strResult, "," & IIf(strTabPrefix <> "", strTabPrefix & ".", "") & "����ҽ��,", ",")
    strResult = Replace(strResult, ",,", "")
    
    strResult = "[[TAB]].ID,[[TAB]].��������,[[TAB]].ҵ��ID,[[TAB]].��������,[[TAB]].�Ŷ�״̬, [[TAB]].�Ŷ����,[[TAB]].�ŶӺ���,[[TAB]].����ҽ�� " & IIf(strResult = "", "", "," & strResult)
    strResult = Replace(strResult, "[[TAB]].", IIf(strTabPrefix <> "", strTabPrefix & ".", ""))
    
    GetValidCols = strResult
End Function


Private Sub InitQueueList(objQueueList As Object, ByVal strGroupCol As String, ByVal strOrderCols As String, _
    ByVal strDisplayCols As String, ByVal strDataCols As String)
'��ʼ���Ŷ��б�

    Dim Column As ReportColumn
    Dim strAllColNames As String
    Dim strQueueColNames() As String
    Dim strCallColNames() As String
    Dim strOrderCondition() As String
    Dim blnIsOrders As Boolean  '�Ƿ������������ֶ�
    Dim i As Integer
    Dim j As Integer
    Dim aryDisplayCols() As String
    Dim strOrderCol As String
    Dim blnIsConfigOrder As Boolean
    Dim aryCurOrderInf() As String
    Dim objCurQueueList As ReportControl
    Dim lngColIndex As Long
    Dim strColPros As String
    
    Err = 0: On Error Resume Next
    
    If objQueueList Is Nothing Then Exit Sub
    Set objCurQueueList = objQueueList
    

    '��ʼ���ŶӶ�����ʾ�ֶ�
    Call objCurQueueList.Records.DeleteAll
    Call objCurQueueList.Columns.DeleteAll
    
    Set objCurQueueList.Icons = zlCommFun.GetPubIcons

    '��ʼ���б��������
    objCurQueueList.AllowColumnRemove = False
    objCurQueueList.ShowItemsInGroups = False
    objCurQueueList.SkipGroupsFocus = True
    objCurQueueList.MultipleSelection = False

    With objCurQueueList.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "���б����϶�����,�ɰ����з���..."
        .NoItemsText = "û�п���ʾ����Ŀ..."
        .VerticalGridStyle = xtpGridSolid
    End With
    
    '��ȡ�Ŷӽк�Ĭ�ϵĲ�ѯ�����У�Ĭ��Ϊ���е������ֶΣ����ظ�ʽΪ�ԡ��������ŷָ����ֶ�����
    strAllColNames = strDataCols
    
    '���û�в�ѯ���ֶΣ����˳�
    If Trim(strAllColNames) = "" Then Exit Sub
    If Trim(strDisplayCols) = "" Then strDisplayCols = strAllColNames
    
    '�����б���ʾ�ֶ�
    If Trim(strAllColNames) <> "" And strAllColNames <> "*" Then
    
        strQueueColNames() = Split(strAllColNames, ",")
        aryDisplayCols() = Split(strDisplayCols, ",")
        
        lngColIndex = 0
        
        strColPros = GetSetting("ZLSOFT", gstrRegPath, objQueueList.Name)
        
        For i = LBound(aryDisplayCols) To UBound(aryDisplayCols)
            If Trim(aryDisplayCols(i)) <> "" Then
                '������Ҫ��ʾ�������ֶ�
                Set Column = objCurQueueList.Columns.Add(lngColIndex, aryDisplayCols(i), GetColumnWidth(strColPros, aryDisplayCols(i)), True)
                lngColIndex = lngColIndex + 1
                
                '�жϸ����Ƿ�������
                If InStr("," & strGroupCol & ",", "," & aryDisplayCols(i) & ",") > 0 Then
                    Column.Groupable = True
                    
                    '���������в�������ʾ
                    Column.Visible = False
                End If
                
                RaiseEvent OnColumnInit(objCurQueueList, Column)
                
            End If
        Next i

        For i = LBound(strQueueColNames) To UBound(strQueueColNames)
            If Trim(strQueueColNames(i)) <> "" And InStr(strDisplayCols, strQueueColNames(i)) <= 0 Then
                '���벻��Ҫ��ʾ���ֶ�
                Set Column = objCurQueueList.Columns.Add(lngColIndex, strQueueColNames(i), 100, True)
                lngColIndex = lngColIndex + 1
                
                If InStr("," & strGroupCol & ",", "," & strQueueColNames(i) & ",") > 0 Then
                    Column.Groupable = True
                End If
                
                Column.Visible = False
            End If
        Next i
        
        '����δ�����Զ��������Ӷκ�ķ�������
        Set Column = objCurQueueList.Columns.Add(lngColIndex, "ORD", 0, False)
        Column.Visible = False
    End If
    
    '���û��ʹ���Զ���������ؼ���������������������ʱʹ�����ݿ��Ĭ������
    blnIsConfigOrder = False
    If Trim(strOrderCols) <> "" Then
        aryCurOrderInf = Split(strOrderCols, ",")
        blnIsConfigOrder = True
    End If
    
    

    '��������Լ�����Ĺ���
    With objCurQueueList
    
        If Trim(strGroupCol) <> "" Then
            .GroupsOrder.DeleteAll

            'ֻ����������һ���ֶν��з���
            For i = 0 To .Columns.Count
                If .Columns(i).Caption = strGroupCol Then
                    .GroupsOrder.Add .Columns(i)
                    Exit For
                End If
            Next i

            '����֮��,��������в���ʾ,�����е������ǲ����
            .GroupsOrder(0).SortAscending = True ' False '
        End If
        
         '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.DeleteAll
        .AllowColumnSort = False 'blnIsConfigOrder
        
        
        If blnIsConfigOrder Then
            '���������ֶ�
            For i = LBound(aryCurOrderInf) To UBound(aryCurOrderInf)
                strOrderCol = Trim(aryCurOrderInf(i))
                
                If strOrderCol <> "" Then
                    If InStr(strOrderCol, "DESC") > 0 Then
                        '����������
                        strOrderCol = Replace(strOrderCol, "DESC", "")
    
                        For j = 0 To .Columns.Count - 1
                            If .Columns(j).Caption = strOrderCol Then
                                .SortOrder.Add .Columns(j)
                                .SortOrder(.SortOrder.Count - 1).SortAscending = False
                                Exit For
                            End If
                        Next j
                    Else
                        '����������
                        strOrderCol = Replace(strOrderCol, "ASC", "")
                        
                        For j = 0 To .Columns.Count - 1
                            If .Columns(j).Caption = strOrderCol Then
                                .SortOrder.Add .Columns(j)
                                .SortOrder(.SortOrder.Count - 1).SortAscending = True
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Next i
        Else
            '���û�����������У���������ݿ�ļ���˳���������
            .SortOrder.Add .Columns(GetColIndex("ORD", objCurQueueList))
            .SortOrder(.SortOrder.Count - 1).SortAscending = True
        End If
        
    End With
    
End Sub

Private Function FormatQueueOrder(ByVal strQueueOrder As String) As String
'��ʽ���Ŷ����
    Dim strQueueNum As String
    Dim strQueueNumFront As String
    Dim strQueueNumBehind As String
    
    FormatQueueOrder = ""
    strQueueNum = strQueueOrder
    
    '�������С��������Ҫ��ȡ����
    If InStr(strQueueNum, ".") > 0 Then
        strQueueNumFront = Mid(strQueueNum, 1, InStr(strQueueNum, ".") - 1)
        strQueueNumBehind = Mid(strQueueNum, InStr(strQueueNum, "."), Len(strQueueNum))
        
        strQueueNumFront = Replace(Space(M_LNG_FORMAT_ORDER_LEN - Len(strQueueNumFront)), " ", "0") & strQueueNumFront
    Else
        strQueueNum = Replace(Space(M_LNG_FORMAT_ORDER_LEN - Len(strQueueNum)), " ", "0") & strQueueNum
    End If
    
    '���ش���������
    FormatQueueOrder = IIf(InStr(strQueueNum, ".") > 0, strQueueNumFront & strQueueNumBehind, strQueueNum)
End Function


Private Sub SetReportRecordItem(rptControl As ReportControl, rptItems As ReportRecord, rsData As ADODB.Recordset)
'����ReportControl�ؼ�������
    On Error GoTo errHandle
    Dim i As Long
    Dim j As Long
    Dim intMaxNumLen As Integer
    Dim strQueueNum As String
    Dim strQueueNumFront As String
    Dim strQueueNumBehind As String
    Dim strValue As String
    Dim lngFitColWidth As Long
    Dim lngQueueState As Long
    Dim lngColIndex As Long
    Dim lngFirstCol As Long
    
    lngQueueState = Val(Nvl(rsData!�Ŷ�״̬))
    lngFirstCol = GetFirstDisplayColIndex(rptControl)
    
    'ѭ�����ظ�����Ԫ������
    For i = 0 To rptControl.Columns.Count - 1
        lngColIndex = rptControl.Columns(i).ItemIndex
        
        '��ReportControl�ؼ��ķ��������������ԣ���Ҫ�ԡ��Ŷ���š�����Ӧ�ַ����Ĵ���������ȷ�Ľ�������
        If Not HasField(rsData, rptControl.Columns(i).Caption) Then
            rptItems(lngColIndex).value = ""
        Else
            strValue = Nvl(rsData("" & rptControl.Columns(i).Caption & "").value)
            
            If (Trim(rptControl.Columns(i).Caption) = "�ŶӺ���" Or Trim(rptControl.Columns(i).Caption) = "�ŶӺ�") And mblnIsReleationQueueTag Then
                strValue = Nvl(rsData!�Ŷӱ��) & strValue
            End If
            
            '�ж��Ƿ��ǡ��Ŷ���š��У�������봦������ֱ�Ӽ���
            If rptControl.Columns(i).Caption = "�Ŷ����" Then
            
                strQueueNum = FormatQueueOrder(strValue)
                
                rptItems(lngColIndex).value = strQueueNum
            ElseIf rptControl.Columns(i).Caption = "�Ŷ�״̬" Then
            
                Select Case Val(rsData!�Ŷ�״̬)
                    Case TQueueState.qsQueueing
                        rptItems(lngColIndex).value = "�Ŷ���"
                    Case TQueueState.qsPause
                        rptItems(lngColIndex).value = "����ͣ"
                    Case TQueueState.qsAbstain
                        rptItems(lngColIndex).value = "������"
                    Case TQueueState.qsComplete
                        rptItems(lngColIndex).value = "�����"
                    Case TQueueState.qsCalled
                        rptItems(lngColIndex).value = "�Ѻ���"
                    Case TQueueState.qsCalling
                        rptItems(lngColIndex).value = "������"
                    Case TQueueState.qsDiagnose
                        rptItems(lngColIndex).value = "������"
                    Case TQueueState.qsWaitCall
                        rptItems(lngColIndex).value = "������"
                End Select
                
            ElseIf rptControl.Columns(i).Caption = "��������" Then
                If InStr(Nvl(rsData!��������), IIf(Trim(mstrLastFixedQueue) <> "", mstrLastFixedQueue, "@<A...B  C.#.D>")) > 0 Then
                    If rptControl.GroupsOrder.Count > 0 Then
                        If rptControl.GroupsOrder(0).SortAscending = True Then
                            '�����������������
                            rptItems(lngColIndex).value = " " & mstrLastFixedQueue
                        Else
                            '��������н�������
                            rptItems(lngColIndex).value = Chr(255) & mstrLastFixedQueue
                        End If
                    End If
                Else
                    rptItems(lngColIndex).value = strValue
                End If
                
            Else
                If IsDate(strValue) Then strValue = Format(strValue, "yyyy-mm-dd hh:mm:ss")
                rptItems(lngColIndex).value = strValue
                
            End If
    
            '���û�������ͼ��
            If lngColIndex = lngFirstCol Then
                If lngQueueState = TQueueState.qsDiagnose Then
                    rptItems(lngColIndex).Icon = M_LNG_ICON_DIAGNOSE
                ElseIf lngQueueState = TQueueState.qsCalling Then
                    rptItems(lngColIndex).Icon = M_LNG_ICON_CALLING
                ElseIf lngQueueState = TQueueState.qsCalled Then
                    rptItems(lngColIndex).Icon = M_LNG_ICON_CALLED
                Else
                    rptItems(lngColIndex).Icon = M_LNG_ICON_QUEUEING
                End If
            End If
            
            rptItems(i).BackColor = vbWhite
        End If
        
    Next i

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Function GetFirstDisplayColIndex(objQueueList As Object) As Long
'ȡ�õ�һ����ʾ���Ŷ���
    Dim i As Long
    
    GetFirstDisplayColIndex = -1
    
    If Nvl(objQueueList.Tag, "") <> "" Then
        GetFirstDisplayColIndex = Val(Nvl(objQueueList.Tag))
        Exit Function
    End If
    
    For i = 0 To objQueueList.Columns.Count - 1
        If objQueueList.Columns(i).Visible Then
            GetFirstDisplayColIndex = objQueueList.Columns(i).ItemIndex
            objQueueList.Tag = objQueueList.Columns(i).ItemIndex
            Exit Function
        End If
    Next i
End Function


Private Sub SetReportRecordColor(objQueueList As Object, rrRow As ReportRecord, ByVal lngColor As Long)
'�����еı�����ɫ
    Dim i As Long
    
    For i = 0 To objQueueList.Columns.Count - 1
        rrRow.Item(objQueueList.Columns(i).ItemIndex).BackColor = lngColor
    Next i
End Sub

Private Sub SetReportRecordBold(objQueueList As Object, rrRow As ReportRecord, ByVal blnBold As Boolean)
'�����е�����Ӵ�
    Dim i As Long
    
    For i = 0 To objQueueList.Columns.Count - 1
        rrRow.Item(objQueueList.Columns(i).ItemIndex).Bold = blnBold
    Next i
End Sub


Public Sub RefreshQueueData(Optional ByVal blSetFocus As Boolean = True )
'blSetFocus �Ƿ��ֹ�����б��㣬Ĭ�� True
'ˢ���ŶӶ�������
On Error GoTo errHandle
    Dim rsData As ADODB.Recordset
    
    Set rsData = QueryQueueData()
    If rsData Is Nothing Then Exit Sub

    Call LoadDataToList(rptQueueList, rsData, blSetFocus)
    Call LoadDataToList(rptCallList, rsData, blSetFocus)
    
    '�ָ������б�
    If mblnIsSelectedCallingList Then
        Call SwitchActiveWindow(mblnIsSelectedCallingList)
    Else
        Call SwitchActiveWindow(mblnIsSelectedCallingList)
    End If
    
    Call ConfigQueueStateSel(True)
    mblnIsFindQueue = False

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub RefreshQueueRowData(ByVal lngQueueId As Long, ByVal strColName As String, ByVal strValue As String)
'ˢ��������
    Dim lngRowIndex As Long
    Dim lngColIndex As Long
    Dim objList As ReportControl
    
    Call LocateQueueRow(lngQueueId, objList, lngRowIndex)
    If lngRowIndex < 0 Then
        Exit Sub
    End If
    
    '������
    lngColIndex = GetColIndex(strColName, objList)
    If lngColIndex < 0 Then Exit Sub
    
    objList.Rows(lngRowIndex).Record(lngColIndex).value = strValue
End Sub


Public Sub RefreshQueueRowState(ByVal lngQueueId As Long, ByVal lngCurState As TQueueState)
'ˢ�¶���������
    Dim lngRowIndex As Long
    Dim objList As ReportControl
    
    lngRowIndex = -1
    
    Call LocateQueueRow(lngQueueId, objList, lngRowIndex)
    If lngRowIndex < 0 Then
        '������ݲ����ڣ���ˢ����ʾ
        Call RefreshQueueData(False)
        Exit Sub
    End If
 
    If mblnIsFindQueue Then
        '����ǲ��Ҷ��У�����Ҫ���¶�Ӧ�����ݵ���ʾ״̬
        Call SetQueueRowState(objList, lngRowIndex, lngCurState)
        Call objList.Populate
    Else
        If objList.Name = rptQueueList.Name Then
            '����б�ѡ���״̬�����ݵ�ǰ״̬��ͬ����ɾ��
            If GetWaitQueueSelState() <> lngCurState Then
                Call DelQueueRecord(qftWaitQueue, lngRowIndex)
                
                '�����ǰ״̬Ϊ�����У�����Ҫˢ�º��ж��н���������ʾ
                If lngCurState = qsWaitCall Then
                    Call LoadCallQueueData
                End If
            Else
                Call SetQueueRowState(objList, lngRowIndex, lngCurState)
                Call objList.Populate
            End If
        Else
            '������Ѻ��ж��У���ֱ�Ӹ�����״̬
            If lngCurState <> qsCalled And lngCurState <> qsCalling And lngCurState <> qsDiagnose Then
                '����Ѿ������ں���״̬�������ɾ��
                Call DelQueueRecord(qftCalledQueue, lngRowIndex)
                
                If GetWaitQueueSelState() = lngCurState Then
                    Call LoadWaitQueueData
                End If
            Else
                Call SetQueueRowState(objList, lngRowIndex, lngCurState)
                Call objList.Populate
            End If
        End If
    End If
End Sub

Private Sub LocateQueueRow(ByVal lngQueueId As Long, objList As ReportControl, ByRef lngRow As Long)
'���ݶ���ID��λ��
    Dim i As Long
    Dim lngColIndex As Long
    
    lngRow = -1
    Set objList = Nothing
    
    '���ŶӶ��п�ʼ����
    lngColIndex = GetColIndex("ID", rptQueueList)
    
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow = False Then
            If rptQueueList.Rows(i).Record(lngColIndex).value = lngQueueId Then
                lngRow = rptQueueList.Rows(i).Index
                Exit For
            End If
        End If
    Next i
    
    If lngRow <> -1 Then
        Set objList = rptQueueList
        Exit Sub
    End If
    
    '�Ӻ��ж��п�ʼ����
    lngColIndex = GetColIndex("ID", rptCallList)
    
    For i = 0 To rptCallList.Rows.Count - 1
        If rptCallList.Rows(i).GroupRow = False Then
            If rptCallList.Rows(i).Record(lngColIndex).value = lngQueueId Then
                lngRow = rptCallList.Rows(i).Index
                Exit For
            End If
        End If
    Next i
    
    If lngRow <> -1 Then Set objList = rptCallList
End Sub

Private Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function


Private Sub DoCmdBarExecute(Control As XtremeCommandBars.ICommandBarControl, ByRef blnUseCustom As Boolean)
    RaiseEvent OnCmdBarExecute(Control, blnUseCustom)
End Sub


Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnUseCustom As Boolean
    
    blnUseCustom = False
    
    Call DoCmdBarExecute(Control, blnUseCustom)
    
    '���ʹ�����Զ����¼�������ִ�к���Ĳ���
    If blnUseCustom Then Exit Sub
    
    Select Case Control.Id
        Case conMenu_Queue_LocateType * 10# + 1 To conMenu_Queue_LocateType * 10# + 50   '��λ
            Call Menu_View_Locate_Type_click(Control)

        Case conMenu_Queue_PrintNumber   '���
            Call comMenu_���

        Case conMenu_Queue_CallNext      '˳��
            Call comMenu_˳��

        Case conMenu_Queue_CallThis      'ֱ��
            Call comMenu_ֱ��

        Case conMenu_Queue_Broadcast     '�㲥
            Call comMenu_�㲥

        Case conMenu_Queue_InsertQueue   '���
            Call comMenu_���
            
        Case conMenu_Queue_RestartQueue     '����
            Call comMenu_����

        Case conMenu_Queue_RecDiagnose   '����
            Call comMenu_����

        Case conMenu_Queue_Pause         '��ͣ
            Call comMenu_��ͣ

        Case conMenu_Queue_Abandon       '����
            Call comMenu_����

        Case conMenu_Queue_Restore       '�ָ�
            Call comMenu_�ָ�

        Case conMenu_Queue_Finaled       '���
            Call comMenu_���

        Case conMenu_Queue_Filter        'ˢ��
            Call comMenu_����
            
        Case conMenu_Queue_Refresh       'ˢ��
            Call comMenu_ˢ��

        Case conMenu_Queue_Update        '�޸�
            Call comMenu_�޸�

        Case conMenu_Queue_Setup         '����
            Call comMenu_����
            
        Case conMenu_Queue_Locate       '��λ
            Call SetLocateState(Control, True)
            
            Call LocateQueueData(mstrLocateType, txtLocateValue.Text)
        Case conMenu_Queue_Find          '����
            Call SetLocateState(Control, False)

            Call FindQueueData(mstrLocateType, txtLocateValue.Text)

    End Select
End Sub


'��ťִ���¼�
Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Call zlExecuteCommandBars(Control)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub SetLocateState(Control As XtremeCommandBars.ICommandBarControl, ByVal blnIsLocate As Boolean)
    Dim objFindControl As XtremeCommandBars.ICommandBarControl
    
    Set objFindControl = cbrMain.FindControl(, IIf(blnIsLocate, conMenu_Queue_Find, conMenu_Queue_Locate), True, True)
    If Not objFindControl Is Nothing Then
        objFindControl.Checked = False
    End If
    
    Control.Checked = True
End Sub


Private Function IsFindModel() As Boolean
'�ж��Ƿ�Ϊ����ģʽ
    Dim cbrFind As CommandBarControl
    
    IsFindModel = False
    
    Set cbrFind = cbrMain.FindControl(, conMenu_Queue_Find, True, True)
    
    If cbrFind Is Nothing Then Exit Function

    IsFindModel = cbrFind.Checked
End Function

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim mintQueueState As Integer
    
    '��ȡ�ŶӶ��е���ʾ״̬
    mintQueueState = GetWaitQueueSelState()

    If Not mblnInitOk Then Exit Sub
    
    Select Case Control.Id
        Case conMenu_Queue_LocateType   '��λ����ҵ�����
          Control.Visible = True
          Control.Enabled = True
          Control.Caption = mstrLocateType
            
        Case conMenu_Queue_LocateType * 10# + 1 To conMenu_Queue_LocateType * 10# + 50
          Control.Visible = True
          Control.Checked = (InStr(Control.Caption, mstrLocateType) > 0)
            
        Case conMenu_Queue_PrintNumber  '���
          Control.Visible = CheckPopedom("���")
          Control.Enabled = Trim(mobjQueueManage.ReportNum) <> ""
          
          
          
        Case conMenu_Queue_CallNext     '˳��           'ֻ�д����Ŷ��е����ݣ����ܽ���˳������
          Control.Visible = CheckPopedom("˳��")
          Control.Enabled = Not mblnIsSelectedCallingList And mintQueueState = TQueueSelState.qss�Ŷ��� And Not mblnIsFindQueue
          
        Case conMenu_Queue_CallThis     'ֱ��           '�Ŷ��б��е����ݣ������Խ���ֱ������
          Control.Visible = CheckPopedom("ֱ��")
          Control.Enabled = Not mblnIsSelectedCallingList And mintQueueState = TQueueSelState.qss�Ŷ��� And Not mblnIsFindQueue
          
        Case conMenu_Queue_Broadcast    '�㲥           'ֻ�к��к�����ݲ��ܽ��й㲥
          Control.Visible = CheckPopedom("�㲥")
          Control.Enabled = mblnIsSelectedCallingList And rptCallList.SelectedRows.Count > 0
          
          
          
          
          
        Case conMenu_Queue_InsertQueue  '���
          Control.Visible = CheckPopedom("���")
          Control.Enabled = Not mblnIsSelectedCallingList And mintQueueState = TQueueSelState.qss�Ŷ��� And Not mblnIsFindQueue
          
        Case conMenu_Queue_RestartQueue '����
          Control.Visible = CheckPopedom("����")
          Control.Enabled = mintQueueState <> -1
          
        Case conMenu_Queue_RecDiagnose  '����
          Control.Visible = CheckPopedom("����")
          Control.Enabled = mblnIsSelectedCallingList Or mblnIsFindQueue Or (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss����� And mintQueueState <> -1 And mintQueueState <> TQueueSelState.qss�Ŷ���)
          
        Case conMenu_Queue_Pause        '��ͣ
          Control.Visible = CheckPopedom("��ͣ")
          Control.Enabled = (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss����ͣ And mintQueueState <> -1) Or mblnIsFindQueue Or mblnIsSelectedCallingList
          
        Case conMenu_Queue_Abandon      '����           '�Ѻ������ݿ��Խ�������
          Control.Visible = CheckPopedom("����")
          Control.Enabled = (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss������ And mintQueueState <> -1) Or mblnIsFindQueue Or mblnIsSelectedCallingList
          
        Case conMenu_Queue_Restore      '�ָ�
          Control.Visible = CheckPopedom("�ָ�")
          Control.Enabled = mblnIsSelectedCallingList Or (mintQueueState <> TQueueSelState.qss�Ŷ��� And mintQueueState <> -1 And Not mblnIsSelectedCallingList) Or mblnIsFindQueue
          
        Case conMenu_Queue_Finaled      '���
          Control.Visible = CheckPopedom("���")
          Control.Enabled = mblnIsSelectedCallingList Or mblnIsFindQueue Or (Not mblnIsSelectedCallingList And mintQueueState <> TQueueSelState.qss����� And mintQueueState <> -1 And mintQueueState <> TQueueSelState.qss�Ŷ���) '


          
        Case conMenu_Queue_Filter       '����
          Control.Visible = CheckPopedom("����")
          Control.Enabled = True
          
'        Case conMenu_Queue_Refresh      'ˢ��
'          Control.Visible = CheckPopedom("ˢ��")
'          Control.Enabled = True
          
        Case conMenu_Queue_Locate       '��λ
          Control.Visible = CheckPopedom("��λ")
          Control.Enabled = True
          
        Case conMenu_Queue_Find         '����
          Control.Visible = CheckPopedom("����")
          Control.Enabled = True
          
        Case conMenu_Queue_Update       '�޸�
          Control.Visible = CheckPopedom("�޸�")
          Control.Enabled = True
        
        Case conMenu_Queue_Setup        '����
          Control.Visible = CheckPopedom("����") Or CheckPopedom("��������")
          Control.Enabled = True
    End Select
    
    Call DoCmdBarUpdate(Control)
End Sub

'��ť�����¼�
Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next
    Call zlUpdateCommandBars(Control)
    
    If mblnIsFindQueue Then
        scQueueInf.Caption = "��ѯ�����"
    Else
        scQueueInf.Caption = "�Ŷ��б�"
    End If
Err.Clear
End Sub

Private Sub DoCmdBarUpdate(Control As XtremeCommandBars.ICommandBarControl)
    RaiseEvent OnCmdBarUpdate(Control)
End Sub


Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    Dim aryKindInfo() As String
    
On Error Resume Next
    If CommandBar.Parent Is Nothing Then Exit Sub
    

    Select Case CommandBar.Parent.Id
        Case conMenu_Queue_LocateType
            With CommandBar.Controls
                If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                    mstrFindWay = Replace(mstrFindWay, "����", "")
                    mstrFindWay = Replace(mstrFindWay, "�ŶӺ�", "")
                    
                    mstrFindWay = "�ŶӺ�,����," & mstrFindWay
                    aryKindInfo = Split(mstrFindWay, ",")
                    
                    For i = 0 To UBound(aryKindInfo)
                        If Trim(aryKindInfo(i)) <> "" Then
                            Set objControl = .Add(xtpControlButton, conMenu_Queue_LocateType * 10# + i + 1, aryKindInfo(i) & "(&" & IIf(i >= 9, Chr(65 + i - 9), i + 1) & ")"): objControl.Category = "CallFind"
                            If i = 0 Then objControl.Checked = True
                        End If
                    Next i
                End If
            End With
    End Select
End Sub


Private Sub Menu_View_Locate_Type_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next

    mstrLocateType = Split(Control.Caption, "(")(0)
    Call SaveSetting("ZLSOFT", gstrRegPath, "��λ��ʽ", mstrLocateType)
    
    cbrMain.RecalcLayout

    txtLocateValue.Text = ""
    txtLocateValue.PasswordChar = ""
    txtLocateValue.SetFocus
End Sub


Private Sub SwitchActiveWindow(ByVal blnIsCalledList As Boolean)
On Error Resume Next

    If blnIsCalledList Then
        scCallInf.GradientColorDark = &HFFC0C0
        scCallInf.GradientColorLight = &HFF8080

        scQueueInf.GradientColorDark = &HC0C0C0
        scQueueInf.GradientColorLight = &H808080
    Else
        scQueueInf.GradientColorDark = &HFFC0C0
        scQueueInf.GradientColorLight = &HFF8080

        scCallInf.GradientColorDark = &HC0C0C0
        scCallInf.GradientColorLight = &H808080
    End If
    
Err.Clear
End Sub

Public Sub comMenu_���()
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim strQueueName As String
    
    '��ȡ�ɺ������ݵ�������
    If CurQueueType = qftFindQueue Or CurQueueType = qftWaitQueue Then
        lngRowIndex = GetWaitQueueIndex()
    Else
        lngRowIndex = GetCalledQueueIndex()
    End If
           
    If lngRowIndex < 0 Then
        MsgBox "û�пɹ���ӡ�Ķ������ݣ���ˢ�º����ԡ�", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "��������")
    
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otPrintNo) = False Then Exit Sub

    '���ô�Ź���
    If mobjQueueManage.PrintQueueNo(lngQueueId) = False Then Exit Sub

    Call DoWorkAfter(lngQueueId, strQueueName, otPrintNo)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Private Function GetWindowCaption() As String
'��ȡ���ڱ���
    GetWindowCaption = mobjOwner.Caption
End Function

Private Sub CopyWaitRowToCallRow(ByVal lngWaitRow As Long)
'���ŶӶ��и������ݵ����ж���
    Dim rptRecord As ReportRecord
    Dim objReportRecordItem As ReportRecordItem
    Dim lngWaitColIndex As Long
    Dim lngFirstColIndex As Long
    Dim i As Long
    
    Set rptRecord = rptCallList.Records.Insert(0)
    lngFirstColIndex = GetFirstDisplayColIndex(rptCallList)
    
    '����Ĭ�ϵĿ�����
    For i = 0 To rptCallList.Columns.Count - 1
        Call rptRecord.AddItem("")
    Next i
    
    For i = 0 To rptCallList.Columns.Count - 1
        '���Һ��ж��ж�Ӧ���Ŷ�������
        lngWaitColIndex = GetColIndex(rptCallList.Columns(i).Caption, rptQueueList)
        
        If lngWaitColIndex >= 0 Then
            rptRecord(rptCallList.Columns(i).ItemIndex).value = rptQueueList.Rows(lngWaitRow).Record(lngWaitColIndex).value
            
            If rptQueueList.Rows(lngWaitRow).Record(lngWaitColIndex).Icon > 0 Then
                rptRecord(lngFirstColIndex).Icon = rptQueueList.Rows(lngWaitRow).Record(lngWaitColIndex).Icon
            End If
        End If
    Next
    
    rptCallList.Populate
End Sub


Private Sub CopyCallRowToWaitRow(ByVal lngCallRow As Long)
'�Ӻ��ж��и������ݵ��ŶӶ���
    Dim rptRecord As ReportRecord
    Dim objReportRecordItem As ReportRecordItem
    Dim lngCallColIndex As Long
    Dim lngFirstColIndex As Long
    Dim i As Long
    
    
    Set rptRecord = rptQueueList.Records.Insert(0)
    lngFirstColIndex = GetFirstDisplayColIndex(rptQueueList)
    
    '����Ĭ�ϵĿ�����
    For i = 0 To rptQueueList.Columns.Count - 1
        Call rptRecord.AddItem("")
    Next i
    
    For i = 0 To rptQueueList.Columns.Count - 1
        lngCallColIndex = GetColIndex(rptQueueList.Columns(i).Caption, rptCallList)
        
        If lngCallColIndex >= 0 Then
            rptRecord(rptQueueList.Columns(i).ItemIndex).value = rptCallList.Rows(lngCallRow).Record(lngCallColIndex).value
            
            If rptCallList.Rows(lngCallRow).Record(lngCallColIndex).Icon > 0 Then
                rptRecord(lngFirstColIndex).Icon = rptCallList.Rows(lngCallRow).Record(lngCallColIndex).Icon
            End If
        End If
    Next
    
    rptQueueList.Populate
End Sub


Private Function CheckIsQueueing(ByVal lngQueueId As Long) As Boolean
'�ж������Ƿ����Ŷ���
    Dim lngCurQueueState As Long
    
    lngCurQueueState = mobjQueueManage.GetQueueState(lngQueueId)
    
    CheckIsQueueing = IIf(lngCurQueueState = 0, True, False)
End Function

Public Function GetColumnIndex(ByVal lngQueueFromType As TQueueFromType, ByVal strColName As String) As Long
'ȡ�ö�Ӧ�б��ָ��������
    If lngQueueFromType = qftFindQueue Or lngQueueFromType = qftWaitQueue Then
        GetColumnIndex = GetColIndex(strColName, rptQueueList)
    Else
        GetColumnIndex = GetColIndex(strColName, rptCallList)
    End If
End Function


Public Function GetRowIndex(ByVal lngQueueFromType As TQueueFromType, _
                            ByVal strColName As String, ByVal strValue As String) As Long
'��ȡ��Ӧֵ���ڵ���
    Dim objList As ReportControl
    Dim lngColIndex As Long
    Dim i As Long
    
    GetRowIndex = -1
    
    If lngQueueFromType = qftFindQueue Or lngQueueFromType = qftWaitQueue Then
        Set objList = rptQueueList
        lngColIndex = GetColIndex(strColName, rptQueueList)
    Else
        Set objList = rptCallList
        lngColIndex = GetColIndex(strColName, rptCallList)
    End If
    
    For i = 0 To objList.Rows.Count - 1
        If objList.Rows(i).GroupRow = False Then
            If objList.Rows(i).Record(lngColIndex).value = strValue Then
                GetRowIndex = objList.Rows(i).Index
                Exit Function
            End If
        End If
    Next i
    
End Function

Public Function GetCalledQueueIndex() As Long
'��ȡ�Ѻ��ж��е���ѡ������
    Dim lngCalledQueueRowIndex As Long
    
    lngCalledQueueRowIndex = -1
    GetCalledQueueIndex = -1
    
    If rptCallList.SelectedRows.Count <= 0 Then Exit Function
    
    If rptCallList.SelectedRows(0).GroupRow <> True Then
        lngCalledQueueRowIndex = rptCallList.SelectedRows(0).Index
    Else
        lngCalledQueueRowIndex = rptCallList.SelectedRows(0).Childs(0).Index
    End If
    
    GetCalledQueueIndex = lngCalledQueueRowIndex
End Function

Public Function GetWaitQueueIndex() As Long
'ȡ��ֱ��������
    Dim lngCallRowIndex As Long
    
    lngCallRowIndex = -1
    GetWaitQueueIndex = -1
    
    If rptQueueList.SelectedRows.Count <= 0 Then Exit Function
    
    If rptQueueList.SelectedRows(0).GroupRow <> True Then
        lngCallRowIndex = rptQueueList.SelectedRows(0).Index
    Else
        lngCallRowIndex = rptQueueList.SelectedRows(0).Childs(0).Index
    End If
    
    GetWaitQueueIndex = lngCallRowIndex
End Function


Public Sub DelQueueRecord(ByVal lngQueueFromType As TQueueFromType, ByVal lngRowIndex As Long)
'ɾ�����м�¼����
    Dim lngRecordIndex As Long
    Dim objQueueList As ReportControl
    
    Select Case lngQueueFromType
        Case qftFindQueue, qftWaitQueue
            Set objQueueList = rptQueueList
        Case qftCalledQueue
            Set objQueueList = rptCallList
    End Select
    
    lngRecordIndex = objQueueList.Rows(lngRowIndex).Record.Index
    objQueueList.Rows(lngRowIndex).Selected = False
    
    Call objQueueList.Records.RemoveAt(lngRecordIndex)
    Call objQueueList.Populate
    
    If objQueueList.Rows.Count > lngRowIndex Then
        objQueueList.Rows(lngRowIndex).Selected = True
    End If
End Sub


Public Function GetListValue(ByVal lngQueueFromType As TQueueFromType, ByVal lngRowIndex As Long, ByVal strColName As String) As String
'��ȡ�б����ж�Ӧ��ֵ
    Dim objCurQueueList As ReportControl
    Dim lngColIndex As Long
    
    GetListValue = ""
    
    Select Case lngQueueFromType
        Case TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue
            Set objCurQueueList = rptQueueList
            lngColIndex = GetColIndex(strColName, rptQueueList)
        Case Else
            Set objCurQueueList = rptCallList
            lngColIndex = GetColIndex(strColName, rptCallList)
    End Select
    
    If objCurQueueList.Rows(lngRowIndex).GroupRow = True Then Exit Function
    
    GetListValue = objCurQueueList.Rows(lngRowIndex).Record(lngColIndex).value
End Function


Public Sub SetListValue(ByVal lngQueueFromType As TQueueFromType, ByVal lngRowIndex As Long, _
                            ByVal strColName As String, ByVal strValue As String)
'�����б����ж�Ӧ��ֵ
    Dim objCurQueueList As ReportControl
    Dim lngColIndex As Long
    
    lngColIndex = -1
    
    Select Case lngQueueFromType
        Case TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue
            Set objCurQueueList = rptQueueList
            lngColIndex = GetColIndex(strColName, rptQueueList)
        Case Else
            Set objCurQueueList = rptCallList
            lngColIndex = GetColIndex(strColName, rptCallList)
    End Select
    
    If lngColIndex < 0 Then Exit Sub
    
    objCurQueueList.Rows(lngRowIndex).Record(lngColIndex).value = strValue
End Sub


Public Sub Populate(Optional ByVal lngQueueFromType As TQueueFromType = -1)
'�����б���ʾ
     If lngQueueFromType = -1 Or lngQueueFromType = qftCalledQueue Then rptCallList.Populate
     If lngQueueFromType = -1 Or lngQueueFromType <> qftCalledQueue Then rptQueueList.Populate
End Sub

Public Function GetOrderCallIndex() As Long
'ȡ�õ�ǰ��ѡ�����µĵ�һ���ɹ����еļ�¼������
'˳����ʱ��ʹ�ô˷������λ�ȡ�����е��Ŷ�����

    Dim i As Long
    Dim lngQueueId As Long
    Dim lngCallRowIndex As Long         '������������
    Dim strQueueName As String
    Dim lngQueueNameColIndex As Long
    Dim lngRowIndex As Long
    Dim lngRecordIndex As Long
    
    
    lngCallRowIndex = -1
    GetOrderCallIndex = -1
    
    strQueueName = ""
    lngQueueNameColIndex = GetColIndex("��������", rptQueueList)
    
    '�ж��Ƿ����ŶӼ�¼��ѡ�У�������ڣ���ʹ��ѡ�еĶ���
    If rptQueueList.SelectedRows.Count > 0 Then
        'ѡ��ļ�¼��Ϊ������
        If rptQueueList.SelectedRows(0).GroupRow <> True Then
            strQueueName = rptQueueList.SelectedRows(0).Record(lngQueueNameColIndex).value
        Else
            strQueueName = rptQueueList.SelectedRows(0).Childs(0).Record(lngQueueNameColIndex).value
        End If
    Else
        If rptQueueList.Rows.Count <= 0 Then
            Exit Function
        End If
    End If
    
    lngRowIndex = 0
    
    '���û�б�ѡ�еļ�¼�����ȡ��һ�����еĵ�һ����¼
    Do While rptQueueList.Rows.Count > 0 And lngRowIndex < rptQueueList.Rows.Count
        If rptQueueList.Rows(lngRowIndex).GroupRow = True Then
            If rptQueueList.Rows(lngRowIndex).Childs(0).Record(lngQueueNameColIndex).value = strQueueName Or strQueueName = "" Then
                lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Childs(0).Record(GetColIndex("ID", rptQueueList)).value)
                
                '�жϸ������Ƿ��ܹ����к���
                If CheckIsQueueing(lngQueueId) Then
'                    '���ﲻ��������ɾ������ֻ�гɹ����к󣬲�ɾ������
'                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Childs(0).Record.Index
'                    Call DelQueueRecord(rptQueueList, lngRecordIndex)
                    
                    lngCallRowIndex = rptQueueList.Rows(lngRowIndex).Childs(0).Index
                    
                    If rptQueueList.Rows.Count - 1 >= lngRowIndex Then
                        If rptQueueList.Rows(lngRowIndex).Childs.Count > 0 Then
                            rptQueueList.Rows(lngRowIndex).Childs(0).Selected = True
                        End If
                    Else
                        If rptQueueList.Rows.Count > 0 Then
                            rptQueueList.Rows(rptQueueList.Rows.Count - 1).Selected = True
                        End If
                    End If
                    
                    Exit Do
                Else
                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Childs(0).Index
                    Call DelQueueRecord(qftWaitQueue, lngRecordIndex)
                End If
            Else
                lngRowIndex = lngRowIndex + 1
            End If
        Else
            If rptQueueList.Rows(lngRowIndex).Record(lngQueueNameColIndex).value = strQueueName Or strQueueName = "" Then
                lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
                
                
                '�ж������Ƿ��ܹ����к���
                If CheckIsQueueing(lngQueueId) Then
'                    '���ﲻ��������ɾ������ֻ�гɹ����к󣬲�ɾ������
'                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Record.Index
'                    Call DelQueueRecord(rptQueueList, lngRecordIndex)
                    
                    lngCallRowIndex = rptQueueList.Rows(lngRowIndex).Index
                    
                    If rptQueueList.Rows.Count - 1 >= lngRowIndex Then
                        rptQueueList.Rows(lngRowIndex).Selected = True
                    Else
                        If rptQueueList.Rows.Count > 0 Then
                            rptQueueList.Rows(rptQueueList.Rows.Count - 1).Selected = True
                        End If
                    End If
                    
                    Exit Do
                Else
                    lngRecordIndex = rptQueueList.Rows(lngRowIndex).Record.Index
                    Call DelQueueRecord(qftWaitQueue, lngRecordIndex)
                End If
            Else
                lngRowIndex = lngRowIndex + 1
            End If
        End If
    Loop
    
    GetOrderCallIndex = lngCallRowIndex
End Function


Private Sub rptCallList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnIsFindQueue Then Exit Sub
    
    If mblnIsSelectedCallingList = True Then
        RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
        Exit Sub
    End If

    mblnIsSelectedCallingList = True
    
    '���б���ı�����������б�������ʾ��ʽ
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Call DoQueueListChangeEvent(TQueueFromType.qftCalledQueue, rptCallList)
    
    RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
End Sub

Private Sub rptCallList_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
'����OnGroupHint�¼�
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim objReportHitTest As ReportHitTestInfo
    Dim lngRowCount As Long
    Dim strGroupName As String
    
    Set objReportRow = Nothing
    
    If Not mblnInitOk Then Exit Sub
    
    lngRowCount = 0
    
    Set objReportHitTest = rptCallList.HitTest(X, Y)
    If Not objReportHitTest Is Nothing Then
        Set objReportRow = objReportHitTest.Row
        
        If Not objReportRow Is Nothing Then
            If objReportRow.GroupRow <> True Then
                lngRowCount = objReportRow.ParentRow.Childs.Count
                strGroupName = objReportRow.Record(GetColIndex("��������", rptCallList)).value
            Else
                '����Ƿ��飬��ֱ�ӻ�ȡ�����µ�������
                lngRowCount = objReportRow.Childs.Count
                strGroupName = objReportRow.Childs(0).Record(GetColIndex("��������", rptCallList)).value
            End If
        End If
    End If
    
    rptCallList.ToolTipText = IIf(lngRowCount <= 0, "", "[" & strGroupName & "] �Ѻ�������Ϊ��" & lngRowCount)
    RaiseEvent OnGroupHint(rptCallList.ToolTipText)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptCallList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    RaiseEvent OnQueueListMouseUp(CurQueueType, Button, Shift, X, Y)
End Sub

Private Sub rptCallList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim lngQueueId As Long
    
    '����Ƕ������ݲ���״̬��������б���ʹ��
    If mblnIsFindQueue Then Exit Sub
    
    lngQueueId = Row.Record(GetColIndex("ID", rptCallList)).value
    
    RaiseEvent OnItemDblClick(qftCalledQueue, lngQueueId, Row, Item)
End Sub

Private Sub rptQueueList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnIsSelectedCallingList = False Then
        RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
        Exit Sub
    End If
    
    mblnIsSelectedCallingList = False
    
    '���б���ı�����������б�������ʾ��ʽ
    Call SwitchActiveWindow(mblnIsSelectedCallingList)
    
    Call DoQueueListChangeEvent(IIf(mblnIsFindQueue, TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue), rptQueueList)
    
    RaiseEvent OnQueueListMouseDown(CurQueueType, Button, Shift, X, Y)
End Sub


Private Sub DoQueueListChangeEvent(ByVal lngListType As TQueueFromType, objQueueList As Object)
    '�����б��л��¼�
    RaiseEvent OnQueueListChange(lngListType, objQueueList)
End Sub



Private Sub rptQueueList_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
'����OnGroupHint�¼�
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim objReportHitTest As ReportHitTestInfo
    Dim lngRowCount As Long
    Dim strGroupName As String
    
    Set objReportRow = Nothing
    
    If Not mblnInitOk Then Exit Sub
    
    lngRowCount = 0
    
    Set objReportHitTest = rptQueueList.HitTest(X, Y)
    If Not objReportHitTest Is Nothing Then
        Set objReportRow = objReportHitTest.Row
        
        If Not objReportRow Is Nothing Then
            If objReportRow.GroupRow <> True Then
                lngRowCount = objReportRow.ParentRow.Childs.Count
                strGroupName = objReportRow.Record(GetColIndex("��������", rptQueueList)).value
            Else
                '����Ƿ��飬��ֱ�ӻ�ȡ�����µ�������
                lngRowCount = objReportRow.Childs.Count
                strGroupName = objReportRow.Childs(0).Record(GetColIndex("��������", rptQueueList)).value
            End If
        End If
    End If
    
    rptQueueList.ToolTipText = IIf(lngRowCount <= 0, "", "[" & strGroupName & "] ʣ���Ŷ�����Ϊ��" & lngRowCount)
    RaiseEvent OnGroupHint(rptQueueList.ToolTipText)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub rptQueueList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    RaiseEvent OnQueueListMouseUp(CurQueueType, Button, Shift, X, Y)
End Sub

Private Sub rptQueueList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim lngQueueId As Long
    
    lngQueueId = Row.Record(GetColIndex("ID", rptQueueList)).value
    
    RaiseEvent OnItemDblClick(IIf(mblnIsFindQueue, qftFindQueue, qftWaitQueue), lngQueueId, Row, Item)
End Sub

Private Sub rptQueueList_SelectionChanged()
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim lngQueueId As Long
    
    Set objReportRow = Nothing
    
    If Not mblnInitOk Then Exit Sub
    
    If mblnIsSelectedCallingList <> False Then
        Call DoQueueListChangeEvent(IIf(mblnIsFindQueue, TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue), rptQueueList)
    End If
    
    lngQueueId = 0
    
    If rptQueueList.SelectedRows.Count > 0 Then
        Set objReportRow = rptQueueList.SelectedRows(0)
        
        If objReportRow.GroupRow <> True Then
            lngQueueId = objReportRow.Record(GetColIndex("ID", rptQueueList)).value
        End If
    End If
    
    RaiseEvent OnSelectionChanged(IIf(mblnIsFindQueue, TQueueFromType.qftFindQueue, TQueueFromType.qftWaitQueue), lngQueueId, rptQueueList, objReportRow)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub rptCallList_SelectionChanged()
On Error GoTo errHandle
    Dim objReportRow As ReportRow
    Dim lngQueueId As Long

    If Not mblnInitOk Then Exit Sub
    If mblnIsFindQueue Then Exit Sub
    
    Set objReportRow = Nothing
    
    If mblnIsSelectedCallingList <> True Then
        Call DoQueueListChangeEvent(TQueueFromType.qftCalledQueue, rptCallList)
    End If

    If rptCallList.SelectedRows.Count > 0 Then
        Set objReportRow = rptCallList.SelectedRows(0)
        
        If objReportRow.GroupRow <> True Then
            lngQueueId = objReportRow.Record(GetColIndex("ID", rptCallList)).value
        End If
        
    End If
    
    RaiseEvent OnSelectionChanged(TQueueFromType.qftCalledQueue, lngQueueId, rptCallList, objReportRow)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_˳��()
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim blnCancel As Boolean
    Dim strCallContext As String
    Dim lngSendHwnd As Long
    Dim lngSendResult As Long
    
    '��ȡ�ɺ������ݵ�������
    lngRowIndex = GetOrderCallIndex()
           
    If lngRowIndex < 0 Then
        MsgBox "û�пɹ����еĶ������ݣ���ˢ�º����ԡ�", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)

    strCallContext = ""
    
    RaiseEvent OnCallPreBefore(lngQueueId, TCallWay.cwOrder, strCallContext, blnCancel)
    If blnCancel = True Then Exit Sub
    
'    If mobjQueueManage.CallTarget <> "" Then
'        '���ú��к����ڵ�����Ŀ�ĵ�
'        Call mobjQueueManage.WriteTarget(lngQueueId)
'    End If
    
    'ִ�к��д���
    If mobjQueueManage.SpecifiedCall(lngQueueId, strCallContext) <= 0 Then Exit Sub
    
    'ˢ���Ժ��ж�������
    Call SetQueueRowState(rptQueueList, lngRowIndex, qsWaitCall)
    Call SetListValue(qftWaitQueue, lngRowIndex, "����ҽ��", mstrLoginUserName)
    Call SetListValue(qftWaitQueue, lngRowIndex, "����ʱ��", Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss"))
    
    Call CopyWaitRowToCallRow(lngRowIndex)
    Call rptCallList.Populate
    
    'ɾ���Ѿ����еļ�¼
    Call DelQueueRecord(qftWaitQueue, lngRowIndex)
    
    RaiseEvent OnCallPreAfter(lngQueueId, TCallWay.cwOrder)
    
    '���к�֪ͨzlQueueShow���Ա�����ʾ�ж�ҳ����ʱ����λ����ǰ���в��ˡ������:85290
    lngSendHwnd = FindWindow(vbNullString, "�Ŷ���ʾ����")
    
    If lngSendHwnd > 0 Then
        lngSendResult = PostMessage(lngSendHwnd, 1025, lngQueueId, 0)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function DoCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay, ByRef strCallContext As String) As Boolean
'ִ��OnCallPreBefore�¼�
    Dim blnCancel As Boolean
    
    DoCallPreBefore = True
    blnCancel = False
    RaiseEvent OnCallPreBefore(lngQueueId, lngCallWay, strCallContext, blnCancel)
    
    DoCallPreBefore = Not blnCancel
End Function


Private Sub DoCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As TCallWay)
'ִ��OnCallPreAfter�¼�
    RaiseEvent OnCallPreAfter(lngQueueId, lngCallWay)
End Sub


Public Sub comMenu_ֱ��()
'������ֱ�Ӻ��з���
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim strCallContext As String
    Dim lngSendHwnd As Long
    Dim lngSendResult As Long
   
    lngRowIndex = GetWaitQueueIndex()
    
    If lngRowIndex < 0 Then
        MsgBox "û����Ҫ���еĶ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
    If Not CheckIsQueueing(lngQueueId) Then
        MsgBox "��ǰ�����ѱ����У���ˢ�º����ԡ�", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    '�����¼�
    strCallContext = ""
    If DoCallPreBefore(lngQueueId, TCallWay.cwSpecify, strCallContext) = False Then Exit Sub
       
'    If mobjQueueManage.CallTarget <> "" Then
'        '���ú��к����ڵ�����Ŀ�ĵ�
'        Call mobjQueueManage.WriteTarget(lngQueueId)
'    End If
    
    'ִ�к��д���
    If mobjQueueManage.SpecifiedCall(lngQueueId, strCallContext) <= 0 Then Exit Sub
    
    'ˢ���Ѻ����б�������ʾ
    Call SetQueueRowState(rptQueueList, lngRowIndex, qsWaitCall)
    Call SetListValue(qftWaitQueue, lngRowIndex, "����ҽ��", mstrLoginUserName)
    Call SetListValue(qftWaitQueue, lngRowIndex, "����ʱ��", Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss"))
    
    Call CopyWaitRowToCallRow(lngRowIndex)
    
    Call rptCallList.Populate
    
    'ɾ���Ѿ������е�������
    Call DelQueueRecord(qftWaitQueue, lngRowIndex)
    
    Call DoCallPreAfter(lngQueueId, TCallWay.cwSpecify)
    
    '���к�֪ͨzlQueueShow���Ա�����ʾ�ж�ҳ����ʱ����λ����ǰ���в��ˡ������:85290
    lngSendHwnd = FindWindow(vbNullString, "�Ŷ���ʾ����")
    
    If lngSendHwnd > 0 Then
        lngSendResult = PostMessage(lngSendHwnd, 1025, lngQueueId, 0)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_�㲥()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strCallContext As String
    
    If mblnIsFindQueue Then
        lngRowIndex = GetWaitQueueIndex()
        If lngRowIndex > 0 Then lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
    Else
        lngRowIndex = GetCalledQueueIndex()
        If lngRowIndex > 0 Then lngQueueId = Val(rptCallList.Rows(lngRowIndex).Record(GetColIndex("ID", rptCallList)).value)
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "û�пɹ����еĶ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If

    'ִ�й㲥����
    If CheckIsQueueing(lngQueueId) Then
        MsgBox "��ǰ���ݴ����Ŷ�״̬������ִ�д˲�����", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    '�����¼�
    strCallContext = ""
    If DoCallPreBefore(lngQueueId, TCallWay.cwBroadcast, strCallContext) = False Then Exit Sub
    
    If mobjQueueManage.BroadcastCall(lngQueueId, strCallContext) <= 0 Then Exit Sub
    
    Call DoCallPreAfter(lngQueueId, TCallWay.cwBroadcast)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
  
Public Sub comMenu_���()
'ִ�ж��в�Ӳ���
On Error GoTo errHandle
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim strQueueName As String
    
    lngRowIndex = GetWaitQueueIndex()
    
    If lngRowIndex < 0 Then
        MsgBox "û����Ҫ��ӵĶ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(qftWaitQueue, lngRowIndex, "ID"))
    strQueueName = GetListValue(qftWaitQueue, lngRowIndex, "��������")
        
    '�����¼�
    If DoWorkBefore(qftWaitQueue, lngRowIndex, lngQueueId, TOperationType.otInsertQueue) = False Then Exit Sub
        
    If frmPriorityCause.ShowPriorityCause(Me, rptQueueList, lngRowIndex, mintWorkType, mstrReason) = True Then
        If GetWaitQueueSelState = qss�Ŷ��� Then
            Call LoadWaitQueueData
        End If
        
        Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otInsertQueue)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Sub comMenu_����()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim lngQueueSelState As Long
    Dim objRestoreQueue As New frmRestoreQueue

    Dim strCurQueueName As String
    Dim dtCurQueueDate As Date
    Dim strNewQueueNo As String
    Dim strQueueOrder As String
    Dim strNewQueueName As String
    Dim dtQueueDate As Date
    
    Dim lngMsgResult As Long
    
    'ִ�����Ų���ʱ����Ҫ�жϵ�ǰ��ѡ�б������Ŷ��б����Ѻ����б����ݲ�ͬ���б�ȡ����Ҫ���ŵ�����
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
    Else
        lngRowIndex = GetWaitQueueIndex()
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "û����Ҫ���ŵĶ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueSelState = GetWaitQueueSelState
    
    '��ȡ���ڶ��е��Ŷ�ID
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strCurQueueName = GetListValue(CurQueueType, lngRowIndex, "��������")
    
    '�����¼�����
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otRestore) = False Then Exit Sub
    
    dtCurQueueDate = Nvl(mobjQueueManage.GetQueueInf(lngQueueId, "�Ŷ�ʱ��")!�Ŷ�ʱ��, Now)
    
    Call objRestoreQueue.ShowRestoreQueueWindow(mstrQueryQueueNames, strCurQueueName, strNewQueueName, dtCurQueueDate, dtQueueDate, Me)
    If Trim(strNewQueueName) = "" Then Exit Sub
    
    '�������Ƹı�����Ŷ�ʱ��ı䣬����Ҫ���������ŶӺ���
    If strNewQueueName <> strCurQueueName Or dtQueueDate <> dtCurQueueDate Then
        '�����˵�ǰ���ŶӶ��У���Ҫ�����µ��ŶӺ���
        strNewQueueNo = mobjQueueManage.GetQueueMaxNo(strNewQueueName, dtQueueDate)
        
        lngMsgResult = MsgBox("���з����仯���Ѳ����µ��ŶӺ��� [" & strNewQueueNo & "], �Ƿ������", vbYesNo, "��ʾ")
        If lngMsgResult = vbNo Then Exit Sub
        
        Call mobjQueueManage.UpdateQueue(lngQueueId, "��������=''" & strNewQueueName & "''" & IIf(strNewQueueName <> strCurQueueName, ",����=''''", "") & ",�ŶӺ���=''" & strNewQueueNo & "'',�Ŷ�״̬=-1,�Ŷ�ʱ��=To_Date(''" & dtQueueDate & "'', ''yyyy-mm-dd hh24:mi:ss'')")
        
        Call DoCreateQueueNo(lngQueueId, strNewQueueName, strNewQueueNo)
    End If
    
        
    '�ŶӶ���û�н��иı䣬�����ڸö����Ŷӣ�����Ҫ�����µ��ŶӺ���
    strQueueOrder = mobjQueueManage.RestoreQueue(lngQueueId)
    If Trim(strQueueOrder) = "" Then Exit Sub
        
    strQueueOrder = FormatQueueOrder(strQueueOrder)
        
    If CurQueueType = qftCalledQueue Then
        Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        If lngQueueSelState = qss�Ŷ��� Then
            Call LoadWaitQueueData
        End If

    ElseIf CurQueueType = qftWaitQueue Then
        If lngQueueSelState <> qss�Ŷ��� Then
            Call DelQueueRecord(qftWaitQueue, lngRowIndex)
        Else
            '���������ŶӶ�������
            Call LoadWaitQueueData
        End If
        
    Else
        '���Ĳ��Ҷ����е�����״̬,������������ڶ��У�������Ϊ��
        If strNewQueueName <> strCurQueueName Then
            Call SetListValue(qftFindQueue, lngRowIndex, "����", "")
        End If
        
        Call SetListValue(qftFindQueue, lngRowIndex, "�ŶӺ���", strNewQueueNo)
        Call SetListValue(qftFindQueue, lngRowIndex, "�Ŷ����", strQueueOrder)
        Call SetListValue(qftFindQueue, lngRowIndex, "��������", strNewQueueName)
        
        Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsQueueing)
        
        Call rptQueueList.Populate
        
    End If
                    
    Call DoWorkAfter(lngQueueId, strNewQueueName, TOperationType.otRestore)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function DoWorkBefore(ByVal lngListType As TQueueFromType, ByVal lngListRow As Long, ByVal lngQueueId As Long, ByVal lngOperationType As TOperationType) As Boolean
'ִ��OnWorkBefore�¼�
    Dim blnCancel As Boolean
    
    DoWorkBefore = True
    blnCancel = False
    RaiseEvent OnWorkBefore(lngListType, lngListRow, lngQueueId, lngOperationType, blnCancel)
    
    DoWorkBefore = Not blnCancel
End Function


Private Sub DoWorkAfter(ByVal lngQueueId As Long, ByVal strCurQueueName As String, ByVal lngOperationType As TOperationType)
'ִ��OnWorkAfter�¼�
    RaiseEvent OnWorkAfter(lngQueueId, strCurQueueName, lngOperationType)
End Sub


Private Sub DoCreateQueueNo(ByVal lngQueueId As Long, ByVal strQueueName As String, ByRef strQueueNo As String)
'ִ��OnCreateQueueNo�¼�
    RaiseEvent OnCreateQueueNo(lngQueueId, strQueueName, strQueueNo)
End Sub


Private Sub AutoComplete()
'�Զ�����ѽ��ﴦ��
'ֻ�ܽ��������Լ����е�����
    Dim i As Long
    Dim lngColIndex As Long
    Dim lngQueueId As Long
    Dim lngIdColIndex As Long
    Dim blnAutoAll As Boolean
    Dim lngQueueSelState As Long
    Dim lngQueueNameIndex As Long
    Dim strQueueName As String
    Dim lngQueueNoIndex As Long
    
    lngIdColIndex = GetColIndex("ID", rptCallList)
    
    lngColIndex = GetColIndex("�Ŷ�״̬", rptCallList)
    lngQueueNameIndex = GetColIndex("��������", rptCallList)
                    
    lngQueueSelState = GetWaitQueueSelState
    blnAutoAll = True
    
    For i = rptCallList.Rows.Count - 1 To 0 Step -1
        If Not rptCallList.Rows(i).GroupRow Then
            If rptCallList.Rows(i).Record(lngColIndex).value = "������" And i <> rptCallList.SelectedRows(0).Index Then
                lngQueueId = rptCallList.Rows(i).Record(lngIdColIndex).value
                strQueueName = rptCallList.Rows(i).Record(lngQueueNameIndex).value
    
                '�жϸļ���Ƿ������Լ����е�����
                If Nvl(mobjQueueManage.GetQueueInf(lngQueueId, "����ҽ��")!����ҽ��) = mstrLoginUserName Then
                
                    '�����¼�
                    If DoWorkBefore(CurQueueType, i, lngQueueId, TOperationType.otComplete) = False Then Exit Sub
        
                    Call mobjQueueManage.CompleteQueue(lngQueueId)
                    
                    Call SetQueueRowState(rptCallList, i, qsComplete)
                    
                    If lngQueueSelState = qss����� Then
                        Call CopyCallRowToWaitRow(i)
                    End If
                    
                    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otComplete)
                Else
                    blnAutoAll = False
                End If
            End If
        End If
    Next i
    
    'ˢ���б�
    Call Populate
    
    If Not blnAutoAll Then
        MsgBox "�����ѽ����������[" & mstrLoginUserName & "]���У�δִ���Զ���ɲ�����", vbOKOnly, "��ʾ"
    End If
End Sub


Private Sub DelCompleteQueue()
'ɾ����ɶ�������
On Error Resume Next
    Dim i As Long
    Dim lngColIndex As Long
    Dim lngQueueNameColIndex As Long
    
    lngColIndex = GetColIndex("�Ŷ�״̬", rptCallList)
        
    For i = rptCallList.Rows.Count - 1 To 0 Step -1
        If Not rptCallList.Rows(i).GroupRow Then
            If rptCallList.Rows(i).Record(lngColIndex).value = "�����" Then
                Call DelQueueRecord(qftCalledQueue, i)
            End If
        End If
    Next i
    
    rptCallList.Populate
Err.Clear
End Sub

Public Sub comMenu_����()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strQueueName As String
    Dim objList As ReportControl
      

    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
        Set objList = rptCallList
    Else
        lngRowIndex = GetWaitQueueIndex()
        Set objList = rptQueueList
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "û�пɹ�����Ķ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "��������")
    
    
    '�ж��Ƿ�������ﴦ��
    If CheckIsQueueing(lngQueueId) Then
        MsgBox "��ǰ���ݴ����Ŷ�״̬������ִ�д˲�����", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    '�����¼�
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otDiagnose) = False Then Exit Sub

    If Not mobjQueueManage.DiagnoseQueue(lngQueueId) Then Exit Sub
        
    If CurQueueType = qftCalledQueue Then
        '���Ĳ��Ҷ����е�����״̬
        
        If mblnAutoComplete Then
            '���ѽ���������޸�Ϊ�����
            Call AutoComplete
        End If
        
        Call SetQueueRowState(objList, lngRowIndex, TQueueState.qsDiagnose)
        
        If mblnAutoComplete Then
            Call DelCompleteQueue
        End If
    Else
        '�ŶӶ�����Ҫ�������ݺ�ת�Ƶ����ж�����ʾ
        Call SetQueueRowState(objList, lngRowIndex, TQueueState.qsDiagnose)
        Call CopyWaitRowToCallRow(lngRowIndex)
        
        Call rptCallList.Populate
        
        'ɾ���Ѿ�����ļ�¼
        Call DelQueueRecord(qftWaitQueue, lngRowIndex)
    End If
        
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otDiagnose)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_��ͣ()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strQueueName As String
    Dim objList As ReportControl
    
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
        Set objList = rptCallList
    Else
        lngRowIndex = GetWaitQueueIndex()
        Set objList = rptQueueList
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "û����Ҫ��ͣ�Ķ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "��������")
    
    '�����¼�����
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otPause) = False Then Exit Sub
    
    If Not mobjQueueManage.PauseQueue(lngQueueId) Then Exit Sub
    
    Select Case CurQueueType
        Case qftCalledQueue
            '���Ѻ��ж������ݣ�ת�Ƶ��ŶӶ���
            If GetWaitQueueSelState = qss����ͣ Then
                Call SetQueueRowState(objList, lngRowIndex, qsPause)
                Call CopyCallRowToWaitRow(lngRowIndex)
                
                Call rptQueueList.Populate
            End If
            
            Call DelQueueRecord(qftCalledQueue, lngRowIndex)
            
        Case qftFindQueue
            '���Ĳ��Ҷ����е�����״̬
            Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsPause)
            Call rptQueueList.Populate
        Case qftWaitQueue
            If optOutQueue(TQueueSelState.qss����ͣ).value Then
                'ֱ�Ӹ����б�״̬
                Call SetQueueRowState(objList, lngRowIndex, qsPause)
                Call rptQueueList.Populate
            Else
                If GetWaitQueueSelState() <> qss����ͣ Then
                    Call DelQueueRecord(qftWaitQueue, lngRowIndex)
                End If
            End If
    End Select
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otPause)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_����()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim strQueueName As String
    Dim objList As ReportControl
    
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
        Set objList = rptCallList
    Else
        lngRowIndex = GetWaitQueueIndex()
        Set objList = rptQueueList
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "û����Ҫ�����Ķ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
       
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "��������")
    
    '�����¼�����
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otAbstain) = False Then Exit Sub
    
    If Not mobjQueueManage.AbstainQueue(lngQueueId) Then Exit Sub
    
    
    Select Case CurQueueType
        Case qftCalledQueue
            If GetWaitQueueSelState() = qss������ Then
                Call SetQueueRowState(objList, lngRowIndex, qsAbstain)
                Call CopyCallRowToWaitRow(lngRowIndex)
                
                Call rptQueueList.Populate
            End If
            
            Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        Case qftWaitQueue
            If optOutQueue(TQueueSelState.qss������).value Then
                'ֱ�Ӹ����б�״̬
                Call SetQueueRowState(objList, lngRowIndex, qsAbstain)
                Call rptQueueList.Populate
            Else
                Call DelQueueRecord(qftWaitQueue, lngRowIndex)
            End If
        Case qftFindQueue
            '���Ĳ��Ҷ����е�����״̬
            Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsAbstain)
            Call rptQueueList.Populate
    End Select
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otAbstain)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_�ָ�()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim lngQueueSelState As Long
    Dim strNewQueueNo As String
    Dim strQueueName As String
    
    'ִ�лָ�����ʱ����Ҫ�жϵ�ǰ��ѡ�б������Ŷ��б����Ѻ����б����ݲ�ͬ���б�ȡ����Ҫ�ָ�������
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
    Else
        lngRowIndex = GetWaitQueueIndex()
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "û����Ҫ�ָ��Ķ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueSelState = GetWaitQueueSelState
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "��������")
    
    '�����¼�����
    If DoWorkBefore(CurQueueType, lngRowIndex, lngQueueId, TOperationType.otStart) = False Then Exit Sub
    
    If Not mobjQueueManage.LineQueue(lngQueueId, strNewQueueNo) Then Exit Sub
    
    If strNewQueueNo <> "" Then
        Call DoCreateQueueNo(lngQueueId, strQueueName, strNewQueueNo)
    Else
        strNewQueueNo = GetListValue(CurQueueType, lngRowIndex, "�ŶӺ���")
    End If
        
    If CurQueueType = qftCalledQueue Then
        'ˢ���Ѻ����б�������ʾ
        Call SetQueueRowState(rptCallList, lngRowIndex, qsQueueing)
        
        If strNewQueueNo <> "" Then Call SetListValue(qftCalledQueue, lngRowIndex, "�ŶӺ���", strNewQueueNo)
        Call CopyCallRowToWaitRow(lngRowIndex)
        
        Call rptQueueList.Populate
        
        'ɾ���Ѿ������е�������
        Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        
    ElseIf CurQueueType = qftWaitQueue Then
        If lngQueueSelState <> qss�Ŷ��� Then
            Call DelQueueRecord(qftWaitQueue, lngRowIndex)
        End If
        
    Else
        '���Ĳ��Ҷ����е�����״̬
        Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsQueueing)
        
        If strNewQueueNo <> "" Then Call SetListValue(qftFindQueue, lngRowIndex, "�ŶӺ���", strNewQueueNo)
        Call SetListValue(qftFindQueue, lngRowIndex, "�Ŷ�ʱ��", Now)
        
        Call rptQueueList.Populate
        
    End If
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otStart)
    
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub SetQueueRowState(objList As ReportControl, ByVal lngRow As Long, ByVal lngState As TQueueState)
'�ָ��Ŷ���ʾ״̬
    Dim lngStateColIndex As Long
        
    lngStateColIndex = GetColIndex("�Ŷ�״̬", objList)
    If lngStateColIndex >= 0 Then
        objList.Rows(lngRow).Record(lngStateColIndex).value = Decode(lngState, _
                                                                    TQueueState.qsAbstain, "������", _
                                                                    TQueueState.qsCalling, "������", _
                                                                    TQueueState.qsCalled, "�Ѻ���", _
                                                                    TQueueState.qsComplete, "�����", _
                                                                    TQueueState.qsPause, "����ͣ", _
                                                                    TQueueState.qsQueueing, "�Ŷ���", _
                                                                    TQueueState.qsDiagnose, "������", _
                                                                    TQueueState.qsWaitCall, "������", _
                                                                    "")
    End If
    
    lngStateColIndex = GetFirstDisplayColIndex(objList)
    
    If lngStateColIndex >= 0 Then
        Select Case lngState
            Case TQueueState.qsDiagnose
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_DIAGNOSE
            Case TQueueState.qsCalling
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_CALLING
            Case TQueueState.qsCalled
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_CALLED
            Case TQueueState.qsQueueing
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_QUEUEING
            Case Else   '��ͣ����ɣ����ž�ʹ����ͬ��ͼ��
                objList.Rows(lngRow).Record(lngStateColIndex).Icon = M_LNG_ICON_QUEUEING
        End Select
        
    End If
End Sub


Public Sub comMenu_���()
On Error GoTo errHandle
    Dim lngRowIndex As Long
    Dim lngQueueId As Long
    Dim lngQueueSelState As Long
    Dim strQueueName As String
    
    'ִ�лָ�����ʱ����Ҫ�жϵ�ǰ��ѡ�б������Ŷ��б����Ѻ����б����ݲ�ͬ���б�ȡ����Ҫ�ָ�������
    If CurQueueType = qftCalledQueue Then
        lngRowIndex = GetCalledQueueIndex()
    Else
        lngRowIndex = GetWaitQueueIndex()
    End If
    
    If lngRowIndex < 0 Then
        MsgBox "û����Ҫ��ɵĶ������ݱ�ѡ��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    lngQueueSelState = GetWaitQueueSelState
    
    lngQueueId = Val(GetListValue(CurQueueType, lngRowIndex, "ID"))
    strQueueName = GetListValue(CurQueueType, lngRowIndex, "��������")
    
    If CheckIsQueueing(lngQueueId) Then
        MsgBox "��ǰ���ݴ����Ŷ�״̬������ִ�д˲�����", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    
    '�����¼�
    If DoWorkBefore(qftCalledQueue, lngRowIndex, lngQueueId, TOperationType.otComplete) = False Then Exit Sub
    
    If Not mobjQueueManage.CompleteQueue(lngQueueId) Then Exit Sub
        
        
    If CurQueueType = qftCalledQueue Then
        If lngQueueSelState = qss����� Then
            Call SetQueueRowState(rptCallList, lngRowIndex, qsComplete)
            Call CopyCallRowToWaitRow(lngRowIndex)
            
            Call rptQueueList.Populate
        End If
        
        '�Ӻ��ж�����ɾ���Ѿ���ɵ�����
        Call DelQueueRecord(qftCalledQueue, lngRowIndex)
        
    ElseIf CurQueueType = qftWaitQueue Then
        If optOutQueue(TQueueSelState.qss�����).value Then
            'ֱ�Ӹ����б�״̬
            Call SetQueueRowState(rptQueueList, lngRowIndex, qsComplete)
            Call rptQueueList.Populate
        Else
            If lngQueueSelState <> qss����� Then
                Call DelQueueRecord(qftWaitQueue, lngRowIndex)
            End If
        End If
    Else
        '���Ĳ��Ҷ����е�����״̬
        Call SetQueueRowState(rptQueueList, lngRowIndex, TQueueState.qsComplete)
        Call rptQueueList.Populate
        
    End If
    
    Call DoWorkAfter(lngQueueId, strQueueName, TOperationType.otComplete)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub comMenu_����()
On Error GoTo errHandle
    Dim strResult As String
    Dim strFilterWhere As String
    Dim strFilterValue As String
    Dim blnCancel As Boolean
    Dim blnUseCustom As Boolean
    Dim rsData As ADODB.Recordset
    
    blnUseCustom = False
    blnCancel = False
    
    RaiseEvent OnFilter(rsData, blnCancel, blnUseCustom)
    
    If blnCancel Then Exit Sub
    
    If Not blnUseCustom Then
        Call frmFilter.ShowFilterWindow(mobjQueueManage.DefQueryCols, strFilterWhere, strFilterValue, Me)
        If Trim(strFilterWhere) = "" Then Exit Sub
        
        RaiseEvent OnFindData(strFilterWhere, strFilterValue, Nothing, rsData, blnUseCustom)
        
        If Not blnUseCustom Then
            Set rsData = DefaultFind(strFilterWhere, strFilterValue)
        End If
        
    End If
    
    'ɾ����ǰ����
    Call rptQueueList.Records.DeleteAll
    Call rptQueueList.Populate
    
    Call rptCallList.Records.DeleteAll
    Call rptCallList.Populate
    
    '���ý���״̬
    Call ConfigQueueStateSel(False)
    mblnIsFindQueue = True

    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub

    '�����ѯ����
    Call LoadFindData(rptQueueList, rsData)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub comMenu_ˢ��()
On Error GoTo errHandle

    rptQueueList.Tag = ""
    rptCallList.Tag = ""
    
    '���¼�������
    Call RefreshQueueData

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub ConfigDefaultInputPro(objInputCfg As Dictionary)
'����Ĭ��¼����Ŀ
    'String,Number,Date,DateTime
    
    objInputCfg.Add "��������", "STRING"
    objInputCfg.Add "�ŶӺ���", "STRING"
    objInputCfg.Add "����", "STRING"
    objInputCfg.Add "��ע", "STRING"
    objInputCfg.Add "�Ŷ�ʱ��", "DATETIME"
    objInputCfg.Add "�Ŷӱ��", "STRING"
End Sub

Public Sub comMenu_�޸�()
On Error GoTo errHandle
    Dim objUpdateWind As New frmUpdateInfo
    Dim lngQueueId As Long
    Dim lngRowIndex As Long
    Dim blnCancel As Boolean
    Dim blnUseCustom As Boolean
    

    Dim objInputCfg As New Dictionary
    Dim objReturn As New Dictionary
    Dim strKey As Variant
  
    lngQueueId = 0
    
    If CurQueueType = qftFindQueue Or CurQueueType = qftWaitQueue Then
        lngRowIndex = GetWaitQueueIndex()
    Else
        lngRowIndex = GetCalledQueueIndex()
    End If
           
    If lngRowIndex < 0 Then
        MsgBox "��ѡ����Ҫ�޸ĵĶ��м�¼��", vbOKOnly Or vbInformation, GetWindowCaption
        Exit Sub
    End If
    
    If CurQueueType = qftCalledQueue Then
        lngQueueId = Val(rptCallList.Rows(lngRowIndex).Record(GetColIndex("ID", rptCallList)).value)
    Else
        lngQueueId = Val(rptQueueList.Rows(lngRowIndex).Record(GetColIndex("ID", rptQueueList)).value)
    End If
    
    blnCancel = False
    blnUseCustom = False
    
    Call ConfigDefaultInputPro(objInputCfg)
    
    RaiseEvent OnModifyBefore(CurQueueType, lngQueueId, objInputCfg, blnCancel, blnUseCustom)
    
    If blnCancel = True Then Exit Sub
    If blnUseCustom = True Then Exit Sub
    
    If objUpdateWind.zlShowMe(lngQueueId, objInputCfg, objReturn, mobjQueueManage, Me) = True Then
                      
        RaiseEvent OnModifyAfter(lngQueueId, objReturn)
        
        'ͬ�������б��е�����
        For Each strKey In objReturn.Keys
            Call SetListValue(CurQueueType, lngRowIndex, strKey, objReturn.Item(strKey))
        Next
                
        Call Populate(CurQueueType)
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Public Sub comMenu_����()
On Error GoTo errHandle
    '���ô򿪲������ý���
    Dim blnUseCustom As Boolean
    
    blnUseCustom = False
    RaiseEvent OnConfigEvent(blnUseCustom)
    
    If blnUseCustom Then Exit Sub
    
    If ShowVoiceConfig Then
        Call ApplyVoiceConfig
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


'��ʼ������ѯ��ʾ
Private Sub StartCall()
    tmrBroadCast.Enabled = True
End Sub


'��ֹ������ѯ��ʾ
Private Sub AbortCall()
    tmrBroadCast.Enabled = False
End Sub


'��ѯ������ʾ�ͺ���
Private Sub LoopPlayVoice()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngVoiceId As Long
    Dim lngQueueId As Long
    Dim strVoiceContext As String
    Dim blnCancel As Boolean
    Dim i As Long
    Dim lngRowIndex As Long
    Dim blnAllowQuery As Boolean
    
    '�ж��Ƿ���Ҫ��ѯ���ݿ�����
    blnAllowQuery = IIf(mrsVoiceContext Is Nothing, True, False)
    
    If Not mrsVoiceContext Is Nothing Then
        blnAllowQuery = IIf(mrsVoiceContext.RecordCount <= 0 Or mrsVoiceContext.EOF, True, False)
    End If
    
    If blnAllowQuery Then
        If Timer < mdtLastVoiceDate + mlngInterval Then Exit Sub
        mdtLastVoiceDate = Timer
        
        '��ѯ��Ҫ���ŵ���������
        strSql = "select id,����ID,��������,����ʱ�� from �Ŷ���������  where վ��=[1] order by ����ʱ��"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ������������", mstrComputerName)
        
        If rsData.RecordCount > 0 Then
            Set mrsVoiceContext = zlDatabase.CopyNewRec(rsData)
            mrsVoiceContext.Sort = "����ʱ�� asc"
        End If
    End If
    
    If mrsVoiceContext Is Nothing Then Exit Sub
    If mrsVoiceContext.RecordCount <= 0 Or mrsVoiceContext.EOF Then Exit Sub
    
'    mdtLastVoiceDate = Timer
    
    lngVoiceId = Val(Nvl(mrsVoiceContext!Id))
    lngQueueId = Val(Nvl(mrsVoiceContext!����ID))
    strVoiceContext = Nvl(mrsVoiceContext!��������)
    
    Call mrsVoiceContext.MoveNext

    blnCancel = False
    RaiseEvent OnPlayVoiceBefore(lngVoiceId, lngQueueId, strVoiceContext, blnCancel)
    
    If blnCancel Then Exit Sub
    
    If lngQueueId <= 0 Then
        '�����Զ���ĺ�������
        Call mobjQueueManage.PlayCustomVoice(lngVoiceId, False, strVoiceContext)
    Else
        '���ú����еĵ�ǰ��ɫ
        lngRowIndex = GetRowIndex(qftCalledQueue, "ID", lngQueueId)
        
        If lngRowIndex >= 0 Then
            Call SetQueueRowState(rptCallList, lngRowIndex, qsCalling)
            Call rptCallList.Populate
        End If
        
        '��������
        Call mobjQueueManage.PlayQueueVoice(lngVoiceId, lngQueueId, False, strVoiceContext)
    End If
    
    RaiseEvent OnPlayVoiceAfter(lngVoiceId, lngQueueId, strVoiceContext)

    '���гɹ���ɾ�����й�������
    Call mobjQueueManage.DelVoiceData(lngVoiceId)
    
    If lngQueueId > 0 Then
        '���ú����еĵ�ǰ��ɫ
        lngRowIndex = GetRowIndex(qftCalledQueue, "ID", lngQueueId)
                
        If lngRowIndex >= 0 Then
            If GetListValue(qftCalledQueue, lngRowIndex, "�Ŷ�״̬") = "������" Then
                Call SetQueueRowState(rptCallList, lngRowIndex, qsCalled)
                Call rptCallList.Populate
            End If
        End If
    End If
    
    
End Sub


Public Sub StartVoice()
'��ʼ��������
    mdtLastVoiceDate = Timer - mlngInterval
    
    tmrBroadCast.Interval = mlngInterval
    tmrBroadCast.Enabled = True
    tmrBroadCast.Tag = 0
End Sub

Public Sub StopVoice()
'������������
On Error GoTo errHandle
    tmrBroadCast.Tag = 1
    tmrBroadCast.Enabled = False
    
    If Not mobjQueueManage Is Nothing Then Call mobjQueueManage.StopVoice
Exit Sub
errHandle:
    Debug.Print "StopVoice Err:" & Err.Description
End Sub


Private Sub timerCard_Timer()
On Error GoTo errHandle
    If GetTickCount - mlngStartTime > 200 Then
        '����200����ʱ���Զ���Ϊˢ������
        timerCard.Enabled = False
        
        mlngStartTime = 0
        mlngAvgTime = 0
        mlngReadCount = 0
        
        Call zlControl.TxtSelAll(txtLocateValue)
        
        If IsFindModel Then
            '�������ݲ���
            Call FindQueueData(mstrLocateType, txtLocateValue.Text)
        Else
            '�������ݶ�λ
            Call LocateQueueData(mstrLocateType, txtLocateValue.Text)
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub LocateQueueData(ByVal strLocateType As String, ByVal strFindValue As String)
    Dim blnUseCustom As Boolean
    Dim lngQueueId As Long
    Dim i As Long, j As Long
    Dim lngIdColIndex As Long
    Dim blnOldBold As Boolean

    
    If Trim(strFindValue) = "" Then Exit Sub
    
    blnUseCustom = False
    RaiseEvent OnLocateData(strLocateType, strFindValue, txtLocateValue, lngQueueId, blnUseCustom)
    
    If Not blnUseCustom Then
        lngQueueId = DefaultLocate(strLocateType, strFindValue)
    End If
    

    Call rptQueueList.SelectedRows.DeleteAll
    Call rptCallList.SelectedRows.DeleteAll
    
    
    If lngQueueId <= 0 Then Exit Sub
    
    
    lngIdColIndex = GetColIndex("ID", rptQueueList)
    
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow = True Then
            For j = 0 To rptQueueList.Rows(i).Childs.Count - 1
                If rptQueueList.Rows(i).Childs(j).Record(lngIdColIndex).value = lngQueueId Then
                    Call zlControl.TxtSelAll(txtLocateValue)
                    
                    rptQueueList.Rows(i).Expanded = True
                    rptQueueList.Rows(i).Childs(j).Selected = True
                    
                    Set rptQueueList.FocusedRow = rptQueueList.Rows(i).Childs(j)
        
                    mblnIsSelectedCallingList = False
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Exit Sub
                End If
            Next j
        End If
    Next i
    
    
    lngIdColIndex = GetColIndex("ID", rptCallList)
    
    For i = 0 To rptCallList.Rows.Count - 1
        If rptCallList.Rows(i).GroupRow = True Then
            For j = 0 To rptCallList.Rows(i).Childs.Count - 1
                If rptCallList.Rows(i).Childs(j).Record(lngIdColIndex).value = lngQueueId Then
                    Call zlControl.TxtSelAll(txtLocateValue)
                    
                    rptCallList.Rows(i).Expanded = True
                    rptCallList.Rows(i).Childs(j).Selected = True
                    
                    Set rptCallList.FocusedRow = rptCallList.Rows(i).Childs(j)
        
                    mblnIsSelectedCallingList = True
                    Call SwitchActiveWindow(mblnIsSelectedCallingList)
                    
                    Exit Sub
                End If
            Next j
        End If
    Next i
    
    
End Sub


Private Sub ConfigQueueStateSel(ByVal blnEnable As Boolean)
''���ö��й���״̬�Ƿ���������
'    lblQueueFilter(0).Enabled = blnEnable
'    lblQueueFilter(1).Enabled = blnEnable
'    lblQueueFilter(2).Enabled = blnEnable
'    lblQueueFilter(3).Enabled = blnEnable
'
'    optOutQueue(0).Enabled = blnEnable And optOutQueue(0).value = 0
'    optOutQueue(1).Enabled = blnEnable And optOutQueue(1).value = 0
'    optOutQueue(2).Enabled = blnEnable And optOutQueue(2).value = 0
'    optOutQueue(3).Enabled = blnEnable And optOutQueue(3).value = 0
'
'    If mblnIsShowCalledQueue = True Then
'        DkpMain.Panes(2).Closed = Not blnEnable
'    End If
    
End Sub


Public Sub FindQueueData(ByVal strLocateType As String, ByVal strFindValue As String)
    Dim blnUseCustom As Boolean
    Dim rsData As ADODB.Recordset
    
    If Trim(strFindValue) = "" Then
        '���û��¼���ѯ���ݣ����ʾˢ������
        Call RefreshQueueData
        Exit Sub
    End If
    
    blnUseCustom = False
    RaiseEvent OnFindData(strLocateType, strFindValue, txtLocateValue, rsData, blnUseCustom)
    
    If Not blnUseCustom Then
        'ʹ��Ĭ�ϵĲ�ѯ
        Set rsData = DefaultFind(strLocateType, strFindValue)
    End If

    Call rptQueueList.Records.DeleteAll
    Call rptQueueList.Populate
    
    Call rptCallList.Records.DeleteAll
    Call rptCallList.Populate
    
    Call ConfigQueueStateSel(False)
    mblnIsFindQueue = True

    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub

    Call LoadFindData(rptQueueList, rsData)
    
End Sub


Private Sub LoadFindData(objQueueList As ReportControl, rsData As ADODB.Recordset)
'�����������

On Error GoTo errHandle
    Dim rptRecord As ReportRecord
    Dim blnCancel As Boolean
    Dim i As Long

'�������ݵ��б�

    Call objQueueList.Records.DeleteAll
    Call objQueueList.Populate
    
    If rsData.RecordCount <= 0 Then Exit Sub

    While Not rsData.EOF

        blnCancel = False
        RaiseEvent OnReadBefore(rsData, TQueueFromType.qftFindQueue, blnCancel)
        
        If Not blnCancel Then
            Set rptRecord = objQueueList.Records.Add
            
            For i = 0 To objQueueList.Columns.Count - 1
                rptRecord.AddItem ""
            Next
    
            Call SetReportRecordItem(objQueueList, rptRecord, rsData)
            
            RaiseEvent OnReadAfter(rsData, TQueueFromType.qftFindQueue, rptRecord)
        End If

        rsData.MoveNext
    Wend

    objQueueList.Populate

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function DefaultFind(ByVal findType As String, ByVal findData As String) As ADODB.Recordset
    Dim strSql As String, strFilter As String
    Dim strQueueNames As String
    Dim varValue1 As Variant
    Dim varValue2 As Variant
    
    On Error GoTo errHandle
    
    strFilter = ""
    varValue1 = findData
    varValue2 = ""
    
    Select Case findType  ' '0-�ŶӺ�;1-����;
    Case "�ŶӺ�", "�ŶӺ���"
        varValue1 = findData
        
        If mblnIsReleationQueueTag Then
            strFilter = " and upper(�Ŷӱ��)||upper(�ŶӺ���) = upper([2])"
        Else
            strFilter = " and upper(�ŶӺ���) = upper([2])"
        End If
    Case "����", "��������"
        varValue1 = findData & "%"
        strFilter = " and upper(��������) Like upper([2])"

    Case Else
        strFilter = " and upper(" & findType & ")=upper([2])"
        
        Select Case True
            Case IsNumeric(findData)
                varValue1 = Val(findData)
            Case IsDate(findData)
                
                If Format(findData, "hh:mm:ss") = "00:00:00" Then
                    varValue1 = CDate(Format(findData, "yyyy-mm-dd 00:00:00"))
                    varValue2 = CDate(Format(findData, "yyyy-mm-dd 23:59:59"))
                    strFilter = " and " & findType & " between [2] and [3] "
                Else
                    varValue1 = CDate(findData)
                    strFilter = " and " & findType & " = [2]"
                End If
                
            Case Else
                varValue1 = findData
        End Select
        
        
    End Select
    
    strQueueNames = mstrQueryQueueNames
    
    If strQueueNames <> "" Then
        strQueueNames = Replace(strQueueNames, ",", "','")
        strFilter = strFilter & " and �������� in ('" & strQueueNames & "') "
    End If
    
    strSql = "select * from �ŶӽкŶ��С�where  ҵ������=[1] " & strFilter & " order by �ŶӺ��� "

    Set DefaultFind = zlDatabase.OpenSQLRecord(strSql, "��Ĭ�Ϸ�ʽ���Ҷ���", mintWorkType, varValue1, varValue2)
     
    Exit Function
errHandle:
    Set DefaultFind = Nothing
    If ErrCenter = 1 Then Resume

End Function


Private Sub tmrBroadCast_Timer()
On Error GoTo errHandle
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    'ֹͣ��ѵ
    Call AbortCall
    
    If Val(tmrBroadCast.Tag) <> 1 Then
        '������ѵ����
        Call LoopPlayVoice
    End If
    
    If Val(tmrBroadCast.Tag) = 1 Then Exit Sub
    
    '��ʼ��ѵ
    Call StartCall

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub InitLocalParas()
On Error GoTo errHandle
    Dim i As Integer
    Dim X As Long, Y As Long, r As Long, b As Long

    mstrLocateType = GetSetting("ZLSOFT", gstrRegPath, "��λ��ʽ", "����")

'    mlngQueueW1 = GetSetting("ZLSOFT", gstrRegPath, "�ŶӶ�����ʾ���", Round(Width / 3 * 2))
'    mlngQueueW2 = GetSetting("ZLSOFT", gstrRegPath, "���ж�����ʾ���", Round(Width / 3))

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function DefaultLocate(ByVal strFindType As String, ByVal strFindData As String) As Long
    Dim i As Integer
    Dim j As Integer
    Dim rsData As ADODB.Recordset
    Dim strSql As String
    Dim lngPatientId As Long
    Dim blnFind As Boolean
    Dim lngFindColIndex As Long
    Dim lngStartIndex As Long
    Dim blnExpandState As Boolean
    
    DefaultLocate = 0
    lngFindColIndex = -1
    
    If strFindType = "�ŶӺ�" Or strFindData = "�ŶӺ���" Then
        lngFindColIndex = GetColIndex("�ŶӺ���", rptQueueList)
    ElseIf strFindType = "����" Or strFindData = "��������" Then
        lngFindColIndex = GetColIndex("��������", rptQueueList)
    Else
        lngFindColIndex = GetColIndex(strFindType, rptQueueList)
    End If
    
    If lngFindColIndex < 0 Then Exit Function
    
    '��ȡ��ʼ���ҵ�������
    If rptQueueList.SelectedRows.Count > 0 Then
        mlngLocateRowIndex = rptQueueList.SelectedRows(0).Index + 1
    End If
    
    If mlngLocateRowIndex >= rptQueueList.Rows.Count + rptCallList.Rows.Count - 1 Then mlngLocateRowIndex = 0
    
   
    blnFind = False
    For i = 0 To rptQueueList.Rows.Count - 1
        If rptQueueList.Rows(i).GroupRow = True Then
            blnExpandState = rptQueueList.Rows(i).Expanded
            rptQueueList.Rows(i).Expanded = True
            
            For j = 0 To rptQueueList.Rows(i).Childs.Count - 1
                If strFindType = "����" Or strFindType = "��������" Then
                    blnFind = IIf(rptQueueList.Rows(i).Childs(j).Index >= mlngLocateRowIndex And UCase(rptQueueList.Rows(i).Childs(j).Record(lngFindColIndex).value) Like UCase(strFindData) & "*", True, False)
                Else
                    blnFind = IIf(rptQueueList.Rows(i).Childs(j).Index >= mlngLocateRowIndex And UCase(rptQueueList.Rows(i).Childs(j).Record(lngFindColIndex).value) = UCase(strFindData), True, False)
                End If
        
                If blnFind Then
                    DefaultLocate = rptQueueList.Rows(i).Childs(j).Record(GetColIndex("ID", rptQueueList)).value
                    Exit Function
                End If
            Next j
            
            rptQueueList.Rows(i).Expanded = blnExpandState
        End If
    Next i


    '��ȡ��ʼ���ҵ�������
    If rptCallList.SelectedRows.Count > 0 Then
        mlngLocateRowIndex = rptQueueList.Rows.Count + rptCallList.SelectedRows(0).Index + 1
    End If
    
    If mlngLocateRowIndex > rptQueueList.Rows.Count + rptCallList.Rows.Count - 1 Then
        mlngLocateRowIndex = 0
        Exit Function
    End If
    
    '���û���ҵ����ݣ�����Ѻ��ж����в���
    For i = 0 To rptCallList.Rows.Count - 1
        If rptCallList.Rows(i).GroupRow = True Then
            blnExpandState = rptCallList.Rows(i).Expanded
            rptCallList.Rows(i).Expanded = True
            
            For j = 0 To rptCallList.Rows(i).Childs.Count - 1
                If strFindType = "����" Or strFindType = "��������" Then
                    blnFind = IIf(rptCallList.Rows(i).Childs(j).Index >= mlngLocateRowIndex - rptQueueList.Rows.Count And UCase(rptCallList.Rows(i).Childs(j).Record(lngFindColIndex).value) Like UCase(strFindData) & "*", True, False)
                Else
                    blnFind = IIf(rptCallList.Rows(i).Childs(j).Index >= mlngLocateRowIndex - rptQueueList.Rows.Count And UCase(rptCallList.Rows(i).Childs(j).Record(lngFindColIndex).value) = UCase(strFindData), True, False)
                End If
        
                If blnFind Then
                    DefaultLocate = rptCallList.Rows(i).Childs(j).Record(GetColIndex("ID", rptCallList)).value
                    Exit Function
                End If
            Next j
            
            rptCallList.Rows(i).Expanded = blnExpandState
        End If
    Next i
    
    mlngLocateRowIndex = 0
    
End Function


Private Sub txtLocateValue_GotFocus()
    On Error Resume Next
    
    Call zlControl.TxtSelAll(txtLocateValue)
End Sub


Private Sub txtLocateValue_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        If IsFindModel Then
            '�������
            Call FindQueueData(mstrLocateType, txtLocateValue.Text)
        Else
            '���붨λ
            Call LocateQueueData(mstrLocateType, txtLocateValue.Text)
        End If
        
        Exit Sub
    End If
    
    If KeyAscii = 8 Then Exit Sub
    
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    mlngReadCount = mlngReadCount + 1
    If mlngStartTime <> 0 Then
        If GetTickCount - mlngStartTime > 200 Then
            mlngReadCount = 1
            mlngAvgTime = 0
        Else
            mlngAvgTime = mlngAvgTime + (GetTickCount() - mlngStartTime)
        End If
    End If
    
    mlngStartTime = GetTickCount
    
    'ȡ����ƽ��¼��ʱ��
    If mlngReadCount = 3 Then
        mlngAvgTime = Fix(mlngAvgTime / 3)
        
        If mlngAvgTime <= 30 Then timerCard.Enabled = True
    End If

End Sub


Private Sub UserControl_Initialize()
    mblnIsShowBars = True
    mblnIsShowCalledQueue = True
    mlngInterval = 30000    'Ĭ��30����ѯһ��
    mblnIsFindQueue = False
    mstrReason = ""
    mblnAutoComplete = True
    mblnShowMySelfCalled = True
    mblnIsReleationQueueTag = False
    
    Set mrsVoiceContext = Nothing
    Set mobjQueueManage = New clsQueueOperation
    
    InitFaceScheme

End Sub

Private Sub UserControl_Resize()
'�����ؼ�λ�÷���
On Error Resume Next
    Call picCallFace_Resize
    Call picQueueFace_Resize
Err.Clear
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '�����ѯ�������ֶ�
    mstrDataFields = PropBag.ReadProperty("DataFields", "")

    '�Ŷ��б���ʾ�ֶ�
    mstrDisplayQueueFields = PropBag.ReadProperty("DisplayQueueColNames", "")

    '�����б���ʾ�ֶ�
    mstrDisplayCallFields = PropBag.ReadProperty("DisplayCallColNames", "")

    '������ʾ�Ķ�������  ע�����Ϊ�գ�����ʾ��ǰҵ�������µ����ж����е��ŶӺͺ�������
    mstrQueryQueueNames = PropBag.ReadProperty("QueryQueueNames", "")

    '����(��������ݷ�����ʾ)
    mstrGroupField = PropBag.ReadProperty("GroupField", "")
    
    '�Զ���������
    mstrCustomOrderColName = PropBag.ReadProperty("CustomOrderName", "")
    
    '�������ҷ�ʽ
    mstrFindWay = PropBag.ReadProperty("FindWay", "")
    
    '������ѯ���ʱ��
    mlngInterval = PropBag.ReadProperty("Interval", 30)
    
    '������Ч��������
    mobjQueueManage.ValidDays = PropBag.ReadProperty("ValidDays", 1)
    
    '�Ƿ���ʾ������
    IsShowBars = PropBag.ReadProperty("IsShowBars", True)
    
    '�Ƿ���ʾ�Ѻ��ж���
    IsShowCalledQueue = PropBag.ReadProperty("IsShowCalledQueue", True)
    
    '���ú��к��Ŀ�ĵ�
    mobjQueueManage.CallTarget = PropBag.ReadProperty("CalledTarget", "")
    
    '�Ƿ���ʾ��ť�ı�
    mlngMenuCaptionStyle = PropBag.ReadProperty("IsShowToolText", xtpButtonIconAndCaption)
    IsShowToolText = IIf(mlngMenuCaptionStyle = xtpButtonIcon, False, True)
    
    '�Ƿ���ʾ��ͼ��
    IsIconLarge = PropBag.ReadProperty("IsIconLarge", True)
    
    '�Ƿ���ʾ���ԭ��
    mstrReason = PropBag.ReadProperty("Reason", "")
End Sub

Private Sub UserControl_Terminate()
    
    Call StopVoice
    
    SaveColWidth rptQueueList
    SaveColWidth rptCallList
    
    If gstrRegPath <> "" Then
        SaveSetting "ZLSOFT", gstrRegPath, "QueueListWidthRate", picQueueFace.Width / ScaleWidth
    End If
        
    Set mobjQueueManage = Nothing

    Unload frmPriorityCause
    Unload frmSetup
    Unload frmFilter
End Sub

Private Sub SaveColWidth(objQueueList As Object)
'�����еĿ��
    Dim strColPro As String
    Dim i As Long
    
    If gstrRegPath = "" Then Exit Sub
    
    For i = 0 To objQueueList.Columns.Count - 1
        If objQueueList.Columns(i).Visible = True Then
            If strColPro <> "" Then strColPro = strColPro & ";"
            
            strColPro = strColPro & objQueueList.Columns(i).Caption & ":" & objQueueList.Columns(i).Width
        End If
    Next i
    
    SaveSetting "ZLSOFT", gstrRegPath, objQueueList.Name, strColPro
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DataFields", mstrDataFields, "")
    Call PropBag.WriteProperty("DisplayQueueColNames", mstrDisplayQueueFields, "")
    Call PropBag.WriteProperty("DisplayCallColNames", mstrDisplayCallFields, "")
    Call PropBag.WriteProperty("QueryQueueNames", mstrQueryQueueNames, "")
    Call PropBag.WriteProperty("GroupField", mstrGroupField, "")
    Call PropBag.WriteProperty("CustomOrderName", mstrCustomOrderColName, "")
    Call PropBag.WriteProperty("FindWay", mstrFindWay, "")
    Call PropBag.WriteProperty("Interval", mlngInterval, 30)
    Call PropBag.WriteProperty("ValidDays", mobjQueueManage.ValidDays, 1)
    Call PropBag.WriteProperty("IsShowBars", mblnIsShowBars, True)
    Call PropBag.WriteProperty("IsShowCalledQueue", mblnIsShowCalledQueue, True)
    Call PropBag.WriteProperty("IsShowToolText", mlngMenuCaptionStyle, xtpButtonIconAndCaption)
    Call PropBag.WriteProperty("IsIconLarge", cbrMain.Options.LargeIcons, True)
    Call PropBag.WriteProperty("Reasons", mstrReason, "")
    Call PropBag.WriteProperty("CalledTarget", mobjQueueManage.CallTarget, "")
End Sub


