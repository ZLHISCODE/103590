VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmCaseTendBody 
   Caption         =   "������ͼ"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBody.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   4605
      ScaleHeight     =   6825
      ScaleWidth      =   5145
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   5175
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   5370
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   915
         Width           =   5160
         _Version        =   589884
         _ExtentX        =   9102
         _ExtentY        =   9472
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton CmdRef 
         Caption         =   "ˢ��"
         Height          =   315
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "ȡ��"
         Top             =   510
         Width           =   555
      End
      Begin VB.CommandButton cmdFilterUserCancle 
         Height          =   315
         Left            =   4530
         Picture         =   "frmCaseTendBody.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "ȡ��"
         Top             =   6435
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterUserOk 
         Height          =   315
         Left            =   3990
         Picture         =   "frmCaseTendBody.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "ȷ��"
         Top             =   6435
         Width           =   450
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   14
         Top             =   510
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   107151363
         CurrentDate     =   37068
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   300
         Index           =   0
         Left            =   885
         TabIndex        =   12
         Top             =   495
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   107151363
         CurrentDate     =   37068
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000005&
         Caption         =   "ϵͳĬ����ȡ�������µ��ļ��Ĵ���ơ���Ժ��ת���ͳ�Ժ3���ڵĲ��ˣ����ڳ�Ժ���˲���Ա����ָ��ʱ�䷶Χ���й��ˡ�"
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   60
         TabIndex        =   0
         Top             =   0
         Width           =   5100
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   2265
         TabIndex        =   1
         Top             =   555
         Width           =   180
      End
      Begin VB.Label lbl��Ժʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   548
         Width           =   705
      End
   End
   Begin VB.PictureBox picCondition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   840
      ScaleHeight     =   345
      ScaleWidth      =   7755
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   165
      Width           =   7755
      Begin VB.PictureBox pic��ʶ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3510
         ScaleHeight     =   345
         ScaleWidth      =   2775
         TabIndex        =   6
         Top             =   0
         Width           =   2775
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��123456"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   1065
         End
         Begin VB.Label lbl���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����������������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1260
            TabIndex        =   8
            Top             =   60
            Width           =   2040
         End
      End
      Begin VB.PictureBox pic���� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   810
         ScaleHeight     =   315
         ScaleWidth      =   1725
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1755
         Begin VB.TextBox txt���� 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   15
            TabIndex        =   5
            Top             =   90
            Width           =   1335
         End
         Begin VB.Image img�����б� 
            Height          =   360
            Left            =   1350
            Picture         =   "frmCaseTendBody.frx":13DE
            Tag             =   "���������������µ��ļ��Ĳ����б�"
            Top             =   -30
            Width           =   360
         End
      End
      Begin VB.PictureBox picסԺ���� 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6330
         ScaleHeight     =   225
         ScaleWidth      =   1335
         TabIndex        =   3
         Top             =   60
         Width           =   1365
         Begin VB.ComboBox cboPages 
            Height          =   315
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   -30
            Width           =   1425
         End
      End
      Begin VB.Label lbl��λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   10
         Top             =   90
         Width           =   720
      End
      Begin VB.Image img��һ�� 
         Height          =   360
         Left            =   2580
         Picture         =   "frmCaseTendBody.frx":1AE0
         Tag             =   "��һ������"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image img��һ�� 
         Height          =   360
         Left            =   2940
         Picture         =   "frmCaseTendBody.frx":21E2
         Tag             =   "��һ������"
         Top             =   0
         Width           =   360
      End
   End
   Begin zl9TemperatureChartSC.usrBodyEditor BodyEdit 
      Height          =   4425
      Left            =   255
      TabIndex        =   15
      Top             =   840
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   7805
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   7080
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBody.frx":28E4
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16642
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
   Begin MSComctlLib.ImageList imgRPT 
      Left            =   240
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBody.frx":3176
            Key             =   "woman"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBody.frx":99D8
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   195
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBody"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************
'���˻�����Ϣ
'***************************************************************
Private Type type_Patient
    lng����ID As Long
    lng��ҳID As Long
    lng����ID As Long
    lng����ID As Long
    lng��Ժ As Long
    lngӤ�� As Long
    lng�༭ As Long
    lng����ȼ� As Long
    lng�ļ�ID As Long
    lngԭʼ��С As Long
    lngPage As Long
End Type

Private T_Info As type_Patient    '��¼��ǰ������Ϣ

Private Enum PATI_COLUMN
    c_ѡ�� = 0
    c_ͼ�� = 1
    c_���� = 2
    c_״̬ = 3
    c_���� = 4
    C_����ID = 5
    c_��ҳID = 6
    c_���� = 7
    c_���� = 8
    c_סԺ�� = 9
    c_��Ժ���� = 10
    c_��Ժ���� = 11
End Enum

Private mblnChildForm As Boolean
Private mcbrToolBar As CommandBar
Private mcbr�鿴 As CommandBarControl
Private mstrPrivs As String
Private mstrSQL As String
Private mblnShowing As Boolean
Private mblnChanged As Boolean
Private mfrmMain As Form
Private mIntDataEditor As Integer
Private mblnMove As Boolean
Private mfrmTendBody As Object '���µ�����
Private mintChange As Integer '�������ת������
Private mdtOutEnd As String '������Ժ��ʾ��ֹʱ��
Private mdtOutBegin As String '������Ժ��ʾ��ʼʱ��
Private mrsPati As New ADODB.Recordset
Private mintPrePage As Integer

Public Event AfterPrint()
Public Event CmdClick(ByVal strParam As String)

'######################################################################################################################
'�Զ��庯������������

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim blnShowing As Boolean
    
    mblnMove = False
    mblnChildForm = False
    mblnChanged = False
    mstrPrivs = strPrivs
    
    blnShowing = mblnShowing
    
    mblnShowing = True
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    On Error GoTo ErrHand
    
    If blnShowing Then
        If Val(varParam(0)) = T_Info.lng����ID Or Val(varParam(1)) = T_Info.lng��ҳID And T_Info.lng����ID = Val(varParam(2)) Then
            Call ShowWindow(Me.hWnd, SW_RESTORE)
            Call BringWindowToTop(Me.hWnd)
            Exit Function
        End If
    End If
    
    Set BodyEdit.ParentForm = Me
    Set mfrmMain = frmMain

    '������ʽ������ID;��ҳID;����ID;�ļ�ID;��Ժ;�༭;Ӥ��;�Ƿ���ߴ����С�Զ�У�����µ���ʽ(1 �� 0 У��)ҳ��(Ĭ����ʾ�ڼ�ҳ,���ҳ�ų�����Χ�Ͱ�ȱʡ��ʾ,0��ȱʡ��ʾ)
    
    '��ʼ������
    
    T_Info.lngӤ�� = 0
    T_Info.lng��Ժ = 0
    T_Info.lng�༭ = 0
    T_Info.lngԭʼ��С = 0
    T_Info.lngPage = 0
    
    T_Info.lng����ID = Val(varParam(0))
    T_Info.lng��ҳID = Val(varParam(1))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng�ļ�ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng��Ժ = Val(varParam(4))
    If UBound(varParam) > 4 Then
        T_Info.lng�༭ = Val(varParam(5))
    Else
        If InStr(1, ";" & mstrPrivs & ";", ";���µ���ͼ;") = 0 Then
            T_Info.lng�༭ = 0
        Else
            T_Info.lng�༭ = 1
        End If
    End If
    If UBound(varParam) > 5 Then T_Info.lngӤ�� = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lngԭʼ��С = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    mintPrePage = T_Info.lng��ҳID
    
    Set RS = New ADODB.Recordset
    If blnShowing = False Then
        Call InitMenuBar
        '���µ�ԭʼ��С�����л�����
        Call RefreshPatiList(RS)
        Call AddPages
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ��Ժ����ID,nvl(����ת��,0) ת��  from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����ID, T_Info.lng��ҳID)
    If RS.BOF = False Then
        T_Info.lng����ID = Val(zlCommFun.Nvl(RS("��Ժ����ID").Value))
        If T_Info.lng��Ժ = 1 Then mblnMove = (Val(RS("ת��")) <> 0)
    End If
    
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    
    If blnShowing = False Then
        Hook Me
        
        If bytMode = 1 Then
            Me.Show , mfrmMain
        Else
            Me.Show 1, mfrmMain
        End If
        ShowEdit = mblnChanged
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlInit() As Boolean
    mblnChildForm = True
End Function

Public Function GetCurvePage() As Long
   GetCurvePage = BodyEdit.intPage
End Function

Public Sub zlDataEditor(ByVal intDataEditor As Integer)
    BodyEdit.DateEditor = intDataEditor
End Sub

Public Function zlRefresh(ByVal frmParent As Form, strParam As String, Optional strPrivs As String) As Boolean

   '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim varParam As Variant
    Dim strPar As String
    Dim strTmp As String
    Dim intBaby As Integer
    
    mblnMove = False
    mstrPrivs = strPrivs
    mblnChildForm = True
    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = False
    picCondition.Visible = False
    picCondition.Enabled = False
    cbsThis.RecalcLayout
    
    mblnChanged = False
    
    Set BodyEdit.ParentForm = frmParent
    
    If strParam <> "" Then varParam = Split(strParam, ";")
    
    On Error GoTo ErrHand
    
    '������ʽ������ID;��ҳID;����ID;�ļ�ID;��Ժ;�༭;Ӥ��;�Ƿ���ߴ����С�Զ�У�����µ���ʽ(1 �� 0У��);ҳ��(Ĭ����ʾ�ڼ�ҳ,���ҳ�ų�����Χ�Ͱ�ȱʡ��ʾ,0��ȱʡ��ʾ)
    If Val(varParam(3)) <> T_Info.lng�ļ�ID Then
        glngCurPage = 0
    Else
        If UBound(varParam) > 5 Then
            intBaby = Val(varParam(6))
        Else
            intBaby = 0
        End If
        
        If T_Info.lngӤ�� <> intBaby Then
            glngCurPage = 0
        End If
    End If
    
    '��ʼ������
    T_Info.lngӤ�� = 0
    T_Info.lng��Ժ = 0
    T_Info.lng�༭ = 0
    T_Info.lngԭʼ��С = 0
    T_Info.lngPage = 0
    
    T_Info.lng����ID = Val(varParam(0))
    T_Info.lng��ҳID = Val(varParam(1))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng����ID = Val(varParam(2))
    T_Info.lng�ļ�ID = Val(varParam(3))
    
    If UBound(varParam) > 3 Then T_Info.lng��Ժ = Val(varParam(4))
    If UBound(varParam) > 4 Then
        T_Info.lng�༭ = Val(varParam(5))
    Else
        If InStr(1, ";" & mstrPrivs & ";", ";���µ���ͼ;") = 0 Then
            T_Info.lng�༭ = 0
        Else
            T_Info.lng�༭ = 1
        End If
    End If
    If UBound(varParam) > 5 Then T_Info.lngӤ�� = Val(varParam(6))
    If UBound(varParam) > 6 Then T_Info.lngԭʼ��С = Val(varParam(7))
    If UBound(varParam) > 7 Then
        T_Info.lngPage = Val(varParam(8))
    Else
        T_Info.lngPage = glngCurPage
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select ��Ժ����ID,nvl(����ת��,0) ת�� from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����ID, T_Info.lng��ҳID)
    If RS.BOF = False Then
        T_Info.lng����ID = Val(zlCommFun.Nvl(RS("��Ժ����ID").Value))
        If T_Info.lng��Ժ = 1 Then mblnMove = (Val(RS("ת��")) <> 0)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If OpenPatientMap = False Then
        Unload Me
        Exit Function
    End If
    
    Call GetTendEidor
    
    zlRefresh = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPatientMap() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim RS As New ADODB.Recordset
    Dim strParam As String
    
    On Error GoTo ErrHand
    
    T_Info.lng����ȼ� = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����ID, T_Info.lng��ҳID)
    If RS.BOF = False Then T_Info.lng����ȼ� = zlCommFun.Nvl(RS("����ȼ�"), 3)
    
    '������ȡ�ļ�ID
    gstrSQL = "select A.ID from ���˻����ļ� A,�����ļ��б� B" & _
       "    where A.����ID=[1] and A.��ҳId=[2] and A.Ӥ��=[3] and A.����ID=[4] and A.��ʽID=B.ID and B.����=3 and B.����=-1"
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lngӤ��, T_Info.lng����ID)
    If mblnMove = True Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    
    If RS.BOF = False Then T_Info.lng�ļ�ID = Val(zlCommFun.Nvl(RS("ID")))
    '��ʼ�����߲˵�
    If InitBodyLine = False Then Exit Function
    
    '����������ID;��ҳID;����ID;�ļ�ID;��Ժ��־;�༭��־;Ӥ��;����ȼ�;ԭʼ��С;ҳ��(Ĭ����ʾ�ڼ�ҳ,���ҳ�ų�����Χ�Ͱ�ȱʡ��ʾ,0��ȱʡ��ʾ)
    strParam = T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����ID & ";" & T_Info.lng�ļ�ID & ";" & _
        T_Info.lng��Ժ & ";" & T_Info.lng�༭ & ";" & T_Info.lngӤ�� & ";" & T_Info.lng����ȼ� & ";" & T_Info.lngԭʼ��С & ";" & T_Info.lngPage
    Call BodyEdit.zlMenuClick("��ʼ��", strParam)
        
    OpenPatientMap = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitBodyLine() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    '--�������ü��
    mstrSQL = "SELECT A.��¼��,A.��Ŀ��� FROM ���¼�¼��Ŀ A,�����¼��Ŀ B " & _
            "WHERE A.��¼�� =1 And A.��Ŀ���=B.��Ŀ��� AND B.����ȼ�>=[1]  And Nvl(b.Ӧ�÷�ʽ,0)=1 " & _
            "ORDER BY A.�������"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng����ȼ�)
    If rsTmp.BOF Then
        MsgBox "�����µ�����������Ŀ�����ڻ�����Ŀ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '--��¼Ƶ��ʱ������ü��
    mstrSQL = " Select Distinct nvl(��¼Ƶ��,2) Ƶ��  From ���¼�¼��Ŀ A,�����¼��Ŀ B" & _
            "   WHERE A.��¼�� =2 And A.��Ŀ���<>3 And  ��Ŀ��ʾ<>4 And A.��Ŀ���=B.��Ŀ��� AND B.����ȼ�>=[1] And Nvl(b.Ӧ�÷�ʽ,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, T_Info.lng����ȼ�)
    
    Do While Not rsTmp.EOF
        strSQL = "select Count(*) ��¼�� From ������ĿƵ�� where Ƶ��=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!Ƶ��))
        If Val(rsData!��¼��) < Val(rsTmp!Ƶ��) Then
            MsgBox "������Ŀ��¼Ƶ��ʱ�����ò����������ڻ�����Ŀ���������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    rsTmp.MoveNext
    Loop
    '--������Ŀʱ������ü��
    mstrSQL = "select count(*) ��¼�� from �������ʱ�� WHERE ����=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    If Val(rsTmp!��¼��) < 3 Then
        MsgBox "�������ʱ�����ò����������ڻ�����Ŀ���������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    InitBodyLine = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "") As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnCur As Boolean
    Dim lngBeginY As Long
    Dim intBeginPage As Integer
    Dim intPrintRange As Integer
    Dim strPage  As String, strParam As String
    
    '�����˴�ӡ������,˵����������ӡ,�Զ��ӵ�1ҳ��ʼ��ӡ,�������κ�ѯ��
    '����:0-ȡ��,2-Ԥ��,1-��ӡ
    
    frmCaseTendBodyPrintSet.cmdPrint.Visible = (bytMode = 1)
    frmCaseTendBodyPrintSet.cmdPreview.Visible = (bytMode = 2)
    
    
    If strPrintDevice = "" Then
        'strParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����Id & ";" & T_Info.lng����Id
        strParam = T_Info.lng�ļ�ID & ";" & Me.BodyEdit.AllPage
        bytMode = frmCaseTendBodyPrintSet.PrintSet(Me, True, strParam, intPrintRange, lngBeginY, intBeginPage, strPage, mstrPrivs)
    Else
        bytMode = 2
        intPrintRange = 2
    End If
    If bytMode = 0 Then Exit Function
    If intBeginPage <= 0 Then intBeginPage = -1
    
    '��ӡ��ǰҳ���뵱ǰҳ��
    If intPrintRange = 0 Then
        strPage = Me.BodyEdit.intPage - 1
    End If
    
    Select Case bytMode
    Case 2  '��ӡ
        Call BodyEdit.PrintState(intPrintRange, True, lngBeginY, intBeginPage, strPrintDevice, strPage)
    Case 1  'Ԥ��
        Call BodyEdit.PrintState(intPrintRange, False, lngBeginY, intBeginPage, strPrintDevice, strPage)
    End Select

End Function

Public Function zlPrintBody(Optional ByVal bytMode As Byte = 2, Optional ByVal strPrintDevice As String) As Long
    '���:1-Ԥ��,2-��ӡ
    '����ֵ:0-ʧ��;1-�ɹ�;2-��ӡ
    gblnPrinted = False
    Call PrintData(IIf(bytMode = 1, 2, 1), strPrintDevice)
    zlPrintBody = IIf(gblnPrinted, 2, 1)
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim objCustom As CommandBarControlCustom
    
    On Error GoTo ErrHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
       
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "��������(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "�ָ�����(&R)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve, "���߱༭(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CurveTable, "���༭(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "������ʾ(&D)")
    End With

    Set mcbr�鿴 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    With mcbr�鿴.CommandBar.Controls
                
'       Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
'
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
                
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    picCondition.Visible = True
    picCondition.Enabled = True
    
    '����������
    Set mcbrToolBar = cbsThis.Add("����������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0

    With mcbrToolBar.Controls
          Set objCustom = .Add(xtpControlCustom, 1, "")
          objCustom.Handle = picCondition.hWnd
    End With
    
    '��λ������
    '------------------------------------------------------------------------------------------------------------------

    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("Q"), conMenu_Edit_Curve
        .Add FCONTROL, Asc("T"), conMenu_Edit_CurveTable
        .Add FCONTROL, Asc("D"), conMenu_Edit_Curve_Show
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    Call InitReprotControl '��ʼ��������Ϣ�б�
    
    InitMenuBar = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub BodyEditCur(ByVal intDataEditor As Integer, Optional ByVal strParam As String = "")
    Call GetTendEidor
    If intDataEditor = 0 Or intDataEditor = -1 Then
        gintEditorCurveState = intDataEditor
        Call BodyEdit.zlMenuClick("�������ݱ༭", strParam)
    ElseIf intDataEditor = 1 Then
         Call BodyEdit.zlMenuClick("����������ʾ����", strParam)
    End If
End Sub

Private Sub BodyEdit_DbClickCur(ByVal intDataEditor As Integer)
    Call BodyEditCur(intDataEditor)
End Sub

Private Sub cboPages_Click()
    If cboPages.ListIndex = -1 Then Exit Sub
    If Val(cboPages.ItemData(cboPages.ListIndex)) = mintPrePage Then Exit Sub
    mintPrePage = Val(cboPages.ItemData(cboPages.ListIndex))
    T_Info.lng��ҳID = cboPages.ItemData(cboPages.ListIndex)
    
    Call GetPatiInfo
    'ˢ������
    T_Info.lngӤ�� = 0: T_Info.lngPage = 0
    Call OpenPatientMap
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo ErrHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.Id
        Case conMenu_File_PrintSet   '��ӡ����
            
            On Error Resume Next
            Call frmPrintSet.ShowMe(Me, 1)
            
        Case conMenu_File_Preview  '��ӡԤ��
            
            Call PrintData(2)
            
        Case conMenu_File_Print  '��ӡ
        
            Call PrintData(1)
        
        Case conMenu_View_ToolBar_Button

'            cbsThis(2).Visible = Not cbsThis(2).Visible
'            cbsThis.RecalcLayout

        Case conMenu_View_ToolBar_Text

'            For Each cbrControl In cbsThis(1).Controls
'                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
'            Next
'
'            cbsThis.RecalcLayout
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
            
        Case conMenu_Edit_Curve '���߱༭
            Call BodyEditCur(0)
        Case conMenu_Edit_CurveTable '���༭
            Call BodyEditCur(-1)
        Case conMenu_Edit_Curve_Show '������ʾ
            Call BodyEditCur(1)
            
        Case conMenu_Edit_Save '��������
            
        Case conMenu_Edit_Reuse '���ݻָ�
            
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.hWnd)
            
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hWnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.Id

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Curve, conMenu_Edit_CurveTable, conMenu_Edit_Curve_Show
        
        Control.Enabled = (T_Info.lng�༭ = 1)
        
    Case conMenu_View_ToolBar_Button
    
        Control.Checked = Me.cbsThis(2).Visible
        
    Case conMenu_View_ToolBar_Text
    
        Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        
    Case conMenu_View_ToolBar_Size
    
        Control.Checked = Me.cbsThis.Options.LargeIcons
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
        
    End Select
End Sub

Private Sub BodyEdit_zlAfterPrint()
    gblnPrinted = True
    RaiseEvent AfterPrint
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With BodyEdit
        .mblnResize = True
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .mblnResize = False
        .Height = lngBottom - lngTop
    End With
    picCondition.Width = Me.picסԺ����.Left + Me.picסԺ����.Width + 100
End Sub

Private Sub cmdFilterUserCancle_Click()
    picPati.Visible = False
End Sub

Private Sub CmdRef_Click()
    Dim RS As New ADODB.Recordset
    Call RefreshPatiList(RS)
    Call img�����б�_MouseDown(1, 0, 0, 0)
End Sub

Private Sub dtpB_Change(Index As Integer)
'ʱ�䷶Χ�ı�ʱˢ��
    If dtpB(Index).Value >= dtpE(Index).Value Then
        MsgBox "��Ժʱ�䷶Χ�Ŀ�ʼʱ��ӦС�ڽ���ʱ��", vbInformation, gstrSysName
        dtpB(Index).Value = dtpB(Index).Tag
        dtpB(Index).SetFocus: Exit Sub
    Else
        dtpB(Index).Tag = dtpB(Index).Value
        If Index = 0 Then mdtOutBegin = dtpB(Index).Value
    End If
End Sub

Private Sub dtpE_Change(Index As Integer)
    If dtpB(Index).Value >= dtpE(Index).Value Then
        MsgBox "��Ժʱ�䷶Χ�Ŀ�ʼʱ��ӦС�ڽ���ʱ��", vbInformation, gstrSysName
        dtpE(Index).Value = dtpE(Index).Tag
        dtpE(Index).SetFocus: Exit Sub
    Else
        dtpE(Index).Tag = dtpE(Index).Value
        If Index = 0 Then mdtOutEnd = dtpE(Index).Value
    End If
End Sub

Private Sub Form_Load()
    Call GetLocalSetting '��ȡ��ز���
    If Not mblnChildForm Then
         Call RestoreWinState(Me, App.ProductName)
    End If
End Sub

Private Sub GetTendEidor()
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    Set gobjTendEditor = Me
End Sub

Private Sub BodyEdit_CmdClick(ByVal strParam As String)
    Dim arrParam() As String
    If mfrmTendBody Is Nothing Then Set mfrmTendBody = New frmCaseTendBody
    
    If mfrmTendBody.ShowEdit(BodyEdit.ParentForm, strParam, 0, mstrPrivs) Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) > 6 Then arrParam(7) = 0
        If UBound(arrParam) > 7 Then
            strParam = arrParam(0) & ";" & arrParam(1) & ";" & arrParam(2) & ";" & arrParam(3) & ";" & arrParam(4) & ";" & arrParam(5) & ";" & arrParam(6) & ";" & arrParam(7)
        Else
            strParam = Join(arrParam, ";")
        End If
        
        Call zlRefresh(BodyEdit.ParentForm, strParam, mstrPrivs)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHook Me
    
    mblnShowing = False
    Set mfrmTendBody = Nothing
    
    If Not mblnChildForm Then
        Call SaveWinState(Me, App.ProductName)
        mblnChanged = True
    End If
    If Not gobjTendEditor Is Nothing Then Set gobjTendEditor = Nothing
    
    Set mcbrToolBar = Nothing
    Set mcbr�鿴 = Nothing
    Set mrsPati = Nothing
    'ж���û��ؼ����� ������ر�ʱ�û��ؼ��� UserControl_Terminate �¼��޷����� ���Է��ڸ�����ر�ִ�� ��
    Call BodyEdit.ReleaseObj
End Sub

Private Sub img�����б�_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngColor As Long
    Dim lngLoop As Long
    Dim objRow As ReportRow
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strPatient As String '�����б���Ϣ
    Dim lngRow As Long, lngID As Long 'VSFѡ��Ĳ���ID
    Dim lngLeft As Long, lngTop  As Long, lngRight As Long, lngBottom As Long
    Dim ArrCode() As String
    Dim blnVisible As Boolean
    
    If Button <> 1 Then Exit Sub
    
    On Error GoTo ErrHand
    
    If rptPati.Records.Count = 0 And mrsPati.RecordCount > 0 Then
        '��ʾ�����б�ѡ��
        With mrsPati
            .MoveFirst
            
            Do While Not .EOF
                Set objRecord = rptPati.Records.Add()
                objRecord.Tag = CStr(!����ID & "," & !��ҳID)
                Set objItem = objRecord.AddItem("")
                objItem.HasCheckbox = True
                objItem.Checked = False
                
                Set objItem = objRecord.AddItem(""): objItem.Icon = IIf(!�Ա� = "��", 1, 0)
                Set objItem = objRecord.AddItem(CStr(!����))
                objItem.Caption = CStr(!���� & !����)
                Set objItem = objRecord.AddItem(CStr(!���� & !����))
                objItem.Caption = CStr(!���� & !����)
                
                Set objItem = objRecord.AddItem(LPAD(Nvl(!����), 10, " "))
                objItem.Caption = Trim(Nvl(!����, " "))
                objRecord.AddItem Val(!����ID)
                objRecord.AddItem Val(!��ҳID)
                objRecord.AddItem CStr(Nvl(!����))
                objRecord.AddItem CStr(Nvl(!����))
                Set objItem = objRecord.AddItem(CStr(Nvl(!סԺ��)))
                objItem.Caption = Nvl(!סԺ��, " ")
                
                Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd"))
                objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd")
                Set objItem = objRecord.AddItem(Format(!��Ժ����, "yyyy-MM-dd"))
                objItem.Caption = Format(!��Ժ����, "yyyy-MM-dd")
                
                '��ȡ�������͵���ɫ
                lngColor = Nvl(!��ɫ, 0)
                If lngColor <> 0 Then objRecord.Item(c_����).ForeColor = lngColor
                
                .MoveNext
            Loop
        End With
    End If
    
    mrsPati.MoveFirst
    mrsPati.Find "����ID=" & T_Info.lng����ID
    
    Call mcbrToolBar.GetWindowRect(lngLeft, lngTop, lngRight, lngBottom)
    rptPati.Populate 'ȱʡ��ѡ���κ���
    picPati.Left = picCondition.Left + Me.pic����.Left
    picPati.Top = lngTop - Me.Top - 60
    picPati.Visible = True
    
    'ѡ�е�ǰ����(���۵���Ļ�,Rows.Countֻ����ĸ�����,�����ȶ�λ,���۵�)
    For lngLoop = 0 To rptPati.Rows.Count - 1
        If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
            If Val(rptPati.Rows(lngLoop).Record.Item(C_����ID).Value) = T_Info.lng����ID Then
                Set rptPati.FocusedRow = rptPati.Rows(lngLoop)
                Exit For
            End If
        End If
    Next
    
    '�۵�������(ѡ�в�����һ�鲻�۵�)
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
            objRow.Expanded = False
        End If
    Next
    rptPati.FocusedRow.EnsureVisible
    rptPati.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub img�����б�_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo Me.pic����.hWnd, img�����б�.Tag
End Sub

Private Sub img��һ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(1)
End Sub

Private Sub img��һ��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hWnd, img��һ��.Tag
End Sub

Private Sub img��һ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(2)
End Sub

Private Sub img��һ��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hWnd, img��һ��.Tag
End Sub

Private Sub lbl����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo pic��ʶ.hWnd, lbl����.Caption
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdFilterUserCancle_Click
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If rptPati.Records.Count = 0 Then Exit Sub
    If rptPati.FocusedRow.Record Is Nothing Then Exit Sub
    
    T_Info.lng����ID = Split(rptPati.FocusedRow.Record.Tag, ",")(0)
    '�����Ҫ���˶�λ����һ��,��һ��ʱ����λǰ��˳��,�ɰѸ�������ε�
    mrsPati.MoveFirst
    mrsPati.Find "����ID=" & T_Info.lng����ID
    
    picPati.Visible = False
    txt����.Text = ""
    mintPrePage = -1
    Call AddPages
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    strInput = Trim(txt����.Text)
    If strInput = "" Then Exit Sub
    
    strInput = " ����='" & LPAD(strInput, 10, " ") & "'"
    mrsPati.Filter = strInput
    If mrsPati.RecordCount = 0 Then
        If Not IsNumeric(Trim(txt����.Text)) Then
            strInput = " ����='" & Trim(txt����.Text) & "'"
        Else
            strInput = " סԺ��=" & Trim(txt����.Text)
        End If
        mrsPati.Filter = strInput
        
        If mrsPati.RecordCount = 0 Then
            mrsPati.Filter = 0
            MsgBox "δ�ҵ��ò��˵���Ч���ݣ����������룡", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    T_Info.lng����ID = mrsPati!����ID
    mrsPati.Filter = 0
    mrsPati.MoveFirst
    mrsPati.Find "����ID=" & T_Info.lng����ID
    
    mintPrePage = -1
    Call AddPages
    
    picPati.Visible = False
End Sub


Private Sub LocatePati(ByVal intType As Integer)
    '����˵��:intType:1-��һ������;2-��һ������
    '���˷�Χ:�ڴ�����ѭ��,���ϰ汣��һ��
    Dim blnExit As Boolean  'ǿ���˳�
    On Error Resume Next
    
redo:
    If intType = 1 Then
        mrsPati.MovePrevious
        If mrsPati.BOF Then mrsPati.MoveLast
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then mrsPati.MoveFirst
    End If
    If mrsPati!����ID <> 0 Then
        If mrsPati!����ID <> T_Info.lng����ID Then
            T_Info.lng����ID = mrsPati!����ID
            
            mintPrePage = -1
            Call AddPages
        Else
            If blnExit Then Exit Sub
            blnExit = True
            GoTo redo
        End If
    Else
        GoTo redo
    End If
    
    picPati.Visible = False
End Sub

Private Sub AddPages()
    Dim i As Integer, j As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    '���ݲ���ID��ȡ�ò��˵�סԺ����

    strSQL = " Select סԺ���� From ������Ϣ Where ����ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡסԺ����", T_Info.lng����ID)
    j = Nvl(rsTemp!סԺ����, 0)

    Me.cboPages.Clear
    For i = j To 1 Step -1
        Me.cboPages.AddItem "��" & i & "��סԺ"
        Me.cboPages.ItemData(Me.cboPages.NewIndex) = i
        If mintPrePage = i Then
            cboPages.ListIndex = cboPages.NewIndex
        End If
    Next
    If cboPages.ListCount > 0 And cboPages.ListIndex = -1 Then Me.cboPages.ListIndex = 0 '��λ�����һ��סԺ
End Sub

Private Sub GetPatiInfo()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo ErrHand
    
    strSQL = "Select A.����,B.��Ժ���� ����, C.��ɫ" & vbNewLine & _
        " From ������Ϣ A, ������ҳ B,�������� C" & vbNewLine & _
        " Where A.����ID=B.����ID And B.����ID=[1] And B.��ҳID=[2] And B.��������=C.����(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", T_Info.lng����ID, T_Info.lng��ҳID)
    
    lbl����.Caption = "��" & Trim(Nvl(rsTemp!����))
    lbl����.Caption = Nvl(rsTemp!����)
    lbl����.ForeColor = Nvl(rsTemp!��ɫ, 0)

    Me.pic��ʶ.Width = lbl����.Width + lbl����.Left
    Me.picסԺ����.Width = Me.cboPages.Width - 50
    Me.picסԺ����.Left = pic��ʶ.Left + pic��ʶ.Width + 50
    picCondition.Width = Me.picסԺ����.Left + Me.picסԺ����.Width + 100
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub GetLocalSetting()
'���ܣ���ע����ȡ��Ժ���˵�ʱ�䷶Χ
    Dim i As Integer
    Dim curDate As Date, intDay As Integer

    '������ʾ��Χ
    mintChange = Val(zlDatabase.GetPara("���ת������", glngSys, pסԺ��ʿվ, 7))
    '�������30���ȡȱʡֵ
    If mintChange > 30 Then mintChange = 7
    
    '��Ժ����ʱ�䷶Χ
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd")
    mdtOutBegin = Format(CDate(mdtOutEnd) - 3, "yyyy-MM-dd")
    dtpE(0).Value = mdtOutEnd
    dtpE(0).Tag = mdtOutEnd
    dtpB(0).Value = mdtOutBegin
    dtpB(0).Tag = mdtOutBegin
End Sub

Public Sub RefreshPatiList(Optional ByVal rsThis As ADODB.Recordset)
    On Error GoTo ErrHand
    
    'ˢ�²����嵥,�Զ�λ����ǰ�����Ĳ�����
    Call LoadPatient(rsThis)
    mrsPati.MoveFirst
    mrsPati.Find ("����ID=" & T_Info.lng����ID)
    rptPati.Records.DeleteAll
    Call GetPatiInfo
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadPatient(ByVal rsThis As ADODB.Recordset)
    Dim strSQL As String
    On Error GoTo ErrHand
    '58890:������,2013-02-26,��Ժ���˶�ȡ�����Ż�(������Ժ���˱���в�ѯ)
    '��Ժ����ƺ�ת�ƴ���Ʋ���(���˿��������Ĳ������ɽ���)
    'c.����id + 0,˵����ͨ��H����������ӹ��˺󣬼�¼�������٣�������B�������
    If rsThis Is Nothing Then
ErrGO:
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.״̬,1,0,Decode(c.��ʼԭ��,3,1,2)) As ����, Decode(Nvl(b.����״̬, 0), 0, 999, b.����״̬) As ����2," & _
            " Decode(B.״̬,1,'��Ժ����ס����',Decode(c.��ʼԭ��,3,'ת�ƴ���ס����','ת��������ס����')) As ����," & _
            " a.����id, b.��ҳid, A.�����,B.סԺ��, a.����, a.�Ա�, b.����," & vbNewLine & _
            " d.���� As ����, c.����id, c.����ҽʦ As סԺҽʦ,b.���λ�ʿ, b.����״̬, lpad(nvl(C.����,' '),10,' ') as ����," & _
            " e.���� As ����ȼ�, b.�ѱ�,b.��ǰ����, b.��Ժ����, b.��Ժ����,B.��Ժ��ʽ, b.��������, b.״̬, b.����, a.���￨��," & vbNewLine & _
            " -1 As ·��״̬,trunc(sysdate)-trunc(b.��Ժ����)+1 as סԺ����,Z.��ɫ" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B, ���˱䶯��¼ C, ���ű� D, �շ���ĿĿ¼ E, �������Ҷ�Ӧ H,�������� Z,��Ժ���� R" & vbNewLine & _
            "Where B.��������=Z.����(+) And A.����ID = R.����ID And a.����id = b.����id And Nvl(b.��ҳid, 0) <> 0 And b.����id = c.����id And b.��ҳid = c.��ҳid And c.����id = d.Id" & vbNewLine & _
            "      And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
            "      And b.����ȼ�id = e.Id(+) And Nvl(c.���Ӵ�λ, 0) = 0 And c.��ֹʱ�� Is Null" & vbNewLine & _
            "      And (c.��ʼԭ�� in(1,3) And c.����id + 0 = h.����id And h.����id = [1] or c.��ʼԭ��=15 And c.����id = [1])" & vbNewLine & _
            "      And ((c.��ʼԭ�� = 1 And b.״̬ = 1) Or (c.��ʼԭ�� in (3,15) And c.��ʼʱ�� Is Null And b.״̬ = 2)) "
    
        '��Ժ����
        
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.״̬,3,4,DECODE(B.��Ժ����, NULL, 3.1,DECODE(B.״̬,2,3.2,3))) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.״̬,3,'Ԥ��Ժ����',DECODE(B.��Ժ����, NULL, '��ͥ����',DECODE(B.״̬,2,'Ԥת�Ʋ���', '��Ժ����'))) as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,A.����,A.�Ա�,B.����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " lpad(nvl(B.��Ժ����,' '),10,' ') as ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,B.��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(B.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(b.��Ժ����)+1 as סԺ����,z.��ɫ" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z,��Ժ���� R" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And A.סԺ����=B.��ҳID And Nvl(B.��ҳID,0)<>0 And Nvl(B.״̬,0)<>1" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL And A.����ID=R.����ID And R.����ID=[1]"
            
        '��Ժ����:��Ժ���˿������ж��סԺ
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.��Ժ��ʽ,'����',6,5) as ����," & _
            " Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2," & _
            " Decode(B.��Ժ��ʽ,'����','��������','��Ժ����') as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,A.����,A.�Ա�,B.����,C.���� as ����,B.��Ժ����ID ����ID,B.סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " lpad(nvl(B.��Ժ����,' '),10,' ') AS ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,B.��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(B.·��״̬,-1) ·��״̬,trunc(b.��Ժ����)-trunc(b.��Ժ����)+1 as סԺ����,z.��ɫ" & _
            " From ������Ϣ A,������ҳ B,���ű� C,�շ���ĿĿ¼ E,�������� Z" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.״̬=0" & _
            " And B.��Ժ����ID=C.ID And B.����ȼ�ID=E.ID(+) And B.��ǰ����ID+0=[1] And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null)" & _
            " And B.��Ժ���� Between [2] And [3] And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
        'ת������:��Ժ,ҽ���ʹ�����ʾ����ת��ǰ��
    
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Distinct 7 as ����,Decode(Nvl(B.����״̬,0),0,999,B.����״̬) as ����2,'ת������' as ����," & _
            " A.����ID,B.��ҳID,A.�����,B.סԺ��,A.����,A.�Ա�,B.����,D.���� as ����,C.����ID,C.����ҽʦ as סԺҽʦ,B.���λ�ʿ,B.����״̬," & _
            " lpad(nvl(C.����,' '),10,' ') as ����,E.���� as ����ȼ�,B.�ѱ�,B.��ǰ����,B.��Ժ����,B.��Ժ����,B.��Ժ��ʽ,B.��������," & _
            " B.״̬,B.����,A.���￨��,Nvl(B.·��״̬,-1) ·��״̬,trunc(sysdate)-trunc(b.��Ժ����)+1 as סԺ����,z.��ɫ" & _
            " From ������Ϣ A,������ҳ B,���˱䶯��¼ C,���ű� D,�շ���ĿĿ¼ E,�������� Z" & _
            " Where B.��������=Z.����(+) And A.����ID=B.����ID And Nvl(B.��ҳID,0)<>0 And B.����ȼ�ID=E.ID(+)" & _
            " And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
            " And B.��ǰ����ID<>[1] And C.����ID+0=[1] And C.����ID=D.ID" & _
            " And Nvl(C.���Ӵ�λ,0)=0 And C.��ֹԭ�� In(3,15) And C.��ֹʱ�� Between Sysdate-[4] And Sysdate" & _
            " And Nvl(B.״̬,0)<>2 And Nvl(B.����״̬,0)<>5 And B.���ʱ�� is NULL"
    
        '��ȡ������Ϣ
        
        strSQL = "SELECT A.����,A.����2,A.����,A.����ID,A.��ҳID,A.�����,A.סԺ��,A.����,A.�Ա�,A.����,A.����,A.����ID,A.סԺҽʦ,A.���λ�ʿ,A.����״̬," & _
                " lpad(nvl(A.����,' '),10,' ') as ����,A.����ȼ�,A.�ѱ�,A.��ǰ����,A.��Ժ����,A.��Ժ����,A.��Ժ��ʽ,A.��������," & _
                " A.״̬,A.����,A.���￨��,A.·��״̬,A.סԺ����,A.��ɫ" & _
                " From (" & strSQL & ") A,���˻����ļ� B,�����ļ��б� C" & _
                " Where A.����ID=B.����ID and A.��ҳID=B.��ҳID And nvl(B.Ӥ��,0)=0 And B.��ʽID=C.ID And C.����=3 And C.����=-1"
        strSQL = strSQL & " Order by A.����,A.����,A.��ҳID DESC"
        
        Screen.MousePointer = 11
        On Error GoTo ErrHand
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����б�", T_Info.lng����ID, _
            CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), _
            mintChange)
        Screen.MousePointer = 0
    Else
        If rsThis.State = 1 Then
            Set mrsPati = rsThis.Clone
        Else
            GoTo ErrGO
        End If
    End If
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitReprotControl()
 '��ʼ������ѡ����
    Dim objCol As ReportColumn
    With rptPati
        Set objCol = .Columns.Add(c_ѡ��, "", 0, False): objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_ͼ��, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_����, "����", 0, True)
        Set objCol = .Columns.Add(c_״̬, "״̬", 0, True)
        Set objCol = .Columns.Add(c_����, "����", 40, True)
        Set objCol = .Columns.Add(C_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_����, "����", 60, True)
        Set objCol = .Columns.Add(c_סԺ��, "סԺ��", 60, True)
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", 70, True)
        Set objCol = .Columns.Add(c_��Ժ����, "��Ժ����", 70, True)
        For Each objCol In .Columns
            If objCol.Index <> c_ѡ�� Then
                objCol.Editable = False
            Else
                objCol.Sortable = True
                objCol.Editable = True
            End If
            objCol.Groupable = (objCol.Index = c_״̬)
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�в���..."
        End With
        
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgRPT
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(c_����)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_����)
    End With
End Sub

