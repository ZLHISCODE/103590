VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmPACSGate 
   AutoRedraw      =   -1  'True
   Caption         =   "Ӱ����շ���"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10995
   Icon            =   "frmPACSGate.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10995
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraUD_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   270
      MousePointer    =   7  'Size N S
      TabIndex        =   2
      Top             =   3480
      Width           =   7635
   End
   Begin VB.Timer Timer1 
      Left            =   5010
      Top             =   3270
   End
   Begin MSComctlLib.ListView lvwSeq 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "fgfg"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   735
      Top             =   2190
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
            Picture         =   "frmPACSGate.frx":058A
            Key             =   "_0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":0B24
            Key             =   "_1"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPACSGate.frx":10BE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11748
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VB.Frame fraLR_s 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   3330
      MousePointer    =   9  'Size W E
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   30
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   1335
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1952
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1B6C
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1D86
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":1FA0
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":21BA
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":28B4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":2FAE
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":36A8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":3DA2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":449C
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":4B96
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":5290
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":598A
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":6084
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":677E
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   1935
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":6E78
            Key             =   "Ԥ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":7092
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":72AC
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":74C6
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":76E0
            Key             =   "��¼"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":7DDA
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":84D4
            Key             =   "���"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":8BCE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":92C8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":99C2
            Key             =   "�ķ�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":A0BC
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":A7B6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":AEB0
            Key             =   "�޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":B5AA
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACSGate.frx":BCA4
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picView 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   8235
      TabIndex        =   4
      Top             =   3720
      Width           =   8295
      Begin DicomObjects.DicomViewer DViewer 
         Height          =   2055
         Left            =   360
         TabIndex        =   5
         Top             =   120
         Width           =   3735
         _Version        =   262147
         _ExtentX        =   6588
         _ExtentY        =   3625
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1191
      ButtonWidth     =   820
      ButtonHeight    =   1138
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgGray"
      HotImageList    =   "imgColor"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ԥ��"
            Key             =   "Ԥ��"
            Object.ToolTipText     =   "Ԥ��"
            Object.Tag             =   "Ԥ��"
            ImageKey        =   "Ԥ��"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "��ӡ"
            Key             =   "��ӡ"
            Object.ToolTipText     =   "��ӡ"
            Object.Tag             =   "��ӡ"
            ImageKey        =   "��ӡ"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "����"
            Object.ToolTipText     =   "��ǰ��������"
            Object.Tag             =   "����"
            ImageKey        =   "����"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�"
            Key             =   "�˳�"
            Object.ToolTipText     =   "�˳�"
            Object.Tag             =   "�˳�"
            ImageKey        =   "�˳�"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mmuProcessLog 
         Caption         =   "��¼������־"
      End
      Begin VB.Menu mmuCommLog 
         Caption         =   "��¼ͨѶ��־"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "��ʾͨѶ��־"
         Index           =   1
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "��ʾ������־"
         Index           =   2
      End
      Begin VB.Menu mmuShowLog 
         Caption         =   "��ʾ��ǰ����"
         Index           =   3
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolItem 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mmuUpdateDB 
         Caption         =   "�������ݿ�(&U)"
      End
      Begin VB.Menu mnuHelp_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmPACSGate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public LastState As Integer
Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mstrPrivs As String         'Ȩ���ִ�
Private mBufferDir As String
Private lngErrCounts As Long
Private mstrMWLModality As String            'worklistͨѶ��ʹ�õ�Ӱ�����
Private mintWLCount As Integer               'worklistͨѶ�������Զ������ļ�����

Private strWhere As String
Private blnNewImg As Boolean '���µ�Ӱ����Ҫˢ���б�
Private strDirURL As String, strHost As String

Private mdtLastAssociation As Date           '������յ�Association��ʱ��

'��ӡ·��
Dim DGlobal As DicomGlobal
Dim PrintRouterDss As DicomDataSets
Dim printerobject As DicomDataSet

Private WithEvents mobjIcon As clsTaskIcon  '������
Attribute mobjIcon.VB_VarHelpID = -1
Private mfrmUpdateDB As frmUpdateDB
Private mfrmShowLog As frmShowLog

Private Sub DViewer_AssociationClosed(ByVal connection As DicomObjects.DicomConnection)
    Dim Session As DicomDataSet
    
    For Each Session In connection.Tag
        subRemove Session.instanceUID
    Next
End Sub

Private Sub DViewer_AssociationRequest2(ByVal connection As DicomObjects.DicomConnection, isOK As Boolean)
    Dim context As DicomContext
    Dim strLog As String, strTmp As String
    Dim i As Integer
    Dim blnMatch As Boolean     '�Ƿ������ķ����ƥ��
    Dim blnNext As Boolean
    
    On Error GoTo ProcError
    
    '��¼���յ����������ʱ��
    mdtLastAssociation = Time
    
    '�����ӡ·�ɣ�Ԥ�ȸ�connection��tag����һ�������ݼ�
    Set connection.Tag = New DicomDataSets
    
    '�����ݶϿ�ʱ��������
    CheckDBConnect
    
    strLog = "�����AE�ǣ�" & connection.CallingAET & _
        ",�����IP�ǣ�" & connection.RemoteIP & ",�����е�AE�ǣ�" & connection.CalledAET
    WriteCommLog "AssociationRequest", "���յ�ͨѶ����", strLog
    
    '�жϸ������Ƿ��Ƿ�������������
    '���ڡ���ӡ·�ɡ��͡���Ƭ���ա�����ֻ�ж�CalledAE
    '����ͼ����ա�worklist��Q/R�����ж��豸IP,�豸AE,����AE��
    
    For i = 1 To UBound(Services)
        If UCase(Services(i).ServiceAE) = UCase(connection.CalledAET) _
            And (UCase(Services(i).SOP) = "PRINT" Or Services(i).SOP = "��Ƭ����") Then
            blnMatch = True
            Exit For
        ElseIf Services(i).DeviceIP = connection.RemoteIP And UCase(Services(i).DeviceAE) = UCase(connection.CallingAET) _
            And UCase(Services(i).ServiceAE) = UCase(connection.CalledAET) Then
            blnMatch = True
            Exit For
        End If
    Next i
    
    If blnMatch = False Then
        isOK = False
        strLog = "����ķ��������֧�ֵķ���ƥ�䣬����AE���ƣ����󱻾ܾ��������AE�ǣ�" & connection.CallingAET & _
        ",�����IP�ǣ�" & connection.RemoteIP & ",�����е�AE�ǣ�" & connection.CalledAET
        WriteCommLog "AssociationRequest", "�ܾ���������ķ���", strLog
        Exit Sub
    End If
        
    '�ܾ����з�DICOM SOP �����������
    strLog = ""
    For Each context In connection.Contexts
        If Left(context.AbstractSyntax, 14) <> "1.2.840.10008." Then
'            context.Reject 3
'            WriteCommLog "AssociationRequest", "���󱻾ܾ�", "������﷨Ϊ��" & _
'                context.AbstractSyntax & ",������1.2.840.10008��"
        Else
            strLog = strLog & ": " & context.AbstractSyntax
            '����ͼ��洢�����Association
            'ֻ�жϵ�һ��������
            '�������Q/R��Worklist��Verify֮������ӣ���������Ӧ�þ���ͼ��洢����
            'Q/R�ĳ����﷨Ϊ ��1.2.840.10008.5.1.4.1.2.x.x��
            'Worklist�ĳ����﷨Ϊ ��1.2.840.10008.5.1.4.31��
            'Verify�ĳ����﷨Ϊ ��1.2.840.10008.1.1��
            'Print �� BasicGrayScalePrint �ĳ����﷨Ϊ ��1.2.840.10008.5.1.1.9��
            If blnNext = False Then
                If context.AbstractSyntax <> "1.2.840.10008.5.1.4.31" And context.AbstractSyntax <> "1.2.840.10008.1.1" _
                    And Left(context.AbstractSyntax, 24) <> "1.2.840.10008.5.1.4.1.2." Then
                    '�������Ӳ���
                    subSaveAssociation connection
                    blnNext = True
                End If
            End If
        End If
    Next context
    If strLog <> "" Then
        WriteCommLog "AssociationRequest", "��������", "����������﷨Ϊ��" & strLog
    End If
    
    Exit Sub
ProcError:
    On Error Resume Next
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog(1, err.Number, err.Description)
End Sub

Private Sub DViewer_ImageReceived(ByVal ReceivedImage As DicomObjects.DicomImage, ByVal Association As Long, isAdded As Boolean, Status As Long)
    Dim blnReceived As Boolean
    
    On Error GoTo ProcError
    
    '��¼���յ�ͼ���ʱ��
    mdtLastAssociation = Time
    
    ReceivedImage.Tag = Association
    '�����ݶϿ�ʱ��������
    CheckDBConnect
    If ImageExist(DViewer.Images, ReceivedImage) Then
        isAdded = False: Status = 0
        Exit Sub
    End If
    
    blnReceived = True
    '��һ��ͼ����Ҫ���⴦��������Intera MR��һ��ͼ����Ϊ�޷��������Բ�����
    If Not IsNull(ReceivedImage.Attributes(&H8, &H60).value) Then
        If UCase(ReceivedImage.Attributes(&H8, &H60).value) = "MR" And Not IsNull(ReceivedImage.Attributes(&H8, &H16).value) Then
            If Left(ReceivedImage.Attributes(&H8, &H16).value, Len(ReceivedImage.Attributes(&H8, &H16).value) - 1) = "1.3.46.670589.11.0.0.12." Or _
                ReceivedImage.Attributes(&H8, &H16).value = "1.2.840.10008.5.1.4.1.1.66" Then
                  '����ΪMR�ģ������Ե�֪�����Sop Class UID ="1.3.46.670589.11.0.0.12.2"��"1.3.46.670589.11.0.0.12.4" ����Ҳ�����κδ���
                  '��������������SOP ClassUID,����ж�ǰ׺��1.3.46.670589.11.0.0.12.xxx��
                  blnReceived = False
            End If
        End If
    End If
    
    If blnReceived = False Then
        isAdded = False: Status = 0
        Exit Sub
    End If
    
    DViewer.Images.Add ReceivedImage
    DoEvents
    isAdded = False: Status = 0: blnNewImg = True
    ProcSave

    Me.stbThis.Panels(2).Text = "������յ���ͼ����Ϣ�����ˣ�" & NVL(ReceivedImage.Name) & " ���ʱ�䣭" & Time
    WriteCommLog "DViewer_ImageReceived", "���յ�ͼ��", Me.stbThis.Panels(2).Text
    Exit Sub
ProcError:
    On Error Resume Next
    
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog(1, err.Number, err.Description)
End Sub

Private Sub DViewer_NormalisedReceived(ByVal connection As DicomObjects.DicomConnection)
    Dim command As DicomDataSet, ds As DicomDataSet, a As DicomAttribute
    Dim operation As Integer, rclass As String, ruid As String, aclass As String, auid As String
    Dim dss As DicomDataSets, ds1 As DicomDataSet, ds2 As DicomDataSet, i As Integer
    Dim sessionUID As String
    Dim lngPrintStatus As Long
    
    
    Set command = connection.command
    
    '�ص�ο�DICOM��׼������
    operation = command.Attributes(0, &H100)    'Command Field ���ÿһ����Ϣ����һ���ض�ֵ
    rclass = command.Attributes(0, 3) & ""      'Requested SOP Class UID
    ruid = command.Attributes(0, &H1001) & ""   'Requested SOP Instance UID
    aclass = command.Attributes(0, 2) & ""      'Affected SOP Class UID
    auid = command.Attributes(0, &H1000) & ""   'Affected SOP Instance UID
    
    On Error Resume Next
    Select Case operation
    Case &H110  'N-GET      ��Ӧһ��N-GET���󣬷�������������ݼ�
        Set ds = funGetDataset(ruid)
        connection.SendData ds, 0
        connection.SendStatus 0
    
    Case &H140  'N-CREATE   ��Ӧһ��N-CREATE����Affected SOP Class UID����Ҫ������SOP Instance ��Class UID
                'Affected Instance UID ����Ҫ������SOP Instance ��Instance UID
                'Connection.Request.Attributes��N-CREATE�������渽�������ݼ�
                '���Է�Open ��ʱ���Ƚ��յ�FilmSession�����Ĳ���
                '���Է�PrintImage��ʱ�򣬽��յ�FilmBox�����Ĳ���
                
        Set ds = NewDataSet         'ds���Ǳ���N-CREAT�����������ݼ�
        '����Ĭ��ֵ��ʹ�������е����ݼ�����һ��
        For Each a In connection.request.Attributes
            ds.Attributes.Add a.group, a.element, a.value
        Next
    
        ds.Attributes.Add 8, &H16, aclass       'SOP Class UID ,���������Class UID
        If auid = "" Then auid = DGlobal.NewUID
        ds.Attributes.Add 8, &H18, auid         'SOP Instance UID,���������Instance UID
        
        '����Film Box������doSOP_BasicFilmBox���͵�N-CREATE���󣬴���FilmBox
        If aclass = doSOP_BasicFilmBox Then
            
            '���Image boxs  ��������Ȼ�󴴽�����
            '��Image Display Format�н���ͼ��Ĳ���,�����ͼ������
            Dim intImgNum As Integer
            DecodeFormat ds.Attributes(&H2010, &H10), intImgNum
            
            
            '����ͼ�������������������Ҷȵ�ImageBox
            'dssΪds�и��ӵ����ݼ�������������Tag�У���ǰ�洴��ds��ʱ���Ѿ�����TAG��ֵ����DicomDataSets
            Set dss = ds.Tag
            For i = 1 To intImgNum
                '����ImageBox�����ݼ���ͨ��NewDataSet���浽PrintRouterDss��������
                Set ds1 = NewDataSet
                ds1.instanceUID = DGlobal.NewUID                            'instanceUID�Ժ���Ϊ��������
                ds1.Attributes.Add 8, &H1155, ds1.instanceUID               'Referenced SOP Instance UID
                ds1.Attributes.Add 8, &H1150, doSOP_BasicGrayscaleImageBox  'Referenced SOP Class UID
                dss.Add ds1
            Next i
            '��ImageBox��������ӵ�ds��
            ds.Attributes.Add &H2010, &H510, dss        'Referenced Image Box Sequence ���Image Box����
            
            '������session
            Dim SessionSeq As DicomDataSets
            Set SessionSeq = ds.Attributes(&H2010, &H500).value     'Referenced Film Session Sequence,ָ��Session ����
            sessionUID = SessionSeq(1).Attributes(8, &H1155)        'Referenced SOP Instance UID
            
            PrintRouterDss(sessionUID).Tag.Add ds                   '���TAG����ָ�����һ��DicomDataSets
        End If
        
        '����session ,��connecion������
        If aclass = doSOP_BasicFilmSession Then             '����doSOP_BasicFilmSession���͵�N-CREATE����
            connection.Tag.Add ds
        End If
        
        connection.SendData ds, 0
    
    Case &H130  '��ӦN-ACTION���󣬿�ʼ��ӡInstance UID��Ӧ������
        
        'ִ�д�ӡ����-------------------------------------
        '�жϵ�ǰAE�Ǵ�ӡ·�ɣ����ǽ�Ƭ����
        
        lngPrintStatus = funPrintOut(ruid, connection)
        connection.SendStatus lngPrintStatus
    
    Case &H150  '��ӦN-DELETE����ɾ�������Instance UID
        subRemove ruid
        connection.SendStatus 0
    Case &H120  '��Ӧһ��N-SET�������ö�Ӧ��SOP Class��Instance UIDָ�������ݼ���ʵ�����ǽ���ͼ��
        
        Set ds = funGetDataset(ruid)
        For Each a In connection.request.Attributes
            ds.Attributes.Add a.group, a.element, a.value
        Next
        connection.SendStatus 0
    End Select
End Sub

Private Sub DViewer_VerifyReceived(Status As Long)
    On Error GoTo ProcError
    Status = 0
    Exit Sub
ProcError:
    On Error Resume Next
'    Status = err.Number
    
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog(1, err.Number, err.Description)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub fraUD_s_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    fraUD_s.BackColor = IIf(Y > 0, vbWhite, RGB(0, 0, 0))
    On Error Resume Next
    If fraUD_s.Top + Y < 2000 Then
        fraUD_s.Top = 2000
    ElseIf Me.ScaleHeight - fraUD_s.Top - Y < 4000 Then
        fraUD_s.Top = Me.ScaleHeight - 4000
    Else
        fraUD_s.Top = fraUD_s.Top + Y
    End If
End Sub

Private Sub fraUD_s_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub

    fraUD_s.BackColor = Me.BackColor
    Form_Resize
End Sub


Private Sub lvwSeq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwSeq, ColumnHeader.Index)
End Sub

Private Sub mmuCommLog_Click()
    mmuCommLog.Checked = Not mmuCommLog.Checked
End Sub



Private Sub mmuProcessLog_Click()
    mmuProcessLog.Checked = Not mmuProcessLog.Checked
    gblnProcessLog = mmuProcessLog.Checked
End Sub

Private Sub mmuShowLog_Click(Index As Integer)
    Set mfrmShowLog = New frmShowLog
    mfrmShowLog.intLogType = Index
    mfrmShowLog.Show 1, Me
    Set mfrmShowLog = Nothing
End Sub

Private Sub mmuShowService_Click()
        
End Sub

Private Sub mmuUpdateDB_Click()
    Set mfrmUpdateDB = New frmUpdateDB
    Set mfrmUpdateDB.m_cnAccess = gcnAccess
    mfrmUpdateDB.Show 1, Me
    Set mfrmUpdateDB = Nothing
End Sub

Private Sub subListenPorts(iType As Integer)
'-----------------------------------------------------------------------------
'����:������ֹͣ�Է���˿ڵ�����
'����: iType = 0 ֹͣ������iType = 1 ����������
'�޸���:�ƽ�
'�޸�����:2007-11-30
'-----------------------------------------------------------------------------
    Dim strPort As String
    Dim i As Integer
    
    '������������˿�
    For i = 1 To UBound(Services)
        If InStr(strPort, Services(i).ServicePort) = 0 Then
            strPort = strPort & "," & Services(i).ServicePort
            If iType = 0 Then   'ֹͣ����
                DViewer.Unlisten Val(Services(i).ServicePort)
                Services(i).Started = False
            ElseIf iType = 1 Then   '��������
                If Not DViewer.Listen(Val(Services(i).ServicePort)) Then
                    Services(i).Started = False
                    MsgBox "�˿ڣ�" & Services(i).ServicePort & "�ѱ�ʹ�ã�" & _
                    "ϵͳ�޷����������������ü����˿ڡ�", vbExclamation, gstrSysName
                Else
                    Services(i).Started = True
                End If
            End If
        Else
            Services(i).Started = IIf(iType = 0, False, True)
        End If
    Next i
End Sub


Private Sub picView_Resize()
    Dim iCols As Integer, iRows As Integer
    
    On Error Resume Next
    With DViewer
        .Left = 0: .Top = 0
        .Width = picView.ScaleWidth: .Height = picView.ScaleHeight
    
        ResizeRegion .Images.Count, .Width, .Height, iRows, iCols
        .MultiColumns = iCols: .MultiRows = iRows
    End With
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "�˳�"
            Me.WindowState = vbMinimized ' mnuFileQuit_Click
        Case "��ӡ"
            mnuFilePrint_Click
        Case "Ԥ��"
            mnuFilePreview_Click
        Case "����"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub mnuFilePrintSet_Click()
'���ܣ���ӡ����
    Call zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
'���ܣ������Excel
    Call OutputList(3)
End Sub

Private Sub mnuFilePreview_Click()
'���ܣ���ӡԤ��
    Call OutputList(2)
End Sub

Private Sub mnuFilePrint_Click()
'���ܣ���ӡ
    Call OutputList(1)
End Sub

Private Sub mnuHelpTitle_Click()
'���ܣ����ð�������
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset

    On Error Resume Next
    
    If WindowState = vbMinimized Then
        LastState = vbNormal
        Me.Hide
    Else
        LastState = WindowState
    End If

    
    '----------��������ͼ��
    Set mobjIcon = New clsTaskIcon
    mobjIcon.frmHwnd = tbrMain.hwnd ' hwnd
    mobjIcon.Icon = Icon.Handle
    mobjIcon.Message = "ZLPACS��������"
    mobjIcon.AddIcon
    '----------��������ͼ��
    
    mstrPrivs = gstrPrivs
    
    With lvwSeq
        With .ColumnHeaders
            .Clear
        
            .Add , , "Ӱ�����", 1000
            .Add , , "����", 800, 1
            .Add , , "����豸", 1500
            .Add , , "����", 1000
            .Add , , "Ӣ����", 1000
            .Add , , "�Ա�", 600
            .Add , , "����", 600, 1
            .Add , , "Ӱ����", 800, 1
            .Add , , "����ʱ��", 2000
            .Add , , "���UID", 2700
            .Add , , "����UID", 3000
        End With
        .ListItems.Add , , "Temp", , 1
        .ListItems.Clear
    End With
    
    '��ʼ�����ز���
    ReDim AEconnections(0) As AEconnection
    
    gstrAccessPath = App.Path & "\ZlPacsLog"
    gstrAccessName = gstrAccessPath & ".mdb"
    '����������Access����־��¼��������
    With gcnAccess
        .ConnectionString = "DBQ=" & gstrAccessName & ";DefaultDir=" & App.Path & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then MsgBox "���ܴ򿪱�����־�ļ���ϵͳ���޷���¼���չ��̣�", vbInformation, gstrSysName
    End With
    
    strBeginDate = Format(Date & " " & Time, "yyyy-MM-dd hh:mm:ss")
    
    '��DGlobal����ֵ
    Set DGlobal = New DicomGlobal
    '��PrintRouterDss���ϸ���ֵ
    Set PrintRouterDss = New DicomDataSets
    
    '������ӡ���ݼ�
    MakePrinterdataset
    
    If Not ReadPara Then Unload Me: Exit Sub
    
    
    '������������˿�
    Call subListenPorts(1)
    
    lngErrCounts = 0
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    
    strWhere = ""
    ListSeq strWhere
    
    Me.WindowState = vbMinimized
    blnNewImg = False
    
    If funCanStartServer = False Then
        Unload Me
    End If
    
    gblnProcessLog = False
    
    '��¼��������
    Call WriteLog(5802, 5802, "�������������ذ汾Ϊ��" & App.Major & "." & App.Minor & "." & App.Revision)
    
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolItem_Click()
    
    mnuViewToolItem.Checked = Not mnuViewToolItem.Checked
    tbrMain.Visible = mnuViewToolItem.Checked
    tbrMain.Enabled = tbrMain.Visible
    mnuViewToolText.Enabled = tbrMain.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer, j As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    
    For i = 1 To tbrMain.Buttons.Count
        tbrMain.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrMain.Buttons(i).Tag, "")
    Next i
    If mnuViewToolText.Checked Then
        tbrMain.TextAlignment = tbrTextAlignBottom
    End If
    tbrMain.Refresh
    Form_Resize
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hwnd
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long, i As Long

    On Error Resume Next
    
    If WindowState = 1 Then
        Me.Hide
        Exit Sub
    End If
    cbrH = IIf(tbrMain.Visible, tbrMain.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With Me.fraUD_s
        If .Top > Me.ScaleHeight Then .Top = cbrH + (Me.ScaleHeight - cbrH) / 2
        .Left = 0: .Width = Me.ScaleWidth
    End With
    
    With picView
        .Left = 0: .Top = fraUD_s.Top + fraUD_s.Height
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - staH - .Top
    End With
    
    With lvwSeq
        .Left = 0
        .Top = cbrH
        .Height = fraUD_s.Top - .Top
        .Width = Me.ScaleWidth
    End With
    
    If WindowState <> vbMinimized Then
        LastState = WindowState
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call WriteLog(5803, 5803, "���عرա�")
    If gcnAccess.State <> adStateClosed Then gcnAccess.Close
    Call SaveWinState(Me, App.ProductName)
    'ֹͣ��������˿�
    Call subListenPorts(0)
    '�������ͼ��
    mobjIcon.DelIcon
    Set mobjIcon = Nothing
End Sub

Private Sub mobjIcon_MouseLeftDBClick()
    '����������ݿ����ʾ��־��ģʽ�����Ѿ����򿪣����˳���������ִ���
    If mfrmUpdateDB Is Nothing And mfrmShowLog Is Nothing Then
        If WindowState <> 1 Then
            WindowState = vbMinimized
            Me.Hide
        Else
            WindowState = vbNormal
            Me.Show
        End If
    End If
End Sub

Private Sub tbrMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mobjIcon.MouseState X
End Sub

Private Sub tbrMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub OutputList(bytStyle As Byte)
'����: ������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrintLvw

    On Error Resume Next
    If lvwSeq.SelectedItem Is Nothing Then Exit Sub
    
    Set objOut.Body.objData = Me.lvwSeq
    objOut.Title.Text = "Ӱ������"
    objOut.UnderAppItems.Add ""
    objOut.UnderAppItems.Add "ʱ�䣺" & strBeginDate & " - " & Format(Date & " " & Time, "yyyy-MM-dd HH:mm:SS")
    If bytStyle = 1 Then
        bytStyle = zlPrintAsk(objOut)
        If bytStyle <> 0 Then zlPrintOrViewLvw objOut, bytStyle
    Else
        zlPrintOrViewLvw objOut, bytStyle
    End If
End Sub

Private Sub ListSeq(ByVal strWhere As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strCurKey As String
    Dim tmpItem As MSComctlLib.ListItem
    Dim i As Integer

    On Error GoTo DBError
    If Not lvwSeq.SelectedItem Is Nothing Then strCurKey = lvwSeq.SelectedItem.Key
    
    If gcnAccess.State = adStateOpen Then
        gstrSQL = "Select Ӱ�����,����,����豸,����,Ӣ����,�Ա�,����," & _
            " Ӱ����,����ʱ��,��Ӧ���,���UID,����UID,ID" & _
            " From Ӱ��������� Where " & _
            IIf(strWhere = "", "����ʱ��>cDate('" & _
            strBeginDate & "')", strWhere) & _
            " Order By ����ʱ�� Desc"
        Set rsTmp = gcnAccess.Execute(gstrSQL)
        
        Me.lvwSeq.ListItems.Clear
        Do While Not rsTmp.EOF
            i = i + 1
            If i > 500 Then Exit Do
            Set tmpItem = lvwSeq.ListItems.Add(, "_" & rsTmp("ID"), NVL(rsTmp("Ӱ�����")))
            With tmpItem
                .SubItems(1) = NVL(rsTmp("����"))
                .SubItems(2) = NVL(rsTmp("����豸"))
                .SubItems(3) = NVL(rsTmp("����"))
                .SubItems(4) = NVL(rsTmp("Ӣ����"))
                .SubItems(5) = NVL(rsTmp("�Ա�"))
                .SubItems(6) = NVL(rsTmp("����"))
                .SubItems(7) = NVL(rsTmp("Ӱ����"))
                .SubItems(8) = NVL(rsTmp("����ʱ��"), Date)
                .SubItems(9) = NVL(rsTmp("���UID"))
                .SubItems(10) = NVL(rsTmp("����UID"))
                
                .SmallIcon = "_" & IIf(NVL(rsTmp("��Ӧ���"), 1), 0, 1)
                
                If .Key = strCurKey Then .Selected = True
            End With
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
DBError:
'    If ErrCenter() = 1 Then Resume
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog(2, err.Number, err.Description)
End Sub

Private Sub ProcSave()
    On Error GoTo ProcError
    If DViewer.Images.Count > 0 Then
        SaveImages DViewer.Images, mBufferDir
    End If
    Exit Sub
ProcError:
    On Error Resume Next
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog(0, err.Number, err.Description)
End Sub

Private Function ReadPara() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objFile As New Scripting.FileSystemObject
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo DBError
    ReadPara = True
    
    gstrSQL = "Select �豸��,�豸�� From Ӱ���豸Ŀ¼ Where ����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(1))
    
    If rsTmp.EOF Then
        MsgBox "δ����Ӱ��洢�豸���뵽Ӱ���豸Ŀ¼�����ã�", vbInformation, gstrSysName
        ReadPara = False: Exit Function
    End If
    
    '���úʹ�����ʱĿ¼
    mBufferDir = App.Path & "\TempImage\"
    If Not objFile.FolderExists(mBufferDir) Then objFile.CreateFolder mBufferDir
    
    Timer1.Interval = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\���շ���", "�洢���", "10")) * 1000

    
    '��ȡ����IP��ַ
    gstrLocalIP = funcGetLocalIP & ",127.0.0.1"
    
    '�����ݿ��ȡ���������
    strSQL = "Select �豸��,Ӱ�����,PACSAE����,PACS�˿�,�豸IP��ַ,�豸AE����,�豸�˿�,������ From Ӱ��dicom����� a ,Ӱ���豸Ŀ¼ b " & _
             " Where a.�豸��=b.�豸�� And (a.PACS��ɫ='SCP' or a.PACS��ɫ='SCU' ) and NVL(b.״̬,0)=1  And instr([1],PACSIP��ַ)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡDICOM�����", gstrLocalIP)
    If rsTmp.RecordCount > 0 Then
        ReDim Services(rsTmp.RecordCount) As Service
        i = 1
    Else
        ReDim Services(0) As Service
    End If
    
    While Not rsTmp.EOF
        Services(i).DeviceAE = NVL(rsTmp!�豸AE����)
        Services(i).DeviceIP = NVL(rsTmp!�豸IP��ַ)
        Services(i).DevicePort = NVL(rsTmp!�豸�˿�)
        Services(i).DeviceName = NVL(rsTmp!�豸��)
        Services(i).ServiceAE = NVL(rsTmp!PACSAE����)
        Services(i).ServicePort = NVL(rsTmp!PACS�˿�)
        Services(i).SOP = NVL(rsTmp!������)
        Services(i).Modality = NVL(rsTmp!Ӱ�����)
        Services(i).Started = False
        i = i + 1
        rsTmp.MoveNext
    Wend
    
    '�����ݿ��ȡ��������
    strSQL = "Select distinct a.�豸IP��ַ,a.PACSAE����,b.��������,b.����ֵ From Ӱ��DICOM����� a,Ӱ��DICOM������� b,Ӱ���豸Ŀ¼ c " & _
             "Where a.����ID = b.����ID and c.�豸��=a.�豸�� And a.PACS��ɫ='SCP' And NVL(c.״̬,0)=1 and  instr([1],a.PACSIP��ַ)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡDICOM�������", gstrLocalIP)
    If rsTmp.RecordCount > 0 Then
        ReDim AEParas(rsTmp.RecordCount) As AEPara
        i = 1
    End If
    While Not rsTmp.EOF
        AEParas(i).AE = rsTmp!PACSAE����
        AEParas(i).IP = rsTmp!�豸IP��ַ
        AEParas(i).ParaName = rsTmp!��������
        AEParas(i).ParaValue = NVL(rsTmp!����ֵ)
        i = i + 1
        rsTmp.MoveNext
    Wend
    
    '�����ݿ��ȡFTP�洢�豸
    strSQL = "Select �豸��,IP��ַ,FTPĿ¼,FTP�û���,FTP���� From Ӱ���豸Ŀ¼ Where ���� =1 And NVL(״̬,0)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡFTP�洢�豸")
    If rsTmp.RecordCount > 0 Then
        ReDim FTPDevices(rsTmp.RecordCount) As FTPDevice
        i = 1
    End If
    While Not rsTmp.EOF
        FTPDevices(i).No = rsTmp!�豸��
        FTPDevices(i).IP = rsTmp!IP��ַ
        FTPDevices(i).FTPDir = NVL(rsTmp!FTPĿ¼)
        FTPDevices(i).User = NVL(rsTmp!FTP�û���)
        FTPDevices(i).Password = NVL(rsTmp!FTP����)
        i = i + 1
        rsTmp.MoveNext
    Wend
    
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function funGetServiceIndex(strServiceAE As String, Optional strDeviceIP As String = "") As Integer
'���ݷ���AE���豸IP���Ҷ�Ӧ�ķ���ID
'������     strServiceAE --- �����AE����
'           strDeviceIP --- �豸��IP��ַ����Ӧ��ӡ·�ɣ�����Ҫ�����豸��IP��ַ����Ϊ�豸��IP��ַû�еǼ�

    Dim i As Integer
    
    funGetServiceIndex = -1
    If strDeviceIP = "" Then
        For i = 1 To UBound(Services)
            If UCase(Services(i).ServiceAE) = UCase(strServiceAE) Then
                funGetServiceIndex = i
                Exit For
            End If
        Next i
    Else
        For i = 1 To UBound(Services)
            If UCase(Services(i).ServiceAE) = UCase(strServiceAE) And Services(i).DeviceIP = strDeviceIP Then
                funGetServiceIndex = i
                Exit For
            End If
        Next i
    End If
End Function

Private Sub DViewer_QueryRequest(ByVal connection As DicomObjects.DicomConnection)
    Dim result As Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim rq As DicomDataSet
    Dim rqs As DicomDataSets
    Dim rq1 As DicomDataSet
    Dim sql As String
    Dim D1 As DicomAttribute
    Dim NullSequence As New DicomDataSets
    Dim root As String
    Dim resultDS As DicomDataSets
    Dim Level As String
    Dim ds As DicomDataSet
    Dim RemoteAET As String
    Dim ResultImages As DicomImages
    Dim iService As Integer             '��ǰ�����Ӧ�ķ���Ա��
    Dim blnAddModality As Boolean       '��¼�Ƿ������Ӱ�����Ĳ���
    '�����������
    Dim intFilterModality As Integer    '�Ƿ�MWL���˷��� 0--��Ӱ�������ˣ�1--��IP��ַ����
    Dim intDayInterval As Integer       '��ѯ��������
    Dim blnUseForceResult As Boolean    '�Ƿ�ʹ��ǿ�ƽ��
    Dim blnCGet As Boolean              '�Ƿ�����C-GET
    Dim intPatientIDMatch As Integer    'QR��ѯʱ��PatientID��ƥ�䷽ʽ0--���ţ�1--סԺ��/����ţ�2--ҽ��ID
    Dim intBodypartType As Integer      '�ಿλ��ʽ��0-�ޣ�1-�ָ�����2-���¼��3-������
    Dim strBodypartSplitter As String   '�ಿλ�ָ���
    Dim strMultiBodypartsName As String     '�ಿλ���ƣ��á�,������
    Dim strMultiBodypartsCode As String     '�ಿλ���룬�á�,������
    Dim intResultFilter As Integer          '��ѯ����������0-ͼ��ɼ���1-������
    Dim i As Integer
    Dim dtCurDate As Date                   '���ⲿ�����ȡ���ݿ�ʱ�䣬����������Ķ༶ѭ�����ظ���ѯ���ݿ�
    
    On Error GoTo ProcError
    iService = -1
    '�����ݶϿ�ʱ��������
    CheckDBConnect
    '��ȡ�����д����������ݼ�
    Set rq = connection.request
    root = connection.root    '��ȡ��������ĸ��������֣�PATIENT;STUDY;PATIENT/STUDY;WORKLIST
    If root = "WORKLIST" Then    '����Worklist������

        '��¼������־
        Call WriteProcessLog("DViewer_QueryRequest", "���յ�Worklist����", "�����IP��ַ�ǣ� " & connection.RemoteIP & "�������е�AE�ǣ�" & connection.CalledAET)
        
        '��ȡ����Ļ�����������
        funGetAEMWLParas connection.CalledAET, connection.RemoteIP, intFilterModality, intDayInterval, blnUseForceResult, _
                    intBodypartType, strBodypartSplitter, intResultFilter
        
        '��¼������־
        Call WriteProcessLog("DViewer_QueryRequest", "��ȡWorklist����������", "intFilterModality = " & intFilterModality & ",intDayInterval= " & intDayInterval _
                        & ", blnUseForceResult = " & blnUseForceResult & ", intBodypartType= " & intBodypartType & ", strBodypartSplitter=" & strBodypartSplitter _
                        & ",intResultFilter = " & intResultFilter)
                        
        subLogDataset rq, "QueryRequest", "WORKLIST�������ݼ�"
        
        sql = "select /*+RULE*/ e.Ӱ�����,a.ҽ��ID,c.ҽ������,a.���ͺ�,a.Ӣ����,b.ִ�м�,b.������,a.��������,a.�Ա�,b.�״�ʱ��,b.ִ�й���,a.���UID,a.���� " & _
              ",Decode(C.������Դ,2,D.סԺ��,D.�����) As ��ʶ��,a.����,a.����,a.����豸,a.����,a.�������� " & _
              "from Ӱ�����¼ a, ����ҽ������ b,����ҽ����¼ C,������Ϣ D, Ӱ���豸Ŀ¼ E where a.ҽ��ID = b.ҽ��ID and a.���ͺ� = b.���ͺ� And A.����豸 =E.�豸�� " & _
              "And B.ҽ��ID=C.ID And C.����ID=D.����ID And A.�Ƿ���=1 And C.������� in('D','E') And B.ִ��״̬=3 AND C.���ID IS NULL "
        '���ݲ�ѯ������������SQL��ѯ����
        If intResultFilter = 1 Then
            sql = sql & "And B.ִ�й���>=2 And B.ִ�й���<6 "
        Else
            sql = sql & "And B.ִ�й���=2 And A.���UID IS Null "
        End If
              
        '���ݲ����������ͣ��������������� Ӱ���౻����IP��ַ
        If intFilterModality = 0 Then
            '������������ Ӱ�����
            If rq.Attributes(&H40, &H100).Exists And Not IsNull(rq.Attributes(&H40, &H100).value) Then
                Set rqs = rq.Attributes(&H40, &H100).value   'һ��Ƕ�׵����ݼ�
                Set rq1 = rqs(1)
                If rq1.Attributes(&H8, &H60).Exists And Not IsNull(rq1.Attributes(&H8, &H60).value) Then
                    If rq1.Attributes(&H8, &H60) <> "*" Then
                        sql = sql & " And UPPER(e.Ӱ�����)='" & UCase(rq1.Attributes(&H8, &H60).value) & "'"
                        blnAddModality = True
                    End If
                End If
            End If
            If blnAddModality = False Then
                If iService = -1 Then iService = funGetServiceIndex(connection.CalledAET, connection.RemoteIP)
                sql = sql & " And UPPER(e.Ӱ�����)='" & UCase(NVL(Services(iService).Modality)) & "'"
            End If
        ElseIf intFilterModality = 1 Then   '����IP��ַ����
            sql = sql & " And E.IP��ַ='" & connection.RemoteIP & "'"
        End If
        
        If rq.Attributes(&H10, &H20).Exists And Not IsNull(rq.Attributes(&H10, &H20).value) Then
            If Trim(rq.Attributes(&H10, &H20).value) <> "*" Then
                sql = sql & " AND (1=2 "
                AddCondition sql, rq.Attributes(&H10, &H20), "a.����", False
                AddIDCondition sql, rq.Attributes(&H10, &H20), "D.�����", "", False
                AddIDCondition sql, rq.Attributes(&H10, &H20), "D.סԺ��", "", False
                sql = sql & ")"
            End If
        End If
        
        sql = sql & " AND B.�״�ʱ��>=SysDate-" & intDayInterval
        
        AddCondition sql, rq.Attributes(&H10, &H10), "Upper(a.Ӣ����)"
        '����
        If rq.Attributes(&H8, &H50).Exists And Not IsNull(rq.Attributes(&H8, &H50).value) Then
            AddCondition sql, rq.Attributes(&H8, &H50), "a.����", True
        End If
        
'        AddLinkedDateTimeCondition sql, rq1.Attributes(&H40, 2), rq1.Attributes(&H40, 3), "b.�״�ʱ��"
        sql = sql & " order by a.ҽ��ID"
        WriteCommLog "QueryRequest", "���յ�WORKLIST����", Replace(sql, "'", "��")
        
        '��Ϊ���ú�֮��ÿ��WORKLIST�����󣬲�������һ���ģ���˾Ͳ���Ҫ�󶨲���
        Set result = zlDatabase.OpenSQLRecord(sql, "��ѯWORKLIST����")
        
        '��ȡ���ݿ�ʱ��
        dtCurDate = zlDatabase.Currentdate
        
        '����DICOM������Ϣ
        '���ǿ����Ϣ
        If blnUseForceResult Then subReturnDataSet connection, 2, dtCurDate
        
        If strBodypartSplitter = "" Then strBodypartSplitter = "|"
        
        While Not result.EOF
            strMultiBodypartsName = ""
            strMultiBodypartsCode = ""
            '���ѡ����벿λ�������Ƕ��¼���Ƿָ�������ѯ��ҽ��ID��Ӧ�Ķ��벿λ���ƺͶ��벿λ����
            If intBodypartType = 1 Or intBodypartType = 2 Or intBodypartType = 3 Then
                funGetBodypartValue result!ҽ��ID, connection, strBodypartSplitter, strMultiBodypartsName, strMultiBodypartsCode
            End If
            
            If intBodypartType = 2 And strMultiBodypartsName <> "" Then '���¼��ʽ���ض�����벿λ����ѭ�����ض����λ
                For i = 0 To UBound(Split(strMultiBodypartsName, strBodypartSplitter))
                    subReturnDataSet connection, 1, dtCurDate, result, Split(strMultiBodypartsName, strBodypartSplitter)(i), _
                            Split(strMultiBodypartsCode, strBodypartSplitter)(i)
                Next i
            ElseIf intBodypartType = 3 And strMultiBodypartsName <> "" Then '�����з�ʽ���ض�����벿λ
                subReturnDataSet connection, 1, dtCurDate, result, strMultiBodypartsName, strMultiBodypartsCode, strBodypartSplitter
            Else
                subReturnDataSet connection, 1, dtCurDate, result, strMultiBodypartsName, strMultiBodypartsCode
            End If
            result.MoveNext
        Wend
        connection.SendStatus 0
        Exit Sub
    ElseIf root = "PATIENT" Or root = "STUDY" Or root = "PATIENT/STUDY" Then '�����ѯ����Query/Retrieve����
        '��¼����Query/Retrieve����
        subLogDataset rq, "QueryRequest", "���յ�Query/Retrieve����"
        
        '��ȡQR�ķ������
        funGetQRParas connection.CalledAET, connection.RemoteIP, blnCGet, intPatientIDMatch
        
        '��ȡ��������Ĳ��,�����֣�PATIENT;STUDY;SERIES;IMAGE
        Level = rq.Attributes(&H8, &H52)
        '����C-FIND���͵��������󣬲�ѯ���ݿ⣬ֻ���ز�ѯ�����������ͼ��
        If connection.operation = "C-FIND" Then
            Set resultDS = New DicomDataSets
            
            If Level = "PATIENT" Then   '�����˼���Ĳ�ѯ,֧�ֲ�������������ID��Ϊ��ѯ����
                '���ݲ���ID��ƥ�䷽ʽ��֯SQL��ѯ���
                If intPatientIDMatch = 0 Then       '����
                    sql = "Select a.ҽ��ID,a.���� as PatientID,a.����,a.�Ա�,a.Ӣ����,a.�������� From Ӱ�����¼ a Where ���uid Is Not Null "
                ElseIf intPatientIDMatch = 1 Then   'סԺ��/�����
                    sql = "Select a.ҽ��ID,Decode(b.������Դ, 2, c.סԺ��, c.�����) as PatientID,a.����,a.�Ա�,a.Ӣ����,a.�������� " _
                          & " From Ӱ�����¼ a,����ҽ����¼ b,������Ϣ c Where a.���uid Is Not Null and a.ҽ��ID=b.Id And b.����ID=c.����ID "
                Else                                '����ҽ��ID
                    sql = "Select a.ҽ��ID,a.ҽ��ID as PatientID,a.����,a.�Ա�,a.Ӣ����,a.�������� From Ӱ�����¼ a Where ���uid Is Not Null "
                End If
               
                AddStringCondition sql, rq.Name, "a.Ӣ����"
                AddDateCondition sql, rq.Attributes(&H10, &H30), "a.��������"
                
                '���PatientID����,��Ϊ* ���ո�����PatientID��ѯ
                If rq.PatientID <> "*" And rq.PatientID <> "" Then
                    If intPatientIDMatch = 0 Then   '����,ȥ�������*��
                        sql = sql & " and a.����= '" & Replace(rq.PatientID, "*", "") & "'"
                    ElseIf intPatientIDMatch = 1 Then   'סԺ�ţ������
                        sql = sql & " and ((c.סԺ��=" & Val(rq.PatientID) & " And b.������Դ=2) Or (c.�����=" _
                            & Val(rq.PatientID) & " And b.������Դ<>2))"
                    Else    'ҽ��ID
                        sql = sql & " and a.ҽ��ID = " & Val(rq.PatientID)
                    End If
                End If
                
                WriteCommLog "QueryRequest", "���˼����Query", Replace(sql, "'", "��")
                Set result = zlDatabase.OpenSQLRecord(sql, "���˼����Query")
                
                '���ز�ѯ������ݼ���Ӣ������PatientID���������ڣ��Ա𣬱�������AE
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H10, &H10, NVL(result!Ӣ����)
                    AddResultItem ds, rq, &H10, &H20, result!PatientID
                    AddResultItem ds, rq, &H10, &H30, NVL(result!��������)
                    AddResultItem ds, rq, &H10, &H40, IIf(NVL(result!�Ա�) = "Ů", "F", IIf(NVL(result!�Ա�) = "��", "M", "O"))
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            '�����鼶��Ĳ�ѯ
            If Level = "STUDY" Then
                '���ݲ���ID��ƥ�䷽ʽ��֯SQL��ѯ���,֧��Ӣ�������������ڣ����UID��������ڣ��������ڣ���PatientID ��Ϊ��ѯ����
                If intPatientIDMatch = 0 Then       '����
                    sql = "Select  a.ҽ��ID,a.���� as PatientID,a.Ӱ�����,a.���uid,a.�Ա�,a.Ӣ����,a.��������,a.�������� " _
                        & " From Ӱ�����¼ a Where ���uid Is Not Null "
                ElseIf intPatientIDMatch = 1 Then   'סԺ�ţ������
                    sql = "Select a.ҽ��ID,Decode(b.������Դ, 2, c.סԺ��, c.�����) as PatientID,a.Ӱ�����,a.���uid,a.�Ա�,a.Ӣ����,a.��������,a.�������� " _
                          & " From Ӱ�����¼ a,����ҽ����¼ b,������Ϣ c Where a.���uid Is Not Null and a.ҽ��ID=b.Id And b.����ID=c.����ID "
                Else                                'ҽ��ID
                    sql = "Select a.ҽ��ID,a.ҽ��ID as PatientID,a.Ӱ�����,a.���uid,a.�Ա�,a.Ӣ����,a.��������,a.�������� From Ӱ�����¼ a Where ���uid Is Not Null "
                End If
                
                If root = "STUDY" Then
                    '�����Լ��Ϊ���Ĳ�ѯ���ڼ�����Ϸ��ص��ǲ������ֺͼ������,�������ʹ�������ͳ�������������
                    AddStringCondition sql, rq.Name, "a.Ӣ����"
                    AddDateCondition sql, rq.Attributes(&H10, &H30), "a.��������"
                End If
                AddStringCondition sql, rq.StudyUID, "a.���uid"
                AddDateTimeCondition sql, rq.Attributes(&H8, &H20), rq.Attributes(&H8, &H30), "a.��������"
                
                '���PatientID����,��Ϊ* ���ո�����PatientID��ѯ
                If rq.PatientID <> "*" And rq.PatientID <> "" Then
                    If intPatientIDMatch = 0 Then   '����
                        sql = sql & " and a.����= '" & Replace(rq.PatientID, "*", "") & "'"
                    ElseIf intPatientIDMatch = 1 Then   'סԺ�ţ������
                        sql = sql & " and ((c.סԺ��=" & Val(rq.PatientID) & " And b.������Դ=2) Or (c.�����=" _
                            & Val(rq.PatientID) & " And b.������Դ<>2))"
                    Else    'ҽ��ID
                        sql = sql & " and a.ҽ��ID = " & Val(rq.PatientID)
                    End If
                End If
                
                WriteCommLog "QueryRequest", "��鼶���Query", Replace(sql, "'", "��")
                Set result = zlDatabase.OpenSQLRecord(sql, "��鼶���Query")
                
                '��֯���ص����ݼ����������UID��������ڣ��������ڣ���Ӣ������PatientID���������ڣ���������AE������Ӱ�����
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H20, &HD, NVL(result!���UID, 1)
                    '��������ǿ�ѡ����ǵ����ݿ����� û��������ݣ�����ݲ�֧�ִ���
                    'AddResultItem ds, rq, &H8, &H1030, result!StudyDescription
                    AddResultItem ds, rq, &H8, &H20, Format(NVL(result!��������, "19000101"), "YYYYMMDD")
                    AddResultItem ds, rq, &H8, &H30, Format(NVL(result!��������, "12:01:01"), "hhmmss")
                    If root = "STUDY" Then
                        AddResultItem ds, rq, &H10, &H10, result!Ӣ����
                        AddResultItem ds, rq, &H10, &H20, result!PatientID
                        AddResultItem ds, rq, &H10, &H30, result!��������
                    End If
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    AddResultItem ds, rq, &H8, &H60, result!Ӱ�����
                    AddResultItem ds, rq, &H8, &H61, result!Ӱ�����
                    AddResultItem ds, rq, &H10, &H40, IIf(NVL(result!�Ա�) = "Ů", "F", IIf(NVL(result!�Ա�) = "��", "M", "O"))
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            '�������м���Ĳ�ѯ,֧�ֵ��������������UID������UID
            If Level = "SERIES" Then
                sql = "select /*+RULE*/ b.����uid,b.��������,b.���к�,a.Ӱ����� from Ӱ�����¼ a ,Ӱ�������� b " _
                            & "where  a.���uid = b.���uid"
                AddStringCondition sql, rq.StudyUID, "a.���uid"
                AddStringCondition sql, rq.SeriesUID, "b.����uid"
                
                WriteCommLog "QueryRequest", "���м����Query", Replace(sql, "'", "��")
                Set result = zlDatabase.OpenSQLRecord(sql, "���м����Query")
                
                '��֯���ص����ݼ�������������UID���������������кţ�Ӱ����𣬱�������AE,ͼ������
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H20, &HE, result!����uid
                    AddResultItem ds, rq, &H8, &H103E, result!��������
                    AddResultItem ds, rq, &H20, &H11, result!���к�
                    AddResultItem ds, rq, &H8, &H60, result!Ӱ�����
                    AddCountItem ds, rq, &H20, &H1209, "SeriesUID", result!����uid, "InstanceUID"
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            '����ͼ�񼶱�Ĳ�ѯ��֧�ֵ���������������UID��ͼ��UID
            If Level = "IMAGE" Then
                sql = "  select /*+RULE*/ t.ͼ��uid,t.����uid,t.ͼ��� from Ӱ����ͼ�� t where 1=1"
                AddStringCondition sql, rq.SeriesUID, "t.����uid"
                AddStringCondition sql, rq.instanceUID, "t.ͼ��uid"
                
                WriteCommLog "QueryRequest", "ͼ�񼶱��Query", Replace(sql, "'", "��")
                Set result = zlDatabase.OpenSQLRecord(sql, "ͼ�񼶱��Query")
                
                '��֯���ص����ݼ���������ͼ��UID��ͼ��ţ���������AE
                Do Until result.EOF
                    Set ds = NewResultItem(rq)
                    AddResultItem ds, rq, &H8, &H18, result!ͼ��uid
                    AddResultItem ds, rq, &H20, &H13, result!ͼ���
                    AddResultItem ds, rq, &H8, &H54, connection.CalledAET
                    resultDS.Add ds
                    result.MoveNext
                Loop
            End If
            
            For Each ds In resultDS
                subLogDataset ds, "QueryRequest", "Query/Retrieve��ѯ���"
            Next ds
            
            '���Ͳ�ѯ���
            connection.SendData resultDS, &HFF00
        ElseIf connection.operation = "C-GET" Or connection.operation = "C-MOVE" Then
            '����C-GET��C-MOVE����������ķ�������ͼ�񷵻ظ�SCU
            If connection.operation = "C-MOVE" Then
                '���AE�����Ƿ��ڱ���ɵ�AE���У�������ڣ���ܾ�����ͼ��
                '��ΪC-MOVE�Ǳ���ʹ��һ���µ�����������ͼ��ģ�������ʹ�����е����ӣ������Ҫ����
                '��������AE���ƣ����ҵ���AE��Ӧ��IP��ַ�Ͷ˿ںš�
                RemoteAET = connection.Destination
                sql = "Select decode(PACS��ɫ,'SCP',PACSIP��ַ,�豸IP��ַ) As IP��ַ,decode(PACS��ɫ,'SCP',PACS�˿�,�豸�˿�) As �˿ں� " & _
                      "From Ӱ��DICOM����� Where ������='ͼ�����' And (PACS��ɫ='SCP' And upper(PACSAE����) =[1]) Or (PACS��ɫ='SCU' And �豸AE����=[1])"
                WriteCommLog "QueryRequest", "C-MOVE����ͼ���ƶ�Ŀ������", Replace(sql, "'", "��")
                WriteCommLog "QueryRequest", "C-MOVE����ͼ���ƶ�Ŀ������--����", "[1] = " & UCase(RemoteAET)
                
                Set result = zlDatabase.OpenSQLRecord(sql, "����C-MOVE��Ŀ�ĵ�", UCase(RemoteAET))
                If result.EOF Then
                    WriteCommLog "QueryRequest", "C-MOVE�������Ҵ���", "ͼ���ƶ���Ŀ�ĵز���֪"
                    connection.Errors.Attributes.Add 0, &H902, "ͼ���ƶ���Ŀ�ĵز���֪��"
                    connection.SendStatus (&HA801)
                    Exit Sub
                End If
                '�����������Ŀ�ĵ��Ǳ���ģ���������ã������ܳɹ��ķ���ͼ��
                '��ΪC-MOVE����Ҫ��һ��Ҫʹ��һ���µ�����������ͼ��
                WriteCommLog "QueryRequest", "C-MOVE�ҵ�ͼ���ƶ�Ŀ������", "IP��ַ��" & result!IP��ַ _
                             & "���˿ںţ�" & result!�˿ں� & ",����AE:" & connection.CalledAET & ",Զ��AE:" & RemoteAET
                          
                connection.SetDestination result!IP��ַ, result!�˿ں�, connection.CalledAET, RemoteAET
            ElseIf connection.operation = "C-GET" Then
                '�����ﴦ������C-GET�����,������C-GET,��ܾ���
                '�ж��Ƿ�֧��C-GET
                If blnCGet = False Then Exit Sub
            End If
            Set ResultImages = New DicomImages
            If Level = "PATIENT" Then
            '���ݲ���ID��ƥ�䷽ʽ��֯SQL��ѯ���
                If intPatientIDMatch = 0 Then       '����
                    sql = "select ����ID from Ӱ�����¼ a ,����ҽ����¼ b where a.ҽ��id=b.id and a.����=[1]"
                ElseIf intPatientIDMatch = 1 Then   'סԺ��/�����
                    sql = "select ����id from  ������Ϣ  where �����=[1] or סԺ��=[1]"
                Else                                '����ҽ��ID
                    sql = "select ����id from  ����ҽ����¼  where ID=[1]"
                End If
                Set rsTmp = zlDatabase.OpenSQLRecord(sql, "��ѯ����ID", IIf(intPatientIDMatch = 0, rq.PatientID, Val(rq.PatientID)))
                
                WriteCommLog "QueryRequest", "C-MOVE���Ҳ���ID", "sql = " & sql & " ,[1] = " & Val(rq.PatientID)
                
                If Not rsTmp.EOF Then
                    Set ResultImages = GetAllImageFiles(Level, rsTmp!����id)
                End If
            End If
            
            If Level = "STUDY" Then Set ResultImages = GetAllImageFiles(Level, rq.StudyUID)
            If Level = "SERIES" Then Set ResultImages = GetAllImageFiles(Level, rq.SeriesUID)
            If Level = "IMAGE" Then Set ResultImages = GetAllImageFiles(Level, rq.instanceUID)
            If ResultImages Is Nothing Then
                WriteCommLog "QueryRequest", "����ͼ��", "δ�����κ�ͼ��"
                connection.SendStatus 0
                WriteCommLog "QueryRequest", "QRͨѶ���", "δ�����κ�ͼ��"
            Else
                WriteCommLog "QueryRequest", "׼������ͼ��", "ͼ������Ϊ��" & ResultImages.Count & _
                         "ͼ����UIDΪ��" & IIf(ResultImages.Count = 0, "��", ResultImages(1).StudyUID)
                connection.SendImages ResultImages
                WriteCommLog "QueryRequest", "����ͼ�����", "ͼ������Ϊ��" & ResultImages.Count
            End If
        End If
    End If
    Exit Sub
ProcError:
    Call WriteLog(1, err.Number, err.Description)
End Sub

Private Function GetAllImageFiles(Level As String, SearchValue As String) As DicomImages
'------------------------------------------------
'���ܣ���Q/R��ѯ��ʹ�ã��������ط���������ͼ�񼯺�
'������ Level ������ѯ����
'       SearchValue������ѯ����
'���أ�DicomImages��ѯ����ͼ�񼯺�
'-----------------------------------------------
    Dim strSQL As String, lngSeqUID As String
    Dim strURL As String
    Dim rsTmp As New ADODB.Recordset
    Dim dblInit As Double
    Dim FrameCount As Integer
    Dim iCols As Integer, iRows As Integer
    Dim Item As MSComctlLib.ListItem
    Dim clsUseFTP1 As New clsFtp
    Dim clsUseFTP2 As New clsFtp
    
    
    Dim aSeriesUIDs() As String     '�������ڻ�ȡͼ�������UID����
    Dim i As Integer                'ѭ��������
    Dim OneSeriesUID As String      '���浥������UID
    Dim lngResult As Long           '���淵��ֵ
    Dim AllImages As New DicomImages
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    
    Dim curImage As DicomImage, GetAllImages As New DicomImages
    
    Dim bln1stDev As Boolean
    bln1stDev = True
    
    On Error GoTo DBError
    Screen.MousePointer = vbHourglass
    
    'Ҫ�ֲ��ˣ���飬���У�ͼ�� ���ֲ������ȡ������ͼ��,���ݲ�εĲ�ͬ����ȡ����UID���ϵķ�����ͬ
    If Level = "PATIENT" Then
        strSQL = "select /*+RULE*/ e.����uid from Ӱ�����¼ c , Ӱ�������� e , " _
                    & "(select a.����id,b.ҽ��id,b.���ͺ� from ����ҽ����¼ a,����ҽ������ b " _
                    & "where a.����id=" & SearchValue & "��AND a.���ID IS NULL  and a.id=b.ҽ��id) d " _
                    & "Where c.ҽ��id = d.ҽ��id And c.���ͺ� = d.���ͺ� and c.���uid = e.���uid"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    ElseIf Level = "STUDY" Then
        strSQL = "select /*+RULE*/ b.����uid from Ӱ�����¼ a, Ӱ�������� b where a.���uid = b.���uid " _
                    & "and a.���uid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    ElseIf Level = "SERIES" Then
        strSQL = "select /*+RULE*/ t.����uid from Ӱ�������� t where t.����uid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    ElseIf Level = "IMAGE" Then
        strSQL = "select /*+RULE*/ t.����uid from Ӱ�������� t ,Ӱ����ͼ�� q where t.����uid = q.����uid " _
                    & "and q.ͼ��uid = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, SearchValue)
    End If
  
    WriteCommLog "GetAllImageFiles", "����ID����ͼ��", "sql = " & strSQL & " ,[1] = " & SearchValue
  
    If rsTmp.RecordCount <= 0 Then Exit Function    'û�н���򷵻�
    '�����ѯ���������ѯ����������UID����aSeriesUIDs������
    ReDim aSeriesUIDs(rsTmp.RecordCount) As String
    i = 1
    While Not rsTmp.EOF
        aSeriesUIDs(i) = rsTmp!����uid
        i = i + 1
        rsTmp.MoveNext
    Wend
    For i = 1 To UBound(aSeriesUIDs)
        OneSeriesUID = aSeriesUIDs(i)
        strSQL = "Select /*+RULE*/ A.ͼ���,D.FTP�û��� as User1 ,D.FTP���� as Psw1 , D.IP��ַ as IP1 , " & _
            "'/'||D.FtpĿ¼||'/' As FtpPath1,D.�豸�� as �豸��1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
            "||C.���UID||'/' As Path,A.ͼ��UID as ImgName , " & _
            "E.FTP�û��� as User2, E.FTP���� as Psw2 , E.IP��ַ as IP2 , " & _
            "'/'||E.FtpĿ¼||'/' As FtpPath2,E.�豸�� as �豸��2 " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
            "And A.����UID= [1] Order By A.ͼ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, OneSeriesUID)
        If rsTmp.RecordCount > 0 Then
            If strDeviceNO1 <> rsTmp("�豸��1") Then
                clsUseFTP1.FuncFtpDisConnect
                clsUseFTP1.FuncFtpConnect rsTmp("IP1"), rsTmp("User1"), rsTmp("Psw1")
                strDeviceNO1 = rsTmp("�豸��1")
            End If
            If strDeviceNO2 <> rsTmp("�豸��2") Then
                clsUseFTP2.FuncFtpDisConnect
                clsUseFTP2.FuncFtpConnect rsTmp("IP2"), rsTmp("User2"), rsTmp("Psw2")
                strDeviceNO2 = rsTmp("�豸��2")
            End If
            
            Do While Not rsTmp.EOF
                '�жϼ����ļ����Ƿ�IMAGE������ǣ�����ͼ���UID
                If Level <> "IMAGE" Or (Level = "IMAGE" And SearchValue = rsTmp("ImgName")) Then
                    If Dir(mBufferDir & rsTmp("ImgName")) = vbNullString Then
                        'ʹ��FTP�ӷ�������ȡͼ��
                        lngResult = clsUseFTP1.FuncDownloadFile(rsTmp("FtpPath1") & rsTmp("Path"), mBufferDir & rsTmp("ImgName"), rsTmp("ImgName"))
                        If lngResult <> 0 Then  '���豸1��û��ͼ��
                            lngResult = clsUseFTP2.FuncDownloadFile(rsTmp("FtpPath2") & rsTmp("Path"), mBufferDir & rsTmp("ImgName"), rsTmp("ImgName"))
                        End If
                    End If
                    AllImages.ReadFile mBufferDir & rsTmp("ImgName")
                End If
                rsTmp.MoveNext
            Loop
        End If
    Next i
    clsUseFTP1.FuncFtpDisConnect
    clsUseFTP2.FuncFtpDisConnect
    Screen.MousePointer = vbDefault
    Set GetAllImageFiles = AllImages
    Exit Function

DBError:
    clsUseFTP1.FuncFtpDisConnect
    clsUseFTP2.FuncFtpDisConnect
    Screen.MousePointer = vbDefault
    
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
    Call WriteLog(2, err.Number, err.Description)
End Function

Private Sub WriteCommLog(logSubName As String, logTitle As String, logDesc As String)
'���ܣ� ��¼ͨѶ��־
'������ logSubName -- ͨѶ���ڵĹ�������
'       logTitle  --  ͨѶ����
'       logDesc --    ��־����

    Dim strSQL As String
    
    On Error Resume Next
    
    If mmuCommLog.Checked Then
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        strSQL = "Insert into DICOMͨѶ��־ (ͨѶʱ��,ͨѶ����,��¼����,��¼����) " & _
            "Values( cDate('" & Date & " " & Time() & "'),'" & logSubName & "','" & logTitle & _
            "','" & logDesc & "')"
        gcnAccess.Execute strSQL
    End If
End Sub

Private Sub subLogDataset(ds As DicomDataSet, logSubName As String, logTitle As String)
    Dim strLog As String
    If mmuCommLog.Checked Then
        AppendAttributes strLog, "", ds.Attributes
        WriteCommLog logSubName, logTitle, Replace(strLog, "'", "��")
    End If
End Sub

Private Sub AppendAttributes(ByRef list As String, prefix As String, ByRef ob As Object)
    Dim at As DicomAttribute
    Dim s As DicomDataSets
    Dim i As Integer
    Dim v As Variant
    For Each at In ob
        list = list & prefix & "(" & hex4(at.group) & "," & hex4(at.element) & ") : "
        list = list & Left(at.Description & Space(30), 30) & ": "
        If (at.group = &H7FE0) Then ' pixel data
            list = list & "Pixel data" & vbCrLf
        ElseIf (VarType(at.value) = 9) Then ' i.e. a sequence
            Set s = at.value
            list = list & "Sequence of " & s.Count & " items:" & vbCrLf
            For i = 1 To s.Count
                list = list & prefix & ">---------------" & vbCrLf
                AppendAttributes list, prefix & ">", s(i).Attributes
            Next
            list = list & prefix & ">---------------" & vbCrLf
        Else
            v = at.value ' could be variant or array
            If (VarType(v) > 8192) Then ' i.e. an array
                list = list & "Multiple values :" & vbCrLf & "              "
                If UBound(v, 1) > 32 Then
                    list = list & "Array of " & UBound(v, 1) & " elements"
                Else
                    For i = LBound(v, 1) To UBound(v, 1)
                        list = list & v(i)
                        If i <> UBound(v, 1) Then list = list & " : "
                    Next
                End If
                list = list & vbCrLf
            Else
                list = list & v & vbCrLf
            End If
        End If
    Next
End Sub

Private Function hex4(ByVal v As Integer) As String
    hex4 = Right("000" & Hex(v), 4)
End Function

Private Sub subReturnDataSet(connection As DicomConnection, intType As Integer, dtCurDate As Date, Optional rsOracle As Recordset, _
    Optional ByVal strBodypartName As String = "", Optional ByVal strBodyparCode As String = "", _
    Optional ByVal strBodypartSplitter As String = "")
'��֯������Worklist�����ݼ�
'������ connection ---Worklist���ڵ�Dicom����
'       intType --- ���ͣ�1-����������ѯ�����2-����ǿ�ƽ��
'       dtCurDate --- ��ǰ���ݿ�ʱ��
'       rsOracle --- Ҫ���ص�Oracle���ݼ�
'       strBodypartName --- ���벿λ����
'       strBodyparCode --- ���벿λ����
'       strBodypartSplitter --- ���벿λ�ķָ���������зָ�������˵����ʹ�������������벿λ�������Ҫ��������

    Dim dsSeqItem(5) As New DicomDataSet
    Dim dsSeqItemTemp As DicomDataSet
    Dim dsResult As New DicomDataSet
    Dim dssSeq(5) As New DicomDataSets
    Dim rq As DicomDataSet
    Dim rq1(5) As DicomDataSet
    Dim rqs(5) As DicomDataSets
    Dim D1 As DicomAttribute
    Dim NullSequence As New DicomDataSets
    Dim NullSeqItem As New DicomDataSet
    Dim strValue As String      '�������ݼ����ص�һ�����
    Dim intValueType As Integer     '��������ͣ�0-��ͨ�����1-���벿λ���룻2-���벿λ����
    Dim strField As String
    Dim i As Integer
    Dim strFieldValue As String
    Dim strSQL As String
    Dim UpID() As Integer
    Dim dbResult As Recordset
    Dim intBodypartNum As Integer   '���벿λ�Ĳ�λ����
    Dim arrBodypartName() As String '���벿λ���ƴ�
    Dim arrBodypartCode() As String '���벿λ���봮
    Dim int�ϼ�ID As Integer        '���ݼ����ϼ�ID
    
    On Error GoTo ProcError
    '��ʼ���������ݼ�
    Set rq = connection.request
    
    '��ʼ�������е�����
    Set dsResult = rq
    
    '���ж��Ƿ�ʹ�������������벿λ�����ʹ�ã���������벿λ������
    If strBodypartSplitter <> "" Then
        arrBodypartName = Split(strBodypartName, strBodypartSplitter)
        arrBodypartCode = Split(strBodyparCode, strBodypartSplitter)
        intBodypartNum = UBound(arrBodypartName) + 1
    End If
    
    '��ʼ����������
    If rq.Attributes(&H8, &H1110).Exists And Not IsNull(rq.Attributes(&H8, &H1110).value) Then
        Set rqs(1) = rq.Attributes(&H8, &H1110).value
        If rqs(1).Count > 0 Then
            Set rq1(1) = rqs(1)(1)
            For Each D1 In rq1(1).Attributes
                dsSeqItem(1).Attributes.Add D1.group, D1.element, D1.value
            Next
            dssSeq(1).Add dsSeqItem(1)
            dsResult.Attributes.Add &H8, &H1110, dssSeq(1)
        End If
    Else
        Set rqs(1) = NullSequence
        Set rq1(1) = NullSeqItem
    End If
    If rq.Attributes(&H8, &H1120).Exists And Not IsNull(rq.Attributes(&H8, &H1120).value) Then
        Set rqs(2) = rq.Attributes(&H8, &H1120).value
        If rqs(2).Count > 0 Then
            Set rq1(2) = rqs(2)(1)
            For Each D1 In rq1(2).Attributes
                dsSeqItem(2).Attributes.Add D1.group, D1.element, D1.value
            Next
            dssSeq(2).Add dsSeqItem(2)
            dsResult.Attributes.Add &H8, &H1120, dssSeq(2)
        End If
    Else
        Set rqs(2) = NullSequence
        Set rq1(2) = NullSeqItem
    End If
    
    '���������Ҫ������벿λ����8,100�������ǲ�λ����
    If rq.Attributes(&H32, &H1064).Exists And Not IsNull(rq.Attributes(&H32, &H1064).value) Then
        Set rqs(3) = rq.Attributes(&H32, &H1064).value
        If rqs(3).Count > 0 Then
            Set rq1(3) = rqs(3)(1)
            If intBodypartNum > 1 Then
                For i = 1 To intBodypartNum
                    Set dsSeqItemTemp = New DicomDataSet
                    For Each D1 In rq1(3).Attributes
                        dsSeqItemTemp.Attributes.Add D1.group, D1.element, D1.value
                    Next
                    dssSeq(3).Add dsSeqItemTemp
                    If i = 1 Then
                        Set dsSeqItem(3) = dsSeqItemTemp
                    End If
                Next i
            Else
                For Each D1 In rq1(3).Attributes
                    dsSeqItem(3).Attributes.Add D1.group, D1.element, D1.value
                Next
                dssSeq(3).Add dsSeqItem(3)
            End If
            dsResult.Attributes.Add &H32, &H1064, dssSeq(3)
        End If
    Else
        Set rqs(3) = NullSequence
        Set rq1(3) = NullSeqItem
    End If
    
    If rq.Attributes(&H40, &H100).Exists And Not IsNull(rq.Attributes(&H40, &H100).value) Then
        Set rqs(5) = rq.Attributes(&H40, &H100).value
        If rqs(5).Count > 0 Then
            Set rq1(5) = rqs(5)(1)
            For Each D1 In rq1(5).Attributes
                dsSeqItem(5).Attributes.Add D1.group, D1.element, D1.value
            Next
            dssSeq(5).Add dsSeqItem(5)
            dsResult.Attributes.Add &H40, &H100, dssSeq(5)
        End If
        
        '���������Ҫ������벿λ����8,100�������ǲ�λ����
        If rq1(5).Attributes(&H40, &H8).Exists And Not IsNull(rq1(5).Attributes(&H40, &H8).value) Then
            Set rqs(4) = rq1(5).Attributes(&H40, &H8).value
            If rqs(4).Count > 0 Then
                Set rq1(4) = rqs(4)(1)
                If intBodypartNum > 1 Then
                    For i = 1 To intBodypartNum
                        Set dsSeqItemTemp = New DicomDataSet
                        For Each D1 In rq1(4).Attributes
                            dsSeqItemTemp.Attributes.Add D1.group, D1.element, D1.value
                        Next
                        dssSeq(4).Add dsSeqItemTemp
                        If i = 1 Then
                            Set dsSeqItem(4) = dsSeqItemTemp
                        End If
                    Next i
                Else
                    For Each D1 In rq1(4).Attributes
                        dsSeqItem(4).Attributes.Add D1.group, D1.element, D1.value
                    Next
                    dssSeq(4).Add dsSeqItem(4)
                End If
                dsSeqItem(5).Attributes.Add &H40, &H8, dssSeq(4)
            End If
        Else
            Set rqs(4) = NullSequence
            Set rq1(4) = NullSeqItem
        End If
    
    Else
        Set rqs(5) = NullSequence
        Set rq1(5) = NullSeqItem
    End If
    
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    '���ȴ������ݼ�����
    strSQL = "Select a.Id, a.���, a.Ԫ�غ� From Ӱ��MWL����� a ,Ӱ��DICOM����� b " & _
             " Where a.ֵ���� = 'SQ' and a.����ID =b.����ID and  a.ѡ�� = 1 and upper(b.PACSAE����)=[1] and b.�豸IP��ַ =[2]  Order by id "
    Set dbResult = zlDatabase.OpenSQLRecord(strSQL, "��ȡMWL���ݼ�����", CStr(UCase(connection.CalledAET)), CStr(connection.RemoteIP))
    Dim lngMin As Long
    Dim lngMax As Long
     
    If dbResult.RecordCount = 0 Then
        '��ѯ�������ݼ�Ϊ0���Ǵ���ģ�������Ҫѡ���� ��40,100��������˳������̣���¼������־
        err.Raise 10, , Replace(strSQL, "'", "��") & vbNewLine & "��ѯ�������ݼ�������Ӧ��ѡ�񷵻أ�40,100��������ݼ�"
    End If
     
    lngMin = dbResult!id
    dbResult.MoveLast
    lngMax = dbResult!id
    ReDim UpID(lngMin To lngMax) As Integer
    
    dbResult.MoveFirst
    While Not dbResult.EOF
        If dbResult!��� = "0008" And dbResult!Ԫ�غ� = "1110" Then
            UpID(dbResult!id) = 1
        ElseIf dbResult!��� = "0008" And dbResult!Ԫ�غ� = "1120" Then
            UpID(dbResult!id) = 2
        ElseIf dbResult!��� = "0032" And dbResult!Ԫ�غ� = "1064" Then
            UpID(dbResult!id) = 3
        ElseIf dbResult!��� = "0040" And dbResult!Ԫ�غ� = "0008" Then
            UpID(dbResult!id) = 4
        ElseIf dbResult!��� = "0040" And dbResult!Ԫ�غ� = "0100" Then
            UpID(dbResult!id) = 5
        End If
        dbResult.MoveNext
    Wend
    
    'ѭ�����������д���ֵ
    strSQL = "Select a.���, a.Ԫ�غ�, a.�ϼ�id, a.����ֵ, a.�Ƿ����,a.ֵ����, a.Ԫ������, a.ǿ�ƽ��ֵ From Ӱ��MWL����� a , " & _
             " Ӱ��DICOM����� b Where a.����ID =b.����ID and  a.ѡ�� = 1 and upper(b.PACSAE����)=[1] and b.�豸IP��ַ =[2]"
    Set dbResult = zlDatabase.OpenSQLRecord(strSQL, "��ȡMWL���ݼ�", CStr(UCase(connection.CalledAET)), CStr(connection.RemoteIP))
    
    While Not dbResult.EOF
        intValueType = 0
        If dbResult!ֵ���� <> "SQ" Then  '�������Ͳ���SQ�ģ��Ž��뷵��ֵ
            '�����ݿ��ж�ȡ����ֵ�ַ���
            If intType = 1 Then         '����������ѯ���
                strValue = NVL(dbResult!����ֵ)
                '���뷵���ַ���
                Do While InStr(strValue, "[") <> 0
                    If InStr(strValue, "]") = 0 Or InStr(strValue, "]") < InStr(strValue, "[") Then
                        strValue = ""
                        Exit Do
                    End If
                    strField = Mid(strValue, InStr(strValue, "[") + 1, InStr(strValue, "]") - InStr(strValue, "[") - 1)
                    
                    strFieldValue = ""
                    On Error Resume Next
                    If strField = "CallingAET" Then
                        strFieldValue = connection.CallingAET
                    ElseIf strField = "���벿λ����" Then
                        strFieldValue = strBodypartName
                        intValueType = 2
                    ElseIf strField = "���벿λ����" Then
                        strFieldValue = strBodyparCode
                        intValueType = 1
                    Else
                        strFieldValue = funGetFieldValue(strField, rsOracle, dtCurDate)
                    End If
                    
                    strValue = Replace(strValue, "[" & strField & "]", strFieldValue)
                Loop
                
            ElseIf intType = 2 Then         '����ǿ�ƽ��
                strValue = NVL(dbResult!ǿ�ƽ��ֵ)
            End If
            '��������Ľ��
            If dbResult!�Ƿ���� = True Then
                strValue = strValue & mintWLCount
                mintWLCount = mintWLCount + 1
            End If
            '����Ԫ������Ϊ1��1C�ģ�����������ؿ�ֵ
            If dbResult!Ԫ������ = "1" Or UCase(dbResult!Ԫ������) = "1C" Then
                If strValue = "" Then strValue = "1"
            End If
        End If
        
        '������������͵�����
        If dbResult!ֵ���� <> "SQ" Then
            If IsNull(dbResult!�ϼ�ID) Then      '�ϼ�IDΪ�գ�ֱ����д���ݼ�
                AddResultItem dsResult, rq, Int("&H" & dbResult!���), Int("&H" & dbResult!Ԫ�غ�), strValue
            Else    '���ϼ�ID��˵����Ƕ���������е�����
                '֪���ϼ�ID����Ҫ���ҵ�ʹ���Ǹ����ݼ�
                int�ϼ�ID = UpID(dbResult!�ϼ�ID)
                If intValueType = 1 And strBodypartSplitter <> "" Then    '���벿λ����
                    For i = 1 To intBodypartNum
                        AddResultItem dssSeq(int�ϼ�ID)(i), rq1(int�ϼ�ID), Int("&H" & dbResult!���), Int("&H" & dbResult!Ԫ�غ�), arrBodypartCode(i - 1)
                    Next i
                ElseIf intValueType = 2 And strBodypartSplitter <> "" Then    '���벿λ����
                    For i = 1 To intBodypartNum
                        AddResultItem dssSeq(int�ϼ�ID)(i), rq1(int�ϼ�ID), Int("&H" & dbResult!���), Int("&H" & dbResult!Ԫ�غ�), arrBodypartName(i - 1)
                    Next i
                Else
                    AddResultItem dsSeqItem(int�ϼ�ID), rq1(int�ϼ�ID), Int("&H" & dbResult!���), Int("&H" & dbResult!Ԫ�غ�), strValue
                End If
            End If
        End If
        dbResult.MoveNext
    Wend
    
    connection.SendData dsResult, &HFF00  '����в�ƥ����ֶΣ�����ʹ��&HFF01
    subLogDataset dsResult, "subReturnDataset", IIf(intType = 2, "WORKLISTǿ�Ʒ������ݼ�", "WORKLIST�������ݼ�")
    Exit Sub
ProcError:
    Call WriteLog(10, err.Number, err.Description)
    On Error Resume Next
    lngErrCounts = lngErrCounts + 1
    Me.stbThis.Panels(3).Text = "����" & Format(lngErrCounts, "@@@@@@") & "��"
End Sub

Private Function funGetFieldValue(strField As String, rsDataSet As Recordset, dtCurDate As Date) As String
    Dim lngAge As Long
    Dim strAge As String
        
    Select Case strField
        Case "�״�����"
            funGetFieldValue = Format(NVL(rsDataSet!�״�ʱ��, "30000101"), "YYYY-MM-DD")
        Case "�״�ʱ��"
            funGetFieldValue = Format(NVL(rsDataSet!�״�ʱ��, "000001"), "HH:MM:SS")
        Case "Ӱ�����"
            funGetFieldValue = rsDataSet!Ӱ�����
        Case "ִ�м�"
            funGetFieldValue = NVL(rsDataSet!ִ�м�, "XX")
        Case "ִ�й���"
            funGetFieldValue = NVL(rsDataSet!ִ�й���, "2")
        Case "ҽ��ID"
            funGetFieldValue = rsDataSet!ҽ��ID
        Case "��鲿λ"
            If InStr(NVL(rsDataSet!ҽ������), ":") > 0 Then
                funGetFieldValue = Split(rsDataSet!ҽ������, ":")(1)
            Else
                funGetFieldValue = rsDataSet!ҽ������
            End If
        Case "���ͺ�"
            funGetFieldValue = rsDataSet!���ͺ�
        Case "����"
            funGetFieldValue = rsDataSet!����
        Case "��ʶ��"
            funGetFieldValue = rsDataSet!��ʶ��
        Case "Ӣ����"
            funGetFieldValue = rsDataSet!Ӣ����
        Case "�Ա�"
            funGetFieldValue = IIf(NVL(rsDataSet!�Ա�) = "��", "M", IIf(NVL(rsDataSet!�Ա�) = "Ů", "F", "O"))
        Case "����"
            If NVL(rsDataSet!��������) <> "" Then
                '���ݳ�������ת��λdicom��ʽ������
                
                '�������
                lngAge = DateDiff("yyyy", CDate(rsDataSet!��������), dtCurDate)
                If lngAge >= 3 Then
                    funGetFieldValue = Format(lngAge, "000") & "Y"
                    Exit Function
                End If
                
                '���¼���
                lngAge = DateDiff("m", CDate(rsDataSet!��������), dtCurDate)
                If lngAge >= 3 Then
                    funGetFieldValue = Format(lngAge, "000") & "M"
                    Exit Function
                End If
                
                
                '���ܼ���
                lngAge = DateDiff("w", CDate(rsDataSet!��������), dtCurDate)
                If lngAge >= 4 Then
                    funGetFieldValue = Format(lngAge, "000") & "W"
                    Exit Function
                End If
                
                '�������
                lngAge = DateDiff("d", CDate(rsDataSet!��������), dtCurDate)
                funGetFieldValue = Format(lngAge, "000") & "D"
                
                Exit Function
            Else
                '����¼�������ת��Ϊdicom��ʽ��������ʽ
                strAge = NVL(rsDataSet!����, "0")
                
                lngAge = Val(strAge)
                
                Select Case True
                    Case (InStr(strAge, "��") > 0), (InStr(UCase(strAge), "Y") > 0):
                        funGetFieldValue = Format(lngAge, "000") & "Y"
                    Case (InStr(strAge, "��") > 0), (InStr(UCase(strAge), "M") > 0):
                        funGetFieldValue = Format(lngAge, "000") & "M"
                    Case (InStr(strAge, "��") > 0), (InStr(UCase(strAge), "W") > 0):
                        funGetFieldValue = Format(lngAge, "000") & "W"
                    Case Else
                        funGetFieldValue = Format(lngAge, "000") & "D"
                End Select
                    
            End If
        Case "��������"
            funGetFieldValue = Format(NVL(rsDataSet!��������), "YYYYMMDD")
        Case "������"
            funGetFieldValue = NVL(rsDataSet!����)
        Case "����豸"
            funGetFieldValue = NVL(rsDataSet!����豸)
        Case "����"
            funGetFieldValue = Val(NVL(rsDataSet!����))
        Case "��������"
            funGetFieldValue = NVL(rsDataSet!��������)
    End Select
End Function

Private Sub Timer1_Timer()
    
    On Error GoTo err
    
    '�����ݶϿ�ʱ��������
    CheckDBConnect
    
    '����ʣ�µ�ͼ��FTP��
    Call ProcSave
    
    'ˢ����ʾ�б�
    If blnNewImg Then
        ListSeq strWhere
        blnNewImg = False
    End If
    
    '�жϵ�ǰ�Ƿ���ͼ�����û��ͼ�񣬶���300��(5����)֮��û��Association�������Association����
    If DateDiff("S", mdtLastAssociation, Time) > 300 And DViewer.Images.Count = 0 Then
        ReDim AEconnections(0) As AEconnection
    End If
    
    '�ж���־�ļ��Ƿ񳬹�600M�������򴴽��µ���־�ļ�
    If FileLen(gstrAccessName) > 600000000 Then
        Call subNewLogFile
    End If
    Exit Sub
err:
    Call WriteLog(5801, err.Number, "Timer �������������ǣ�" & err.Description)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub subRemove(instanceUID As String)
    Dim children As DicomDataSets, child As DicomDataSet
    Dim thisDataSet As DicomDataSet
    
    ' Object may already have been removed by delete so trap and ignore errors
    On Error GoTo er1
    Set thisDataSet = PrintRouterDss(instanceUID)
    
    Set children = thisDataSet.Tag
    For Each child In children
        subRemove child.instanceUID
    Next
    
    PrintRouterDss.Remove instanceUID
    
cont:
    Exit Sub
    
er1:
    Resume cont
End Sub

Private Function funGetDataset(uID As String)
    '������������ڲ�ͬ�����ݼ����б��治ͬ�����ݼ���
    ' this function would allow youto keep different classes of dataset in different collections if you wished
    Set funGetDataset = PrintRouterDss(uID)
End Function

Private Function NewDataSet() As DicomDataSet
    Set NewDataSet = PrintRouterDss.AddNew
    Set NewDataSet.Tag = New DicomDataSets
End Function

Private Sub DecodeFormat(ByVal strFormat As String, ImgCount As Integer)
'����ͼ���ʽ�ַ���,�����ͼ�������
'��Ҫ����STANDARD��ROW��COL�ĸ�ʽ
'���ø�ʽ��ʾ����Ϊ����STANDARD\2,2������ROW\1,2,2������COL\1,2,2����

    Dim strNum() As String       '��Ÿ�ʽ��ͼ������������
    Dim strPrintFormat As String
    Dim strImageCount As String
    Dim i As Integer
    
    strPrintFormat = Left(strFormat, InStr(strFormat, "\") - 1)
    strImageCount = Right(strFormat, Len(strFormat) - InStr(strFormat, "\"))
    
    strNum = Split(strImageCount, ",")
    
    On Error Resume Next
    
    If strPrintFormat = "STANDARD" Then
        ImgCount = Val(strNum(0)) * Val(strNum(1))
    ElseIf strPrintFormat = "ROW" Or strPrintFormat = "COL" Then
        For i = 0 To UBound(strNum)
            ImgCount = ImgCount + Val(strNum(i))
        Next i
    Else
        ImgCount = 0
    End If
End Sub


Private Function funSaveFilmImages(ByVal ruid As String, connection As DicomConnection, ByVal iService As Integer)
    '���ս�Ƭ���ѽ�Ƭ�е�ͼ��ȫ����������
    Dim DFilmBox As DicomDataSet
    Dim DImageBoxs As DicomDataSets
    Dim DSessions As DicomDataSets
    Dim DSession As DicomDataSet
    Dim intImageCount As Integer
    Dim strFormat As String
    Dim DImageds As DicomDataSet
    Dim DImageAtt As DicomAttribute
    Dim DImages As DicomImages
    Dim DImage As DicomImage
    Dim DTempImage As New DicomImage
    Dim strStudyUID As String
    Dim strSeriesUID As String
    Dim curDate As Date
    
    Dim i As Integer
    
    On Error GoTo err1
    
    '����ruid�ӹ������ݼ�PrintRouterDss�л��һ��ָ��FilmBox�����ݼ�
    Set DFilmBox = PrintRouterDss(ruid)
    '��ȡImageBox���� Referenced Image Box Sequence
    Set DImageBoxs = DFilmBox.Attributes(&H2010, &H510).value
    '��ȡsession  Referenced Film Session Sequence
    Set DSessions = DFilmBox.Attributes(&H2010, &H500).value
    Set DSession = DSessions(1)
    Set DSession = PrintRouterDss(DSession.Attributes(8, &H1155).value)
    
    '��ȡͼ���ӡ��ʽ
    If DFilmBox.Attributes(&H2010, &H10).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H10)) Then
        strFormat = DFilmBox.Attributes(&H2010, &H10)
    Else
        strFormat = "STANDARD\1,1"
    End If
    '���ݸ�ʽ������ͼ������
    DecodeFormat strFormat, intImageCount
    
    '��ȡ�´���ͼ��� ���UID������UID
    strStudyUID = DTempImage.StudyUID
    strSeriesUID = DTempImage.SeriesUID
    
    '��ǰ��ȡ���ݿ�ʱ�䣬�����������ѭ���ж�β�ѯ���ݿ�
    curDate = zlDatabase.Currentdate
    
    'ѭ��ImageBoxs ��ӡÿһ��ͼ��
    For i = 1 To intImageCount
        '��ȡͼ�����ݼ�
        '(8,&h1155) = Referenced SOP Instance UID
        Set DImageds = PrintRouterDss(DImageBoxs(i).Attributes(8, &H1155).value)
        
        '�ȳ��Զ�ȡ�Ҷ�ͼ�� (&H2020, &H110) = Basic Grayscale Image Sequence
        'ÿһ��DImageAtt���汣��һ��DicomImage��ʵ����DImageAtt�Ǵ�DImageds�����������
        Set DImageAtt = DImageds.Attributes(&H2020, &H110)
        
        '����Ҷ�ͼ��û�ж�ȡ�ɹ������ȡ��ɫͼ��
        '(&H2020, &H111) = Basic Colour Image Sequence
        If Not DImageAtt.Exists Then
            Set DImageAtt = DImageds.Attributes(&H2020, &H111)
        End If
        
        '�ҵ�ͼ����ʼ����ͼ��
        If DImageAtt.Exists Then
            Set DImages = New DicomImages
            DImages.Add DImageAtt.value.Item(1)
            Set DImage = DImages(1)
            '��ͼ��ŵ�DViewer�У��ȴ�����
            subWriteDicomPara DImage, iService, Str(i), curDate
            DImage.StudyUID = strStudyUID
            DImage.SeriesUID = strSeriesUID
            DImage.Tag = connection.Association
'            DImage.Tag = "��Ƭ����-" & iService
            
            '�������Ӳ���
            subSaveAssociation connection
            
            DViewer.Images.Add DImage
        End If
    Next i
    
    blnNewImg = True
    funSaveFilmImages = 0
    Exit Function
err1:
    funSaveFilmImages = 1
End Function

Private Sub subWriteDicomPara(img As DicomImage, ByVal iService As Integer, strImageNum As String, curDate As Date)
'------------------------------------------------
'���ܣ��������ͼ����дDICOM�ļ�ͷ��Ϣ
'������img���������DICOM�ļ�
'       iService--������Ϣ���������
'       strImageNum--ͼ���
'���أ��ޣ�ֱ���ļ�ͷ��Ϣд��img���ļ�ͷ
'------------------------------------------------
    Dim g As New DicomGlobal
    
    img.instanceUID = g.NewUID
    img.Attributes.Add &H8, &H8, ""                             'ImageType  ��
    img.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"   'SOP Class  UID�����β�׽
    img.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date �������
    img.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date ��������
    img.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date �ɼ�����
    img.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   ͼ������
    img.Attributes.Add &H8, &H30, Format(curDate, "HH:MM:SS")     'Study Time   ���ʱ��
    img.Attributes.Add &H8, &H31, Format(curDate, "HH:MM:SS")     'Series Time  ����ʱ��
    img.Attributes.Add &H8, &H32, Format(curDate, "HH:MM:SS")     'Acquisition Time  �ɼ�ʱ��
    img.Attributes.Add &H8, &H33, Format(curDate, "HH:MM:SS")     'Image Time  ͼ��ʱ��
    img.Attributes.Add &H8, &H50, ""                            'Accession Number ��
    img.Attributes.Add &H8, &H60, Services(iService).Modality   'Modality Ӱ�����
    img.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer ����
    img.Attributes.Add &H8, &H80, gstr��λ����                  'Institution Name ��λ����
    img.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name ��
    img.Attributes.Add &H8, &H1030, ""                          'Study Description ������� ��
    img.Attributes.Add &H10, &H10, ""                           'Name ����
    img.Attributes.Add &H10, &H20, ""                           'Patient ID ����ID
    img.Attributes.Add &H10, &H30, ""                           'BirthDate ����
    img.Attributes.Add &H10, &H40, ""                           'Sex �Ա�
    img.Attributes.Add &H10, &H1010, ""                         'Age ����
    img.Attributes.Add &H10, &H4000, ""                         'Patient Comment ����ע��
    img.Attributes.Add &H20, &H10, ""                           'Study ID ���ID
    img.Attributes.Add &H20, &H11, "1"                          'Series Number ���к�
    img.Attributes.Add &H20, &H13, strImageNum                         'ImageNumber ͼ���
    img.Attributes.Add &H20, &H20, ""                           'Orientation ��
End Sub


Private Function funPrintOut(ByVal ruid As String, connection As DicomConnection) As Long
    '�Ѵ�ӡ������ת����ʵ�Ĵ�ӡ��,����ת����ͼ�����
    
    Dim DPrinter As New DicomPrint
    Dim DFilmBox As DicomDataSet
    Dim DImageBoxs As DicomDataSets
    Dim DSessions As DicomDataSets
    Dim DSession As DicomDataSet
    Dim strOrientation As String
    Dim strFilmSize As String
    Dim intCopies As Integer
    Dim intImageCount As Integer
    Dim DImageds As DicomDataSet
    Dim DImageAtt As DicomAttribute
    Dim DImages As DicomImages
    Dim DImage As DicomImage
    Dim iService As Integer
    
    Dim i As Integer
    
    iService = -1
    
    '��ȡ����������IP��ַ���˿ڣ�AE���Ƶ�
    If iService = -1 Then iService = funGetServiceIndex(connection.CalledAET)
    
    If iService = -1 Then
        '�Ҳ�����ӡ·�ɵ����������ã��˳���ӡ
        funPrintOut = 1
        Exit Function
    Else
        '����ǡ���Ƭ���ա���ת���ӳ���ִ��
        If Services(iService).SOP = "��Ƭ����" Then
            funPrintOut = funSaveFilmImages(ruid, connection, iService)
            Exit Function
        End If
        
        DPrinter.Node = Services(iService).DeviceIP
        DPrinter.Port = Services(iService).DevicePort
        DPrinter.CalledAE = Services(iService).DeviceAE
        DPrinter.CallingAE = Services(iService).ServiceAE
    End If
    
    On Error GoTo err1
    
    '����ruid�ӹ������ݼ�PrintRouterDss�л��һ��ָ��FilmBox�����ݼ�
    Set DFilmBox = PrintRouterDss(ruid)
    '��ȡImageBox���� Referenced Image Box Sequence
    Set DImageBoxs = DFilmBox.Attributes(&H2010, &H510).value
    '��ȡsession  Referenced Film Session Sequence
    Set DSessions = DFilmBox.Attributes(&H2010, &H500).value
    Set DSession = DSessions(1)
    Set DSession = PrintRouterDss(DSession.Attributes(8, &H1155).value)
    
    
    
    '��ӡͼ���λ��������
    '�������޷�֪��ͼ���λ����������Ϊ��ӡ·�ɣ����յ���ͼ�����Ѿ�����ÿ���ֱ�Ӵ�ӡ��ͼ���ˣ�
    '��˲��ٴ����ӡͼ���λ���������ڴ�ӡÿһ��ͼ���ʱ��PrintImage��ʹ��Raw=True�Ĳ�����
    'ʹ�� ImageBox �е�����ȡ (0028,0101) : Bits Stored
    '''''''''''''''''''Printer��ֱ�Ӳ���''''''''''''''''''''''''''''''''''
    

    '''''''''''''''''''Session�Ĳ���''''''''''''''''''''''''''''''''''
    '��ӡ����������
    If Not IsNull(DSession.Attributes(&H2000, &H10)) Then
        intCopies = DSession.Attributes(&H2000, &H10)           '��ȡNumber of Copies
        DPrinter.Copies = intCopies
    Else
        DPrinter.Copies = 1
    End If
    
    'Print Priority ���ȼ�����ѡ
    If DSession.Attributes(&H2000, &H20).Exists And Not IsNull(DSession.Attributes(&H2000, &H20)) Then
        DPrinter.Session.Attributes.Add &H2000, &H20, DSession.Attributes(&H2000, &H20)
    End If
    
    'Medium Type �������ͣ���ѡ
    If DSession.Attributes(&H2000, &H30).Exists And Not IsNull(DSession.Attributes(&H2000, &H30)) Then
        DPrinter.Session.Attributes.Add &H2000, &H20, DSession.Attributes(&H2000, &H30)
    End If
    
    
    'Film Destination ����Ŀ�꣬��ѡ
    If DSession.Attributes(&H2000, &H40).Exists And Not IsNull(DSession.Attributes(&H2000, &H40)) Then
        DPrinter.Session.Attributes.Add &H2000, &H20, DSession.Attributes(&H2000, &H40)
    End If
    
    
    '''''''''''''''''''''''''''''''''''''    '�򿪴�ӡ��'''''''''''''''''''''''''''
    DPrinter.Open
    
    '''''''''''''''''''FilmBox�Ĳ���''''''''''''''''''''''''''''''''''
    '��Ƭ���� ������
    If Not IsNull(DFilmBox.Attributes(&H2010, &H40)) Then
        strOrientation = DFilmBox.Attributes(&H2010, &H40)      '��ȡFilm Orientation
        DPrinter.Orientation = strOrientation
    Else
        DPrinter.Orientation = "PORTRAIT"
    End If
    
    '��ӡ��ʽ������
    If DFilmBox.Attributes(&H2010, &H10).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H10)) Then
        DPrinter.Format = DFilmBox.Attributes(&H2010, &H10)
    Else
        DPrinter.Format = "STANDARD\1,1"
    End If
    
    '��Ƭ��С,Ĭ��ֵΪ�գ�ʹ�ô�ӡ����Ĭ��ֵ
    If Not IsNull(DFilmBox.Attributes(&H2010, &H50)) Then
        strFilmSize = DFilmBox.Attributes(&H2010, &H50)
        DPrinter.FilmSize = strFilmSize
    End If
    
    '�Ŵ�ʽ,����
    If DFilmBox.Attributes(&H2010, &H60).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H60)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H60, DFilmBox.Attributes(&H2010, &H60)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H60, "CUBIC"
    End If
    
    'Smoothing Type 'ƽ��,��ѡ
    If DFilmBox.Attributes(&H2010, &H80).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H80)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H80, DFilmBox.Attributes(&H2010, &H80)
    End If
    
    'border density ��Ե�ܶȣ�����
    If DFilmBox.Attributes(&H2010, &H100).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H100)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H100, DFilmBox.Attributes(&H2010, &H100)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H100, "BLACK"
    End If
    
    'empty image density �հ��ܶȣ�����
    If DFilmBox.Attributes(&H2010, &H110).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H110)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H110, DFilmBox.Attributes(&H2010, &H110)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H110, "BLACK"
    End If

    '���з�ʽ
    If DFilmBox.Attributes(&H2010, &H140).Exists And Not IsNull(DFilmBox.Attributes(&H2010, &H140)) Then
        DPrinter.FilmBox.Attributes.Add &H2010, &H140, DFilmBox.Attributes(&H2010, &H140)
    Else
        DPrinter.FilmBox.Attributes.Add &H2010, &H140, "NO"
    End If
        
    'Polarity ����,��ѡ
    If DFilmBox.Attributes(&H2020, &H20).Exists And Not IsNull(DFilmBox.Attributes(&H2020, &H20)) Then
        DPrinter.FilmBox.Attributes.Add &H2020, &H20, DFilmBox.Attributes(&H2020, &H20)
    End If
        
    'Requested Resolution ID �ֱ��ʣ���ѡ
    If DFilmBox.Attributes(&H2020, &H50).Exists And Not IsNull(DFilmBox.Attributes(&H2020, &H50)) Then
        DPrinter.FilmBox.Attributes.Add &H2020, &H50, DFilmBox.Attributes(&H2020, &H50)
    End If
        
    'ѭ��ImageBoxs ��ӡÿһ��ͼ��
    DecodeFormat DPrinter.Format, intImageCount
    For i = 1 To intImageCount
        '��ȡͼ�����ݼ�
        '(8,&h1155) = Referenced SOP Instance UID
        Set DImageds = PrintRouterDss(DImageBoxs(i).Attributes(8, &H1155).value)
        
        '�ȳ��Զ�ȡ�Ҷ�ͼ�� (&H2020, &H110) = Basic Grayscale Image Sequence
        'ÿһ��DImageAtt���汣��һ��DicomImage��ʵ����DImageAtt�Ǵ�DImageds�����������
        Set DImageAtt = DImageds.Attributes(&H2020, &H110)
        
        '����Ҷ�ͼ��û�ж�ȡ�ɹ������ȡ��ɫͼ��
        '(&H2020, &H111) = Basic Colour Image Sequence
        If Not DImageAtt.Exists Then
            Set DImageAtt = DImageds.Attributes(&H2020, &H111)
        End If
        
        '�ҵ�ͼ����ʼ��ӡ
        If DImageAtt.Exists Then
            Set DImages = New DicomImages
            DImages.Add DImageAtt.value.Item(1)
            Set DImage = DImages(1)
            DPrinter.PrintImage DImage, True, False
        End If
    Next i
    DPrinter.PrintFilm
    DPrinter.Close
    
    funPrintOut = 0
    Exit Function
err1:
    funPrintOut = 1
    
End Function

Private Sub MakePrinterdataset()
    Dim p As DicomDataSet
    Set p = NewDataSet
    
    p.instanceUID = doInstance_Printer
    p.Attributes.Add 8, &H16, doSOP_Printer             'SOP Class UID
    p.Attributes.Add 8, &H70, "ZLSOFT"                  'Manufacturer
    p.Attributes.Add 8, &H1090, "Demo Printer SCP"      'Manufacturer's Model Name
    p.Attributes.Add &H18, &H1000, "serial no 1234"     'Device Serial Number
    p.Attributes.Add &H18, &H1020, DGlobal.Version            'Software Version(s)
    Set printerobject = p
End Sub

Private Function funCanStartServer() As Boolean
'����豸�����Ƿ񳬹�����,����������������������
'������
'����ֵ��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    gint��Ƭ��ӡ������ = getLicenseCount(LOGIN_TYPE_��Ƭ��ӡ��)
    gintDICOM�豸���� = getLicenseCount(LOGIN_TYPE_DICOM�豸)
    strSQL = "select �豸��,�豸��,���� from Ӱ���豸Ŀ¼ Where  NVL(״̬,0)=1 and (����=3 Or ����=4 )"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�豸����")
    
    rsTemp.Filter = "����=3"
    If (rsTemp.RecordCount > gint��Ƭ��ӡ������ And gint��Ƭ��ӡ������ <> -1) Or gint��Ƭ��ӡ������ = 0 Then
        MsgBox LOGIN_TYPE_��Ƭ��ӡ�� & "�������������������" & gint��Ƭ��ӡ������ & "���������޷����������������Ӧ����ϵ", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rsTemp.Filter = "����=4"
    If (rsTemp.RecordCount > gintDICOM�豸���� And gintDICOM�豸���� <> -1) Or gintDICOM�豸���� = 0 Then
        MsgBox LOGIN_TYPE_DICOM�豸 & "�������������������" & gintDICOM�豸���� & "���������޷����������������Ӧ����ϵ", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    funCanStartServer = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function funGetBodypartValue(lngOrderID As Long, connection As DicomConnection, strBodypartSplitter As String, strBodypartName As String, _
    strBodypartCode As String) As Boolean
'����ҽ��ID����ѯ���벿λ���ƺͶ��벿λ����
'������ lngOrderID��IN�� --- ҽ��ID
'       connection��IN�� --- DICOM����
'       strBodyPartSplitter��IN�� --- �ಿλ�ķָ���
'       strBodypartName ��OUT��--- ���벿λ���ƴ�����strBodyPartSplitter�ָ�
'       strBodypartCode ��OUT��--- ���벿λ���봮����strBodyPartSplitter�ָ�
'����ֵ��   True-�ɹ���False-ʧ��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    'PACS��λ�ǡ��걾��λ+��鷽������ͬ��ɲ�λ�����еġ�PACS��λ���ơ�
    strSQL = "Select x.��λ���� ,b.�豸��λ����, b.�豸��λ���� " & _
             " From  (Select a.�걾��λ||a.��鷽�� As ��λ���� From ����ҽ����¼ a Where ���id = [1]) x, " & _
             " Ӱ��mwl��λ���� b, Ӱ��dicom����� c " & _
             " Where x.��λ���� =b.Pacs��λ���� And b.����id = c.����id And Upper(c.Pacsae����) = [2] and c.�豸ip��ַ = [3] " & _
             " order by �豸��λ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngOrderID, UCase(CStr(connection.CalledAET)), CStr(connection.RemoteIP))
    
    While rsTemp.EOF = False
        If InStr(strBodypartCode, NVL(rsTemp!�豸��λ����)) = 0 Then
            strBodypartName = strBodypartName & strBodypartSplitter & NVL(rsTemp!�豸��λ����)
            strBodypartCode = strBodypartCode & strBodypartSplitter & NVL(rsTemp!�豸��λ����)
        End If
        rsTemp.MoveNext
    Wend
    
    If strBodypartName <> "" Then
        strBodypartName = Mid(strBodypartName, 2)
    End If
    
    If strBodypartCode <> "" Then
        strBodypartCode = Mid(strBodypartCode, 2)
    End If
    funGetBodypartValue = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    funGetBodypartValue = False
End Function
