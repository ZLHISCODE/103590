VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiSelect 
   Caption         =   "����ѡ��"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15000
   Icon            =   "frmPatiSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15000
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdExit 
      Caption         =   "ȡ��(&E)"
      Height          =   495
      Left            =   12480
      TabIndex        =   26
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   495
      Left            =   10800
      TabIndex        =   25
      Top             =   7680
      Width           =   1575
   End
   Begin VB.PictureBox picת�� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   5880
      ScaleHeight     =   5415
      ScaleWidth      =   9855
      TabIndex        =   4
      Top             =   4200
      Width           =   9855
      Begin XtremeReportControl.ReportControl rptת�� 
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   3360
         _Version        =   589884
         _ExtentX        =   5927
         _ExtentY        =   4260
         _StockProps     =   0
         BorderStyle     =   2
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.TextBox txtChange 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   900
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "7"
         Top             =   120
         Width           =   285
      End
      Begin VB.Frame fraChange 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   870
         TabIndex        =   22
         Top             =   330
         Width           =   300
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "ˢ��"
         Height          =   255
         Left            =   2625
         TabIndex        =   21
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblת�� 
         AutoSize        =   -1  'True
         Caption         =   "��ʾ���    ���ת������"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   150
         Width           =   2160
      End
   End
   Begin VB.PictureBox pic��Ժ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   3600
      ScaleHeight     =   5415
      ScaleWidth      =   9855
      TabIndex        =   2
      Top             =   2640
      Width           =   9855
      Begin XtremeReportControl.ReportControl rpt��Ժ 
         Height          =   2415
         Left            =   0
         TabIndex        =   17
         Top             =   480
         Width           =   3360
         _Version        =   589884
         _ExtentX        =   5927
         _ExtentY        =   4260
         _StockProps     =   0
         BorderStyle     =   2
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cboSelectTime 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label lbl��Ժʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox picHLDJ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1200
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic��Ժ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   840
      ScaleHeight     =   5415
      ScaleWidth      =   9855
      TabIndex        =   1
      Top             =   360
      Width           =   9855
      Begin VB.PictureBox picIn��Ժ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   5895
         TabIndex        =   5
         Top             =   0
         Width           =   5895
         Begin VB.PictureBox picPati 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   999
            Left            =   0
            Picture         =   "frmPatiSelect.frx":6852
            ScaleHeight     =   1455
            ScaleWidth      =   1395
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   1395
            Begin VB.Label lbl���� 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   210
               Index           =   999
               Left            =   1740
               TabIndex        =   14
               Top             =   1620
               Width           =   105
            End
            Begin VB.Label lbl����� 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Height          =   180
               Index           =   999
               Left            =   1800
               TabIndex        =   13
               Top             =   840
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label lbl���� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "09123"
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
               Index           =   999
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   675
            End
            Begin VB.Label lblסԺ�� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "027647132"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   999
               Left            =   120
               TabIndex        =   11
               Top             =   840
               Width           =   810
            End
            Begin VB.Label lbl�Ա� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               ForeColor       =   &H00C00000&
               Height          =   180
               Index           =   999
               Left            =   630
               TabIndex        =   10
               Top             =   1125
               Width           =   180
            End
            Begin VB.Label lbl���� 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "33"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   999
               Left            =   930
               TabIndex        =   9
               Top             =   1125
               Width           =   525
            End
            Begin VB.Label lbl���� 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "���������л����񹲺͹�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Index           =   999
               Left            =   75
               TabIndex        =   8
               Top             =   375
               Width           =   1215
            End
            Begin VB.Label lblSplit 
               BackColor       =   &H008080FF&
               Height          =   60
               Index           =   999
               Left            =   45
               TabIndex        =   7
               Top             =   735
               Width           =   1260
            End
            Begin VB.Image img����ȼ� 
               Appearance      =   0  'Flat
               Height          =   240
               Index           =   999
               Left            =   1065
               Picture         =   "frmPatiSelect.frx":DA34
               Stretch         =   -1  'True
               Top             =   60
               Width           =   240
            End
            Begin VB.Label lblSelect 
               BackColor       =   &H00FFC0C0&
               Height          =   360
               Index           =   999
               Left            =   45
               TabIndex        =   15
               Top             =   375
               Visible         =   0   'False
               Width           =   1260
            End
         End
      End
      Begin VB.VScrollBar HScr 
         Height          =   3945
         LargeChange     =   1000
         Left            =   9600
         Max             =   0
         SmallChange     =   20
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.Timer tmrOpen 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer tmrClose 
      Left            =   480
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5265
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9287
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList imgHLDJ 
      Index           =   999
      Left            =   0
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":DD76
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":E310
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":E8AA
            Key             =   "�ȴ����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":EE44
            Key             =   "�ܾ����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":F3DE
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":F978
            Key             =   "���ڳ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1038A
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":10D9C
            Key             =   "��鷴��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":11336
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":11D48
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1275A
            Key             =   "δ����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":12CF4
            Key             =   "ִ����"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1328E
            Key             =   "������"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":13CA0
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1423A
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":147D4
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":14D6E
            Key             =   "������"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1B5D0
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1BB6A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":1C104
            Key             =   "Fbaby"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
'Const LWA_COLORKEY = &H1
Private lngAlpha As Integer
Private j As Integer
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As PointAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" _
    (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const ALTERNATE = 1
Private mobjFileSys As New FileSystemObject
Private mlng����ID As Long
Private mblnCardOrder As Boolean
Private mColBad As New Collection
Private mblnCancle As Boolean
Private mintIndex As Integer
Private mlngColor As Long
Private mstr����IDs As String
Private mstrסԺ�� As String
Private mstr���� As String
Private mstr���� As String
Private mblnOK As Boolean
Private Enum PATIREPORT_COLUMN
    COL_��� = 0
    COL_����ID = 1
    COL_��ҳID = 2
    COL_���� = 3
    COL_סԺ�� = 4
    COL_���� = 5
    col_�Ա� = 6
    col_���� = 7
    col_�������� = 8
    col_��Ժ���� = 9
    col_��Ժ���� = 10
    col_סԺ���� = 11
End Enum
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintOutPreTime As Integer

Public Function ShowMe(objParent As Object, ByVal lng����ID As Long, str����IDs As String, Optional str���� As String, Optional strסԺ�� As String, Optional str���� As String) As Boolean
    mlng����ID = lng����ID
    mstr����IDs = ""
    mstrסԺ�� = ""
    mstr���� = ""
    mstr���� = ""
    mblnOK = False
    Me.Show 1, objParent
    ShowMe = mblnOK
    If mblnOK Then
        str����IDs = mstr����IDs
        strסԺ�� = mstrסԺ��
        str���� = mstr����
        str���� = mstr����
    End If
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim objSel As ReportRow
    
    For i = 0 To picPati.Count - 2
        If lblSelect(i).Visible Then
            mstr����IDs = mstr����IDs & "," & Val(Split(picPati(i).Tag, ",")(0)) & ":" & Val(Split(picPati(i).Tag, ",")(1))
            mstrסԺ�� = mstrסԺ�� & "," & lblסԺ��(i).Caption
            mstr���� = mstr���� & "," & lbl����(i).Caption
            mstr���� = mstr���� & "," & lbl����(i).Caption
        End If
    Next
    
    For Each objSel In rpt��Ժ.SelectedRows
        mstr����IDs = mstr����IDs & "," & Val(objSel.Record.Item(COL_����ID).value) & ":" & Val(objSel.Record.Item(COL_��ҳID).value)
        mstrסԺ�� = mstrסԺ�� & "," & objSel.Record.Item(COL_סԺ��).value
        mstr���� = mstr���� & "," & Trim(objSel.Record.Item(COL_����).value)
        mstr���� = mstr���� & "," & objSel.Record.Item(COL_����).value
    Next
    
    For Each objSel In rptת��.SelectedRows
        mstr����IDs = mstr����IDs & "," & Val(objSel.Record.Item(COL_����ID).value) & ":" & Val(objSel.Record.Item(COL_��ҳID).value)
        mstrסԺ�� = mstrסԺ�� & "," & objSel.Record.Item(COL_סԺ��).value
        mstr���� = mstr���� & "," & Trim(objSel.Record.Item(COL_����).value)
        mstr���� = mstr���� & "," & objSel.Record.Item(COL_����).value
    Next
    
    mstr����IDs = Mid(mstr����IDs, 2)
    mstrסԺ�� = Mid(mstrסԺ��, 2)
    mstr���� = Mid(mstr����, 2)
    mstr���� = Mid(mstr����, 2)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdRef_Click()
    loadת����Ժ���� rptת��
End Sub

Private Sub Form_Load()
    '���õ���
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA  '150Ϊ͸����(0-255)
    tmrClose.Interval = 10
    tmrOpen.Interval = 10
    tmrOpen.Enabled = True
    tmrClose.Enabled = False
    lngAlpha = 5
    mblnCancle = False
    InitColor
    
    mblnCardOrder = (Val(zlDatabase.GetPara("��λ��Ƭ����ʽ", glngSys, P�°滤ʿվ, 0)) = 0)
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(0, "��Ժ����", pic��Ժ.hwnd, 0).Tag = "��Ժ"
        .InsertItem(1, "��Ժ����", pic��Ժ.hwnd, 0).Tag = "��Ժ"
        .InsertItem(2, "���ת��", picת��.hwnd, 0).Tag = "ת��"
       
        .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
        'ֻ����ѡ����Ӵ���
        Form_Resize
        Call tbcSub_SelectedChanged(.Selected)
    End With
    InitSelectTime
    InitReportColumn rpt��Ժ
    InitReportColumn rptת��
End Sub

Private Sub InitColor()
    Dim strValue As String
    Dim lng�ؼ� As Long, lngһ�� As Long, lng���� As Long, lng���� As Long
    Const c��ɫ As Long = 8388736
    Const c��ɫ As Long = 255
    Const c��ɫ As Long = 16711680
    Const c��ɫ As Long = 16777215
    
    Call DeleteFile
    mintIndex = 0
    imgHLDJ(999).ListImages.Clear
    '��ȡ����ȼ���������(����ȡȱʡ����)
    strValue = zlDatabase.GetPara("�ؼ�������ɫ", glngSys, 1265, "")
    lng�ؼ� = IIF(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("һ��������ɫ", glngSys, 1265, "")
    lngһ�� = IIF(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("����������ɫ", glngSys, 1265, "")
    lng���� = IIF(strValue = "", c��ɫ, Val(strValue))
    strValue = zlDatabase.GetPara("����������ɫ", glngSys, 1265, "")
    lng���� = IIF(strValue = "", c��ɫ, Val(strValue))
    
    '��ͼ
    mlngColor = lng�ؼ�
    Call DrawPoly
    mlngColor = lngһ��
    Call DrawPoly
    mlngColor = lng����
    Call DrawPoly
    mlngColor = lng����
    Call DrawPoly
End Sub

Private Sub AddColor()
    Dim strFile As String
    mintIndex = mintIndex + 1
    '������Ϊ�ļ�,���������ͼƬʱ,���뵽imagelist���ʼ��ֻ�����һ��,Ӧ��������image�б������ͼƬID���
    
    strFile = App.Path & "\HLDJTMP" & mintIndex & ".BMP"
    SavePicture picHLDJ.Image, strFile
    picHLDJ.Picture = LoadPicture(strFile)
    imgHLDJ(999).ListImages.Add , "K_" & mintIndex, picHLDJ.Picture
End Sub

Private Sub DrawPoly()
    Dim lngRgn As Long, lngBrush As Long
    Dim lngPen As Long, lngOldPen As Long
    Dim PtInPoly() As PointAPI

    '������򲢻�����
    ReDim PtInPoly(4) As PointAPI
    PtInPoly(1).X = 0
    PtInPoly(1).Y = 0
    PtInPoly(2).X = picHLDJ.ScaleWidth
    PtInPoly(2).Y = 0
    PtInPoly(3).X = picHLDJ.ScaleWidth
    PtInPoly(3).Y = picHLDJ.ScaleHeight
    PtInPoly(4).X = PtInPoly(1).X
    PtInPoly(4).Y = PtInPoly(1).Y
    
    '����ϵͳˢ��
    picHLDJ.Cls
    lngBrush = CreateSolidBrush(mlngColor)

    '�������ˢ�ӳɹ�,��ѡ��
    If lngBrush <> 0 Then
        lngRgn = CreatePolygonRgn(PtInPoly(1), UBound(PtInPoly), ALTERNATE)
        FillRgn picHLDJ.hDC, lngRgn, lngBrush
        Call DeleteObject(lngRgn)
        Call DeleteObject(lngBrush)
    End If
    picHLDJ.Refresh
    
    Call AddColor
End Sub

Private Function Get����ȼ�(ByVal str����ȼ� As String) As Integer
    '�������޵ȼ�ʱ,����3
    If InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "��") <> 0 Then
        Get����ȼ� = 0
    ElseIf InStr(1, str����ȼ�, "III") <> 0 Then
        Get����ȼ� = 3
    ElseIf InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "2") <> 0 Or InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "II") <> 0 Then
        Get����ȼ� = 2
    ElseIf InStr(1, str����ȼ�, "һ") <> 0 Or InStr(1, str����ȼ�, "1") <> 0 Or InStr(1, str����ȼ�, "��") <> 0 Or InStr(1, str����ȼ�, "I") <> 0 Then
        Get����ȼ� = 1
    Else
        Get����ȼ� = 3
    End If
End Function

Private Sub DeleteFile()
    Dim objFile As File
    For Each objFile In mobjFileSys.GetFolder(App.Path).Files
        If Left(objFile.Name, 7) = "HLDJTMP" Then
            mobjFileSys.DeleteFile objFile.Path, True
        End If
    Next
End Sub

Private Sub Load��Ժ����()
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Long, lngWcount As Long, lngWidth As Long, lngHeigh As Long, lngHcount As Long
    Dim int����ȼ� As Integer
    
    For i = mColBad.Count To 1 Step -1
        Unload lbl����(mColBad(i).Index)
        Unload img����ȼ�(mColBad(i).Index)
        Unload lbl����(mColBad(i).Index)
        Unload lblסԺ��(mColBad(i).Index)
        Unload lbl�Ա�(mColBad(i).Index)
        Unload lbl����(mColBad(i).Index)
        Unload lblSplit(mColBad(i).Index)
        Unload lblSelect(mColBad(i).Index)
        Unload mColBad(i)
        mColBad.Remove i
    Next
    
    strSQL = "Select distinct a.����id, a.��ҳid,A.��Ժ���� as ����, a.סԺ��, a.����, a.����, a.�Ա�, a.���ʱ��,A.��������," & vbNewLine & _
            "       Trunc(Sysdate) - Trunc(Decode(a.���ʱ��, Null, a.��Ժ����, a.���ʱ��)) || '��' As סԺ����,C.���� as ����ȼ�" & IIF(mblnCardOrder, "", ",d.���� as ��λ����") & vbNewLine & _
            "From ������ҳ A, ��Ժ���� B,�շ���ĿĿ¼ C" & IIF(mblnCardOrder, "", ",��λ���Ʒ��� D,��λ״����¼ F") & vbNewLine & _
            "Where a.����id = b.����id And a.��ҳid = b.��ҳid and A.����ȼ�ID = C.ID(+) " & IIF(mblnCardOrder, "", " and b.����id=f.����id(+) and f.��λ����=D.����(+)") & " And b.����id = [1]" & vbNewLine & _
            "order by " & IIF(mblnCardOrder, "", "��λ����,") & "LPAD(A.��Ժ����,10,' ')"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    lngWidth = picPati(999).Width
    lngHeigh = picPati(999).Height
    picIn��Ժ.Move 0, 0, pic��Ժ.Width - HScr.Width - 250
    
    lngWcount = (picIn��Ժ.Width) \ (lngWidth + 50)
    lngHcount = (pic��Ժ.Height) \ (lngHeigh + 50)

    
    For i = 0 To rsTmp.RecordCount - 1
        '�տ�Ƭ
        Load picPati(i)
        picPati(i).Move ((i Mod lngWcount)) * lngWidth + ((i Mod lngWcount) + 1) * 50, (i \ lngWcount) * lngHeigh + (i \ lngWcount + 1) * 50
        mColBad.Add picPati(i)
        picPati(i).Tag = rsTmp!����ID & "," & rsTmp!��ҳID
        picPati(i).Visible = True
        
        Load lbl����(i)
        Set lbl����(i).Container = picPati(i)
        lbl����(i).Caption = rsTmp!���� & ""
        lbl����(i).Visible = True
        
        Load img����ȼ�(i)
        img����ȼ�(i).Visible = True
        Set img����ȼ�(i).Container = picPati(i)
        Set img����ȼ�(i).Picture = Nothing
        img����ȼ�(i).ZOrder 1
        '���û���ȼ�(�ؼ���,һ����,������,������)
        int����ȼ� = Get����ȼ�(rsTmp!����ȼ� & "")
        Set img����ȼ�(i).Picture = imgHLDJ(999).ListImages(int����ȼ� + 1).Picture
        
        Load lbl����(i)
        Set lbl����(i).Container = picPati(i)
        lbl����(i).Caption = rsTmp!���� & ""
        lbl����(i).Visible = True
        
        Load lblסԺ��(i)
        Set lblסԺ��(i).Container = picPati(i)
        lblסԺ��(i).Caption = rsTmp!סԺ�� & ""
        lblסԺ��(i).Visible = True
        
        Load lbl�Ա�(i)
        Set lbl�Ա�(i).Container = picPati(i)
        lbl�Ա�(i).Caption = rsTmp!�Ա� & ""
        lbl�Ա�(i).Visible = True
        
        Load lbl����(i)
        Set lbl����(i).Container = picPati(i)
        lbl����(i).Caption = rsTmp!���� & ""
        lbl����(i).Visible = True
        
        Load lblSplit(i)
        Set lblSplit(i).Container = picPati(i)
        lblSplit(i).ZOrder 1
        lblSplit(i).BackColor = IIF(NVL(rsTmp!��������) = "��ͨ����", &HFFFFFF, zlDatabase.GetPatiColor(NVL(rsTmp!��������)))
        lblSplit(i).Visible = True
        
        Load lblSelect(i)
        Set lblSelect(i).Container = picPati(i)
        
        rsTmp.MoveNext
    Next
    picIn��Ժ_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tbcSub.Move 0, 0, Me.Width, Me.Height - cmdOK.Height - 700
    cmdOK.Move Me.Width - cmdOK.Width - cmdExit.Width - 600, Me.Height - cmdOK.Height - 630
    cmdExit.Move Me.Width - cmdExit.Width - 400, cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    If mblnCancle = False Then
        Cancel = 1
        lngAlpha = 250
        tmrClose.Enabled = True
    End If
    For i = mColBad.Count To 1 Step -1
        mColBad.Remove i
    Next
End Sub

Private Sub HScr_Change()
    picIn��Ժ.Top = -1 * (HScr.value / HScr.Max) * (picIn��Ժ.Height - pic��Ժ.Height + 800)
End Sub

Private Sub img����ȼ�_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub img����ȼ�_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lblSplit_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lblSplit_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl����_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl����_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl����_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl����_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl��Ժ����_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl��Ժ����_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl�Ա�_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl�Ա�_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lbl����_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lbl����_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lblסԺ��_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lblסԺ��_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub lblסԺ����_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub lblסԺ����_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub picIn��Ժ_Resize()
    Dim i As Long, lngWcount As Long, lngWidth As Long, lngHeigh As Long, lngHcount As Long
    
    On Error Resume Next
    If mColBad.Count > 0 Then
        lngWidth = picPati(999).Width
        lngHeigh = picPati(999).Height
        
        lngWcount = (picIn��Ժ.Width) \ (lngWidth + 50)
        lngHcount = (pic��Ժ.Height) \ (lngHeigh + 50)
        picIn��Ժ.Height = ((mColBad.Count \ lngWcount) + 1) * (lngHeigh + 50)
        
        For i = 0 To mColBad.Count - 1
            mColBad(i + 1).Move ((i Mod lngWcount)) * lngWidth + ((i Mod lngWcount) + 1) * 50, (i \ lngWcount) * lngHeigh + (i \ lngWcount + 1) * 50
        Next
        If mColBad.Count > lngHcount * lngWcount Then
            HScr.Visible = True
            HScr.Max = HScr.LargeChange * (mColBad.Count / (lngHcount * lngWcount)) - HScr.LargeChange
        Else
            HScr.Visible = False
        End If
    End If
End Sub

Private Sub picPati_Click(Index As Integer)
    SelectPati Index
End Sub

Private Sub SelectPati(Index As Integer)
    lblSelect(Index).Visible = Not lblSelect(Index).Visible
End Sub

Private Sub picPati_DblClick(Index As Integer)
    cmdOK_Click
End Sub

Private Sub pic��Ժ_Resize()
    rpt��Ժ.Move 0, rpt��Ժ.Top, pic��Ժ.Width, pic��Ժ.Height
End Sub

Private Sub pic��Ժ_Resize()
    On Error Resume Next
    HScr.Move pic��Ժ.Width - HScr.Width - 250, 0, HScr.Width, pic��Ժ.Height
    picIn��Ժ.Move 0, 0, pic��Ժ.Width - HScr.Width - 250
End Sub

Private Sub picת��_Resize()
    rptת��.Move 0, rptת��.Top, picת��.Width, picת��.Height
End Sub

Private Sub rpt��Ժ_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdOK_Click
End Sub

Private Sub rptת��_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdOK_Click
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Tag = "��Ժ" Then
        If pic��Ժ.Tag = "" Then
            Load��Ժ����
            pic��Ժ.Tag = "1"
        End If
    ElseIf Item.Tag = "��Ժ" Then
        If pic��Ժ.Tag = "" Then
            loadת����Ժ���� rpt��Ժ
            pic��Ժ.Tag = "1"
        End If
    ElseIf Item.Tag = "ת��" Then
        If picת��.Tag = "" Then
            loadת����Ժ���� rptת��
            picת��.Tag = "1"
        End If
    End If
End Sub

Private Sub loadת����Ժ����(objrpt As ReportControl)
    Dim strSQL As String, rsPati As Recordset
    Dim i As Long
    Dim objRecord As ReportRecord
    If Me.Visible = False Then Exit Sub
    With objrpt
        On Error GoTo errH
        If objrpt.Name = "rpt��Ժ" Then
            strSQL = "Select distinct a.����id, a.��ҳid,LPAD(A.��Ժ����,10,' ') as ����, a.סԺ��, a.����, a.����, a.�Ա�, a.���ʱ��,A.��Ժ����,A.��������," & vbNewLine & _
                "       Trunc(A.��Ժ����) - Trunc(Decode(a.���ʱ��, Null, a.��Ժ����, a.���ʱ��)) As סԺ����" & vbNewLine & _
                "From ������ҳ A" & vbNewLine & _
                "Where A.��ǰ����ID=[1]" & vbNewLine & _
                " And A.��Ժ���� Between to_date([2],'YYYY-MM-DD HH24:MI:SS') And to_date([3],'YYYY-MM-DD HH24:MI:SS')" & _
                "order by A.��Ժ���� Desc"
        ElseIf objrpt.Name = "rptת��" Then
            strSQL = "Select Distinct a.����id, a.��ҳid, LPad(a.��Ժ����, 10, ' ') As ����, a.סԺ��, a.����, a.����, a.�Ա�, a.���ʱ��, C.��ֹʱ�� as ��Ժ����, a.��������," & vbNewLine & _
                "                Trunc(Sysdate) - Trunc(Decode(a.���ʱ��, Null, a.��Ժ����, a.���ʱ��)) As סԺ����" & vbNewLine & _
                "From ������ҳ A, ���˱䶯��¼ C" & vbNewLine & _
                "Where a.����id = c.����id And a.��ҳid = c.��ҳid And a.��ǰ����id <> [1] And c.����id + 0 = [1] And Nvl(c.���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "      c.��ֹԭ�� In (3, 15) And  C.��ֹʱ�� Between Sysdate-[4] And Sysdate" & vbNewLine & _
                "Order By C.��ֹʱ�� Desc"

        End If
            
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Format(mdtOutBegin, "yyyy-MM-dd 00:00:00"), Format(mdtOutEnd, "yyyy-MM-dd 23:59:59"), Val(txtChange.Text))
        
        .Records.DeleteAll
        For i = 1 To rsPati.RecordCount
            Set objRecord = .Records.Add()
            objRecord.AddItem i
            objRecord.Tag = Val(rsPati!����ID) & "," & Val(rsPati!��ҳID)
            objRecord.AddItem Val(rsPati!����ID)
            objRecord.AddItem Val(rsPati!��ҳID)
            objRecord.AddItem CStr(NVL(rsPati!����))
            objRecord.AddItem CStr(NVL(rsPati!סԺ��))
            objRecord.AddItem CStr(NVL(rsPati!����))
            objRecord.AddItem CStr(NVL(rsPati!�Ա�))
            objRecord.AddItem CStr(NVL(rsPati!����))
            objRecord.AddItem CStr(NVL(rsPati!��������))
            objRecord.AddItem CStr(NVL(rsPati!���ʱ��))
            objRecord.AddItem CStr(NVL(rsPati!��Ժ����))
            objRecord.AddItem CStr(NVL(rsPati!סԺ����))
            objRecord.Item(COL_����).ForeColor = zlDatabase.GetPatiColor(NVL(rsPati!��������))
            
            rsPati.MoveNext
        Next
        .Populate 'ȱʡ��ѡ���κ���
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date
    
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mdtOutEnd = datCurr
    mdtOutBegin = mdtOutEnd - 1
    
    cboSelectTime.Clear '��Ժ
    With cboSelectTime
        .AddItem "������"
        .ItemData(.NewIndex) = 0
        .AddItem "������"
        .ItemData(.NewIndex) = 1
        .AddItem "ǰ����"
        .ItemData(.NewIndex) = 2
        .AddItem "һ����"
        .ItemData(.NewIndex) = 7
        .AddItem "30����"
        .ItemData(.NewIndex) = 30
        .AddItem "60����"
        .ItemData(.NewIndex) = 60
        .AddItem "[ָ��...]"
        .ItemData(.NewIndex) = -1
    End With
    If cboSelectTime.ListCount > 0 Then cboSelectTime.ListIndex = 0
End Sub

Private Sub cboSelectTime_Click()
'���ܣ���ʱ�䷶Χ��ָ���ǣ�����ʱ��ѡ����
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    intDateCount = cboSelectTime.ItemData(cboSelectTime.ListIndex)
    datCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    If cboSelectTime.ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
    If intDateCount = -1 Then
        If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboSelectTime) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call cbo.SetIndex(cboSelectTime.hwnd, mintOutPreTime)
            Exit Sub
        End If
    Else
        mdtOutEnd = datCurr
        mdtOutBegin = mdtOutEnd - intDateCount
    End If
    If mdtOutBegin = CDate(0) Or mdtOutEnd = CDate(0) Then
        cboSelectTime.ToolTipText = ""
    Else
        cboSelectTime.ToolTipText = "��Χ��" & Format(mdtOutBegin, "yyyy-MM-dd") & " �� " & Format(mdtOutEnd, "yyyy-MM-dd")
    End If
    '�����������֤ÿ���ط���ȡ�ĳ�Ժ���˶�����ͬһʱ�䷶Χ�ڣ�72783��
    mintOutPreTime = cboSelectTime.ListIndex
    loadת����Ժ���� rpt��Ժ
End Sub

Private Sub tmrOpen_Timer()
    lngAlpha = lngAlpha + 10
     SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
     SetLayeredWindowAttributes Me.hwnd, 0, lngAlpha, LWA_ALPHA  '150Ϊ͸����(0-255)
     If lngAlpha >= 255 Then tmrOpen.Enabled = False
End Sub

Private Sub tmrClose_Timer()
    lngAlpha = lngAlpha - 10
     SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
     SetLayeredWindowAttributes Me.hwnd, 0, lngAlpha, LWA_ALPHA  '150Ϊ͸����(0-255)
     If lngAlpha <= 5 Then tmrClose.Enabled = False: mblnCancle = True: Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = HScr.value
    lngMin = HScr.Min
    lngMax = HScr.Max
    
    If KeyCode = vbKeyPageDown Then '��
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            HScr.value = lngCur + (lngMax - lngMin) / 10
        Else
            HScr.value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '��
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            HScr.value = lngCur - (lngMax - lngMin) / 10
        Else
            HScr.value = lngMin
        End If
    End If
    
End Sub

Private Sub Form_Activate()
'������
    If picIn��Ժ.Visible Then
        glngPreHWnd = GetWindowLong(picIn��Ժ.hwnd, GWL_WNDPROC)
        SetWindowLong picIn��Ժ.hwnd, GWL_WNDPROC, AddressOf FlexScroll
    End If
End Sub

Private Sub InitReportColumn(obj As Object)
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With obj
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(COL_���, "���", 30, True)
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_��ҳID, "��ҳID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 80, True)
        Set objCol = .Columns.Add(COL_סԺ��, "סԺ��", 90, True)
        Set objCol = .Columns.Add(COL_����, "����", 70, True)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 50, True)
        Set objCol = .Columns.Add(col_����, "����", 50, True)
        Set objCol = .Columns.Add(col_��������, "��������", 120, True)
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 120, True)
        Set objCol = .Columns.Add(col_��Ժ����, IIF(obj.Name = "rpt��Ժ", "��Ժ����", "ת������"), 120, True)
        Set objCol = .Columns.Add(col_סԺ����, "סԺ����", 60, True)
      

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        '.MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.imgPati
        
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(col_��Ժ����)
        .SortOrder(0).SortAscending = False
    End With
    
    
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    loadת����Ժ���� rptת��
End Sub

Private Sub txtChange_GotFocus()
    Call zlControl.TxtSelAll(txtChange)
End Sub
