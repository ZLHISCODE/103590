VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmReportImageEdit 
   Caption         =   "����ͼƬ�༭"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12960
   Icon            =   "frmReportImageEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12960
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCboDropDown 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5520
      Picture         =   "frmReportImageEdit.frx":0E42
      ScaleHeight     =   375
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   6600
      Width           =   255
   End
   Begin VB.ListBox lstMemoText 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3200
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame frmLabels 
      Height          =   5175
      Left            =   6120
      TabIndex        =   12
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "�Զ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   11
         Left            =   2865
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "�Զ������ֱ�ע"
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtUserLabelText 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   960
         TabIndex        =   34
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "Xn=ֱ�ӻ�첿λ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   10
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "���ֱ�ע"
         Top             =   4020
         Width           =   3000
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "P=��״Ѫ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   9
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "���ֱ�ע"
         Top             =   3480
         Width           =   3000
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "V=�ǵ���Ѫ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   8
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "���ֱ�ע"
         Top             =   2940
         Width           =   3000
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "AT=�쳣ת����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   7
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "���ֱ�ע"
         Top             =   2400
         Width           =   3000
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "W=�����ɫ��Ƥ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   6
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "���ֱ�ע"
         Top             =   1860
         Width           =   3000
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "I=�����԰�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   5
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "���ֱ�ע"
         Top             =   1320
         Width           =   1500
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "C=ʪ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   4
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "���ֱ�ע"
         Top             =   1320
         Width           =   1500
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "L=ճĤ�װ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "���ֱ�ע"
         Top             =   780
         Width           =   1500
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "M=��Ƕ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "���ֱ�ע"
         Top             =   780
         Width           =   1500
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "E=������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "���ֱ�ע"
         Top             =   240
         Width           =   1500
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   9
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "���ֱ��9"
         Top             =   4560
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   8
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "���ֱ��8"
         Top             =   4080
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   7
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "���ֱ��7"
         Top             =   3600
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   6
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "���ֱ��6"
         Top             =   3120
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   5
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "���ֱ��5"
         Top             =   2640
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "���ֱ��4"
         Top             =   2160
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "���ֱ��3"
         Top             =   1680
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "���ֱ��2"
         Top             =   1200
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "���ֱ��1"
         Top             =   720
         Width           =   450
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "�Զ��������ֱ��"
         Top             =   240
         Width           =   450
      End
      Begin VB.CommandButton cmdTextLabel 
         Caption         =   "Po=Ϣ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "���ֱ�ע"
         ToolTipText     =   "���ֱ�ע"
         Top             =   240
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdFont 
      Height          =   375
      Left            =   7680
      Picture         =   "frmReportImageEdit.frx":119E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "���õ�ǰ��ע���塣"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��"
      Height          =   400
      Left            =   2400
      TabIndex        =   9
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCur 
      Caption         =   "��һ��"
      Height          =   400
      Left            =   1080
      TabIndex        =   8
      Top             =   6600
      Width           =   1100
   End
   Begin VB.ComboBox cbxMemoText 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2280
      TabIndex        =   7
      Top             =   5880
      Width           =   5055
   End
   Begin VB.CommandButton cmdInsert 
      Height          =   375
      Left            =   7320
      Picture         =   "frmReportImageEdit.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "����ǰ��ע����Ϊ���ñ�ע"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   400
      Left            =   9120
      TabIndex        =   5
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���"
      Height          =   400
      Left            =   7680
      TabIndex        =   4
      Top             =   6600
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   4680
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtInputText 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   3495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      _Version        =   262147
      _ExtentX        =   9128
      _ExtentY        =   6165
      _StockProps     =   35
      BackColor       =   -2147483638
      UseScrollBars   =   0   'False
   End
   Begin VB.Label lblMemoText 
      AutoSize        =   -1  'True
      Caption         =   "��ӱ�ע���֣�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   11
      Top             =   5955
      Width           =   1470
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportImageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TPoint
  X As Integer
  Y As Integer
End Type

Private mlngModule As Long
Private mImage As DicomImage
Private mintMouseState As TMouseState
Private mblnDcmViewDown As Boolean
Private mMouseDownPoint As TPoint
Private mInitScrollPoint As TPoint
Private mCorpSize As TPoint             '�϶�������ƫ��λ��

'����������ʹ�õ�����׼λ��
Private mlngBaseXX As Long
Private mlngBaseYY As Long
'�ƶ���עʹ�õ�����׼λ��
Private mlngBaseX As Long
Private mlngBaseY As Long

Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע
Private mMovingLabel As DicomLabel      '��ǰѡ��Ҫ�ƶ�����ɾ���ı�ע

Private mblnOK As Boolean
Private mOldImage As DicomImage
Private mintCurImgIndex As Integer      '������ѡ������ͼ������
Private mfrmParent As Object            '������ģ�����
Private mSelViewerIndex As Integer      '�����屻ѡ�еı���ͼ���ID����1��ʼ����
Private mblnIsMark As Boolean           '�Ǳ��ͼ
Private mintTextIndex As Integer        '���ֱ�ע��ť������
Private mintNumberIndex As Integer      '���ֱ�Ű�ť������
Private mintAutoNumber As Integer       '�Զ�������ŵ�������

Private mrsTmp As ADODB.Recordset       'ͼ��ע��¼��

Private Enum TMouseState
    msNone = 0          '��״̬
    msWinLevel = 1      '����λ
    msZoom = 2          '����
    msRectangle = 3     '��ѡ����
    msArrow = 11        '��ͷ
    msEllipse = 12      '��Բ
    msText = 13         '����
    msDrag = 14         '�����϶�
    msNumber = 15       '���ֱ��
    msFixText = 16      '���ְ�ť
    msMove = 17         '�ƶ���ɾ����ע
End Enum

Public Sub zlShowMe(ByVal img As DicomImage, frmParent As frmReportImage, _
    intCurImgIndex As Integer, SelViewerIndex As Integer, ByVal lngModule As Long)
    
    On Error GoTo err
    
    Dim i As Integer
    'ȥ���߿�
    For i = 1 To img.Labels.Count
        If img.Labels(i).tag = "SELECT" Or img.Labels(i).tag = "BORDER" Then
            img.Labels(i).Visible = False
        End If
    Next
    
    Set mOldImage = img

    mlngModule = lngModule
    mintCurImgIndex = intCurImgIndex
    mSelViewerIndex = SelViewerIndex
    Set mfrmParent = frmParent
    
    '��ͼ���ע��������ͼ
    If mintCurImgIndex = 0 Then
        mblnIsMark = True
    Else
        mblnIsMark = False
    End If
    
    cmdNext.Visible = Not mblnIsMark
    cmdCur.Visible = Not mblnIsMark
    Me.lblMemoText.Visible = Not mblnIsMark
    Me.cbxMemoText.Visible = Not mblnIsMark
    Me.picCboDropDown.Visible = Not mblnIsMark
    Me.cmdInsert.Visible = Not mblnIsMark
    Me.cmdFont.Visible = Not mblnIsMark
    
    Me.DViewer.Images.Clear
    Me.DViewer.Images.Add img
    '�ؽ���ע֮��Ĺ���
    Call subLabelCopyRebuild(img, Me.DViewer.Images(1))
    
    Me.Show 0, frmParent
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '�������ö�
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ChangeImage(intType As Integer)
'intType �л����� 1 --��һ��ͼ��2--��һ��ͼ
    Dim i As Integer
    
    Me.DViewer.Images.Clear
    If intType = 1 Then  '��һ��ͼ
        If mintCurImgIndex <= 1 Then
            Call mfrmParent.MovePage(mtLast)
            mintCurImgIndex = mfrmParent.ImageCount
        Else
            mintCurImgIndex = mintCurImgIndex - 1
        End If
        
        Me.DViewer.Images.Add mfrmParent.dcmImages(mintCurImgIndex)
    ElseIf intType = 2 Then   '��һ��ͼ
        If mintCurImgIndex >= mfrmParent.ImageCount Then
            Call mfrmParent.MovePage(mtNext)
            mintCurImgIndex = 1
        Else
            mintCurImgIndex = mintCurImgIndex + 1
        End If
        
        
        Me.DViewer.Images.Add mfrmParent.dcmImages(mintCurImgIndex)
    End If
    
    '���ѡ��ͼ�εı߿���ɫ
    Me.DViewer.Images(1).BorderColour = vbRed
    
    '�Ը���������ͼ�ı߿���ɫ���д���
    For i = 1 To mfrmParent.ImageCount
        mfrmParent.dcmImages(i).BorderColour = vbWhite
    Next i
    
    Set mfrmParent.mSelMiniImg = mfrmParent.dcmImages(mintCurImgIndex)
    mfrmParent.mSelMiniImg.BorderColour = vbRed
    
    '���ComboBox�ı�
    zlControl.CboSetIndex cbxMemoText.hWnd, -1
    
    '�ر�������
    If lstMemoText.Visible Then lstMemoText.Visible = False
End Sub

Private Function getListIndex() As Integer
'���ݼ���������ȡ����
    Dim i As Integer

    getListIndex = -1
    
    If mrsTmp.RecordCount <= 0 Then Exit Function

    mrsTmp.MoveFirst
    
    If cbxMemoText.Text = "" Then Exit Function

    For i = 0 To mrsTmp.RecordCount - 1
        If InStr(Trim(Nvl(mrsTmp!����)), UCase(cbxMemoText.Text)) > 0 Or InStr(Trim(Nvl(mrsTmp!����)), UCase(cbxMemoText.Text)) > 0 Then
            getListIndex = i
            
            Exit For
        End If

        mrsTmp.MoveNext
    Next
End Function

Private Sub cbxMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
    End If
End Sub

Private Sub cmdNum_Click(Index As Integer)
    mintNumberIndex = Index
    subSetMouseState msNumber
    
    Call setCmdLabelColor
    
    cmdNum(Index).BackColor = &HC0C000
End Sub

Private Sub cmdTextLabel_Click(Index As Integer)
    
    If Index = 11 Then '�Զ��壬���ж��Ƿ�����������
        If Trim(txtUserLabelText.Text) = "" Then
            MsgBoxD Me, "�������Զ����ע��", vbOKOnly, gstrSysName
            txtUserLabelText.SetFocus
            Exit Sub
        End If
    End If
    mintTextIndex = Index
    subSetMouseState msFixText
    
    Call setCmdLabelColor
    
    cmdTextLabel(Index).BackColor = &HC0C000
End Sub

Private Sub setCmdLabelColor()
    Dim i As Integer
    
    For i = 0 To cmdTextLabel.Count - 1
        cmdTextLabel(i).BackColor = &H8000000F
    Next i
    
    For i = 0 To cmdNum.Count - 1
        cmdNum(i).BackColor = &H8000000F
    Next i
End Sub

Private Sub DViewer_DblClick()
    Dim ls As DicomLabels
    Dim l As DicomLabel
    
    On Error GoTo err
    
    If mintMouseState = msMove Then
        Set ls = DViewer.LabelHits(mlngBaseXX, mlngBaseYY, False, False, True)
        If ls.Count > 0 Then
            If MsgBoxD(Me, "�Ƿ�ɾ�������ע��", vbOKCancel, gstrSysName) = vbOK Then
                Set l = ls(1)
                If l.tag <> "" Then
                    '�Ǳ�ű�ע����Ҫͬʱɾ��������ע����ɾ������
                    If DViewer.Images(1).Labels.IndexOf(l.TagObject.TagObject) <> 0 Then
                        Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l.TagObject.TagObject))
                    End If
                    If DViewer.Images(1).Labels.IndexOf(l.TagObject) <> 0 Then
                        Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l.TagObject))
                    End If
                End If
                '����ͨ��ע�����߱�ŵ����һ����ע��ֱ��ɾ������
                Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l))
                DViewer.Refresh
            End If
        End If
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub lstMemoText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
End Sub

Private Sub picCboDropDown_Click()
    lstMemoText.Visible = Not lstMemoText.Visible
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex

    If lstMemoText.Visible Then lstMemoText.SetFocus
End Sub

Private Sub cbxMemoText_Change()
    If Not lstMemoText.Visible Then lstMemoText.Visible = True
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
End Sub

Private Sub cbxMemoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cbxMemoText.ListIndex = lstMemoText.ListIndex
        lstMemoText.Visible = False
        
        cbxMemoText.SelStart = 0
        cbxMemoText.SelLength = Len(cbxMemoText.Text)
        cbxMemoText.SetFocus
    End If
    
    If KeyAscii = vbKeyEscape Then lstMemoText.Visible = False
End Sub

Private Sub cmdCur_Click()
'��һ��ͼ��
On Error GoTo errH
 
    Call ChangeImage(1)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdFont_Click()
On Error GoTo ErrHandle
    diaFont.flags = 1
    diaFont.FontBold = Me.Font.Bold
    diaFont.FontItalic = Me.Font.Italic
    diaFont.FontName = Me.Font.Name
    diaFont.FontSize = Me.Font.Size
    diaFont.FontStrikethru = Me.Font.Strikethrough
    diaFont.FontUnderline = Me.Font.Underline

    
    diaFont.ShowFont
    
    Me.Font.Bold = diaFont.FontBold
    Me.Font.Italic = diaFont.FontItalic
    Me.Font.Name = diaFont.FontName
    Me.Font.Size = diaFont.FontSize
    Me.Font.Strikethrough = diaFont.FontStrikethru
    Me.Font.Underline = diaFont.FontUnderline
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdNext_Click()
'��һ��ͼ��
On Error GoTo errH
 
    Call ChangeImage(2)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAdd_Click()
'------------------------------------------------
'���ܣ���Ӳ��������رմ���
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    Dim dcmGlobal As New DicomGlobal
    
    dcmGlobal.RegString("UIDRoot") = "1"
    mblnOK = True
    'ƴ�ӷ���
    Call subAddMemoText
    
    If mblnOK Then
        If Me.DViewer.Images.Count = 1 Then
            Set mImage = Me.DViewer.Images(1)
            mImage.InstanceUID = dcmGlobal.NewUID   'ͼ������ͼ��󣬾������µ�InstanceUID
        Else
            Set mImage = Nothing
        End If
    Else
        Set mImage = Nothing
    End If
    
    If mblnIsMark = True Then   '���ͼ������ӱ��ͼ��ֱ���˳�
        Call mfrmParent.DcmAddMarkImage(mImage)
        Unload Me
        Exit Sub
    Else    '�ɼ���ͼ����
        '��ƴ�Ӻ��ͼ��ı߿���д���
         If Me.DViewer.Images.Count > 0 Then
             With Me.DViewer.Images(1)
                .BorderWidth = 3
                .BorderStyle = 2
                .BorderColour = vbRed
            End With
        End If
        
        Call mfrmParent.DcmAddImage(mImage, mSelViewerIndex)
    End If
    
    Me.DViewer.Refresh
    
    '���ComboBox�ı�
    cbxMemoText.Text = ""
    
    '�ر�������
    lstMemoText.Visible = False
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Sub cmdExit_Click()
'���Viewer�ؼ�����ж�ش���
   ' Me.DViewer.Images.Clear
    Unload Me
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
  On Error GoTo ErrHandle
    Select Case control.ID
        Case conMenu_Process_Window         '���ȶԱȶ�
            subSetMouseState 1
            'Control.Checked = True
            
        Case conMenu_Process_Zoom           '����
            subSetMouseState 2
            'Control.Checked = True
            
        Case conMenu_Process_RectZoom       '�ü�����
            subSetMouseState msRectangle
            'Control.Checked = True
        
        Case conMenu_Process_RectCapture         '�ü���ɼ�
            Call CaptureFrameSelectImage
            
        Case conMenu_Process_RRotate        '˳ʱ����ת
            subSetRotate True
            
        Case conMenu_Process_LRotate        '��ʱ����ת
            subSetRotate False
            
        Case conMenu_Process_Sharpness      '��
            subSetSharp True
            
        Case conMenu_Process_Filter         'ƽ��
            subSetSharp False
            
        Case conMenu_Process_Corp          '�϶�
           subSetMouseState msDrag
            
        Case conMenu_Process_Arrow          '��ͷ��ע
            subSetMouseState msArrow
            
        Case conMenu_Process_Ellipse        'Բ�α�ע
            subSetMouseState msEllipse
            
        Case conMenu_Process_Text           '���ֱ�ע
            subSetMouseState msText
        
        Case conMenu_Process_DelAllLabels   '�����ע
            DViewer.Images(1).Labels.Clear
            DViewer.Refresh
            
        Case conMenu_Process_MoveLabel      '�ƶ���ע
            subSetMouseState msMove
            
        Case conMenu_Process_LabelSetUp     '��ע����
            Call subSetTextLabel
            
        Case conMenu_Process_Restore        '�ָ�
            DViewer.Images.Clear
            DViewer.Images.Add mOldImage
            '�ؽ���ע֮��Ĺ���
            Call subLabelCopyRebuild(mOldImage, Me.DViewer.Images(1))
            mintAutoNumber = 0  '�ָ���ͼ��ʱ��������
    End Select
    
    If control.ID <> conMenu_Process_LabelSetUp Then
        Call setCmdLabelColor
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub subSetSharp(blnSharp As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ���ƽ������
'������blnSharp��ʾͼ����ķ���True=�񻯣�False=ƽ��
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If DViewer.Images.Count > 0 Then
        If blnSharp = True Then
            '�񻯴���
            If DViewer.Images(1).FilterLength <= 0 Then
                DViewer.Images(1).FilterLength = 0
                '��ǰû��ƽ������ֱ�ӽ����񻯴���
                DViewer.Images(1).UnsharpEnhancement = DViewer.Images(1).UnsharpEnhancement + 0.1
            Else
                '�����ǰ�Ѿ���ƽ���������ȵ���ƽ��Ч��
                DViewer.Images(1).FilterLength = DViewer.Images(1).FilterLength - 1
            End If
        Else
            'ƽ������
            '�ж�Zoom�Ƿ�1������ǣ����޸�Ϊ0.9999
            If DViewer.Images(1).ActualZoom = 1 Then
                DViewer.Images(1).Zoom = 0.9999
            End If
            
            If DViewer.Images(1).UnsharpEnhancement <= 0 Then
                DViewer.Images(1).UnsharpEnhancement = 0
                '��ǰû���񻯴���ֱ�ӿ�ʼƽ��
                '�ж�FilterLength�Ƿ�0����ǣ�����2/ActualZoom��2��FilterLength֮����е���
                If DViewer.Images(1).FilterLength = 0 Then
                    DViewer.Images(1).FilterLength = 2 / DViewer.Images(1).ActualZoom + 1
                Else    '���������FilterLength��1
                    DViewer.Images(1).FilterLength = DViewer.Images(1).FilterLength + 1
                End If
            Else
                '��ǰ�Ѿ������񻯴����ȵ����񻯵�Ч��
                DViewer.Images(1).UnsharpEnhancement = DViewer.Images(1).UnsharpEnhancement - 0.1
            End If
        End If
    End If
End Sub

Private Sub subSetRotate(blnClockwise As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ�����ת
'������blnClockwise��ת�ķ���,True=˳ʱ����ת��False=��ʱ����ת
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If DViewer.Images.Count > 0 Then
        Dim iRotateState As Integer
        
        iRotateState = DViewer.Images(1).RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        If iRotateState = -1 Then iRotateState = 3
        iRotateState = iRotateState Mod 4
        DViewer.Images(1).RotateState = iRotateState
    End If
End Sub


'DicomViewer�ü���ɼ�ͼ��
Private Sub CaptureFrameSelectImage()
    Dim imgResult As DicomImage
    Dim imgs As New DicomImages
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim iMax As Integer
    Dim img As DicomImage
    Dim lblFrame As DicomLabel
    
    If Me.DViewer.Images.Count <> 1 Then Exit Sub
    If Me.DViewer.Images(1).Labels.Count < 1 Then Exit Sub
    
    Set img = Me.DViewer.Images(1)
    Set lblFrame = Me.DViewer.Images(1).Labels(Me.DViewer.Images(1).Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgBoxD Me, "��ѡ��ͼ��������ٱ���", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    'ͼ�������=300
    iMax = 300
    
    '����label����ȡ����ѡ�е�ͼ��
    'ͼ��λ��,�ڰ�ͼ��Ϊ1����ɫͼ��Ϊ3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).value = "RGB" Or img.Attributes(&H28, &H4).value = "YBR_FULL_422" Then
            iPlane = 3
        End If
    End If
    
    'ͼ����λ��
    If lblFrame.Width >= 0 Then
        iLeft = lblFrame.Left
        iRight = iLeft + lblFrame.Width
    Else
        iLeft = lblFrame.Left + lblFrame.Width
        iRight = lblFrame.Left
    End If
    
    If lblFrame.Height >= 0 Then
        iTop = lblFrame.Top
        iBottom = iTop + lblFrame.Height
    Else
        iTop = lblFrame.Top + lblFrame.Height
        iBottom = lblFrame.Top
    End If
    
    '���ƽ��ͼ��Ĵ�С��300*300֮��
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.SizeY - iBottom, img.SizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X��Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, img.SizeY - iBottom, img.SizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    DViewer.Images.Clear
    DViewer.Images.Add imgResult
    
End Sub

Private Sub subSetMouseState(intMoustState As TMouseState)
'------------------------------------------------
'���ܣ��������״̬��ͬʱ���¹�������ť��ѡ��״̬
'������intMoustState -- ���״̬
'���أ���
'------------------------------------------------
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_MoveLabel).Checked = False
        
    '�ı䵱ǰ���״̬
    If mintMouseState = intMoustState And ((intMoustState <> msFixText) And _
        (intMoustState <> msNumber) And (intMoustState <> msMove)) Then
        mintMouseState = msNone
    Else
        mintMouseState = intMoustState
        
        Select Case mintMouseState
            Case msWinLevel: cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = True
            Case msZoom: cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = True
            Case msRectangle: cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = True
            Case msArrow: cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = True
            Case msEllipse: cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = True
            Case msText: cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = True
            Case msDrag: cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = True
            Case msMove: cbrMain.FindControl(xtpControlButton, conMenu_Process_MoveLabel).Checked = True
        End Select
    End If
    
End Sub

Private Sub cbrMain_Resize()
    '������ʾ�Ŀͻ�����
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    On Error Resume Next
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    '�ڷ�DViewer
    Me.DViewer.Left = lngLeft
    Me.DViewer.Top = lngTop
    Me.DViewer.Width = Abs(lngRight - lngLeft - Me.frmLabels.Width)
    If mblnIsMark = True Then
        Me.DViewer.Height = Abs(lngBottom - lngTop - 800)
    Else
        Me.DViewer.Height = Abs(lngBottom - lngTop - 1300)
    End If
    
    '�ڷű�ע��ť
    Me.frmLabels.Left = Me.DViewer.Width
    Me.frmLabels.Height = Me.DViewer.Height
    Me.frmLabels.Top = Me.DViewer.Top
    
    '�ڷű�ע����
    Me.lblMemoText.Left = 100
    Me.lblMemoText.Top = Me.ScaleHeight - 1100

    Me.cbxMemoText.Left = Me.lblMemoText.Left + Me.lblMemoText.Width
    Me.cbxMemoText.Top = Me.lblMemoText.Top - 100
    Me.cbxMemoText.Width = Abs(Me.ScaleWidth - Me.cbxMemoText.Left - 250 - cmdInsert.Width - cmdFont.Width)

    Me.lstMemoText.Left = Me.cbxMemoText.Left
    Me.lstMemoText.Top = Me.cbxMemoText.Top - Me.lstMemoText.Height
    Me.lstMemoText.Width = Me.cbxMemoText.Width - 10
    
    Me.picCboDropDown.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width - 260
    Me.picCboDropDown.Top = Me.cbxMemoText.Top + 30
    
    Me.cmdInsert.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width
    Me.cmdInsert.Top = Me.cbxMemoText.Top

    Me.cmdFont.Left = Me.cmdInsert.Left + Me.cmdInsert.Width
    Me.cmdFont.Top = Me.cmdInsert.Top

    '�ڷš���ӡ������˳�����ť
    Me.cmdAdd.Left = Me.ScaleWidth - Me.cmdAdd.Width * 3
    Me.cmdAdd.Top = Me.ScaleHeight - 600

    Me.cmdExit.Left = Me.ScaleWidth - Me.cmdExit.Width * 1.8
    Me.cmdExit.Top = Me.cmdAdd.Top

    '�ڷš���һ����������һ������ť
    Me.cmdCur.Left = Me.ScaleWidth / 15
    Me.cmdCur.Top = Me.ScaleHeight - 600

    Me.cmdNext.Left = Me.cmdCur.Width + Me.cmdCur.Left + 200
    Me.cmdNext.Top = Me.cmdAdd.Top
    
End Sub



Private Sub cmdInsert_Click()
    Dim strSQL As String, i As Integer
    
    If Trim(cbxMemoText.Text) = "" Then
        MsgBoxD Me, "�����뱸ע���ݡ�", vbInformation, gstrSysName
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    End If
    If cbxMemoText.ListIndex <> -1 Then
        MsgBoxD Me, "�ñ�ע�����Ѿ��ڳ��ñ�ע�С�", vbInformation, gstrSysName
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    Else
        For i = 0 To cbxMemoText.ListCount - 1
            If UCase(Trim(cbxMemoText.list(i))) = UCase(Trim(cbxMemoText.Text)) Then
                MsgBoxD Me, "�ñ�ע���Ѿ��ڳ��ñ�ע�С�", vbInformation, gstrSysName
                If cbxMemoText.Enabled Then cbxMemoText.SetFocus
                Exit Sub
            End If
        Next
    End If
        
    On Error GoTo errH
    
    strSQL = zlCommFun.zlGetSymbol(cbxMemoText.Text)
    strSQL = "zl_Ӱ��ͼ��ע_Insert('" & Replace(cbxMemoText.Text, "'", "''") & "','" & strSQL & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem cbxMemoText.hWnd, CB_ADDSTRING, 0, cbxMemoText.Text
    lstMemoText.AddItem cbxMemoText.Text
    MsgBoxD Me, "������Ϊ���ñ�ע��", vbInformation, gstrSysName
    If cbxMemoText.Enabled Then cbxMemoText.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subAddMemoText()
'------------------------------------------------
'���ܣ���ͼ����ӱ�ע����
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    Dim img As DicomImage
    Dim iLeft As Integer
    Dim iWidth As Integer
    Dim iTop As Integer
    Dim iHeight As Integer
    Dim imgResult As New DicomImage
    Dim iPlane As Integer
    Dim lngWhiteX As Long
    Dim lngWhiteY As Long
    Dim lngFontHeight As Long
    
    If Me.DViewer.Images.Count <> 1 Then Exit Sub
    
    If Trim(cbxMemoText.Text) = "" Then Exit Sub
    
    lngFontHeight = ScaleY(TextHeight(cbxMemoText.Text), vbTwips, vbPixels) + 6
    
    '�ѱ�ע������ӵ�ͼ����
    Set img = Me.DViewer.Images(1)
    
    iLeft = 0
    iTop = 0
    iWidth = img.SizeX
    iHeight = img.SizeY + lngFontHeight

    'ʹ��PrinterImage���������Խ�ͼ���ϵı�ǩ����עͬʱ���л���
    Set imgResult = img.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight - lngFontHeight)
'

    '��ӱ�ע
    Dim dlMemoText As New DicomLabel
    
    dlMemoText.LabelType = doLabelText
    dlMemoText.ImageTied = True
    dlMemoText.Transparent = False
    dlMemoText.AutoSize = False
    dlMemoText.Left = 0
    dlMemoText.Top = img.SizeY
    dlMemoText.Width = iWidth
    dlMemoText.Height = lngFontHeight
    
    dlMemoText.BackColour = vbWhite
    dlMemoText.ForeColour = vbBlack
            
    dlMemoText.Font.Name = Me.Font.Name
    dlMemoText.Font.Italic = Me.Font.Italic
    dlMemoText.Font.Strikethrough = Me.Font.Strikethrough
    dlMemoText.Font.Underline = Me.Font.Underline
    dlMemoText.Font.Size = Me.Font.Size
    dlMemoText.Font.Bold = Me.Font.Bold
    dlMemoText.FontName = Me.Font.Name
    dlMemoText.FontSize = Me.Font.Size
    dlMemoText.ShowTextBox = True
    
    dlMemoText.Text = Me.cbxMemoText.Text & "                                                                                                                                 "
    
    imgResult.Labels.Add dlMemoText
    
    Set imgResult = imgResult.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight)

    '����ͼ��
    Me.DViewer.Images.Clear
    Me.DViewer.Images.Add imgResult
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim ls As DicomLabels
    Dim lngLeftD As Long
    
    If Button = 1 And DViewer.Images.Count > 0 Then
        Dim intLabelType As Integer
        
        mMouseDownPoint.X = DViewer.Images(1).ActualScrollX
        mMouseDownPoint.Y = DViewer.Images(1).ActualScrollY
          
        mInitScrollPoint.X = DViewer.Images(1).ScrollX + X
        mInitScrollPoint.Y = DViewer.Images(1).ScrollY + Y
        
        mblnDcmViewDown = True
        If mintMouseState <> msNone Then
            '��¼��ǰ���λ��
            mlngBaseXX = X
            mlngBaseYY = Y
            Select Case mintMouseState
                Case msArrow, msEllipse, msText, msRectangle, msFixText, msNumber      '��ͷ����Բ�����֣���ѡ���̶����֣�˳����
                    If mintMouseState = msArrow Then
                        intLabelType = doLabelArrow
                    ElseIf mintMouseState = msEllipse Or mintMouseState = msNumber Then
                        intLabelType = doLabelEllipse
                    ElseIf mintMouseState = msText Or mintMouseState = msFixText Then
                        intLabelType = doLabelText
                    ElseIf mintMouseState = msRectangle Then
                        intLabelType = doLabelRectangle
                    End If
                    
                    If mintMouseState = msFixText Then
                        '����ǵ������֣�λ�Ƶ���Ҫ����
                        If mintTextIndex = 11 Then
                            lngLeftD = IIf(Len(txtUserLabelText.Text) = 1, 3, 7)
                        Else
                            lngLeftD = IIf(Len(Left(cmdTextLabel(mintTextIndex).Caption, InStr(cmdTextLabel(mintTextIndex).Caption, "=") - 1)) = 1, 3, 7)
                        End If
                    Else
                        lngLeftD = 7
                    End If
                    DViewer.Images(1).Labels.Add GetNewLabel(intLabelType, DViewer.ImageXPosition(X, Y) - lngLeftD, DViewer.ImageYPosition(X, Y) - 7, 0, 0)
                    Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
                    If intLabelType = doLabelArrow Then
                        '��ͷ��Ҫʹ���߿�=2
                        mdcmSelectLabel.LineWidth = 4
                    ElseIf intLabelType = doLabelText Then
                        mdcmSelectLabel.XOR = False
                        mdcmSelectLabel.ForeColour = vbBlack
                        If mblnIsMark = False Then
                            '���Ǳ��ͼ������������ӱ��������ͼ�������ӣ���Ϊ���Ӳ�����֧�֣���ӡ��ʱ��Ͳ�֧��
                            mdcmSelectLabel.Transparent = False
                            mdcmSelectLabel.ForeColour = vbWhite
                            mdcmSelectLabel.BackColour = vbBlack
                        End If
                        '���������С
                        If DViewer.Images(1).SizeX <= 256 Then
                            mdcmSelectLabel.FontSize = 10
                        ElseIf DViewer.Images(1).SizeX <= 512 Then
                            mdcmSelectLabel.FontSize = 15
                        Else
                            mdcmSelectLabel.FontSize = 18
                        End If
                        
                    End If
                Case msMove     '�ƶ���ע
                    Set ls = DViewer.LabelHits(X, Y, False, False, True)
                    mlngBaseX = DViewer.ImageXPosition(X, Y)
                    mlngBaseY = DViewer.ImageYPosition(X, Y)
                    If ls.Count > 0 Then    '���ѡ�����κ�һ����ע
                        '���Tag=""˵���Ǽ򵥱�ע���ǿ�˵�������ֱ�ű�ע����Ҫ�ҵ����ֱ�ע
                        Set mMovingLabel = ls(1)
                        If mMovingLabel.tag <> "" Then
                            If mMovingLabel.tag = m_LabelTag_Back Then
                                Set mMovingLabel = mMovingLabel.TagObject
                            ElseIf mMovingLabel.tag = m_LabelTag_Circle Then
                                Set mMovingLabel = mMovingLabel.TagObject.TagObject
                            End If
                        End If
                    End If
            End Select
        End If
    End If
End Sub

Private Sub DViewer_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        Select Case mintMouseState
            Case msWinLevel   '���ȶԱȶ�
                DViewer.Images(1).Width = DViewer.Images(1).Width + (X - mlngBaseXX)
                DViewer.Images(1).Level = DViewer.Images(1).Level + (Y - mlngBaseYY)
                mlngBaseXX = X
                mlngBaseYY = Y
            Case msZoom   '����
                Dim dblZoom As Double
                dblZoom = DViewer.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseYY) * 0.001)
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom DViewer.Images(1), DViewer, dblZoom, mCorpSize
                End If
                mlngBaseYY = Y
'            Case msRectangle  '�ü�����
'                Dim dcmLabel As DicomLabel
'                dcmView.Labels.Clear
'                Set dcmLabel = dcmView.Labels.AddNew
'                dcmLabel.LabelType = doLabelRectangle
'                dcmLabel.Left = mlngBaseXX
'                dcmLabel.Top = mlngBaseYY
'                dcmLabel.Width = x - mlngBaseXX
'                dcmLabel.Height = y - mlngBaseYY
            Case msArrow, msEllipse, msRectangle    '��ͷ��ע'Բ�α�ע,��ѡ
                mdcmSelectLabel.Width = DViewer.ImageXPosition(X, Y) - mdcmSelectLabel.Left
                mdcmSelectLabel.Height = DViewer.ImageYPosition(X, Y) - mdcmSelectLabel.Top
            Case msDrag
                '�϶�ͼ��......
                DViewer.Images(1).ScrollX = mInitScrollPoint.X - X
                DViewer.Images(1).ScrollY = mInitScrollPoint.Y - Y
            Case msMove
                '�ƶ���ע
                If Not mMovingLabel Is Nothing Then
                    subaCorrectCursor DViewer, DViewer.Images(1), X, Y  '����ƶ��������ͼ��Χ�����������λ��
                    subMoveLable mMovingLabel, DViewer.ImageXPosition(X, Y) - mlngBaseX, DViewer.ImageYPosition(X, Y) - mlngBaseY
                    mlngBaseX = DViewer.ImageXPosition(X, Y)
                    mlngBaseY = DViewer.ImageYPosition(X, Y)
                End If
        End Select
        
        DViewer.Refresh
    End If
End Sub

Private Sub DViewer_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        mblnDcmViewDown = False
        If mintMouseState = msText Then      '���ֱ�ע
            
            txtInputText.Left = Me.ScaleX(X, vbPixels, vbTwips) + DViewer.Left
            txtInputText.Top = Me.ScaleY(Y, vbPixels, vbTwips) + DViewer.Top
            
            txtInputText.Text = ""
            txtInputText.Visible = True
            txtInputText.SetFocus
        ElseIf mintMouseState = msRectangle Then   '�ü�����
            
            '��ʾͼ�񱣴�˵�
            Call ShowFrameSelectImagePopup
            'ɾ����ѡ�õ���ʱ��ע
            If DViewer.Images(1).Labels.Count > 0 Then
                DViewer.Images(1).Labels.Remove DViewer.Images(1).Labels.Count
            End If
            
            Set mdcmSelectLabel = Nothing
            
'            dcmView.Labels.Clear
'            RectangleZoom dcmView, dcmView.Images(1), mlngBaseXX, mlngBaseYY, x - mlngBaseXX, y - mlngBaseYY
        ElseIf mintMouseState = msDrag Then
            '����ͼ�����ε�ƫ��λ��
            mCorpSize.X = mCorpSize.X + (DViewer.Images(1).ActualScrollX - mMouseDownPoint.X)
            mCorpSize.Y = mCorpSize.Y + (DViewer.Images(1).ActualScrollY - mMouseDownPoint.Y)
        ElseIf mintMouseState = msFixText Then
            '��ӹ̶�����
            If mintTextIndex = 11 Then  '�Զ������ֱ�ע
                mdcmSelectLabel.Text = txtUserLabelText.Text
            Else
                mdcmSelectLabel.Text = Left(cmdTextLabel(mintTextIndex).Caption, InStr(cmdTextLabel(mintTextIndex).Caption, "=") - 1)
            End If
        ElseIf mintMouseState = msNumber Then
            Dim intText As Integer
            
            If mintNumberIndex = 0 Then '�Զ�˳����
                mintAutoNumber = mintAutoNumber + 1
                intText = mintAutoNumber
            Else
                intText = mintNumberIndex
            End If
            '���˳����
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.BackColour = glngColor(intText Mod 9 + 1)
            mdcmSelectLabel.Transparent = False
            mdcmSelectLabel.Width = 14
            mdcmSelectLabel.Height = 14
            mdcmSelectLabel.tag = m_LabelTag_Back
            
            '���˳����Բ�ε��������ӱ�ע��Բ�ο������
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelEllipse, mdcmSelectLabel.Left, mdcmSelectLabel.Top, 14, 14)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.tag = m_LabelTag_Circle
            mdcmSelectLabel.TagObject = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 1)
            
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelText, mdcmSelectLabel.Left + 1, mdcmSelectLabel.Top, 0, 0)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.tag = m_LabelTag_Number
            mdcmSelectLabel.FontSize = 8
            mdcmSelectLabel.FontName = "Arial Bold"
            mdcmSelectLabel.AutoSize = True
            mdcmSelectLabel.Text = intText
            If mdcmSelectLabel.Text < 10 Then
                mdcmSelectLabel.Left = mdcmSelectLabel.Left + 3
            End If
            mdcmSelectLabel.TagObject = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 1)
            DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 2).TagObject = mdcmSelectLabel    'TagObject�γɱջ�
        End If
        
        DViewer.Refresh
    End If
End Sub

Public Sub ShowFrameSelectImagePopup()
'------------------------------------------------
'���ܣ�������ѡͼ���ʱ�� ������Ҽ��ĵ����˵�
'������
'���أ���
'------------------------------------------------

Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '����Ҽ������˵�
    Set cbrToolBar = Me.cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectCapture, "ȷ�ϲü�")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'���ܣ���ͼ��������š��Ե�ǰviewer���ĵ�Ϊ�������ĵ㡣
'������
'       img -- �������ŵ�ͼ��
'       viewer ���� ͼ�����ڵ�viewer
'       dblZoom ����ͼ���µ����ű���
'���أ��ޣ�ֱ�ӵ���ͼ������ű���
'�ϼ���������̣�frmViewer.Viewer_MouseMove
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ� �ƽ� 2006-2-10
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False

            
    img.ScrollX = (img.SizeX * img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub


Private Sub Form_Load()
    On Error GoTo err

    '�ָ�����λ��
    Call RestoreWinState(Me, App.ProductName)
    
    '����������
    Call InitCommandBars
    
    Call LoadMemoFontStyle
    
    mCorpSize.X = 0
    mCorpSize.Y = 0
    mblnOK = False
    mintAutoNumber = 0
    
    'ͼ����������Ĭ���ǵ�����ͼ���ע��������Ĭ�����ƶ���ע
    If mblnIsMark = True Then
        Call subSetMouseState(msMove)
    Else
        Call subSetMouseState(msWinLevel)
    End If
    
    Call ReadEnjoin
    
    Call subLoadTextLabel
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

'���뱸ע������ʽ
Private Sub LoadMemoFontStyle()
    Dim strFontStyle As String
    Dim aryFontStyle() As String
    
    '������,12,B,U,S,I��
    
    strFontStyle = zlDatabase.GetPara("ͼ��ע����", glngSys, mlngModule, "")
    
    strFontStyle = strFontStyle & ",,,,,,"
    
    aryFontStyle = Split(strFontStyle, ",")
    
    If aryFontStyle(0) <> "" Then Me.Font.Name = aryFontStyle(0)
    If Val(aryFontStyle(1)) <> 0 Then Me.Font.Size = Val(aryFontStyle(1))
    If UCase(aryFontStyle(2)) = "B" Then Me.Font.Bold = True
    If UCase(aryFontStyle(3)) = "U" Then Me.Font.Underline = True
    If UCase(aryFontStyle(4)) = "S" Then Me.Font.Strikethrough = True
    If UCase(aryFontStyle(5)) = "I" Then Me.Font.Italic = True
End Sub


Private Sub SaveMemoFontStyle()
    Dim strFontStyle As String
    
    strFontStyle = Me.Font.Name & "," & _
        Me.Font.Size & "," & _
        IIf(Me.Font.Bold, "B", "") & "," & _
        IIf(Me.Font.Underline, "U", "") & "," & _
        IIf(Me.Font.Strikethrough, "S", "") & "," & _
        IIf(Me.Font.Italic, "I", "")

    Call zlDatabase.SetPara("ͼ��ע����", strFontStyle, glngSys, mlngModule)
End Sub


Private Function ReadEnjoin() As Boolean
'���ܣ���ȡ�����볣�ñ�ע
    Dim strSQL As String, strPre As String
        
    On Error GoTo errH
    
    '��������
    strPre = cbxMemoText.Text '����󱣳�ԭ��ֵ
    cbxMemoText.Clear
    
    strSQL = _
        " Select ����,���� From Ӱ��ͼ��ע Where ���� is Not Null And ��Ա=[1]" & _
        " Union" & _
        " Select ����,���� From Ӱ��ͼ��ע Where ���� is Not Null And ��Ա is Null" & _
        " Order by ����"
    Set mrsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.����)
    Do While Not mrsTmp.EOF
        AddComboItem cbxMemoText.hWnd, CB_ADDSTRING, 0, mrsTmp!����
        
        lstMemoText.AddItem mrsTmp!����
        mrsTmp.MoveNext
    Loop
    cbxMemoText.Text = strPre
    
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Unload(Cancel As Integer)
    '���洰��λ��
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveMemoFontStyle
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
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
        .SetIconSize True, 32, 32
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    'ͼ���������������
    Set cbrToolBar = Me.cbrMain.Add("ͼ�������", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True '�ı���ʾ��ͼ���·�
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Window, "����"): cbrControl.ToolTipText = "��������/�Աȶ�": cbrControl.Visible = Not mblnIsMark
        cbrControl.Checked = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Zoom, "����"): cbrControl.ToolTipText = "����ͼ��": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Corp, "�϶�"): cbrControl.ToolTipText = "�϶�ͼ��": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectZoom, "�ü�"): cbrControl.ToolTipText = "�ü��ɼ�ͼ��": cbrControl.IconId = 3201: cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "˳ʱ"): cbrControl.ToolTipText = "˳ʱ����ת": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "��ʱ"): cbrControl.ToolTipText = "��ʱ����ת": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Sharpness, "��"): cbrControl.ToolTipText = "��": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Filter, "ƽ��"): cbrControl.ToolTipText = "ƽ��": cbrControl.Visible = Not mblnIsMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Arrow, "��ͷ"): cbrControl.ToolTipText = "��ͷ��ע": cbrControl.Visible = Not mblnIsMark
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Ellipse, "Բ��"): cbrControl.ToolTipText = "Բ�α�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Text, "����"): cbrControl.ToolTipText = "���ֱ�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_MoveLabel, "�ƶ���ע"): cbrControl.ToolTipText = "��������ק�ƶ���ע��˫��ɾ����ע"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LabelSetUp, "���ñ�ע"): cbrControl.ToolTipText = "�������ֱ�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_DelAllLabels, "�����ע"): cbrControl.ToolTipText = "���ȫ����ע"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Restore, "�ָ�"): cbrControl.ToolTipText = "�ָ�ͼ�񵽳�ʼ״̬"
        cbrControl.BeginGroup = True
    End With
    For Each cbrControl In cbrToolBar.Controls
         cbrControl.Style = xtpButtonIconAndCaption
         cbrControl.Category = "Main" '���ó�������˵�
    Next
    cbrToolBar.Position = xtpBarTop
End Sub

Private Sub lstMemoText_DblClick()
    cbxMemoText.Text = lstMemoText.list(lstMemoText.ListIndex)
    lstMemoText.Visible = False
    
    cbxMemoText.SelStart = 0
    cbxMemoText.SelLength = Len(cbxMemoText.Text)
    cbxMemoText.SetFocus
End Sub

Private Sub lstMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
    End If
End Sub

Private Sub lstMemoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
    End If
    
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then lstMemoText.Visible = False
End Sub

Private Sub picCboDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 1
End Sub

Private Sub picCboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 0
End Sub

Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then  '''ESC�ͻس����˳�����
        txtInputText.Visible = False
        If Trim(txtInputText.Text) = "" Then
            'ɾ�����ֱ�ע
            DViewer.Images(1).Labels.Remove DViewer.Images(1).Labels.Count
            txtInputText = "1 "
        Else
            mdcmSelectLabel.Text = txtInputText.Text
            DViewer.Refresh
        End If
    End If
End Sub

Private Sub subaCorrectCursor(v As DicomViewer, im As DicomImage, xx As Long, Yy As Long)
'------------------------------------------------
'���ܣ�����ƶ��������ͼ��Χ�����������λ��
'������v--ͼ�����ڵ�viewer��im--������ڵ�ͼ��xx--������ڵ�x����λ�ã������곬��ͼ���򽫴�ֵ�޸ĵ�ͼ��֮�ڣ�
'      yy--������ڵ�y����λ�ã������곬��ͼ���򽫴�ֵ�޸ĵ�ͼ��֮�ڣ�
'���أ���
'------------------------------------------------
    Dim X As Integer, Y As Integer, w As Long, h As Long
    Dim i As DicomImage
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    w = v.Width / v.MultiColumns / Screen.TwipsPerPixelX - v.CellSpacing * 2
    h = v.Height / v.MultiRows / Screen.TwipsPerPixelY - v.CellSpacing * 2
    X = im.OriginX + v.CellSpacing
    Y = im.OriginY + v.CellSpacing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If xx < X Then xx = X
    If xx > X + w Then xx = X + w
    If Yy < Y Then Yy = Y
    If Yy > Y + h Then Yy = Y + h
End Sub

Public Sub subMoveLable(la As DicomLabel, X As Long, Y As Long)
'------------------------------------------------
'���ܣ��ƶ�һ����ע
'������la--���ƶ��ı�ע��x--x�����ƶ���ͼ�����ؾ��룻y--y�����ƶ���ͼ�����ؾ���
'���أ���
'------------------------------------------------
    
    la.Left = la.Left + X
    la.Top = la.Top + Y
    
    '��������ֱ�ţ���Ҫͬʱ�ƶ�������ע
    If la.tag <> "" And Not la.TagObject Is Nothing Then
        la.TagObject.Left = la.TagObject.Left + X
        la.TagObject.Top = la.TagObject.Top + Y
        la.TagObject.TagObject.Left = la.TagObject.TagObject.Left + X
        la.TagObject.TagObject.Top = la.TagObject.TagObject.Top + Y
    End If
       
End Sub

Private Sub subSetTextLabel()
'------------------------------------------------
'���ܣ��������ֱ�ע��������
'������
'���أ���
'------------------------------------------------
    Dim strTemp As String
    Dim i As Integer

    On Error GoTo err
    
    If mintMouseState <> msFixText Or mintTextIndex = 11 Then
        MsgBoxD Me, "����ѡ��һ�����ֱ�ע��ť��Ȼ��������á�", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    strTemp = InputBox("�������µ����ֱ�ע���ã���ʽΪ������=˵������", "���ֱ�ע����", cmdTextLabel(mintTextIndex).Caption)
    
    If strTemp = "" Then Exit Sub
    
    If InStr(strTemp, "=") = 0 Then
        MsgBoxD Me, "����ĸ�ʽ����ȷ��Ӧ�ð��ա�����=˵������ʽ���룬������������á�", vbOKOnly, gstrSysName
        Exit Sub
    End If
     
    '����ɹ���ʹ������µ����ֱ�ע��ͬʱ���浽ע�����
    cmdTextLabel(mintTextIndex).Caption = strTemp
    
    strTemp = ""
    For i = 0 To cmdTextLabel.Count - 2
        strTemp = strTemp & "[+]" & cmdTextLabel(i).Caption
    Next i
    strTemp = Mid(strTemp, 4)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmReportImageEdit", "�������ֱ�ע", strTemp)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subLoadTextLabel()
'------------------------------------------------
'���ܣ���ȡ���ֱ�ע
'������
'���أ���
'------------------------------------------------
    Dim strTemp As String
    Dim strText() As String
    Dim i  As Integer
    
    On Error GoTo err
    
    strTemp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmReportImageEdit", "�������ֱ�ע", "")
    
    If strTemp = "" Then
        'ʹ��Ĭ��ֵ������Ҫ����
        Exit Sub
    End If
    
    strText = Split(strTemp, "[+]")
    If UBound(strText) <> 10 Then
        '���ݲ����ϸ�ʽ��ʹ��Ĭ��ֵ
        Exit Sub
    End If
    
    For i = 0 To 10
        cmdTextLabel(i).Caption = strText(i)
    Next i
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subLabelCopyRebuild(Simg As DicomImage, oImg As DicomImage)
'------------------------------------------------
'���ܣ��ؽ�ͼ��ı�ע������ϵ
'������sImg--Դͼ��oImg--Ŀ��ͼ��
'���أ���
'------------------------------------------------
    Dim l As DicomLabel
    For Each l In oImg.Labels
        If Not l.TagObject Is Nothing Then
            If Simg.Labels.IndexOf(l.TagObject) <> 0 Then
                Set l.TagObject = oImg.Labels(Simg.Labels.IndexOf(l.TagObject))
            End If
        End If
    Next
End Sub
