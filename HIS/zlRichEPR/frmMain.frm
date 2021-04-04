VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7D52C334-5021-43A4-8EB4-86CC21862ABF}#1.2#0"; "zlTable.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "zlRichEPR"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   7170
   ScaleWidth      =   9930
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtFeedBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   300
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   500
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Left            =   4560
      MaxLength       =   500
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   5000
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2310
      Left            =   6450
      ScaleHeight     =   2310
      ScaleWidth      =   3690
      TabIndex        =   20
      Top             =   1110
      Width           =   3690
      Begin VB.PictureBox picHistoryInfo 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   105
         ScaleHeight     =   360
         ScaleWidth      =   1500
         TabIndex        =   22
         Top             =   60
         Width           =   1500
         Begin VB.Label lblHistoryInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ǰ 5 �����ʷ����"
            Height          =   180
            Left            =   60
            TabIndex        =   23
            Top             =   75
            Width           =   1530
         End
      End
      Begin zlRichEditor.Editor edtThis 
         Height          =   2610
         Left            =   270
         TabIndex        =   21
         Top             =   645
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   4604
         ShowRuler       =   0   'False
      End
   End
   Begin zlRichEPR.ucPacsImgCanvas ucPacsImgCanvas1 
      Height          =   915
      Left            =   2160
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
      _extentx        =   1296
      _extenty        =   1614
   End
   Begin zlRichEPR.ucPictureEditor ucPictureEditor1 
      Height          =   975
      Left            =   1080
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   855
      _extentx        =   1508
      _extenty        =   1720
   End
   Begin VB.Timer tmrAutoSaveEPR 
      Interval        =   1000
      Left            =   9000
      Top             =   765
   End
   Begin VB.PictureBox picPatiInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   45
      MouseIcon       =   "frmMain.frx":058A
      ScaleHeight     =   375
      ScaleWidth      =   9735
      TabIndex        =   7
      Top             =   3735
      Width           =   9735
      Begin VB.Label lblPatiInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         Height          =   180
         Left            =   90
         TabIndex        =   12
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lblPatiIns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00008000&
         Height          =   180
         Index           =   1
         Left            =   7410
         TabIndex        =   11
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lblPatiIns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����:"
         Height          =   180
         Index           =   0
         Left            =   6750
         TabIndex        =   10
         Top             =   90
         Width           =   630
      End
      Begin VB.Label lblPatiState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   9480
         TabIndex        =   9
         Top             =   90
         Width           =   90
      End
      Begin VB.Label lblPatiState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:"
         Height          =   180
         Index           =   0
         Left            =   8985
         TabIndex        =   8
         Top             =   90
         Width           =   450
      End
   End
   Begin VB.PictureBox picPenInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   5315
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   4185
      Visible         =   0   'False
      Width           =   1030
      Begin VB.TextBox txtPenInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   20
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   20
         Width           =   1005
      End
   End
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8370
      Top             =   720
   End
   Begin VB.PictureBox picTMP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   7815
      ScaleHeight     =   420
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   5700
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picDropDown 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   7875
      ScaleHeight     =   315
      ScaleWidth      =   270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5175
      Visible         =   0   'False
      Width           =   330
   End
   Begin zlRichEditor.Editor edtBuff 
      Height          =   600
      Left            =   360
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
   End
   Begin zlRichEPR.F1ColorPicker ColorFillColor 
      Height          =   2190
      Left            =   5940
      TabIndex        =   2
      Top             =   1170
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      AutoColor       =   16777215
   End
   Begin zlTable.Table tblThis 
      Height          =   1230
      Left            =   5400
      TabIndex        =   4
      Top             =   4905
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   2170
      SingleLine      =   0   'False
   End
   Begin zlRichEPR.ColorPicker ColorPaperBackColor 
      Height          =   2190
      Left            =   5760
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      AutoColor       =   16777215
   End
   Begin zlRichEPR.ColorPicker ColorHighlight 
      Height          =   2190
      Left            =   5580
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   810
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      AutoColor       =   16777215
   End
   Begin zlRichEPR.ColorPicker ColorForeColor 
      Height          =   2190
      Left            =   5400
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      Color           =   0
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   2790
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   8370
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C74
            Key             =   "HIGHLIGHT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DE9
            Key             =   "FORECOLOR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F36
            Key             =   "FILLCOLOR"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   17
      Top             =   6792
      Width           =   9936
      _ExtentX        =   17515
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4339
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2716
            MinWidth        =   2716
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1658
            MinWidth        =   1658
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "Ins"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
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
   Begin zlRichEditor.Editor Editor1 
      Height          =   2715
      Left            =   255
      TabIndex        =   13
      Top             =   810
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4789
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   5160
      MousePointer    =   7  'Size N S
      Top             =   2925
      Width           =   5115
   End
   Begin XtremeCommandBars.CommandBars cbrThis 
      Left            =   1215
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DkpThis 
      Bindings        =   "frmMain.frx":10AA
      Left            =   300
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'## ȫ�ֱ���
'######################################################################################################################
Public Document As cEPRDocument                         '�ĵ������
Public glngCurEleKey As Long                            '��ǰԪ��ID

Public WithEvents mfrmCompends As frmCompends           '�ĵ��ṹͼ
Attribute mfrmCompends.VB_VarHelpID = -1

Private WithEvents mfrmSentenceDetailed As frmSentenceDetailed       'ʾ���ʾ䴰��
Attribute mfrmSentenceDetailed.VB_VarHelpID = -1
Private WithEvents mfrmSegments As frmSegmentList        'ʾ��Ƭ�δ���
Attribute mfrmSegments.VB_VarHelpID = -1
Private WithEvents mfrmModElement As frmElementEdit     '���ݱ༭����
Attribute mfrmModElement.VB_VarHelpID = -1
Private WithEvents mfrmInsElement As frmInsElement      '��������Ҫ�ش���
Attribute mfrmInsElement.VB_VarHelpID = -1
Private WithEvents mfrmDicSelect As frmDicSelect        '�����ֵ���Ŀ
Attribute mfrmDicSelect.VB_VarHelpID = -1
Private WithEvents mfrmStyleMan As frmStyleMan          '������ʽά��
Attribute mfrmStyleMan.VB_VarHelpID = -1
Private WithEvents cPicEditor As cPictureEditor         'λͼ�༭������
Attribute cPicEditor.VB_VarHelpID = -1
Private WithEvents mfrmMultiDocView As frmMultiDocView  '���ĵ���ϲ���
Attribute mfrmMultiDocView.VB_VarHelpID = -1
Private WithEvents mfrmPacsPic As frmPACSImg            'PACSͼƬ�б���
Attribute mfrmPacsPic.VB_VarHelpID = -1
Private WithEvents mfrmMainError As frmMainMsg
Attribute mfrmMainError.VB_VarHelpID = -1
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mfrmPreview As frmPrintPreview
Attribute mfrmPreview.VB_VarHelpID = -1
Private WithEvents mfrmDocksymbol As frmDockSymbol      '�������
Attribute mfrmDocksymbol.VB_VarHelpID = -1
Private WithEvents mfrmHistoryReport As frmDockReportHistory       '��ʷ���洰��
Attribute mfrmHistoryReport.VB_VarHelpID = -1

Private cDropDown As cDropDownToolWindow                '���ѡ����

Private mlngHP As Long, blnSpaceEvent As Boolean        '��¼�Զ����ӿո��λ�ã�
Private lngYOld As Long, lngXOld As Long, blnIsDown As Boolean
Private lngIndex As Long                                '�ĵ�Key���½�Ϊ��δ�����ĵ�1������δ�����ĵ�2���������ȣ��������ƣ�
Private lngPicPosition As Long                          'ͼƬ�༭λ��
Private lngTablePosition As Long                        '���༭λ��
Private mblnExistHistroy As Boolean                     '�Ƿ�����ʷ�����б�
Private mlngSelHightlightColor As OLE_COLOR             '��ǰѡ�е����ָ���ɫ
Private mlngSelForeColor As OLE_COLOR                   '��ǰѡ�е�����ǰ��ɫ
Private mlngCellFillColor As OLE_COLOR                  '��Ԫ�����ɫ

Private mParaFmt As cParaFormat                         '�洢��ʽˢ��������
Private mFontFmt As cFontFormat                         '�洢��ʽˢ��������
Private mblnFmtBrushDown As Boolean                     '��ʽˢ�������

Private mblnIsMultiMode As Boolean                      '�Ƿ��Ƕ��ĵ��༭ģʽ
Private mintStyle As Integer                              'Ƕ��༭ -1 ��ģ̬ vbModeless=0 ģ̬ vbModal=1

Private Type PatiInfor
    ����    As String
    ���֤�� As String
End Type
Private mPatiInfor As PatiInfor
Private mblnPatiSign As Boolean
Private mblnEnPtSign As Boolean

Private Type UndoInfo
    Filename As String
    SelStart As Long
    SelEnd As Long
End Type

Private Type EleLimit
    �䶯ԭ�� As Byte
    ԭ��Ҫ��id As Long
    ԭ��Ҫ�� As String
    ԭ������ As String
    �䶯��� As Byte
    ������id As Long
    ���Ҫ��id As Long
    ���Ҫ�� As String
    ���ֵ�� As String
    ԭʼֵ�� As String
End Type
Private mEleLimit() As EleLimit
Private mlDiseaseID As Long
Private mlDiagnoseID As Long

Private UndoList() As UndoInfo                          '��ʷ�ļ��б�1��ʼ���
Private p_Undo As Long                                  '��ǰ��Undoָ��
Private mblnAutosave As Boolean                         '�Ƿ����Զ�����
Private mlngUndoLimit As Long                           'Undo�������Ĭ��20��
Private mlngSaveInterval As Long                        '�Զ�����ʱ��������
Private mblnAutoSaveEPR As Boolean                      '�Ƿ����Զ�����
Private mlngSaveIntervalEPR As Long                     '�Զ�����ʱ����������
Private mblnAutoPageCount As Boolean                    '�Զ���ҳ����
Private mblnAutoPageNote As Boolean                     '�Զ���ҳ����
Private mintSharePages As Integer                       '��ʾ����ҳ���ļ����ݵ�����
Private mblnNoAsk As Boolean                            '��Ĭ��ӡ
Private mblnSignAutoAlter As Boolean                    '���Ƶ���,ǩ���Զ���λ
Private DT1_EPR As Date, DT2_EPR As Date
Private DT1 As Date, DT2 As Date, mblnChange As Boolean
Private mbEditInTable As Boolean                        '�Ƿ�ǰ�ڱ��༭״̬
Private mblnǩ��Ҫ�� As Boolean                         '�Ƿ��п��õ�ǩ��Ҫ��
Private mblnChildMode As Boolean                        '�Ƿ���Ƕ��༭���Ӵ���
Private mblnCanPrint As Boolean                         '�Ƿ���Դ�ӡԤ��
Private mblnReadOnly As Boolean                         '�Ƿ�ֻ��
Private mstrSex As String                               '���˵��Ա�
Private mblnPrecess As Boolean                          '�Ƿ����ڴ�������,�����н�ֹ�رմ���
Private mbln���޴��� As Boolean                         '�Ƿ��Ƿ��޴���Ⱦ�����濨
Private mblnFBContentChanged As Boolean                 '���޴���˵���Ƿ��޸���

'######################################################################################################################
Public Property Get ReadOnly() As Boolean
    ReadOnly = mblnReadOnly
End Property

Public Property Let ReadOnly(vData As Boolean)
    mblnReadOnly = vData
    Editor1.ReadOnly = vData
    If vData Then
        DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Close
        DkpThis.FindPane(ID_VIEW_SEGMENT).Close
    Else
        DkpThis.ShowPane ID_VIEW_PHRASEDEMO
        DkpThis.ShowPane ID_VIEW_SEGMENT
    End If
End Property

Public Property Get CanPrint() As Boolean
    CanPrint = mblnCanPrint
End Property

Public Property Let CanPrint(vData As Boolean)
    mblnCanPrint = vData
End Property

Public Property Get ChildMode() As Boolean
    ChildMode = mblnChildMode
End Property

Public Property Let ChildMode(vData As Boolean)
    mblnChildMode = vData
    If mblnChildMode Then
        Me.BorderStyle = 0
        SetWindowLong Me.hwnd, GWL_STYLE, GetWindowLong(Me.hwnd, GWL_STYLE) Xor WS_BORDER Xor WS_THICKFRAME Xor WS_DLGFRAME
    Else
        Me.BorderStyle = 2
    End If
End Property

Private Function AutoMoveSignPos() As Boolean

    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long
    
    With Editor1
        lS = .Selection.StartPos
        If .SelLength > 0 Then Exit Function
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            Editor1.SelStart = lKEE
        End If
'        If sKeyType = "S" Then
'            If Editor1.Selection.StartPos >= lKSS And Editor1.Selection.StartPos <= lKSE And lKSS > 0 And lKSE > 0 Then
'                If lKSS > 0 Then Editor1.SelStart = lKSS
'            End If
'            If Editor1.Selection.StartPos >= lKES And Editor1.Selection.StartPos <= lKEE And lKES > 0 And lKEE > 0 Then
'                Editor1.SelStart = lKEE
'            End If
'        End If
    End With
                    
    AutoMoveSignPos = True
    
End Function
Private Function ShowSharePageHistory(ByVal Document As cEPRDocument, Optional ByVal intNumber As Integer = 5) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ����ҳ���ļ���ʷ����
    '������ Document                    ��ǰ�ĵ�
    '       intNumber                   Ҫ��ʾ��ʷ�Ĵ���
    '���أ� ���򷵻��棬�����
    '******************************************************************************************************************
    Dim lngTMP As Long, rsTemp As New ADODB.Recordset, strTime As String, strIDs As String, varPar() As String
    Dim strFile As String, strZipFile As String, lngLen2 As Long, lngLen1 As Long, lngStart As Long
    Dim objEPRFileInfo As New cEPRFileDefineInfo
    Dim strTmpClipboard As String
    On Error GoTo errHand
    
    strTmpClipboard = Clipboard.GetText '��ʱ��¼ճ�������ݣ��������ڲ����õ�ճ���������ճ������
    If mintSharePages = 0 Then Exit Function
    
    edtThis.ReadOnly = False
    edtBuff.ReadOnly = False
    edtThis.Freeze
    edtThis.NewDoc
    edtBuff.NewDoc
    
    strTime = Format(Document.EPRPatiRecInfo.����ʱ��, "yyyy-MM-dd HH:mm:ss")
    lblHistoryInfo.Caption = "ǰ " & intNumber & " �ε���ʷ����"
    '���ҵ�ǰ����ǰ���intNumber�ݲ���
    strIDs = GetFileRange(Document.EPRFileInfo.ID, Document.EPRPatiRecInfo.ID, strTime, Document.EPRPatiRecInfo.��������, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, False, Document.EPRPatiRecInfo.ҽ��id)
    
    gstrSQL = "Select /*+ rule*/ a.Id, a.�ļ�id, a.��������, a.��������, a.����id, a.��ҳid,a.����ʱ��, a.���汾, a.������, a.���ʱ��, a.����ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ A," & LongIDsTable(strIDs, varPar, 2) & vbNewLine & _
                "Where a.Id = b.Id" & vbNewLine & _
                "Order By a.���, a.����ʱ�� Desc"
    gstrSQL = "Select ID,����ʱ�� From (" & gstrSQL & ") Where RowNum<=[1] Order by ����ʱ��" '�޴�������ʱ�䷴������
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", intNumber, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9))
    If rsTemp.BOF = False Then
        edtThis.ForceEdit = True
        edtBuff.ForceEdit = True

        Do While Not rsTemp.EOF
            strZipFile = zlBlobRead(5, Val(rsTemp!ID))

            If gobjFSO.FileExists(strZipFile) Then
                strFile = zlFileUnzip(strZipFile)
                If gobjFSO.FileExists(strFile) Then

                    edtBuff.OpenDoc strFile
                    
                    lngTMP = Val(rsTemp!ID)

                    lngLen1 = Len(edtBuff.Text)
                    lngLen2 = Len(edtThis.Text)

                    edtThis.Range(lngLen2, lngLen2).Selected
                    edtBuff.SelectAll

                    edtBuff.CopyWithFormat
                    edtThis.PasteWithFormat

                    lngStart = Len(edtThis.Text)

                    If rsTemp.AbsolutePosition < rsTemp.RecordCount Then
                        'ĩβ��֤��һ���س�
                        If edtThis.Range(lngStart - 2, lngStart) = vbCrLf Then
                            edtThis.Range(lngStart - 2, lngStart).Font.Hidden = False
                        Else
                            edtThis.Range(lngStart, lngStart).Text = vbCrLf
                            edtThis.Range(lngStart, lngStart + 2).Font.Hidden = False
                        End If
                    End If
                    edtThis.TOM.TextDocument.Range(lngStart, lngStart).Para = edtBuff.TOM.TextDocument.Range(lngLen1, lngLen1).Para
                End If
                
                If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile
                If gobjFSO.FileExists(strZipFile) Then gobjFSO.DeleteFile strZipFile
            End If

            rsTemp.MoveNext
        Loop

        gstrSQL = "Select c.ID, a.��ʽ From   ����ҳ���ʽ a, �����ļ��б� b, ���Ӳ�����¼ c " & _
                " Where  c.�ļ�id = b.id And a.���� = b.���� And a.��� = b.ҳ�� And c.ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", lngTMP)
        If Not rsTemp.EOF Then
            objEPRFileInfo.��ʽ = zlCommFun.NVL(rsTemp("��ʽ").Value)
            objEPRFileInfo.SetFormat edtThis, objEPRFileInfo.��ʽ
            edtThis.ResetWYSIWYG
        End If
                    
        If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile
        If gobjFSO.FileExists(strZipFile) Then gobjFSO.DeleteFile strZipFile
                
        '����λ����ʼλ��
        edtThis.Range(1, 1).Selected
        
        edtThis.ForceEdit = False
        edtBuff.ForceEdit = False
        
        ShowSharePageHistory = True
    End If
    
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ReadOnly = True
    edtThis.ReadOnly = True
    
    Set objEPRFileInfo = Nothing
    
    If Trim(strTmpClipboard) <> "" Then '�ָ�ճ��������
        DoEvents
        Clipboard.SetText strTmpClipboard
    Else
        DoEvents
        Clipboard.Clear
    End If
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    Set objEPRFileInfo = Nothing
End Function

'################################################################################################################
'## ���ܣ�  ��ȡϵͳĬ����ʱ·��
'################################################################################################################
Private Function GetSysTmpPath() As String
    GetSysTmpPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
End Function

'################################################################################################################
'## ���ܣ�  �Ƿ���Գ����༭
'################################################################################################################
Public Function CanUndo() As Boolean
    CanUndo = (p_Undo > 1) Or (p_Undo = 1 And UndoList(1).Filename <> "" And UBound(UndoList) = 1)
End Function

'################################################################################################################
'## ���ܣ�  �Ƿ���������༭
'################################################################################################################
Public Function CanRedo() As Boolean
    CanRedo = (p_Undo > 0) And (p_Undo < UBound(UndoList))
End Function

'################################################################################################################
'## ���ܣ�  �����༭һ��
'################################################################################################################
Public Sub Undo()
    If CanUndo = False Then Exit Sub
    If mblnChange Then AddUndoAction
    p_Undo = p_Undo - 1
    Me.Editor1.Tag = "Undo"
    If p_Undo = 0 Then p_Undo = 1
    Me.Document.ImportFromXMLFile Me.Editor1, UndoList(p_Undo).Filename, False, True
    Me.Editor1.Range(UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd).Selected
    Me.Editor1.Tag = ""
    mblnChange = False
    DT1 = Now

    On Error Resume Next
    'ˢ������б�
    RefCompends
    Editor1_SelChange Me.Editor1.ViewMode, UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd   '�ֹ�ˢ��ʾ���ʾ�
End Sub

'################################################################################################################
'## ���ܣ�  �����༭һ��
'################################################################################################################
Public Sub Redo()
    If CanRedo = False Then Exit Sub
    p_Undo = p_Undo + 1
    Me.Editor1.Tag = "Redo"
    Me.Document.ImportFromXMLFile Me.Editor1, UndoList(p_Undo).Filename, False, True
    Me.Editor1.Range(UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd).Selected
    mblnChange = False
    Me.Editor1.Tag = ""
    DT1 = Now

    On Error Resume Next
    'ˢ������б�
    RefCompends
    Editor1_SelChange Me.Editor1.ViewMode, UndoList(p_Undo).SelStart, UndoList(p_Undo).SelEnd   '�ֹ�ˢ��ʾ���ʾ�
End Sub

'################################################################################################################
'## ���ܣ�  �������Undo������ʱ�ļ�
'################################################################################################################
Public Sub ClearUndoList()
    Dim i As Long
    For i = 1 To UBound(UndoList)
        If gobjFSO.FileExists(UndoList(i).Filename) Then gobjFSO.DeleteFile UndoList(i).Filename, True
    Next
    ReDim UndoList(1 To 1) As UndoInfo
    p_Undo = 0
    DT1 = Now
End Sub

'################################################################################################################
'## ���ܣ�  ������õ�Undo�б�����Undo������������������ʱ�����Undo������б�
'################################################################################################################
Private Sub ClearNoUseUndoList()
    If mblnAutosave Then
        If p_Undo < 1 Then Exit Sub
        Dim i As Long
        If p_Undo < UBound(UndoList) Then
            For i = p_Undo + 1 To UBound(UndoList)
               '����ļ�
               If gobjFSO.FileExists(UndoList(i).Filename) Then gobjFSO.DeleteFile UndoList(i).Filename, True
            Next
            ReDim Preserve UndoList(1 To p_Undo) As UndoInfo
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ����һ����ʷ�ļ�
'################################################################################################################
Public Sub AddUndoAction()
    If Me.Document Is Nothing Then Exit Sub
    If mblnAutosave = False Then Exit Sub

    Dim i As Long, j As Long, k As Long, strF As String
    '��ȡһ������ļ���
    Do
        '���У��������ļ����а�����ǰ���λ�õķ�ʽ
        k = Val(gfrmPublic.Tag)
        strF = GetSysTmpPath & "\EPRUndo_" & App.ThreadID & "_" & k & "_" & CLng(Rnd(Timer) * 1000) & ".xml"

        j = j + 1
        If j = 100 Then
            MsgBox "������ʱ�ļ�ʱ�����޷�������ʱ�ļ�", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
        k = k + 1
        gfrmPublic.Tag = k
    Loop While gobjFSO.FileExists(strF)

    ClearNoUseUndoList
    If Me.Document.ExportToXMLFile(Me.Editor1, strF) Then
        If UBound(UndoList) = mlngUndoLimit + 1 Then
            '�Ѿ�����洢����

            If gobjFSO.FileExists(UndoList(1).Filename) Then gobjFSO.DeleteFile UndoList(1).Filename    '�����һ���ļ�
            For i = 1 To UBound(UndoList) - 1
                UndoList(i).Filename = UndoList(i + 1).Filename
                UndoList(i).SelStart = UndoList(i + 1).SelStart
                UndoList(i).SelEnd = UndoList(i + 1).SelEnd
            Next

            p_Undo = mlngUndoLimit + 1
        Else
            p_Undo = p_Undo + 1
            ReDim Preserve UndoList(1 To p_Undo) As UndoInfo
        End If
        UndoList(p_Undo).Filename = strF
        UndoList(p_Undo).SelStart = Me.Editor1.Selection.StartPos
        UndoList(p_Undo).SelEnd = Me.Editor1.Selection.EndPos
        mblnChange = False
        DT1 = Now
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ˢ�µ�ǰ�༭�ĵ��ı���ͼ�����ڲɼ�վ�ڲɼ��µ�ͼ��֮�����
'################################################################################################################
Public Sub RefPacsPic()
    mfrmPacsPic.zlRefresh Document.EPRPatiRecInfo.ҽ��id, Document.EPRFileInfo.lngModule
    mfrmPacsPic.Tag = "Loaded"
End Sub

'################################################################################################################
'## ���ܣ�  �����ǰ�༭�ĵ��ı���ͼ��ˢ�±�־���������ĵ�����ʱ����ˢ��
'################################################################################################################
Public Sub ClsPacsPic()
    mfrmPacsPic.Tag = ""
End Sub

'################################################################################################################
'## ���ܣ�  ���¼����ĵ�ҳ��
'## ������  blnEstopNote-��ֹ��ҳ���ѣ������������ļ�ʱ����Ӧ��ֹ
'################################################################################################################
Public Sub RecountPage(Optional blnEstopNote As Boolean)
    Dim lngPageCount As Long
    
    If mblnAutoPageCount = False Then Exit Sub
    
    If Me.Visible = False Then Exit Sub                 '���ɼ�ʱ������
    If Me.Editor1.ReadOnly Then Exit Sub                'ֻ��ʱ������
    If Me.Editor1.ViewMode <> cprNormal Then Exit Sub   '����ͨģʽ������
    If Me.Editor1.Tag <> "" Then Exit Sub               '������������в�����
    If Me.Editor1.InProcessing Then Exit Sub            'ͬ����ʾ�ڴ��������
    
    lngPageCount = Me.Editor1.PageCount
    Call Me.Editor1.DoVirtualPrint
    stbThis.Panels(2).Text = Editor1.CurrentLine & " ��,  " & Editor1.CurrentColumn & " ��,  ��" & Editor1.LineCount & " ��,  �� " & Me.Editor1.PageCount & " ҳ"
    
    If blnEstopNote = True Then Exit Sub
    If mblnAutoPageNote = False Then Exit Sub
    If Me.Editor1.PageCount - lngPageCount <= 0 Then Exit Sub
    MsgBox "����ҳ�����ѣ��ĵ����ݴ�" & lngPageCount & "ҳ����Ϊ" & Me.Editor1.PageCount & "ҳ��", vbInformation, gstrSysName
    
End Sub
Private Function CanSetFormat() As Boolean
'���ܣ����������ģʽ���ж�ѡ�������Ƿ��ǵ�ǰ�汾��ֻҪѡ����������һ���ֲ��ǵ�ǰ�汾�����ļ�����False
Dim lStar As Long, lEnd As Long, i As Long, COLOR As OLE_COLOR
    With Editor1
        If .SelLength = 0 Then CanSetFormat = True: Exit Function
        lEnd = .Selection.EndPos: lStar = .Selection.StartPos
        For i = lStar To lEnd
            COLOR = .Range(i, i + 1).Font.ForeColor
            If Not Me.Document.IsNewCharColor(COLOR) Then Exit Function
            i = i + 1
        Next
    End With
    CanSetFormat = True
End Function

Private Sub Editor1_BeforeKeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Me.Editor1.ReadOnly Then Exit Sub
'    Debug.Print KeyCode, Shift
    Select Case KeyCode
    Case 0, 16, 17, 18, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyEscape, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital

    Case Else
        If UBound(UndoList) = 1 And p_Undo = 0 And Me.Editor1.Tag = "" Then
            '�״α���
            AddUndoAction
        End If
    End Select
End Sub

Private Sub Editor1_Change(ViewMode As zlRichEditor.ViewModeEnum)
    If Me.Editor1.ReadOnly Then Exit Sub
    mblnChange = True
    If Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos + 1) = Chr(32) Then
        Dim blnForce As Boolean
        If Me.Editor1.Tag <> "" Then Exit Sub       '�Ѿ���������������У���Ӧȥ���ո񣻷����´ʾ����Ŀո�ȥ������������
        blnForce = Me.Editor1.ForceEdit
        Me.Editor1.Tag = "Change"
        Me.Editor1.ForceEdit = True
        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos + 1) = ""
        Me.Editor1.ForceEdit = blnForce
        Me.Editor1.Tag = ""
    End If
    If (DateDiff("s", DT1, DT2) > mlngSaveInterval And Me.Editor1.Tag = "" And Me.Editor1.ForceEdit = False) Or (UBound(UndoList) = 1 And p_Undo = 0 And Me.Editor1.Tag = "" And Me.Editor1.ForceEdit = False) Then
        '����
        AddUndoAction
    ElseIf Me.Editor1.Tag = "" Then
        ClearNoUseUndoList
    End If
    
    Call RecountPage
End Sub

Private Sub Editor1_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Editor1_Resize(ViewMode As zlRichEditor.ViewModeEnum)
    If Me.ucPictureEditor1.Inited = False Then Exit Sub
    If Editor1.UIVisibled Then Editor1.ShowUIInterface
End Sub

Private Sub Editor1_UIClick(ViewMode As zlRichEditor.ViewModeEnum)
    If tblThis.Visible And tblThis.Enabled Then tblThis.SetFocus
End Sub

Private Sub Editor1_UIClose(UIhWnd As Long)
    If ucPacsImgCanvas1.Visible Then
        Dim lKey As Long
        lKey = Val(ucPacsImgCanvas1.Tag)
        
        If ucPictureEditor1.Visible Then
            ucPictureEditor1.Visible = False
            ucPictureEditor1.CloseMe ucPacsImgCanvas1.mMarkedPicture
            ucPacsImgCanvas1.LayoutPictures False
        End If
        
        ucPacsImgCanvas1.SavePictures
        If lKey > 0 Then Document.Tables("K" & lKey).Refresh Editor1
        ucPacsImgCanvas1.CloseMe
        If DkpThis.FindPane(ID_VIEW_SEGMENT).Closed = False Then DkpThis.ShowPane ID_VIEW_SEGMENT
        If DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Closed = False Then DkpThis.ShowPane ID_VIEW_PHRASEDEMO
    Else
        If ucPictureEditor1.Visible Then
            ucPictureEditor1.Visible = False
            ucPictureEditor1.CloseMe
        End If
        
        If tblThis.Visible Then
            Dim lStart As Long, lEnd As Long
            lStart = Editor1.Selection.StartPos
            lEnd = Editor1.Selection.EndPos
            
            If Val(tblThis.Tag) <= 0 Then Exit Sub
            If tblThis.Modified And Me.Document.Tables.Count > 0 Then
                SaveUIToTable Me.Document.Tables("K" & tblThis.Tag), False
            End If
    '        Editor1.Range(lStart, lEnd).Selected
            
            tblThis.Visible = False
            mbEditInTable = False
            tblThis.SelectedCellKey = 0
            tblThis.Tag = ""
        End If
    End If
End Sub
Private Sub Editor1_UIOpen(UIhWnd As Long, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long)
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenAnyKeys(Me.Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys And sKeyType = "T" Then
        If Me.Document.Tables("K" & lKey).TableType = tte_����ͼƬ�� Then
            ucPacsImgCanvas1.ShowMe Me, UIhWnd, cbrThis, Me.Document.Tables("K" & lKey), lngLeft, lngTop, lngWidth, lngHeight
            ucPacsImgCanvas1.Tag = lKey
        Else
            If Val(tblThis.Tag) <= 0 Then Exit Sub
            ReadTableToUI Me.Document.Tables("K" & tblThis.Tag)
            lngWidth = tblThis.Width + 2 * lngLeft
            SetParent tblThis.hwnd, UIhWnd
            tblThis.Move lngLeft, lngTop
            tblThis.hWndBound = Editor1.hWndRTB
            tblThis.OffsetX = tblThis.Left + Editor1.UILeft - 390
            tblThis.OffsetY = tblThis.Top + Editor1.UITop
            tblThis.Visible = True
            mbEditInTable = True
        End If
    ElseIf bBeteenKeys And sKeyType = "P" Then
        If Me.Document.Pictures("K" & lKey).PictureType = EPRSignPicture Or Me.Document.Pictures("K" & lKey).PictureType = EPRPatiSign Then
            Me.Editor1.CloseUIInterface
            Exit Sub
        Else
            ucPictureEditor1.ShowMe Me, UIhWnd, cbrThis, Me.Document.Pictures("K" & lKey), lngLeft, lngTop, lngWidth, lngHeight, False
        End If
    End If
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, y As Single)
    Dim Popup As CommandBar
    Dim Control As CommandBarControl

    Set Popup = cbrThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000

    cbrThis.RecalcLayout
End Sub


Private Sub mfrmDocksymbol_GetPosFontSize()
    Dim lngSize As Long
    On Error Resume Next
    lngSize = Editor1.Selection.Font.Size
    mfrmDocksymbol.PicFontSize = lngSize
End Sub
'���뷶��
Private Sub mfrmDocksymbol_InsertEPRDemo(lngEPRDemoID As Long)
    If Editor1.ReadOnly Then Exit Sub
    If lngEPRDemoID > 0 Then
                Call AddUndoPoint  '�ֶ�����
                Me.Document.ImportEPRDemo Me.Editor1, lngEPRDemoID
                Call ClearNoUseUndoList
                Call RecountPage(True)
            End If
End Sub

Private Sub mfrmDocksymbol_InsertPicSymbol(strInfor As String, picSy As StdPicture, strReturn As String)
    Dim blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    
    On Error GoTo ErrHandle
    If Editor1.ReadOnly Then Exit Sub
    If tblThis.Visible Then
        If strReturn = "" Then
            MsgBox "��ǰλ�ò�֧�ִ����������", vbInformation, gstrSysName: Exit Sub
        Else
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = strReturn
            tblThis.Modified = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Else
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False Then
            Call AddUndoPoint  '�ֶ�����
            Editor1.Tag = "InsertPicSymbol"
            InsertPicture EPRFormulaPicture, picSy, picSy.Width, picSy.Height, strInfor
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
        Editor1.SetFocus
    End If
    Call RecountPage
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mfrmDocksymbol_InsertSymbol(strSymbol As String, intStrLen As Integer)
'intStrLen �� strSymbol �����ַ����ĳ��ȣ������������չ��λ�ã�����ҽѧ��λ������ҩ���������Ϊʵ�ʳ��ȣ�����Ĭ��Ϊ1
    Dim blnForce As Boolean, strFont As String, lngSelStart As Long, lngSymbolPos As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    On Error GoTo ErrHandle
    If Editor1.ReadOnly Then Exit Sub
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = tblThis.Cells("K" & tblThis.SelectedCellKey).Text & strSymbol
                tblThis.Modified = True
                tblThis.Refresh True, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        End If
        tblThis.SetFocus
    Else
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
            Call AddUndoPoint  '�ֶ�����
            blnForce = Editor1.ForceEdit
            Editor1.ForceEdit = True
            Call Editor1_KeyDown(cprNormal, 32, 0)
            Editor1.Tag = "InsertstrSymbol"
            If Me.Editor1.AuditMode Then
                Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
                '���������ԣ����������ı���
                On Error Resume Next
                Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
                Me.Editor1.OriginRTB.SelStrikeThru = False
            End If
            strFont = Editor1.Selection.Font.Name
            lngSelStart = Editor1.SelStart
            Editor1.SelText = strSymbol
            lngSymbolPos = lngSelStart + intStrLen
            Editor1.SelStart = lngSymbolPos
            Editor1.Range(lngSelStart, lngSymbolPos).Font.Name = strFont 'Toshma������ڱ��桢ǩ��ʱ������Ϊ��UTF�ַ�ռλ3���ֽ�
            If intStrLen = 1 And Editor1.Range(lngSelStart, lngSymbolPos).Font.Name = "Tahoma" Then
                Editor1.Range(lngSelStart, lngSymbolPos).Font.Name = "����"
            End If
            Editor1.Selection.Font.Name = strFont
            
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
        Editor1.SetFocus
    End If
    Call RecountPage
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mfrmDocksymbol_SetFouse()
    Me.Editor1.Enabled = False
    Me.Editor1.Enabled = True
End Sub

Private Sub mfrmHistoryReport_CopyClick(ByVal strContent As String)
    Clipboard.SetText strContent
    On Error Resume Next
    If cbrThis.FindControl(, ID_EDIT_PASTE).Enabled Then
        cbrThis.FindControl(, ID_EDIT_PASTE).Execute
    End If
End Sub

Private Sub mfrmHistoryReport_ReportCountChange(ByVal lngReportCount As Long)
On Error Resume Next
    DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Title = "��ʷ���(" & lngReportCount & ")"
End Sub

Private Sub mfrmMainError_Location(ByVal Key As Long)

    Call Me.Document.Elements("K" & Key).Selected(Me.Editor1)

End Sub

Private Sub mfrmPacsPic_InsertPicture(pic As stdole.StdPicture, ByVal strUid As String, ByVal lngAdviceID As Long)
    If Editor1.ReadOnly Then Exit Sub
    If ucPacsImgCanvas1.Visible = False Then Exit Sub
    ucPacsImgCanvas1.AddPacsPicture pic, strUid, lngAdviceID
    Call RecountPage
End Sub

Private Sub mfrmPreview_PrintEpr(ByVal lngRecordId As Long)
    Me.Document.AfterPrinted Me.Document.EPRPatiRecInfo.ID
End Sub

Private Sub mfrmSegments_ModifiedOrDeleted(Action As Integer)
    If Me.Document.EditType = cprET_ȫ��ʾ���༭ Then
        Err = 0: On Error Resume Next
        Call Me.Document.mfrmParent.RefreshList
    End If
End Sub

Private Sub mfrmSegments_RowDblClick(ByVal Row As XtremeReportControl.IReportRow)
    Dim rsTemp As New ADODB.Recordset, lngDemoId As Long
    Dim rsText As New ADODB.Recordset, strVSql As String
    Dim oCompend As cEPRCompend, lngStart As Long, lngTail As Long
    Dim oCell As cEPRCell, oElement As cEPRElement, oPicture As cEPRPicture
    Dim StrText As String, lngLen As Long, lngKey As Long, aryProp() As String
    
    If Me.Editor1.ViewMode <> cprNormal Or Me.Editor1.ReadOnly Then Exit Sub
    If Row Is Nothing Then Exit Sub
    lngDemoId = Row.Record(1).Value

    gstrSQL = "Select Id, �������id From ������������ Where �ļ�id = [1] And �������� = 1 Order By �������"
    strVSql = "Select Id, ��������, ��������, �����ı�, �Ƿ���, Ҫ������, ����Ҫ��id, �滻��, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ," & vbNewLine & _
            "       Ҫ�ر�ʾ, Ҫ��ֵ��, ������̬" & vbNewLine & _
            "From ������������" & vbNewLine & _
            "Where �ļ�id = [1] And ��id = [2]" & vbNewLine & _
            "Order By �������"
    
    Me.Editor1.ForceEdit = True
    Me.Editor1.Tag = "mfrmSegments_RowDblClick"
    Me.Editor1.Freeze
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", lngDemoId)
    Do While Not rsTemp.EOF
        lngStart = 0: lngTail = 0
        For Each oCompend In Me.Document.Compends
            If oCompend.�������ID > 0 And oCompend.�������ID = Val("" & rsTemp!�������ID) Then
                oCompend.GetPosition Me.Editor1, lngStart, lngTail
                Exit For
            End If
        Next
        If lngTail > 0 Then
            With Me.Editor1
                If .Range(lngTail - 2, lngTail) <> vbCrLf Then
                    .Range(lngTail, lngTail) = vbCrLf
                    With .Range(lngTail, lngTail + 2).Font
                        .Protected = False: .Hidden = False: .Strikethrough = False: .BackColor = tomAutoColor
                        .ForeColor = IIf(Me.Editor1.AuditMode, GetCharColor(Me.Document.Ŀ��汾, 0), tomAutoColor)
                    End With
                End If
                .SelStart = lngTail
            End With
            Set rsText = zlDatabase.OpenSQLRecord(strVSql, "��ȡ��Ϣ", lngDemoId, CLng(rsTemp!ID))
            Do While Not rsText.EOF
                lngTail = Me.Editor1.SelStart
                Select Case rsText!��������
                Case 2  '�ı�
                    StrText = rsText!�����ı� & IIf(Val("" & rsText!�Ƿ���) = 1, vbCrLf, "")
                    lngLen = Len(StrText)
                    With Me.Editor1
                        .Range(lngTail, lngTail) = StrText
                        With .Range(lngTail, lngTail + lngLen).Font
                            .Protected = False: .Hidden = False: .Strikethrough = False: .BackColor = tomAutoColor
                            .ForeColor = IIf(Me.Editor1.AuditMode, GetCharColor(Me.Document.Ŀ��汾, 0), tomAutoColor)
                        End With
                        .Range(lngTail + lngLen, lngTail + lngLen).Selected
                    End With
                Case 3  '���
                    If Me.Document.EditType = cprET_ȫ��ʾ���༭ Or Me.Document.EditType = cprET_�������༭ Then
                        lngKey = Me.Document.Tables.Add
                        With Me.Document.Tables("K" & lngKey)
                            Call .GetTableFromDB(cprET_ȫ��ʾ���༭, lngDemoId, rsText!ID, False)
                            .ID = 0: .�ļ�ID = 0: .��ID = 0: .��ʼ�� = Me.Document.Ŀ��汾
                            For Each oCell In .Cells
                                oCell.ID = 0: oCell.�ļ�ID = 0: oCell.��ID = 0: oCell.��ʼ�� = Me.Document.Ŀ��汾
                            Next
                            For Each oElement In .Elements
                                oElement.ID = 0: oElement.�ļ�ID = 0: oElement.��ID = 0: oElement.��ʼ�� = Me.Document.Ŀ��汾
                                If oElement.�滻�� = 1 And Me.Document.EditType = cprET_�������༭ Then
                                    oElement.�����ı� = GetReplaceEleValue(oElement.Ҫ������, _
                                        Me.Document.EPRPatiRecInfo.����ID, _
                                        Me.Document.EPRPatiRecInfo.��ҳID, _
                                        Me.Document.EPRPatiRecInfo.������Դ, _
                                        Me.Document.EPRPatiRecInfo.ҽ��id, _
                                        Me.Document.EPRPatiRecInfo.Ӥ��)
                                        For Each oCell In .Cells
                                            If oCell.ElementKey = oElement.Key Then oCell.�����ı� = oElement.�����ı�: Exit For
                                        Next
                                End If
                            Next
                            For Each oPicture In .Pictures
                                oPicture.ID = 0: oPicture.�ļ�ID = 0: oPicture.��ID = 0: oPicture.��ʼ�� = Me.Document.Ŀ��汾
                            Next
                            .InsertIntoEditor Me.Editor1, lngTail
                        End With
                    End If
                Case 4  'Ԫ��
                    lngKey = Me.Document.Elements.Add
                    With Me.Document.Elements("K" & lngKey)
                        .ID = 0
                        .�����ı� = "" & rsText!�����ı�
                        .Ҫ������ = "" & rsText!Ҫ������
                        .����Ҫ��ID = Val("" & rsText!����Ҫ��ID)
                        .�滻�� = Val("" & rsText!�滻��)
                        .Ҫ������ = Val("" & rsText!Ҫ������)
                        .Ҫ�س��� = Val("" & rsText!Ҫ�س���)
                        .Ҫ��С�� = Val("" & rsText!Ҫ��С��)
                        .Ҫ�ص�λ = "" & rsText!Ҫ�ص�λ
                        .Ҫ�ر�ʾ = Val("" & rsText!Ҫ�ر�ʾ)
                        .Ҫ��ֵ�� = "" & rsText!Ҫ��ֵ��
                        .������̬ = Val("" & rsText!������̬)
                        .�Ƿ��� = Val("" & rsText!�Ƿ���)
                        If .�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                            .�����ı� = GetReplaceEleValue(.Ҫ������, _
                                Me.Document.EPRPatiRecInfo.����ID, _
                                Me.Document.EPRPatiRecInfo.��ҳID, _
                                Me.Document.EPRPatiRecInfo.������Դ, _
                                Me.Document.EPRPatiRecInfo.ҽ��id, _
                                Me.Document.EPRPatiRecInfo.Ӥ��)
                        End If
                        .��ʼ�� = Me.Document.Ŀ��汾
                        .InsertIntoEditor Me.Editor1, lngTail, , True
                    End With
                Case 5  'ͼ��
                    If Me.Document.EditType = cprET_ȫ��ʾ���༭ Or Me.Document.EditType = cprET_�������༭ Then
                        lngKey = Me.Document.Pictures.Add
                        With Me.Document.Pictures("K" & lngKey)
                            Call .GetPictureFromDB(cprET_ȫ��ʾ���༭, lngDemoId, rsText!ID, False)
                            .ID = 0: .�ļ�ID = 0: .��ID = 0
                            .InsertIntoEditor Me.Editor1, lngTail, True
                        End With
                    End If
                Case 7  '���
                    aryProp = Split("" & rsText!��������, ";")
                    lngKey = Me.Document.Diagnosises.Add
                    With Me.Document.Diagnosises("K" & lngKey)
                        .ID = 0
                        .���� = "" & rsText!�����ı�
                        .���� = Val(aryProp(0))
                        .��ҽ = Val(aryProp(1))
                        .����id = Val(aryProp(2))
                        .���id = Val(aryProp(3))
                        .֤��id = Val(aryProp(4))
                        .���� = Val(aryProp(5))
                        .���� = Format(Now(), "yyyy-mm-dd hh:mm:ss")
                        .��ʼ�� = Me.Document.Ŀ��汾
                        .InsertIntoEditor Me.Editor1, lngTail, True
                    End With
                End Select
                rsText.MoveNext
            Loop
        End If
        rsTemp.MoveNext
    Loop
    Me.Editor1.ForceEdit = False
    Me.Editor1.Tag = ""
    Me.Editor1.UnFreeze
    Call RecountPage
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.Editor1.ForceEdit = False
End Sub

Private Sub mfrmSentenceDetailed_ShiftFocus()
    Me.Editor1.Enabled = False
    Me.Editor1.Enabled = True
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Me.Document.AfterPrinted Me.Document.EPRPatiRecInfo.ID
End Sub

Private Sub picHistoryInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If X > 0 And X < picHistoryInfo.ScaleWidth And y > 0 And y < picHistoryInfo.ScaleHeight Then
        If picHistoryInfo.Tag = "" Then
            SetCapture picHistoryInfo.hwnd
            picHistoryInfo.Cls
            picHistoryInfo.BackColor = &HD2BDB6    ' &HD8D5D4 ' &HD2BDB6
            picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H6A240A, B
            picHistoryInfo.Tag = "Captured"
        End If
    Else
        ReleaseCapture
        picHistoryInfo.Cls
        picHistoryInfo.BackColor = &H8000000F
        picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
        picHistoryInfo.Tag = ""
    End If
End Sub

Private Sub picHistoryInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    ReleaseCapture
    picHistoryInfo.Cls
    picHistoryInfo.BackColor = &H8000000F
    picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picHistoryInfo.Tag = ""
End Sub

Private Sub picHistoryInfo_Resize()
    picHistoryInfo.Cls
    picHistoryInfo.BackColor = &H8000000F
    picHistoryInfo.Line (0, 0)-(picHistoryInfo.ScaleWidth - Screen.TwipsPerPixelX, picHistoryInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picHistoryInfo.Tag = ""
End Sub

Private Sub picPane_Resize()
    On Error Resume Next
    
    picHistoryInfo.Move 15, 15, picPane.Width - 30
    edtThis.Move 15, picHistoryInfo.Top + picHistoryInfo.Height, picPane.Width - 30, picPane.Height - 15 - (picHistoryInfo.Top + picHistoryInfo.Height)
End Sub

Private Sub tblThis_CancelEdit()
    Editor1.Modified = True
End Sub

Private Sub tblThis_SelectionChange(ByVal lrow As Long, ByVal lCol As Long)
    Dim lngKey As Long
    If Me.Editor1.AuditMode = True Then Exit Sub
    lngKey = Val(tblThis.Cell(lrow, lCol).Tag)
    If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
    If lngKey > 0 Then
        If Not tblThis.Cell(lrow, lCol).Picture Is Nothing Then
            'ͼƬ
            If Val(tblThis.Tag) > 0 Then
                '�༭ͼƬ
                Dim LL As Long, lT As Long, lW As Long, lH As Long
                tblThis.Cell(lrow, lCol).GetCellPictureBorder LL, lT, lW, lH
                
                    ucPictureEditor1.ShowMe Me, tblThis.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey), _
                        LL, lT, lW, lH, True, Me.Document.Tables("K" & tblThis.Tag)
                
                mblnChange = True
                tblThis.Modified = True
            Else
                If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
            End If
        Else
            If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
        End If
    Else
        If ucPictureEditor1.Visible Then ucPictureEditor1.CloseMe: tblThis.Modified = True
    End If
End Sub

Private Sub tblThis_ModifyProtected(ByVal lKey As Long)
    Dim lLeft As Long, lTOp As Long, lRight As Long, lBottom As Long, lngKey As Long

    tblThis.Cells("K" & lKey).GetCellTextBorder lLeft, lTOp, lRight, lBottom

    lngKey = Val(tblThis.Cells("K" & lKey).Tag)
    If lngKey > 0 Then
        If tblThis.Cells("K" & lKey).Picture Is Nothing Then
            '����Ҫ��
            If Val(tblThis.Tag) > 0 Then
                If Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).�滻�� = 2 Then
                    '�ֵ���Ŀ
                    mfrmDicSelect.Tag = lngKey
                    mfrmDicSelect.ShowMe Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).Ҫ������, Me.Left + Editor1.Left + Editor1.UILeft + tblThis.Left + lLeft * 15 + 30, _
                        Me.Top + Editor1.Top + Editor1.UITop + tblThis.Top + lBottom * 15 + 500, IIf(mintStyle = -1, 0, mintStyle), Me, Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).�����ı�
                Else
                    '����Ҫ��
                    mfrmModElement.Tag = lngKey
                    mfrmModElement.ShowMe Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey), _
                        Me.Left + Editor1.Left + Editor1.UILeft + tblThis.Left + lLeft * 15 + 30, _
                        Me.Top + Editor1.Top + Editor1.UITop + tblThis.Top + lBottom * 15 + 500, IIf(mintStyle = -1, 0, mintStyle), Me, Me.Document.EditType
                End If
            End If
        Else
            'ͼƬ
            If Me.Editor1.AuditMode = True Then Exit Sub
            If Val(tblThis.Tag) > 0 Then
                '�༭ָ����ͼƬ
                If Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).PictureType = EPRMarkedPicture Then
                    '�༭ͼƬ
                    Dim LL As Long, lT As Long, lW As Long, lH As Long
                    tblThis.Cells("K" & lKey).GetCellPictureBorder LL, lT, lW, lH
                    ucPictureEditor1.ShowMe Me, tblThis.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey), _
                        LL, lT, lW, lH, True, Me.Document.Tables("K" & tblThis.Tag)
                ElseIf Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).PictureType = EPROutPicture Then
                    cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).OrigPic, _
                        lngKey, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).��������, Me, False
                    '�����ⲿͼƬ�ı�����cPicEditor�����pOK�¼��д���
                End If
            End If
        End If
        mblnChange = True
    End If
End Sub

Private Sub tblThis_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Dim Popup As CommandBar
        Dim cbpPopup As CommandBarPopup
        Dim Control As CommandBarControl
        Dim lngKey As Long

        Set Popup = cbrThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "����(&X)")
            Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
            Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "ճ��(&V)")
            Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "ɾ��(&D)")
            Set Control = .Add(xtpControlButton, ID_TABLE_MERGE, "�ϲ���Ԫ��(&M)"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, ID_TABLE_DELETETABLE, "ɾ�����(&T)")

            If tblThis.SelStartRow = tblThis.SelEndRow And tblThis.SelStartCol = tblThis.SelEndCol Then
                If Val(tblThis.Tag) > 0 And tblThis.SelectedCellKey > 0 Then
                    '�༭ָ����ͼƬ
                    If Not tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing Then
                        lngKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                        Set Control = .Add(xtpControlButton, ID_EDIT_MARKEDPIC, "����޸�(&M)"): Control.BeginGroup = True
                        If Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lngKey).PictureType = EPROutPicture Then
                            .Add xtpControlButton, ID_EDIT_OUTERPIC, "��ͼ����(&D)"
                        End If
                    End If
                End If
            End If

            Set cbpPopup = .Add(xtpControlPopup, ID_TABLE_CELLALIGNMENT, "��Ԫ����뷽ʽ")
            cbpPopup.CommandBar.SetTearOffPopup "��Ԫ����뷽ʽ", ID_TABLE_CELLALIGNMENT, 100
            cbpPopup.CommandBar.SetPopupToolBar True
            cbpPopup.BeginGroup = True
            cbpPopup.CommandBar.Width = 70
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "���������"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "���Ͼ���"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "�����Ҷ���"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "�в������"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "�в�����"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "�в��Ҷ���"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "���������"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "���¾���"
            cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "�����Ҷ���"

            Set Control = .Add(xtpControlButton, ID_TABLE_PROPERTY, "�������(&R)..."): Control.BeginGroup = True

            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub tblThis_Resize(ByVal lWidth As Long, ByVal lHeight As Long)
    '�༭�����ж�̬�ı����С
    Editor1.ResizeUIInterface lWidth, lHeight
    If ucPictureEditor1.Visible Then ucPictureEditor1.Visible = False
    If Val(tblThis.Tag) > 0 Then
        Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, bFinded As Boolean, bNeeded As Boolean
        Dim lW As Long
        bFinded = FindKey(Editor1, "T", tblThis.Tag, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then
            Editor1.InProcessing = True
            lW = Me.Editor1.PaperWidth - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - Me.ScaleX(Me.Editor1.Range(lSE, lES).Para.LeftIndent + Me.Editor1.Range(lSE, lES).Para.FirstLineIndent, vbPixels, vbTwips) - 130
            picTMP.Width = IIf(tblThis.Width > lW, lW, tblThis.Width)
            picTMP.Height = tblThis.Height
            tblThis.DrawToDC picTMP.hDC
            picTMP.Picture = picTMP.Image
            'ˢ�²�ѡ�иı��ͼƬ
            Me.Document.Tables("K" & tblThis.Tag).Refresh Editor1, picTMP.Picture, True
'            Editor1.InProcessing = True
'            Editor1.Range(lSE, lES).Selected
'            Editor1.RefreshUIInterface
            Editor1.InProcessing = False
            mblnChange = True
        End If
    End If
End Sub

Private Sub tmrAutoSaveEPR_Timer()
    DT2_EPR = Now
    If mblnAutoSaveEPR Then
        If DateDiff("n", DT1_EPR, DT2_EPR) > mlngSaveIntervalEPR Then
            '�Զ������ļ�
            If (Editor1.Modified) And (Me.Document.Ŀ��汾 <= 16) Then
                Call SaveEMRDoc(True)
                DT1_EPR = Now
            End If
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ��ʱ�Զ�����
'################################################################################################################
Private Sub tmrThis_Timer()
    If Me.Document Is Nothing Then Exit Sub
    If mblnAutosave Then DT2 = Now
End Sub

'################################################################################################################
'## ���ܣ�  �ֹ����ӻ�ԭ��
'################################################################################################################
Private Sub AddUndoPoint()
    If Me.Document Is Nothing Then Exit Sub
    If mblnChange Then AddUndoAction
End Sub
'################################################################################################################
'## ���ܣ�  ˢ�����
'################################################################################################################
Public Sub RefCompends()
    Document.Compends.UpdateOrdersFromText Editor1
    Document.Compends.FillTree mfrmCompends.Tree
    Call RefSentenceList
End Sub

'################################################################################################################
'## ���ܣ�  ���������ص�ʾ���ʾ�
'################################################################################################################
Public Sub RefSentenceList()
    Dim lngCompend As Long, lngPatient As Long, lngVisit As Long, lngAdvice As Long
    Dim blnForce As Boolean         '�ļ�����ʱ��ǿ��ˢ��
    Dim strLimit As String          '�ɲ���,Ҫ������Ĵʾ�����
    If mfrmCompends.Tree.SelectedItem Is Nothing Then Exit Sub
    
    If Me.Document.EditType = cprET_�����ļ����� Then
        If mfrmCompends.Tree.SelectedItem Is Nothing Then
            lngCompend = 0
        Else
            lngCompend = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).ID
        End If
        lngPatient = 0: lngVisit = 0: lngAdvice = 0
        blnForce = True
    ElseIf Me.Document.EditType = cprET_ȫ��ʾ���༭ Then
        If mfrmCompends.Tree.SelectedItem Is Nothing Then
            lngCompend = 0
        Else
            lngCompend = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).�������ID
        End If
        lngPatient = 0: lngVisit = 0: lngAdvice = 0
        blnForce = False
    Else
        If mfrmCompends.Tree.SelectedItem Is Nothing Then
            lngCompend = 0
        Else
            lngCompend = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).�������ID
        End If
        lngPatient = Me.Document.EPRPatiRecInfo.����ID
        lngVisit = Me.Document.EPRPatiRecInfo.��ҳID
        lngAdvice = Me.Document.EPRPatiRecInfo.ҽ��id
        blnForce = False
    End If
    strLimit = MakeSentenceLimit(lngCompend)
    Call mfrmSentenceDetailed.zlRefFromCompend(Me, lngCompend, lngPatient, lngVisit, lngAdvice, blnForce, strLimit)
End Sub

'################################################################################################################
'## ���ܣ�  ��ʾ���༭������
'##
'## ������  frmParent       :������
'##         blnFirst        :�Ƿ��ǵ�һ�δ򿪴��壨���ļ��༭ʱʹ�ã��л���ǰ�ļ�ʱ��ΪFalse����ʾ��ˢ������ĵ���
'##         blnCanPrint     :�Ƿ�����Ԥ������ӡ
'################################################################################################################
Public Sub ShowMe(frmParent As Object, Optional blnFirst As Boolean = True, Optional blnCanPrint As Boolean = True, Optional ByVal byteStyle As Integer)
    '���ô�����ʾ״̬
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    mblnIsMultiMode = Me.Document.IsMultiEPRDoc
    mblnCanPrint = blnCanPrint
    mblnPrecess = False
    mintStyle = byteStyle
    mblnPatiSign = HavedPatiSign
    Call SetStateInfo
    Select Case Document.EditType
        Case cprET_�������༭
            stbThis.Panels(5).Visible = False
            DkpThis.FindPane(ID_VIEW_STRUCTURE).Close
            CommBar(ID_BAR_SIGN).Visible = True
            Call mfrmSegments.zlRefresh(Me)
        Case cprET_���������
            Editor1.AuditMode = True
            DkpThis.FindPane(ID_VIEW_STRUCTURE).Close
            CommBar(ID_BAR_SIGN).Visible = True
            Call mfrmSegments.zlRefresh(Me)
        Case cprET_�����ļ�����
            stbThis.Panels(5).Visible = False
            DkpThis.FindPane(ID_VIEW_SEGMENT).Close
            CommBar(ID_BAR_SIGN).Visible = False
        Case cprET_ȫ��ʾ���༭
            stbThis.Panels(5).Visible = False
            If Document.EPRDemoInfo.���� <> 0 Then
                CommBar(ID_BAR_FORMAT).Visible = False
                CommBar(ID_BAR_FORMAT).Delete
            End If
            CommBar(ID_BAR_SIGN).Visible = False
            Call mfrmSegments.zlRefresh(Me)
    End Select

    If Document.EditType = cprET_�������༭ Or Document.EditType = cprET_��������� Then
        If gobjESign Is Nothing Then
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
        If Not gobjESign Is Nothing Then
            mblnEnPtSign = gobjESign.EnabledPatiSign
        End If
    End If
        
    If Me.Document.EPRFileInfo.���� = cpr���Ʊ��� Then
        mfrmPacsPic.zlRefresh Document.EPRPatiRecInfo.ҽ��id, Document.EPRFileInfo.lngModule
        DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Title = "��ʷ���(" & mfrmHistoryReport.zlRefresh(Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.ID) & ")"
    Else
        DkpThis.FindPane(ID_VIEW_PACSPIC).Close
        DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Close
    End If

    If mblnIsMultiMode And blnFirst Then
        If mfrmMultiDocView Is Nothing Then Set mfrmMultiDocView = New frmMultiDocView
        '���ļ���ϲ��Ĵ���ĳ�ʼ��
        mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
        DkpThis.ShowPane ID_VIEW_MULTIDOCVIEW
    End If
    
    If mblnIsMultiMode Then
        mblnExistHistroy = ShowSharePageHistory(Me.Document, mintSharePages)
    End If
    If mblnExistHistroy = False Then picPane.Visible = mblnExistHistroy
    
    If mblnChildMode Then
        stbThis.Visible = False
        picPatiInfo.Visible = False
    Else
        stbThis.Visible = True
        picPatiInfo.Visible = True
    End If
    Call Me.Editor1.DoVirtualPrint
    stbThis.Panels(2).Text = Editor1.CurrentLine & " ��,  " & Editor1.CurrentColumn & " ��,  ��" & Editor1.LineCount & " ��"
    If mblnAutoPageCount Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & ",  �� " & Me.Editor1.PageCount & " ҳ"
    
    Dim intLoop As Integer
    
    For intLoop = 1 To Me.Document.Elements.Count
        If Me.Document.Elements(intLoop).ǩ��Ҫ�� Then
            mblnǩ��Ҫ�� = True
            Exit For
        End If
    Next
    
    If Me.Document.EditType = cprET_�������༭ Then  '���ݲ��֡�Ҫ����ѡ����δѡҪ�ؿ���ѡ��
        Call ReadElementLimit
        Call CheckLastDiagnose '�������²��ܶԼ��е�Ҫ�ط����䶯
    Else '���������������ά��
        ReDim mEleLimit(0) As EleLimit
    End If
    'Ϊˢ�·����б�������
    Call mfrmDocksymbol.SetItems(Document.EPRPatiRecInfo.�ļ�ID, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.ҽ��id)
    mbln���޴��� = False
    mblnFBContentChanged = True
    If Document.EPRPatiRecInfo.�������� = cpr������� Then
        strSQL = "Select b.�������� From �����걨��¼ A, �������淴�� B" & vbNewLine & _
             "Where a.�ļ�id = b.�ļ�id and A.����״̬=4 And a.�ļ�id = [1] And B.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From �������淴�� Where �ļ�id = [1])"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Me.Document.EPRPatiRecInfo.ID)
        If rs.RecordCount > 0 Then
            mbln���޴��� = True
            txtFeedBack.Text = NVL(rs!��������)
            cbrThis.Item(2).Controls.Find(, 99999901).Visible = True
            cbrThis.Item(2).Controls.Find(, 99999902).Visible = True
            cbrThis.Item(2).Controls.Find(, 99999903).Visible = True
            cbrThis.Item(2).Controls.Find(, 99999904).Visible = True
        End If
    End If
    If mintStyle = vbModeless Or mintStyle = vbModal Then
        Me.Show mintStyle, frmParent
    End If
End Sub
Private Sub ReadElementLimit()
Dim rsTemp As ADODB.Recordset, i As Integer
Dim strLastDiagnose As String
'�ṹ�����0ά����ֵ,�ӵ�һά��ʼ��Ϊ��Ч
    On Error GoTo errHand
    If Not (Document.EPRFileInfo.���� = cpr���ﲡ�� Or Document.EPRFileInfo.���� = cprסԺ����) Then Exit Sub 'ֻ������ﲡ����סԺ����
    
    gstrSQL = "Select a.�䶯ԭ��,a.ԭ��Ҫ��id, a.ԭ��Ҫ��, a.ԭ������, b.�䶯���, b.�������id ������id, b.���Ҫ��id, b.���Ҫ��, b.���ֵ��,b.ԭʼֵ��" & vbNewLine & _
                "From �����䶯ԭ�� A, �����䶯��� B" & vbNewLine & _
                "Where a.�����ļ�id = [1] And b.�䶯ԭ��id = a.Id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�䶯", Document.EPRFileInfo.ID)
    ReDim mEleLimit(rsTemp.RecordCount) As EleLimit
    For i = 1 To rsTemp.RecordCount
        mEleLimit(i).�䶯ԭ�� = rsTemp!�䶯ԭ��
        mEleLimit(i).ԭ��Ҫ��id = NVL(rsTemp!ԭ��Ҫ��id, 0)
        mEleLimit(i).ԭ��Ҫ�� = NVL(rsTemp!ԭ��Ҫ��, "")
        mEleLimit(i).ԭ������ = NVL(rsTemp!ԭ������, "")
        mEleLimit(i).�䶯��� = rsTemp!�䶯���
        mEleLimit(i).������id = NVL(rsTemp!������id, 0)
        mEleLimit(i).���Ҫ��id = NVL(rsTemp!���Ҫ��id, 0)
        mEleLimit(i).���Ҫ�� = NVL(rsTemp!���Ҫ��, "")
        mEleLimit(i).���ֵ�� = NVL(rsTemp!���ֵ��, "")
        mEleLimit(i).ԭʼֵ�� = NVL(rsTemp!ԭʼֵ��, "")
        rsTemp.MoveNext
    Next
    
    strLastDiagnose = GetReplaceEleValue("������ID", Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.������Դ, Document.EPRPatiRecInfo.ҽ��id, Me.Document.EPRPatiRecInfo.Ӥ��)
    mlDiseaseID = Val(Split(strLastDiagnose, "|")(0))
    mlDiagnoseID = Val(Split(strLastDiagnose, "|")(1))
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub CheckLastDiagnose()
'���ܣ��ڲ�����ʱ�Բ�������Ҫ��ѡ����м�������
Dim intLm As Integer, intEl, lKey As Long
    On Error GoTo errHand
    If Not (Document.EPRFileInfo.���� = cpr���ﲡ�� Or Document.EPRFileInfo.���� = cprסԺ����) Then Exit Sub 'ֻ������ﲡ����סԺ����
    
    For intLm = 1 To UBound(mEleLimit)
        Select Case mEleLimit(intLm).�䶯ԭ��
            Case 1 'Ҫ������ı仯���ڴ˴����ڱ༭ʱ�����޸������Ҫ�ص�ֵ����ȷ�����������²��ɴ���
            Case 2, 3 '��������ı仯
                If (mlDiseaseID = mEleLimit(intLm).ԭ��Ҫ��id Or mlDiagnoseID = mEleLimit(intLm).ԭ��Ҫ��id) Then
                    For intEl = 1 To Document.Elements.Count
                        lKey = Document.Elements(intEl).Key
                        If Document.Elements("K" & lKey).Ҫ������ = mEleLimit(intLm).���Ҫ�� And Document.Elements("K" & lKey).����Ҫ��ID = mEleLimit(intLm).���Ҫ��id Then
                            Select Case mEleLimit(intLm).�䶯���
                                Case 1 '����Ҫ��ѡ��仯
                                    If Document.Elements("K" & lKey).������̬ = 1 Then
                                        If InStr(Document.Elements("K" & lKey).�����ı�, "��") = 0 And InStr(Document.Elements("K" & lKey).�����ı�, "��") = 0 Then
                                            '���Ҫ��û��ѡ�ѡ�вŸ���ֵ�����ݣ���ˢ����ʾ
                                            Document.Elements("K" & lKey).Ҫ��ֵ�� = mEleLimit(intLm).���ֵ��
                                            If Document.Elements("K" & lKey).Ҫ�ر�ʾ = 2 Then
                                                Document.Elements("K" & lKey).�����ı� = "��" & Replace(mEleLimit(intLm).���ֵ��, ";", "��")
                                            Else
                                                Document.Elements("K" & lKey).�����ı� = "��" & Replace(mEleLimit(intLm).���ֵ��, ";", "��")
                                            End If
                                            Document.Elements("K" & lKey).Refresh Editor1
                                        End If
                                    Else                                                             '���Ҫ��û��ѡ�ѡ�У�����ֵ��
                                        If Document.Elements("K" & lKey).�����ı� = "" Then Document.Elements("K" & lKey).Ҫ��ֵ�� = mEleLimit(intLm).���ֵ��
                                    End If
                                Case 3 'ɾ��Ҫ��
                                    If Document.Elements("K" & lKey).������̬ = 1 Then
                                        If InStr(Document.Elements("K" & lKey).�����ı�, "��") = 0 And InStr(Document.Elements("K" & lKey).�����ı�, "��") = 0 Then
                                            Document.Elements("K" & lKey).DeleteFromEditor Editor1
                                            Exit For
                                        End If
                                    Else
                                        If Document.Elements("K" & lKey).�����ı� = "" Then Document.Elements("K" & lKey).DeleteFromEditor Editor1: Exit For
                                    End If
                            End Select
                        End If
                    Next
                End If
        End Select
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub UpdateSameELement(ByVal lKey As Long)
'���ܣ���ǰ���ڱ༭״̬����Ҫ��ѡ��ʱ������Ҫ�ؽ���ѡ���������
Dim leKey As Long
Dim strElementName As String, strElementValue As String, lEleId As Long, intLm As Integer, intEl As Integer, intLn As Integer, blnLm As Boolean
    On Error GoTo errHand
    
    If Not (Document.EPRFileInfo.���� = cpr���ﲡ�� Or Document.EPRFileInfo.���� = cprסԺ����) Then Exit Sub 'ֻ������ﲡ����סԺ����
    If Document.EditType = cprET_��������� Then Exit Sub
    
    strElementName = Document.Elements("K" & lKey).Ҫ������
    strElementValue = Document.Elements("K" & lKey).�����ı�
    lEleId = Document.Elements("K" & lKey).����Ҫ��ID
    
    For intLm = 1 To UBound(mEleLimit)
        If mEleLimit(intLm).�䶯ԭ�� = 4 Then
            If strElementName = mEleLimit(intLm).ԭ��Ҫ�� And lEleId = mEleLimit(intLm).ԭ��Ҫ��id Then
                For intEl = 1 To Document.Elements.Count
                    leKey = Document.Elements(intEl).Key
                    If Document.Elements("K" & leKey).Ҫ������ = mEleLimit(intLm).���Ҫ�� And Document.Elements("K" & leKey).����Ҫ��ID = mEleLimit(intLm).���Ҫ��id Then
                        If Document.Elements("K" & leKey).������̬ = 1 Then
                            '���Ҫ��û��ѡ�ѡ�вŸ���ֵ�����ݣ���ˢ����ʾ
                            Document.Elements("K" & leKey).�����ı� = strElementValue
                            Document.Elements("K" & leKey).Refresh Editor1
                        Else
                            '���Ҫ��û��ѡ�ѡ�вŸ���ֵ�����ݣ���ˢ����ʾ
                            Document.Elements("K" & leKey).�����ı� = strElementValue
                            Document.Elements("K" & leKey).Refresh Editor1
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub CheckElementLimit(ByVal lKey As Long)
'���ܣ���ǰ���ڱ༭״̬����Ҫ��ѡ��ʱ������Ҫ�ؽ���ѡ���������
Dim leKey As Long
Dim strElementName As String, strElementValue As String, lEleId As Long, intLm As Integer, intEl As Integer, intLn As Integer, blnLm As Boolean
Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean

    On Error GoTo errHand
    If Not (Document.EPRFileInfo.���� = cpr���ﲡ�� Or Document.EPRFileInfo.���� = cprסԺ����) Then Exit Sub 'ֻ������ﲡ����סԺ����
    If Document.EditType = cprET_��������� Then Exit Sub
    strElementName = Document.Elements("K" & lKey).Ҫ������
    strElementValue = Document.Elements("K" & lKey).�����ı�
    lEleId = Document.Elements("K" & lKey).����Ҫ��ID
    
    For intLm = 1 To UBound(mEleLimit)
        Select Case mEleLimit(intLm).�䶯ԭ��
            Case 2, 3
            Case 1, 4
                If strElementName = mEleLimit(intLm).ԭ��Ҫ�� And lEleId = mEleLimit(intLm).ԭ��Ҫ��id Then
                    Select Case mEleLimit(intLm).�䶯���
                        Case 1
                            '��ѡ������ԭ������:���ơ�ID��ѡ��
                            If (strElementValue = mEleLimit(intLm).ԭ������ Or InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0 Or InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0) Then
                                'ԭ��Ҫ�ر�ѡ��
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).Ҫ������ = mEleLimit(intLm).���Ҫ�� And Document.Elements("K" & leKey).����Ҫ��ID = mEleLimit(intLm).���Ҫ��id Then
                                        If Document.Elements("K" & leKey).������̬ = 1 Then
                                            Document.Elements("K" & leKey).Ҫ��ֵ�� = mEleLimit(intLm).���ֵ��
                                            If Document.Elements("K" & leKey).Ҫ�ر�ʾ = 2 Then
                                                Document.Elements("K" & leKey).�����ı� = "��" & Replace(mEleLimit(intLm).���ֵ��, ";", "  ��")
                                            Else
                                                Document.Elements("K" & leKey).�����ı� = "��" & Replace(mEleLimit(intLm).���ֵ��, ";", "  ��")
                                            End If
                                            Document.Elements("K" & leKey).Refresh Editor1
                                        Else
                                            Document.Elements("K" & leKey).Ҫ��ֵ�� = mEleLimit(intLm).���ֵ��
                                        End If
                                    End If
                                Next
                            ElseIf strElementValue = "" Or Not (InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0 Or InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0) Then
                                'ԭ��Ҫ��û��ѡ�� ��û��ѡ��ԭ��Ҫ�ص�ָ������
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).Ҫ������ = mEleLimit(intLm).���Ҫ�� And Document.Elements("K" & leKey).����Ҫ��ID = mEleLimit(intLm).���Ҫ��id Then
                                        '����Ƿ�����ԭ�����ƣ��类���ƣ��򲻻�ԭ;������ͬһԭ��Ҫ�ز�ͬѡ��
                                        blnLm = False
                                        For intLn = 1 To UBound(mEleLimit)
                                            With Document.Elements("K" & leKey)
                                                If .Ҫ������ = mEleLimit(intLn).���Ҫ�� And .����Ҫ��ID = mEleLimit(intLn).���Ҫ��id _
                                                    And mEleLimit(intLn).�䶯��� = 1 And .Ҫ��ֵ�� = mEleLimit(intLn).���ֵ�� Then
                                                    If Document.Elements("K" & lKey).������̬ = 1 Then 'չ����
                                                        If InStr(strElementValue, "��") > 0 Then blnLm = True
                                                        If InStr(strElementValue, "��" & mEleLimit(intLn).ԭ������) > 0 Then blnLm = True
                                                    Else                                               '������
                                                        If InStr(strElementValue, mEleLimit(intLn).ԭ������) > 0 Then blnLm = True
                                                    End If
                                                     If blnLm = True Then Exit For
                                                End If
                                            End With
                                        Next
                                    
                                        If Not blnLm Then
                                            If Document.Elements("K" & leKey).������̬ = 1 Then
                                                Document.Elements("K" & leKey).Ҫ��ֵ�� = mEleLimit(intLm).ԭʼֵ��
                                                If Document.Elements("K" & leKey).Ҫ�ر�ʾ = 2 Then
                                                    Document.Elements("K" & leKey).�����ı� = "��" & Replace(mEleLimit(intLm).ԭʼֵ��, ";", "  ��")
                                                Else
                                                    Document.Elements("K" & leKey).�����ı� = "��" & Replace(mEleLimit(intLm).ԭʼֵ��, ";", "  ��")
                                                End If
                                                Document.Elements("K" & leKey).Refresh Editor1
                                            Else
                                                Document.Elements("K" & leKey).Ҫ��ֵ�� = mEleLimit(intLm).ԭʼֵ��
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        Case 3 'ɾ��Ҫ��
                            If (strElementValue = mEleLimit(intLm).ԭ������ Or InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0 Or InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0) Then
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).Ҫ������ = mEleLimit(intLm).���Ҫ�� And Document.Elements("K" & leKey).����Ҫ��ID = mEleLimit(intLm).���Ҫ��id Then
                                        If Document.Elements("K" & leKey).������̬ = 1 Then
                                            If InStr(Document.Elements("K" & leKey).�����ı�, "��") = 0 And InStr(Document.Elements("K" & leKey).�����ı�, "��") = 0 Then
                                                '���Ҫ��û��ѡ�ѡ�в�ɾ������ˢ����ʾ
                                                Document.Elements("K" & leKey).DeleteFromEditor Editor1
                                                Exit For
                                            End If
                                        Else
                                            If Document.Elements("K" & leKey).�����ı� = "" Then Document.Elements("K" & leKey).DeleteFromEditor Editor1: Exit For
                                        End If
                                    End If
                                Next
                            End If
                        Case 4 'ͬʱ���
                            If strElementName = mEleLimit(intLm).ԭ��Ҫ�� And lEleId = mEleLimit(intLm).ԭ��Ҫ��id Then
                                For intEl = 1 To Document.Elements.Count
                                    leKey = Document.Elements(intEl).Key
                                    If Document.Elements("K" & leKey).Ҫ������ = mEleLimit(intLm).���Ҫ�� And Document.Elements("K" & leKey).����Ҫ��ID = mEleLimit(intLm).���Ҫ��id Then
                                        If Document.Elements("K" & leKey).������̬ = 1 Then
                                            '���Ҫ��û��ѡ�ѡ�вŸ���ֵ�����ݣ���ˢ����ʾ
                                            Document.Elements("K" & leKey).�����ı� = strElementValue
                                            Document.Elements("K" & leKey).Refresh Editor1
                                        Else
                                            '���Ҫ��û��ѡ�ѡ�вŸ���ֵ�����ݣ���ˢ����ʾ
                                            Document.Elements("K" & leKey).�����ı� = strElementValue
                                            Document.Elements("K" & leKey).Refresh Editor1
                                        End If
                                    End If
                                Next
                            End If
                        Case 5
                            If (strElementValue = mEleLimit(intLm).ԭ������ Or InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0 Or InStr(strElementValue, "��" & mEleLimit(intLm).ԭ������) > 0) Then
                                If FindKey(Editor1, "E", lKey, lSS, lSE, lES, lEE, bNeeded) Then
                                    Editor1.ForceEdit = True
                                    Editor1.Range(lEE, lEE).Selected
                                    Editor1.Range(lEE, lEE).Font.Protected = False
                                    If Editor1.Range(lEE, lEE + 1).Text <> "��" Or Editor1.Range(lEE, lEE + 1).Text <> "��" _
                                       Or Editor1.Range(lEE, lEE + 1).Text <> "," Or Editor1.Range(lEE, lEE + 1).Text <> "." Then
                                        Editor1.Range(lEE + 1, lEE + 1).Selected
                                    Else
                                        Editor1.Range(lEE, lEE).Selected
                                    End If
                                    mfrmSentenceDetailed_RowDblClick mEleLimit(intLm).���Ҫ��id
                                End If
                            End If
                    End Select
                End If
        End Select
    Next
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function MakeSentenceLimit(ByVal lngCompend As Long) As String
'���ܣ�ͨ���Ƚϲ������ƴʾ䣬Ҫ�����ƴʾ䣬�����Զ��ŷָ���,�Զ��ſ�ʼ���Զ��Ž���
Dim strReturn As String, intLm As Integer, intEl As Integer, leKey As Long
    On Error GoTo errHand
    If Not (Document.EPRFileInfo.���� = cpr���ﲡ�� Or Document.EPRFileInfo.���� = cprסԺ����) Then Exit Function 'ֻ������ﲡ����סԺ����
    If Document.EditType = cprET_��������� Then Exit Function
    
    For intLm = 1 To UBound(mEleLimit)
        Select Case mEleLimit(intLm).�䶯ԭ��
            Case 2, 3 '��������ʾ�仯
                If (mlDiseaseID = mEleLimit(intLm).ԭ��Ҫ��id Or mlDiagnoseID = mEleLimit(intLm).ԭ��Ҫ��id) And _
                    mEleLimit(intLm).�䶯��� = 2 And mEleLimit(intLm).������id = lngCompend Then
                    strReturn = strReturn & "," & mEleLimit(intLm).���Ҫ��id
                End If
            Case 1 '��Ҫ������ʾ�仯,��Ҫ�˶�ԭ��Ҫ���Ƿ�����ֵ�����ơ�ID������λ��
                If mEleLimit(intLm).�䶯��� = 2 And mEleLimit(intLm).������id = lngCompend Then
                    For intEl = 1 To Document.Elements.Count
                        leKey = Document.Elements(intEl).Key
                        If Document.Elements("K" & leKey).�����ı� <> "" Then
                            If Document.Elements("K" & leKey).Ҫ������ = mEleLimit(intLm).ԭ��Ҫ�� And Document.Elements("K" & leKey).����Ҫ��ID = mEleLimit(intLm).ԭ��Ҫ��id Then
                                If Document.Elements("K" & leKey).�����ı� = mEleLimit(intLm).ԭ������ Or _
                                    InStr(Document.Elements("K" & leKey).�����ı�, "��" & mEleLimit(intLm).ԭ������) > 0 Or _
                                    InStr(Document.Elements("K" & leKey).�����ı�, "��" & mEleLimit(intLm).ԭ������) > 0 Then
                                        strReturn = strReturn & "," & mEleLimit(intLm).���Ҫ��id
                                End If
                            End If
                        End If
                    Next
                End If
        End Select
    Next
    MakeSentenceLimit = Decode(strReturn, "", "", strReturn & ",")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'################################################################################################################
'## ���ܣ�  ���ñ�����ɫ
'################################################################################################################
Private Sub ColorFillColor_pOK()
    SendKeys "{ESCAPE}"
    mlngCellFillColor = IIf(ColorFillColor.COLOR = tomAutoColor, -1, ColorFillColor.COLOR)
    If tblThis.Visible Then
        Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long, i As Long, j As Long
        If tblThis.SelStartRow > tblThis.SelEndRow Then
            lRow1 = tblThis.SelEndRow
            lRow2 = tblThis.SelStartRow
        Else
            lRow1 = tblThis.SelStartRow
            lRow2 = tblThis.SelEndRow
        End If
        If tblThis.SelStartCol > tblThis.SelEndCol Then
            lCol1 = tblThis.SelEndCol
            lCol2 = tblThis.SelStartCol
        Else
            lCol1 = tblThis.SelStartCol
            lCol2 = tblThis.SelEndCol
        End If
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then Exit Sub
        For i = lRow1 To lRow2
            For j = lCol1 To lCol2
                tblThis.Cell(i, j).BackColor = mlngCellFillColor
            Next j
        Next i
        SetColorIcon "FILLCOLOR", ID_DRAW_FILLCOLOR, mlngCellFillColor
    End If
    tblThis.Modified = True
    tblThis.Refresh False, False
    SendKeys "{ESCAPE}"
End Sub

'################################################################################################################
'## ���ܣ�  ��������ǰ��ɫ
'################################################################################################################
Private Sub ColorForeColor_pOK()
    SendKeys "{ESCAPE}"
    mlngSelForeColor = ColorForeColor.COLOR
    If tblThis.Visible Then
        Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long, i As Long, j As Long
        If tblThis.SelStartRow > tblThis.SelEndRow Then
            lRow1 = tblThis.SelEndRow
            lRow2 = tblThis.SelStartRow
        Else
            lRow1 = tblThis.SelStartRow
            lRow2 = tblThis.SelEndRow
        End If
        If tblThis.SelStartCol > tblThis.SelEndCol Then
            lCol1 = tblThis.SelEndCol
            lCol2 = tblThis.SelStartCol
        Else
            lCol1 = tblThis.SelStartCol
            lCol2 = tblThis.SelEndCol
        End If
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then Exit Sub
        For i = lRow1 To lRow2
            For j = lCol1 To lCol2
                tblThis.Cell(i, j).ForeColor = mlngSelForeColor
            Next j
        Next i
        SetColorIcon "FORECOLOR", ID_FORMAT_FORECOLOR, mlngSelForeColor
    End If
    tblThis.Modified = True
    tblThis.Refresh False, False
    SendKeys "{ESCAPE}"
End Sub

'################################################################################################################
'## ���ܣ�  �������屳��ɫ
'################################################################################################################
Private Sub ColorHighlight_pOK()
    SendKeys "{ESCAPE}"
    mlngSelHightlightColor = ColorHighlight.COLOR
    If Editor1.Selection.Font.Protected = False And Editor1.Selection.Font.Hidden = False Then
        Editor1.Tag = "ColorHighlight_pOK"
        Editor1.ForceEdit = True
        Editor1.Selection.Font.BackColor = mlngSelHightlightColor
        Editor1.ForceEdit = False
        Editor1.Tag = ""
        SetColorIcon "HIGHLIGHT", ID_FORMAT_HIGHLIGHT, IIf(mlngSelHightlightColor = tomAutoColor, vbWhite, mlngSelHightlightColor)
    End If
    If tblThis.Visible Then
        tblThis.SetFocus
    Else
        If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ����ҳ�汳��ɫ
'################################################################################################################
Private Sub ColorPaperBackColor_pOK()
    Editor1.PaperColor = IIf(ColorPaperBackColor.COLOR = tomAutoColor, vbWhite, ColorPaperBackColor.COLOR)
End Sub

'################################################################################################################
'## ���ܣ�  ���˲���
'################################################################################################################
Private Sub ExecBackSpace()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim i As Long, j As Long, lLen As Long
    Dim lS As Long, lE As Long, lSS As Long, lSS2 As Long
    Dim lf As Long, LL As Long, lR As Long, W As Long
    
    If Editor1.UIVisibled Then Editor1.CloseUIInterface

    Call AddUndoPoint  '�ֶ�����

    With Editor1
        .Tag = "ExecBackSpace"
        If .AuditMode Then
            i = .Selection.StartPos
            j = .Selection.StartPos + .SelLength

            '�˸������
            If i = j Then
                If .Range(i - 1, i).Font.Protected Or .Range(i - 1, i).Font.Hidden Then Exit Sub
                If Me.Document.IsNewCharColor(.Range(i - 1, i).Font.ForeColor) And .Range(i - 1, i).Font.Strikethrough = False Then
                    'ǰ��һ���ַ��Ѿ��������ı�����ֱ��ɾ��֮
                    .Range(i - 1, i).Text = ""
                Else
                    '���򣬱��ǰ���ı�Ϊɾ���ı�
                    .Range(i - 1, i).Font.Strikethrough = True
                    .Range(i - 1, i).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i - 1, i).Font.ForeColor)
                    .Range(i - 1, i - 1).Selected
                End If
            Else
                If Me.Document.IsNewCharColor(.Range(i, j).Font.ForeColor) And .Range(i, j).Font.Strikethrough = False Then
                    'ѡ���ı�Ϊ�����ı���ֱ��ɾ��֮
                    .Range(i, j) = ""
                ElseIf Me.Document.IsNewCharColor(.Range(i, j).Font.ForeColor) = False And Me.Document.IsDelCharColor(.Range(i, j).Font.ForeColor) = False And .Range(i, j).Font.ForeColor <> tomUndefined Then
                    '�������Ϊ��ͨ�ı���ֱ�ӱ��Ϊɾ��
                    .Range(i, j).Font.Strikethrough = True
                    .Range(i, j).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i, j).Font.ForeColor)
                ElseIf .Range(i, j).Font.ForeColor = tomUndefined Then
                    '�������Ϊ����ı����򲻴���
                    .Range(j, j).Selected
                End If
            End If
        Else
            '��ͨ��дģʽ
            lS = .Selection.StartPos
            lE = .Selection.EndPos
            lSS = IIf(lS - 2 > 0, lS - 2, 0)
            lSS2 = IIf(lS - 16 > 0, lS - 16, 0)
            If .Range(lSS, lS) = vbCrLf Or lS = 0 Or (.Range(lSS2, lSS2 + 3) = "OE(" And .Range(lSS2, lSS2 + 3).Font.Hidden = True) Then
                '���ף�������������
                lf = .Range(lS, lE).Para.FirstLineIndent
                LL = .Range(lS, lE).Para.LeftIndent
                lR = .Range(lS, lE).Para.RightIndent
                If lf = tomUndefined Then lf = 0
                If LL = tomUndefined Then LL = 0
                If lR = tomUndefined Then lR = 0

                W = (.PaperWidth - .MarginLeft - .MarginRight - 3000) * .ZoomFactor / 20

                If lf > 0 Then
                    lf = 0
                Else
                    LL = LL - .DefaultTabStop
                End If
                If LL < 0 Then LL = 0
                .ForceEdit = True
                .Range(lS, lE).Para.SetIndents lf, LL, lR
                .ForceEdit = False
            ElseIf .Range(lE - 1, lS).Font.Protected = False Then
                .ForceEdit = True
                .Range(lE - 1, lS) = ""
                .ForceEdit = False
            End If
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If
        End If
        .Tag = ""
    End With
    Call ClearNoUseUndoList
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵ�ճ������������Ҫ�عؼ��֣�����ɾ����Ҫ��Ҫ����Ϊ�������޶��ı�Ҳͳһ��Ϊ�����ı���
'################################################################################################################
Private Sub ExecPaste(ByRef edtThis As Object)
    If edtThis.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim i As Long, bForce As Boolean, bFinded As Boolean, strTmp As String, lS As Long, lE As Long, lngLen As Long
    Dim ParaFmt As New cParaFormat
    
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected Then Exit Sub
            If tblThis.InEdit Then
                tblThis.InsertText Clipboard.GetText
            Else
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Clipboard.GetText
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
            tblThis.Modified = True
        End If
        Exit Sub
    End If

    bBeteenKeys = IsBetweenAnyKeys(edtThis, edtThis.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then Exit Sub    '������ճ����Ԫ���ڲ�
'    Call AddUndoPoint  '�ֶ�����

    If edtThis.Selection.Font.ForeColor = tomUndefined Or edtThis.Selection.Font.Protected Then Exit Sub

    '���������Ϊ�գ���ô��ճ���ڲ�����
    Dim strClipboard As String
    strClipboard = Clipboard.GetText
    If Len(Trim(strClipboard)) > 0 Then
        'ճ������������
        lS = edtThis.Selection.StartPos
        lE = lS + Len(strClipboard)
        edtThis.Tag = "ExecPaste"
        edtThis.ForceEdit = True
        edtThis.Range(lS, edtThis.Selection.EndPos).Text = strClipboard
        edtThis.Range(lS, lE).Font.Strikethrough = False
        edtThis.Range(lS, lE).Font.Protected = False
        edtThis.Range(lS, lE).Font.ForeColor = IIf(Me.Document.EditType = cprET_���������, Me.Document.GetNewCharColor(vbBlack), tomAutoColor)
        edtThis.ForceEdit = False
        edtThis.Tag = ""
        edtThis.Range(lE, lE).Selected
        Exit Sub
    End If

    '�������ؼ���
    gfrmPublic.edtPublic.ForceEdit = True
    '�滻��Ϊu����ֹ����
'    gfrmPublic.edtPublic.Text = Replace(gfrmPublic.edtPublic.Text, "��", "u") '�ᵼ�±༭���������ֶε����Զ�ʧ����ʱ���Σ��ҵ��õĽ�����������޸�
    For i = 1 To gfrmPublic.Elements.Count
        '����Ҫ��
        lKey = Me.Document.Elements.AddExistNode(gfrmPublic.Elements(i).Clone, False)
        Me.Document.Elements("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾
        Me.Document.Elements("K" & lKey).��ֹ�� = 0     'ȥ����ֹ��
        Me.Document.Elements("K" & lKey).�������� = False
        Me.Document.Elements("K" & lKey).ID = 0
        '�����ؼ���
        bFinded = FindKey(gfrmPublic.edtPublic, "E", gfrmPublic.Elements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            strTmp = Format(lKey, "00000000") & "," & IIf(Me.Document.Elements("K" & lKey).��������, 1, 0) & ",0)"
            gfrmPublic.edtPublic.Range(lKSS, lKSE) = "ES(" & strTmp
            gfrmPublic.edtPublic.Range(lKES, lKEE) = "EE(" & strTmp
            gfrmPublic.Elements(i).Key = lKey '�����ı���ͬʱ������Key
        End If
    Next

    '����RTF���ݣ����ǰ��ɫ��ɾ����
    bForce = edtThis.ForceEdit
    edtThis.Tag = "ExecPaste"
    edtThis.Freeze
    edtThis.ForceEdit = True

    lS = 0: lE = Len(gfrmPublic.edtPublic.Text)
    If Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_��������� Then
        For i = lS To lE - 1
            '����ģʽ�£�ȫ������Ϊ�����ı���ȥ������
            If gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected Then
                '�����ı���Ϊ�����ı�
                gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected = False
            End If
            gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = IIf(Me.Document.EditType = cprET_���������, Me.Document.GetNewCharColor(vbBlack), tomAutoColor)
        Next
    Else
        For i = lS To lE - 1
            '����ģʽ�£�ȫ������Ϊ�����ı���ȥ������
            If gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected Then
                '�����ı�����
            Else
                gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = tomAutoColor
            End If
        Next
    End If

    gfrmPublic.edtPublic.SelectAll
    gfrmPublic.edtPublic.Selection.Font.Strikethrough = False
    lS = edtThis.Selection.StartPos
    '1����������ʽ
    Set ParaFmt = edtThis.Range(lS, lS).Para.GetParaFmt
    '2��ͬʱ����Tab�Ʊ�λ
    Dim j As Long
    Dim iT As Single, lA As Long, lLd As Long, LL As Long
    Dim iTabPos() As Long, lAlign() As Byte, lLeader() As Long
    j = edtThis.Range(lS, lS).Para.TabCount

    If j = tomUndefined Then j = 0
    ReDim iTabPos(0 To j) As Long
    ReDim lAlign(0 To j) As Byte
    ReDim lLeader(0 To j) As Long
    For i = 0 To j - 1
        edtThis.TOM.TextDocument.Range(lS, lS).Para.GetTab i, iT, lA, LL
        iTabPos(i) = iT * 20
        lAlign(i) = lA * 20
        lLeader(i) = lLd * 20
    Next

    lngLen = Len(gfrmPublic.edtPublic.Text)
    If lngLen > 0 Then
'        gfrmPublic.edtPublic.CopyWithFormat
'        edtThis.PasteWithFormat
        edtThis.TOM.TextDocument.Selection.FormattedText = gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText
        '3���ָ������ʽ
        edtThis.Range(lS, lS).Para.SetParaFmt ParaFmt
        '4���ָ�Tab�Ʊ�λ
        For i = 0 To UBound(iTabPos)
            If iTabPos(i) > 0 Then edtThis.TOM.TextDocument.Range(lS, lS).Para.AddTab iTabPos(i) / 20, lAlign(i), tomSpaces
        Next i
        'ȥ��ĩβ�Ļس����з�
        If edtThis.Range(lS + lngLen, lS + lngLen + 2) = vbCrLf And edtThis.Range(lS + lngLen, lS + lngLen + 2).Font.Protected = False Then
            edtThis.Range(lS + lngLen, lS + lngLen + 2) = ""
        End If
        edtThis.Range(lS + lngLen, lS + lngLen).Selected
    End If
'    Clipboard.Clear

    edtThis.ForceEdit = bForce
    edtThis.UnFreeze
    edtThis.Tag = ""
    Call ClearNoUseUndoList
'    '��������б�
'    Me.Document.Compends.UpdateOrdersFromText edtThis
'    Me.Document.Compends.FillTree mfrmCompends.Tree
End Sub

'################################################################################################################
'## ���ܣ�  ���Ϊʾ���ʾ�
'################################################################################################################
Private Sub ExecSaveAsPhrase()
    Dim lngCompendID As Long, lngRetuId As Long, lngClassId As Long
    Dim rsTemp As ADODB.Recordset
    Err = 0: On Error GoTo errHand
    '��ȡ�����id
    If mfrmCompends.Tree.SelectedItem Is Nothing Then Exit Sub
    If Me.Document.EditType = cprET_�����ļ����� Then
        lngCompendID = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).ID
    Else
        lngCompendID = Me.Document.Compends(mfrmCompends.Tree.SelectedItem.Key).�������ID
    End If
    If lngCompendID = 0 Then MsgBox "��ʱ��ٲ��ܶ���ʾ���ʾ䣡", vbInformation, gstrSysName: Exit Sub
    
    '��ȡ��ʾ����id
    gstrSQL = "Select �ʾ����id From ������ٴʾ� Where ���id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", lngCompendID)
    If rsTemp.RecordCount > 0 Then
        lngClassId = rsTemp.Fields(0).Value
    Else
        MsgBox "��ǰ���û�����ôʾ�ʾ�������Ӧ������ϵ����Ա��ʼ���������ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    '��ȡ����ǰ�û��Ƿ��дʾ�ʾ������Ȩ��
    If InStr(1, gstrPrivsEpr, "ȫԺ�����ʾ�") = 0 And InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") = 0 And InStr(1, gstrPrivsEpr, "���˲����ʾ�") = 0 Then
        MsgBox "�㲻�߱��ʾ�ʾ�������Ȩ�ޣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    '������Ӵʾ䴰��
    lngRetuId = frmSentenceEdit.ShowMe(Me, True, 0, lngClassId, , True)
    If lngRetuId = 0 Then Exit Sub
    'ˢ�´ʾ��б�
    Call mfrmSentenceDetailed.zlSubRefList(lngRetuId)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵĸ��Ʋ����������ı���Ҫ�أ�
'################################################################################################################
Private Sub ExecCopy(ByRef edt As Object)

    If edt.ReadOnly Then Exit Sub
    
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngLen As Long, lngSum As Long

    If tblThis.Visible Then
        If tblThis.Row > 0 And tblThis.Col > 0 Then
            Clipboard.Clear
            Clipboard.SetText tblThis.Cell(tblThis.Row, tblThis.Col).Text
        End If
        Exit Sub
    End If

    '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
    lS = edt.Selection.StartPos
    lE = edt.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '�ȿ���RTF����
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
'    edt.Range(lS, lE).Selected
'    edt.CopyWithFormat
'    gfrmPublic.edtPublic.PasteWithFormat
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = edt.TOM.TextDocument.Range(lS, lE).FormattedText
    '����Ҫ�أ���������Ԫ�أ�ͼƬ����ϡ����ȣ����ؼ���ҲҪ������ȥ����֤�����ݵ����عؼ���Keyֵһ�£�
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(edt, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '��Χ�ڴ��ڹؼ���
                If sKeyType = "E" Then
                    '�����Ҫ�أ���ô������������
                    gfrmPublic.Elements.AddExistNode Me.Document.Elements("K" & lKey).Clone(True), True
                Else
                    '���������Ԫ�أ������֮����gfrmPublic.edtPublic�����������¼��ǰλ�ã���
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '��¼ɾ�����ݵ��ܳ���
                End If
            Else
                '���򣬳�����Χ���˳�ѭ��
                Exit For
            End If
            i = lKEE - 1
        Else
            '�������κ�Ԫ�أ���ô�˳�ѭ��
            Exit For
        End If
    Next
    Clipboard.Clear
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵļ��в����������ı���Ҫ�أ�
'################################################################################################################
Private Sub ExecCut()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngNum As Long, lngSum As Long

    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected Then Exit Sub
            Clipboard.Clear
            Clipboard.SetText tblThis.Cells("K" & tblThis.SelectedCellKey).Text
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = ""
            tblThis.Modified = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
        Exit Sub
    End If

    Call AddUndoPoint  '�ֶ�����

    '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
    lS = Editor1.Selection.StartPos
    lE = Editor1.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(Editor1, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(Editor1, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    'ĩβλ�û����ܿ�Խ���
    bFinded = FindNextKey(Editor1, lS + 1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        If lKSS < lE Then
            If Editor1.Range(lKSS - 2, lKSS) = vbCrLf Then
                lE = lKSS - 2
            Else
                lE = lKSS
            End If
        End If
    End If
    If Editor1.Range(lE - 2, lE) = vbCrLf Then lE = lE - 2

    '�ȿ���RTF����
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
'    Me.Editor1.Range(lS, lE).Selected
'    Me.Editor1.CopyWithFormat
'    gfrmPublic.edtPublic.PasteWithFormat
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = Me.Editor1.TOM.TextDocument.Range(lS, lE).FormattedText
    '����Ҫ�أ���������Ԫ�أ�ͼƬ����ϡ����ȣ����ؼ���ҲҪ������ȥ����֤�����ݵ����عؼ���Keyֵһ�£�
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '��Χ�ڴ��ڹؼ���
                If sKeyType = "E" Then
                    '�����Ҫ�أ���ô������������
                    gfrmPublic.Elements.AddExistNode Me.Document.Elements("K" & lKey), True
                Else
                    '���������Ԫ�أ������֮����gfrmPublic.edtPublic�����������¼��ǰλ�ã���
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '��¼ɾ�����ݵ��ܳ���
                End If
            Else
                '���򣬳�����Χ���˳�ѭ��
                Exit For
            End If
            i = lKEE - 1
        Else
            '�������κ�Ԫ�أ���ô�˳�ѭ��
            Exit For
        End If
    Next

    'ɾ��ѡ������
    Dim bForce As Boolean, COLOR As OLE_COLOR, bProtect1 As Boolean, bProtect2 As Boolean
    bForce = Me.Editor1.ForceEdit
    Me.Editor1.Freeze
    Me.Editor1.Tag = "ExecCut"
    Me.Editor1.ForceEdit = True
    If Me.Editor1.AuditMode Then
        '���ģʽ�Ļ�����Ҫ������ɫ�Ͱ汾���⴦��
        '����Ԫ��
        For i = lS To lE
            bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then
                    '��Χ�ڴ��ڹؼ���
                    Select Case sKeyType
                    Case "E"    'Ҫ��
                        If Me.Document.Elements("K" & lKey).�������� = False Then
                            If Me.Document.Elements("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                                '���ΰ汾�½���Ҫ��
                                Me.Editor1.Range(lKSS, lKEE) = ""
                                lE = lE - (lKEE - lKSS)
                                i = lKSS - 1
                            ElseIf Me.Document.Elements("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 And _
                                Me.Document.Elements("K" & lKey).��ֹ�� = 0 Then
                                Me.Document.Elements("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                                Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                                i = lKEE - 1
                            Else
                                i = lKEE - 1
                            End If
                        Else
                            i = lKEE - 1
                        End If
                    Case "D"    '���
                        If Me.Document.Diagnosises("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                            '���ΰ汾�½���Ҫ��
                            Me.Editor1.Range(lKSS, lKEE) = ""
                            lE = lE - (lKEE - lKSS)
                            i = lKSS - 1
                        ElseIf Me.Document.Diagnosises("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 And _
                            Me.Document.Diagnosises("K" & lKey).��ֹ�� = 0 Then
                            Me.Document.Diagnosises("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                            Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                            i = lKEE - 1
                        Else
                            i = lKEE - 1
                        End If
                    Case Else
                       '���������Ԫ�أ��򲻴���
                       i = lKEE - 1
                    End Select
                Else
                    '���򣬳�����Χ���˳�ѭ��
                    Exit For
                End If
            Else
                '�������κ�Ԫ�أ���ô�˳�ѭ��
                Exit For
            End If
        Next

        '��������
        For i = lS To lE - 1
            If Editor1.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And Editor1.Range(i, i + 1).Font.Protected Then
                '�����ı���������ɾ����������
            ElseIf Editor1.Range(i, i + 1).Font.Protected = False Then
                COLOR = IIf(Editor1.Range(i, i + 1).Font.ForeColor = tomAutoColor Or Editor1.Range(i, i + 1).Font.ForeColor = tomUndefined, vbBlack, Editor1.Range(i, i + 1).Font.ForeColor)
                If Me.Document.IsNewCharColor(COLOR) And Editor1.Range(i, i + 1).Font.Strikethrough = False Then
                    '����һ���ַ��������ı�����ֱ��ɾ��֮
                    Editor1.Range(i, i + 1).Text = ""
                    lE = lE - 1
                    i = i - 1
                ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                    '�������ǰ���ı�Ϊ��ǰ�汾��ɾ���ı��������κδ���
                Else
                    '������Ϊɾ���ı�
                    Editor1.Range(i, i + 1).Font.Strikethrough = True
                    Editor1.Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Range(i, i + 1).Font.ForeColor)
                End If
            End If
        Next
        Me.Editor1.UnFreeze
        Me.Editor1.Range(lS, lE).Selected
    Else
        '���޶�ģʽ�����������Ҫ�ء�ͼƬ�������ϣ�����ɾ�����
        lngSum = 0
        For i = lS To lE - 1
            bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then   '��Χ�ڴ��ڹؼ���
                    '1���ȴ���ǰ�������
                    lngNum = DelTextRange(Me.Editor1, i, lKSS)
                    lE = lE - lngNum
                    lngSum = lngSum + lngNum
                    i = lKSS - lngNum - 1
                    '2���������һ��Ҫ�ء�ͼƬ��������
                    Select Case sKeyType
                    Case "E"    'Ҫ��
                        If Me.Document.Elements("K" & lKey).�������� = False Then
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Elements.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Else
                            i = lKEE - lngNum - 1
                        End If
                    Case "P"    'ͼƬ
                        Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                        Me.Document.Pictures.Remove "K" & lKey
                        lngSum = lngSum + (lKEE - lKSS)
                        lE = lE - (lKEE - lKSS)
                    Case "T"    '���
                        Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                        Me.Document.Tables.Remove "K" & lKey
                        lngSum = lngSum + (lKEE - lKSS)
                        lE = lE - (lKEE - lKSS)
                    Case "D"    '���
                        Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                        Me.Document.Diagnosises.Remove "K" & lKey
                        lngSum = lngSum + (lKEE - lKSS)
                        lE = lE - (lKEE - lKSS)
                    Case Else
                       '���������Ԫ�أ��򲻴���
                       i = lKEE - lngNum - 1
                    End Select
                Else
                    '���򣬳�����Χ���˳�ѭ��
                    Exit For
                End If
            Else
                '�������κ�Ԫ�أ���ô�˳�ѭ��
                Exit For
            End If
        Next
        If i < lE Then
            lngNum = DelTextRange(Me.Editor1, i, lE)
        End If
        Me.Editor1.UnFreeze
        Me.Editor1.SelLength = 0
        Me.Editor1.Range(lS, lS).Selected
    End If
    Me.Editor1.Tag = ""
    Me.Editor1.ForceEdit = bForce
    Clipboard.Clear
    Call ClearNoUseUndoList
End Sub
'################################################################################################################
'## ���ܣ�  �޶����ݵ�ɾ������
'################################################################################################################
Private Sub ExecAuditDelete()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim i As Long, j As Long, blnForce As Boolean
    Dim lngStart As Long, lngEnd As Long, lngLen As Long, blnWithEles As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strQuestion As String, COLOR As OLE_COLOR, lS As Long, lE As Long, bFinded As Boolean
    Dim bProtect1 As Boolean, bProtect2 As Boolean, lStart As Long
    Dim lngNum As Long, lngSum As Long, lIndex As Long
     'ѡ�����ݷǿ�
        '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
        lS = Editor1.Selection.StartPos
        lE = Editor1.Selection.EndPos
        '���ѡ��λ��������ĩβ���򲻴���
        If lS = Len(Editor1.Text) Then Exit Sub
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lS = lKSS
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lE = lKEE
        'ĩβλ�û����ܿ�Խ���
        bFinded = FindNextKey(Editor1, lS + 1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                If Editor1.Range(lKSS - 2, lKSS) = vbCrLf Then
                    lE = lKSS - 2
                Else
                    lE = lKSS
                End If
            End If
        End If
        If Editor1.Range(lE - 2, lE) = vbCrLf Then lE = lE - 2
    If Me.Editor1.AuditMode Then
            '���ģʽ�Ļ�����Ҫ������ɫ�Ͱ汾���⴦��
            '����Ԫ��
            For i = lS To lE
                bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bFinded Then
                    If lKSS < lE Then
                        '��Χ�ڴ��ڹؼ���
                        Select Case sKeyType
                        Case "E"    'Ҫ��
                            If Me.Document.Elements("K" & lKey).�������� = False Then
                                If Me.Document.Elements("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                                    '���ΰ汾�½���Ҫ��
                                    Me.Editor1.Range(lKSS, lKEE) = ""
                                    lE = lE - (lKEE - lKSS)
                                    i = lKSS - 1
                                ElseIf Me.Document.Elements("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 And _
                                    Me.Document.Elements("K" & lKey).��ֹ�� = 0 Then
                                    Me.Document.Elements("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                                    Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                                    i = lKEE - 1
                                Else
                                    i = lKEE - 1
                                End If
                            Else
                                i = lKEE - 1
                            End If
                        Case "D"    '���
                            If Me.Document.Diagnosises("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                                '���ΰ汾�½���Ҫ��
                                Me.Editor1.Range(lKSS, lKEE) = ""
                                lE = lE - (lKEE - lKSS)
                                i = lKSS - 1
                            ElseIf Me.Document.Diagnosises("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 And _
                                Me.Document.Diagnosises("K" & lKey).��ֹ�� = 0 Then
                                Me.Document.Diagnosises("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                                Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                                i = lKEE - 1
                            Else
                                i = lKEE - 1
                            End If
                        Case Else
                           '���������Ԫ�أ��򲻴���
                           i = lKEE - 1
                        End Select
                    Else
                        '���򣬳�����Χ���˳�ѭ��
                        Exit For
                    End If
                Else
                    '�������κ�Ԫ�أ���ô�˳�ѭ��
                    Exit For
                End If
            Next

            '��������
            For i = lS To lE - 1
                If Editor1.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And Editor1.Range(i, i + 1).Font.Protected Then
                    '�����ı���������ɾ����������
                ElseIf Editor1.Range(i, i + 1).Font.Protected = False Then
                    COLOR = IIf(Editor1.Range(i, i + 1).Font.ForeColor = tomAutoColor Or Editor1.Range(i, i + 1).Font.ForeColor = tomUndefined, vbBlack, Editor1.Range(i, i + 1).Font.ForeColor)
                    If Me.Document.IsNewCharColor(COLOR) And Editor1.Range(i, i + 1).Font.Strikethrough = False Then
                        '����һ���ַ��������ı�����ֱ��ɾ��֮
                        Editor1.Range(i, i + 1).Text = ""
                        lE = lE - 1
                        i = i - 1
                    ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                        '�������ǰ���ı�Ϊ��ǰ�汾��ɾ���ı��������κδ���
                    ElseIf Editor1.Range(i, i + 2).Text = vbCrLf Then
                        '�ǻس����з��Ҳ����ڱ���״̬����ֱ��ɾ��
                        If Not (Editor1.Range(i, i + 2).Font.Protected Or Editor1.Range(i, i + 2).Font.Hidden) Then
                            Editor1.Range(i, i + 2).Text = ""
                            If lE > lS Then lE = lE - 2
                            i = i - 1
                        End If
                    Else
                        '������Ϊɾ���ı�
                        Editor1.Range(i, i + 1).Font.Strikethrough = True
                        Editor1.Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Range(i, i + 1).Font.ForeColor)
                    End If
                End If
            Next
            Me.Editor1.UnFreeze
            Me.Editor1.Range(lE, lE).Selected
      End If
End Sub
'################################################################################################################
'## ���ܣ�  ���ݵ�ɾ������
'################################################################################################################
Private Sub ExecDelete()
    If Me.Editor1.ReadOnly Then Exit Sub
    Dim i As Long, j As Long, blnForce As Boolean
    Dim lngStart As Long, lngEnd As Long, lngLen As Long, blnWithEles As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strQuestion As String, COLOR As OLE_COLOR, lS As Long, lE As Long, bFinded As Boolean
    Dim bProtect1 As Boolean, bProtect2 As Boolean, lStart As Long
    Dim lngNum As Long, lngSum As Long, lIndex As Long
    
    If tblThis.Visible Then
        '����ڲ���ɾ��
        lKey = tblThis.SelectedCellKey
        If lKey <= 0 Then Exit Sub
        If tblThis.Cells("K" & lKey).Protected = True And Me.Document.EditType <> cprET_�����ļ����� Then Exit Sub
        If tblThis.Cells("K" & lKey).Text = "" And tblThis.Cells("K" & lKey).Protected Then
            'ɾ��Ҫ��/ͼƬ
            If Val(tblThis.Tag) > 0 Then
                If Val(tblThis.Cells("K" & lKey).Tag) > 0 Then
                    If tblThis.Cells("K" & lKey).Picture Is Nothing Then
                        Me.Document.Tables("K" & tblThis.Tag).Elements.Remove "K" & tblThis.Cells("K" & lKey).Tag
                    Else
                        Me.Document.Tables("K" & tblThis.Tag).Pictures.Remove "K" & tblThis.Cells("K" & lKey).Tag
                        Set tblThis.Cells("K" & lKey).Picture = Nothing
                    End If
                End If
            End If
            tblThis.Cells("K" & lKey).Tag = ""
            tblThis.Cells("K" & lKey).ToolTipText = ""
            tblThis.Cells("K" & lKey).Protected = False
            tblThis.Refresh False, True, lKey
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            'ɾ���ı�
            If tblThis.InEdit Then
                '����ڱ༭״̬
                tblThis.PressDelKey
                tblThis.Refresh False, True, lKey
                tblThis_Resize tblThis.Width, tblThis.Height
            Else
                tblThis.Cells("K" & lKey).Text = ""
                If Val(tblThis.Tag) > 0 Then
                    If Val(tblThis.Cells("K" & lKey).Tag) > 0 Then
                        If tblThis.Cells("K" & lKey).Picture Is Nothing Then
                            Me.Document.Tables("K" & tblThis.Tag).Elements("K" & tblThis.Cells("K" & lKey).Tag).�����ı� = ""
                        End If
                    End If
                End If
                tblThis.Refresh False, True, lKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        End If
        tblThis.Modified = True
        Exit Sub
    End If
    
    If Editor1.UIVisibled Then Editor1.CloseUIInterface
    Call AddUndoPoint  '�ֶ�����

    blnForce = Editor1.ForceEdit
    Editor1.Tag = "ExecDelete"
    Editor1.ForceEdit = True

    If Me.Editor1.SelLength > 0 Then
        'ѡ�����ݷǿ�
        '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
        lS = Editor1.Selection.StartPos
        lE = Editor1.Selection.EndPos
        '���ѡ��λ��������ĩβ���򲻴���
        If lS = Len(Editor1.Text) Then Exit Sub
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lS = lKSS
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lE = lKEE
        'ĩβλ�û����ܿ�Խ���
        bFinded = FindNextKey(Editor1, lS + 1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                If Editor1.Range(lKSS - 2, lKSS) = vbCrLf Then
                    lE = lKSS - 2
                Else
                    lE = lKSS
                End If
            End If
        End If
        If Editor1.Range(lE - 2, lE) = vbCrLf Then lE = lE - 2

        'ɾ��ѡ������
        Me.Editor1.Freeze
        If Me.Editor1.AuditMode Then
            '���ģʽ�Ļ�����Ҫ������ɫ�Ͱ汾���⴦��
            '����Ԫ��
            For i = lS To lE
                bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bFinded Then
                    If lKSS < lE Then
                        '��Χ�ڴ��ڹؼ���
                        Select Case sKeyType
                        Case "E"    'Ҫ��
                            If Me.Document.Elements("K" & lKey).�������� = False Then
                                If Me.Document.Elements("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                                    '���ΰ汾�½���Ҫ��
                                    Me.Editor1.Range(lKSS, lKEE) = ""
                                    lE = lE - (lKEE - lKSS)
                                    i = lKSS - 1
                                ElseIf Me.Document.Elements("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 And _
                                    Me.Document.Elements("K" & lKey).��ֹ�� = 0 Then
                                    Me.Document.Elements("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                                    Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                                    i = lKEE - 1
                                Else
                                    i = lKEE - 1
                                End If
                            Else
                                i = lKEE - 1
                            End If
                        Case "D"    '���
                            If Me.Document.Diagnosises("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                                '���ΰ汾�½���Ҫ��
                                Me.Editor1.Range(lKSS, lKEE) = ""
                                lE = lE - (lKEE - lKSS)
                                i = lKSS - 1
                            ElseIf Me.Document.Diagnosises("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 And _
                                Me.Document.Diagnosises("K" & lKey).��ֹ�� = 0 Then
                                Me.Document.Diagnosises("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                                Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                                i = lKEE - 1
                            Else
                                i = lKEE - 1
                            End If
                        Case Else
                           '���������Ԫ�أ��򲻴���
                           i = lKEE - 1
                        End Select
                    Else
                        '���򣬳�����Χ���˳�ѭ��
                        Exit For
                    End If
                Else
                    '�������κ�Ԫ�أ���ô�˳�ѭ��
                    Exit For
                End If
            Next

            '��������
            For i = lS To lE - 1
                If Editor1.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And Editor1.Range(i, i + 1).Font.Protected Then
                    '�����ı���������ɾ����������
                ElseIf Editor1.Range(i, i + 1).Font.Protected = False Then
                    COLOR = IIf(Editor1.Range(i, i + 1).Font.ForeColor = tomAutoColor Or Editor1.Range(i, i + 1).Font.ForeColor = tomUndefined, vbBlack, Editor1.Range(i, i + 1).Font.ForeColor)
                    If Me.Document.IsNewCharColor(COLOR) And Editor1.Range(i, i + 1).Font.Strikethrough = False Then
                        '����һ���ַ��������ı�����ֱ��ɾ��֮
                        Editor1.Range(i, i + 1).Text = ""
                        lE = lE - 1
                        i = i - 1
                    ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                        '�������ǰ���ı�Ϊ��ǰ�汾��ɾ���ı��������κδ���
                    ElseIf Editor1.Range(i, i + 2).Text = vbCrLf Then
                        '�ǻس����з��Ҳ����ڱ���״̬����ֱ��ɾ��
                        If Not (Editor1.Range(i, i + 2).Font.Protected Or Editor1.Range(i, i + 2).Font.Hidden) Then
                            Editor1.Range(i, i + 2).Text = ""
                            If lE > lS Then lE = lE - 2
                            i = i - 1
                        End If
                    Else
                        '������Ϊɾ���ı�
                        Editor1.Range(i, i + 1).Font.Strikethrough = True
                        Editor1.Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Range(i, i + 1).Font.ForeColor)
                    End If
                End If
            Next
            Me.Editor1.UnFreeze
            Me.Editor1.Range(lE, lE).Selected
        Else
            '���޶�ģʽ�����������Ҫ�ء�ͼƬ�������ϣ�����ɾ�����
            lngSum = 0
            For i = lS To lE - 1
                bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bFinded Then
                    If lKSS < lE Then   '��Χ�ڴ��ڹؼ���
                        '1���ȴ���ǰ�������
                        lngNum = DelTextRange(Me.Editor1, i, lKSS)
                        lE = lE - lngNum
                        lngSum = lngSum + lngNum
                        i = lKSS - lngNum - 1
                        '2���������һ��Ҫ�ء�ͼƬ��������
                        Select Case sKeyType
                        Case "E"    'Ҫ��
                            If Me.Document.Elements("K" & lKey).�������� = False Then
                                Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                                Me.Document.Elements.Remove "K" & lKey
                                lngSum = lngSum + (lKEE - lKSS)
                                lE = lE - (lKEE - lKSS)
                            Else
                                i = lKEE - lngNum - 1
                            End If
                        Case "P"    'ͼƬ
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Pictures.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Case "T"    '���
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Tables.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Case "D"    '���
                            Me.Editor1.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Me.Document.Diagnosises.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Case Else
                           '���������Ԫ�أ��򲻴���
                           i = lKEE - lngNum - 1
                        End Select
                    Else
                        '���򣬳�����Χ���˳�ѭ��
                        Exit For
                    End If
                Else
                    '�������κ�Ԫ�أ���ô�˳�ѭ��
                    Exit For
                End If
            Next
            If i < lE Then
                lngNum = DelTextRange(Me.Editor1, i, lE)
            End If
            Me.Editor1.UnFreeze
            Me.Editor1.SelLength = 0
            Me.Editor1.Range(lE - lngNum, lE - lngNum).Selected
        End If
        Clipboard.Clear
    Else
        'û��ѡ���ı�
        lS = Editor1.Selection.StartPos
        lE = Editor1.Selection.EndPos
        '���ѡ��λ��������ĩβ���򲻴���
        If lS = Len(Editor1.Text) Then Exit Sub
        If Editor1.AuditMode Then
            '���ģʽ
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys Then
                'ɾ����������Ҫ�ء�ͼƬ���߱��
                Select Case sKeyType
                Case "E"
                    If Document.Elements("K" & lKey).��ֹ�� > 0 Then Exit Sub
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "�Ƿ�ɾ��������Ҫ�أ�"
                Case "D"
                    If Document.Diagnosises("K" & lKey).��ֹ�� > 0 Then Exit Sub
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "�Ƿ�ɾ������ϣ�"
                Case Else
                    GoTo LL
                End Select
    '            If MsgBox(strQuestion, vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                Select Case sKeyType
                Case "E"
                    If Document.Elements("K" & lKey).�������� = False Then
                        If Document.Elements("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                            Document.Elements("K" & lKey).DeleteFromEditor Me.Editor1
                            Document.Elements.Remove "K" & lKey
                            Me.Editor1.SelLength = 0
                        Else
                            Document.Elements("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                            Document.Elements("K" & lKey).Refresh Me.Editor1
                            Me.Editor1.Range(lKEE, lKEE).Selected
                        End If
                    End If
                Case "D"
                    If Document.Diagnosises("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                        Document.Diagnosises("K" & lKey).DeleteFromEditor Me.Editor1
                        Document.Diagnosises.Remove "K" & lKey
                        Me.Editor1.SelLength = 0
                    Else
                        Document.Diagnosises("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1
                        Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                        Me.Editor1.Range(lKEE, lKEE).Selected
                    End If
                Case Else
                    GoTo LL
                End Select
    '            End If
            Else
                '�ı��ı༭
                With Editor1
                    i = .Selection.StartPos
                    j = .Selection.StartPos + .SelLength
                    If .Range(i, j).Font.Protected Or .Range(i, j).Font.Hidden Then GoTo LL
                    If .Range(i, i + 2) = vbCrLf Then
'                        COLOR = IIf(.Range(i, i + 2).Font.ForeColor = tomAutoColor Or .Range(i, i + 2).Font.ForeColor = tomUndefined, vbBlack, .Range(i, i + 2).Font.ForeColor)
                        If .Range(i, i + 2).Font.Protected And .Range(i, i + 2).Font.Hidden Then GoTo LL
'                        If Me.Document.IsNewCharColor(COLOR) And .Range(i, i + 2).Font.Strikethrough = False Then
                            '����һ���ַ��������ı�����ֱ��ɾ��֮
                            .Range(i, i + 2).Text = ""
'                        ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
'                            '�������ǰ���ı�Ϊ��ǰ�汾��ɾ���ı��������κδ���
'                            .Range(i + 2, i + 2).Selected
'                        Else
'                            '������Ϊɾ���ı�
'                            .Range(i, i + 2).Font.Strikethrough = True
'                            .Range(i, i + 2).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i, i + 2).Font.ForeColor)
'                            .Range(i + 2, i + 2).Selected   '����һλ
'                        End If
                    Else
                        COLOR = IIf(.Range(i, i + 1).Font.ForeColor = tomAutoColor, vbBlack, .Range(i, i + 1).Font.ForeColor)
                        If .Range(i, i + 1).Font.Protected And .Range(i, i + 1).Font.Hidden Then GoTo LL
                        If Me.Document.IsNewCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = False Then
                            '����һ���ַ��������ı�����ֱ��ɾ��֮
                            .Range(i, i + 1).Text = ""
                        ElseIf rgbBlue(COLOR) <> 0 And Me.Document.IsDelCharColor(COLOR) = False Then
                            '�������ǰ���ı�Ϊ��ǰ�汾��ɾ���ı��������κδ���
                            .Range(i + 1, i + 1).Selected
                        Else
                            '������Ϊɾ���ı�
                            .Range(i, i + 1).Font.Strikethrough = True
                            .Range(i, i + 1).Font.ForeColor = Me.Document.GetDelCharColor(.Range(i, i + 1).Font.ForeColor)
                            .Range(i + 1, i + 1).Selected   '����һλ
                        End If
                    End If
                End With
            End If
        Else
            '��дģʽ������Ϊ���ı�
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys Then
                'ɾ����������Ҫ�ء�ͼƬ���߱��
                Select Case sKeyType
                Case "T"
                    If Document.Tables("K" & lKey).Ԥ�����ID <> 0 Then
                        MsgBox "����ɾ��Ԥ������й���Ԫ�أ�����ɾ����Ԥ����ٱ���", vbOKOnly + vbInformation, gstrSysName
                        GoTo LL
                    Else
                        strQuestion = "�Ƿ�ɾ���ñ��"
                    End If
                Case "P"
                    strQuestion = "�Ƿ�ɾ����ͼƬ��"
                Case "E"
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "�Ƿ�ɾ��������Ҫ�أ�"
                Case "D"
                    Editor1.Range(lKSE, lKES).Selected
                    strQuestion = "�Ƿ�ɾ������ϣ�"
                Case Else
                    GoTo LL
                End Select
    '            If MsgBox(strQuestion, vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
                Select Case sKeyType
                Case "T"
                    Document.Tables.Remove "K" & lKey
                Case "P"
                    Document.Pictures.Remove "K" & lKey
                Case "E"
                    If Document.Elements("K" & lKey).�������� = False Or Me.Document.EditType = cprET_�����ļ����� Then
                        Document.Elements.Remove "K" & lKey
                    Else
                        GoTo LL
                    End If
                Case "D"
                    Document.Diagnosises.Remove "K" & lKey
                Case Else
                    GoTo LL
                End Select
                Editor1.Range(lKSS, lKEE) = ""
                If Editor1.Range(lKSS - 2, lKSS) = vbCrLf And Editor1.Range(lKSS - 2, lKSS).Font.Protected Then
                    Editor1.Range(lKSS - 2, lKSS) = ""
                    Editor1.Range(lKSS - 2, lKSS - 2).Font.Protected = False
                Else
                    Editor1.Range(lKSS, lKSS).Font.Protected = False
                End If
    '            End If
            Else
                'ɾ���ı�
                i = Editor1.Selection.StartPos
                j = Len(Editor1.Text)

                If Editor1.Range(i, i + 2).Font.Protected = False And Editor1.Range(i, i + 2) = vbCrLf And _
                    Editor1.Range(i + 2, i + 5) = "OS(" And Editor1.Range(i + 2, i + 5).Font.Hidden Then
                    '���Ȳ�����ɾ�����ǰ��Ļس������������ÿ���俪ʼλ�ã�
                    Editor1.Range(i + 2, i + 2).Selected
                ElseIf Editor1.Range(i, i + 1).Font.Protected = False And (Editor1.Range(i + 1, i + 2).Font.Protected = True Or i = j - 1) Then
                    Editor1.Range(i, i + 1) = ""
                ElseIf Editor1.Range(i - 1, i).Font.Protected = True And Editor1.Range(i, i + 1).Font.Protected = False Then
                    If Editor1.Range(i, i + 2) = vbCrLf And Editor1.Range(i, i + 2).Font.Protected = False Then
                        Editor1.Range(i, i + 2) = ""
                        Editor1.Range(i, i).Font.Protected = False
                    Else
                        Editor1.Delete
                    End If
                ElseIf Editor1.Range(i, i + 2) = vbCrLf And Editor1.Range(i, i + 2).Font.Protected = False Then
                    Editor1.Range(i, i + 2) = ""
                    Editor1.Range(i, i).Font.Protected = False
                ElseIf Editor1.Range(i, i).Font.Protected = False And Editor1.Range(i, i + 1).Font.Protected = False Then
                    Editor1.Delete
                ElseIf Editor1.Range(i, i + 2) = vbCrLf And Editor1.Range(i, i + 2).Font.Protected Then
                    Editor1.Range(i + 2, i + 2).Selected
                Else
                    Editor1.Range(i + 1, i + 1).Selected
                End If
            End If
        End If
    End If
    Call ClearNoUseUndoList
LL:
    Editor1.ForceEdit = blnForce
    Editor1.Tag = ""
End Sub

'################################################################################################################
'## ���ܣ�  ɾ��ָ����Χ���ı����޳��ܱ������ı�������ɾ�����ַ���
'################################################################################################################
Private Function DelTextRange(ByRef edtThis As Editor, ByVal lS As Long, ByVal lE As Long) As Long
    Dim i As Long, j As Long, lStart As Long, lNum As Long, lSum As Long
    Dim bProtect1 As Boolean, bProtect2 As Boolean
    edtThis.Tag = "DelTextRange"
    If Me.Document.EditType <> cprET_�����ļ����� Then
        '�޳������ı�
        lStart = lS
        For i = lS To lE - 1
            bProtect1 = edtThis.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(i, i + 1).Font.Protected
            bProtect2 = edtThis.Range(i + 1, i + 2).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(i + 1, i + 2).Font.Protected
            If bProtect1 = bProtect2 Then
                '�ı�״̬��ͬ��������
            Else
                '�ı�״̬��ͬ
                If bProtect1 Then
                    'ǰһλ��Ϊ�����ı���������
                    lStart = i + 1  '��¼�Ǳ����ı�����ʼλ��
                Else
                    'ǰһλ��Ϊ�Ǳ����ı��������֮
                    edtThis.Range(lStart, i + 1) = ""
                    lNum = i + 1 - lStart   '����ɾ�����ַ���
                    lSum = lSum + lNum
                    lE = lE - lNum
                    i = lStart - 1          'lStart����
                End If
            End If
        Next
        '���һֱ��״̬��ͬ
        If (bProtect1 = bProtect2) And (bProtect1 = False) And lStart < lE Then
            edtThis.Range(lStart, lE) = ""
            lNum = lE - lStart
            lSum = lSum + lNum
        End If
        DelTextRange = lSum
    Else
        edtThis.Range(lS, lE) = ""
        DelTextRange = lE - lS
    End If
    edtThis.Tag = ""
End Function

'################################################################################################################
'## ���ܣ�  �����û����Ի����ò˵��͹�����
'################################################################################################################
Private Sub cbrThis_Customization(ByVal Options As XtremeCommandBars.ICustomizeOptions)
    Dim Controls As CommandBarControls
    Set Controls = cbrThis.DesignerControls

    If (Controls.Count = 0) Then
        AddButton Controls, xtpControlButton, ID_FILE_CLEAR, "���", , "��յ�ǰ�ļ���������", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_IMPORT, "����...", , "�����ⲿ������ļ�", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_SAVE, "����", , "���浱ǰ�༭���ļ�", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_SAVE_QUIT, "�����˳�", , "���浱ǰ�༭���ļ����˳��༭��", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEASEPRDEMO, "���Ϊ����...", , "���Ϊ����", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEASSEGMENT, "���ΪƬ��...", , "���Ϊʾ��Ƭ��", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_SAVEAS, "����ΪRTF�ļ�...", , "����ΪRTF�ļ�", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_EXPORTTOXML, "����ΪXML�ļ�...", , "����ΪXML�ļ�", xtpButtonAutomatic, "�ļ�"
'        AddButton Controls, xtpControlButton, ID_FILE_EXPORTTOHTML, "����ΪHTML�ļ�...", , "����ΪHTML�ļ�", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_IMPORTFROMXML, "��XML�ļ�����...", , "��XML�ļ�����", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_PAGESETUP, "ҳ������...", , "ҳüҳ�š�ҳ��ߴ�����", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_PRINTPREVIEW, "��ӡԤ��", , "��ӡԤ��", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_PRINT, "��ӡ...", , "��ӡ��ǰ�ļ�", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_PRINTINWORD, "ͨ��Word��ӡ", , "ͨ��Word��ӡ��ǰ�ļ�", xtpButtonAutomatic, "�ļ�"
        AddButton Controls, xtpControlButton, ID_FILE_EXIT, "�˳�", , "�˳�ϵͳ", xtpButtonAutomatic, "�ļ�"

        AddButton Controls, xtpControlButton, ID_EDIT_UNDO, "����", , "�������һ�α༭", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_REDO, "����", , "�ظ����һ�α༭", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_CUT, "����", , "����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_COPY, "����", , "����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_PASTE, "ճ��", , "ճ��", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_FORMATBRUSH, "��ʽˢ", , "��ʽˢ", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_REFCOMPEND, "ˢ�����", , "ˢ�����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_ADDCOMPEND, "�������", , "�������", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_DELCOMPEND, "ɾ�����", , "ɾ�����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_MODCOMPEND, "�޸����", , "�޸����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_DELETE, "ɾ��", , "ɾ����ѡ����", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_SELECTALL, "ȫѡ", , "ȫѡ", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_FIND, "����...", , "����...", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_FINDNEXT, "������һ��", , "������һ��", xtpButtonAutomatic, "�༭"
        AddButton Controls, xtpControlButton, ID_EDIT_REPLACE, "�滻...", , "�滻...", xtpButtonAutomatic, "�༭"

        AddButton Controls, xtpControlButton, ID_VIEW_STRUCTURE, "�ĵ��ṹͼ", , "�ĵ��ṹͼ", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_PHRASEDEMO, "ʾ���ʾ��б�", , "ʾ���ʾ��б�", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_SEGMENT, "ʾ��Ƭ���б�", , "ʾ��Ƭ���б�", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_SEGMENT, "����ͼ�б�", , "����ͼ�б�", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_HISTORYWINDOW, "��ʷ�����б�", , "��ʷ�����б�", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_HISTORYREPORT, "��ʷ�����б�", , "��ʷ�����б�", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, XTP_ID_TOOLBARLIST, "�������б�", , "�������б�", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_HEADFOOT, "ҳüҳ��", , "ҳüҳ��", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_CHARCOUNT, "����ͳ��", , "����ͳ��", xtpButtonAutomatic, "��ͼ"
'        AddButton Controls, xtpControlButton, ID_VIEW_GRID, "������", , "������", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_RULER, "���", , "���", xtpButtonAutomatic, "��ͼ"
        AddButton Controls, xtpControlButton, ID_VIEW_PENWINDOW, "��д���봰��", , "��д���봰��", xtpButtonAutomatic, "��ͼ"

        AddButton Controls, xtpControlButton, ID_INSERT_DATETIME, "���ں�ʱ��", , "���ں�ʱ��", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_INSERT_DATE, "��������", , "��������", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_INSERT_TIME, "����ʱ��", , "����ʱ��", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_INSERT_SPECIALCHAR, "�������", , "�������", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_TABLE_INSERTTABLE, "������", , "������", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_INSERT_ELEMENT, "����Ҫ��", , "����Ҫ��", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_EDIT_ADDCOMPEND, "�������", , "�������", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_INSERT_EPRDEMO, "���뷶��", , "���뷶��", xtpButtonAutomatic, "����"

        AddButton Controls, xtpControlButton, ID_FORMAT_FONT, "����...", , "����...", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_PARA, "����...", , "����...", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_BOLD, "����", , "����", xtpButtonAutomatic, "��ʽ"
'        AddButton Controls, xtpControlButton, ID_FORMAT_ITALIC, "б��", , "б��", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_SUPER, "�ϱ�", , "�ϱ�", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_SUB, "�±�", , "�±�", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_NONE, "���»���", , "���»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_THIN, "ϸ�»���", , "ϸ�»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_THICK, "���»���", , "���»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_WAVE, "�����»���", , "�����»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DOT, "���»���", , "���»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DASH, "���»���", , "���»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT, "�㻮�»���", , "�㻮�»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT2, "˫�㻮�»���", , "˫�㻮�»���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_ALIGNLEFT, "�����", , "�����", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_ALIGNCENTER, "���ж���", , "���ж���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_ALIGNRIGHT, "�Ҷ���", , "�Ҷ���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTNONE, "����Ŀ��������", , "����Ŀ��������", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTBULLETS, "��Ŀ����(��)", , "��Ŀ����(��)", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTARABIC, "����������(1,2,3,...)", , "����������(1,2,3,...)", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTLCHAR, "Сд��ĸ(a,b,c,...)", , "Сд��ĸ(a,b,c,...)", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTUCHAR, "��д��ĸ(A,B,C,...)", , "��д��ĸ(A,B,C,...)", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTLROME, "Сд��������(i,ii,iii,...)", , "Сд��������(i,ii,iii,...)", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTUROME, "��д��������(I,II,III,...)", , "��д��������(I,II,III,...)", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LISTSETUP, "�Զ����ʽ...", , "�Զ����ʽ...", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE1, "1.0", , "1.0", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE2, "1.3", , "1.3", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE3, "1.5", , "1.5", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE4, "2.0", , "2.0", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE5, "2.5", , "2.5", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE6, "3.0", , "3.0", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_LINESPACE7, "����...", , "����...", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_SPACEBEFORE, "��ǰ���", , "��ǰ���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_SPACEAFTER, "�κ���", , "�κ���", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_FIRSTINDENT, "��������", , "��������", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_FIRSTHUNGING, "��������", , "��������", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_INDENTDECREASE, "����������", , "����������", xtpButtonAutomatic, "��ʽ"
        AddButton Controls, xtpControlButton, ID_FORMAT_INDENTINCREASE, "����������", , "����������", xtpButtonAutomatic, "��ʽ"

        AddButton Controls, xtpControlButton, ID_HELP_CONTENT, "��������", , "��������", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_HELP_ONLINE, gstrProductName & "����", , gstrProductName & "����", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_HELP_WEBFORUM, gstrProductName & "��̳", , gstrProductName & "��̳", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_HELP_CONTACT, "���ͷ���", , "���ͷ���", xtpButtonAutomatic, "����"
        AddButton Controls, xtpControlButton, ID_HELP_ABOUT, "����...", , "����...", xtpButtonAutomatic, "����"

        AddButton Controls, xtpControlButton, ID_REVISION_PREV, "ǰһ���޶�", , "ǰһ���޶�", xtpButtonAutomatic, "�޶�"
        AddButton Controls, xtpControlButton, ID_REVISION_NEXT, "��һ���޶�", , "��һ���޶�", xtpButtonAutomatic, "�޶�"
        AddButton Controls, xtpControlButton, ID_REVISION_RESET, "����޶�", , "����޶�", xtpButtonAutomatic, "�޶�"

        AddButton Controls, xtpControlButton, ID_DIAGNOSIS, "���", , "���", xtpButtonAutomatic, "���"

        AddButton Controls, xtpControlButton, ID_PATISIGN, "����ǩ��", , "����ǩ��", xtpButtonAutomatic, "ǩ��"
        AddButton Controls, xtpControlButton, ID_SIGN, "ǩ��", , "ǩ��", xtpButtonAutomatic, "ǩ��"
        AddButton Controls, xtpControlButton, ID_UNTREAD, "����", , "����", xtpButtonAutomatic, "ǩ��"
        AddButton Controls, xtpControlButton, ID_SIGN_QUIT, "ǩ���˳�", , "ǩ�����˳��༭", xtpButtonAutomatic, "ǩ��"
    End If
End Sub

'################################################################################################################
'## ���ܣ�  �˵�&������ִ���¼�
'################################################################################################################

Private Function AskOutputMode(ByRef blnOrigMode As Boolean, ByVal blnPreview As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ��ж��Ƿ�������ո�ʽ����ԭʼ��ʽ��ӡ/Ԥ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strAsk As String
    
    If Document.Ŀ��汾 > 1 And Document.EPRFileInfo.���� = cpr���Ʊ��� Then
        If zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1" Then
            blnOrigMode = False
        Else
            strAsk = "��" & Document.EPRFileInfo.���� & "�����й��޶���"
            strAsk = strAsk & "���԰�����ʽ��ԭʼ��ʽ" & IIf(blnPreview, "Ԥ��", "��ӡ") & "��"
            strAsk = strAsk & vbCrLf & "    ���ո�ʽ���������޸ĺۼ�������ʽ"
            strAsk = strAsk & vbCrLf & "    ԭʼ��ʽ�������޸ĺۼ��Ĳݸ��ʽ"
            strAsk = strAsk & vbCrLf & "�������ո�ʽ��ģʽ" & IIf(blnPreview, "Ԥ��", "��ӡ") & "��"
            
            Select Case MsgBox(strAsk, vbYesNoCancel + vbQuestion, gstrSysName)
            Case vbYes
                blnOrigMode = False
            Case vbNo
                blnOrigMode = True
            Case Else
                Exit Function
            End Select
        End If
        
        AskOutputMode = True
    Else
        AskOutputMode = True
    End If
        
End Function

Private Sub cbrThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    On Error Resume Next
    Dim i As Long, j As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim bFinded As Boolean
    Dim blnForce As Boolean, blnModified As Boolean
    Dim lngLen As Long
    Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long, blnTmp As Boolean, sText As String, strPic As String, objPic As StdPicture
    Dim blnOrigMode As Boolean
    Dim lngVkState As Long
    
    If mblnPrecess Then
        Exit Sub
    End If
    
    mblnPrecess = True
    If tblThis.Visible Then
        If tblThis.SelStartRow > tblThis.SelEndRow Then
            lRow1 = tblThis.SelEndRow
            lRow2 = tblThis.SelStartRow
        Else
            lRow1 = tblThis.SelStartRow
            lRow2 = tblThis.SelEndRow
        End If
        If tblThis.SelStartCol > tblThis.SelEndCol Then
            lCol1 = tblThis.SelEndCol
            lCol2 = tblThis.SelStartCol
        Else
            lCol1 = tblThis.SelStartCol
            lCol2 = tblThis.SelEndCol
        End If
    End If

    blnForce = Me.Editor1.ForceEdit
    Select Case Control.ID
    Case ID_FILE_CLEAR: Call ClearDoc: Call RecountPage(True)
    Case ID_FILE_IMPORT: Call ImportEPRDoc: Call RecountPage(True)
    Case ID_FILE_SAVE, ID_FILE_SAVE_QUIT
        If SaveEMRDoc And Control.ID = ID_FILE_SAVE_QUIT Then
            mblnPrecess = False: Unload Me: Exit Sub
        End If
        If Editor1.Enabled Then
            Editor1.SetFocus
        End If
    Case ID_FILE_SAVEAS
        If SaveDocToFile Then MsgBox "�����ɹ���", vbOKOnly + vbInformation, gstrSysName
    Case ID_FILE_SAVEASEPRDEMO: Call SaveDocAsEPRDemo
    Case ID_FILE_SAVEASSEGMENT: Call SaveDocAsSegment
    Case conMenu_File_Parameter
        '��������
        Dim frmSetup As New frmAutoSaveSetup
        
        If frmSetup.ShowMe(Me, gstrPrivsEpr) Then
            mblnAutosave = zlDatabase.GetPara("AutoSave", glngSys, 1070, 1) = 1
            mlngUndoLimit = zlDatabase.GetPara("UndoLimit", glngSys, 1070, 20)
            mlngSaveInterval = zlDatabase.GetPara("SaveInterval", glngSys, 1070, 60)
            mblnAutoSaveEPR = zlDatabase.GetPara("AutoSaveEPR", glngSys, 1070, 0) = 1
            mlngSaveIntervalEPR = zlDatabase.GetPara("SaveIntervalEPR", glngSys, 1070, 5)
            mblnAutoPageCount = zlDatabase.GetPara("AutoPageCount", glngSys, 1070, 0) = 1
            mblnAutoPageNote = zlDatabase.GetPara("AutoPageNote", glngSys, 1070, 0) = 1
            
            If mintSharePages <> Val(zlDatabase.GetPara("SharePageCount", glngSys, 1070, 5)) Then
                mintSharePages = Val(zlDatabase.GetPara("SharePageCount", glngSys, 1070, 5))
                mblnExistHistroy = ShowSharePageHistory(Me.Document, mintSharePages)
            End If
        End If
    Case ID_FILE_EXPORTTOXML:
        Call ExportXML
    Case ID_FILE_EXPORTTOHTML
        '������HTML��
        Dim strHTML As String
        Select Case Me.Document.EditType
        Case cprET_�����ļ�����
            dlgThis.Filename = "����_" & Me.Document.EPRFileInfo.���� & ".htm"
        Case cprET_ȫ��ʾ���༭
            dlgThis.Filename = "����_" & Me.Document.EPRFileInfo.���� & "_" & Me.Document.EPRDemoInfo.���� & ".htm"
        Case cprET_�������༭, cprET_���������
            dlgThis.Filename = "��¼_" & Me.Document.EPRFileInfo.���� & "(" & Me.Document.EPRPatiRecInfo.ID & "," & Me.Document.Ŀ��汾 & ").htm"
        End Select

        dlgThis.Filter = "*.htm|*.htm|*.html|*.html|*.*|*.*"
        dlgThis.CancelError = True
        On Error GoTo out
        dlgThis.ShowSave
        strHTML = dlgThis.Filename
        If gobjFSO.FileExists(strHTML) Then
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then GoTo out
        End If
        If Me.Document.ExportToHTML(Me.Editor1, strHTML) Then
            MsgBox "�ɹ�����ΪHTML�ļ���" & vbCrLf & "�ļ���:" & strHTML, vbOKOnly + vbInformation, gstrSysName
        End If
    Case ID_FILE_IMPORTFROMXML
        '��XML�ļ�����
        Dim strXML As String
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error GoTo out
        dlgThis.ShowOpen
        strXML = dlgThis.Filename
        If gobjFSO.FileExists(strXML) Then
            If Me.Document.EPRPatiRecInfo.ǩ������ > cprSL_�հ� Or Me.Document.EPRPatiRecInfo.���汾 > 1 Then
                MsgBox "ֻ��������д�����ļ�����ʱ����XML���뵼��������", vbOKOnly + vbInformation, gstrSysName
                GoTo out
            End If
            If Me.Document.Signs.Count > 0 And _
                (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                MsgBox "����������ǩ�����ļ��н���XML���������", vbOKOnly + vbInformation, gstrSysName
                GoTo out
            Else
                Call AddUndoPoint  '�ֶ�����
                If Me.Document.ImportFromXMLFile(Me.Editor1, strXML) Then
                    Me.Document.Compends.UpdateOrdersFromText Me.Editor1
                    Me.Document.Compends.FillTree mfrmCompends.Tree
                End If
                Call ClearNoUseUndoList
            End If
            Call RecountPage(True)
        End If
    Case ID_FILE_PAGESETUP
        Editor1.ShowPageSetupDlg
        Call RecountPage(True)
    Case ID_FILE_PRINTPREVIEW
        blnModified = Me.Editor1.Modified
        If AskOutputMode(blnOrigMode, False) Then
            Call PrintEPRDoc(True, Not blnOrigMode)
        End If
        Me.Editor1.Modified = blnModified
    Case ID_FILE_PRINT
        blnModified = Me.Editor1.Modified
        If AskOutputMode(blnOrigMode, False) Then
            Call PrintEPRDoc(False, Not blnOrigMode)
        End If
        Me.Editor1.Modified = blnModified
    Case ID_FILE_PRINTINWORD
        Call PrintInWord
    Case ID_FILE_EXIT, ID_COMMON_CANCEL
        mblnPrecess = False: Unload Me: Exit Sub
    Case ID_EDIT_UNDO
'        Editor1.Undo
        Call Undo
        Call RecountPage
    Case ID_EDIT_REDO
'        Editor1.Redo
        Call Redo
        Call RecountPage
    Case ID_EDIT_CUT
        gstrCopyPID = CStr(Document.EPRPatiRecInfo.����ID)
        Call ExecCut
        Call RecountPage
    Case ID_EDIT_COPY
        If Control.Enabled And Control.Visible Then '��ݼ�ִ��ʱ��Ҫ�ж�
            gstrCopyPID = CStr(Document.EPRPatiRecInfo.����ID)
            If Me.ActiveControl Is edtThis Then
                edtThis.Copy    '�������ı���ʽ�������������򣨷ŵ������壩
                Call ExecCopy(Me.edtThis)   '�������ݣ��ؼ���δ������
            Else
                Editor1.Copy    '�������ı���ʽ�������������򣨷ŵ������壩
                Call ExecCopy(Me.Editor1)   '�������ݣ��ؼ���δ������
            End If
            Call RecountPage
        End If
    Case ID_EDIT_COPYSELF  'ר�ø���
            Call SpicalCopy(Control.Enabled, Control.Visible)
    Case ID_EDIT_COPYOUT  '���Ƶ�ճ����
        Editor1.Copy
        gstrCopyPID = CStr(Document.EPRPatiRecInfo.����ID)
    Case ID_EDIT_SAVEASPHRASE
        '��Ϊʾ���ʾ�
        Call ExecSaveAsPhrase
    Case ID_EDIT_PASTE
        If Control.Enabled And Control.Visible Then '��ݼ�ִ��ʱ��Ҫ�ж�
            If Control.Parent Is Nothing Then
                'Control.ParentΪ�ձ�ʾ�ɰ��ȼ�����
                lngVkState = GetAsyncKeyState(Asc("V"))
                
                'lngVkState���Ϊ0��ʾCtrl+v��v��û�б����£������Ʒ��ʹ����HooK�����ɿ�v��ʱ��Ҳ�ᴥ�����ȼ�
                If lngVkState = 0 Then
                    GoTo out
                End If
            End If
            
            If gstrCopyPID <> "" And gstrCopyPID <> CStr(Document.EPRPatiRecInfo.����ID) And Document.EPRFileInfo.���� <> cpr���Ʊ��� And InStr(gstrPrivsEpr, "�������˲���") <= 0 Then
                MsgBox "�������ò������������ݣ���ֹ�������˲�����", vbExclamation, gstrSysName
                gstrCopyPID = ""
                On Error Resume Next
                Clipboard.Clear
                gfrmPublic.edtPublic.NewDoc: Set gfrmPublic.Elements = New cEPRElements
                GoTo out
            End If
            Call ExecAuditDelete
            Call ExecPaste(Me.Editor1)   'ճ�����ݣ������ؼ��֣�
            Call RecountPage
        End If
    Case ID_EDIT_DELETE
        If Editor1.ViewMode = cprNormal Then
            Call ExecDelete
            Call RecountPage
        End If
    Case ID_EDIT_BACKSPACE
        If Editor1.ViewMode = cprNormal Then
            Call ExecBackSpace
            Call RecountPage
        End If
    Case ID_EDIT_SELECTALL
        Editor1.SelectAll
    Case ID_EDIT_FIND
        Editor1.ShowFindReplaceDlg 0
    Case ID_EDIT_FINDNEXT
        Editor1.FindNext
    Case ID_EDIT_REPLACE
        Editor1.ShowFindReplaceDlg IIf(Me.Editor1.AuditMode, -1, 1)
        Call RecountPage(True)
    Case ID_VIEW_STRUCTURE
        If mfrmCompends.Visible Then
            DkpThis.FindPane(ID_VIEW_STRUCTURE).Close
        Else
            DkpThis.ShowPane ID_VIEW_STRUCTURE
        End If
    Case ID_VIEW_PHRASEDEMO
        If mfrmSentenceDetailed.Visible Then
            DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Close
        Else
            DkpThis.ShowPane ID_VIEW_PHRASEDEMO
        End If
    Case ID_VIEW_SEGMENT
        If mfrmSegments.Visible Then
            DkpThis.FindPane(ID_VIEW_SEGMENT).Close
        Else
            DkpThis.ShowPane ID_VIEW_SEGMENT
        End If
    Case ID_VIEW_PACSPIC
        If mfrmPacsPic.Visible Then
            DkpThis.FindPane(ID_VIEW_PACSPIC).Close
        Else
            DkpThis.ShowPane ID_VIEW_PACSPIC
        End If
    Case ID_VIEW_HISTORYREPORT
        If mfrmHistoryReport.Visible Then
            DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Close
        Else
            DkpThis.ShowPane ID_VIEW_HISTORYREPORT
        End If
    Case ID_VIEW_HISTORYWINDOW
        
        picPane.Visible = Not picPane.Visible
        Call picHistoryInfo_Resize
        
        cbrThis.RecalcLayout
        
    Case ID_EDIT_REFCOMPEND
        Me.Document.Compends.UpdateOrdersFromText Me.Editor1
        Me.Document.Compends.FillTree mfrmCompends.Tree
    Case ID_EDIT_ADDCOMPEND
        If Editor1.ViewMode = cprNormal Then
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
                Call AddUndoPoint  '�ֶ�����
                Dim f_InsCompend As New frmInsCompend
                f_InsCompend.ShowMe Me, Me.Editor1, Me.Document.Compends
                Call ClearNoUseUndoList
                Call RecountPage(True)
            End If
        End If
    Case ID_EDIT_MODCOMPEND
        If Not mfrmCompends.Tree.SelectedItem Is Nothing Then
            lKey = mfrmCompends.Tree.SelectedItem.Tag
            If lKey > 0 Then
                If Document.Compends("K" & lKey).Ԥ�����ID <> 0 And Document.EditType <> cprET_�����ļ����� Then
                    MsgBox "������༭������٣�", vbOKOnly + vbInformation, gstrSysName
                    GoTo out
                End If
                Call AddUndoPoint  '�ֶ�����
                Dim f_ModCompend As New frmInsCompend
                f_ModCompend.ShowMe Me, Me.Editor1, Me.Document.Compends, Me.Document.Compends("K" & lKey)
                Call ClearNoUseUndoList
            End If
        End If
    Case ID_EDIT_DELCOMPEND
        If Not mfrmCompends.Tree.SelectedItem Is Nothing Then
            lKey = mfrmCompends.Tree.SelectedItem.Tag
            If lKey > 0 Then
                Call AddUndoPoint  '�ֶ�����
                Me.DeleteOutline lKey
                Call ClearNoUseUndoList
                Call RecountPage(True)
            End If
        End If
    Case ID_VIEW_HEADFOOT
        Editor1.Foot = Document.EPRFileInfo.ҳ��
        Editor1.Head = Document.EPRFileInfo.ҳü
        If Editor1.ShowHeadFootDlg Then
            Document.EPRFileInfo.ҳ�� = Editor1.Foot
            Document.EPRFileInfo.ҳü = Editor1.Head
            If Me.Editor1.ViewMode = cprPaper Then
                '���·�ҳ
                Me.Editor1.Freeze
                Me.Editor1.ViewMode = cprNormal
                Me.Editor1.ViewMode = cprPaper
                Me.Editor1.UnFreeze
            End If
            Me.Editor1.Modified = True
            Call RecountPage(True)
        End If
    Case ID_VIEW_CHARCOUNT
        Editor1.ShowCharCountDlg
    Case ID_VIEW_RULER
        Control.Checked = Not Control.Checked
        Editor1.ShowRuler = Control.Checked
    Case ID_VIEW_PENWINDOW
        If picPenInput.Visible Then
            picPenInput.Visible = False
        Else
            picPenInput.Visible = True
            If txtPenInput.Visible And txtPenInput.Enabled Then txtPenInput.SetFocus
        End If
    Case ID_INSERT_DATETIME
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        sText = Editor1.ShowInsertDateTimeDlg(, , , False)
                        If sText <> "" Then
                            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = sText
                            tblThis.Modified = True
                            tblThis.Refresh False, True, tblThis.SelectedCellKey
                            tblThis_Resize tblThis.Width, tblThis.Height
                        End If
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Call AddUndoPoint  '�ֶ�����
                    Editor1.ShowInsertDateTimeDlg
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_DATE
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Format(Now, "YYYY��MM��DD��")
                        tblThis.Modified = True
                        tblThis.Refresh False, True, tblThis.SelectedCellKey
                        tblThis_Resize tblThis.Width, tblThis.Height
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
                    Call AddUndoPoint  '�ֶ�����
                    Editor1.ForceEdit = True
                    Editor1.Tag = "cbrThis_ExeCute"
                    If Me.Editor1.AuditMode Then
                        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
                        '���������ԣ����������ı���
                        On Error Resume Next
                        Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
                        Me.Editor1.OriginRTB.SelStrikeThru = False
                    End If
                    Editor1.Selection.Text = Format(Now, "YYYY��MM��DD��")
                    Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
                    Me.Editor1.ForceEdit = blnForce
                    Editor1.Tag = ""
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_TIME
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Format(Now, "HHʱmm��")
                        tblThis.Modified = True
                        tblThis.Refresh False, True, tblThis.SelectedCellKey
                        tblThis_Resize tblThis.Width, tblThis.Height
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False And Editor1.Selection.Font.Protected = False Then
                    Call AddUndoPoint  '�ֶ�����
                    Editor1.ForceEdit = True
                    Editor1.Tag = "cbrThis_ExeCute"
                    If Me.Editor1.AuditMode Then
                        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
                        '���������ԣ����������ı���
                        On Error Resume Next
                        Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
                        Me.Editor1.OriginRTB.SelStrikeThru = False
                    End If
                    Editor1.Selection.Text = Format(Now, "HHʱmm��")
                    Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
                    Me.Editor1.ForceEdit = blnForce
                    Editor1.Tag = ""
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_SPECIALCHAR
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False Then
                        sText = Editor1.ShowInsertSymbolDlg(False, IIf(InStr(mstrSex, "��") > 0, 1, IIf(InStr(mstrSex, "Ů") > 0, 2, 0)), True, strPic, objPic)
                        If sText = "" Then GoTo out
                        If sText <> "" Then
                            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = sText
                            tblThis.Modified = True
                            tblThis.Refresh False, True, tblThis.SelectedCellKey
                            tblThis_Resize tblThis.Width, tblThis.Height
                        End If
                    End If
                End If
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Call AddUndoPoint  '�ֶ�����
                    sText = Editor1.ShowInsertSymbolDlg(True, IIf(InStr(mstrSex, "��") > 0, 1, IIf(InStr(mstrSex, "Ů") > 0, 2, 0)), False, strPic, objPic)
                    If Not objPic Is Nothing Then 'ͼƬ��ʽ
                        Editor1.Tag = "cbrThis_ExeSPECIALCHAR"
                        InsertPicture EPRFormulaPicture, objPic, objPic.Width, objPic.Height, strPic
                        Editor1.Tag = ""
                    End If
                    Call ClearNoUseUndoList
                End If
            End If
            Call RecountPage
        End If
    Case ID_INSERT_TABLE
        If Editor1.ViewMode = cprNormal Then
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then
                Call AddUndoPoint  '�ֶ�����
                Dim ScreenPoint As POINTAPI
                GetCursorPos ScreenPoint
                ShowTablePicker ScreenPoint.X * 15, ScreenPoint.y * 15
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTTABLE
        If Editor1.ViewMode = cprNormal Then
            If tblThis.Visible = False Then
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Dim frmInsTb As New frmInsTable, lR As Long, lC As Long
                    If frmInsTb.ShowMe(Me, lR, lC) Then
                        Call AddUndoPoint  '�ֶ�����

                        lKey = Me.Document.Tables.Add
                        tblThis.AutoHeight = True
                        tblThis.Redraw = False
                        tblThis.SingleClickEdit = False
                        tblThis.HighlightMode = HMFilledRectAlpha
                        tblThis.Width = Me.Editor1.PaperWidth - Me.Editor1.Selection.Para.LeftIndent - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - 800
                        tblThis.Init lR, lC
                        tblThis.CellMargin = 10
                        For i = 1 To lR
                            For j = 1 To lC
                                Me.Document.Tables("K" & lKey).Cells.Add , i, j
                            Next j
                        Next i
                        tblThis.Tag = lKey
                        tblThis.ShowToolTipText = True
                        tblThis.MinRowHeight = 300
                        tblThis.Redraw = True
                        tblThis.Refresh
                        SaveUIToTable Me.Document.Tables("K" & lKey), True

                        Call ClearNoUseUndoList
                        Call RecountPage
                    End If
                    Unload frmInsTb: Set frmInsTb = Nothing
                End If
            End If
        End If
    Case ID_INSERT_PICTURE
        If Editor1.ViewMode = cprNormal Then
            Dim frmInsertPic  As New frmInsertPicture
            If mbEditInTable Then
                frmInsertPic.ShowMe Me
            ElseIf ucPacsImgCanvas1.Visible Then
                frmInsertPic.ShowMe Me
            Else
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then
                    Call AddUndoPoint  '�ֶ�����
                    frmInsertPic.ShowMe Me, lngMaxWidth:=(Me.Editor1.PaperWidth - Me.Editor1.MarginLeft - Me.Editor1.MarginRight) / Screen.TwipsPerPixelX, lngMaxHeight:=(Me.Editor1.PaperHeight - Me.Editor1.MarginTop - Me.Editor1.MarginBottom) / Screen.TwipsPerPixelY
                    Call ClearNoUseUndoList
                End If
            End If
        End If
    Case ID_INSERT_ELEMENT
        If mbEditInTable Then
            If tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing Then
                lKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                If lKey > 0 Then
                    mfrmInsElement.Tag = lKey
                End If
                If mfrmInsElement.Tag = "" Then
                    '����
                    mfrmInsElement.ShowMe Me
                Else
                    '�޸�
                    If Val(tblThis.Tag) > 0 Then mfrmInsElement.ShowMe Me, Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lKey), True
                End If
            End If
        Else
            bBeteenKeys = IsBetweenAnyKeys(Me.Editor1, Me.Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            Call AddUndoPoint  '�ֶ�����
            If bBeteenKeys Then
                '�޸�Ҫ��
                If sKeyType = "E" Then
                    If Me.Document.EditType <> cprET_�����ļ����� And Document.Elements("K" & lKey).�������� And InStr(1, gstrPrivsEpr, "�����ı�����") = 0 Then
                        '�Ƕ���ģʽ�£��������޸ı���������Ҫ�أ�
                        MsgBox "�����޸ı���������Ҫ�أ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                    If Document.Elements("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                        mfrmInsElement.Tag = lKey
                        mfrmInsElement.ShowMe Me, Me.Document.Elements("K" & lKey), True, (Me.Document.EditType = cprET_�����ļ�����)
                    End If
                End If
            Else
                '����Ҫ��
                If mfrmInsElement.Visible Then
                    mfrmInsElement.Hide
                Else
                    mfrmInsElement.ShowMe Me, , , (Me.Document.EditType = cprET_�����ļ�����)
                End If
            End If
            Call ClearNoUseUndoList
        End If
    Case ID_INSERT_EPRDEMO
        '��������
        Dim f_EPRDemo As New frmImportEPRDemo, lngEPRDemoID As Long
        lngEPRDemoID = f_EPRDemo.ShowMe(Me)
        If lngEPRDemoID > 0 Then
            Call AddUndoPoint  '�ֶ�����
            Me.Document.ImportEPRDemo Me.Editor1, lngEPRDemoID
            Call ClearNoUseUndoList
            Call RecountPage(True)
        End If
    Case ID_INSERT_DOCADVISE
        Call AddUndoPoint  '�ֶ�����
        Call ImportDocAdvice
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_INSERT_PACSPIC
        Call AddUndoPoint  '�ֶ�����
        Call InsertPacsPicTable
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_STYLEWINDOW
        '��ʽ����
        If mfrmStyleMan.Visible Then
            DkpThis.FindPane(ID_FORMAT_STYLEWINDOW).Close
        Else
            DkpThis.ShowPane ID_FORMAT_STYLEWINDOW
        End If
    Case ID_FORMAT_STYLE
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        If Control.Type = xtpControlComboBox Then
            Call AddUndoPoint  '�ֶ�����
            If Control.Text = "����..." Then
                DkpThis.ShowPane ID_FORMAT_STYLEWINDOW
            Else
                SetCommonStyle Editor1, Control.Text, Editor1.Selection.StartPos, Editor1.Selection.EndPos, True
            End If
            Call ClearNoUseUndoList
            Call RecountPage
        End If
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
    Case ID_FORMAT_FONTNAME
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            blnTmp = Not tblThis.Cell(lRow1, lCol1).FontBold
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FontName = Control.Text
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            If Control.Type = xtpControlComboBox Then
                Call AddUndoPoint  '�ֶ�����
                Editor1.Selection.Font.Name = Control.Text
                Call ClearNoUseUndoList
            End If
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
        End If
        Call RecountPage
    Case ID_FORMAT_FONTSIZE
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            blnTmp = Not tblThis.Cell(lRow1, lCol1).FontBold
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FontSize = GetFontSizeNumber(Control.Text)
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            If Control.Type = xtpControlComboBox Then
                Call AddUndoPoint  '�ֶ�����
                Editor1.Selection.Font.Size = GetFontSizeNumber(Control.Text)
                Call ClearNoUseUndoList
            End If
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
        End If
        Call RecountPage
    Case ID_FORMAT_FONT
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.ShowFontDlg 2 ^ 5 + 2 ^ 4 + 2 ^ 3 + 2 ^ 2 + 2 ^ 1 + 2 ^ 0
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_PARA
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.ShowParaDlg False
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_PROTECT
        Call AddUndoPoint  '�ֶ�����
        If Control.Checked Then
            'ȡ������
            If Editor1.Selection.Font.ForeColor <> PROTECT_FORECOLOR Then
                GoTo out
            Else
                Editor1.ForceEdit = True
                Editor1.Tag = "cbrThis_ExeCute"
                Editor1.Selection.Font.Protected = False
                Editor1.Selection.Font.ForeColor = tomAutoColor
                Me.Editor1.ForceEdit = blnForce
                Editor1.Tag = ""
            End If
        Else
            '���ñ���
            If Editor1.Selection.Font.Protected = True Or Editor1.Selection.Font.Hidden = True Or _
                Editor1.Selection.Font.BackColor <> tomAutoColor Then
                GoTo out
            Else
                Editor1.ForceEdit = True
                Editor1.Tag = "cbrThis_ExeCute"
                Editor1.Selection.Font.Protected = True
                Editor1.Selection.Font.ForeColor = PROTECT_FORECOLOR
                Me.Editor1.ForceEdit = blnForce
                Editor1.Tag = ""
            End If
        End If
        Call ClearNoUseUndoList
    Case ID_FORMAT_BOLD
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            blnTmp = Not tblThis.Cell(lRow1, lCol1).FontBold
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FontBold = blnTmp
                    tblThis.Cell(i, j).FontWeight = 0
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '�ֶ�����
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Font.Bold = Not Editor1.Selection.Font.Bold
            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
        Call RecountPage
    Case ID_FORMAT_SUPER
        If tblThis.Visible Then GoTo out
        If Editor1.AuditMode Then '���ģʽ�£�ֻ�е�ǰ�汾�����Ŀ��Ը��ģ�ԭ�е����ֲ��ܸ���
            If Not CanSetFormat Then
                MsgBox "��ǰΪ���ģʽ�����±깦��ֻ��Ӧ���ڱ���������������ݣ����顣", vbInformation, gstrSysName
                GoTo out  'ѡ��������ֻҪ��һ������ԭ�����ּ����ɱ��
            End If
        End If
            
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Font.Superscript = Not Editor1.Selection.Font.Superscript
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_SUB
        If tblThis.Visible Then GoTo out
        If Editor1.AuditMode Then '���ģʽ�£�ֻ�е�ǰ�汾�����Ŀ��Ը��ģ�ԭ�е����ֲ��ܸ���
            If Not CanSetFormat Then
                MsgBox "��ǰΪ���ģʽ�����±깦��ֻ��Ӧ���ڱ���������������ݣ����顣", vbInformation, gstrSysName
                GoTo out  'ѡ��������ֻҪ��һ������ԭ�����ּ����ɱ��
            End If
        End If
        
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Font.Subscript = Not Editor1.Selection.Font.Subscript
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_UNDERLINE, ID_FORMAT_UNDERLINE_NONE, ID_FORMAT_UNDERLINE_THIN
        Call ExecuteUnderLine(Control, blnForce)
    Case ID_FORMAT_UNDERLINE_THICK, ID_FORMAT_UNDERLINE_WAVE, ID_FORMAT_UNDERLINE_DOT
        Call ExecuteUnderLine(Control, blnForce)
    Case ID_FORMAT_UNDERLINE_DASH, ID_FORMAT_UNDERLINE_DASHDOT, ID_FORMAT_UNDERLINE_DASHDOT2
        Call ExecuteUnderLine(Control, blnForce)
    Case ID_FORMAT_ALIGNLEFT
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '�ֶ�����
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Para.Alignment = cprHALeft
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If

            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
    Case ID_FORMAT_ALIGNCENTER
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '�ֶ�����
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Para.Alignment = cprHACenter
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If

            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
    Case ID_FORMAT_ALIGNRIGHT
        If tblThis.Visible Then
            If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            Call AddUndoPoint  '�ֶ�����
            Editor1.ForceEdit = True
            Editor1.Tag = "cbrThis_ExeCute"
            Editor1.Selection.Para.Alignment = cprHARight
            If tblThis.Visible Then
                tblThis.SetFocus
            Else
                If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
            End If

            Me.Editor1.ForceEdit = blnForce
            Editor1.Tag = ""
            Call ClearNoUseUndoList
        End If
    Case ID_FORMAT_LISTNONE
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListType = cprLTNone
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTLCHAR
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 25
        Editor1.Selection.Para.ListType = cprLTNumberAsLCLetter
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTUCHAR
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 25
        Editor1.Selection.Para.ListType = cprLTNumberAsUCLetter
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTLROME
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 30
        Editor1.Selection.Para.ListType = cprLTNumberAsLCRoman
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTUROME
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.ListStart = 1
        Editor1.Selection.Para.ListTab = 30
        Editor1.Selection.Para.ListType = cprLTNumberAsUCRoman
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTSETUP
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.ShowItemNumberDlg
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_SPACEBEFORE
        Dim l1 As Single
        l1 = Val(InputBox("�����ǰ����ֵ����λ������", gstrSysName, "0"))
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Call AddUndoPoint  '�ֶ�����
        Editor1.Selection.Para.SpaceBefore = l1
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_SPACEAFTER
        Dim l2 As Single
        l2 = Val(InputBox("�����ǰ����ֵ����λ������", gstrSysName, "0"))
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Call AddUndoPoint  '�ֶ�����
        Editor1.Selection.Para.SpaceAfter = l2
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_FIRSTINDENT
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.SetIndents 21, 0, 0
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_FIRSTHUNGING
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.SetIndents -21, 21, 0
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_INDENTDECREASE
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.LeftIndent = IIf(Editor1.Selection.Para.LeftIndent - 21 <= 0, 0, Editor1.Selection.Para.LeftIndent - 21)
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_INDENTINCREASE
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        Editor1.Selection.Para.LeftIndent = IIf(Editor1.Selection.Para.LeftIndent + 21 >= 300, 300, Editor1.Selection.Para.LeftIndent + 21)
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTARABIC
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        If Control.Checked Then
            Editor1.Selection.Para.ListType = cprLTNone
        Else
            Editor1.Selection.Para.ListStart = 1
            Editor1.Selection.Para.ListTab = 25
            Editor1.Selection.Para.ListType = cprLTNumberAsArabic
        End If
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LISTBULLETS
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        If Control.Checked Then
            Editor1.Selection.Para.ListType = cprLTNone
        Else
            Editor1.Selection.Para.ListTab = 12
            Editor1.Selection.Para.ListType = cprLTBullet
        End If
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
        Call RecountPage
    Case ID_FORMAT_LINESPACE, ID_FORMAT_LINESPACE1, ID_FORMAT_LINESPACE2, ID_FORMAT_LINESPACE3
        Call ExecuteLineSpace(Control, True)
    Case ID_FORMAT_LINESPACE4, ID_FORMAT_LINESPACE5, ID_FORMAT_LINESPACE6, ID_FORMAT_LINESPACE7
        Call ExecuteLineSpace(Control, True)
    Case ID_FORMAT_HIGHLIGHT
        Call AddUndoPoint  '�ֶ�����
        Editor1.ForceEdit = True
        Editor1.Tag = "cbrThis_ExeCute"
        ColorHighlight_pOK
        Me.Editor1.ForceEdit = blnForce
        Editor1.Tag = ""
        Call ClearNoUseUndoList
    Case ID_TABLE_CELLALIGNMENT
        Debug.Print "ID_TABLE_CELLALIGNMENT"

    Case ID_DRAW_FILLCOLOR
        ColorFillColor_pOK
    Case ID_FORMAT_FORECOLOR
        ColorForeColor_pOK
    'Public Const ID_HELP_CONTENT = 500
    'Public Const ID_HELP_ASSISTANT = 501
    'Public Const ID_HELP_CONTACT = 502
    'Public Const ID_HELP_ONLINE = 503
    'Public Const ID_HELP_ABOUT = 504
    Case ID_HELP_CONTENT
        ShowHelp App.ProductName, Me.hwnd, "frmMain", Int((glngSys) / 100)
    Case ID_HELP_CONTACT
        Call zlMailTo(Me.hwnd)
    Case ID_HELP_ONLINE
        Call zlHomePage(Me.hwnd)
    Case ID_HELP_WEBFORUM
        Call zlWebForum(Me.hwnd)
    Case ID_HELP_ABOUT
        ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
    Case ID_EDIT_FORMATBRUSH
        If mblnFmtBrushDown = False Then
            Call AddUndoPoint  '�ֶ�����
            mblnFmtBrushDown = True
            Me.Editor1.OriginRTB.MousePointer = 99
            Me.Editor1.OriginRTB.MouseIcon = picPatiInfo.MouseIcon
            Dim lS As Long, lE As Long
            With Editor1
                lS = .Selection.StartPos
                lE = .Selection.EndPos
                If lE > lS + 1 Then
                    If .Range(lE - 2, lE) = vbCrLf Then
                        '�����������
                        Set mParaFmt = New cParaFormat
                        Set mFontFmt = New cFontFormat
                        Set mParaFmt = Editor1.Range(lS, lS + 1).Para.GetParaFmt
                        Set mFontFmt = Editor1.Range(lS, lS + 1).Font.GetFontFmt
                    Else
                        'ֻ������������
                        Set mParaFmt = Nothing
                        Set mFontFmt = New cFontFormat
                        Set mFontFmt = Editor1.Range(lS, lS + 1).Font.GetFontFmt
                    End If
                Else
                    'ֻ������������
                    Set mParaFmt = Nothing
                    Set mFontFmt = New cFontFormat
                    Set mFontFmt = Editor1.Range(lS, lS + 1).Font.GetFontFmt
                End If
            End With
            Call ClearNoUseUndoList
            Call RecountPage
        Else
            Me.Editor1.OriginRTB.MousePointer = 0
            mblnFmtBrushDown = False
            Set mParaFmt = Nothing
            Set mFontFmt = Nothing
        End If
    Case ID_INSERT_AUTORECOGNISE
        '�Զ�ʶ������Ҫ�ػ����ֵ���Ŀ
        Dim strAuto As String
        If tblThis.Visible Then
            If Val(tblThis.Tag) > 0 Then
                If tblThis.InEdit Then tblThis.EndEdit
                strAuto = Trim(tblThis.Cells("K" & tblThis.SelectedCellKey).Text)
                If strAuto = "" Then GoTo out
                If Len(strAuto) > 100 Then strAuto = Left(strAuto, 100)
                ShowAutoRecSelector strAuto
            End If
        Else
            strAuto = Trim(Me.Editor1.SelText)
            If strAuto = "" Then GoTo out
            If Len(strAuto) > 100 Then strAuto = Left(strAuto, 100)
            Call AddUndoPoint  '�ֶ�����
            ShowAutoRecSelector strAuto
            Call ClearNoUseUndoList
        End If
        Call RecountPage
    Case ID_EDIT_MARKEDPIC
        If tblThis.Visible Then
            If Val(tblThis.Tag) > 0 Then
                '���ͼ
                lKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                If lKey > 0 Then
'                    Dim frmPictureEditor1 As New frmPictureEditor
'                    If frmPictureEditor1.ShowMe(Me, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey), True) Then
'                        '������ͼƬ�������
'                        Set tblThis.Cells("K" & tblThis.SelectedCellKey).Picture = Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey).DrawFinalPic
'                        tblThis.Modified = True
'                    End If
                    '�༭ͼƬ
                    Dim LL As Long, lT As Long, lW As Long, lH As Long
                    tblThis.Cells("K" & tblThis.SelectedCellKey).GetCellPictureBorder LL, lT, lW, lH
                    ucPictureEditor1.ShowMe Me, tblThis.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey), _
                        LL, lT, lW, lH, True, Me.Document.Tables("K" & tblThis.Tag)
                End If
            End If
        Else
            If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And (Me.Editor1.AuditMode = False) Then
                '���ҹؼ��� ID ��
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then GoTo out
                If sKeyType = "P" Then
                    Call AddUndoPoint  '�ֶ�����
'                    Dim frmPictureEditor2 As New frmPictureEditor
'                    frmPictureEditor2.ShowMe Me, Me.Document.Pictures("K" & lKey)
                    '�༭ͼƬ
                    Editor1.ShowUIInterface
                    ucPictureEditor1.ShowMe Me, Editor1.hwnd, cbrThis, Me.Document.Pictures("K" & lKey), _
                        Editor1.UILeft, Editor1.UITop, Editor1.UIWidth, Editor1.UIHeight, False

                    Call ClearNoUseUndoList
                End If
            End If
        End If
    Case ID_EDIT_OUTERPIC
        If tblThis.Visible Then
            If Val(tblThis.Tag) > 0 Then
                '���ͼ
                lKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
                If lKey > 0 Then
                    cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey).OrigPic, _
                        lKey, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey).��������, Me, False
                    '�����ⲿͼƬ�ı�����cPicEditor�����pOK�¼��д���
                End If
            End If
        Else
            If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And (Me.Editor1.AuditMode = False) Then
                '���ҹؼ��� ID ��
                bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                If bBeteenKeys = False Then GoTo out
                If sKeyType = "P" Then
                    Call AddUndoPoint  '�ֶ�����
                    cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Pictures("K" & lKey).OrigPic, lKey, Me.Document.Pictures("K" & lKey).��������, Me, False
                    Call ClearNoUseUndoList
                End If
            End If
        End If
    Case ID_PATISIGN
        Call PatiSign(Control)
    Case ID_SIGN, ID_SIGN_QUIT
        If AddSign Then
            Call RelateFeedback(True)
            If Control.ID = ID_SIGN_QUIT Then 'ǩ�����˳�
                mblnPrecess = False
                Unload Me
                Exit Sub
            End If
        End If
        
        Call RecountPage
    Case ID_UNTREAD
        Call DoUntread
        txtContent.Enabled = True
        Call RelateFeedback(False)
        Call RecountPage
    Case ID_ELEMENT_UPDATE                  '����Ҫ��

        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False Then GoTo out
        If sKeyType = "E" Then
        
            Call AddUndoPoint  '�ֶ�����
            With Me.Document.Elements("K" & lKey)
                If .�滻�� = 1 Then
                                        
                    .�����ı� = GetReplaceEleValue(.Ҫ������, Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID, Me.Document.EPRPatiRecInfo.������Դ, Me.Document.EPRPatiRecInfo.ҽ��id, Me.Document.EPRPatiRecInfo.Ӥ��)
                    .Refresh Me.Editor1
                    
                    If .�Զ�ת�ı� Then
                        Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lKey), False      '�Զ�ת��Ϊ���ı�����ʱ��ɾ����Ҫ�أ�
                    End If
                End If
            End With
            Call ClearNoUseUndoList
        End If
        
        Call RecountPage
        
    Case ID_ELEMENT_CLEAR
    
        '���Ҫ��
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys = False Then GoTo out
        If sKeyType = "E" Then
            Call AddUndoPoint  '�ֶ�����
            Me.Document.Elements("K" & lKey).�����ı� = ""
            Me.Document.Elements("K" & lKey).Refresh Me.Editor1
            Call ClearNoUseUndoList
        End If
        Call RecountPage
        
    Case ID_ELEMENT_TOSTRING
        'Ҫ��ת��Ϊ���ı�
        If MsgBox("�Ƿ񽫸�����Ҫ��ת��Ϊ�����ṹ����Ϣ�Ĵ��ı���", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then GoTo out
            If sKeyType = "E" Then
                Call AddUndoPoint  '�ֶ�����
                Dim str���� As String
                str���� = IIf(Me.Document.Elements("K" & lKey).�����ı� = "", "  ", Me.Document.Elements("K" & lKey).�����ı�)
                lngLen = Len(str����)
                With Me.Editor1
                    .Freeze
                    .ForceEdit = True
                    .Tag = "cbrThis_ExeCute"
                    .Range(lKSS, lKEE) = str����
                    .Range(lKSS, lKSS + lngLen).Font.Protected = False
                    .Range(lKSS, lKSS + lngLen).Font.Hidden = False
                    .Range(lKSS, lKSS + lngLen).Font.BackColor = tomAutoColor
                    .Range(lKSS, lKSS + lngLen).Font.Underline = cprNone
                    .ForceEdit = False
                    .Tag = ""
                    .UnFreeze
                End With
                Me.Document.Elements.Remove "K" & lKey
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_REVISION_PREV
        Call GotoPrevRevision
    Case ID_REVISION_NEXT
        Call GotoNextRevision
    Case ID_REVISION_RESET
        Call ResetRevision
'        Me.Editor1.ResetAuditText
    Case ID_DIAGNOSIS
        '���
        Call AddUndoPoint  '�ֶ�����
        Call AddDiagnosis
        Call ClearNoUseUndoList
        Call RecountPage
    Case conMenu_Tool_Reference
        '��ϲο�
        bFinded = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "D", lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            Call Me.Document.Event_ClickDiagRef(Me.Document.Diagnosises("K" & lKey).���id, vbModal)
        End If
    Case ID_INSERT_TABLE
        Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.Selection.Font.Protected = False And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
        Call RecountPage
    Case ID_TABLE_CELLALIGNMENT1
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                    tblThis.Cell(i, j).VAlignment = VALignTop
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT2
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                    tblThis.Cell(i, j).VAlignment = VALignTop
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT3
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                    tblThis.Cell(i, j).VAlignment = VALignTop
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT4
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                    tblThis.Cell(i, j).VAlignment = VALignVCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT5
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                    tblThis.Cell(i, j).VAlignment = VALignVCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT6
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                    tblThis.Cell(i, j).VAlignment = VALignVCentre
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT7
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignLeft
                    tblThis.Cell(i, j).VAlignment = VALignBottom
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT8
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignCentre
                    tblThis.Cell(i, j).VAlignment = VALignBottom
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CELLALIGNMENT9
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).HAlignment = HALignRight
                    tblThis.Cell(i, j).VAlignment = VALignBottom
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False, False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_CURRENCY
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FormatString = "��#,0.00"
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_PERCENT
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FormatString = "0.0%"
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_KILOBIT
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.Visible Then
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    tblThis.Cell(i, j).FormatString = "#,0.00"
                    tblThis.Cell(i, j).HAlignment = HALignRight
                Next j
            Next i
            tblThis.Modified = True
            tblThis.Refresh False
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    Case ID_TABLE_MERGE
        If tblThis.Visible Then
            If Control.Checked Then
                tblThis.DisMergeCells tblThis.Row, tblThis.Col
            Else
                tblThis.MergeSelectedCells
            End If
            tblThis.Modified = True
            tblThis.Refresh False
            '����UI����
            tblThis_Resize tblThis.Width, tblThis.Height
            '���汳��ͼƬ
            If Val(tblThis.Tag) <= 0 Then GoTo out
            If tblThis.Modified Then SaveUIToTable Me.Document.Tables("K" & tblThis.Tag)
            Call RecountPage
        End If
    Case ID_TABLE_CELLPROTECTED
        If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then GoTo out
        If tblThis.SelectedCellKey > 0 Then
            tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = Not tblThis.Cells("K" & tblThis.SelectedCellKey).Protected
        End If
        tblThis.Refresh False, False
    Case ID_TABLE_PROPERTY
        If tblThis.Visible Then
            tblThis.ShowProperty Me, tblThis
        End If
    Case ID_TABLE_INSERTCOLLEFT
        If tblThis.Visible Then
            If tblThis.Col > 0 And Val(tblThis.Tag) > 0 Then
                '��������ϲ���Ԫ����ô���������
                For i = 1 To tblThis.RowCount
                    If Len(tblThis.Cell(i, tblThis.Col).MergeInfo) > 0 Then
                        MsgBox "��Ϊ���а����ϲ���Ԫ�����Բ�������������У�����ȡ���ϲ������ԣ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '�ֶ�����
                '�ڱ��ؼ��в���հ���
                Me.Document.Tables("K" & tblThis.Tag).InsertCol tblThis.Col - 1
                tblThis.InsertCol tblThis.Col - 1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTCOLRIGHT
        If tblThis.Visible Then
            If tblThis.Col > 0 And Val(tblThis.Tag) > 0 Then
                '��������ϲ���Ԫ����ô���������
                For i = 1 To tblThis.RowCount
                    If Len(tblThis.Cell(i, tblThis.Col).MergeInfo) > 0 Then
                        MsgBox "��Ϊ���а����ϲ���Ԫ�����Բ�������������У�����ȡ���ϲ������ԣ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '�ֶ�����
                '�ڱ��ؼ��в���հ���
                Me.Document.Tables("K" & tblThis.Tag).InsertCol tblThis.Col
                tblThis.InsertCol tblThis.Col
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTROWUP
        If tblThis.Visible Then
            If tblThis.Row > 0 And Val(tblThis.Tag) > 0 Then
                '��������ϲ���Ԫ����ô���������
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "��Ϊ���а����ϲ���Ԫ�����Բ�������������У�����ȡ���ϲ������ԣ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '�ֶ�����
                '�ڱ��ؼ��в���հ���
                Me.Document.Tables("K" & tblThis.Tag).InsertRow tblThis.Row - 1
                tblThis.InsertRow tblThis.Row - 1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTROWDOWN
        If tblThis.Visible Then
            If tblThis.Row > 0 And Val(tblThis.Tag) > 0 Then
                '��������ϲ���Ԫ����ô���������
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "��Ϊ���а����ϲ���Ԫ�����Բ�������������У�����ȡ���ϲ������ԣ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '�ֶ�����
                '�ڱ��ؼ��в���հ���
                Me.Document.Tables("K" & tblThis.Tag).InsertRow tblThis.Row
                tblThis.InsertRow tblThis.Row
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_INSERTINHERITROW
        '����̳���
        Dim lTag As Long
        If tblThis.Visible Then
            If tblThis.Row > 0 And Val(tblThis.Tag) > 0 Then
                '��������ϲ���Ԫ����ô���������
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "��Ϊ���а����ϲ���Ԫ�����Բ�������������У�����ȡ���ϲ������ԣ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '�ֶ�����
                '���ڱ��ؼ��в���հ���
                lRow1 = tblThis.Row
                Me.Document.Tables("K" & tblThis.Tag).InsertRow lRow1
                tblThis.InsertRow lRow1
                'Ȼ������һ������
                For i = 1 To tblThis.ColCount
                    With tblThis.Cell(lRow1 + 1, i)
                        .Margin = tblThis.Cell(lRow1, i).Margin
                        .SingleLine = tblThis.Cell(lRow1, i).SingleLine
                        .Visibled = tblThis.Cell(lRow1, i).Visibled
                        .Width = tblThis.Cell(lRow1, i).Width
                        .Height = tblThis.Cell(lRow1, i).Height
                        .FixedWidth = tblThis.Cell(lRow1, i).FixedWidth
                        .AutoHeight = tblThis.Cell(lRow1, i).AutoHeight
                        .Icon = tblThis.Cell(lRow1, i).Icon
                        .Text = tblThis.Cell(lRow1, i).Text

'                        .Tag = tblThis.Cell(lRow1, i).Tag
                        If Val(tblThis.Cell(lRow1, i).Tag) > 0 Then
                            If tblThis.Cell(lRow1, i).Picture Is Nothing Then
                                '����Ҫ��
                                lKey = Me.Document.Tables("K" & tblThis.Tag).Elements.AddExistNode(Me.Document.Tables("K" & tblThis.Tag).Elements("K" & tblThis.Cell(lRow1, i).Tag), False)
                                .Tag = lKey
                                Me.Document.Tables("K" & tblThis.Tag).Cell(lRow1 + 1, i).ElementKey = lKey
                            Else
                                '����ͼƬ
                                lKey = Me.Document.Tables("K" & tblThis.Tag).Pictures.AddExistNode(Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & tblThis.Cell(lRow1, i).Tag), False)
                                .Tag = lKey
                                Me.Document.Tables("K" & tblThis.Tag).Cell(lRow1 + 1, i).PictureKey = lKey
                            End If
                        End If
                        .ToolTipText = tblThis.Cell(lRow1, i).ToolTipText
                        .FormatString = tblThis.Cell(lRow1, i).FormatString
                        .Indent = tblThis.Cell(lRow1, i).Indent
                        .HAlignment = tblThis.Cell(lRow1, i).HAlignment
                        .VAlignment = tblThis.Cell(lRow1, i).VAlignment
                        .ForeColor = tblThis.Cell(lRow1, i).ForeColor
                        .BackColor = tblThis.Cell(lRow1, i).BackColor
                        .GridLineColor = tblThis.Cell(lRow1, i).GridLineColor
                        .GridLineWidth = tblThis.Cell(lRow1, i).GridLineWidth
                        .FontName = tblThis.Cell(lRow1, i).FontName
                        .FontSize = tblThis.Cell(lRow1, i).FontSize
                        .FontBold = tblThis.Cell(lRow1, i).FontBold
                        .FontItalic = tblThis.Cell(lRow1, i).FontItalic
                        .FontStrikeout = tblThis.Cell(lRow1, i).FontStrikeout
                        .FontUnderline = tblThis.Cell(lRow1, i).FontUnderline
                        .FontWeight = tblThis.Cell(lRow1, i).FontWeight
                        .Protected = tblThis.Cell(lRow1, i).Protected
                        Set .Picture = tblThis.Cell(lRow1, i).Picture
                    End With
                Next
                tblThis.Refresh
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_FORMATCELL
        If tblThis.Visible Then
            tblThis.ShowProperty Me, tblThis, 3
        End If
    Case ID_TABLE_SAMECOLWIDTH
        '��ͬ�п�
        Dim lSum As Long, lEvery As Long
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                If lCol1 <> lCol2 Then
                    For i = lCol1 To lCol2
                        lSum = lSum + tblThis.ColWidth(i)
                    Next
                    lEvery = lSum / (lCol2 - lCol1 + 1)
                    For i = lCol1 To lCol2
                        tblThis.ColWidth(i) = lEvery
                    Next
                    tblThis.Modified = True
                    tblThis.Refresh
                    tblThis_Resize tblThis.Width, tblThis.Height
                End If
            End If
        End If
    Case ID_TABLE_DELETEROW
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 And tblThis.RowCount > 1 And Val(tblThis.Tag) > 0 Then
                '��������ϲ���Ԫ����ô���������
                For i = 1 To tblThis.ColCount
                    If Len(tblThis.Cell(tblThis.Row, i).MergeInfo) > 0 Then
                        MsgBox "��Ϊ���а����ϲ���Ԫ�����Բ�����ɾ��������ȡ���ϲ������ԣ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '�ֶ�����
                lRow1 = tblThis.Row
                Me.Document.Tables("K" & tblThis.Tag).DeleteRow lRow1
                tblThis.DeleteRow lRow1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_DELETECOL
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 And tblThis.ColCount > 1 And Val(tblThis.Tag) > 0 Then
                '��������ϲ���Ԫ����ô���������
                For i = 1 To tblThis.RowCount
                    If Len(tblThis.Cell(i, tblThis.Col).MergeInfo) > 0 Then
                        MsgBox "��Ϊ���а����ϲ���Ԫ�����Բ�����ɾ��������ȡ���ϲ������ԣ�", vbOKOnly + vbInformation, gstrSysName
                        GoTo out
                    End If
                Next

                Call AddUndoPoint  '�ֶ�����
                lCol1 = tblThis.Col
                Me.Document.Tables("K" & tblThis.Tag).DeleteCol lCol1
                tblThis.DeleteCol lCol1
                tblThis.Modified = True
                Editor1.Modified = True
                mblnChange = True
                tblThis_Resize tblThis.Width, tblThis.Height
                Call ClearNoUseUndoList
                Call RecountPage
            End If
        End If
    Case ID_TABLE_DELETETABLE
        If tblThis.Visible And Val(tblThis.Tag) > 0 Then
            Call AddUndoPoint  '�ֶ�����

            lKey = Val(tblThis.Tag)
            Me.Document.Tables("K" & lKey).DeleteFromEditor Me.Editor1
            Me.Document.Tables.Remove "K" & lKey
            Editor1.CloseUIInterface
            Editor1.Modified = True
            tblThis.Visible = False
            mblnChange = True

            Call ClearNoUseUndoList
            Call RecountPage
        End If
    Case Else
        If Control.ID >= conMenu_Tool_PlugIn_Item + 1 And Control.ID <= conMenu_Tool_PlugIn_Item + 99 And Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.ExecuteFunc(glngSys, 1070, Control.Parameter, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.ID, Document.EPRFileInfo.ID)
            Err.Clear: On Error GoTo 0
        End If
    End Select
    
out: mblnPrecess = False
End Sub

'################################################################################################################
'## ���ܣ�  ��һ���޶�
'################################################################################################################
Private Sub GotoPrevRevision()
    On Error Resume Next
    Dim i As Long, lS As Long
    Dim lState1 As Long, lState2 As Long
    Dim lngStart As Long, lngEnd As Long
    Dim lng��ʼ��1 As Long, lng��ֹ��1 As Long
    Dim lng��ʼ��2 As Long, lng��ֹ��2 As Long

    With Me.Editor1
        .Freeze
        lS = .Selection.StartPos
        lState1 = Me.Document.GetTextState(Me.Editor1, lS, lS + 1, lng��ʼ��1, lng��ֹ��1) '��ȡ��ǰ�ı�״̬
        If lS = 0 Then Exit Sub
        For i = lS - 1 To 0 Step -1   'ѭ�����Ҳ�ͬ״̬�ı�
            lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng��ʼ��2, lng��ֹ��2) '��ȡ�ı�״̬
            If lState2 <> lState1 Or lng��ʼ��1 <> lng��ʼ��2 Or lng��ֹ��1 <> lng��ֹ��2 Then
                If (lng��ʼ��2 = Me.Document.Ŀ��汾 Or lng��ֹ��2 = Me.Document.Ŀ��汾 - 1) Then
                    '״̬��ͬ��
                    lState1 = lState2
                    lngEnd = i + 1
                    Exit For
                Else
                    lState1 = lState2
                    lng��ʼ��1 = lng��ʼ��2
                    lng��ֹ��1 = lng��ֹ��2
                End If
            End If
        Next
        If lngEnd > 0 Then
            For i = lngEnd - 1 To 0 Step -1
                lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng��ʼ��2, lng��ֹ��2) '��ȡ�ı�״̬
                If lState2 <> lState1 Or lng��ʼ��1 <> lng��ʼ��2 Or lng��ֹ��1 <> lng��ֹ��2 Then
                    '״̬��ͬ��
                    lState1 = lState2
                    lng��ʼ��1 = lng��ʼ��2
                    lng��ֹ��1 = lng��ֹ��2
                    lngStart = i + 1
                    Exit For
                End If
            Next
        End If
        If lngStart <> lngEnd Then .Range(lngStart, lngEnd).Selected
        .UnFreeze
    End With
End Sub

'################################################################################################################
'## ���ܣ�  ��һ���޶�
'################################################################################################################
Private Sub GotoNextRevision()
    On Error Resume Next
    Dim i As Long, lS As Long, lngLen As Long
    Dim lState1 As Long, lState2 As Long
    Dim lngStart As Long, lngEnd As Long
    Dim lng��ʼ��1 As Long, lng��ֹ��1 As Long
    Dim lng��ʼ��2 As Long, lng��ֹ��2 As Long

    With Me.Editor1
        .Freeze
        lS = .Selection.StartPos
        lngLen = Len(.Text)
        lState1 = Me.Document.GetTextState(Me.Editor1, lS, lS + 1, lng��ʼ��1, lng��ֹ��1) '��ȡ��ǰ�ı�״̬
        If lS = 0 Then Exit Sub
        For i = lS To lngLen - 1
            lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng��ʼ��2, lng��ֹ��2) '��ȡ�ı�״̬
            If lState2 <> lState1 Or lng��ʼ��1 <> lng��ʼ��2 Or lng��ֹ��1 <> lng��ֹ��2 Then
                If (lng��ʼ��2 = Me.Document.Ŀ��汾 Or lng��ֹ��2 = Me.Document.Ŀ��汾 - 1) Then
                    '״̬��ͬ��
                    lState1 = lState2
                    lng��ʼ��1 = lng��ʼ��2
                    lng��ֹ��1 = lng��ֹ��2
                    lngStart = i
                    Exit For
                Else
                    lState1 = lState2
                    lng��ʼ��1 = lng��ʼ��2
                    lng��ֹ��1 = lng��ֹ��2
                End If
            End If
        Next
        If lngStart < lngLen Then
            For i = lngStart + 1 To lngLen - 1
                lState2 = Me.Document.GetTextState(Me.Editor1, i, i + 1, lng��ʼ��2, lng��ֹ��2) '��ȡ�ı�״̬
                If lState2 <> lState1 Or lng��ʼ��1 <> lng��ʼ��2 Or lng��ֹ��1 <> lng��ֹ��2 Then
                    '״̬��ͬ��
                    lState1 = lState2
                    lng��ʼ��1 = lng��ʼ��2
                    lng��ֹ��1 = lng��ֹ��2
                    lngEnd = i
                    Exit For
                End If
            Next
        End If
        If lngStart <> lngEnd Then .Range(lngStart, lngEnd).Selected
        .UnFreeze
    End With
End Sub

'################################################################################################################
'## ���ܣ�  ��ȡ����ǩ����Դ�ı�����ȥ����дǩ����Ԥ����ٵ����������ı����ݣ�
'################################################################################################################
Public Function GetSignSourceString(ByRef edtThis As zlRichEditor.Editor) As String
    Dim lSS As Long, lSE As Long, lES As Long, lEE As Long, bNeeded As Boolean, bFinded As Boolean, lKey As Long
    Dim i As Long, strR As String, lS As Long, lE As Long, strS As String, lngLen As Long, lPos As Long
    
    edtThis.SaveDoc App.Path & "\tmp.RTF"
    gfrmPublic.edtBuff.OpenDoc App.Path & "\tmp.RTF"
    gobjFSO.DeleteFile App.Path & "\tmp.RTF"
    gfrmPublic.edtBuff.Freeze
    gfrmPublic.edtBuff.ForceEdit = True
    'ȥ������S�ؼ��ֵ�ǩ������
    lPos = 0
    Do
        bFinded = FindNextKey(gfrmPublic.edtBuff, lPos, "S", lKey, lSS, lSE, lES, lEE, bNeeded)
        If bFinded Then gfrmPublic.edtBuff.Range(lSS, lEE).Text = ""
    Loop Until bFinded = False
    gfrmPublic.edtBuff.ForceEdit = False
    gfrmPublic.edtBuff.UnFreeze
    strS = gfrmPublic.edtBuff.Text
    strS = Replace(strS, Chr(32), "")
    strS = Replace(strS, vbCr, "")
    strS = Replace(strS, vbLf, "")
    
'   '���ַ�������������ǩ��λ�û��Ҿͻ��޷���֤-----------�˶����������豣��
'    lngLen = Len(edtThis.Text)
'    If Me.Document.Signs.Count = 0 Then
'        strS = edtThis.Text
'    Else
'        For i = 1 To Me.Document.Signs.Count
'            bFinded = FindKey(edtThis, "S", Me.Document.Signs(i).Key, lSS, lSE, lES, lEE, bNeeded)
'            If bFinded Then
'                '�޳�ǩ��
'                If i = 1 Then
'                    strS = edtThis.Range(0, lSS)
'                Else
'                    strS = strS & edtThis.Range(lS, lSS)
'                End If
'                lS = lEE
'            End If
'        Next
'        If lEE < lngLen Then
'            '����ĩ���ı�
'            strS = strS & edtThis.Range(lEE, lngLen).Text
'        End If
'    End If
    GetSignSourceString = strS
End Function
Private Sub PatiSign(Control As XtremeCommandBars.ICommandBarControl)
    If Control.Caption = "����ǩ��" Then
        Call PatiDoSign(Control)
    Else
        Call PatiUnDoSign(Control)
    End If
End Sub
Private Sub PatiUnDoSign(Control As XtremeCommandBars.ICommandBarControl)
'���ܣ�������дǩ��ͼƬ
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, blnNeeded As Boolean, lCurPos As Long, intLoop As Integer
    On Error GoTo errHand
    
    If mblnPatiSign Then
        lCurPos = 1
        Do Until False
            If FindNextKey(Editor1, lCurPos, "P", lKey, lKSS, lKSE, lKES, lKEE, blnNeeded) Then
                If Document.Pictures("K" & lKey).PictureType = EPRPatiSign Then '�ҵ�ͼƬ������Ƿ�����ǩͼ
                    Exit Do
                End If
            Else 'û�ҵ��κ�ͼ
                For intLoop = 1 To Document.Pictures.Count
                    If Document.Pictures(intLoop).PictureType = EPRPatiSign Then
                        Document.Pictures.Remove (intLoop)
                        Exit Do
                    End If
                Next

                GoTo undosign
            End If
            lCurPos = lKEE
        Loop
        Call Document.Pictures("K" & lKey).DeleteFromEditor(Editor1)
    End If

undosign:
    mblnPatiSign = False
    Call RecountPage
    Control.Caption = "����ǩ��"
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function HavedPatiSign() As Boolean
'���ܣ�����Ƿ��Ѿ���ǩ��
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, blnNeeded As Boolean, lCurPos As Long
    On Error GoTo errHand
    
    lCurPos = 1
    Do Until False
        If FindNextKey(Editor1, lCurPos, "P", lKey, lKSS, lKSE, lKES, lKEE, blnNeeded) Then
            If Document.Pictures("K" & lKey).PictureType = EPRPatiSign Then '�ҵ�ͼƬ������Ƿ�����ǩͼ
                HavedPatiSign = True
                Exit Function
            End If
        Else 'û�ҵ��κ�ͼ
            Exit Function
        End If
        lCurPos = lKEE
    Loop

    HavedPatiSign = False
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub PatiDoSign(Control As XtremeCommandBars.ICommandBarControl)
'���ܣ���ȡ��дǩ��ͼƬ��ǩ����֤������Ϣ
'������strSource ǩ��Դ��
'        strName ����������ȱʡǩ����
'        strIdentifyNo �������֤�ţ�ȱʡǩ����֤����
'        strOtherParms  Ϊ�պ󲻱�����������������������ܵĲ���
'        strSignInfo ǩ��������֤��Ϣ
'        strPenSignBase64 ������дǩ��ͼƬBASE64����
'        objPenSignPic ������дǩ��ͼƬ
Dim strSource As String, strName As String, strIdentifyNo As String, strOtherParms As String, strSignInfo As String, strPenSignBase64 As String, objPenSignPic As Object
Dim blnReturn As Boolean, lWidth As Long, lHeight As Long
    On Error GoTo errHand
    strSource = GetSignSourceString(Me.Editor1)
    strName = mPatiInfor.����
    strIdentifyNo = mPatiInfor.���֤��
    strOtherParms = "" '
    blnReturn = gobjESign.PenSignature(strSource, strName, strIdentifyNo, strOtherParms, strSignInfo, strPenSignBase64, objPenSignPic)
    If blnReturn And Not objPenSignPic Is Nothing Then
        '����ͼƬ
        Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
        If Editor1.ReadOnly Then Exit Sub
        If tblThis.Visible Then
            MsgBox "��ǰλ�ò�֧�ֲ���ǩ��", vbInformation, gstrSysName: Exit Sub
        Else
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then
                Call AddUndoPoint  '�ֶ�����
                Editor1.Tag = "InsertPatiSign"
                '��ͼƬ���󱣴浽�������
                lKey = Document.Pictures.Add()
                Set Document.Pictures("K" & lKey).OrigPic = objPenSignPic
                Document.Pictures("K" & lKey).Width = lWidth
                Document.Pictures("K" & lKey).Height = lHeight
                Document.Pictures("K" & lKey).OrigWidth = lWidth
                Document.Pictures("K" & lKey).OrigHeight = lHeight
                Document.Pictures("K" & lKey).PictureType = EPRPatiSign
                Document.Pictures("K" & lKey).InsertIntoEditor Editor1
                Document.Pictures("K" & lKey).�����ı� = strName & "|" & strIdentifyNo & "|" & strSignInfo   '������Ϣ
                Editor1.Tag = ""
                Call ClearNoUseUndoList
            End If
            Editor1.SetFocus
        End If
        mblnPatiSign = True
        Call RecountPage
        Control.Caption = "���߳�ǩ"
        '�������
        '�޷�����,��������ˣ�ǩ�����޷���������
    End If
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'################################################################################################################
'## ���ܣ�  ����ǩ��
'################################################################################################################
Private Function AddSign() As Boolean
    If Me.Editor1.ViewMode <> cprNormal Then Exit Function
    Dim strTmp As String, lLen As Long, lngKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strSource As String, blnR As Boolean, picSign As StdPicture, lngPicKey As Long
    Dim frmSign As New frmEPRSign, oSign As cEPRSign
    Dim blnModified As Boolean '�Ƿ���ǩ��ǰ�Ѿ��޸�������
    Dim lS As Long, l As Long
    Dim strSQL As String, strTime As String
    
    If AutoMoveSignPos = False Then Exit Function
    If Me.Editor1.Selection.Font.Protected Then
        MsgBox "�Բ����������ڵ�ǰλ�ý���ǩ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Me.Document.�û�ǩ������ = cprSL_�հ� Then
        MsgBox "��ǰ�û���δ����ǩ������������Ա�����е���Ƹ��ְ��", vbInformation, gstrSysName: Exit Function
    End If
    For l = 1 To Document.Signs.Count
        If Document.Signs(l).ǩ������ > Me.Document.�û�ǩ������ Then
            MsgBox "��ǰ�������и��߼����ǩ������ǰǩ��������Ȩ��ǩ������", vbInformation, gstrSysName
            Exit Function
        End If
    Next
    
    If Not CheckAllObjects(True) Then Exit Function '������Ҫ��
    
            
    If Not gobjPlugIn Is Nothing Then 'ǩ��ǰ�������
        On Error Resume Next
        If Not gobjPlugIn.SignEMRBefore(glngSys, 1070, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.ID) Then Exit Function
        Err.Clear: On Error GoTo 0
    End If

    With Editor1
        '���û�б��ΰ汾��ǩ��λ�ã����ڵ�ǰλ�ò���ǩ��
        .Tag = "ǩ��"
        blnModified = .Modified
        lS = .Selection.StartPos
        
        If .SelLength > 0 And mblnǩ��Ҫ�� = False Then .Tag = "": Exit Function
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        
        If Editor1.Selection.StartPos >= lKSS And Editor1.Selection.StartPos <= lKSE Then
            Editor1.SelStart = IIf(lKSS = 0, 0, lKSS - 1)
'            Editor1.Selection.StartPos = lKSS - 1
        End If
        If .Selection.Font.Protected And mblnǩ��Ҫ�� = False Then
            .Tag = ""
            Exit Function
        End If
        
        If bBeteenKeys And mblnǩ��Ҫ�� = False Then
            .Tag = ""
            Exit Function    '��֤���ܲ���ؼ����ڲ�
        Else
            strSource = GetSignSourceString(Me.Editor1)
            Set oSign = frmSign.ShowMe(Me.Editor1, Me, strSource, picSign)
            If Not oSign Is Nothing Then
                Me.Editor1.Modified = blnModified
                If (Me.Editor1.Modified Or (Me.Editor1.AuditMode And Me.Document.EPRPatiRecInfo.ǩ������ = cprSL_�հ�)) Or Me.Editor1.AuditMode = False Then
                    oSign.��ʼ�� = Me.Document.Ŀ��汾
                Else
                    oSign.��ʼ�� = Me.Document.Ŀ��汾 - 1
                End If
                If oSign.��ʼ�� > 16 Then
                    MsgBox "Ŀǰϵͳ֧�ֵ����汾��Ϊ16������˻�����������", vbOKOnly + vbInformation, gstrSysName
                    .Tag = ""
                    Exit Function
                End If
                lngKey = Me.Document.Signs.AddExistNode(oSign)
                
                If Me.Document.Signs("K" & lngKey).InsertIntoEditor(Me.Editor1, , True, Me.Document) = True Then
                    If oSign.ǩ��ͼƬ And Not picSign Is Nothing Then
                        lngPicKey = Document.Pictures.Add()
                        Set Document.Pictures("K" & lngPicKey).OrigPic = picSign
                        Document.Pictures("K" & lngPicKey).Width = Me.ScaleX(picSign.Width, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).Height = Me.ScaleY(picSign.Height, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).OrigWidth = Me.ScaleX(picSign.Width, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).OrigHeight = Me.ScaleY(picSign.Height, vbHimetric, vbTwips)
                        Document.Pictures("K" & lngPicKey).PictureType = EPRSignPicture
                        Document.Pictures("K" & lngPicKey).��ʼ�� = oSign.��ʼ��
                        Document.Pictures("K" & lngPicKey).InsertIntoEditor Editor1
                        Call FindKey(Editor1, "S", lngKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                        If Editor1.ForceEdit = False Then Editor1.ForceEdit = True
                        Editor1.Range(lKSS, lKEE).Font.Hidden = True
                    End If
                    If oSign.ǩ����ʽ <> 2 Then Call AutoAlterSignInPage '����ǩ��λ�� ����ǩ�����������ǩ��λ��
                    Me.Editor1.Modified = blnModified
                    blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "�޸�ҳ������") > 0)
                    
    '                If oSign.ǩ����ʽ = 2 And oSign.ǩ������ = 3 And blnR Then
    '                    Dim strSign As String, lngCertID As Long, strʱ��� As String
    '                    '��ǩ�����ڣ�ǩ�������ѽ��г�ʼ�����Ա����ʼ��ʧ�ܵ������ѱ������������ֱ��ʹ�ü���
    '                    strSource = GetSignSourceFromDB(Me.Document.EPRPatiRecInfo.ID, lngKey)
    '                    On Error Resume Next
    '                    strSign = gobjESign.signature(strSource, UCase(oSign.ǩ����Ϣ), lngCertID, strʱ���) '����ǩ����Ϣ,lngCertID����ǩ��ʹ�õ�֤���¼ID
    '                    If strSign <> "" Then
    '                        oSign.ǩ����Ϣ = strSign    'Ҫ��ֵ��
    '                        oSign.֤��ID = lngCertID    '��������
    '                        oSign.ʱ��� = strʱ���    'Ҫ�ص�λ
    '                        gstrSQL = "zl_���Ӳ�������_����ǩ��(" & Me.Document.EPRPatiRecInfo.ID & "," & lngKey & ",'" & oSign.�������� & "','" & oSign.ǩ����Ϣ & "','" & oSign.ʱ��� & "')  '���ù��̱���"
    '                        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ǩ����Ϣ")
    '                    ElseIf strSign = "" Or Err.Number <> 0 Then 'ǩ��ʧ����Ҫ����
    '                        Call DoUntread(True, lngKey)
    '                        MsgBox "����ǩ��ʧ�ܣ�����ǩ�����Զ�����,������ǩ����", vbCritical, gstrSysName
    '                    End If
    '                End If

                    Call ClearUndoList      '���Undo���У�
                    DT1_EPR = Now
                    Me.Editor1.Modified = False
                    AddSign = blnR
                End If
                
                If Not gobjPlugIn Is Nothing Then 'ǩ����������
                    On Error Resume Next
                    Call gobjPlugIn.SignEMRAfter(glngSys, 1070, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.ID, oSign.����)
                    Err.Clear: On Error GoTo 0
                End If
            End If
        End If
    End With
    If mbln���޴��� And mblnFBContentChanged Then
        strTime = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "Zl_�����걨��¼_Update(" & Me.Document.EPRPatiRecInfo.ID & ",5,null,null,null,'" & gstrUserName & "'," & strTime & ",'" & Trim(txtContent.Text) & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mblnFBContentChanged = False
    End If
    If mblnIsMultiMode And Not mfrmMultiDocView Is Nothing Then
        mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
    End If
    Call SetStateInfo
    Editor1.Tag = ""
End Function
Private Sub AutoAlterSignInPage()
'������Ʊ���,�Զ�����ǩ����ҳ����λ��
'�鵽�ĵ�һ��ǩ��ǰ׷�ӻس����з�����ǩ�����������λ�Ƶ�ҳ��ֱ����ҳ��ֻ��һҳ����ʾ���棩
    If Not mblnSignAutoAlter Then Exit Sub
    If Document.EPRFileInfo.���� <> cpr���Ʊ��� Then Exit Sub
    
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim blnForce As Boolean, l As Long
    
    With Editor1
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
    
        Call Editor1.DoVirtualPrint
        If Editor1.PageCount > 1 Then Exit Sub '�����Ѿ�����һҳ������
    
        If Not FindNextKey(Editor1, 1, "S", lKey, lKSS, lKSE, lKES, lKEE, bNeeded) Then Exit Sub 'û�ҵ�ǩ�����󲻴���
        .Range(lKSS, lKSS).Font.Protected = False
        For l = 1 To 40
            .Range(lKSS, lKSS).Text = vbCrLf
            .Range(lKSS, lKSS + 2).Font.Protected = False
            Call Editor1.DoVirtualPrint
            If Editor1.PageCount > 1 Then '����һҳ��ȡ���ղ�׷�ӵ������س�����
                If .Range(lKSS, lKSS + 2) = vbCrLf Then .Range(lKSS, lKSS + 2).Text = ""
                lKSS = lKSS - 2
                If .Range(lKSS, lKSS + 2) = vbCrLf Then .Range(lKSS, lKSS + 2).Text = ""
                lKSS = lKSS - 2
                If .Range(lKSS, lKSS + 2) = vbCrLf Then .Range(lKSS, lKSS + 2).Text = ""
                Exit For
            Else
                lKSS = lKSS + 2
            End If
        Next
        
        .ForceEdit = blnForce
        .UnFreeze
    End With
    
End Sub
'################################################################################################################
'## ���ܣ�  �������
'################################################################################################################
Private Function AddDiagnosis() As Boolean
    Dim strTmp As String, lLen As Long, lngKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim strSource As String, blnR As Boolean
    Dim frmDiagnosis As New frmInsDiagnosis, oDiagnosis As cEPRDiagnosis

    With Editor1
        '�ڵ�ǰλ�ò������
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            AddDiagnosis = False: Exit Function    '��֤���ܲ���ؼ����ڲ�
        Else
            Set oDiagnosis = frmDiagnosis.ShowMe(Me.Editor1, Me)
            If Not oDiagnosis Is Nothing Then
                oDiagnosis.��ʼ�� = Me.Document.Ŀ��汾
                lngKey = Me.Document.Diagnosises.AddExistNode(oDiagnosis)
                Me.Document.Diagnosises("K" & lngKey).InsertIntoEditor Me.Editor1
            End If
            AddDiagnosis = True
        End If
    End With
End Function

'################################################################################################################
'## ���ܣ�  ȡ����ǰѡ�����ݵ��޶�
'################################################################################################################
Private Sub ResetRevision()
    '�ָ���ѡ�ı��޶�����
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, lG As Long, COLOR As OLE_COLOR
    With Me.Editor1
        .Tag = "ResetRevision"
        .InProcessing = True
        .Freeze
        .ForceEdit = True
        lS = .Selection.StartPos
        lE = .Selection.EndPos

        '�ȴ���Ҫ�غ����
        For i = lS To lE
            bFinded = FindNextAnyKey(Editor1, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then
                    '��Χ�ڴ��ڹؼ���
                    If sKeyType = "E" Then
                        'Ҫ��
                        If Me.Document.Elements("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                            '����������Ҫ�أ���ôɾ��֮��
                            .Range(lKSS, lKEE).Text = ""
                            Me.Document.Elements.Remove "K" & lKey
                            lE = lE - (lKEE - lKSS)
                            i = i - 1
                        ElseIf Me.Document.Elements("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1 Then
                            '����ɾ����Ҫ�أ���ô�ָ�֮��
                            Me.Document.Elements("K" & lKey).��ֹ�� = 0
                            Me.Document.Elements("K" & lKey).Refresh Me.Editor1
                            i = lKEE - 1
                        Else
                            '������
                            i = lKEE - 1
                        End If
                    ElseIf sKeyType = "D" Then
                        '���
                        If Me.Document.Diagnosises("K" & lKey).��ʼ�� = Me.Document.Ŀ��汾 Then
                            '����������Ҫ�أ���ôɾ��֮��
                            .Range(lKSS, lKEE).Text = ""
                            Me.Document.Diagnosises.Remove "K" & lKey
                            lE = lE - (lKEE - lKSS)
                            i = i - 1
                        ElseIf Me.Document.Diagnosises("K" & lKey).��ֹ�� = Me.Document.Ŀ��汾 - 1 Then
                            '����ɾ����Ҫ�أ���ô�ָ�֮��
                            Me.Document.Diagnosises("K" & lKey).��ֹ�� = 0
                            Me.Document.Diagnosises("K" & lKey).Refresh Me.Editor1
                            i = lKEE - 1
                        Else
                            '������
                            i = lKEE - 1
                        End If
                    Else
                        '���������Ԫ�أ�������
                        i = lKEE - 1
                    End If
                Else
                    '���򣬳�����Χ���˳�ѭ��
                    Exit For
                End If
            Else
                '�������κ�Ԫ�أ���ô�˳�ѭ��
                Exit For
            End If
        Next

        i = lS
        Do While i < lE
            If .Range(i, i + 1).Font.Protected = False And .Range(i, i + 1).Font.Hidden = False Then
                COLOR = IIf(.Range(i, i + 1).Font.ForeColor = tomAutoColor, vbBlack, .Range(i, i + 1).Font.ForeColor)
                If Me.Document.IsNewCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = False Then
                    '��һ���ַ�Ϊ�����ı�����ֱ��ɾ��֮��
                    .Range(i, i + 1) = ""
                    lE = lE - 1
                ElseIf Me.Document.IsDelCharColor(COLOR) And .Range(i, i + 1).Font.Strikethrough = True Then
                    '��һ���ַ�Ϊɾ���ı�����ָ��ı�Ϊ����ɾ���ߣ�ɾ��ǰ����ɫ����
                    lG = rgbGreen(COLOR)
                    If lG <> 0 Then
                        '��ʾ���ı���ɾ��ǰ�������ı�����ôӦ�ûָ�Ϊ����״̬
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.ForeColor = RGB(255, lG, 0)
                    Else
                        '����ָ�Ϊ��ɫ
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.ForeColor = tomAutoColor
                    End If
                    i = i + 1
                Else
                    i = i + 1
                End If
            Else
                '��Ϊ����/�����ı�����ֱ�Ӻ���һλ��
                i = i + 1
            End If
        Loop
        .InProcessing = False
        .Range(i, i).Selected
        .UnFreeze
        .Tag = ""
    End With
End Sub

'################################################################################################################
'## ���ܣ�  ȡ��ǩ�������˲�����
'################################################################################################################
Private Sub DoUntread(Optional blnImmediate As Boolean = False, Optional lSignKey As Long)
    If Me.Editor1.ViewMode <> cprNormal Then Exit Sub
    Dim lngVersion As Long, lngSignKey As Long
    Dim frmUntread As New frmEPRUntread
    Dim i As Long, lngKey As Long, lngLen As Long, COLOR As OLE_COLOR
    Dim blnForce As Boolean, lng��ʼ�� As Long
    Dim lngKeys() As Long, lngCount As Long, blnReadOnly As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    If blnImmediate Then '������ǩ��ʧ��ʱ����
        lngSignKey = lSignKey
    Else
        If frmUntread.ShowMe(Me.Document.EPRPatiRecInfo.ID, Me.Document.EditType, lngVersion, lngSignKey, Me) = False Then Exit Sub
        If lngSignKey > 0 Or lngVersion > 0 Then
            If MsgBox("ע�⣺���˲��������ɻָ����Ƿ������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    '���л��˴������±���
    On Error GoTo errHand
    If lngSignKey > 0 Then
        If Me.Document.Signs("K" & lngSignKey).ǩ����ʽ = 2 And Not blnImmediate Then
            '����ǩ����֤
            If gobjESign Is Nothing Then
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
            If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    
        '���ǩ��
        blnReadOnly = Me.Editor1.ReadOnly
        Me.Editor1.ReadOnly = False
        Editor1.Tag = "DoUntread"
        If Me.Document.Signs("K" & lngSignKey).ǩ��ͼƬ Then
            '���ǩ��ͼƬ,��ǩ��ͼƬ��ǩ����ֱ��������ϵ,ǩ���汾��û���ĵ�����²�������������ֻ����ǩ�����һ��ǩ��ͼ���
            If FindKey(Editor1, "S", Me.Document.Signs("K" & lngSignKey).Key, lKSS, lKSE, lKES, lKEE, bNeeded) Then
                If FindNextKey(Editor1, lKSS, "P", lKey, lKSS, lKSE, lKES, lKEE, bNeeded) Then
                    If Me.Document.Pictures("K" & lKey).PictureType = EPRSignPicture Then
                        Me.Document.Pictures("K" & lKey).DeleteFromEditor Me.Editor1
                        Me.Document.Pictures.Remove "K" & lKey
                    End If
                End If
            End If
        End If
        
        Me.Document.Signs("K" & lngSignKey).DeleteFromEditor Me.Editor1, Me.Document
        Me.Document.Signs.Remove "K" & lngSignKey
        Editor1.Tag = ""
        Me.Editor1.ReadOnly = blnReadOnly
    ElseIf lngVersion > 0 Then
        Editor1.Tag = "DoUntread"
        Editor1.InProcessing = True
        
        '�������е�ǰ�汾��Ԫ�����»�ȡ
        Dim oTable As cEPRTable
        For Each oTable In Me.Document.Tables
            If oTable.TableType = tte_Ĭ�� Then Call oTable.ReGetCellsFromDB(lngVersion)
        Next
    
        '�������Ҫ��
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Elements.Count
            If Me.Document.Elements(i).��ʼ�� >= lngVersion And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Elements(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Elements("K" & lngKey).DeleteFromEditor Me.Editor1
            Me.Document.Elements.Remove "K" & lngKey
        Next
        '�ָ�ɾ��������Ҫ��
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Elements.Count
            If Me.Document.Elements(i).��ֹ�� >= lngVersion - 1 And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Elements(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Elements("K" & lngKey).��ֹ�� = 0
        Next
        '������
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Diagnosises.Count
            If Me.Document.Diagnosises(i).��ʼ�� >= lngVersion And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Diagnosises(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Diagnosises("K" & lngKey).DeleteFromEditor Me.Editor1
            Me.Document.Diagnosises.Remove "K" & lngKey
        Next
        '�ָ�ɾ�������
        ReDim Preserve lngKeys(0 To 0) As Long
        For i = 1 To Me.Document.Diagnosises.Count
            If Me.Document.Diagnosises(i).��ֹ�� >= lngVersion - 1 And lngVersion > 1 Then
                lngCount = UBound(lngKeys) + 1
                ReDim Preserve lngKeys(0 To lngCount) As Long
                lngKeys(lngCount) = Me.Document.Diagnosises(i).Key
            End If
        Next
        For i = 1 To UBound(lngKeys)
            lngKey = lngKeys(i)
            Me.Document.Diagnosises("K" & lngKey).��ֹ�� = 0
        Next
        '����ı�
        With Me.Editor1
            lngLen = Len(.Text)
            blnForce = .ForceEdit
            .Tag = "DoUntread"
            .Freeze
            .ForceEdit = True
            For i = 0 To lngLen - 1
                '�ж� .Range(i, i + 1).Font.ForeColor ��ɫֵ������ȷ���ı��汾
                COLOR = .Range(i, i + 1).Font.ForeColor
                If Me.Document.IsNewCharColor(COLOR) Then
                    '���������ı�����ô������ı�
                    If .Range(i, i + 1).Font.Hidden And .Range(i, i + 3).Text = "TS(" Then
                        i = i + InStr(1, .Range(i, i + 100), ")")
                    Else
                        .Range(i, i + 1) = ""
                        lngLen = lngLen - 1
                        i = i - 1
                    End If
                ElseIf Me.Document.IsDelCharColor(COLOR) Then
                    '����ɾ���ı�����ô��ԭ���ı�
                    lng��ʼ�� = Get��ʼ��(COLOR)
                    .Range(i, i + 1).Font.ForeColor = GetCharColor(lng��ʼ��, 0)
                    .Range(i, i + 1).Font.Strikethrough = False
                    lngLen = lngLen - 1
                End If
            Next
            .ForceEdit = blnForce
            .UnFreeze
            .Tag = ""
        End With
        Editor1.Tag = ""
        Editor1.InProcessing = False
    End If
    '�����ļ�
    Me.Editor1.Modified = False
    If Me.Document.SaveEPRDoc(Me.Editor1, InStr(1, gstrPrivsEpr, "�޸�ҳ������") > 0) Then
        Call ClearUndoList      '���Undo���У�
        DT1_EPR = Now
'        MsgBox "���˲����ɹ���", vbOKOnly + vbInformation, gstrSysName
        If mblnIsMultiMode And Not mfrmMultiDocView Is Nothing Then
            mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'################################################################################################################
'## ���ܣ�  ��ʾ�Զ�ʶ������Ҫ�ػ����ֵ���Ŀ��ѡ����
'##
'## ������  strAuto     :IN     �����ѯ�ؼ���
'################################################################################################################
Private Sub ShowAutoRecSelector(ByVal strF As String)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys And tblThis.Visible = False Then Exit Sub    '��֤���ܲ���ؼ����ڲ�
    If Me.Editor1.Selection.Font.Protected And tblThis.Visible = False Then Exit Sub

    Dim rs As New ADODB.Recordset
    Dim lLeft As Long, lTOp As Long, lRight As Long, lBottom As Long

    '�������������Ӣ�����������������һЩ��
    gstrSQL = "select  ID,����,������ As ����,��λ,decode(�滻��,2,'�ֵ���Ŀ',1,'�滻��Ŀ','�ⲿ������') As ���� " & _
        "From ����������Ŀ " & _
        "Where ������ Like '%" & strF & "%' Or Ӣ���� Like '%" & UCase(strF) & "%' " & _
        "Order By ����"

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ")
    If rs.EOF Then Exit Sub
    Dim pt As POINTAPI, arrPara As String, T As Variant, lngId As Long
    Dim f As New frmSelectChild

    pt.X = 0
    pt.y = 0
    ClientToScreen Editor1.OriginRTB.hwnd, pt
    '��ȡ��ʼλ������
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then
            tblThis.Cells("K" & tblThis.SelectedCellKey).GetCellTextBorder lLeft, lTOp, lRight, lBottom
            lLeft = Me.Left + Editor1.Left + Editor1.UILeft + tblThis.Left + lLeft * 15 + 30
            lTOp = Me.Top + Editor1.Top + Editor1.UITop + tblThis.Top + lBottom * 15 + 500
            arrPara = "0;830;2500;700;1000"
            strF = f.ShowSelectChild(Me, lLeft, lTOp, 5550, 3000, rs, arrPara)
        Else
            Exit Sub
        End If
    Else
        Editor1.Range(Editor1.Selection.StartPos, Editor1.Selection.StartPos + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp
        Call AddUndoPoint  '�ֶ�����
        arrPara = "0;830;2500;700;1000"
        strF = f.ShowSelectChild(Me, pt.X * Screen.TwipsPerPixelX + lLeft, pt.y * Screen.TwipsPerPixelY + lTOp, 5550, 3000, rs, arrPara)
    End If


    If strF = "" Then
        Exit Sub
    Else
        T = Split(strF, ";")
        lngId = T(0)
        rs.Close
        gstrSQL = "Select ID, ������, ����, ����, С��, ��λ, ��ʾ��, �滻��, ��ʼֵ, ��ֵ�� From ����������Ŀ Where ID =[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", lngId)
        If Not rs.EOF Then
            '����Ԫ��
            Dim Ele As New cEPRElement, aryTemp() As String, lngKey As Long, lngCount As Long
            With Ele
                .Ҫ������ = NVL(rs("������"))
                .����Ҫ��ID = NVL(rs("ID"), 0)
                .Ҫ������ = NVL(rs("����"), 1)
                .Ҫ�س��� = NVL(rs("����"), 0)
                .Ҫ��С�� = NVL(rs("С��"), 0)
                .Ҫ�ص�λ = NVL(rs("��λ"))
                .Ҫ�ر�ʾ = IIf(NVL(rs("��ʾ��"), 0) = 4, 2, NVL(rs("��ʾ��"), 0))
                .�滻�� = NVL(rs("�滻��"), 0)      '0-�ⲿ������Ŀ��1-�滻��Ŀ��2-�ֵ���Ŀ
                .�����ı� = Trim(NVL(rs("��ʼֵ")))
                If .Ҫ������ = 0 Then
                    Select Case .Ҫ�ر�ʾ
                    Case 0, 1
                        If Trim(NVL(rs("��ֵ��"))) = "" Then
                            .Ҫ��ֵ�� = ""
                        Else
                            aryTemp = Split(NVL(rs("��ֵ��")), ";")
                            .Ҫ��ֵ�� = Val(aryTemp(0)) & ";" & Val(aryTemp(1))
                        End If
                    Case 2
                        aryTemp = Split(NVL(rs("��ֵ��")), ";")
                        For lngCount = 0 To UBound(aryTemp)
                            aryTemp(lngCount) = Val(aryTemp(lngCount))
                        Next
                        .Ҫ��ֵ�� = Join(aryTemp(0), ";")
                    Case Else
                        .Ҫ��ֵ�� = ""
                    End Select
                Else
                    Select Case .Ҫ�ر�ʾ
                    Case 2, 3
                        .Ҫ��ֵ�� = NVL(rs("��ֵ��"))
                    Case Else
                        .Ҫ��ֵ�� = ""
                    End Select
                End If
                .������̬ = IIf(.Ҫ�ر�ʾ = 2 Or .Ҫ�ر�ʾ = 3, 1, 0) '0-�ı� 1-���� 2-��ѡ 3-��ѡ   ���Ϊ��ѡ����ѡ��������Ĭ��ֵΪչ����Ŀ   0-����;1-չ��
            End With
            If tblThis.Visible Then
                If Val(tblThis.Tag) > 0 Then
                    lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.AddExistNode(Ele)
                    If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                    Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).Ҫ������, _
                        Me.Document.EPRPatiRecInfo.����ID, _
                        Me.Document.EPRPatiRecInfo.��ҳID, _
                        Me.Document.EPRPatiRecInfo.������Դ, _
                        Me.Document.EPRPatiRecInfo.ҽ��id, _
                        Me.Document.EPRPatiRecInfo.Ӥ��)
                    End If
                    '���浽��Ԫ����
                    tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�����ı�
                    tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
                    tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).Ҫ������
                    tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                    tblThis.Modified = True
                    tblThis.Refresh False, True, tblThis.SelectedCellKey
                    tblThis_Resize tblThis.Width, tblThis.Height
                End If
            Else
                lngKey = Me.Document.Elements.AddExistNode(Ele)
                '�滻��Ŀ
                If Me.Document.Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                    Me.Document.Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).Ҫ������, _
                        Me.Document.EPRPatiRecInfo.����ID, _
                        Me.Document.EPRPatiRecInfo.��ҳID, _
                        Me.Document.EPRPatiRecInfo.������Դ, _
                        Me.Document.EPRPatiRecInfo.ҽ��id, _
                        Me.Document.EPRPatiRecInfo.Ӥ��)
                End If
                Me.Document.Elements("K" & lngKey).��ʼ�� = Me.Document.Ŀ��汾

                '��������Ҫ�ص��༭����
                Dim blnForce As Boolean
                blnForce = Me.Editor1.ForceEdit
                Me.Editor1.ForceEdit = True
                Me.Editor1.Tag = "ShowAutoRecSelector"
                Me.Editor1.SelText = ""
                Me.Document.Elements("K" & lngKey).InsertIntoEditor Me.Editor1, , True
                Me.Editor1.ForceEdit = blnForce
                Me.Editor1.Tag = ""
                'ͬʱ�����༭���幩����
                If (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) And Me.Document.Elements("K" & lngKey).�滻�� <> 1 Then
                    bInKeys = FindKey(Editor1, "E", lngKey, lSS, lSE, lES, lEE, bNeeded)
                    If bInKeys Then
                        '��λ
                        Me.Editor1.Range(lSE, lES).Selected
                        ShowEleEditor 0, 0
                    End If
                End If
            End If
        End If
    End If
    If tblThis.Visible = False Then Call ClearNoUseUndoList
End Sub

Private Sub cbrThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

'################################################################################################################
'## ���ܣ�  �༭��λ�õ���
'################################################################################################################
Private Sub cbrThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    On Error Resume Next
    cbrThis.GetClientRect Left, Top, Right, Bottom

    If Right >= Left And Bottom >= Top Then
        If Not Me.Document Is Nothing Then
            
            If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000 Then
                imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000
            End If
            
            imgX_S.Move Left, imgX_S.Top, Right - Left
            
            If picPane.Visible Then
            
                If imgX_S.Top > Bottom - Top - 1000 Then imgX_S.Top = Bottom - Top - 1000
                            
                picPane.Move Left, Top, Right - Left, imgX_S.Top - Top
                If imgX_S.Top < 0 Then imgX_S.Top = picPane.Top + picPane.Height
            End If
            
            If ChildMode = False And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                
                If picPane.Visible Then
                    picPatiInfo.Move Left, imgX_S.Top + imgX_S.Height, Right - Left
                    Editor1.Move Left, picPatiInfo.Top + picPatiInfo.Height, Right - Left, Bottom - Top - (picPane.Height + imgX_S.Height + picPatiInfo.Height)
                Else
                    picPatiInfo.Move Left, Top, Right - Left
                    Editor1.Move Left, picPatiInfo.Top + picPatiInfo.Height, Right - Left, Bottom - Top - picPatiInfo.Height
                End If
                picPatiInfo.Visible = True
            Else
                If picPane.Visible Then
                    Editor1.Move Left, imgX_S.Top + imgX_S.Height, Right - Left, Bottom - Top - picPane.Height - imgX_S.Height
                Else
                    Editor1.Move Left, Top, Right - Left, Bottom - Top
                End If
                picPatiInfo.Visible = False
            End If
        End If
    Else
        Editor1.Move 0, 0, 0, 0
    End If
    If Editor1.ViewMode = cprNormal Then
        picPenInput.Move Editor1.Width - picPenInput.Width - 300, Editor1.Height - picPenInput.Height - 300
    Else
        picPenInput.Visible = False
    End If
End Sub

'################################################################################################################
'## ���ܣ�  �˵�&�����������¼�
'################################################################################################################
Private Sub cbrThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, eEPRType As EPRDocTypeEnum
    
    If Me.Visible = False Then Exit Sub
    If Me.Document Is Nothing Then Exit Sub
    If mblnReadOnly And Control.ID <> ID_FILE_EXIT And Control.ID <> ID_COMMON_CANCEL Then
        'ֻ������ģʽ�����в˵���Ч�������˳��˵���
        Control.Enabled = False
        Exit Sub
    End If
    eEPRType = Me.Document.EPRFileInfo.����
    
    Select Case Control.ID
    Case ID_FILE_PRINTPREVIEW, ID_FILE_PRINT, ID_FILE_PRINTINWORD
        Control.Enabled = mblnCanPrint
        If Control.Enabled And (eEPRType = cprסԺ���� Or eEPRType = cpr���ﲡ��) Then
            Control.Enabled = IIf(Document.Signs.Count = 0, InStr(1, gstrPrivsEpr, "δǩ����ӡ") > 0, InStr(1, gstrPrivsEpr, "������ӡ") > 0)
        End If
    Case ID_FILE_EXIT, ID_COMMON_CANCEL
        Control.Enabled = (mblnChildMode = False)
        Control.Visible = Control.Enabled
        If mintStyle = -1 Then
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_EDIT_UNDO
        Control.Enabled = CanUndo And mblnAutosave And (Me.Editor1.ReadOnly = False)
        Control.Visible = mblnAutosave
    Case ID_EDIT_REDO
        Control.Enabled = CanRedo And mblnAutosave And (Me.Editor1.ReadOnly = False)
        Control.Visible = mblnAutosave
    Case ID_VIEW_STRUCTURE
        Control.Checked = mfrmCompends.Visible
    Case ID_VIEW_PHRASEDEMO
        Control.Checked = mfrmSentenceDetailed.Visible
    Case ID_VIEW_SEGMENT
        Control.Enabled = (Me.Document.EditType <> cprET_�����ļ�����)
        Control.Checked = mfrmSegments.Visible
    Case ID_VIEW_PACSPIC
        Control.Visible = (eEPRType = cpr���Ʊ���)
        Control.Enabled = (eEPRType = cpr���Ʊ���)
        Control.Checked = mfrmSegments.Visible
    Case ID_VIEW_HISTORYREPORT
        Control.Visible = (eEPRType = cpr���Ʊ���)
        Control.Enabled = (eEPRType = cpr���Ʊ���)
        Control.Checked = mfrmHistoryReport.Visible
    Case ID_VIEW_HISTORYWINDOW
        Control.Enabled = mblnExistHistroy
        Control.Visible = Control.Enabled
        Control.Checked = picHistoryInfo.Visible
    Case ID_FILE_CLEAR
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Document Is Nothing Then
                Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
            Else
                If Me.Document.EditType = cprET_��������� Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
                End If
            End If
        End If
    Case ID_FILE_IMPORT
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Document Is Nothing Then
                Control.Enabled = Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False And Editor1.AuditMode = False
            Else
                If Me.Document.EditType = cprET_�����ļ����� Then
                    Control.Enabled = False
                Else
                    Control.Enabled = Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False And Editor1.AuditMode = False
                End If
            End If
        End If
        If Control.Enabled Then Control.Enabled = InStr(gstrPrivsEpr, "��ʷ�ļ�") > 0
    Case ID_FILE_IMPORTFROMXML
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Document Is Nothing Then
                Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
            Else
                If Me.Document.EditType = cprET_��������� Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False)
                End If
            End If
        End If
        If Control.Enabled Then Control.Enabled = InStr(gstrPrivsEpr, "����/��XML�ļ�") > 0
    Case ID_FILE_SAVE, ID_FILE_SAVE_QUIT
        Control.Enabled = (Editor1.Modified) And (Me.Document.Ŀ��汾 <= 16) And (Editor1.ViewMode = cprNormal)
        If Control.ID = ID_FILE_SAVE_QUIT And mintStyle = -1 Then
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_FILE_SAVEAS
        Control.Enabled = (Editor1.ViewMode = cprNormal And mblnCanPrint)
        If Control.Enabled Then Control.Enabled = InStr(gstrPrivsEpr, "����RTF�ļ�") > 0
    Case ID_FILE_EXPORTTOHTML
        Control.Enabled = (Editor1.ViewMode = cprNormal And mblnCanPrint)
    Case ID_FILE_EXPORTTOXML
        Control.Enabled = (Editor1.ViewMode = cprNormal And mblnCanPrint)
        Control.Enabled = InStr(gstrPrivsEpr, "����/��XML�ļ�") > 0
    Case ID_FILE_SAVEASEPRDEMO, ID_FILE_SAVEASSEGMENT
        Control.Enabled = (Editor1.ViewMode = cprNormal And Me.Document.EditType <> cprET_�����ļ�����)
    Case ID_EDIT_CUT:
        Control.Enabled = (tblThis.Visible) Or (Editor1.CanCopy And Editor1.ViewMode = cprNormal And (Me.Editor1.ReadOnly = False))
    Case ID_EDIT_COPY
        If Me.ActiveControl Is edtThis Then
            Control.Enabled = (edtThis.CanCopy And edtThis.ViewMode = cprNormal)
        Else
            Control.Enabled = (tblThis.Visible) Or (Editor1.CanCopy And Editor1.ViewMode = cprNormal)
        End If
        Control.Visible = InStr(gstrPrivsEpr, "���ݸ���") > 0
    Case ID_EDIT_COPYSELF
         Control.Enabled = (tblThis.Visible) Or (Me.Editor1.Selection.Font.ForeColor <> tomUndefined And Me.Editor1.Selection.Font.Protected = False) And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False
         Control.Visible = InStr(gstrPrivsEpr, "ר�ø���") > 0
    Case ID_EDIT_COPYOUT
        If Me.ActiveControl Is edtThis Then
            Control.Enabled = (edtThis.CanCopy And edtThis.ViewMode = cprNormal)
        Else
            Control.Enabled = (tblThis.Visible) Or (Editor1.CanCopy And Editor1.ViewMode = cprNormal)
        End If
        Control.Visible = InStr(gstrPrivsEpr, "���ݸ���") > 0
    Case ID_EDIT_PASTE:
        Control.Enabled = (tblThis.Visible) Or (Me.Editor1.Selection.Font.ForeColor <> tomUndefined And Me.Editor1.Selection.Font.Protected = False) And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False
        Control.Visible = InStr(gstrPrivsEpr, "���ݸ���") > 0
    Case ID_EDIT_DELETE
        Control.Enabled = Editor1.ViewMode = cprNormal
    Case ID_EDIT_FORMATBRUSH
        Control.Enabled = (tblThis.Visible = False) And (Editor1.ViewMode = cprNormal) And Editor1.AuditMode = False And Editor1.ReadOnly = False
        Control.Checked = mblnFmtBrushDown
    Case ID_EDIT_FIND, ID_EDIT_FINDNEXT     ', ID_EDIT_UNDO, ID_EDIT_REDO
        Control.Enabled = (Editor1.ViewMode = cprNormal)
    Case ID_EDIT_REPLACE
        Control.Enabled = (Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_INSERT_DATETIME, ID_INSERT_DATE, ID_INSERT_TIME, ID_INSERT_SPECIALCHAR
        Control.Enabled = Editor1.ViewMode = cprNormal And Me.Editor1.ReadOnly = False
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                On Error Resume Next
                Control.Enabled = Control.Enabled And (tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = False)
            End If
        Else
            Control.Enabled = Control.Enabled And Editor1.Selection.Font.Protected = False
        End If
    Case ID_INSERT_ELEMENT
        Control.Enabled = Editor1.ViewMode = cprNormal And Me.Editor1.ReadOnly = False
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                On Error Resume Next
                Control.Enabled = Control.Enabled And (tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing)
            End If
        End If
    Case ID_INSERT_PICTURE
        Control.Enabled = (Editor1.ViewMode = cprNormal And (Editor1.Selection.Font.Protected = False Or mbEditInTable Or ucPacsImgCanvas1.Visible) And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
        If tblThis.Visible Then
            If tblThis.SelectedCellKey > 0 Then
                On Error Resume Next
                Control.Enabled = Control.Enabled And ((Not tblThis.Cells("K" & tblThis.SelectedCellKey).Picture Is Nothing) Or (tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = ""))
            End If
        End If
    Case ID_INSERT_DOCADVISE
        Control.Enabled = Not Editor1.AuditMode
        If Control.Enabled Then Control.Enabled = Not Editor1.ReadOnly
        Control.Visible = (eEPRType = cpr���ﲡ��)
    Case ID_INSERT_EPRDEMO
        If Me.Editor1.AuditMode Then
            Control.Enabled = False
        Else
            If Me.Document Is Nothing Then
                Control.Enabled = (Editor1.ViewMode = cprNormal) And Me.Editor1.ReadOnly = False
            Else
                Control.Enabled = (Editor1.ViewMode = cprNormal) And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) And Me.Editor1.ReadOnly = False
            End If
        End If
    Case ID_INSERT_PACSPIC
        Control.Enabled = (Document.EPRFileInfo.���� = cpr���Ʊ��� And Me.Editor1.ReadOnly = False)
        Control.Visible = Control.Enabled
    Case ID_EDIT_ADDCOMPEND, ID_EDIT_MODCOMPEND, ID_EDIT_REFCOMPEND, ID_EDIT_DELCOMPEND
        If Me.Editor1.AuditMode Then
            If Control.ID <> ID_EDIT_REFCOMPEND Then Control.Enabled = False
        Else
            If Editor1.ViewMode = cprNormal Then
                If Control.ID = ID_EDIT_DELCOMPEND Then
                    If mfrmCompends.Tree.SelectedItem Is Nothing Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Me.Editor1.ReadOnly = False)
                    End If
                ElseIf Control.ID = ID_EDIT_MODCOMPEND Then
                    If mfrmCompends.Tree.SelectedItem Is Nothing Then
                        Control.Enabled = False
                    Else
                        Control.Enabled = (Me.Editor1.ReadOnly = False)
                    End If
                ElseIf Control.ID = ID_EDIT_ADDCOMPEND Then
                    Control.Enabled = (Editor1.Selection.Font.Protected = False)
                Else
                    Control.Enabled = (Me.Editor1.ReadOnly = False)
                End If
            Else
                Control.Enabled = False
            End If
        End If
        If Not Me.Document Is Nothing Then
            If Me.Document.EditType <> cprET_�����ļ����� And Control.ID <> ID_EDIT_REFCOMPEND Then 'ֻ���ڶ���ʱ�޸����
                Control.Enabled = False
                Control.Visible = False
            End If
        End If
    Case ID_Main_FORMAT
        Control.Enabled = Control.Visible And (InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0)
        Control.Visible = Control.Enabled
        If Control.Visible Then
            If Not Document Is Nothing Then
                Control.Visible = Not (Document.EPRDemoInfo.���� <> 0 And Document.EditType = cprET_ȫ��ʾ���༭)
            End If
        End If
    Case ID_FORMAT_FONTNAME
        Control.Visible = (Editor1.AuditMode = False) And (InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0)
        
        If Control.Type = xtpControlComboBox Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    On Error Resume Next
                    Control.Text = (tblThis.Cells("K" & tblThis.SelectedCellKey).FontName)
                End If
            Else
                Control.Text = CStr(Editor1.Selection.Font.Name)
            End If
        End If
    Case ID_FORMAT_PARA
        If Not Me.Document Is Nothing Then
            Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
            Control.Visible = (Editor1.AuditMode = False)
        End If
    Case ID_FORMAT_FONT
        Control.Visible = (Editor1.AuditMode = False) And (InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0)
        If Not Me.Document Is Nothing Then Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_FONTSIZE
        Control.Visible = (Editor1.AuditMode = False) And (InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0)
        If Control.Type = xtpControlComboBox Then
            If tblThis.Visible Then
                If tblThis.SelectedCellKey > 0 Then
                    On Error Resume Next
                    Control.Text = (tblThis.Cells("K" & tblThis.SelectedCellKey).FontSize)
                End If
            Else
                Control.Text = CStr(IIf(Editor1.Selection.Font.Size = tomUndefined, "", GetFontSizeChinese(Editor1.Selection.Font.Size)))
            End If
        End If
    Case ID_FORMAT_BOLD
        Control.Visible = Control.Visible And (InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0)
        Control.Checked = Editor1.Selection.Font.Bold
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE
        Control.Visible = Control.Visible And (InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0)
        Control.Checked = (Editor1.Selection.Font.Underline <> cprNone)
        Control.Enabled = (Editor1.AuditMode = False And Editor1.Selection.Font.Protected = False)
    Case ID_FORMAT_ALIGNLEFT
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0
        Control.Checked = (Editor1.Selection.Para.Alignment = cprHALeft)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_ALIGNCENTER
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0
        Control.Checked = (Editor1.Selection.Para.Alignment = cprHACenter)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_ALIGNRIGHT
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0
        Control.Checked = (Editor1.Selection.Para.Alignment = cprHARight)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_LINESPACE
        Control.Visible = InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0
        Control.Checked = (Editor1.Selection.Para.LineSpacing > 1#)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_SPACEBEFORE
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_SPACEAFTER
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_FIRSTINDENT
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_FIRSTHUNGING
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTARABIC
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsArabic)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTBULLETS
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTBullet)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTLCHAR
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsLCLetter)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTLROME
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsLCRoman)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTUCHAR
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsUCLetter)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTUROME
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNumberAsUCRoman)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTNONE
        Control.Checked = (Editor1.Selection.Para.ListType = cprLTNone)
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_LISTSETUP
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_PROTECT
        Control.Checked = (Editor1.Selection.Font.Protected And Editor1.Selection.Font.ForeColor = PROTECT_FORECOLOR)
        Control.Enabled = (Me.Document.EditType = cprET_�����ļ����� Or Me.Document.EditType = cprET_ȫ��ʾ���༭ Or InStr(1, gstrPrivsEpr, "�����ı�����") > 0)
    Case ID_FORMAT_SUPER
        Control.Checked = Editor1.Selection.Font.Superscript
        Control.Enabled = (tblThis.Visible = False)
    Case ID_FORMAT_SUB
        Control.Checked = Editor1.Selection.Font.Subscript
        Control.Enabled = (tblThis.Visible = False)
    Case ID_FORMAT_UNDERLINE_DASH
        Control.Checked = (Editor1.Selection.Font.Underline = cprDash)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_DASHDOT
        Control.Checked = (Editor1.Selection.Font.Underline = cprDashDot)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_DASHDOT2
        Control.Checked = (Editor1.Selection.Font.Underline = cprDashDotDot)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_DOT
        Control.Checked = (Editor1.Selection.Font.Underline = cprDotted)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_THIN
        Control.Checked = (Editor1.Selection.Font.Underline = cprHair)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_THICK
        Control.Checked = (Editor1.Selection.Font.Underline = cprThick)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_WAVE
        Control.Checked = (Editor1.Selection.Font.Underline = cprWave)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_UNDERLINE_NONE
        Control.Checked = (Editor1.Selection.Font.Underline = cprNone)
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_BACKGROUND
        Control.Enabled = (Editor1.AuditMode = False)
    Case ID_FORMAT_HIGHLIGHT
        Control.Enabled = (Editor1.AuditMode = False And Editor1.Selection.Font.Protected = False)
    Case ID_VIEW_RULER
        Control.Checked = Editor1.ShowRuler
    Case ID_VIEW_PENWINDOW
        Control.Checked = picPenInput.Visible
    Case ID_FORMAT_STYLE
        Control.Visible = (InStr(";" & gstrPrivsEpr & ";", ";�����ʽ����;") > 0) And (Editor1.AuditMode = False)
        If Control.Visible Then Control.Enabled = (tblThis.Visible = False)
    Case ID_FORMAT_INDENTDECREASE, ID_FORMAT_INDENTINCREASE
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
    Case ID_FORMAT_STYLEWINDOW
        '��ʽ����
        Control.Checked = mfrmStyleMan.Visible
        Control.Enabled = (Editor1.AuditMode = False) And (tblThis.Visible = False)
        Control.Visible = (Editor1.AuditMode = False)
    Case ID_VIEW_HEADFOOT
        Control.Enabled = (Me.Document.EditType = cprET_�����ļ�����) And (Editor1.AuditMode = False)
    Case ID_FILE_PAGESETUP
        Control.Enabled = (InStr(1, gstrPrivsEpr, "�޸�ҳ������") > 0) Or ((Me.Document.EditType = cprET_�����ļ�����) And (Editor1.AuditMode = False))
    Case ID_SIGN, ID_SIGN_QUIT
        If Me.Document.EditType = cprET_�������༭ Then
            Control.Enabled = (Me.Document.Signs.Count = 0 And Me.Editor1.ReadOnly = False)
        ElseIf Me.Document.EditType = cprET_��������� Then
            Control.Enabled = (Me.Document.Ŀ��汾 <= 16 And Me.Editor1.ReadOnly = False)
        Else
            Control.Enabled = False: Control.Visible = False
        End If
        If Control.ID = ID_SIGN_QUIT And mintStyle = -1 Then
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_PATISIGN
        Control.Visible = False
        If Me.Document.EditType = cprET_�������༭ Then
            If eEPRType = cpr֪���ļ� And mblnEnPtSign Then
                Control.Visible = True
            End If
            Control.Enabled = (Editor1.AuditMode = False)
            If Control.Enabled Then Control.Enabled = (Me.Document.Signs.Count = 0 And Me.Editor1.ReadOnly = False)
            Control.Caption = IIf(mblnPatiSign, "���߳�ǩ", "����ǩ��")
        Else
            Control.Enabled = False: Control.Visible = False 'Control.Visible = Me.Document.EditType = cprET_�������༭ ������DockPane���ܻ�ԭ
        End If
    Case ID_UNTREAD
        If Me.Document.Signs.Count > 0 Then
            If Me.Document.Signs("K" & Document.Signs.GetMaxKey).���� <> gstrUserName And Me.Document.Signs("K" & Document.Signs.GetMaxKey).���� <> gstrSignName _
                And InStr(gstrPrivsEpr, "��������ǩ��") = 0 Then  '��������ǩ��Ȩ�޿���
                Control.Visible = False
            Else
                Control.Visible = True
            End If
        End If
        
        If Me.Document.EditType = cprET_�������༭ Then
            Control.Enabled = IIf(mblnIsMultiMode, Me.Document.EPRPatiRecInfo.���汾 = 1, Me.Document.Signs.Count > 0) And Me.Editor1.Modified = False And Editor1.ReadOnly = True
        ElseIf Me.Document.EditType = cprET_��������� Then
            Control.Enabled = (Me.Document.EPRPatiRecInfo.���汾 > 1 Or Me.Document.Signs.Count > 1) And Me.Editor1.Modified = False And Editor1.ReadOnly = False
        Else
            Control.Enabled = False: Control.Visible = False
        End If
    Case ID_REVISION_PREV, ID_REVISION_NEXT, ID_REVISION_RESET
        Control.Enabled = Me.Editor1.AuditMode And (Me.Editor1.ReadOnly = False)
        Control.Visible = Control.Enabled
    Case ID_DIAGNOSIS '���
        Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.Selection.Font.Protected = False) And Me.Editor1.ReadOnly = False And Me.Document.EditType <> cprET_�����ļ�����
    Case ID_EDIT_MARKEDPIC, ID_EDIT_OUTERPIC
        Control.Enabled = Me.Editor1.ReadOnly = False And (Editor1.ViewMode = cprNormal) And (Editor1.AuditMode = False)
    Case ID_INSERT_TABLE
        Control.Enabled = (Editor1.ViewMode = cprNormal And Editor1.Selection.Font.Protected = False And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_FORMAT_FORECOLOR, ID_TABLE_CELLALIGNMENT, ID_TABLE_CURRENCY, ID_TABLE_PERCENT, ID_TABLE_KILOBIT, ID_TABLE_CELLPROTECTED, ID_TABLE_INSERTPICTURE, ID_TABLE_BEELEMENTS
        Control.Enabled = tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_MERGE
        Control.Enabled = tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
        If tblThis.SelectedCellKey > 0 Then
            On Error Resume Next
            Control.Checked = (Len(tblThis.Cells("K" & tblThis.SelectedCellKey).MergeInfo) = 16)
        End If
    Case ID_TABLE_DELETETABLE, ID_TABLE_DELETECOL, ID_TABLE_DELETEROW, ID_TABLE_FORMATCELL, ID_TABLE_PROPERTY
        Control.Enabled = tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_INSERTROWDOWN, ID_TABLE_INSERTROWUP, ID_TABLE_INSERTCOLLEFT, ID_TABLE_INSERTCOLRIGHT, ID_TABLE_INSERTINHERITROW
        Control.Enabled = tblThis.SelectedCellKey > 0 And tblThis.Visible And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_INSERTTABLE
        Control.Enabled = (tblThis.Visible = False) And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_FORMATROWHEIGHT, ID_TABLE_FORMATCOLWIDTH
        Control.Enabled = (tblThis.Visible = True) And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_TABLE_CELLALIGNMENT1, ID_TABLE_CELLALIGNMENT2, ID_TABLE_CELLALIGNMENT3, ID_TABLE_CELLALIGNMENT4, ID_TABLE_CELLALIGNMENT5, ID_TABLE_CELLALIGNMENT6, ID_TABLE_CELLALIGNMENT7, ID_TABLE_CELLALIGNMENT8, ID_TABLE_CELLALIGNMENT9
        Control.Enabled = (tblThis.Visible = True) And (Editor1.ViewMode = cprNormal And Editor1.AuditMode = False) And Me.Editor1.ReadOnly = False
    Case ID_Main_HELP
        Control.Visible = Not mblnChildMode
    End Select
    
    If Not Control.Visible Then Control.Enabled = False
End Sub

Private Sub cPicEditor_pOK(ByRef FinalPicture As StdPicture, ByVal lngWidth As Long, ByVal lngHeight As Long)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim blnForce As Boolean

    lKey = cPicEditor.lngKeyOfPic
    If lKey > 0 And FinalPicture <> 0 Then
        If tblThis.Visible Then
            '����е�ͼƬ
            If Val(tblThis.Tag) > 0 Then
                '��ͼƬ���󱣴浽�������
                Dim ctlPic As VB.PictureBox
                Set ctlPic = gfrmPublic.Controls.Add("VB.PictureBox", "ctlPic" & CLng(Timer * 1000))
                ctlPic.AutoRedraw = True
                ctlPic.BorderStyle = 0
                ctlPic.Height = lngHeight
                ctlPic.Width = lngWidth
                ShowPicMarks ctlPic, FinalPicture, Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).PicMarks

                Set Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).OrigPic = FinalPicture
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).Width = lngWidth
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).Height = lngHeight
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).OrigWidth = lngWidth
                Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lKey).OrigHeight = lngHeight

                '���浽��Ԫ����
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = ""
                tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lKey
                tblThis.Cells("K" & tblThis.SelectedCellKey).Picture = ctlPic.Picture
                tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = ""
                tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                tblThis.Modified = True
                tblThis.Refresh
                tblThis_Resize tblThis.Width, tblThis.Height
                Editor1.Modified = True

                gfrmPublic.Controls.Remove ctlPic
                Set ctlPic = Nothing
            End If
        Else
            '�滻ͼƬ
            Set Me.Document.Pictures("K" & lKey).OrigPic = FinalPicture

            Me.Document.Pictures("K" & lKey).OrigWidth = lngWidth
            Me.Document.Pictures("K" & lKey).OrigHeight = lngHeight
            Me.Document.Pictures("K" & lKey).Width = lngWidth
            Me.Document.Pictures("K" & lKey).Height = lngHeight
            Me.Document.Pictures("K" & lKey).Modified = True

            bInKeys = FindKey(Me.Editor1, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
            If bInKeys = False Then Exit Sub
            Dim ParaFmt As New cParaFormat

            With Me.Editor1
                blnForce = .ForceEdit
                .Freeze
                .Tag = "cPicEditor_pOK"
                .ForceEdit = True
                Set ParaFmt = .Range(lSE, lES).Para.GetParaFmt

                If Me.Document.Pictures("K" & lKey).�Ƿ��� Then
                    .Range(lSS, lEE + 2).Text = ""
                Else
                    .Range(lSS, lEE).Text = ""
                End If
                Me.Document.Pictures("K" & lKey).InsertIntoEditor Me.Editor1, lSS, True

                .Range(lSE, lES).Para.SetParaFmt ParaFmt
                .Range(lSS, lEE).Font.Protected = True
                .ForceEdit = blnForce
                .Tag = ""
                .UnFreeze
            End With
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ��ӿ�ͣ������
'################################################################################################################
Private Sub DkpThis_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case ID_VIEW_PHRASEDEMO     'ʾ���ʾ䴰��
            Item.Handle = mfrmSentenceDetailed.hwnd
        Case ID_VIEW_SEGMENT        'ʾ�����䴰��
            Item.Handle = mfrmSegments.hwnd
        Case ID_VIEW_STRUCTURE      '�ĵ��ṹͼ����
            Item.Handle = mfrmCompends.hwnd
        Case ID_VIEW_HISTORYWINDOW  '����ҳ���ļ�����
            Item.Handle = picPane.hwnd
        Case ID_FORMAT_STYLEWINDOW  '������ʽά��
            Item.Handle = mfrmStyleMan.hwnd
        Case ID_VIEW_PACSPIC
            Item.Handle = mfrmPacsPic.hwnd
        Case ID_VIEW_MULTIDOCVIEW
            If mfrmMultiDocView Is Nothing Then
                If mblnIsMultiMode Then
                    Set mfrmMultiDocView = New frmMultiDocView
                    Item.Handle = mfrmMultiDocView.hwnd
                ElseIf Not DkpThis.FindPane(ID_VIEW_MULTIDOCVIEW) Is Nothing Then
                    DkpThis.FindPane(ID_VIEW_MULTIDOCVIEW).Hide
                End If
            Else
                Item.Handle = mfrmMultiDocView.hwnd
            End If
        Case ID_VIEW_HISTORYREPORT
            Item.Handle = mfrmHistoryReport.hwnd
        Case ID_VIEW_Assistant '��������
            Item.Handle = mfrmDocksymbol.hwnd
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵ�ǰѡ�����ݣ�˫����ͼƬ�����༭��
'################################################################################################################
Private Sub Editor1_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    If Editor1.ReadOnly Then Exit Sub
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And Editor1.AuditMode = False Then
        bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded) '���ҹؼ��� ID ��
        If bInKeys = False Then Exit Sub
        If sType = "P" Then '�༭ͼƬ
            If Me.Document.Pictures("K" & lKey).PictureType = EPRMarkedPicture Then
                Editor1.ShowUIInterface
                ucPictureEditor1.ShowMe Me, Editor1.hwnd, cbrThis, Me.Document.Tables("K" & tblThis.Tag).Pictures("K" & lKey), _
                    Editor1.UILeft, Editor1.UITop, Editor1.UIWidth, Editor1.UIHeight, False
            ElseIf Me.Document.Pictures("K" & lKey).PictureType = EPROutPicture Then
                cPicEditor.ShowPicEditor glngSys, gcnOracle, Me.Document.Pictures("K" & lKey).OrigPic, lKey, Me.Document.Pictures("K" & lKey).��������, Me, False
            End If
        End If
    ElseIf Editor1.ViewMode = cprNormal Then
        bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then
            Select Case sType
            Case "E"
                If Me.Document.Elements("K" & lKey).������̬ = 1 Then Exit Sub
                Me.Editor1.Range(lSE, lES).Selected
                ShowEleEditor 0, 0
            Case "S" '��дǩ������
                Me.Editor1.Range(lSE, lES).Selected
            End Select
        End If
    End If
End Sub

Private Sub Editor1_GetDelCharColor(COLOR As stdole.OLE_COLOR)
    If Editor1.AuditMode Then
        COLOR = Me.Document.GetDelCharColor(COLOR)
    End If
End Sub

Private Sub Editor1_GetNewCharColor(COLOR As stdole.OLE_COLOR)
    If Editor1.AuditMode Then
        COLOR = Me.Document.GetNewCharColor(COLOR)
    End If
End Sub

Private Sub Editor1_IsDelCharColor(ByVal COLOR As stdole.OLE_COLOR, blnIsDelCharColor As Boolean)
    If Editor1.AuditMode Then
        blnIsDelCharColor = Me.Document.IsDelCharColor(COLOR)
    End If
End Sub

Private Sub Editor1_IsNewCharColor(ByVal COLOR As stdole.OLE_COLOR, blnIsNewCharColor As Boolean)
    If Editor1.AuditMode Then
        blnIsNewCharColor = Me.Document.IsNewCharColor(COLOR)
    End If
End Sub

'################################################################################################################
'## ���ܣ�  �������������Ĵ���
'##         ���ѡ������Ҫ�أ��򵯳�����Ҫ�ر༭����
'################################################################################################################
Private Sub Editor1_KeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Editor1.ReadOnly Then Exit Sub
    Dim i As Long, blnForce As Boolean
    If ViewMode = cprPaper Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyEscape Then Exit Sub
    Editor1.Tag = "Editor1.KeyDown"

    If Me.Editor1.AuditMode Then
        If Shift <> 0 Then Editor1.Tag = "": Exit Sub
        Select Case KeyCode
        Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
            vbKeyEscape, vbKeyDelete, vbKeyBack, vbKeyTab, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
            vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
            vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12

            Editor1.Tag = ""
            Exit Sub
        End Select

        If Editor1.SelLength > 0 Then
            If Editor1.Selection.Font.Protected = False Then
                Editor1.ForceEdit = True
                Editor1.Selection.Font.ForeColor = Me.Document.GetDelCharColor(Editor1.Selection.Font.ForeColor)
                Editor1.Selection.Font.Strikethrough = True
                Editor1.Selection.Font.Underline = False
                Editor1.ForceEdit = False
            End If
            Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
            Editor1.SelLength = 0
        End If
        '���ģʽ�µ������������Ĵ���
        '���������عؼ��ֺ��棬���Զ���һ���ո񣨷Ǳ������������ԣ�
        With Editor1
            blnForce = .ForceEdit
            If .SelLength = 0 And .Selection.Font.ForeColor = PROTECT_FORECOLOR Then
                .ForceEdit = True
                .Selection.Font.ForeColor = tomAutoColor
                .ForceEdit = blnForce
            End If
            i = .Selection.StartPos
LL1:
            If .Range(i - 1, i).Font.Hidden And _
                .Range(i, i + 1).Font.Hidden = False And _
                .Range(i, i + 1).Font.Protected = False Then
                'A���⣺�������ı���|��ͨ�ı�
                .ForceEdit = True
                .Range(i, i).Text = " "
                .Range(i, i + 1).Font.Protected = False
                .Range(i, i + 1).Font.Hidden = False
                .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(tomAutoColor)
                .Range(i, i + 1).Font.BackColor = tomAutoColor
                .Range(i, i + 1).Font.Underline = cprNone
                .Range(i, i + 1).Font.Strikethrough = False
                .Range(i, i + 1).Selected
                .ForceEdit = blnForce
            Else
                If .Range(i - 1, i).Font.Hidden And _
                    .Range(i, i + 1).Font.Hidden = False And _
                    .Range(i, i + 1).Font.Protected And .Range(i, i + 1).Font.ForeColor <> PROTECT_FORECOLOR Then
                    'B����1����ͨ�ı��������ı���|�������ı����������ı�����ͨ�ı�
                    i = i - 16
                    If .Range(i - 1, i + 3) Like ")?S(" And _
                        .Range(i - 1, i + 3).Font.Hidden = True Then
                        'C���⣺�������ı����������ı����������ı���|�������ı����������ı����������ı���
                        mlngHP = -1
                        .ForceEdit = True
                        .Range(i, i).Font.Protected = False
                        .Range(i, i).Font.Hidden = False
                        .Range(i - 1, i).Font.ForeColor = vbBlack
                        blnSpaceEvent = True
                        .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")

                        '�����ı��ĸ�ʽ����
                        .Range(i, i + 1).Font.Protected = False
                        .Range(i, i + 1).Font.Hidden = False
                        .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.Underline = cprNone

                        .Range(i + 1, i + 1).Selected
                        .ForceEdit = blnForce
                    ElseIf .Range(i + 1, i + 3) = "E(" And .Range(i, i + 3).Font.Protected And .Range(i, i + 3).Font.ForeColor <> PROTECT_FORECOLOR And _
                        .Range(i + 16, i + 18) = vbCrLf And .Range(i + 16, i + 18).Font.Protected And .Range(i + 16, i + 18).Font.ForeColor <> PROTECT_FORECOLOR Then
                        'D���⣺��ٺ����ͼƬ����֮��û������ʱ���޷�������������
                        i = i + 16
                        mlngHP = -1
                        .ForceEdit = True
                        .Range(i, i).Font.Protected = False
                        .Range(i, i).Font.Hidden = False
                        .Range(i - 1, i).Font.ForeColor = vbBlack
                        blnSpaceEvent = True
                        .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")

                        '�����ı��ĸ�ʽ����
                        .Range(i, i + 1).Font.Protected = False
                        .Range(i, i + 1).Font.Hidden = False
                        .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.Underline = cprNone

                        If (.Range(i - 16, i - 14) <> "EE") Then
                            .Range(i, i + 1).Selected
                        Else
                            .Range(i + 1, i + 1).Selected
                        End If
                        .ForceEdit = blnForce
                    Else
                        .Range(i, i).Selected
                        On Error Resume Next
                        .OriginRTB.SelColor = Me.Document.GetNewCharColor(.OriginRTB.SelColor)
                        .OriginRTB.SelStrikeThru = False    'ȥ��ɾ����
                        .OriginRTB.SelUnderline = False     'ȥ���»���
                    End If
                ElseIf .Range(i - 1, i).Font.Hidden = False And _
                    .Range(i - 1, i).Font.Protected And .Range(i - 1, i).Font.ForeColor <> PROTECT_FORECOLOR And _
                    .Range(i, i + 1).Font.Hidden Then
                    'B����2����ͨ�ı��������ı����������ı���|�������ı�����ͨ�ı�
                    i = i + 16
                    If .Range(i - 1, i + 3) Like ")?S(" And _
                        .Range(i - 1, i + 3).Font.Hidden = True Then
                        'C���⣺�������ı����������ı����������ı���|�������ı����������ı����������ı���
                        mlngHP = -1
                        .ForceEdit = True
                        .Range(i, i).Font.Protected = False
                        .Range(i, i).Font.Hidden = False

                        '�����ı��ĸ�ʽ����
                        .Range(i - 1, i).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i - 1, i).Font.ForeColor)
                        .Range(i - 1, i).Font.Strikethrough = False
                        .Range(i - 1, i).Font.Underline = cprNone

                        blnSpaceEvent = True
                        .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")

                        '�����ı��ĸ�ʽ����
                        .Range(i, i + 1).Font.Protected = False
                        .Range(i, i + 1).Font.Hidden = False
                        .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                        .Range(i, i + 1).Font.Strikethrough = False
                        .Range(i, i + 1).Font.Underline = cprNone

                        .Range(i + 1, i + 1).Selected
                        .ForceEdit = blnForce
                    Else
                        GoTo LL1
                    End If
                ElseIf .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 2) = vbCrLf And .Range(i, i + 2).Font.Protected And .Range(i, i + 2).Font.ForeColor <> PROTECT_FORECOLOR Then
                    .ForceEdit = True
                    .Range(i, i) = " "
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = Me.Document.GetNewCharColor(.Range(i, i + 1).Font.ForeColor)
                    .Range(i, i + 1).Font.Strikethrough = False
                    .Range(i, i + 1).Font.Underline = cprNone
                    .Range(i, i).Selected
                    .Selection.Font.ForeColor = Me.Document.GetNewCharColor(.Selection.Font.ForeColor)
                    .ForceEdit = blnForce
                Else
                    On Error Resume Next
                    .OriginRTB.SelColor = Me.Document.GetNewCharColor(.OriginRTB.SelColor)
                    .OriginRTB.SelStrikeThru = False    'ȥ��ɾ����
                    .OriginRTB.SelUnderline = False     'ȥ���»���
                    .OriginRTB.SelItalic = True
                    If KeyCode = 1 Then KeyCode = 0
                End If
            End If
        End With
        Editor1.Tag = ""
        Exit Sub
    End If

    If Editor1.SelLength > 0 Then GoTo LL4
    If Shift <> 0 Then GoTo LL4
    Select Case KeyCode
    Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyEscape, vbKeyDelete, vbKeyBack, vbKeyTab, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
        vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12

        GoTo LL4
    End Select

    '���������عؼ��ֺ��棬���Զ���һ���ո񣨷Ǳ������������ԣ�
    With Editor1
        blnForce = .ForceEdit
        i = .Selection.StartPos
        If .Range(i, i + 2) = vbCrLf Then
            .ForceEdit = True
            .Selection.Font.Protected = False
            .Selection.Font.Hidden = False
            .Selection.Font.ForeColor = tomAutoColor
            .Selection.Font.BackColor = tomAutoColor
            .Selection.Font.Underline = cprNone
            .Selection.Font.Strikethrough = False
            .ForceEdit = blnForce
        End If
LL3:
        If .Range(i - 1, i).Font.Hidden And _
            .Range(i, i + 1).Font.Hidden = False And _
            .Range(i, i + 1).Font.Protected = False Then
            'A���⣺�������ı���|��ͨ�ı�
            .ForceEdit = True
            .Range(i, i).Text = " "
            .Range(i, i + 1).Font.Protected = False
            .Range(i, i + 1).Font.Hidden = False
            .Range(i, i + 1).Font.ForeColor = tomAutoColor
            .Range(i, i + 1).Font.BackColor = tomAutoColor
            .Range(i, i + 1).Font.Underline = cprNone
            .Range(i, i + 1).Font.Strikethrough = False
            .Range(i, i + 1).Selected
            .ForceEdit = blnForce
        Else
            If .Range(i - 1, i).Font.Hidden And _
                .Range(i, i + 1).Font.Hidden = False And _
                .Range(i, i + 1).Font.Protected Then
                'B����1����ͨ�ı��������ı���|�������ı����������ı�����ͨ�ı�
                i = i - 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C���⣺�������ı����������ı����������ı���|�������ı����������ı����������ı���
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                ElseIf .Range(i + 1, i + 3) = "E(" And .Range(i, i + 3).Font.Protected And .Range(i, i + 3).Font.ForeColor <> PROTECT_FORECOLOR And _
                    .Range(i + 16, i + 18) = vbCrLf And .Range(i + 16, i + 18).Font.Protected And .Range(i + 16, i + 18).Font.ForeColor <> PROTECT_FORECOLOR Then
                    'D���⣺��ٺ����ͼƬ����֮��û������ʱ���޷�������������
                    i = i + 16
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    If (.Range(i - 16, i - 14) <> "EE") Then
                        .Range(i, i + 1).Selected
                    Else
                        .Range(i + 1, i + 1).Selected
                    End If
                    .ForceEdit = blnForce
                Else
                    .Range(i, i).Selected
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And _
                .Range(i - 1, i).Font.Protected And .Range(i - 1, i).Font.ForeColor <> PROTECT_FORECOLOR And _
                .Range(i, i + 1).Font.Hidden Then
                'B����2����ͨ�ı��������ı����������ı���|�������ı�����ͨ�ı�
                i = i + 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C���⣺�������ı����������ı����������ı���|�������ı����������ı����������ı���
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                Else
                    GoTo LL3
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 2) = vbCrLf And .Range(i, i + 2).Font.Protected And .Range(i, i + 2).Font.ForeColor <> PROTECT_FORECOLOR Then
                mlngHP = -1
                .ForceEdit = True
                .Range(i, i).Font.Protected = False
                .Range(i, i).Font.Hidden = False
                .Range(i - 1, i).Font.ForeColor = vbBlack
                blnSpaceEvent = True
                .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                .Range(i, i + 1).Font.Protected = False
                .Range(i, i + 1).Font.Hidden = False
                .Range(i, i + 1).Font.ForeColor = vbBlack
                If (.Range(i - 16, i - 14) <> "EE") Then
                    .Range(i, i + 1).Selected
                Else
                    .Range(i + 1, i + 1).Selected
                End If
                .ForceEdit = blnForce
            End If
        End If
    End With

LL4:
    If ViewMode = cprNormal Then
        If KeyCode = vbKeyF2 Then
            '��ʾ����Ҫ�ر༭��
            If ViewMode = cprNormal Then Call ShowEleEditor(0, Shift)
        ElseIf KeyCode = vbKeyReturn Then
            If Me.Editor1.Selection.GetType = cprSTPicture Then
                Editor1_DblClick ViewMode
            End If
        ElseIf KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            
        End If
    End If
    Editor1.Tag = ""
End Sub

Private Sub Editor1_KeyUp(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Editor1.ReadOnly Then Exit Sub
    If txtPenInput.Visible And txtPenInput.Enabled Then txtPenInput.SetFocus: Exit Sub
End Sub

'################################################################################################################
'## ���ܣ�  �û���ͼ�޸ı����ı���
'##         �����ǰ������Ҫ�أ��򵯳�����Ҫ�ر༭����
'################################################################################################################
Private Sub Editor1_ModifyProtected(ViewMode As zlRichEditor.ViewModeEnum, bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)
    If Editor1.ReadOnly Then bAllowDoIt = False: Exit Sub
    bAllowDoIt = False

    '���������Ҫ���У��򵯳��༭��
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenAnyKeys(Editor1, lStart, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then
        Select Case sKeyType
        Case "E"
            ShowEleEditor KeyAscii, Shift
        End Select
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵ�ǰλ�ã�ͬ��������ʾ������ڵ㡣
'################################################################################################################
Private Sub Editor1_MouseUp(ViewMode As zlRichEditor.ViewModeEnum, Button As Integer, Shift As Integer, X As Single, y As Single)
    If Editor1.ReadOnly Then Exit Sub
    On Error Resume Next
    If ViewMode <> cprNormal Then Exit Sub
    If mblnFmtBrushDown Then
        '��ʽˢ��Ӧ��
        With Me.Editor1
            .Tag = "Editor1_MouseUp"
            .ForceEdit = True
            Dim lS As Long, lE As Long
            lS = .Selection.StartPos
            lE = .Selection.EndPos
            If lE > lS + 1 Then
                If .Range(lE - 2, lE) = vbCrLf Then
                    If Not mParaFmt Is Nothing Then
                        '���ö�������
                        .Selection.Para.Alignment = mParaFmt.Alignment
                        .Selection.Para.FirstLineIndent = mParaFmt.FirstLineIndent
                        .Selection.Para.LeftIndent = mParaFmt.LeftIndent
                        .Selection.Para.SetLineSpacing mParaFmt.LineSpacingRule, mParaFmt.LineSpacing
                        .Selection.Para.ListAlignment = mParaFmt.ListAlignment
                        .Selection.Para.ListStart = mParaFmt.ListStart
                        .Selection.Para.ListTab = mParaFmt.ListTab
                        .Selection.Para.ListType = mParaFmt.ListType
                        .Selection.Para.RightIndent = mParaFmt.RightIndent
                        .Selection.Para.SpaceAfter = mParaFmt.SpaceAfter
                        .Selection.Para.SpaceBefore = mParaFmt.SpaceBefore
                    End If
                End If
            End If
            If Not mFontFmt Is Nothing Then
                '������������
                .Selection.Font.Bold = mFontFmt.Bold
                .Selection.Font.Italic = mFontFmt.Italic
                .Selection.Font.Name = mFontFmt.Name
                .Selection.Font.Size = mFontFmt.Size
                .Selection.Font.Subscript = mFontFmt.Subscript
                .Selection.Font.Superscript = mFontFmt.Superscript
            End If
            .ForceEdit = False
            .Tag = ""
        End With
        mblnFmtBrushDown = False
        Me.Editor1.OriginRTB.MousePointer = 0
        Exit Sub
    End If

    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bFinded As Boolean, bNeeded As Boolean
    'ͬ����λ���
    Document.HighlightCurCompend Editor1, mfrmCompends.Tree

    If txtPenInput.Visible And txtPenInput.Enabled Then txtPenInput.SetFocus: Exit Sub
    If Editor1.SelLength > 0 Then Exit Sub
    bFinded = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "E", lSS, lSE, lES, lEE, lKey, bNeeded)
    If bFinded Then
        '�������Ԫ���ڲ������ʾѡ��ĳ��ѡ��
        If Me.Document.Elements("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 Or Me.Document.Elements("K" & lKey).��ֹ�� > 0 Then Exit Sub
        If Me.Document.Elements("K" & lKey).������̬ = 1 Then
            'չ����ʽ��Ҫ��¼��     '������
            Dim strTmp As String, p As Long, P1 As Long, P2 As Long, blnForce As Boolean, lSP As Long
            With Editor1
                blnForce = .ForceEdit
                .Tag = "Editor1_MouseUp"
                .Freeze
                .ForceEdit = True
                strTmp = .Range(lSE, lES)
                p = .Selection.StartPos: lSP = p - lSE '����λ�ã�����λ�ú͹ؼ���������
                If Me.Document.Elements("K" & lKey).Ҫ�ر�ʾ = 2 Then '��ѡ
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "��", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "��", P2)
                    If P1 > P2 And P1 > 0 Then
                        strTmp = Replace(strTmp, "��", "��")
                        Mid(strTmp, P1, 1) = "��"
                        .Range(lSE, lES) = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        strTmp = Replace(strTmp, "��", "��")
                        Mid(strTmp, P2, 1) = "��"
                        .Range(lSE, lES) = strTmp
                    End If
                    
                    If Me.Document.Elements("K" & lKey).��̬�� = 1 Then '�Զ�̬��ĵ�������
                        If InStrRev(strTmp, "��") > InStrRev(strTmp, "��") Then '���һ��ѡ�ѡ��
                            Dim strSin As String
                            strSin = Trim(InputBox("��¼���Զ���Ҫ��ѡ��" & vbCrLf & "������볤��200������", "�������"))
                            If strSin <> "" Then
                                Me.Document.Elements("K" & lKey).�����ı� = Mid(strTmp, 1, InStrRev(strTmp, "��")) & strSin
                            Else
                                Me.Document.Elements("K" & lKey).�����ı� = Mid(strTmp, 1, InStrRev(strTmp, "��") - 1) & "���Զ���"
                            End If
                        Else '���һ��û�б�ѡ��,������ ���Զ���
                            Me.Document.Elements("K" & lKey).�����ı� = Mid(strTmp, 1, InStrRev(strTmp, "��")) & "�Զ���"
                        End If
                        Me.Document.Elements("K" & lKey).Refresh Editor1
                    End If
                ElseIf Me.Document.Elements("K" & lKey).Ҫ�ر�ʾ = 3 Then '��ѡ
                    P1 = .Selection.StartPos - lSE + 1
                    P1 = InStrRev(strTmp, "��", P1)
                    P2 = .Selection.StartPos - lSE + 1
                    P2 = InStrRev(strTmp, "��", P2)
                    If P1 > P2 And P1 > 0 Then
                        Mid(strTmp, P1, 1) = "��"
                        .Range(lSE, lES) = strTmp
                    ElseIf P2 > P1 And P2 > 0 Then
                        Mid(strTmp, P2, 1) = "��"
                        .Range(lSE, lES) = strTmp
                    End If
                    
                    If Me.Document.Elements("K" & lKey).��̬�� = 1 And p > InStrRev(strTmp, "��") + lSE - 2 Then '�Զ�̬��ĵ��������������λ�õ��
                        If InStrRev(strTmp, "��") > InStrRev(strTmp, "��") Then '���һ��ѡ�ѡ��
                            Dim strMul As String
                            strMul = Trim(InputBox("��¼���Զ���Ҫ��ѡ��" & vbCrLf & "������볤��200������", "�������"))
                            If strMul <> "" Then
                                Me.Document.Elements("K" & lKey).�����ı� = Mid(strTmp, 1, InStrRev(strTmp, "��")) & strMul
                            Else
                                Me.Document.Elements("K" & lKey).�����ı� = Mid(strTmp, 1, InStrRev(strTmp, "��") - 1) & "���Զ���"
                            End If
                        Else '���һ��û�б�ѡ��,������ ���Զ���
                            Me.Document.Elements("K" & lKey).�����ı� = Mid(strTmp, 1, InStrRev(strTmp, "��")) & "�Զ���"
                        End If
                        Me.Document.Elements("K" & lKey).Refresh Editor1
                    End If
                End If
                
                Call FindKey(Editor1, "E", lKey, lSS, lSE, lES, lEE, bNeeded)
                strTmp = .Range(lSE, lES)
                If (InStr(strTmp, "��") = 0 And Me.Document.Elements("K" & lKey).Ҫ�ر�ʾ = 2) _
                    Or (InStr(strTmp, "��") = 0 And Me.Document.Elements("K" & lKey).Ҫ�ر�ʾ = 3) Then '��ûѡ���κ�ѡ������Ӳ�����
                    .Range(lSE, lES).Font.Underline = cprWave
                Else
                    .Range(lSE, lES).Font.Underline = cprNone
                End If
                Me.Document.Elements("K" & lKey).�����ı� = strTmp
                
                Call CheckElementLimit(lKey)
                
                Call FindKey(Editor1, "E", lKey, lSS, lSE, lES, lEE, bNeeded)
                lSP = lSE + lSP
                
                .Range(lSP, lSP).Selected
                .ForceEdit = blnForce
                .UnFreeze
                .Tag = ""
            End With
        End If
    End If
End Sub
'################################################################################################################
'## ���ܣ�  �û���Tab���Ĵ���
'##
'## ˵����  ��ǰ������Ҫ�أ���������һ������Ҫ��λ�ô���
'################################################################################################################
Private Sub Editor1_PressTabKey()
    If Editor1.ReadOnly Then Exit Sub
    If Editor1.ViewMode = cprNormal Then
        Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBetweenKeys As Boolean, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
        bBetweenKeys = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "E", lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBetweenKeys Then
            Call AddUndoPoint  '�ֶ�����
            bFinded = FindNextKey(Editor1, Editor1.Selection.StartPos + 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                Editor1.Range(lKSE, lKES - Len(Me.Document.Elements("K" & lKey).Ҫ�ص�λ)).Selected
            Else
                bFinded = FindNextKey(Editor1, 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded Then
                    Editor1.Range(lKSE, lKES - Len(Me.Document.Elements("K" & lKey).Ҫ�ص�λ)).Selected
                End If
            End If
            Call ClearNoUseUndoList
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ�  �Ҽ��˵�����
'################################################################################################################
Private Sub Editor1_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, y As Single)
    If Editor1.ReadOnly Then Exit Sub
    Dim Popup As CommandBar
    Dim Control As CommandBarControl, bFinded As Boolean, bOK As Boolean
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean

    Set Popup = cbrThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal Then
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys = False Then Exit Sub
                        If sType = "P" Then
                                Set Control = .Add(xtpControlButton, ID_EDIT_MARKEDPIC, "����޸�(&M)")
                                Control.BeginGroup = True
                                If Me.Document.Pictures("K" & lKey).PictureType = EPRMarkedPicture Then
                                        .Add xtpControlButton, ID_EDIT_OUTERPIC, "��ͼ����(&D)"
                                End If
                                Popup.ShowPopup
                        End If
        Else
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys Then
                If sType = "D" Then
                    '���
                    Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "ɾ��(&D)"): Control.BeginGroup = True
                    Set Control = .Add(xtpControlButton, conMenu_Tool_Reference, "������ϲο�(&R)..."): Control.BeginGroup = True
                    Popup.ShowPopup
                ElseIf sType = "E" Then
                    '����Ҫ��
                    bFinded = FindPrevKey(Editor1, Editor1.Selection.StartPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
                    If bFinded Then
                        If Me.Document.Compends("K" & lKey).�������ID <> 0 And Me.Editor1.SelLength > 0 Then bOK = True
                    End If
                    If mfrmModElement.Visible Then Exit Sub     '����ڱ༭�У���ô���ܵ����˵�
                    Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "����(&X)")
                    Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
                    Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "ճ��(&V)    ")
                    Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "ɾ��(&D)")
                    If bOK Then
                        Set Control = .Add(xtpControlButton, ID_EDIT_SAVEASPHRASE, "��Ϊʾ���ʾ�(&S)..."): Control.BeginGroup = True
                    End If
                    bFinded = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
                    If bFinded Then
                        If sType = "E" Then
                            If Me.Document.Elements("K" & lKey).������̬ = 0 Then
                                
                                Set Control = .Add(xtpControlButton, ID_ELEMENT_UPDATE, "��������(&U)"): Control.BeginGroup = True
                                
                                If Me.Document.Elements("K" & lKey).�����ı� <> "" Then
                                '��չ����Ҫ�أ��������Ҫ�����ݣ���Ҫ�����ֵ���Ŀ���ı���գ�
                                    Set Control = .Add(xtpControlButton, ID_ELEMENT_CLEAR, "�������(&R)")
                                End If
                                
                            End If
                            If Me.Document.Elements("K" & lKey).������̬ = 0 And InStr(1, gstrPrivsEpr, "�����ı�����") > 0 Then
                                Set Control = .Add(xtpControlButton, ID_ELEMENT_TOSTRING, "תΪ�ı�(&T)"): Control.BeginGroup = True
                            End If
                        End If
                    End If
                    Popup.ShowPopup
                End If
            Else
'                Set Control = .Add(xtpControlButton, ID_EDIT_UNDO, "����(&U)")
'                Control.BeginGroup = True
'                .Add xtpControlButton, ID_EDIT_REDO, "����(&R)"
                bFinded = FindPrevKey(Editor1, Editor1.Selection.StartPos + 1, "O", lKey, lSS, lSE, lES, lEE, bNeeded)
                If bFinded Then
                    If Me.Document.Compends("K" & lKey).�������ID <> 0 And Me.Editor1.SelLength > 0 Then bOK = True
                End If
                Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "����(&X)")
                Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
                Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "ճ��(&V)    ")
                Set Control = .Add(xtpControlButton, ID_EDIT_DELETE, "ɾ��(&D)")
                Set Control = .Add(xtpControlButton, ID_EDIT_COPYSELF, "ר�ø���(&I)"): Control.BeginGroup = True
                Set Control = .Add(xtpControlButton, ID_EDIT_COPYOUT, "���Ƶ�ճ����(&U)")
                If bOK Then
                    Set Control = .Add(xtpControlButton, ID_EDIT_SAVEASPHRASE, "��Ϊʾ���ʾ�(&S)..."): Control.BeginGroup = True
                End If
                Set Control = .Add(xtpControlButton, ID_EDIT_SELECTALL, "ȫѡ(&A)"): Control.BeginGroup = True
                Popup.ShowPopup
            End If
        End If
    End With
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵ�ǰλ�ã���ʾ��ǰ�С���λ�á�
'################################################################################################################
Private Sub Editor1_SelChange(ViewMode As zlRichEditor.ViewModeEnum, ByVal lStart As Long, ByVal lEnd As Long)
    If Me.Document Is Nothing Then Exit Sub
    
    Dim COLOR As OLE_COLOR
    If Me.Document.EditType = cprET_��������� Then
        COLOR = IIf(Editor1.Selection.Font.ForeColor = tomAutoColor, vbBlack, Editor1.Selection.Font.ForeColor)
        If COLOR = tomUndefined Then
            stbThis.Panels(5).Text = "����ı�"
        Else
            If Get��ֹ��(COLOR) > 0 Then
                'ɾ���ı�
                stbThis.Panels(5).Text = "��" & Get��ֹ��(COLOR) + 1 & "��ɾ��"
            ElseIf Get��ʼ��(COLOR) > 0 Then
                '�����ı�
                stbThis.Panels(5).Text = "��" & Get��ʼ��(COLOR) & "������"
            Else
                stbThis.Panels(5).Text = ""
            End If
        End If
    End If

    On Error Resume Next
    If Editor1.InProcessing Then Exit Sub
    stbThis.Panels(2).Text = Editor1.CurrentLine & " ��,  " & Editor1.CurrentColumn & " ��,  ��" & Editor1.LineCount & " ��"
    If mblnAutoPageCount Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & ",  �� " & Me.Editor1.PageCount & " ҳ"
    If Editor1.Tag = "" Then
        Document.HighlightCurCompend Editor1, mfrmCompends.Tree
        'ˢ��ʾ���ʾ�
        If Not mfrmCompends.Tree.SelectedItem Is Nothing Then
            Call RefSentenceList
        End If
    End If

    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    If Editor1.Tag = "" And Editor1.AuditMode = False Then
        If tblThis.Visible Then
            Editor1.CloseUIInterface
        ElseIf ucPictureEditor1.Visible Then
            Editor1.CloseUIInterface
        ElseIf ucPacsImgCanvas1.Visible Then
            Editor1.CloseUIInterface
        End If
        If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False Then
            
            '���ҹؼ��� ID ��
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys = False Then Exit Sub

            If sType = "T" Then
                '�༭���
                If Me.Document.Tables("K" & lKey).TableType = tte_����ͼƬ�� Then
                    '����ͼƬ��
                    DkpThis.ShowPane ID_VIEW_PACSPIC
                    tblThis.Tag = lKey
                    Editor1.ShowUIInterface
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
                ElseIf Me.Document.Tables("K" & lKey).Ԥ�����ID = 0 Then
                    If Me.Document.Tables("K" & lKey).TableType = tte_ҽ����Ŀ�� Then
                    Else
                        '��ȡ���ݵ����ؼ��У�
                        tblThis.Tag = lKey
                    End If
                    Editor1.ShowUIInterface
                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
                End If
            ElseIf sType = "P" Then
                '�༭ͼƬ
                Editor1.ShowUIInterface
            End If
        End If
    ElseIf Editor1.Tag = "" And Editor1.AuditMode = True And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False Then
        If tblThis.Visible Then Editor1.CloseUIInterface
        If Editor1.Selection.GetType = cprSTPicture And Editor1.ViewMode = cprNormal And Editor1.ReadOnly = False Then
            '���ҹؼ��� ID ��
            bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
            If bInKeys = False Then Exit Sub
            If sType <> "T" Then Exit Sub
            If Me.Document.Tables("K" & lKey).TableType = tte_����ͼƬ�� Then Exit Sub
            If Me.Document.Tables("K" & lKey).TableType = tte_ҽ����Ŀ�� Then Exit Sub
            If Me.Document.Tables("K" & lKey).Ԥ�����ID = 0 Then
                '��ȡ���ݵ����ؼ��У�
                tblThis.Tag = lKey
                Editor1.ShowUIInterface
                Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
            End If
        End If
    End If
End Sub

Private Function ReadTableToUI(ByRef oTable As cEPRTable) As Boolean
    On Error Resume Next
    If oTable Is Nothing Then ReadTableToUI = False: Exit Function
'    Dim t1 As Date, T2 As Date
'    t1 = Timer

    Dim i As Long, j As Long, strMerge As String, R1 As Long, C1 As Long, R2 As Long, C2 As Long
    Dim T As Variant, strColWidth As String, lKey As Long

    tblThis.Redraw = False
    tblThis.SingleClickEdit = False
    tblThis.HighlightMode = HMFilledRectAlpha
    tblThis.BorderWidth = oTable.BorderWidth
    tblThis.AutoHeight = oTable.AutoHeight
    tblThis.Init oTable.Rows, oTable.Cols
    tblThis.ExtendTag = oTable.ExtendTag
    tblThis.UserTag = oTable.���
    strColWidth = oTable.ColWidthString
    T = Split(strColWidth, "|")
    On Error Resume Next
    If UBound(T) = -1 Then
        If oTable.Rows > 0 Then
            For i = 1 To oTable.Cols
                tblThis.ColWidth(i) = oTable.Cell(1, i).Width
            Next
        End If
    Else
        For i = 0 To UBound(T)
            tblThis.ColWidth(i + 1) = Val(T(i))
        Next
    End If

    For i = 1 To oTable.Rows
        tblThis.ROWHEIGHT(i) = oTable.Cell(i, 1).Height
        For j = 1 To oTable.Cols
            lKey = tblThis.CellKey(i, j)
            With oTable.Cell(i, j)
                tblThis.Cells("K" & lKey).Text = oTable.Cell(i, j).�����ı�
                tblThis.Cells("K" & lKey).Margin = .Margin
                tblThis.Cells("K" & lKey).Width = .Width
                tblThis.Cells("K" & lKey).Height = .Height
'                tblThis.Cells("K" & lKey).MergeInfo = .MergeNo
                tblThis.Cells("K" & lKey).SingleLine = .SingleLine
                tblThis.Cells("K" & lKey).ForeColor = .ForeColor
                tblThis.Cells("K" & lKey).BackColor = .BackColor
                tblThis.Cells("K" & lKey).GridLineColor = .GridLineColor
                tblThis.Cells("K" & lKey).GridLineWidth = .GridLineWidth
                tblThis.Cells("K" & lKey).FixedWidth = .FixedWidth
                tblThis.Cells("K" & lKey).AutoHeight = .AutoHeight
                tblThis.Cells("K" & lKey).FontName = .FontName
                tblThis.Cells("K" & lKey).FontSize = .FontSize
                tblThis.Cells("K" & lKey).FontBold = .FontBold
                tblThis.Cells("K" & lKey).FontItalic = .FontItalic
                tblThis.Cells("K" & lKey).FontStrikeout = .FontStrikeout
                tblThis.Cells("K" & lKey).FontUnderline = .FontUnderline
                tblThis.Cells("K" & lKey).FontWeight = .FontWeight
                tblThis.Cells("K" & lKey).FormatString = .FormatString
                tblThis.Cells("K" & lKey).Indent = .Indent
                tblThis.Cells("K" & lKey).HAlignment = .HAlignment
                tblThis.Cells("K" & lKey).VAlignment = .VAlignment
                tblThis.Cells("K" & lKey).Protected = .Protected
                If oTable.Cell(i, j).ElementKey > 0 Then
                    tblThis.Cells("K" & lKey).ToolTipText = oTable.Elements("K" & oTable.Cell(i, j).ElementKey).Ҫ������
                    tblThis.Cells("K" & lKey).Tag = oTable.Cell(i, j).ElementKey
                End If
                If .PictureKey > 0 Then
                    oTable.Pictures("K" & .PictureKey).Row = i
                    oTable.Pictures("K" & .PictureKey).Col = j
                    Set tblThis.Cells("K" & lKey).Picture = oTable.Pictures("K" & .PictureKey).DrawFinalPic(oTable)
                    tblThis.Cells("K" & lKey).Tag = oTable.Cell(i, j).PictureKey
                End If
            End With
        Next
    Next

    For i = 1 To oTable.Cells.Count
        strMerge = oTable.Cells(i).MergeNo              '�ָ���Ԫ��ĺϲ�
        If strMerge <> "" Then
            R1 = Val(Left(strMerge, 4))
            C1 = Val(Mid(strMerge, 5, 4))
            R2 = Val(Mid(strMerge, 9, 4))
            C2 = Val(Mid(strMerge, 13))
            tblThis.MergeCells R1, C1, R2, C2, False
        End If
    Next

    tblThis.ShowToolTipText = True
    tblThis.MinRowHeight = 300
    tblThis.Redraw = True
    tblThis.Refresh
    tblThis.FixCellsWidth
    If (Not tblThis.AutoHeight) Then tblThis.Height = oTable.Height

    Editor1.ResizeUIInterface tblThis.Width, tblThis.Height
    tblThis.Refresh True, False

'    T2 = Timer
'    Debug.Print "��ȡ��ʱ��" & Format(T2 - t1, "0.00000000") & ",��Ԫ��������" & tblThis.Cells.Count
End Function

Private Function SaveUIToTable(ByRef oTable As cEPRTable, Optional ByVal bFirst As Boolean) As Boolean
    If oTable Is Nothing Then SaveUIToTable = False: Exit Function
    Dim strColWidth As String

    Dim i As Long, j As Long, lKey As Long
    For i = 1 To tblThis.ColCount
        If i = 1 Then
            strColWidth = tblThis.ColWidth(i)
        Else
            strColWidth = strColWidth & "|" & tblThis.ColWidth(i)
        End If
    Next

    oTable.Rows = tblThis.RowCount
    oTable.Cols = tblThis.ColCount
    oTable.ColWidthString = strColWidth
'    oTable.Width = tblThis.Width
    oTable.Height = tblThis.Height
    oTable.SingleLine = tblThis.SingleLine
    oTable.AlternateRowBackColor = tblThis.AlternateRowBackColor
    oTable.BackColor = tblThis.BackColor
    oTable.GridLineColor = tblThis.GridLineColor
    oTable.GridLineWidth = tblThis.GridLineWidth
    oTable.BorderColor = tblThis.BorderColor
    oTable.BorderWidth = tblThis.BorderWidth
    oTable.ForeColor = tblThis.ForeColor
    oTable.FontQuality = tblThis.FontQuality
    oTable.AutoHeight = tblThis.AutoHeight
    oTable.WordEllipsis = tblThis.WordEllipsis
    oTable.CellMargin = tblThis.CellMargin
    oTable.CellIndent = tblThis.CellIndent
    oTable.ExtendTag = tblThis.ExtendTag
    oTable.��� = tblThis.UserTag
    For i = 1 To oTable.Rows
        For j = 1 To oTable.Cols
            If oTable.Cell(i, j) Is Nothing Then
                lKey = 0
                Call oTable.Cells.Add(lKey, i, j)
                oTable.Cell(i, j).ID = 0
                oTable.Cell(i, j).��ʼ�� = Me.Document.Ŀ��汾
            End If
            lKey = tblThis.CellKey(i, j)
            If lKey > 0 And Not oTable.Cell(i, j) Is Nothing Then
                With oTable.Cell(i, j)
                    If .�����ı� <> tblThis.Cells("K" & lKey).Text And oTable.TableType = tte_Ĭ�� And Me.Document.EditType = cprET_��������� Then
                        If .��ʼ�� <> Me.Document.Ŀ��汾 And .ID <> 0 Then .ID = 0
                        .��ʼ�� = Me.Document.Ŀ��汾
                    End If
                    .�����ı� = tblThis.Cells("K" & lKey).Text
                    .Margin = tblThis.Cells("K" & lKey).Margin
'                    .Width = tblThis.Cells("K" & lKey).Width
'                    .Height = tblThis.Cells("K" & lKey).Height
                    .Width = tblThis.ColWidth(j)
                    .Height = tblThis.ROWHEIGHT(i)
                    .MergeNo = tblThis.Cells("K" & lKey).MergeInfo
                    .SingleLine = tblThis.Cells("K" & lKey).SingleLine
                    .ForeColor = tblThis.Cells("K" & lKey).ForeColor
                    .BackColor = tblThis.Cells("K" & lKey).BackColor
                    .GridLineColor = tblThis.Cells("K" & lKey).GridLineColor
                    .GridLineWidth = tblThis.Cells("K" & lKey).GridLineWidth
                    .FixedWidth = tblThis.Cells("K" & lKey).FixedWidth
                    .AutoHeight = tblThis.Cells("K" & lKey).AutoHeight
                    .FontName = tblThis.Cells("K" & lKey).FontName
                    .FontSize = tblThis.Cells("K" & lKey).FontSize
                    .FontBold = tblThis.Cells("K" & lKey).FontBold
                    .FontItalic = tblThis.Cells("K" & lKey).FontItalic
                    .FontStrikeout = tblThis.Cells("K" & lKey).FontStrikeout
                    .FontUnderline = tblThis.Cells("K" & lKey).FontUnderline
                    .FontWeight = tblThis.Cells("K" & lKey).FontWeight
                    .FormatString = tblThis.Cells("K" & lKey).FormatString
                    .Indent = tblThis.Cells("K" & lKey).Indent
                    .VAlignment = tblThis.Cells("K" & lKey).VAlignment
                    .HAlignment = tblThis.Cells("K" & lKey).HAlignment
                    .Protected = tblThis.Cells("K" & lKey).Protected
                    If tblThis.Cells("K" & lKey).Picture Is Nothing Then
                        .ElementKey = Val(tblThis.Cells("K" & lKey).Tag)
                        .PictureKey = 0
                    Else
                        .ElementKey = 0
                        .PictureKey = Val(tblThis.Cells("K" & lKey).Tag)
                    End If
                End With
            End If
        Next j
    Next i

    '����Ҫ����ͼƬ��Ȳ��ܳ���ҳ����
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, bFinded As Boolean, bNeeded As Boolean
    Dim lW As Long
    bFinded = FindKey(Editor1, "T", oTable.Key, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        lW = Me.Editor1.PaperWidth - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - Me.ScaleX(Me.Editor1.Range(lSE, lES).Para.LeftIndent + Me.Editor1.Range(lSE, lES).Para.FirstLineIndent, vbPixels, vbTwips) - 130
        picTMP.Width = IIf(tblThis.Width > lW, lW, tblThis.Width)
    Else
        picTMP.Width = tblThis.Width
    End If
    picTMP.Height = tblThis.Height

    oTable.Width = picTMP.Width '����ʵ�ʱ���ȣ���������

    tblThis.DrawToDC picTMP.hDC
    picTMP.Picture = picTMP.Image
    If bFirst Then
        Dim frmT As New frmTablePicCreator
        Me.Document.Tables("K" & tblThis.Tag).InsertIntoEditor Me.Editor1, , frmT.GetFinalPic(Me.Document.Tables("K" & tblThis.Tag)), True
        Unload frmT
        Set frmT = Nothing
    Else
        oTable.Refresh Editor1, picTMP.Picture, True
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next
    If tblThis.Visible Then
        tblThis.SetFocus
    ElseIf ActiveControl Is Editor1 Then
        If Editor1.Visible And Editor1.Enabled Then Editor1.SetFocus
    End If
     
    Err.Clear
End Sub

'################################################################################################################
'## ���ܣ�  �����ʼ��
'################################################################################################################
Private Sub Form_Load()
    Dim i As Long, j As Long

    mblnAutosave = (zlDatabase.GetPara("AutoSave", glngSys, 1070, 1) = "1")
    mlngUndoLimit = zlDatabase.GetPara("UndoLimit", glngSys, 1070, 20)
    mlngSaveInterval = zlDatabase.GetPara("SaveInterval", glngSys, 1070, 60)
    mblnAutoSaveEPR = (zlDatabase.GetPara("AutoSaveEPR", glngSys, 1070, 0) = "1")
    mlngSaveIntervalEPR = zlDatabase.GetPara("SaveIntervalEPR", glngSys, 1070, 5)
    mblnAutoPageCount = (zlDatabase.GetPara("AutoPageCount", glngSys, 1070, 0) = "1")
    mblnAutoPageNote = (zlDatabase.GetPara("AutoPageNote", glngSys, 1070, 0) = "1")
    mintSharePages = zlDatabase.GetPara("SharePageCount", glngSys, 1070, 5)
    mblnSignAutoAlter = (zlDatabase.GetPara("ǩ���Զ�λ��", glngSys, 1070, 0) = "1")
            
    ReDim UndoList(1 To 1) As UndoInfo
    p_Undo = 0
    mblnChange = False


    '## �˵���ʼ��
    Dim cbrMenu As CommandBarPopup                      '���˵�
    Dim cbpPopup As CommandBarPopup                     '��������
    Dim cbpPopupSub As CommandBarPopup                  '��ʱ����
    Dim objControl As CommandBarControl                 '�������ؼ�
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�
    Dim Combo As CommandBarComboBox                     '������������ؼ�
    Dim cbrBar As CommandBar                           '������
    Dim cbrCustom As CommandBarControlCustom

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbrThis.Icons = gfrmPublic.ImageManager.Icons
    cbrThis.Options.ShowExpandButtonAlways = False
    cbrThis.EnableCustomization (False)
    cbrThis.Options.UseDisabledIcons = True
    cbrThis.Options.AlwaysShowFullMenus = True
    cbrThis.StatusBar.Visible = False
    cbrThis.ActiveMenuBar.Title = "�˵���"

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�ļ�(&F)"): cbrMenu.ID = ID_Main_FILE
    With cbrMenu.CommandBar.Controls
        .Add xtpControlButton, ID_FILE_CLEAR, "���(&C)"

        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "����(&S)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "�����˳�(&Q)")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVEASEPRDEMO, "���Ϊ����(&D)...")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVEASSEGMENT, "���ΪƬ��(&G)...")
        
        Set cbpPopup = .Add(xtpControlPopup, 0, "����(&A)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILE_SAVEAS, "����Ϊ&RTF�ļ�..."
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FILE_EXPORTTOXML, "����ΪX&ML�ļ�..."

        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORTFROMXML, "��XM&L�ļ�����")
        Set objControl = .Add(xtpControlButton, ID_FILE_PAGESETUP, "ҳ������(&U)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILE_PRINTPREVIEW, "��ӡԤ��(&V)"
        .Add xtpControlButton, ID_FILE_PRINT, "��ӡ(&P)..."
        .Add xtpControlButton, ID_FILE_PRINTINWORD, "ͨ��Word��ӡ(&W)"

        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&T)"): objControl.BeginGroup = True

        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "�˳�(&X)")
        objControl.BeginGroup = True
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "�༭(&E)"): cbrMenu.ID = ID_Main_EDIT
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "����(&U)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_REDO, "����(&R)"
'
        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "����(&X)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_COPY, "����(&C)"
        
        .Add xtpControlButton, ID_EDIT_PASTE, "ճ��(&V)"
        Set objControl = .Add(xtpControlButton, ID_EDIT_COPYSELF, "ר�ø���(&I)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_COPYOUT, "���Ƶ�ճ����(&U)"
    
        Set cbpPopup = .Add(xtpControlPopup, 0, "���(&M)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_REFCOMPEND, "ˢ�����(&R)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_ADDCOMPEND, "�������(&A)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_DELCOMPEND, "ɾ�����(&D)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_EDIT_MODCOMPEND, "�޸����(&M)"

        Set cbpPopup = .Add(xtpControlPopup, 0, "ǩ�����޶�(&S)")
        cbpPopup.BeginGroup = True
'            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_PATISIGN, "����ǩ��(&S)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_SIGN, "ǩ��(&S)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_UNTREAD, "����(&C)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_SIGN_QUIT, "ǩ���˳�(&Q)"): objControl.STYLE = xtpButtonIconAndCaption
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_REVISION_PREV, "ǰһ���޶�(&P)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_REVISION_NEXT, "��һ���޶�(&N)")
            Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_REVISION_RESET, "ȡ����ѡ�޶�(&E)")

        Set objControl = .Add(xtpControlButton, ID_EDIT_DELETE, "ɾ��(&D)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_SELECTALL, "ȫѡ(&A)"

        Set objControl = .Add(xtpControlButton, ID_EDIT_FIND, "����(&F)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_FINDNEXT, "������һ��(&N)"
        .Add xtpControlButton, ID_EDIT_REPLACE, "�滻(&R)..."
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "��ͼ(&V)"): cbrMenu.ID = ID_Main_VIEW
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_VIEW_STRUCTURE, "�ĵ��ṹͼ(&D)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_VIEW_PHRASEDEMO, "ʾ���ʾ��б�(&S)"
        .Add xtpControlButton, ID_VIEW_SEGMENT, "ʾ��Ƭ���б�(&G)"
        .Add xtpControlButton, ID_VIEW_PACSPIC, "����ͼ�б�(&P)"
        .Add xtpControlButton, ID_VIEW_HISTORYWINDOW, "��ʷ�����б�(&P)"
        .Add xtpControlButton, ID_VIEW_HISTORYREPORT, "��ʷ�����б�(&R)"
        
        Set cbpPopup = .Add(xtpControlPopup, 0, "������(&T)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, "�������б�"

        Set objControl = .Add(xtpControlButton, ID_VIEW_HEADFOOT, "ҳüҳ��(&H)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_VIEW_CHARCOUNT, "����ͳ��(&N)..."
'        .Add xtpControlButton, ID_VIEW_GRID, "������(&G)"
        Set objControl = .Add(xtpControlButton, ID_VIEW_RULER, "���(&R)")
        objControl.Checked = True
        .Add xtpControlButton, ID_VIEW_PENWINDOW, "��д���봰��(&W)"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&I)"): cbrMenu.ID = ID_Main_INSERT
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "���ں�ʱ��(&D)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_SPECIALCHAR, "�������(&S)..."

        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "ͼƬ(&P)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_TABLE_INSERTTABLE, "���(&T)"
        .Add xtpControlButton, ID_INSERT_ELEMENT, "Ҫ��(&E)"
        .Add xtpControlButton, ID_EDIT_ADDCOMPEND, "���(&C)"
        .Add xtpControlButton, ID_INSERT_PACSPIC, "����ͼ(&R)"
        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "��ʷ�ļ�(&H)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_EPRDEMO, "���뷶��(&F)..."
        .Add xtpControlButton, ID_INSERT_DOCADVISE, "����ҽ��(&A)"
        .Add xtpControlButton, ID_DIAGNOSIS, "���(&D)"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "��ʽ(&O)"): cbrMenu.ID = ID_Main_FORMAT
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_FORMAT_FONT, "����(&F)...")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FORMAT_PARA, "����(&P)..."
        Set cbpPopup = .Add(xtpControlPopup, ID_FORMAT_BACKGROUND, "����ɫ(&K)")
        cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, ID_FORMAT_BACKGROUND, "")
        objCustControl.Handle = ColorPaperBackColor.hwnd

        Set objControl = .Add(xtpControlButton, ID_FORMAT_BOLD, "����(&B)")
        objControl.BeginGroup = True
'        .Add xtpControlButton, ID_FORMAT_ITALIC, "б��(&I)"
        .Add xtpControlButton, ID_FORMAT_SUPER, "�ϱ�(&R)"
        .Add xtpControlButton, ID_FORMAT_SUB, "�±�(&S)"
        .Add xtpControlButton, ID_FORMAT_PROTECT, "����(&P)"

        Set cbpPopup = .Add(xtpControlPopup, ID_FORMAT_UNDERLINE, "�»���"): cbpPopup.ID = ID_FORMAT_UNDERLINE
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.CommandBar.SetIconSize 60, 8
        cbpPopup.CommandBar.Width = 60
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_NONE, "<���»���>"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THIN, "ϸ��"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THICK, "����"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_WAVE, "������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DOT, "����"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASH, "����"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT, "�㻮��"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT2, "˫�㻮��"

        Set cbpPopup = .Add(xtpControlPopup, 0, "���뷽ʽ(&A)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_ALIGNLEFT, "�����(&L)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_ALIGNCENTER, "���ж���(&C)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_ALIGNRIGHT, "�Ҷ���(&R)"

        Set cbpPopup = .Add(xtpControlPopup, 0, "��Ŀ��������(&E)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTNONE, "��"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTBULLETS, "��Ŀ����(��)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTARABIC, "����������(1,2,3,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTLCHAR, "Сд��ĸ(a,b,c,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTUCHAR, "��д��ĸ(A,B,C,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTLROME, "Сд��������(i,ii,iii,...)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LISTUROME, "��д��������(I,II,III,...)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_LISTSETUP, "�Զ����ʽ...")
        objControl.BeginGroup = True

        Set cbpPopup = .Add(xtpControlPopup, ID_FORMAT_SPACE, "���(&L)"): cbpPopup.ID = ID_FORMAT_SPACE
        Set cbpPopupSub = cbpPopup.CommandBar.Controls.Add(xtpControlSplitButtonPopup, ID_FORMAT_LINESPACE, "�м��(&L)")
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE1, "1.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE2, "1.3"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE3, "1.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE4, "2.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE5, "2.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE6, "3.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE7, "����..."

        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_SPACEBEFORE, "��ǰ���(&B)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_SPACEAFTER, "�κ���(&A)"

        Set cbpPopup = .Add(xtpControlPopup, 0, "����(&D)")
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_FIRSTINDENT, "��������(&F)"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_FIRSTHUNGING, "��������(&H)"
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_FORMAT_INDENTDECREASE, "����������(&D)")
        objControl.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_INDENTINCREASE, "����������(&I)"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "���(&T)"): cbrMenu.ID = ID_Main_TABLE
    With cbrMenu.CommandBar.Controls
        Set cbpPopup = .Add(xtpControlPopup, 0, "����(&I)")
'        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTTABLE, "���(&T)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLLEFT, "��(�����)(&L)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTCOLRIGHT, "��(���Ҳ�)(&R")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWUP, "��(���Ϸ�)(&A)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTROWDOWN, "��(���·�)(&B)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_INSERTINHERITROW, "����̳���(&I)")

        Set cbpPopup = .Add(xtpControlPopup, 0, "ɾ��(&D)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETETABLE, "���(&T)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETECOL, "��(&C)"): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_DELETEROW, "��(&R)")

        Set cbpPopup = .Add(xtpControlPopup, 0, "��ʽ(&F)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_FORMATCELL, "��Ԫ��(&E)..."): objControl.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_SAMECOLWIDTH, "��ͬ�п�(&C)")
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_TABLE_MERGE, "�ϲ�(&M)"): objControl.BeginGroup = True
        Set cbpPopup = cbpPopup.CommandBar.Controls.Add(xtpControlPopup, ID_TABLE_CELLALIGNMENT, "��Ԫ����뷽ʽ")
        cbpPopup.CommandBar.SetTearOffPopup "��Ԫ����뷽ʽ", ID_TABLE_CELLALIGNMENT, 100
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Width = 70
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "���������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "���Ͼ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "�����Ҷ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "�в������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "�в�����"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "�в��Ҷ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "���������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "���¾���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "�����Ҷ���"
    End With

    Set cbrMenu = cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "����(&H)"): cbrMenu.ID = ID_Main_HELP
    With cbrMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "��������(&H)")
        objControl.BeginGroup = True
        Set cbpPopupSub = .Add(xtpControlPopup, 0, "&Web�ϵ�" & gstrProductName)
        objControl.BeginGroup = True
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_ONLINE, gstrProductName & "��ҳ(&H)"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_WEBFORUM, gstrProductName & "��̳(&F)"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_HELP_CONTACT, "���ͷ���(&M)"
        Set objControl = .Add(xtpControlButton, ID_HELP_ABOUT, "����(&A)...")
        objControl.BeginGroup = True
    End With

    '## ��������ʼ��

    Set cbrBar = cbrThis.Add("����", xtpBarTop): cbrBar.BarId = ID_BAR_NORMAL
    cbrBar.EnableDocking xtpFlagStretched
    With cbrBar.Controls
        Set objControl = .Add(xtpControlButton, ID_FILE_CLEAR, "���")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE, "����")
        Set objControl = .Add(xtpControlButton, ID_FILE_SAVE_QUIT, "�����˳�")
        

        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "��ӡ")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_FILE_PRINTPREVIEW, "��ӡԤ��"

        Set objControl = .Add(xtpControlButton, ID_EDIT_CUT, "����")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_COPY, "����"
        .Add xtpControlButton, ID_EDIT_COPYSELF, "ר�ø���"
        .Add xtpControlButton, ID_EDIT_PASTE, "ճ��"
        .Add xtpControlButton, ID_EDIT_FORMATBRUSH, "��ʽˢ"

        Set objControl = .Add(xtpControlButton, ID_EDIT_UNDO, "����")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_EDIT_REDO, "����"

        Set objControl = .Add(xtpControlButton, ID_INSERT_DATETIME, "����������ʱ��")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_DATE, "��������"
        .Add xtpControlButton, ID_INSERT_TIME, "����ʱ��"
        .Add xtpControlButton, ID_INSERT_SPECIALCHAR, "�����������"

        Set objControl = .Add(xtpControlButton, ID_INSERT_PICTURE, "����ͼ��"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_ELEMENT, "����Ҫ��"
        .Add xtpControlButton, ID_INSERT_PACSPIC, "���뱨��ͼ"

        Set objControl = .Add(xtpControlButton, ID_FILE_IMPORT, "��ʷ�ļ�"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_INSERT_EPRDEMO, "���뷶��"
        .Add xtpControlButton, ID_INSERT_DOCADVISE, "����ҽ��"
        .Add xtpControlButton, ID_DIAGNOSIS, "�������"

        Set cbpPopup = .Add(xtpControlPopup, 0, "���"): cbpPopup.BeginGroup = True
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_REFCOMPEND, "ˢ�����"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_ADDCOMPEND, "�������"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_DELCOMPEND, "ɾ�����"): objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = cbpPopup.CommandBar.Controls.Add(xtpControlButton, ID_EDIT_MODCOMPEND, "�޸����"): objControl.STYLE = xtpButtonIconAndCaption

        Set objControl = .Add(xtpControlButton, ID_VIEW_STRUCTURE, "�ĵ��ṹͼ"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_VIEW_PHRASEDEMO, "ʾ���ʾ��б�"
        .Add xtpControlButton, ID_VIEW_SEGMENT, "ʾ��Ƭ���б�"
        .Add xtpControlButton, ID_VIEW_PACSPIC, "����ͼ�б�"
        .Add xtpControlButton, ID_VIEW_HISTORYWINDOW, "��ʷ�����б�"
        
        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "zlRichEMR ����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_COMMON_CANCEL, "�˳�(&Q)"): objControl.BeginGroup = True: objControl.STYLE = xtpButtonIconAndCaption
     
        Set objControl = .Add(xtpControlLabel, 99999901, "��������:")
        objControl.flags = xtpFlagRightAlign
        objControl.Visible = False
        Set cbrCustom = .Add(xtpControlCustom, 99999902, "��������")
        cbrCustom.Handle = Me.txtFeedBack.hwnd
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Visible = False
        
        Set objControl = .Add(xtpControlLabel, 99999903, "����˵��:")
        objControl.flags = xtpFlagRightAlign
        objControl.Visible = False
        Set cbrCustom = .Add(xtpControlCustom, 99999904, "����˵��")
        cbrCustom.Handle = Me.txtContent.hwnd
        cbrCustom.flags = xtpFlagRightAlign
        cbrCustom.Visible = False
    End With
    
    If Not gobjPlugIn Is Nothing Then '����˵���������
        Dim strFunc As String, lngFuncID As Long, strFuncName As String
        strFunc = gobjPlugIn.GetFuncNames(glngSys, 1070)
        '���Ӳ˵�����
        If strFunc <> "" Then
            With cbrThis.ActiveMenuBar.Controls
                Set cbrMenu = .Find(, ID_Main_HELP)
                If Not cbrMenu Is Nothing Then
                    i = cbrMenu.Index
                Else
                    i = -1
                End If
                Set cbrMenu = .Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ����", i, False)
                With cbrMenu.CommandBar.Controls
                    For i = 0 To UBound(Split(strFunc, ","))
                        lngFuncID = conMenu_Tool_PlugIn_Item + i + 1
                        strFuncName = Split(strFunc, ",")(i)
                        
                        If UCase(strFuncName) Like UCase("Auto:*") Then
                            strFuncName = Mid(strFuncName, 6)
                        End If
                        
                        Set objControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                        If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                        objControl.IconId = conMenu_Tool_PlugIn_Item
                        objControl.Parameter = strFuncName
                    Next
                End With
                
                If .Count > 1 Then .Item(2).BeginGroup = True
            End With
            '���ӹ���������
            Set cbrBar = cbrThis(2)
            Set objControl = cbrBar.FindControl(, ID_HELP_CONTENT)
            If Not objControl Is Nothing Then
                objControl.BeginGroup = True
                i = objControl.Index
            Else
                i = -1
            End If
            
            Set cbpPopup = cbrBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ����", i, False)
            cbpPopup.ID = conMenu_Tool_PlugIn
            cbpPopup.IconId = conMenu_Tool_PlugIn
            cbpPopup.BeginGroup = True
            With cbpPopup.CommandBar.Controls
                For i = 0 To UBound(Split(strFunc, ","))
                    lngFuncID = conMenu_Tool_PlugIn_Item + i + 1
                    strFuncName = Split(strFunc, ",")(i)
                    
                    If UCase(strFuncName) Like UCase("Auto:*") Then
                        strFuncName = Mid(strFuncName, 6)
                    End If
                    
                    Set objControl = .Add(xtpControlButton, lngFuncID, strFuncName)
                    If i <= 9 Then objControl.Caption = objControl.Caption & "(&" & IIf(i = 9, 0, i + 1) & ")"
                    objControl.IconId = conMenu_Tool_PlugIn_Item
                    objControl.Parameter = strFuncName
                Next
            End With
        End If
    End If

    Set cbrBar = cbrThis.Add("��ʽ", xtpBarTop): cbrBar.BarId = ID_BAR_FORMAT
    With cbrBar.Controls
        .Add xtpControlButton, ID_FORMAT_STYLEWINDOW, "��ʽ����"

        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_STYLE, "������ʽ")
        Dim rs As New ADODB.Recordset
        Set rs = zlDatabase.OpenSQLRecord("select ���� from ����������ʽ order by ���", "��ȡ��Ϣ", "")
        i = 0
        Do While Not rs.EOF
            i = i + 1
            Combo.AddItem rs("����")
            If rs("����") = "����" Then Combo.ListIndex = i
            rs.MoveNext
        Loop
        Combo.AddItem "����..."
        Combo.Width = 50
        Combo.DropDownWidth = 220
        Combo.DropDownListStyle = True

        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTNAME, "��������")
        Combo.BeginGroup = True
        For i = 0 To gfrmPublic.cmbFont.ListCount - 1
            Combo.AddItem gfrmPublic.cmbFont.List(i), i + 1
            If gfrmPublic.cmbFont.List(i) = "����" Then Combo.ListIndex = i + 1
        Next
        Combo.Width = 90
        Combo.DropDownWidth = 250
        Combo.DropDownListStyle = True

        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTSIZE, "����ߴ�")
        '�ֺ��б�
        Combo.AddItem "����", 1
        Combo.AddItem "С��", 2
        Combo.AddItem "һ��", 3
        Combo.AddItem "Сһ", 4
        Combo.AddItem "����", 5
        Combo.AddItem "С��", 6
        Combo.AddItem "����", 7
        Combo.AddItem "С��", 8
        Combo.AddItem "�ĺ�", 9
        Combo.AddItem "С��", 10
        Combo.AddItem "���", 11
        Combo.AddItem "С��", 12
        Combo.AddItem "����", 13
        Combo.AddItem "С��", 14
        Combo.AddItem "�ߺ�", 15
        Combo.AddItem "�˺�", 16
        Combo.AddItem 5, 17
        Combo.AddItem 5.5, 18
        Combo.AddItem 6.5, 19
        Combo.AddItem 7.5, 20
        Combo.AddItem 8, 21
        Combo.AddItem 9, 22
        Combo.AddItem 10, 23
        Combo.AddItem 10.5, 24
        Combo.AddItem 11, 25
        Combo.AddItem 12, 26
        Combo.AddItem 14, 27
        Combo.AddItem 16, 28
        Combo.AddItem 18, 29
        Combo.AddItem 20, 30
        Combo.AddItem 22, 31
        Combo.AddItem 24, 32
        Combo.AddItem 26, 33
        Combo.AddItem 28, 34
        Combo.AddItem 36, 35
        Combo.AddItem 48, 36
        Combo.AddItem 72, 37

        Combo.ListIndex = 12
        Combo.Width = 50
        Combo.DropDownWidth = 80
        Combo.DropDownListStyle = True

        Set objControl = .Add(xtpControlButton, ID_FORMAT_BOLD, "����")
        objControl.BeginGroup = True

        Set cbpPopupSub = .Add(xtpControlSplitButtonPopup, ID_FORMAT_UNDERLINE, "�»���"): cbpPopupSub.ID = ID_FORMAT_UNDERLINE
        cbpPopupSub.CommandBar.SetPopupToolBar True
        cbpPopupSub.CommandBar.SetIconSize 60, 8
        cbpPopupSub.CommandBar.Width = 60
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THIN, "ϸ��"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_THICK, "����"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_WAVE, "������"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DOT, "����"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASH, "����"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT, "�㻮��"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_UNDERLINE_DASHDOT2, "˫�㻮��"

        .Add xtpControlButton, ID_FORMAT_SUPER, "�ϱ�"
        .Add xtpControlButton, ID_FORMAT_SUB, "�±�"
        .Add xtpControlButton, ID_FORMAT_PROTECT, "����"

        Set objControl = .Add(xtpControlButton, ID_FORMAT_ALIGNLEFT, "�����"): objControl.BeginGroup = True
        .Add xtpControlButton, ID_FORMAT_ALIGNCENTER, "����"
        .Add xtpControlButton, ID_FORMAT_ALIGNRIGHT, "�Ҷ���"

        Set cbpPopupSub = .Add(xtpControlSplitButtonPopup, ID_FORMAT_LINESPACE, "�о�")
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE1, "1.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE2, "1.3"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE3, "1.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE4, "2.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE5, "2.5"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE6, "3.0"
        cbpPopupSub.CommandBar.Controls.Add xtpControlButton, ID_FORMAT_LINESPACE7, "����..."

        Set objControl = .Add(xtpControlButton, ID_FORMAT_INDENTDECREASE, "����������"): objControl.BeginGroup = True
        objControl.Visible = False
        Set objControl = .Add(xtpControlButton, ID_FORMAT_INDENTINCREASE, "����������")
        objControl.Visible = False

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_HIGHLIGHT, "ͻ����ʾ"): cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorHighlight.hwnd
    End With
    cbpPopupSub.CommandBar.FindControl(, ID_FORMAT_LINESPACE1).Checked = True

    Set cbrBar = cbrThis.Add("ǩ��", xtpBarTop): cbrBar.BarId = ID_BAR_SIGN
    cbrBar.EnableDocking xtpFlagHideWrap
    cbrBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    With cbrBar.Controls
        Set objControl = .Add(xtpControlButton, ID_REVISION_PREV, "ǰһ���޶�")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_REVISION_NEXT, "��һ���޶�"
        Set objControl = .Add(xtpControlButton, ID_REVISION_RESET, "����޶�")
        objControl.STYLE = xtpButtonIconAndCaption

        
        Set objControl = .Add(xtpControlButton, ID_PATISIGN, "����ǩ��")
        objControl.STYLE = xtpButtonIconAndCaption
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_SIGN, "ǩ��")
        objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_UNTREAD, "����")
        objControl.STYLE = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_SIGN_QUIT, "ǩ���˳�")
        objControl.STYLE = xtpButtonIconAndCaption
    End With

    Set cbrBar = cbrThis.Add("���", xtpBarTop): cbrBar.BarId = ID_BAR_TABLE
    cbrBar.EnableDocking xtpFlagHideWrap
    With cbrBar.Controls
        Set objControl = .Add(xtpControlButton, ID_INSERT_TABLE, "������"): objControl.BeginGroup = True

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FILLCOLOR, "�����ɫ")
        cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorFillColor.hwnd

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_FORMAT_FORECOLOR, "������ɫ")
        cbpPopup.BeginGroup = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorForeColor.hwnd

        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_TABLE_CELLALIGNMENT, "��Ԫ����뷽ʽ")
        cbpPopup.CommandBar.SetTearOffPopup "��Ԫ����뷽ʽ", ID_TABLE_CELLALIGNMENT, 100
        cbpPopup.CommandBar.SetPopupToolBar True
        cbpPopup.CommandBar.Width = 70
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT1, "���������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT2, "���Ͼ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT3, "�����Ҷ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT4, "�в������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT5, "�в�����"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT6, "�в��Ҷ���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT7, "���������"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT8, "���¾���"
        cbpPopup.CommandBar.Controls.Add xtpControlButton, ID_TABLE_CELLALIGNMENT9, "�����Ҷ���"

        Set objControl = .Add(xtpControlButton, ID_TABLE_CURRENCY, "������ʽ"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_TABLE_PERCENT, "�ٷֱ���ʽ")
        Set objControl = .Add(xtpControlButton, ID_TABLE_KILOBIT, "ǧλ�ָ���ʽ")

        Set objControl = .Add(xtpControlButton, ID_TABLE_MERGE, "�ϲ���Ԫ��"): objControl.BeginGroup = True
        objControl.BeginGroup = True
        If Not Document Is Nothing Then
            If Document.EditType = cprET_�����ļ����� Then
                Set objControl = .Add(xtpControlButton, ID_TABLE_CELLPROTECTED, "������Ԫ��")
            End If
        End If
    End With
    
    '������λ�õ���
    DockingRightOf CommBar(ID_BAR_FORMAT), CommBar(ID_BAR_SIGN)

    '�����ǲ����õ����ز˵�
    cbrThis.Options.AddHiddenCommand ID_FILE_IMPORT
    cbrThis.Options.AddHiddenCommand ID_FILE_SAVEAS
    cbrThis.Options.AddHiddenCommand ID_FILE_SAVE_QUIT
    cbrThis.Options.AddHiddenCommand ID_FILE_PRINTPREVIEW
    cbrThis.Options.AddHiddenCommand ID_EDIT_FINDNEXT
    cbrThis.Options.AddHiddenCommand ID_EDIT_REPLACE
'    cbrThis.Options.AddHiddenCommand ID_VIEW_GRID
    cbrThis.Options.AddHiddenCommand ID_INSERT_DATETIME
    cbrThis.Options.AddHiddenCommand ID_FORMAT_PROTECT
    cbrThis.Options.AddHiddenCommand ID_TABLE_SHOWGRID

    '�ȼ���
    cbrThis.KeyBindings.Add FCONTROL, Asc("S"), ID_FILE_SAVE
    cbrThis.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    cbrThis.KeyBindings.Add FCONTROL, Asc("Z"), ID_EDIT_UNDO
    cbrThis.KeyBindings.Add FCONTROL, Asc("Y"), ID_EDIT_REDO
    cbrThis.KeyBindings.Add FCONTROL, Asc("X"), ID_EDIT_CUT
    cbrThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    cbrThis.KeyBindings.Add FCONTROL, Asc("V"), ID_EDIT_PASTE
    cbrThis.KeyBindings.Add FCONTROL, Asc("A"), ID_EDIT_SELECTALL
    cbrThis.KeyBindings.Add FCONTROL, Asc("F"), ID_EDIT_FIND
    cbrThis.KeyBindings.Add FCONTROL, Asc("H"), ID_EDIT_REPLACE
    cbrThis.KeyBindings.Add FCONTROL, Asc("D"), ID_VIEW_STRUCTURE
    cbrThis.KeyBindings.Add FCONTROL, Asc("M"), ID_VIEW_PHRASEDEMO
    cbrThis.KeyBindings.Add FCONTROL, Asc("G"), ID_VIEW_SEGMENT
    cbrThis.KeyBindings.Add FCONTROL, Asc("E"), ID_INSERT_ELEMENT
    cbrThis.KeyBindings.Add FCONTROL, Asc("B"), ID_FORMAT_BOLD
'    cbrThis.KeyBindings.Add FCONTROL, Asc("I"), ID_FORMAT_ITALIC
    cbrThis.KeyBindings.Add FCONTROL, Asc("Q"), ID_FILE_EXIT
    cbrThis.KeyBindings.Add FCONTROL, Asc("J"), ID_INSERT_AUTORECOGNISE     '����ʶ��
    cbrThis.KeyBindings.Add FCONTROL, Asc("T"), ID_VIEW_PENWINDOW           '��д���봰��
    cbrThis.KeyBindings.Add FCONTROL, Asc("O"), ID_SIGN_QUIT                'ǩ�����˳��༭��

    cbrThis.KeyBindings.Add 0, VK_F1, ID_HELP_CONTENT
    cbrThis.KeyBindings.Add 0, VK_F3, ID_EDIT_FINDNEXT
    cbrThis.KeyBindings.Add 0, VK_F5, ID_INSERT_PICTURE
    cbrThis.KeyBindings.Add 0, VK_F6, ID_TABLE_INSERTTABLE
    cbrThis.KeyBindings.Add 0, VK_F7, ID_INSERT_ELEMENT
    cbrThis.KeyBindings.Add 0, VK_F8, ID_EDIT_ADDCOMPEND
    cbrThis.KeyBindings.Add 0, VK_F9, ID_INSERT_EPRDEMO
    cbrThis.KeyBindings.Add 0, VK_F10, ID_DIAGNOSIS                         '���
    cbrThis.KeyBindings.Add 0, VK_F11, ID_SIGN                              'ǩ��
    cbrThis.KeyBindings.Add 0, VK_F12, ID_INSERT_AUTORECOGNISE              '����ʶ��
    cbrThis.KeyBindings.Add 0, VK_DELETE, ID_EDIT_DELETE
'    cbrThis.KeyBindings.Add 0, VK_BACK, ID_EDIT_BACKSPACE

    '���α༭����Ĭ�Ͽ�ݼ���������
    cbrThis.KeyBindings.Add FCONTROL, Asc("R"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("L"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("L"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("="), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("1"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("2"), -1
    cbrThis.KeyBindings.Add FCONTROL, Asc("5"), -1
    cbrThis.KeyBindings.Add FCONTROL + FSHIFT, Asc("A"), -1
    cbrThis.KeyBindings.Add FCONTROL + FSHIFT, Asc("L"), -1

    '## ��ͣ����������

    Set mfrmSentenceDetailed = New frmSentenceDetailed
    Set mfrmSegments = New frmSegmentList
    Set mfrmCompends = New frmCompends: mfrmCompends.SetParent Me
    Set mfrmModElement = New frmElementEdit
    Set mfrmInsElement = New frmInsElement
    Set mfrmDicSelect = New frmDicSelect
    Set mfrmStyleMan = New frmStyleMan
    Set cPicEditor = New cPictureEditor
    Set mfrmPacsPic = New frmPACSImg
    Set mfrmHistoryReport = New frmDockReportHistory
    Set mfrmDocksymbol = New frmDockSymbol
    DkpThis.SetCommandBars Me.cbrThis
    DkpThis.Options.ThemedFloatingFrames = True
    DkpThis.TabPaintManager.Position = xtpTabPositionTop

    Dim PaneCompend As XtremeDockingPane.Pane           '�ĵ��ṹͼ
    Dim PaneSentence As XtremeDockingPane.Pane          '�ʾ�ʾ��
    Dim PaneSegment As XtremeDockingPane.Pane           'ʾ������
    Dim PaneStyleMan As XtremeDockingPane.Pane          '������ʽά��
    Dim PaneMultiDocView As XtremeDockingPane.Pane      '���ĵ�Ԥ��
    Dim PanePacsPic As XtremeDockingPane.Pane           'PacsͼƬ��
    Dim PaneSharePage As XtremeDockingPane.Pane           '
    Dim PaneHistoryReport As XtremeDockingPane.Pane     '��������ʷ����
    Dim PaneDockSymbol As XtremeDockingPane.Pane        '�������
    
    '����ҳ�没���鿴
    Set PaneSharePage = DkpThis.CreatePane(ID_VIEW_HISTORYWINDOW, 200, 140, DockTopOf, Nothing)
    PaneSharePage.Title = "����ҳ��"
    PaneSharePage.Options = PaneNoFloatable Or PaneNoHideable
    PaneSharePage.Close

    '���ĵ�Ԥ��
    Set PaneMultiDocView = DkpThis.CreatePane(ID_VIEW_MULTIDOCVIEW, 200, 140, DockLeftOf, Nothing)
    PaneMultiDocView.Title = "�����б�"
    PaneMultiDocView.Options = PaneNoCloseable
    
    '�ĵ��ṹͼ
    Set PaneCompend = DkpThis.CreatePane(ID_VIEW_STRUCTURE, 200, 140, DockLeftOf, Nothing)
    PaneCompend.Title = "�ĵ��ṹ"
    PaneCompend.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneCompend.Hide
    DkpThis.AttachPane PaneCompend, PaneMultiDocView
    PaneMultiDocView.Close

    'ʾ���ʾ��б�
    Set PaneSentence = DkpThis.CreatePane(ID_VIEW_PHRASEDEMO, 200, 140, DockBottomOf, PaneCompend)
    PaneSentence.Title = "�ʾ�ʾ��"
    PaneSentence.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneSentence.Hide
    DkpThis.AttachPane PaneSentence, PaneCompend

    'ʾ��Ƭ���б�
    Set PaneSegment = DkpThis.CreatePane(ID_VIEW_SEGMENT, 200, 140, DockBottomOf, PaneSentence)
    PaneSegment.Title = "ʾ��Ƭ��"
    PaneSegment.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneSegment.Hide
    DkpThis.AttachPane PaneSegment, PaneSentence


    '����ͼƬ�б�
    Set PanePacsPic = DkpThis.CreatePane(ID_VIEW_PACSPIC, 200, 140, DockBottomOf, PaneSentence)
    PanePacsPic.Title = "����ͼƬ"
    PanePacsPic.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PanePacsPic.Hide
    DkpThis.AttachPane PanePacsPic, PaneSentence
    
    '��������ʷ����
    Set PaneHistoryReport = DkpThis.CreatePane(ID_VIEW_HISTORYREPORT, 200, 140, DockTopOf, Nothing)
    PaneHistoryReport.Title = "��ʷ���"
    PaneHistoryReport.Options = PaneNoCloseable
    If Screen.Width / Screen.TwipsPerPixelX <= 800 Then PaneHistoryReport.Hide
    DkpThis.AttachPane PaneHistoryReport, PaneSentence

    '������ʽά��
    Set PaneStyleMan = DkpThis.CreatePane(ID_FORMAT_STYLEWINDOW, 230, 140, DockRightOf, Nothing)
    PaneStyleMan.Title = "������ʽ"
    PaneHistoryReport.Options = PaneNoCloseable
    
    Set PaneDockSymbol = DkpThis.CreatePane(ID_VIEW_Assistant, 230, 140, DockRightOf, Nothing)
    PaneDockSymbol.Title = "��������"
    PaneDockSymbol.Options = PaneNoCloseable
    PaneDockSymbol.MaxTrackSize.Width = 280: PaneDockSymbol.MinTrackSize.Width = 230
    If Screen.Width / Screen.TwipsPerPixelX <= 1024 Then PaneDockSymbol.Hide
    DkpThis.AttachPane PaneStyleMan, PaneDockSymbol
    PaneStyleMan.Close
    
    '## ������ʼ������
    Editor1.Modified = False

    SetParent picPenInput.hwnd, Editor1.hwnd

    ColorForeColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ForeColor", vbBlack)
    ColorHighlight.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "HighlightColor", vbYellow)
    ColorFillColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "CellFillColor", vbWhite)
    Editor1.ShowRuler = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ShowRuler", False)
    SetColorIcon "FILLCOLOR", ID_DRAW_FILLCOLOR, ColorFillColor.COLOR
    SetColorIcon "FORECOLOR", ID_FORMAT_FORECOLOR, ColorForeColor.COLOR
    SetColorIcon "HIGHLIGHT", ID_FORMAT_HIGHLIGHT, IIf(ColorHighlight.COLOR = tomAutoColor, vbWhite, ColorHighlight.COLOR)

    Call RestoreWinState(Me, App.ProductName)   '�ڴ������ж�����ٻָ����岼��

    If zlDatabase.GetPara("ʹ�ø��Ի����", glngSys, , 0) = 0 Then
        Me.WindowState = vbMaximized
    End If

    CommBar(ID_BAR_TABLE).Visible = False
    CommBar(ID_BAR_NORMAL).FindControl(, ID_COMMON_CANCEL).STYLE = xtpButtonIconAndCaption     '�ָ�ͼ�꣫�ı�����
    CommBar(ID_BAR_SIGN).FindControl(, ID_REVISION_RESET).STYLE = xtpButtonIconAndCaption
    CommBar(ID_BAR_SIGN).FindControl(, ID_SIGN).STYLE = xtpButtonIconAndCaption
    CommBar(ID_BAR_SIGN).FindControl(, ID_UNTREAD).STYLE = xtpButtonIconAndCaption
    CommBar(ID_BAR_SIGN).FindControl(, ID_SIGN_QUIT).STYLE = xtpButtonIconAndCaption
    
    If imgX_S.Top < 0 Then
        imgX_S.Top = 5460
    End If
    
    '���������������ʾ˳�򣬱�֤������дʾ�ʾ������ǰ��ʾ
    DkpThis.ShowPane ID_VIEW_PHRASEDEMO
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneCompendHided", False) Then PaneCompend.Close
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSentenceHided", False) Then PaneSentence.Close
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSegmentHided", False) Then PaneSegment.Close
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PanePacsPicHided", False) Then PanePacsPic.Close
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneHistoryReportHided", False) Then PaneHistoryReport.Close
    If GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneDockSymbol", False) Then PaneDockSymbol.Hide

    '## ���ѡ����
    Set cDropDown = New cDropDownToolWindow
    cDropDown.Create picDropDown

    tmrThis.Enabled = True
    DT1 = Now
    DT1_EPR = Now
End Sub

'################################################################################################################
'## ���ܣ�  ȷ���Ƿ񱣴��޸ĺ���ļ�
'################################################################################################################
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If mblnPrecess Then
        Cancel = 1: Exit Sub
    End If
    If Editor1.Modified Then
        Dim r As Long
        r = MsgBox("�Ƿ񱣴�� """ & Me.Document.EPRFileInfo.���� & """ �ĸ��ģ�", vbYesNoCancel + vbExclamation, gstrSysName)
        If r = vbYes Then
            '�����ļ�
            If Me.Document.Ŀ��汾 > 16 Then
                MsgBox "Ŀǰϵͳ֧�ֵ����汾��Ϊ16������ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
                Cancel = 1
                Exit Sub
            End If
            If SaveEMRDoc = False Then Cancel = 1
        ElseIf r = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ����Ƿ��޸ģ�����ʾ�Ƿ񱣴�
'################################################################################################################
Public Function CheckModified(Optional blnCannotCancel As Boolean = False) As Boolean
    CheckModified = True
    If Editor1.Modified And Me.Document.EPRFileInfo.ID <> 0 Then
        Dim r As Long
        If blnCannotCancel Then
            r = MsgBox("�Ƿ񱣴�� """ & Me.Document.EPRFileInfo.���� & """ �ĸ��ģ�", vbYesNo + vbExclamation, gstrSysName)
        Else
            r = MsgBox("�Ƿ񱣴�� """ & Me.Document.EPRFileInfo.���� & """ �ĸ��ģ�", vbYesNoCancel + vbExclamation, gstrSysName)
        End If
        If r = vbYes Then
            '�����ļ�
            If Me.Document.Ŀ��汾 > 16 Then
                MsgBox "Ŀǰϵͳ֧�ֵ����汾��Ϊ16������ʧ�ܣ�", vbOKOnly + vbInformation, gstrSysName
                CheckModified = False
                Exit Function
            End If
            If SaveEMRDoc = False Then CheckModified = False
        ElseIf r = vbCancel Then
            CheckModified = False
        End If
    End If
End Function

'################################################################################################################
'## ���ܣ�  ���������Ϣ���رմ���
'################################################################################################################
Private Sub Form_Unload(Cancel As Integer)
    '����λ����Ϣ
    If mblnPrecess Then
        Cancel = 1
    End If
    Dim i As Long
    If Me.WindowState <> vbMinimized Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ForeColor", ColorForeColor.COLOR
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "HighlightColor", ColorHighlight.COLOR
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "CellFillColor", ColorFillColor.COLOR
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "ShowRuler", Me.Editor1.ShowRuler
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneCompendHided", DkpThis.FindPane(ID_VIEW_STRUCTURE).Closed
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSentenceHided", DkpThis.FindPane(ID_VIEW_PHRASEDEMO).Closed
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneSegmentHided", DkpThis.FindPane(ID_VIEW_SEGMENT).Closed
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PanePacsPicHided", DkpThis.FindPane(ID_VIEW_PACSPIC).Closed
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneHistoryReportHided", DkpThis.FindPane(ID_VIEW_HISTORYREPORT).Closed
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & "frmMain", "PaneDockSymbol", DkpThis.FindPane(ID_VIEW_Assistant).Hidden
    

    Call SaveWinState(Me, App.ProductName)
    
    On Error Resume Next
    If Not Me.Document Is Nothing Then
        Call Me.Document.AfterClosed(Me.Document.EPRPatiRecInfo.ҽ��id) '�����ر��¼�
    End If
       
        Dim objTmp As Object
    For Each objTmp In Me.Controls
        Me.Controls.Remove objTmp
    Next

    '�����ʱ�ļ�
    For i = 1 To UBound(UndoList)
        If gobjFSO.FileExists(UndoList(i).Filename) Then gobjFSO.DeleteFile UndoList(i).Filename
    Next

    Erase mEleLimit
    Erase UndoList
    Set mParaFmt = Nothing
    Set mFontFmt = Nothing
    
    If Not mfrmCompends Is Nothing Then Unload mfrmCompends
    If Not mfrmSentenceDetailed Is Nothing Then Unload mfrmSentenceDetailed
    If Not mfrmSegments Is Nothing Then Unload mfrmSegments
    If Not mfrmModElement Is Nothing Then Unload mfrmModElement
    If Not mfrmInsElement Is Nothing Then Unload mfrmInsElement
    If Not mfrmDicSelect Is Nothing Then Unload mfrmDicSelect
    If Not mfrmStyleMan Is Nothing Then Unload mfrmStyleMan
    If Not mfrmMultiDocView Is Nothing Then Unload mfrmMultiDocView
    If Not mfrmPacsPic Is Nothing Then Unload mfrmPacsPic
    If Not mfrmHistoryReport Is Nothing Then Unload mfrmHistoryReport
    If Not mfrmMainError Is Nothing Then Unload mfrmMainError
    If Not mfrmPreview Is Nothing Then Unload mfrmPreview
    If Not mfrmDocksymbol Is Nothing Then Unload mfrmDocksymbol
'    If Not gfrmPublic Is Nothing Then Unload gfrmPublic
     
    Set mfrmCompends = Nothing
    Set mfrmSentenceDetailed = Nothing
    Set mfrmSegments = Nothing
    Set mfrmModElement = Nothing
    Set mfrmInsElement = Nothing
    Set mfrmDicSelect = Nothing
    Set mfrmStyleMan = Nothing
    Set cPicEditor = Nothing
    Set mfrmMultiDocView = Nothing
    Set mfrmPacsPic = Nothing
    Set mfrmHistoryReport = Nothing
    Set mfrmMainError = Nothing
    Set mfrmPreview = Nothing
    Set mfrmDocksymbol = Nothing
    Set mobjReport = Nothing
    Set cDropDown = Nothing

    Set Me.Document = Nothing

'    ImageManager.Icons.RemoveAll
    imgColor.ListImages.Clear
    ImageList_Destroy imgColor.hImageList
    Set imgX_S.Picture = Nothing
    Set picDropDown.Picture = Nothing
    Set picHistoryInfo.Picture = Nothing
    Set picPane.Picture = Nothing
    Set picPatiInfo.Picture = Nothing
    Set picPenInput.Picture = Nothing
    Set picTMP.Picture = Nothing
'    '�ֶ��ͷ��ڴ�
'    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
End Sub

Public Sub ShowTablePicker(ByVal X As Long, ByVal y As Long)
   '��ʾ�������ѡ����
   DefaultSize picDropDown
   cDropDown.Show X, y
   picDropDown.Visible = True
   DrawTableChooser cDropDown, 0, 0, 0
End Sub

Private Sub HideDropDown()
   cDropDown.Hide
End Sub

Private Sub picDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' Mouse down handling:
    If (cDropDown.IsShown) Then
        ' Drop down window is visible
        If Not (cDropDown.InRect(X, y)) Then
            ' Mouse down outside drop-down area:
            HideDropDown
        Else
            ' Mouse down inside the drop down:
            DrawTableChooser cDropDown, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, Button
        End If
    End If
End Sub

Private Sub picDropDown_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    ' Mouse move.  Note that because all mouse messages are captured,
    ' x and y may be outside the limits of picDropdown.  This is
    ' handled in the draw routine.
    DrawTableChooser cDropDown, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, Button
End Sub

Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim xCellHit As Long, yCellHit As Long, bIn As Boolean

    ' Mouse up.  Determine whether mouse up over a cell:
    DrawTableChooser cDropDown, X \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, Button, bIn, xCellHit, yCellHit
    ' Hide the drop down:
    HideDropDown
    ' If an item selected, say what it was:
    If (bIn) Then
        Dim strTmp As String, lLen As Long, lngKey As Long, i As Long, j As Long, lKey2 As Long
        Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean

        With Editor1
            bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.SelStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bBeteenKeys = False Then
                lKey = Me.Document.Tables.Add
                tblThis.Redraw = False
                tblThis.SingleClickEdit = False
                tblThis.HighlightMode = HMFilledRectAlpha
                tblThis.Width = Me.Editor1.PaperWidth - Me.Editor1.Selection.Para.LeftIndent - Me.Editor1.MarginLeft - Me.Editor1.MarginRight - 800
                tblThis.Init yCellHit, xCellHit
                tblThis.CellMargin = 10
                For i = 1 To yCellHit
                    For j = 1 To xCellHit
                        Me.Document.Tables("K" & lKey).Cells.Add , i, j
'                        tblThis.Cell(i, j).Width = 100
                    Next j
                Next i
                tblThis.Tag = lKey
                tblThis.ShowToolTipText = True
                tblThis.MinRowHeight = 300
                tblThis.Redraw = True
                tblThis.Refresh
                SaveUIToTable Me.Document.Tables("K" & lKey), True
            End If
        End With
    End If
End Sub

Private Sub mfrmCompends_NodeSelected(lngCompendID As Long)
    'ͬ��ʾ���ʾ����ʾ
    Call RefSentenceList
End Sub

Private Sub mfrmDicSelect_pOK(strR As String)
    '�ֵ���Ŀ����ֵ
    Dim strTmp As String, T As Variant
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    If tblThis.Visible Then
        '����б༭Ҫ��
        If Val(tblThis.Tag) > 0 And Val(mfrmDicSelect.Tag) > 0 Then
            T = Split(strR, ";")
            If UBound(T) > 0 And Me.Document.Tables("K" & tblThis.Tag).Elements("K" & mfrmDicSelect.Tag).�滻�� = 2 Then
                Me.Document.Tables("K" & tblThis.Tag).Elements("K" & mfrmDicSelect.Tag).�����ı� = T(1)
                '���浽��Ԫ����
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = T(1)
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
                tblThis.Modified = True
                If mintStyle <> 0 Then 'ģ̬״̬����ʹ��
                    tblThis.SetFocus
                End If
            End If
        End If
        Exit Sub
    End If

    bFinded = FindKey(Editor1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        T = Split(strR, ";")
        If UBound(T) > 0 And Me.Document.Elements("K" & glngCurEleKey).�滻�� = 2 Then
            strTmp = T(1)
            Me.Document.Elements("K" & glngCurEleKey).�����ı� = strTmp

            Me.Document.Elements("K" & glngCurEleKey).Refresh Me.Editor1

            Call UpdateSameELement(glngCurEleKey)
            Call FindKey(Editor1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded) '��ΪUpdateSameELement�п��ܸı��ı����ȣ��Ӷ��ı�ѡ��λ��
            
            If Trim(Me.Document.Elements("K" & glngCurEleKey).�����ı�) <> "" Then
                '�Զ���λ����һ��Ҫ��λ��
                bFinded = FindNextKey(Editor1, Editor1.Selection.StartPos + 1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded Then
                    Editor1.Range(lKSE, lKES).Selected
                End If
            Else
                Editor1.Range(lKSE, lKSE + Me.Document.Elements("K" & glngCurEleKey).GetValidTextLength).Selected
            End If
            Editor1.ForceEdit = False
            Editor1.UnFreeze
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ȡ������Ҫ�صĲ���
'################################################################################################################
Private Sub mfrmInsElement_pCancel()
    mfrmInsElement.Hide
    mfrmInsElement.Tag = ""
End Sub

'################################################################################################################
'## ���ܣ�  ��������Ҫ�صĲ���
'################################################################################################################
Private Sub mfrmInsElement_pOK(Ele As cEPRElement)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lngKey As Long, bInKeys As Boolean, bNeeded As Boolean, lKey2 As Long
    If mfrmInsElement.Tag <> "" Then
        '�޸�ģʽ
        If mbEditInTable Then
            '����е�Ҫ��
            If Val(tblThis.Tag) > 0 Then
                Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.Remove "K" & mfrmInsElement.Tag
                lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.AddExistNode(Ele, True)
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                    Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).Ҫ������, _
                        Me.Document.EPRPatiRecInfo.����ID, _
                        Me.Document.EPRPatiRecInfo.��ҳID, _
                        Me.Document.EPRPatiRecInfo.������Դ, _
                        Me.Document.EPRPatiRecInfo.ҽ��id, _
                        Me.Document.EPRPatiRecInfo.Ӥ��)
                End If
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                    If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�Զ�ת�ı� Then Me.Document.EleToString Me.Editor1, Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey)  '�Զ�ת��Ϊ���ı�
                End If
                '���浽��Ԫ����
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�����ı�
                tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
                tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).Ҫ������
                tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                tblThis.Modified = True
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        Else
            '�ı��е�Ҫ��
            Me.Document.Elements.Remove "K" & mfrmInsElement.Tag
            lngKey = Me.Document.Elements.AddExistNode(Ele, True)
            If Me.Document.Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                Me.Document.Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).Ҫ������, _
                    Me.Document.EPRPatiRecInfo.����ID, _
                    Me.Document.EPRPatiRecInfo.��ҳID, _
                    Me.Document.EPRPatiRecInfo.������Դ, _
                    Me.Document.EPRPatiRecInfo.ҽ��id, _
                    Me.Document.EPRPatiRecInfo.Ӥ��)
            End If
            Me.Document.Elements("K" & lngKey).Refresh Me.Editor1
            If Me.Document.Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                If Me.Document.Elements("K" & lngKey).�Զ�ת�ı� Then Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lngKey)  '�Զ�ת��Ϊ���ı�
            End If
            bInKeys = FindKey(Me.Editor1, "E", lngKey, lSS, lSE, lES, lEE, bNeeded)
            If bInKeys Then
                If Me.Document.Elements("K" & lngKey).������̬ = 0 Then
                    Editor1.Range(lSE, lES).Selected
                Else
                    Editor1.Range(lSE + 1, lSE + 1).Selected
                End If
            End If
        End If
    Else
        If mbEditInTable Then
            '����е�Ҫ��
            If Val(tblThis.Tag) > 0 Then
                lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements.AddExistNode(Ele)
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                    Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).Ҫ������, _
                        Me.Document.EPRPatiRecInfo.����ID, _
                        Me.Document.EPRPatiRecInfo.��ҳID, _
                        Me.Document.EPRPatiRecInfo.������Դ, _
                        Me.Document.EPRPatiRecInfo.ҽ��id, _
                        Me.Document.EPRPatiRecInfo.Ӥ��)
                End If
                Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).��ʼ�� = Me.Document.Ŀ��汾
                If Val(tblThis.Tag) <= 0 Then Exit Sub
                If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                    If Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�Զ�ת�ı� Then Me.Document.EleToString Me.Editor1, Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey)  '�Զ�ת��Ϊ���ı�
                End If
                '���浽��Ԫ����
                tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�����ı�
                tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
                tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).Ҫ������
                tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
                tblThis.Modified = True
                tblThis.Refresh False, True, tblThis.SelectedCellKey
                tblThis_Resize tblThis.Width, tblThis.Height
            End If
        Else
            lngKey = Me.Document.Elements.AddExistNode(Ele)
            If Me.Document.Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                Me.Document.Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).Ҫ������, _
                    Me.Document.EPRPatiRecInfo.����ID, _
                    Me.Document.EPRPatiRecInfo.��ҳID, _
                    Me.Document.EPRPatiRecInfo.������Դ, _
                    Me.Document.EPRPatiRecInfo.ҽ��id, _
                    Me.Document.EPRPatiRecInfo.Ӥ��)
            End If
            Me.Document.Elements("K" & lngKey).��ʼ�� = Me.Document.Ŀ��汾
            Me.Document.Elements("K" & lngKey).InsertIntoEditor Me.Editor1, , True
            If Me.Document.Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                If Me.Document.Elements("K" & lngKey).�Զ�ת�ı� Then Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lngKey)  '�Զ�ת��Ϊ���ı�
            End If
        End If
    End If
    mfrmInsElement.Tag = ""
End Sub

'################################################################################################################
'## ���ܣ�  ȡ������Ҫ�صı༭
'################################################################################################################
Private Sub mfrmModElement_pCancel()
    On Error Resume Next
    Unload mfrmModElement
'    mfrmModElement.Hide
    Err.Clear
End Sub

'################################################################################################################
'## ���ܣ�  ��������Ҫ�صı༭
'################################################################################################################
Private Sub mfrmModElement_pOK()
    '��������Ҫ�ر༭���
    Dim strTmp As String, lngKey As Long, Ele As cEPRElement, lS As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    If tblThis.Visible Then
        '����б༭Ҫ��
        If Val(tblThis.Tag) > 0 Then
            Me.Document.Tables("K" & tblThis.Tag).Elements.Remove "K" & tblThis.Cells("K" & tblThis.SelectedCellKey).Tag
            Set Ele = mfrmModElement.Element.Clone(True)
            lngKey = Me.Document.Tables("K" & tblThis.Tag).Elements.AddExistNode(Ele, True)
            If Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).�滻�� = 1 And mfrmModElement.Element.�����ı� = "" Then
                Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Tables("K" & tblThis.Tag).Elements("K" & lngKey).Ҫ������, Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID, Me.Document.EPRPatiRecInfo.������Դ, Me.Document.EPRPatiRecInfo.ҽ��id, Me.Document.EPRPatiRecInfo.Ӥ��)
            End If
            '���浽��Ԫ����
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).�����ı�
            tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
            tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = Me.Document.Tables("K" & Val(tblThis.Tag)).Elements("K" & lngKey).Ҫ������
            tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
            tblThis.Modified = True
            tblThis.SetFocus
        End If
        Exit Sub
    End If

    bFinded = FindKey(Editor1, "E", glngCurEleKey, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        lS = lKEE
        If mfrmModElement.Element.�滻�� = 1 And (Me.Document.EditType = cprET_�����ļ�����) Then
            '�Զ���λ����һ��Ҫ��λ��
            bFinded = FindNextKey(Editor1, Editor1.Selection.StartPos + 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then
                '������ELE_JUMP_LIMIT���ַ�֮��ĲŶ�λ��ȥ
                If lKSS - lS < ELE_JUMP_LIMIT Then Editor1.Range(lKSE, lKES).Selected
            End If
        Else
            Me.Document.Elements.Remove "K" & glngCurEleKey
            Set Ele = mfrmModElement.Element.Clone(True)
            lngKey = Me.Document.Elements.AddExistNode(Ele, True)
            If Me.Document.Elements("K" & lngKey).�滻�� = 1 And mfrmModElement.Element.�����ı� = "" Then
                Me.Document.Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).Ҫ������, _
                    Me.Document.EPRPatiRecInfo.����ID, _
                    Me.Document.EPRPatiRecInfo.��ҳID, _
                    Me.Document.EPRPatiRecInfo.������Դ, _
                    Me.Document.EPRPatiRecInfo.ҽ��id, _
                    Me.Document.EPRPatiRecInfo.Ӥ��)
            End If

            Me.Document.Elements("K" & lngKey).Refresh Me.Editor1
            bFinded = FindKey(Editor1, "E", lngKey, lKSS, lKSE, lKES, lKEE, bNeeded)
            If bFinded Then lS = lKEE   '����lS

            If Me.Document.Elements("K" & lngKey).������̬ = 0 Then
                '�Զ�����ĵ�������
                If InStr(Trim(Me.Document.Elements("K" & lngKey).�����ı�), "�Զ���") > 0 And Me.Document.Elements("K" & lngKey).��̬�� = 1 Then
                    strTmp = Trim(InputBox("��¼���Զ���Ҫ��ѡ��" & vbCrLf & "������볤��200������", "�������"))
                    If strTmp <> "" Then
                        Me.Document.Elements("K" & lngKey).�����ı� = Replace(Me.Document.Elements("K" & lngKey).�����ı�, "�Զ���", strTmp)
                    Else
                        Me.Document.Elements("K" & lngKey).�����ı� = Replace(Me.Document.Elements("K" & lngKey).�����ı�, "�Զ���", "")
                    End If
                    Me.Document.Elements("K" & lngKey).Refresh Me.Editor1
                End If
                
                Call CheckElementLimit(lngKey)
                bFinded = FindKey(Editor1, "E", lngKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                If bFinded Then lS = lKEE   '����lS,��ΪCheckElementLimit�п��ܸı��ı����ȣ��Ӷ��ı�ѡ��λ��
                
                If Trim(Me.Document.Elements("K" & lngKey).�����ı�) <> "" Then
                    '�Զ���λ����һ��Ҫ��λ��
                    bFinded = FindNextKey(Editor1, lKEE, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
                    If bFinded Then
                        '������ELE_JUMP_LIMIT���ַ�֮��ĲŶ�λ��ȥ
                        If lKSS - lS < ELE_JUMP_LIMIT Then Editor1.Range(lKSE, lKES).Selected
                    End If
                Else
                    Editor1.Range(lKSE, lKSE + Me.Document.Elements("K" & lngKey).GetValidTextLength).Selected
                End If
            End If
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ�  ��������A���õ�������B��ͬһ��
'##
'## ������  BarToDock   ������Ĺ�����
'##         BarOnLeft   ��λ����ߵĹ�����
'################################################################################################################
Private Sub DockingRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    cbrThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbrThis.DockToolBar BarToDock, 0, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

'################################################################################################################
'## ���ܣ�  ����ĵ�
'################################################################################################################
Private Sub ClearDoc()
    If Len(Trim(Editor1.Text)) > 0 Then
        Dim r As Long
        If Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_ȫ��ʾ���༭ Then
            r = MsgBox("���棺��յ�ǰ�ĵ����ĵ����ָ����ļ������ʼ״̬�������Ѿ�¼�����Ϣ����ʧ��" & vbCrLf & "ȷ��Ҫ�������������", vbYesNo + vbExclamation, gstrSysName)
        ElseIf Me.Document.EditType = cprET_�����ļ����� Then
            r = MsgBox("���棺��յ�ǰ�ĵ�����ʧ�����Ѿ�¼�����Ϣ��" & vbCrLf & "ȷ��Ҫ�������������", vbYesNo + vbExclamation, gstrSysName)
        Else
            Exit Sub
        End If
        If r = vbYes Then
            Call AddUndoPoint  '�ֶ�����
            Editor1.InProcessing = True
            Editor1.Freeze
            Editor1.Tag = "ClearDoc"
            Editor1.NewDoc
            Set Me.Document.Compends = New cEPRCompends
            Set Me.Document.Elements = New cEPRElements
            Set Me.Document.Tables = New cEPRTables
            Set Me.Document.Pictures = New cEPRPictures
            If Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_ȫ��ʾ���༭ Then
                Me.Document.ReadInitFileStructure Me.Editor1
            End If
            Editor1.UnFreeze
            Editor1.Tag = ""
            Me.RefCompends
            Call ClearNoUseUndoList
        Else
            Exit Sub
        End If
    End If
    Call SetStateInfo
    Editor1.Filename = ""
    Editor1.Modified = True
    Editor1.InProcessing = False
    If tblThis.Visible = False Then Editor1.SetFocus
End Sub

'################################################################################################################
'## ���ܣ�  ���õ�ǰ״̬��Ϣ����������״̬�������ã�
'################################################################################################################
Public Sub SetStateInfo()
    Select Case Document.EditType
    Case cprET_�����ļ�����
        stbThis.Panels(3).Text = "�ļ�����"
        Me.Caption = Document.EPRFileInfo.����
    Case cprET_ȫ��ʾ���༭
        stbThis.Panels(3).Text = "���ı༭"
        Me.Caption = Document.EPRFileInfo.����
    Case cprET_�������༭
        stbThis.Panels(3).Text = "�ļ��༭"
        Me.Caption = Document.EPRFileInfo.���� & " (��" & Document.Ŀ��汾 & "��) "
    Case cprET_���������
        stbThis.Panels(3).Text = "�ļ��޶�"
        Me.Caption = Document.EPRFileInfo.���� & " (��" & Document.Ŀ��汾 & "��)"
    End Select
    Select Case Document.EditMode
    Case cprEM_����
        stbThis.Panels(4).Text = "��������"
    Case cprEM_�޸�
        stbThis.Panels(4).Text = "���޸ġ�"
    End Select
    
    Me.Caption = gstrUserName & "��" & Me.Caption
End Sub

'################################################################################################################
'## ���ܣ�  ������ʷ�ļ���
'################################################################################################################
Private Sub ImportEPRDoc()
On Error GoTo errHand

    Dim f As New frmEPRSearchMan, lngR As Long
    lngR = f.ShowSearchFile(Me, Me.Document.EPRFileInfo.ID, Me.Document.EPRPatiRecInfo.����ID)
    If lngR > 0 Then
        Dim rsTemp As New ADODB.Recordset
        gstrSQL = "Select Zl_Fun_ImportEnable([1]) CopyEnable From Dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngR)
        If rsTemp!CopyEnable <> 1 Then
            MsgBox "ѡ���Ĳ����ļ�����������", vbInformation, gstrSysName
            Exit Sub
        End If

        '����ָ������ʷ�ļ�
        If Me.Document.ImportOldEPRFile(Me.Editor1, lngR) Then
            MsgBox "�ɹ�������ʷ�ļ���", vbOKOnly + vbInformation, gstrSysName
        End If
    End If
    Exit Sub
errHand:
    MsgBox Err.Description, vbInformation, gstrSysName
End Sub

'################################################################################################################
'## ���ܣ�  ��һ��RTF�ĵ���
'################################################################################################################
Private Sub OpenRTFDoc()
    If Editor1.ViewMode <> cprNormal Then Exit Sub
    If Editor1.Modified Then
        Dim r As Long
        r = MsgBox("�Ƿ񱣴�� """ & Editor1.Title & """ �ĸ���?            ", vbYesNoCancel + vbExclamation, gstrSysName)
        If r = vbYes Then
            '�����ļ�
            If Editor1.Filename = "" Then
                dlgThis.Filename = ""
                dlgThis.Filter = "*.rtf|*.rtf|*.*|*.*"
                dlgThis.ShowSave
                If dlgThis.Filename <> "" Then
                    Editor1.SaveDoc dlgThis.Filename
                Else
                    Exit Sub
                End If
            Else
                Editor1.SaveDoc
            End If
        ElseIf r = vbCancel Then
            Exit Sub
        End If
    End If
    dlgThis.Filename = ""
    dlgThis.Filter = "*.rtf|*.rtf|*.txt|*.txt|*.html|*.html|*.htm|*.htm|*.*|*.*"
    dlgThis.ShowOpen
    If dlgThis.Filename <> "" Then
        Editor1.OpenDoc dlgThis.Filename
        Me.Caption = Editor1.Title & " - zlRichEMR"
    End If
End Sub

'################################################################################################################
'## ���ܣ�  �����ĵ������ݿ�
'################################################################################################################
Public Function SaveEMRDoc(Optional ByVal blnNoAsk As Boolean = False) As Boolean
    Dim blnR As Boolean, eEditMode As EditModeEnum
    Dim strSQL As String, strTime As String
    eEditMode = Me.Document.EditMode
    If Me.Editor1.Modified Then
        If blnNoAsk Then
            blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "�޸�ҳ������") > 0)

            Editor1.Modified = Not blnR
        Else
            If tblThis.Visible Then
                Editor1_UIClose 0
                Editor1.CloseUIInterface
            End If
            blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "�޸�ҳ������") > 0)
            Editor1.Modified = Not blnR
        End If
    Else
        If (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
            If blnNoAsk Then
                blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "�޸�ҳ������") > 0)

                Me.Editor1.Modified = False
            Else
                If tblThis.Visible Then
                    Editor1_UIClose 0
                    Editor1.CloseUIInterface
                End If
                blnR = Me.Document.SaveEPRDoc(Editor1, InStr(1, gstrPrivsEpr, "�޸�ҳ������") > 0)
                Me.Editor1.Modified = False
            End If
        End If
    End If
    
    If mbln���޴��� And mblnFBContentChanged Then
        strTime = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "Zl_�����걨��¼_Update(" & Me.Document.EPRPatiRecInfo.ID & ",5,null,null,null,'" & gstrUserName & "'," & strTime & ",'" & Trim(txtContent.Text) & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mblnFBContentChanged = False
    End If
    SaveEMRDoc = blnR
    If blnR And mblnIsMultiMode And Not mfrmMultiDocView Is Nothing Then
        mfrmMultiDocView.InitData Me, Me.Document, Me.Document.EPRPatiRecInfo.ID
    End If
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  ����Ƿ����б�������Ķ��󶼲�Ϊ�գ�����пյĸ�����ʾ
'##
'## ���أ�  ����û�ȡ��������������򷵻�False�����򷵻�True����ʾǿ�Ʊ��棩
'################################################################################################################
Private Function CheckAllObjects(Optional CheckItemMust As Boolean) As Boolean
    If Me.Editor1.AuditMode Or Me.Document.EditType = cprET_�����ļ����� And Me.Document.EPRFileInfo.���� <> 3 Or Me.Document.EditType = cprET_ȫ��ʾ���༭ Then
        CheckAllObjects = True: Exit Function
    End If
    Dim i As Long, blnOK As Boolean, lngIndex As Long, strMsg As String
    
    If mfrmMainError Is Nothing Then Set mfrmMainError = New frmMainMsg
    
    CheckAllObjects = mfrmMainError.ShowNotice(Me, CheckItemMust)

End Function

'################################################################################################################
'## ���ܣ�  �����ĵ�Ϊ����
'################################################################################################################
Private Sub SaveDocAsEPRDemo()
    If Editor1.Modified Then MsgBox "������֮ǰ�����ȱ��汾�α༭��", vbExclamation, gstrSysName: Exit Sub
    Select Case Me.Document.EditType
    Case cprET_ȫ��ʾ���༭
        Call frmEPRModelSaveAs.ShowMe(1, Me.Document.EPRDemoInfo.ID)
        Err = 0: On Error Resume Next
        Call Me.Document.mfrmParent.RefreshList
    Case cprET_�������༭, cprET_���������
        If Me.Document.EPRPatiRecInfo.ID = 0 Then MsgBox "������֮ǰ�����ȱ��汾�α༭��", vbExclamation, gstrSysName: Exit Sub
        Call frmEPRModelSaveAs.ShowMe(2, Me.Document.EPRPatiRecInfo.ID)
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  �����ĵ�ΪƬ��
'################################################################################################################
Private Sub SaveDocAsSegment()
    Dim strCompends As String, objNode As MSComctlLib.Node
    
    If Editor1.Modified Then MsgBox "������֮ǰ�����ȱ��汾�α༭��", vbExclamation, gstrSysName: Exit Sub
    strCompends = ""
    For Each objNode In mfrmCompends.Tree.Nodes
        If objNode.Checked = True And Me.Document.Compends("K" & objNode.Tag).ID > 0 And Me.Document.Compends("K" & objNode.Tag).�������ID > 0 Then
            strCompends = strCompends & "," & Me.Document.Compends("K" & objNode.Tag).ID
        End If
    Next
    If strCompends = "" Then
        MsgBox "��Ҫ����ѡ����٣��Ա�ȷ������Щ��ٵ��������ΪƬ�Σ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    strCompends = Mid(strCompends, 2)
    
    Select Case Me.Document.EditType
    Case cprET_ȫ��ʾ���༭
        Call frmEPRModelSaveAs.ShowMe(1, Me.Document.EPRDemoInfo.ID, strCompends)
        Err = 0: On Error Resume Next
        Call Me.Document.mfrmParent.RefreshList
    Case cprET_�������༭, cprET_���������
        If Me.Document.EPRPatiRecInfo.ID = 0 Then MsgBox "������֮ǰ�����ȱ��汾�α༭��", vbExclamation, gstrSysName: Exit Sub
        Call frmEPRModelSaveAs.ShowMe(2, Me.Document.EPRPatiRecInfo.ID, strCompends)
        Call mfrmSegments.zlRefresh(Me)
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ��ȡҳüҳ���滻����ı���ֻ���ڵ�����Word��ӡ
'################################################################################################################
Private Function GetReplacedHeadFootStr(strIn As String) As String
    Dim strR As String, strUnitName As String
    strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
    strR = strIn
    strR = Replace(strR, "{��λ����}", strUnitName)
    strR = Replace(strR, "{����}", Editor1.Title)
    strR = Replace(strR, "{·��}", Left(Editor1.Filename, InStrRev(Editor1.Filename, "\")))
    strR = Replace(strR, "{�ļ���}", Mid(Editor1.Filename, InStrRev(Editor1.Filename, "\") + 1))
    strR = Replace(strR, "{����}", Format(Now(), "yyyy��mm��dd��"))
    strR = Replace(strR, "{ʱ��}", Format(Now(), "hh:MM:ss"))
    strR = Replace(strR, "{��ӡ����}", Format(Now(), "yyyy��mm��dd��"))
    strR = Replace(strR, "{��ӡʱ��}", Format(Now(), "hh:MM:ss"))
    GetReplacedHeadFootStr = strR
End Function

'################################################################################################################
'## ���ܣ�  ����ΪRTF��Ȼ��ͨ��Word��ӡ��ǰ�ļ�
'################################################################################################################
Private Sub PrintInWord()
    On Error Resume Next
    Dim strF As String, strPicFile As String, Fd As Object
    strF = GetSysTmpPath & "\PrintInWord_TMP" & App.ThreadID & ".rtf"
    If gobjFSO.FileExists(strF) Then gobjFSO.DeleteFile strF, True
    '�������������Ϊ���˶���
    Dim i As Long, j As Long
    Do
        i = InStr(i + 1, Editor1.Text, vbCrLf)
        If i > 0 Then
            If Editor1.TOM.TextDocument.Range(i - 2, i - 2).Para.Alignment = tomAlignLeft Then
                Editor1.TOM.TextDocument.Range(i - 2, i - 2).Para.Alignment = tomAlignJustify
            End If
        End If
    Loop Until i <= 0
    
    If SaveDocToFile(strF) Then
        If gobjFSO.FileExists(strF) Then
            Dim WordApp As Object   'Word.Application
            Dim WordDoc As Object   'Word.Document
            Set WordApp = CreateObject("Word.Application")
            Set WordDoc = WordApp.Documents.Open(strF)      '��RTF�ĵ�
            
            If WordApp Is Nothing Then
                MsgBox "�޷�����Word�����밲װ Microsoft Office Word ��Ʒ��", vbOKOnly + vbInformation, gstrSysName
            Else
                zlCommFun.ShowFlash "���Ժ�..."
                Screen.MousePointer = vbHourglass
                
                WordApp.Visible = False
                WordApp.ScreenUpdating = False
                
                'ҳ���С����
                WordDoc.PageSetup.LeftMargin = Me.ScaleX(Editor1.MarginLeft, vbTwips, vbPoints)
                WordDoc.PageSetup.RightMargin = Me.ScaleX(Editor1.MarginRight, vbTwips, vbPoints)
                WordDoc.PageSetup.TopMargin = Me.ScaleY(Editor1.MarginTop, vbTwips, vbPoints)
                WordDoc.PageSetup.BottomMargin = Me.ScaleY(Editor1.MarginBottom, vbTwips, vbPoints)
                WordDoc.PageSetup.PageWidth = Me.ScaleX(Editor1.PaperWidth, vbTwips, vbPoints)
                WordDoc.PageSetup.PageHeight = Me.ScaleY(Editor1.PaperHeight, vbTwips, vbPoints)
                
                If WordApp.ActiveWindow.ActivePane.View.Type = 1 Or WordApp.ActiveWindow.ActivePane.View.Type = 2 Then
                    WordApp.ActiveWindow.ActivePane.View.Type = 3
                    'wdNormalView=1     wdOutlineView=2     wdPrintView=3
                End If
                
                WordApp.ActiveWindow.View = 5   'wdMasterView
                '��ӵ�ǰ��ҳüҳ�ŵ�RTF�ļ���
                WordApp.ActiveWindow.View.SeekView = 9  'wdSeekCurrentPageHeader
                WordApp.Selection.ParagraphFormat.Alignment = 0     'wdAlignParagraphLeft
                If Not (Editor1.Picture Is Nothing) Then
                    If Editor1.Picture.Handle <> 0 Then
                        strPicFile = GetSysTmpPath & "\zlDocHead" & App.ThreadID & ".BMP"
                        If gobjFSO.FileExists(strPicFile) Then gobjFSO.DeleteFile strPicFile, True
                        SavePicture Editor1.Picture, strPicFile
                        If gobjFSO.FileExists(strPicFile) Then
                            WordApp.Selection.InlineShapes.AddPicture Filename:=strPicFile, LinkToFile:=False, SaveWithDocument:=True
                            gobjFSO.DeleteFile strPicFile, True
                            WordApp.Selection.TypeParagraph
                        End If
                    End If
                End If
                
                edtBuff.HeadFileTextRTF = Editor1.HeadFileTextRTF: edtBuff.DocHeadReplaceKey: edtBuff.DocHeadCopyWithFormat
                WordApp.Selection.Paste
                'ȥ�� ���е���ҳ��,ҳ��
                WordApp.Selection.Start = 0
                WordApp.Selection.End = 99999
                If WordApp.Selection.Find.Execute("{ҳ��}") Then
                    WordApp.Selection.Start = 99999
                    Set Fd = WordApp.Selection.Fields.Add(Range:=WordApp.Selection.Range, Type:=33)   'wdFieldPage
                    Fd.Copy
                    
                    WordApp.Selection.Start = 0
                    WordApp.Selection.End = 99999
                    WordApp.Selection.Find.Execute FindText:="{ҳ��}", ReplaceWith:="^c"
                    Fd.Cut
                    Clipboard.Clear
                End If
    
                WordApp.Selection.Start = 0
                WordApp.Selection.End = 99999
                If WordApp.Selection.Find.Execute("{��ҳ��}") Then
                    WordApp.Selection.Start = 99999
                    Set Fd = WordApp.Selection.Fields.Add(Range:=WordApp.Selection.Range, Type:=26)   'wdFieldNumPages
                    Fd.Copy
                    
                    WordApp.Selection.Start = 0
                    WordApp.Selection.End = 99999
                    WordApp.Selection.Find.Execute FindText:="{��ҳ��}", ReplaceWith:="^c"
                    Fd.Cut
                    Clipboard.Clear
                End If
                
                WordApp.ActiveWindow.View.SeekView = 10 'wdSeekCurrentPageFooter'ҳ��
                edtBuff.FootFileTextRTF = Editor1.FootFileTextRTF: edtBuff.DocFootReplaceKey: edtBuff.DocFootCopyWithFormat
                WordApp.Selection.Paste
                'ȥ�� ���е���ҳ��,ҳ��
                WordApp.Selection.Start = 0
                WordApp.Selection.End = 99999
                If WordApp.Selection.Find.Execute("{ҳ��}") Then
                    WordApp.Selection.Start = 99999
                    Set Fd = WordApp.Selection.Fields.Add(Range:=WordApp.Selection.Range, Type:=33)   'wdFieldPage
                    Fd.Copy
                    
                    WordApp.Selection.Start = 0
                    WordApp.Selection.End = 99999
                    WordApp.Selection.Find.Execute FindText:="{ҳ��}", ReplaceWith:="^c"
                    Fd.Cut
                    Clipboard.Clear
                End If
    
                WordApp.Selection.Start = 0
                WordApp.Selection.End = 99999
                If WordApp.Selection.Find.Execute("{��ҳ��}") Then
                    WordApp.Selection.Start = 99999
                    Set Fd = WordApp.Selection.Fields.Add(Range:=WordApp.Selection.Range, Type:=26)   'wdFieldNumPages
                    Fd.Copy
                    
                    WordApp.Selection.Start = 0
                    WordApp.Selection.End = 99999
                    WordApp.Selection.Find.Execute FindText:="{��ҳ��}", ReplaceWith:="^c"
                    Fd.Cut
                    Clipboard.Clear
                End If

                Set Fd = Nothing
                WordApp.ActiveWindow.View.SeekView = 3      'wdPrintView
                WordDoc.PrintPreview
                WordApp.ScreenUpdating = True
                WordApp.Visible = True
                WordApp.Activate
                
                Do
                    DoEvents
                    If Not WordDoc.Windows.Item(WordDoc.Windows.Count).View = 4 Then Exit Do    'wdPrintPreview=4
                Loop
                
                zlCommFun.StopFlash
                Screen.MousePointer = vbDefault
            End If
            
            WordDoc.Close False
            WordApp.Quit
            Set WordDoc = Nothing
            Set WordApp = Nothing
            gobjFSO.DeleteFile strF, True
        End If
    End If
End Sub

Private Sub InsertHeadFootInWord(WordApp As Object, strR As String)
    Dim blnFinded As Boolean, i As Long, j As Long, k As Long
    i = 1
    j = Len(strR)
    k = i
    Do While (k <= j)
        If Mid(strR, k, 1) = "{" Then
            If Mid(strR, k, 4) = "{ҳ��}" Then
                WordApp.Selection.Fields.Add Range:=WordApp.Selection.Range, Type:=33   'wdFieldPage
                WordApp.Selection.Start = 999999
                k = k + 4
            ElseIf Mid(strR, k, 5) = "{��ҳ��}" Then
                WordApp.Selection.Fields.Add Range:=WordApp.Selection.Range, Type:=26   'wdFieldNumPages
                WordApp.Selection.Start = 999999
                k = k + 5
            Else
                WordApp.Selection.TypeText Mid(strR, k, 1)
                k = k + 1
            End If
        ElseIf Mid(strR, k, 2) = vbCrLf Then
            WordApp.Selection.TypeText Mid(strR, k, 1)
            k = k + 2
        Else
            WordApp.Selection.TypeText Mid(strR, k, 1)
            k = k + 1
        End If
    Loop
End Sub


'################################################################################################################
'## ���ܣ�  �����ĵ���RTF�ļ�
'##
'## ������  strFileName     ���ļ���
'##         blnClearMode    ���Ƿ������ģʽ
'################################################################################################################
Public Function SaveDocToFile(Optional ByVal strFileName As String = "", _
    Optional blnClearMode As Boolean = True, _
    Optional blnClearKeywords As Boolean = True) As Boolean

    On Error GoTo LL
    Dim strF As String
    If strFileName = "" Then
        If Editor1.ViewMode = cprPaper Then Exit Function
        Select Case Me.Document.EditType
        Case cprET_�����ļ�����
            dlgThis.Filename = "����_" & Me.Document.EPRFileInfo.���� & ".rtf"
        Case cprET_ȫ��ʾ���༭
            dlgThis.Filename = "����_" & Me.Document.EPRFileInfo.���� & "_" & Me.Document.EPRDemoInfo.���� & ".rtf"
        Case cprET_�������༭, cprET_���������
            dlgThis.Filename = "��¼_" & Me.Document.EPRFileInfo.���� & "(" & Me.Document.EPRPatiRecInfo.ID & "," & Me.Document.Ŀ��汾 & ").rtf"
        End Select
        dlgThis.Filter = "*.rtf|*.rtf|*.*|*.*"
        dlgThis.ShowSave
        strF = dlgThis.Filename
    Else
        strF = strFileName
    End If
    If strF <> "" Then
        '=================================================================================================
        Dim lngLen As Long, blnReadOnly As Boolean
        With Me.edtBuff
            .NewDoc
            blnReadOnly = .ReadOnly
            .ReadOnly = False
            .ForceEdit = True
            lngLen = Len(Me.Editor1.Text)
            'RTF���ݸ�ֵ
            Me.Editor1.SaveDoc strF
            .OpenDoc strF
'            .TOM.TextDocument.Selection.FormattedText = Me.Editor1.TOM.TextDocument.Range(0, lngLen).FormattedText
'            .TOM.TextDocument.Range(lngLen, lngLen).Para = Me.Editor1.TOM.TextDocument.Range(lngLen, lngLen).Para.Duplicate

            '������йؼ���
            Dim i As Long
            Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
            i = 0
            If blnClearKeywords Then
                bFinded = FindNextAnyKey(Me.edtBuff, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Do While bFinded
                    .Range(lKSS, lKSE) = ""
                    .Range(lKSS + lKES - lKSE, lKSS + lKES - lKSE + 16) = ""
                    i = lKSS + lKES - lKSE
                    bFinded = FindNextAnyKey(Me.edtBuff, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
                Loop
            End If
            .SelectAll
            If blnClearMode Then
                .AuditMode = True
                .AcceptAuditText    '���ģʽ
            End If
            lngLen = Len(.Text)
            For i = 0 To lngLen - 1
                'ֻ������ɫΪҪ�ر���ɫ��ɫȥ��
                If .Range(i, i + 1).Font.BackColor = ELE_BACKCOLOR Then
                    .Range(i, i + 1).Font.BackColor = tomAutoColor
                End If
                If .Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR Then
                    .Range(i, i + 1).Font.ForeColor = tomAutoColor
                End If
            Next
            .ReadOnly = blnReadOnly
            '���浽�ļ�
            .SaveDoc strF
        End With
    End If
    SaveDocToFile = True
    Exit Function
LL:
    SaveDocToFile = False
End Function
'################################################################################################################
'## ���ܣ�  ��ӡ����ĵ�
'################################################################################################################
Private Function PrintEPRDoc(ByVal blnPreview As Boolean, Optional ByVal blnClearMode As Boolean = False) As Boolean
'������blnPreview��Ԥ��
'      blnClearMode-���ո�ʽ(���޸ĺۼ�)
    Dim intLoop As Integer
    Dim lngLen As Long, strF As String
    Dim rsTemp As ADODB.Recordset
    Dim strBillNo As String
    Dim strExseNo As String, intExseKind As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim strPicPath As String, strPicFile As String
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim aryPara(19) As String, intPCount As Integer
    Dim aryFlagPara(1) As String
    Dim intRows As Integer, intCols As Integer
    Dim dcmImages As New DicomImages, dcmResultImage As DicomImage
    Dim blnNoAsk As Boolean
    
    zlCommFun.ShowFlash "���Ժ�..."
    Screen.MousePointer = vbHourglass
    Err.Clear
    On Error GoTo errHand
    
    blnNoAsk = (zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1")

    If Me.Document.EPRFileInfo.���� = cpr���Ʊ��� And Me.Document.EPRFileInfo.ͨ�� = 2 Then
        strBillNo = "ZLCISBILL" & Format(Document.EPRFileInfo.���, "00000") & "-2"
    
        gstrSQL = "Select ��¼����, No From ����ҽ������ Where ҽ��id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡNO", CLng(Document.EPRPatiRecInfo.ҽ��id))
        If rsTemp.RecordCount = 0 Then zlCommFun.StopFlash: Screen.MousePointer = vbDefault: Exit Function
        strExseNo = "" & rsTemp!NO
        intExseKind = Val("" & rsTemp!��¼����)
        
        If mobjReport Is Nothing Then Set mobjReport = New clsReport
        If Not blnNoAsk Then
            If mobjReport.ReportPrintSet(gcnOracle, glngSys, strBillNo, Me) = False Then zlCommFun.StopFlash: Screen.MousePointer = vbDefault: Exit Function
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
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡID", CLng(Document.EPRPatiRecInfo.ID))
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If cTable.GetTableFromDB(cprET_���������, CLng(Document.EPRPatiRecInfo.ID), Val("" & rsTemp!���Id)) Then
                For intLoop = 1 To cTable.Pictures.Count
                    strPicFile = strPicPath & "PACSPic" & intLoop & ".JPG"
                    If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True
                    If cTable.Pictures(intLoop).PictureType = EPRMarkedPicture Then
                        Set oPicture = cTable.Pictures(intLoop).DrawFinalPic
                    Else
                        Set oPicture = cTable.Pictures(intLoop).OrigPic
                    End If
                    SavePicture oPicture, strPicFile
                    If objFile.FileExists(strPicFile) Then
                        '������ͼ��ͼ���·��
                        If cTable.Pictures(intLoop).PictureType = EPRMarkedPicture Then
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
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", strBillNo)
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
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����", strBillNo)
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
        For intLoop = intPCount To UBound(aryPara) '�����е�ͼ�αȱ�������
            If aryPara(intLoop) Like "*=*" Then aryPara(intLoop) = ""
        Next
            
        '���ñ���
        Call mobjReport.ReportOpen(gcnOracle, glngSys, strBillNo, Me, _
            "NO=" & strExseNo, "����=" & intExseKind, "ҽ��ID=" & CLng(Document.EPRPatiRecInfo.ҽ��id), aryFlagPara(0), _
            aryPara(0), aryPara(1), aryPara(2), aryPara(3), aryPara(4), aryPara(5), _
            aryPara(6), aryPara(7), aryPara(8), aryPara(9), aryPara(10), aryPara(11), _
            aryPara(12), aryPara(13), aryPara(14), aryPara(15), aryPara(16), _
            aryPara(17), aryPara(18), aryPara(19), IIf(blnPreview, 1, 2))
    Else
        Set mfrmPreview = New frmPrintPreview
        Call mfrmPreview.DoSingleDocPreview(Me.Editor1, Me, Me.Document, blnClearMode, blnPreview, blnNoAsk)
        Unload mfrmPreview
        Set mfrmPreview = Nothing
    End If
    '=================================================================================================
    zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    PrintEPRDoc = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'################################################################################################################
'## ���ܣ�  �����ٵ�ָ��λ��
'################################################################################################################
Public Function InsertCompend(ByVal lStart As Long, ByVal lEnd As Long, ByRef objCompend As cEPRCompend, Optional blnFirstIns As Boolean = True) As Boolean
    Dim strTmp As String, lLen As Long, lngKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean

    With Editor1
        bBeteenKeys = IsBetweenAnyKeys(Editor1, lStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            '��ǰλ����ĳ���ؼ��ֶ�֮�䣬�����������٣�
            InsertCompend = False
            Exit Function
        End If
        '������
        lLen = Len(objCompend.����)
        objCompend.InsertIntoEditor Editor1, , blnFirstIns, Me.Document
        mfrmCompends_NodeSelected objCompend.Ԥ�����ID
    End With
    InsertCompend = True
End Function

'################################################################################################################
'## ���ܣ�  �����޸ĺ�����
'################################################################################################################
Public Function ModifyCompend(objCompend As cEPRCompend) As Boolean
    Dim strTmp As String, lIndex As Long, lLen As Long
    Dim lKey As Long, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, sKeyType As String, bNeeded As Boolean, bFinded As Boolean

    With Editor1
        If .ViewMode <> cprNormal Then ModifyCompend = False: Exit Function
        lKey = objCompend.Key
        bFinded = FindKey(Editor1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded = False Then
            ModifyCompend = False
            Exit Function
        End If
        .Freeze
        .Tag = "��ֹͬ��"
        .ForceEdit = True
        .Range(lKSS, lKEE) = ""
        objCompend.InsertIntoEditor Editor1, lKSS, False, Me.Document

        .Tag = ""
        .ForceEdit = False
        .UnFreeze
    End With
    ModifyCompend = True
End Function

'################################################################################################################
'## ���ܣ�  ɾ��һ�����
'################################################################################################################
Public Function DeleteOutline(lKey As Long) As Boolean
    Dim strTmp As String, lIndex As Long, lLen As Long, lLevel As Long, lNextKey As Long
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, sKeyType As String, bNeeded As Boolean, bFinded As Boolean

    With Editor1
        If .ViewMode <> cprNormal Then DeleteOutline = False: Exit Function
        'ȷ���ĵ��д��ڸ����
        bFinded = FindKey(Editor1, "O", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded = False Then
            DeleteOutline = False
            Exit Function
        End If
        Dim lngR As Long

        If Document.Compends("K" & lKey).�������� = True And Me.Document.EditType <> cprET_�����ļ����� Then
            MsgBox "����ɾ��������٣�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If

'        If Document.Compends("K" & lKey).Ԥ�����ID <> 0 Then
'            lngR = MsgBox("ȷ��ɾ�����Ƿ�ͬʱɾ��Ԥ����� [" & Document.Compends("K" & lKey).���� & "] ���������¼����ݣ�" _
'                , vbYesNo + vbQuestion + vbDefaultButton2, "ȷ��ɾ��")
'            If lngR = vbNo Then lngR = vbCancel
'        Else
        lngR = MsgBox("ȷ��ɾ�����Ƿ�ͬʱɾ����� [" & Document.Compends("K" & lKey).���� & "] �������¼����ݣ�" & vbCrLf & _
            "ע�������������ϲ���һ����ٵ������С�", vbYesNoCancel + vbQuestion + vbDefaultButton3, "ȷ��ɾ��")
'        End If
        If lngR = vbNo Then
            .InProcessing = True
            .Tag = "DeleteOutline"
            .Freeze
            .ForceEdit = True
            lLevel = Document.Compends("K" & lKey).Level
            Document.Compends.Remove "K" & lKey
            .Range(lKSS, lKEE) = ""
            .ForceEdit = False
            .Tag = ""
            .UnFreeze
            .InProcessing = False
            Document.Compends.CheckValidParentKeys '��鸸Key����Ч�ԣ�
            Document.Compends.FillTree mfrmCompends.Tree
            DeleteOutline = True
            Editor1.SelLength = 0
        ElseIf lngR = vbYes Then
            .InProcessing = True
            .Tag = "DeleteOutline"
            .Freeze
            .ForceEdit = True
            lLevel = Document.Compends("K" & lKey).Level
            Document.Compends.Remove "K" & lKey
            Dim i As Long, sText As String
            sText = Editor1.Text
            lLen = Len(sText)
            i = lKEE
LL1:
            '�����м����������Ԫ�أ���ɾ����

            i = InStr(i, sText, "OS", vbTextCompare)
            If i <> 0 Then
                If .Range(i - 1, i).Font.Hidden = False Then   '��Ϊ�ؼ��֣��������������ܱ����ġ�
                    i = i + 1
                    GoTo LL1
                End If
                lNextKey = Val(.Range(i + 2, i + 10))
                If Document.Compends("K" & lNextKey).��Key = 0 Then
                    Document.Compends("K" & lNextKey).Level = 1
                End If
                If Document.Compends("K" & lNextKey).Level > lLevel Then
                    '���С�ڵ�ǰ��Σ�����������٣�����
                    Document.Compends.Remove "K" & lNextKey
                    i = i + 1
                    GoTo LL1
                End If
                Call ClearObjectsInArea(sText, lKSS, i - 1)
                .Range(lKSS, i - 1) = ""
            Else
                Call ClearObjectsInArea(sText, lKSS, lLen)
                .Range(lKSS, lLen) = ""
            End If
            .ForceEdit = False
            .UnFreeze
            .Tag = ""
            .InProcessing = False
            Document.Compends.CheckValidParentKeys '��鸸Key����Ч�ԣ�
            Document.Compends.FillTree mfrmCompends.Tree
            DeleteOutline = True
            Editor1.SelLength = 0
        Else
            '������ɾ������
        End If
        Call RefSentenceList
    End With
End Function

'################################################################################################################
'   ��;��  �������(lngStart,lngEnd)�ڵ�����ͼƬ������Ҫ�غͱ����󡣣��������ĳ������ڲ����ж���
'################################################################################################################
Private Sub ClearObjectsInArea(ByRef StrText As String, ByVal lngStart As Long, ByVal lngEnd As Long)
    Dim lLen As Long, i As Long, lngKey As Long, blnForce As Boolean
    With Editor1
        blnForce = .ForceEdit
        .Tag = "ClearObjectinArea"
        .Freeze
        .ForceEdit = True
        lLen = Len(StrText)
        '�����м����������Ԫ�أ���ɾ����
        i = IIf(lngStart = 0, 1, lngStart)
        i = InStr(i, StrText, "ES", vbTextCompare)
        Do While i > lngStart And i < lngEnd
            If .Range(i - 1, i).Font.Hidden Then    '��Ϊ�ؼ��֣��������������ܱ����ġ�
                lngKey = Val(.Range(i + 2, i + 10))
                Document.Elements.Remove "K" & lngKey
            End If
            i = i + 1
            i = InStr(i, StrText, "ES", vbTextCompare)
        Loop
        i = IIf(lngStart = 0, 1, lngStart)
        i = InStr(i, StrText, "PS", vbTextCompare)
        Do While i > lngStart And i < lngEnd
            If .Range(i - 1, i).Font.Hidden Then    '��Ϊ�ؼ��֣��������������ܱ����ġ�
                lngKey = Val(.Range(i + 2, i + 10))
                Document.Pictures.Remove "K" & lngKey
            End If
            i = i + 1
            i = InStr(i, StrText, "PS", vbTextCompare)
        Loop
        i = IIf(lngStart = 0, 1, lngStart)
        i = InStr(i, StrText, "TS", vbTextCompare)
        Do While i > lngStart And i < lngEnd
            If .Range(i - 1, i).Font.Hidden Then    '��Ϊ�ؼ��֣��������������ܱ����ġ�
                lngKey = Val(.Range(i + 2, i + 10))
                Document.Tables.Remove "K" & lngKey
            End If
            i = i + 1
            i = InStr(i, StrText, "TS", vbTextCompare)
        Loop
        .ForceEdit = blnForce
        .Tag = ""
        .UnFreeze
    End With
End Sub

'################################################################################################################
'   ��;��  ����ͼƬ��
'################################################################################################################
Public Function InsertPicture(bytPicType As Integer, objPic As StdPicture, lWidth As Long, lHeight As Long, Optional strOther As String) As Boolean
    'IsMarked:�Ƿ��Ǳ��ͼ
    'objPic  :ͼƬ����
    'lRow    :���ڱ���е�ͼƬ�󶨣���ʾ��
    'lCol    :���ڱ���е�ͼƬ�󶨣���ʾ��
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngKey As Long
    If tblThis.Visible Then '        '����еĲ���ͼ��
        If Val(tblThis.Tag) > 0 Then
            '��ͼƬ���󱣴浽�������
            If Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag) = 0 Then
                lngKey = Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures.Add
            Else
                lngKey = Val(tblThis.Cells("K" & tblThis.SelectedCellKey).Tag)
            End If
            Set Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).OrigPic = objPic
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Width = lWidth
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Height = lHeight
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).OrigWidth = lWidth
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).OrigHeight = lHeight
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).PictureType = bytPicType
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Row = tblThis.Cells("K" & tblThis.SelectedCellKey).Row
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).Col = tblThis.Cells("K" & tblThis.SelectedCellKey).Col
            Me.Document.Tables("K" & Val(tblThis.Tag)).Pictures("K" & lngKey).�����ı� = strOther
            
            '���浽��Ԫ����
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = ""
            tblThis.Cells("K" & tblThis.SelectedCellKey).Tag = lngKey
            tblThis.Cells("K" & tblThis.SelectedCellKey).Picture = objPic
            tblThis.Cells("K" & tblThis.SelectedCellKey).ToolTipText = ""
            tblThis.Cells("K" & tblThis.SelectedCellKey).Protected = True
            tblThis.Modified = True
            tblThis.Refresh
            tblThis_Resize tblThis.Width, tblThis.Height
        End If
    ElseIf ucPacsImgCanvas1.Visible Then '�ڱ�����в���ͼ��
        If ucPictureEditor1.Visible Then ucPictureEditor1.Modified = False: ucPictureEditor1.CloseMe
        If ucPacsImgCanvas1.MarkedPicPosition = 0 Then ucPacsImgCanvas1.MarkedPicPosition = 1
        ucPacsImgCanvas1.AddMarkedPicture objPic, ucPacsImgCanvas1.MarkedPicPosition
    Else '�����е�ͼ
        bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then InsertPicture = False: Exit Function    '��֤���ܲ���ؼ����ڲ�
        '��ͼƬ���󱣴浽�������
        lngKey = Document.Pictures.Add()
        Set Document.Pictures("K" & lngKey).OrigPic = objPic
        Document.Pictures("K" & lngKey).Width = lWidth
        Document.Pictures("K" & lngKey).Height = lHeight
        Document.Pictures("K" & lngKey).OrigWidth = lWidth
        Document.Pictures("K" & lngKey).OrigHeight = lHeight
        Document.Pictures("K" & lngKey).PictureType = bytPicType
        Document.Pictures("K" & lngKey).InsertIntoEditor Editor1
        Document.Pictures("K" & lngKey).�����ı� = strOther
    End If
    InsertPicture = True
End Function

'################################################################################################################
'## ���ܣ�  ��ʾ����Ҫ�ر༭��  '������
'################################################################################################################
Private Sub ShowEleEditor(KeyAscii As Integer, Shift As Integer)
    On Error Resume Next
    If Editor1.ViewMode <> cprNormal Then Exit Sub
'    If Me.Editor1.AuditMode Then Exit Sub
    glngCurEleKey = 0
'    picSmartSignal.Visible = False
    '�жϵ�ǰλ���Ƿ��� CS �� CE ֮�䣺
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenKeys(Editor1, Editor1.Selection.StartPos + 1, "E", lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    
    
    'ǩ��Ҫ�ؽ�ֹ�༭
    '------------------------------------------------------------------------------------------------------------------
    If Document.Elements("K" & lKey).�滻�� = 1 Then
        Select Case Document.Elements("K" & lKey).Ҫ������
        Case "����ҽʦǩ��", "����ҽʦǩ��", "����ҽʦǩ��"
            Exit Sub
        End Select
    End If
    
    Dim pt As POINTAPI
    pt.X = 0
    pt.y = 0
    ClientToScreen Editor1.OriginRTB.hwnd, pt

    If bBeteenKeys Then
        '��ʱ��¼����λ��
        If Me.Editor1.AuditMode Then
            If Document.Elements("K" & lKey).��ʼ�� < Me.Document.Ŀ��汾 Then Exit Sub
        End If
        glngCurEleKey = lKey
        Dim lLeft As Long, lTOp As Long, lRight As Long
        '��ȡ��ʼλ������
        Editor1.Range(Editor1.Selection.StartPos, Editor1.Selection.StartPos + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp

        '��ʾ�༭�ؼ�
        If Editor1.Range(lKEE, lKEE + 2) = vbCrLf Then
            Document.Elements("K" & lKey).�Ƿ��� = True
        Else
            Document.Elements("K" & lKey).�Ƿ��� = False
        End If
        If Document.Elements("K" & lKey).�滻�� = 2 Then
            '�ֵ���Ŀ
            mfrmDicSelect.ShowMe Document.Elements("K" & lKey).Ҫ������, pt.X * Screen.TwipsPerPixelX + lLeft, _
                pt.y * Screen.TwipsPerPixelY + lTOp, vbModeless, Me, Document.Elements("K" & lKey).�����ı�
        Else
            '����Ҫ��
            mfrmModElement.Tag = Editor1.Selection.StartPos
            mfrmModElement.ShowMe Document.Elements("K" & lKey), _
                pt.X * Screen.TwipsPerPixelX + lLeft, _
                pt.y * Screen.TwipsPerPixelY + lTOp, IIf(mintStyle = -1, 0, mintStyle), Me, Me.Document.EditType
        End If
        If Chr(KeyAscii) <> " " And Chr(KeyAscii) <> Chr(13) And KeyAscii <> 0 Then SendKeys Chr(KeyAscii)
    Else
        '���򣬶�λ����һ������Ҫ�ص�λ��
        bBeteenKeys = FindNextKey(Editor1, 1, "E", lKey, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bBeteenKeys Then
            Editor1.Range(lKSE, lKES).Selected
        End If

        glngCurEleKey = 0
    End If
End Sub
Private Sub mfrmSentenceDetailed_RowDblClick(ByVal lngSentenceID As Long)
    '˫������ʾ���ʾ�
    If Me.Editor1.ViewMode <> cprNormal Or Me.Editor1.ReadOnly Then Exit Sub
    If Me.Editor1.Selection.Font.Protected And tblThis.Visible = False Then Exit Sub
    If tblThis.Visible Then
        If tblThis.SelectedCellKey > 0 Then If tblThis.Cells("K" & tblThis.SelectedCellKey).Protected Then Exit Sub
    End If

    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys And tblThis.Visible = False Then Exit Sub
    Dim blnForce As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, strTmp As String, rsTemp As New ADODB.Recordset
    Dim lngStartPos As Long, lngEndPos As Long, sText As String

    If lngSentenceID <= 0 Then Exit Sub
    mfrmSentenceDetailed.Tag = lngSentenceID    '����ԭ���ļ�¼ID

    '�ʾ����ݻָ�
    gstrSQL = "Select �ʾ�id, ���д���, ��������, �����ı�, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, Ҫ��ֵ��, ������̬, ��������" & vbNewLine & _
                "From �����ʾ����" & vbNewLine & _
                "Where �ʾ�id = [1]" & vbNewLine & _
                "Order By ���д���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", lngSentenceID)
    With Editor1
        .Tag = "mfrmSentenceDetailed_RowDblClick"
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
        lngStartPos = .Selection.StartPos
        Do While Not rsTemp.EOF
            Select Case rsTemp("��������")
            Case 0 '��������
                '�ָ�RTF����
                lngStart = .Selection.StartPos
                strTmp = NVL(rsTemp("�����ı�"))
                lngLen = Len(strTmp)

                If tblThis.Visible Then
                    sText = sText & strTmp
                Else
                    .Range(lngStart, lngStart) = strTmp
                    .Range(lngStart, lngStart + lngLen).Font.Protected = False
                    .Range(lngStart, lngStart + lngLen).Font.Hidden = False
                    .Range(lngStart, lngStart + lngLen).Font.ForeColor = IIf(Me.Editor1.AuditMode, GetCharColor(Me.Document.Ŀ��汾, 0), tomAutoColor)
                    .Range(lngStart, lngStart + lngLen).Font.Strikethrough = False
                    .Range(lngStart, lngStart + lngLen).Font.BackColor = tomAutoColor
                    .Range(lngStart + lngLen, lngStart + lngLen).Selected
                End If
            Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                If tblThis.Visible Then
                    If NVL(rsTemp("�滻��"), 0) = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                        strTmp = GetReplaceEleValue(NVL(rsTemp("Ҫ������")), Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID, Me.Document.EPRPatiRecInfo.������Դ, Me.Document.EPRPatiRecInfo.ҽ��id, Me.Document.EPRPatiRecInfo.Ӥ��)
                    Else
                        strTmp = "{" & NVL(rsTemp("Ҫ������")) & "}"
                    End If
                    sText = sText & strTmp
                Else
                    lngStart = .Selection.StartPos
                    lngKey = Me.Document.Elements.Add
                    Me.Document.Elements("K" & lngKey).ID = 0       '���ǣ� NVL(rsTemp("�ʾ�ID"), 0) �����IDֵ��ͬ������
                    Me.Document.Elements("K" & lngKey).�����ı� = NVL(rsTemp("�����ı�"))
                    Me.Document.Elements("K" & lngKey).Ҫ������ = NVL(rsTemp("Ҫ������"))
                    Me.Document.Elements("K" & lngKey).����Ҫ��ID = NVL(rsTemp("����Ҫ��ID"), 0)
                    Me.Document.Elements("K" & lngKey).�滻�� = NVL(rsTemp("�滻��"), 0)
                    Me.Document.Elements("K" & lngKey).Ҫ������ = NVL(rsTemp("Ҫ������"), 0)
                    Me.Document.Elements("K" & lngKey).Ҫ�س��� = NVL(rsTemp("Ҫ�س���"), 0)
                    Me.Document.Elements("K" & lngKey).Ҫ��С�� = NVL(rsTemp("Ҫ��С��"), 0)
                    Me.Document.Elements("K" & lngKey).Ҫ�ص�λ = NVL(rsTemp("Ҫ�ص�λ"))
                    Me.Document.Elements("K" & lngKey).Ҫ�ر�ʾ = NVL(rsTemp("Ҫ�ر�ʾ"), 0)
                    Me.Document.Elements("K" & lngKey).Ҫ��ֵ�� = NVL(rsTemp("Ҫ��ֵ��"))
                    Me.Document.Elements("K" & lngKey).������̬ = NVL(rsTemp("������̬"), 0)
                    Me.Document.Elements("K" & lngKey).�Ƿ��� = False
                    Me.Document.Elements("K" & lngKey).�������� = NVL(rsTemp!��������)
                    If Me.Document.Elements("K" & lngKey).�滻�� = 1 And (Me.Document.EditType = cprET_�������༭ Or Me.Document.EditType = cprET_���������) Then
                        Me.Document.Elements("K" & lngKey).�����ı� = GetReplaceEleValue(Me.Document.Elements("K" & lngKey).Ҫ������, _
                            Me.Document.EPRPatiRecInfo.����ID, _
                            Me.Document.EPRPatiRecInfo.��ҳID, _
                            Me.Document.EPRPatiRecInfo.������Դ, _
                            Me.Document.EPRPatiRecInfo.ҽ��id, _
                            Me.Document.EPRPatiRecInfo.Ӥ��)
    '                    If Me.Document.Elements("K" & lngKey).�����ı� = "" Then Me.Document.Elements("K" & lngKey).�����ı� = "    "
                    End If
                    Me.Document.Elements("K" & lngKey).��ʼ�� = Me.Document.Ŀ��汾
                    Me.Document.Elements("K" & lngKey).InsertIntoEditor Editor1, lngStart
                End If
            End Select
            rsTemp.MoveNext
        Loop
        lngEndPos = .Selection.StartPos
        .ForceEdit = False
        If tblThis.Visible Then
            sText = tblThis.Cells("K" & tblThis.SelectedCellKey).Text & sText
            tblThis.Cells("K" & tblThis.SelectedCellKey).Text = sText
            tblThis.Modified = True
            tblThis.Refresh False, True, tblThis.SelectedCellKey
            tblThis_Resize tblThis.Width, tblThis.Height
        Else
            .Range(lngEndPos, lngEndPos).Selected
        End If
        .UnFreeze
        .Tag = ""
        .SetFocus
    End With
    Call RecountPage
    If Me.Editor1.Visible And Me.Editor1.Enabled And tblThis.Visible = False Then Me.Editor1.SetFocus
End Sub

Private Sub mfrmStyleMan_DblClick(ByVal lngStyleCode As Long)
    '�ı䵱ǰѡ�����ݵĶ�����ʽ
    SetCommonStyle Me.Editor1, lngStyleCode, Me.Editor1.Selection.StartPos, Me.Editor1.Selection.EndPos, True
    Call RecountPage
End Sub

Private Sub picPatiInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If X > 0 And X < picPatiInfo.ScaleWidth And y > 0 And y < picPatiInfo.ScaleHeight Then
        If picPatiInfo.Tag = "" Then
            SetCapture picPatiInfo.hwnd
            picPatiInfo.Cls
            picPatiInfo.BackColor = &HD2BDB6    ' &HD8D5D4 ' &HD2BDB6
            picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H6A240A, B
            picPatiInfo.Tag = "Captured"
        End If
    Else
        ReleaseCapture
        picPatiInfo.Cls
        picPatiInfo.BackColor = &H8000000F
        picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
        picPatiInfo.Tag = ""
    End If
End Sub

Private Sub picPatiInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    ReleaseCapture
    picPatiInfo.Cls
    picPatiInfo.BackColor = &H8000000F
    picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picPatiInfo.Tag = ""
End Sub

Private Sub picPatiInfo_Resize()
'    ReleaseCapture
    picPatiInfo.Cls
    picPatiInfo.BackColor = &H8000000F
    picPatiInfo.Line (0, 0)-(picPatiInfo.ScaleWidth - Screen.TwipsPerPixelX, picPatiInfo.ScaleHeight - Screen.TwipsPerPixelY), &H999999, B
    picPatiInfo.Tag = ""
End Sub

'################################################################################################################
'   ��;��  ��̬���¹���������ɫ��ͼ�ꡣ
'################################################################################################################
Private Sub SetColorIcon(Key As String, ID As Long, COLOR As OLE_COLOR)
    Dim ctlPictureBox As VB.PictureBox
    Set ctlPictureBox = Controls.Add("VB.PictureBox", "ctlPictureBox1")
    Dim ListImage As ListImage
    Set ListImage = imgColor.ListImages(Key)

    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = imgColor.MaskColor

    ctlPictureBox.Picture = ListImage.ExtractIcon

    If COLOR = vbWhite Then COLOR = RGB(254, 254, 254)
    ctlPictureBox.Line (1, ctlPictureBox.Height * 0.6)-(ctlPictureBox.Width, ctlPictureBox.Height), COLOR, BF
    ctlPictureBox.Refresh

    'Replace icon
    imgColor.ListImages.Remove imgColor.ListImages(Key).Index
    imgColor.ListImages.Add 1, Key, ctlPictureBox.Image
'    Set imgColor.ListImages(Key).Picture = ctlPictureBox.Image

    'OK Now replace Tag property
    imgColor.ListImages(1).Tag = ID

    cbrThis.AddImageList imgColor
    cbrThis.RecalcLayout

        Set ctlPictureBox.Picture = Nothing
    Me.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

'################################################################################################################
'   ��;��  ˢ�²�����Ϣ
'################################################################################################################
Public Sub RefreshPatiInfo()
    If Me.Document.EditType <> cprET_�������༭ And Me.Document.EditType <> cprET_��������� Then Exit Sub

    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo errHand
    If Me.Document.EPRPatiRecInfo.������Դ <> 2 Then
        gstrSQL = "Select ����,���֤��,Rpad('�����:' || �����, 18) || Rpad('����:' || ����, 18) ||" & vbNewLine & _
         "        Rpad('�Ա�:' || �Ա�, 10) || Rpad('����:' || ����,10) As ��Ϣ, ҽ����,�Ա� " & vbNewLine & _
         "From ������Ϣ" & vbNewLine & _
         "Where ����id = [1]"
    ElseIf (Document.EPRPatiRecInfo.ҽ��id <> 0 And Document.EPRPatiRecInfo.Ӥ�� <> 0) Then
        gstrSQL = "Select Decode([2], 2, RPad('סԺ��:' || c.סԺ��, 18) || RPad('����:' || c.��Ժ����, 15), RPad('�����:' || a.�����, 18)) ||" & vbNewLine & _
                    "        RPad('����:' || Nvl(b.Ӥ������, a.���� || '֮��'), 18) || RPad('�Ա�:' || b.Ӥ���Ա�, 10) || '����:' ||" & vbNewLine & _
                    "        To_Char(b.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ��Ϣ, A.ҽ����, b.Ӥ���Ա� �Ա�,a.����,a.���֤��" & vbNewLine & _
                    "From ������Ϣ A, ������ҳ C, ������������¼ B" & vbNewLine & _
                    "Where c.����id = [1] And c.��ҳid = [3] And a.����id = c.����id And c.����id = b.����id And c.��ҳid = b.��ҳid And b.��� = [4]"
    ElseIf (Document.EPRPatiRecInfo.Ӥ�� = 0) Then
           gstrSQL = "Select Decode([2], 2, RPad('סԺ��:' || b.סԺ��, 18) || RPad('����:' || b.��Ժ����, 15), RPad('�����:' || a.�����, 18)) ||" & vbNewLine & _
                        "        RPad('����:' || a.����, 18) || RPad('�Ա�:' || a.�Ա�, 10) || RPad('����:' || a.����, 10) As ��Ϣ, a.ҽ����, a.�Ա�,a.����,a.���֤��" & vbNewLine & _
                        "From ������Ϣ A, ������ҳ B" & vbNewLine & _
                        "Where b.����id = [1] And b.��ҳid = [3] And a.����id = b.����id"
    Else
        gstrSQL = "Select Decode([2], 2, RPad('ĸ��סԺ��:' || c.סԺ��, 18) || RPad('ĸ�״���:' || c.��Ժ����, 15), RPad('ĸ�������:' || a.�����, 18)) ||" & vbNewLine & _
                    "        RPad('����:' || Nvl(b.Ӥ������, a.���� || '֮Ӥ' || b.���), 30) || RPad('�Ա�:' || Nvl(b.Ӥ���Ա�, 'δ֪'), 10) || '����:' ||" & vbNewLine & _
                    "        To_Char(b.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ��Ϣ, a.ҽ����, a.�Ա�,a.����,a.���֤��" & vbNewLine & _
                    "From ������Ϣ A, ������ҳ C, ������������¼ B" & vbNewLine & _
                    "Where c.����id = [1] And c.��ҳid = [3] And a.����id = c.����id And c.����id = b.����id And c.��ҳid = b.��ҳid And b.��� = [4]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.������Դ, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.Ӥ��)
    If rsTemp.RecordCount > 0 Then
        mPatiInfor.���� = "" & rsTemp!����
        mPatiInfor.���֤�� = "" & rsTemp!���֤��
        Me.lblPatiInfo.Caption = "" & rsTemp!��Ϣ
        Me.lblPatiIns(1).Caption = "" & rsTemp!ҽ����
        mstrSex = "" & rsTemp!�Ա�
        mfrmDocksymbol.HideSomeThing IIf(InStr(mstrSex, "��") > 0, 1, IIf(InStr(mstrSex, "Ů") > 0, 2, 0))
    Else
        mstrSex = ""
        Me.lblPatiInfo.Caption = ""
        Me.lblPatiIns(1).Caption = ""
    End If
    Err = 0: On Error Resume Next
    lblPatiIns(0).Left = lblPatiInfo.Left + lblPatiInfo.Width + 50
    lblPatiIns(1).Left = lblPatiIns(0).Left + lblPatiIns(0).Width
    lblPatiState(0).Left = lblPatiIns(1).Left + lblPatiIns(1).Width
    lblPatiState(1).Left = lblPatiState(0).Left + lblPatiState(0).Width
    Err.Clear: On Error GoTo errHand
    
    If Me.Document.EPRPatiRecInfo.ҽ��id = 0 Then
        Me.lblPatiState(0).Caption = "����:"
        Select Case Me.Document.EPRPatiRecInfo.������Դ
        Case cprPF_����
            gstrSQL = "Select r.���� From ���˹Һż�¼ r Where r.Id = [1] and r.��¼����=1  and r.��¼״̬=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", Me.Document.EPRPatiRecInfo.��ҳID)
            If rsTemp.RecordCount > 0 Then
                Me.lblPatiState(1).Caption = IIf(Val("" & rsTemp!����) = 1, "��", "")
            Else
                Me.lblPatiState(1).Caption = ""
            End If
        Case cprPF_סԺ
            gstrSQL = "Select ��Ժ����, ��Ժ����, ��Ժ��ʽ From ������ҳ Where ����id = [1] And Nvl(��ҳid, 0) = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID)
            If rsTemp.RecordCount > 0 Then
                If IsNull(rsTemp!��Ժ����) Then
                    If "" & rsTemp!��Ժ���� = "һ��" Then
                        Me.lblPatiState(1).ForeColor = Me.lblPatiState(0).ForeColor
                    Else
                        Me.lblPatiState(1).ForeColor = RGB(255, 0, 0)
                    End If
                    Me.lblPatiState(1).Caption = "" & rsTemp!��Ժ����
                Else
                    If "" & rsTemp!��Ժ��ʽ <> "����" Then
                        Me.lblPatiState(1).ForeColor = Me.lblPatiState(0).ForeColor
                    Else
                        Me.lblPatiState(1).ForeColor = RGB(255, 0, 0)
                    End If
                    Me.lblPatiState(1).Caption = rsTemp!��Ժ��ʽ & "(��Ժ)"
                End If
            Else
                Me.lblPatiState(1).Caption = ""
            End If
        End Select
    Else
        Me.lblPatiState(0).Caption = "Ҫ��:"
        gstrSQL = "Select ������־ From ����ҽ����¼ Where Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ϣ", Me.Document.EPRPatiRecInfo.ҽ��id)
        If rsTemp.RecordCount > 0 Then
            Me.lblPatiState(1).Caption = IIf(Val("" & rsTemp!������־) = 1, "��", "")
        Else
            Me.lblPatiState(1).Caption = ""
        End If
    End If

    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'################################################################################################################
'   ��;��  ���뱾�ξ���ҽ�����༭���е�ǰλ��
'################################################################################################################
Public Function ImportDocAdvice() As Boolean
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
Dim rsTemp As New ADODB.Recordset

On Error GoTo errHand
    With Me.Editor1
        If .Selection.Font.Protected Then Exit Function
        bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then Exit Function
        gstrSQL = "Select ID,����,С��,��λ,�滻��,����,��̬��,���� From ����������Ŀ where ������='����ҽ��'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪ��")
        
        lKey = Me.Document.Elements.Add
        With Me.Document.Elements("K" & lKey)
            .Ҫ������ = "����ҽ��"
            .����Ҫ��ID = rsTemp!ID
            .Ҫ������ = rsTemp!����
            .Ҫ�س��� = NVL(rsTemp!����, 2)
            .Ҫ��С�� = NVL(rsTemp!С��, 2)
            .Ҫ�ص�λ = NVL(rsTemp!��λ, 2)
            .�滻�� = NVL(rsTemp!�滻��, 0)
            .���� = NVL(rsTemp!����, 0)
            .��̬�� = NVL(rsTemp!��̬��, 0)
            .�����ı� = GetReplaceEleValue(.Ҫ������, Document.EPRPatiRecInfo.����ID, Document.EPRPatiRecInfo.��ҳID, Document.EPRPatiRecInfo.������Դ, Document.EPRPatiRecInfo.ҽ��id, Me.Document.EPRPatiRecInfo.Ӥ��)
            .��ʼ�� = Me.Document.Ŀ��汾
            .InsertIntoEditor Me.Editor1, , True
        End With
    End With
    Me.Document.EleToString Me.Editor1, Me.Document.Elements("K" & lKey)
    ImportDocAdvice = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 '################################################################################################################
'   ��;��  ����һ��PacseͼƬ�飨���
'################################################################################################################
Public Sub InsertPacsPicTable()
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    bBeteenKeys = IsBetweenAnyKeys(Editor1, Editor1.Selection.StartPos + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then Exit Sub
    
    lKey = Me.Document.Tables.Add
    Me.Document.Tables("K" & lKey).TableType = tte_����ͼƬ��
    ucPacsImgCanvas1.ReadPicturesFromTable Me.Document.Tables("K" & lKey)
    
    Dim frmT As New frmTablePicCreator
    Me.Document.Tables("K" & lKey).InsertIntoEditor Editor1, , , True
    Unload frmT
    Set frmT = Nothing
End Sub

Private Sub txtContent_Change()
    If txtContent.Text <> txtContent.Tag Then
        mblnFBContentChanged = True
    End If
    txtContent.Tag = txtContent.Text
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtPenInput_Change()
    If Editor1.ReadOnly Then GoTo LL
    If Me.Editor1.Selection.Font.Protected = False And Me.Editor1.Selection.Font.Hidden = False And txtPenInput.Tag = "" Then
        txtPenInput.Tag = "InProcessing"
        Me.Editor1.ForceEdit = True
        If Me.Editor1.AuditMode Then
            On Error Resume Next
            Editor1.Range(Editor1.Selection.StartPos + Len(Editor1.Selection.Text), Editor1.Selection.StartPos + Len(Editor1.Selection.Text)).Selected
            Me.Editor1.SelLength = 0
            Me.Editor1.OriginRTB.SelColor = Me.Document.GetNewCharColor(Me.Editor1.OriginRTB.SelColor)
            Me.Editor1.OriginRTB.SelStrikeThru = False    'ȥ��ɾ����
            Me.Editor1.OriginRTB.SelUnderline = False     'ȥ���»���
            Me.Editor1.SelText = txtPenInput.Text
        Else
            Me.Editor1.SelText = txtPenInput.Text
        End If
        Me.Editor1.Range(Me.Editor1.Selection.EndPos, Me.Editor1.Selection.EndPos).Selected
        Me.Editor1.ForceEdit = False
    End If
LL:
    txtPenInput.Text = ""
    txtPenInput.Tag = ""
End Sub

Private Sub txtPenInput_GotFocus()
    txtPenInput.Tag = "InProcessing"
    txtPenInput.Text = ""
    txtPenInput.Tag = ""
End Sub

Private Sub txtPenInput_KeyPress(KeyAscii As Integer)
    If Editor1.ReadOnly Then Exit Sub
    Select Case KeyAscii
    Case vbKeyBack
        If Me.Editor1.AuditMode Then Exit Sub
        Dim lngStart As Long, lngEnd As Long
        lngStart = Editor1.Selection.StartPos
        lngEnd = Editor1.Selection.EndPos
        If lngStart <> lngEnd Then
            If Me.Editor1.Range(lngStart, lngEnd).Font.Protected = False And Me.Editor1.Range(lngStart, lngEnd).Font.Hidden = False Then
                Editor1.TOM.TextDocument.Range(lngStart, lngEnd) = ""
            End If
        Else
            If Me.Editor1.Range(lngStart - 1, lngStart).Font.Protected = False And Me.Editor1.Range(lngStart - 1, lngStart).Font.Hidden = False Then
                Editor1.TOM.TextDocument.Range(lngStart - 1, lngStart) = ""
            End If
        End If
    Case vbKeyEscape
        SendKeys "{F11}"
    End Select
End Sub

Private Sub mfrmMultiDocView_RequestModifyDoc(ByVal lngFileID As Long)
'����༭ָ���ļ�
    Dim strPrivs As String
    If Editor1.Modified Then
        Dim r As Long
        r = MsgBox("��ǰ�ļ��Ѿ����޸ģ��Ƿ��ȱ��棿", vbYesNoCancel + vbQuestion, gstrSysName)
        If r = vbCancel Then
            Exit Sub
        ElseIf r = vbYes Then
            If SaveEMRDoc = False Then Exit Sub
        ElseIf r = vbNo Then
            '
        End If
    End If
    '���³�ʼ��Doc����
    Me.Editor1.ReadOnly = False
    Me.Document.ClearAllIDs
    Me.Document.InitEPRDoc cprEM_�޸�, Me.Document.EditType, lngFileID, _
        Me.Document.EPRPatiRecInfo.��������, Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID, _
        Me.Document.EPRPatiRecInfo.Ӥ��, Me.Document.EPRPatiRecInfo.����ID
    Me.Document.OpenEPRDoc Me.Editor1
    Me.Editor1.Modified = False
    Me.RefreshPatiInfo
    If Me.Document.EditType = cprET_�������༭ Then
        '�޸�
        Me.Editor1.ReadOnly = Not (Me.Document.EPRPatiRecInfo.���汾 = 1 And Me.Document.EPRPatiRecInfo.ǩ������ = cprSL_�հ�)
        If Not Me.Editor1.ReadOnly Then
            Select Case Me.Document.EPRFileInfo.����
            Case cprסԺ����
                strPrivs = GetPrivFunc(glngSys, 1251)
                Me.Editor1.ReadOnly = Not (Me.Document.EPRPatiRecInfo.������ = gstrUserName Or InStr(1, strPrivs, "���˲���") > 0)
            Case cpr������
                strPrivs = GetPrivFunc(glngSys, 1255)
                Me.Editor1.ReadOnly = Not (Me.Document.EPRPatiRecInfo.������ = gstrUserName Or InStr(1, strPrivs, "���˻�����") > 0)
            End Select
        End If
    Else
        '����
        Me.Editor1.ReadOnly = (Me.Document.EPRPatiRecInfo.���汾 = 1 And Me.Document.EPRPatiRecInfo.ǩ������ = cprSL_�հ�)
    End If
    Me.ShowMe Me.Document.mfrmParent, False
End Sub

Private Sub ucPacsImgCanvas1_Resize(lngWidth As Long, lngHeight As Long)
    Dim lKey As Long
    lKey = Val(ucPacsImgCanvas1.Tag)
    If lKey > 0 Then
        Document.Tables("K" & lKey).Refresh Editor1, ucPacsImgCanvas1.FinalPic, True
    End If
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, bFinded As Boolean, bNeeded As Boolean
    Dim lW As Long
    bFinded = FindKey(Editor1, "T", tblThis.Tag, lSS, lSE, lES, lEE, bNeeded)
    If bFinded Then
        Editor1.InProcessing = True
        Editor1.Range(lSE, lES).Selected
        Editor1.InProcessing = False
    End If
    Editor1.ResizeUIInterface lngWidth, lngHeight
End Sub

Private Sub ucPacsImgCanvas1_SelectedMarkedPic(lLeft As Long, lTOp As Long, lWidth As Long, lHeight As Long)
    'ѡ�б��ͼ������ʾͼƬ�༭��
    If ucPacsImgCanvas1.Visible = False Then Exit Sub
    Dim lKey As Long
    lKey = Val(ucPacsImgCanvas1.Tag)
    If lKey > 0 Then
        ucPacsImgCanvas1.SavePictures
        If Me.Document.Tables("K" & lKey).Pictures.Count > 0 Then
            ucPictureEditor1.ShowMe Me, ucPacsImgCanvas1.hwnd, cbrThis, _
                Me.Document.Tables("K" & lKey).Pictures(1), _
                lLeft, lTOp, lWidth, lHeight, True, Me.Document.Tables("K" & lKey)
        End If
    End If
End Sub

Private Sub ucPacsImgCanvas1_SelectedPacsPic()
    '������ͼ
    If Not ucPacsImgCanvas1.mMarkedPicture Is Nothing Then
        If ucPictureEditor1.Visible Then
            ucPictureEditor1.CloseMe ucPacsImgCanvas1.mMarkedPicture
            ucPacsImgCanvas1.LayoutPictures False
        End If
    ElseIf ucPacsImgCanvas1.Visible And ucPictureEditor1.Visible Then
        ucPictureEditor1.Visible = False
    End If
End Sub

Private Sub ucPictureEditor1_DblClick()
Dim objPic As StdPicture, strPic As String, lKey As String
'�༭��ʽͼ������
    If ucPictureEditor1.mcPicture.PictureType <> EPRFormulaPicture Then Exit Sub
    
    strPic = ucPictureEditor1.mcPicture.�����ı�
    lKey = ucPictureEditor1.mcPicture.Key
    ucPictureEditor1.Visible = False
    ucPictureEditor1.CloseMe
    Editor1.CloseUIInterface
    Call Editor1.ShowInsertSymbolDlg(False, IIf(InStr(mstrSex, "��") > 0, 1, IIf(InStr(mstrSex, "Ů") > 0, 2, 0)), False, strPic, objPic)
    If objPic Is Nothing Then Exit Sub
    
    Editor1.Tag = "������ű༭"
    Call Document.Pictures("K" & lKey).DeleteFromEditor(Editor1)
    Call Document.Pictures.Remove("K" & lKey)
    InsertPicture EPRFormulaPicture, objPic, objPic.Width, objPic.Height, strPic
    Editor1.Tag = ""

End Sub
Private Sub ExportXML()
'������XML�ļ�
Dim strF As String, i As Integer
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
Dim bFinded As Boolean
    Select Case Me.Document.EditType
    Case cprET_�����ļ�����
        dlgThis.Filename = "����_" & Me.Document.EPRFileInfo.���� & ".xml"
    Case cprET_ȫ��ʾ���༭
        dlgThis.Filename = "����_" & Me.Document.EPRFileInfo.���� & "_" & Me.Document.EPRDemoInfo.���� & ".xml"
    Case cprET_�������༭, cprET_���������
        dlgThis.Filename = "��¼_" & Me.Document.EPRFileInfo.���� & "(" & Me.Document.EPRPatiRecInfo.ID & "," & Me.Document.Ŀ��汾 & ").xml"
    End Select

    dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
    dlgThis.CancelError = True
    On Error GoTo LL
    dlgThis.ShowSave
    strF = dlgThis.Filename
    If gobjFSO.FileExists(strF) Then
        If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
    End If

    Editor1.Freeze
    Editor1.ForceEdit = True
    Editor1.Tag = "cbrThis_ExeCute"
    '�ڵ�ǰ�ؼ��д���ͼƬ�ͱ����滻
    Me.Document.PreSavingRTFText Me.Editor1

    If Me.Document.ExportToXMLFile(Me.Editor1, strF) Then
        MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
    End If
    '�ָ�ͼƬ�ͱ��
    Dim ParaFmt As New cParaFormat, FontFmt As New cFontFormat
    For i = 1 To Me.Document.Pictures.Count
        bFinded = FindKey(Editor1, "P", Me.Document.Pictures(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            '��ԭͼƬ
            Set ParaFmt = Editor1.Range(lKSE, lKES).Para.GetParaFmt
            Set FontFmt = Editor1.Range(lKSE, lKES).Font.GetFontFmt

            Me.Document.Pictures(i).�Ƿ��� = False
            Editor1.Range(lKSS, lKEE).Text = ""
            Me.Document.Pictures(i).InsertIntoEditor Editor1, lKSS, True

            Editor1.Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            Editor1.Range(lKSE, lKES).Font.SetFontFmt FontFmt
            Editor1.Range(lKSS, lKEE).Font.Protected = True
        End If
    Next
    For i = 1 To Me.Document.Tables.Count
        bFinded = FindKey(Editor1, "T", Me.Document.Tables(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            '��ԭ���
            Set ParaFmt = Editor1.Range(lKSE, lKES).Para.GetParaFmt
            Set FontFmt = Editor1.Range(lKSE, lKES).Font.GetFontFmt

            Me.Document.Tables(i).�Ƿ��� = False
            Editor1.Range(lKSS, lKEE).Text = ""
            Me.Document.Tables(i).InsertIntoEditor Editor1, lKSS, , , True

            Editor1.Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            Editor1.Range(lKSE, lKES).Font.SetFontFmt FontFmt
            Editor1.Range(lKSS, lKEE).Font.Protected = True
        End If
    Next

    Editor1.ForceEdit = False
    Editor1.UnFreeze
    Editor1.Tag = ""
LL:
End Sub
Public Function CommBar(ByVal BarId As Long) As XtremeCommandBars.CommandBar
    For Each CommBar In cbrThis
        If CommBar.BarId = BarId Then Exit Function
    Next
End Function
Private Sub ExecuteUnderLine(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal blnForce As Boolean)
    If Editor1.Selection.Font.Protected Or Editor1.Selection.Font.Hidden Then Exit Sub
    
    Dim BarUnderLine As CommandBarPopup, objControl As CommandBarControl
    Call AddUndoPoint  '�ֶ�����
    Editor1.ForceEdit = True
    Editor1.Tag = "cbrThis_ExeCute"
    Set BarUnderLine = CommBar(ID_BAR_FORMAT).FindControl(xtpControlSplitButtonPopup, ID_FORMAT_UNDERLINE)
    
    
    Select Case Control.ID
        Case ID_FORMAT_UNDERLINE
            If BarUnderLine.Checked Then
                Editor1.Selection.Font.Underline = cprNone
            Else
                Select Case True
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_THIN).Checked
                    Editor1.Selection.Font.Underline = cprHair
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_THICK).Checked
                    Editor1.Selection.Font.Underline = cprThick
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_WAVE).Checked
                    Editor1.Selection.Font.Underline = cprWave
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DOT).Checked
                    Editor1.Selection.Font.Underline = cprDotted
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DASH).Checked
                    Editor1.Selection.Font.Underline = cprDash
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DASHDOT).Checked
                    Editor1.Selection.Font.Underline = cprDashDot
                Case BarUnderLine.CommandBar.FindControl(, ID_FORMAT_UNDERLINE_DASHDOT2).Checked
                    Editor1.Selection.Font.Underline = cprDashDotDot
                Case Else
                    Editor1.Selection.Font.Underline = cprHair
                End Select
            End If
        Case ID_FORMAT_UNDERLINE_NONE
                Editor1.Selection.Font.Underline = cprNone
        Case ID_FORMAT_UNDERLINE_THIN
                Editor1.Selection.Font.Underline = cprHair
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_THICK
                Editor1.Selection.Font.Underline = cprThick
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_WAVE
                Editor1.Selection.Font.Underline = cprWave
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DOT
                Editor1.Selection.Font.Underline = cprDotted
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DASH
                Editor1.Selection.Font.Underline = cprDash
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DASHDOT
                Editor1.Selection.Font.Underline = cprDashDot
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
        Case ID_FORMAT_UNDERLINE_DASHDOT2
                Editor1.Selection.Font.Underline = cprDashDotDot
                For Each objControl In BarUnderLine.CommandBar.Controls
                    If objControl.ID = Control.ID Then
                        objControl.Checked = True
                    Else
                        objControl.Checked = False
                    End If
                Next
    End Select
    
    Me.Editor1.ForceEdit = blnForce
    Editor1.Tag = ""
    Call ClearNoUseUndoList
    Call RecountPage
End Sub
Private Sub ExecuteLineSpace(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal blnForce As Boolean)
    Dim BarSpace As CommandBarPopup, objControl As CommandBarControl, BarLineSpace As CommandBarPopup, BarLineSpace1 As CommandBarPopup
    On Error GoTo errHand
    Call AddUndoPoint  '�ֶ�����
    Editor1.ForceEdit = True
    Editor1.Tag = "cbrThis_Execute"
    Set BarSpace = Me.cbrThis.FindControl(, ID_Main_FORMAT)
    If Not BarSpace Is Nothing Then Set BarLineSpace1 = BarSpace.CommandBar.FindControl(, ID_FORMAT_SPACE).CommandBar.FindControl(, ID_FORMAT_LINESPACE)
    Set BarLineSpace = CommBar(ID_BAR_FORMAT).FindControl(xtpControlSplitButtonPopup, ID_FORMAT_LINESPACE)
    If BarLineSpace Is Nothing Then Exit Sub
    Select Case Control.ID
        Case ID_FORMAT_LINESPACE
            If BarLineSpace.Checked Then
                Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 0
            Else
                Select Case True
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE1).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1#
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE2).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.3
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE3).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.5
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE4).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2#
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE5).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2.5
                Case BarLineSpace.CommandBar.FindControl(, ID_FORMAT_LINESPACE6).Checked
                    Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 3#
                End Select
            End If
        Case ID_FORMAT_LINESPACE1
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1
        Case ID_FORMAT_LINESPACE2
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.3
        Case ID_FORMAT_LINESPACE3
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 1.5
        Case ID_FORMAT_LINESPACE4
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2
        Case ID_FORMAT_LINESPACE5
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 2.5
        Case ID_FORMAT_LINESPACE6
            Editor1.Selection.Para.SetLineSpacing cprLSMultiple, 3
        Case ID_FORMAT_LINESPACE7
            Editor1.ShowParaDlg True
    End Select
    If Control.ID <> ID_FORMAT_LINESPACE Then
        Call CheckMenu(Control.ID, BarLineSpace)  '��������ѡ��
        Call CheckMenu(Control.ID, BarLineSpace1) '�˵�ѡ��
    End If
    Me.Editor1.ForceEdit = blnForce
    Editor1.Tag = ""
    Call ClearNoUseUndoList
    Call RecountPage
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function CheckMenu(ByVal ID As Long, ByVal obj As CommandBarPopup)
    Dim objControl As CommandBarControl
    For Each objControl In obj.CommandBar.Controls
        If objControl.ID = ID Then
            objControl.Checked = True
        Else
            objControl.Checked = False
        End If
    Next
End Function

Private Sub SpicalCopy(ByVal blnEnabled As Boolean, ByVal blnVisible As Boolean)
    '-----------------------
    'ר�ø���
    On Error GoTo errHand
    Dim frm As New frmContentCopy
    Dim blnCan As Boolean
    blnCan = frm.ShowMe(Me, Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID, Me.Document.EPRPatiRecInfo.������Դ)
    If blnCan Then
        If blnEnabled And blnVisible Then '��ݼ�ִ��ʱ��Ҫ�ж�
            Call ExecPaste(Me.Editor1)   'ճ�����ݣ������ؼ��֣�
            Call RecountPage
        End If
    End If
    Clipboard.Clear
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function RelateFeedback(ByVal isRelated As Boolean) As Boolean
'���ܣ���Ⱦ�����濨���������Խ��������������ȡ������
'������isRelated  true-������false-ȡ������
    Dim objDisease As Object
On Error GoTo errHand
    If Me.Document.EPRPatiRecInfo.�������� <> cpr������� Or mbln���޴��� Then Exit Function
    Set objDisease = CreateObject("zl9Disease.cDockDisease")
    If objDisease Is Nothing Then Exit Function
    Call objDisease.InitDockDisease(glngSys, gcnOracle)
    Call objDisease.RelateFeedback(Me, Me.Document.EPRPatiRecInfo.ID, Me.Document.EPRPatiRecInfo.����ID, Me.Document.EPRPatiRecInfo.��ҳID, Me.Document.EPRPatiRecInfo.������Դ, isRelated)
    Set objDisease = Nothing
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
