VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditAssistant 
   AutoRedraw      =   -1  'True
   Caption         =   "�ʾ�ѡ��"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10950
   Icon            =   "frmEditAssistant.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDef 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3780
      Left            =   2955
      Picture         =   "frmEditAssistant.frx":058A
      ScaleHeight     =   3780
      ScaleWidth      =   5790
      TabIndex        =   17
      Top             =   1455
      Visible         =   0   'False
      Width           =   5790
      Begin VB.CommandButton cmdAdd 
         Height          =   270
         Left            =   5280
         Picture         =   "frmEditAssistant.frx":47BAC
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "�����ֶ�(ALT+A)"
         Top             =   2745
         Width           =   270
      End
      Begin VB.ComboBox cbo�ֶ� 
         Height          =   300
         Left            =   1125
         TabIndex        =   24
         Top             =   2730
         Width           =   4125
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "���(&K)"
         Height          =   350
         Left            =   2175
         TabIndex        =   23
         Top             =   3285
         Width           =   1100
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4470
         TabIndex        =   22
         Top             =   3285
         Width           =   1100
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3375
         TabIndex        =   21
         Top             =   3285
         Width           =   1100
      End
      Begin VB.TextBox txtAdvice 
         Height          =   1125
         Left            =   1125
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1575
         Width           =   4440
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1245
         Width           =   4440
      End
      Begin VB.PictureBox picTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1545
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   18
         Top             =   90
         Width           =   240
         Begin VB.Image imgTip 
            Height          =   240
            Left            =   0
            Picture         =   "frmEditAssistant.frx":47C76
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Label lblDefTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���������Զ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   105
         Width           =   1365
      End
      Begin VB.Image imgClose 
         Height          =   285
         Left            =   5445
         Picture         =   "frmEditAssistant.frx":4E4C8
         Stretch         =   -1  'True
         Top             =   45
         Width           =   270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   30
         X2              =   6090
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   -60
         X2              =   6000
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   105
         X2              =   6165
         Y1              =   3150
         Y2              =   3150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   15
         X2              =   6075
         Y1              =   3165
         Y2              =   3165
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditAssistant.frx":4E932
         Height          =   645
         Left            =   450
         TabIndex        =   29
         Top             =   495
         Width           =   5040
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         Height          =   180
         Left            =   330
         TabIndex        =   28
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label lbl���ݸ�ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ݸ�ʽ"
         Height          =   180
         Left            =   330
         TabIndex        =   27
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl�ֶ���Ŀ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ֶ���Ŀ"
         Height          =   180
         Left            =   330
         TabIndex        =   26
         Top             =   2790
         Width           =   1455
      End
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   3465
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   3465
      MousePointer    =   7  'Size N S
      TabIndex        =   12
      Top             =   3150
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   3375
      TabIndex        =   11
      Top             =   2865
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   4125
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   10950
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6525
      Width           =   10950
      Begin VB.CommandButton cmdDef 
         Caption         =   "���������Զ���(&F)"
         Height          =   350
         Left            =   6030
         TabIndex        =   31
         Top             =   150
         Width           =   1740
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "��λ(&L)"
         Height          =   350
         Left            =   2715
         TabIndex        =   7
         Top             =   135
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   350
         Left            =   870
         TabIndex        =   6
         Top             =   135
         Width           =   1845
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9375
         TabIndex        =   9
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   8160
         TabIndex        =   8
         Top             =   135
         Width           =   1100
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "�������������"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   3975
         TabIndex        =   16
         Top             =   210
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�ʾ����"
         Height          =   180
         Left            =   90
         TabIndex        =   15
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.Frame fraUD 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3465
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   3765
      Width           =   5475
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϸ����"
         Height          =   180
         Left            =   105
         TabIndex        =   14
         Top             =   30
         Width           =   720
      End
   End
   Begin RichTextLib.RichTextBox rtfSentence 
      Height          =   1245
      Left            =   3540
      TabIndex        =   2
      Top             =   4680
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   2196
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmEditAssistant.frx":4E9DD
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   3285
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   2400
      Left            =   3390
      TabIndex        =   1
      Top             =   225
      Width           =   6315
      _cx             =   11139
      _cy             =   4233
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmEditAssistant.frx":4EA7A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgList 
         Left            =   420
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":4EAEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":4F089
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEditAssistant.frx":4F623
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2310
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
            Picture         =   "frmEditAssistant.frx":4FBBD
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditAssistant.frx":50157
            Key             =   "Expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5865
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   10345
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Line lin 
      Index           =   0
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line lin 
      Index           =   1
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line lin 
      Index           =   2
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line lin 
      Index           =   3
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3045
      Y2              =   3045
   End
   Begin VB.Line lin 
      Index           =   4
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Line lin 
      Index           =   5
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Line lin 
      Index           =   6
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line lin 
      Index           =   7
      Visible         =   0   'False
      X1              =   5445
      X2              =   6120
      Y1              =   3165
      Y2              =   3165
   End
End
Attribute VB_Name = "frmEditAssistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'===============================================================================================
Public mblnShow As Boolean '�ô����Ƿ�������ʾ
Private mstrInput As String
Private mstrSentence As String
Private mstrLike As String
Private mintType As Integer
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintӤ�� As Integer
Private mblnOK As Boolean

Private mlngPreY As Long

Private mrsPati As New ADODB.Recordset
Private mrsFind As New ADODB.Recordset
Private mrsField As ADODB.Recordset
Private mrsFormat As ADODB.Recordset
Private mobjPublicLis As Object
Private mobjXML As Object
Private mstrXmlVersion As String
Private mintIndex As Integer
Private mobjVBA As Object
Private mobjScript As clsScript

Private Type LisItem
    ���鱨��id As String
    ����id As String
    ������־ As Integer
    ������Ŀ As String
    �걾��� As String
    �Ƿ�΢���� As Integer
    ������� As Integer
    ������ As String
    ����� As String
    ���ʱ�� As String
    ����ʱ�� As String
End Type

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, Optional ByVal strInput As String, Optional ByVal intType As Integer = 3) As String
    mstrSentence = ""
    mstrInput = strInput
    mintType = intType
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mintӤ�� = intӤ��
    
    On Error Resume Next
    Me.Show 1, frmParent
    Err.Clear: On Error GoTo 0
    
    If mblnOK Then
        ShowMe = mstrSentence
    Else
        ShowMe = mstrInput
    End If
End Function

Private Function ShowTree() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strMatch As String
    Dim strXMLLIS As String
    Dim objXMLNodeList As Object, objXMLNode As Object
    Dim lngParentID As Long, lngID As Long
    Dim strFirstName As String
    Dim rsItem As New ADODB.Recordset
    Dim rsLisItem As ADODB.Recordset
    Dim L_Item As LisItem
    
    On Error GoTo errH
        
    Screen.MousePointer = 11
    
    strMatch = "f_Sentence_Matched(ID,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10])=1"
        
    '98483:������,2016-11-30,�����Ż�
    strSQL = _
        " Select Max(Level) As ����, a.Id, a.�ϼ�id, a.����, a.����, a.˵��, Max(b.����id) ����id" & vbNewLine & _
        " From �����ʾ���� a," & vbNewLine & _
        "     (Select ����id" & vbNewLine & _
        "       From (Select a.Id, a.����id" & vbNewLine & _
        "              From �����ʾ���� b, �����ʾ�ʾ�� a" & vbNewLine & _
        "              Where a.����id = b.Id And Nvl(Substr(b.��Χ, [1], 1), '0') = '1' And" & vbNewLine & _
        "                    ((Nvl(a.ͨ�ü�, 0) = 0 Or a.ͨ�ü� = 1 And a.����id In (Select a.����id From ������Ա a Where a.��Աid = [11]) Or" & vbNewLine & _
        "                    a.ͨ�ü� = 2 And a.��Աid = [11])))" & vbNewLine & _
        "       Where " & strMatch & vbNewLine & _
        "       Group By ����id) b" & vbNewLine & _
        " Where a.Id = b.����id(+)" & vbNewLine & _
        " Start With a.Id In (b.����id)" & vbNewLine & _
        " Connect By Prior a.�ϼ�id = a.Id" & vbNewLine & _
        " Group By a.Id, a.�ϼ�id, a.����, a.����, a.˵��" & vbNewLine & _
        " Order By ���� Desc, ����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mintType, CStr(NVL(mrsPati!�Ա�)), CStr(NVL(mrsPati!����״��)), _
        CStr(NVL(mrsPati!סԺĿ��)), CStr(NVL(mrsPati!���˲���)), CStr(NVL(mrsPati!��Ժ��ʽ)), "", "", "", "", glngUserId)
    
    '��Ӵʾ����
    tvw_s.Nodes.Clear
    Set objNode = tvw_s.Nodes.Add(, , "_", "���дʾ�", "Close")
    objNode.ExpandedImage = "Expend"
    objNode.Expanded = True
    Do While Not rsTmp.EOF
        Set objNode = tvw_s.Nodes.Add("_" & NVL(rsTmp!�ϼ�ID), tvwChild, "_" & rsTmp!ID, "[" & rsTmp!���� & "]" & rsTmp!����, "Close")
        objNode.Tag = NVL(rsTmp!����id, 0)
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Loop

    'ǿ�����ҽ����ؽ��
    Set objNode = tvw_s.Nodes.Add(, , "=", "����ҽ��", "Close")
    objNode.ExpandedImage = "Expend"
    objNode.Expanded = True
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=1", "��Һ��", "Close")
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=2", "ע����", "Close")
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=4", "�ڷ���", "Close")
    Set objNode = tvw_s.Nodes.Add("=", tvwChild, "=0", "������", "Close")
    objNode.ExpandedImage = "Expend"
    '120692:��Ӽ�����Ŀ
    If Not mobjPublicLis Is Nothing Then
        Call Record_Init(rsLisItem, "id," & adBigInt & ",18|parent_id," & adBigInt & ",18|node_name, " & adVarChar & ",50|node_value," & adVarChar & ",4000")
        Set objNode = tvw_s.Nodes.Add(, , "��", "������Ŀ", "Close")
        objNode.ExpandedImage = "Expend"
        objNode.Expanded = True
        strXMLLIS = mobjPublicLis.GetLaboratoryReportList(mlng����ID, mlng��ҳID)
        If strXMLLIS <> "" Then
            If OpenXMLDocument(strXMLLIS) = True Then
                'LIS���ص�XML��Ϣ
'                <���鱨���б�>
'                    <���鱨��id>54603972</���鱨��id>
'                    <����id>7199230</����id>
'                    <������־>0</������־>
'                    <������Ŀ>Ѫ�� Ѫ����23��(����Ѫ)</������Ŀ>
'                    <�걾���>7199229</�걾���>
'                    <�Ƿ�΢����>0</�Ƿ�΢����>
'                    <�������>0</�������>
'                    <������>��Ц��</������>
'                    <�����>�ֽ�</�����>
'                    <���ʱ��>2008/3/15 17:13:15</���ʱ��>
'                    <����ʱ��>2008/3/15 15:29:00</����ʱ��>
'                    <���鱨��id>56511459</���鱨��id>
'                    ����
'                <���鱨���б�>
                Set objXMLNodeList = mobjXML.selectNodes(".//���鱨���б�").Item(0).childNodes
                strFirstName = objXMLNodeList.Item(0).nodename
                lngID = 0
                For Each objXMLNode In objXMLNodeList
                    lngID = lngID + 1
                    If objXMLNode.nodename = strFirstName Then 'ÿ���Լ��鱨��ID������һ����Ŀ
                        lngParentID = lngID
                        rsLisItem.AddNew
                        rsLisItem!ID = lngID
                        rsLisItem!parent_id = 0
                        rsLisItem!node_name = objXMLNode.nodename
                        rsLisItem!node_value = objXMLNode.Text
                        rsLisItem.Update
                        lngID = lngID + 1
                    End If
                    rsLisItem.AddNew
                    rsLisItem!ID = lngID
                    rsLisItem!parent_id = lngParentID
                    rsLisItem!node_name = objXMLNode.nodename
                    rsLisItem!node_value = objXMLNode.Text
                    rsLisItem.Update
                Next
            Else
                MsgBox "LIS�ӿڷ��صļ�����XML��ʽ����ȷ�����ܼ��ؼ�������Ϣ��", vbInformation, gstrSysName
            End If
            Set rsItem = zlDatabase.CopyNewRec(rsLisItem)
            rsLisItem.Filter = "parent_id=0"
            
            lngID = 1
            Do While Not rsLisItem.EOF
                rsItem.Filter = "parent_id=" & rsLisItem!ID
                Do While Not rsItem.EOF
                    Select Case rsItem!node_name & ""
                        Case "���鱨��id"
                            L_Item.���鱨��id = rsItem!node_value & ""
                        Case "����id"
                            L_Item.����id = rsItem!node_value & ""
                        Case "������־"
                            L_Item.������־ = Val(rsItem!node_value & "")
                        Case "������Ŀ"
                            L_Item.������Ŀ = rsItem!node_value & ""
                        Case "�걾���"
                            L_Item.�걾��� = rsItem!node_value & ""
                        Case "�Ƿ�΢����"
                            L_Item.�Ƿ�΢���� = Val(rsItem!node_value & "")
                        Case "�������"
                            L_Item.������� = Val(rsItem!node_value & "")
                        Case "������"
                            L_Item.������ = rsItem!node_value & ""
                        Case "�����"
                            L_Item.����� = rsItem!node_value & ""
                        Case "���ʱ��"
                            L_Item.���ʱ�� = Format(rsItem!node_value & "", "YYYY-MM-DD HH:mm:SS")
                        Case "����ʱ��"
                            L_Item.����ʱ�� = Format(rsItem!node_value & "", "YYYY-MM-DD HH:mm:SS")
                    End Select
                    rsItem.MoveNext
                Loop
                lngID = lngID + 1
                Set objNode = tvw_s.Nodes.Add("��", tvwChild, "��" & L_Item.���鱨��id & "_" & lngID, L_Item.������Ŀ & "[" & L_Item.���ʱ�� & "]", "Close")
                objNode.Tag = L_Item.���鱨��id & "'" & L_Item.����id & "'" & L_Item.������־ & "'" & L_Item.�걾��� & "'" & L_Item.�Ƿ�΢���� & "'" & _
                        L_Item.������� & "'" & L_Item.������ & "'" & L_Item.����� & "'" & L_Item.���ʱ�� & "'" & L_Item.����ʱ��
                objNode.ExpandedImage = "Expend"
                rsLisItem.MoveNext
            Loop
        End If
    End If
    
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Selected = True
    End If
    If Not tvw_s.SelectedItem Is Nothing Then
        tvw_s.SelectedItem.Expanded = True
        tvw_s.SelectedItem.EnsureVisible
    End If
    
    Screen.MousePointer = 0
    ShowTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowList(Optional ByVal lng����id As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim intִ�з��� As Integer
    Dim strSQL As String, i As Long
    Dim strMatch As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If Mid(tvw_s.SelectedItem.Key, 1, 1) = "_" Then
        Call InitVsf(0)
        strMatch = "f_Sentence_Matched(A.ID,[2],[3],[4],[5],[6],[7],[8],[9],[10],[11])=1"
        If lng����id <> 0 Then
            '�����ζ�ȡ����
            strSQL = "Select A.ID,A.���,A.����,A.ͨ�ü�,Trim(B.�����ı�) as �����ı�" & _
                " From �����ʾ���� B,�����ʾ�ʾ�� A" & _
                " Where A.ID=B.�ʾ�ID(+) And B.���д���(+)=1 And A.����ID=[1] And " & strMatch & _
                "   And ((Nvl(A.ͨ�ü�, 0) = 0" & _
                "       Or A.ͨ�ü� = 1 And A.����id In (Select A.����id From ������Ա A Where A.��Աid =[12])" & _
                "       Or A.ͨ�ü� = 2 And A.��Աid =[12])) Order by A.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����id, mintType, CStr(NVL(mrsPati!�Ա�)), CStr(NVL(mrsPati!����״��)), _
                CStr(NVL(mrsPati!סԺĿ��)), CStr(NVL(mrsPati!���˲���)), CStr(NVL(mrsPati!��Ժ��ʽ)), "", "", "", "", glngUserId)
        Else
            '�������ȡ����
            strSQL = "Select A.ID,A.���,A.����,A.ͨ�ü�,LPad(B.���д���,3,'0')||Trim(B.�����ı�) as �����ı�" & _
                " From �����ʾ���� C,�����ʾ���� B,�����ʾ�ʾ�� A" & _
                " Where A.ID=B.�ʾ�ID And Nvl(B.��������,0)=0 And A.����ID=C.ID And Nvl(Substr(C.��Χ, [1], 1), '0') = '1'" & _
                "   And (A.��� Like [1]||'%'" & _
                "       Or A.���� Like " & IIF(mstrLike <> "", "'%'||", "") & "[1]||'%'" & _
                "       Or B.�����ı� Like " & IIF(mstrLike <> "", "'%'||", "") & "[1]||'%')" & _
                "   And ((Nvl(A.ͨ�ü�, 0) = 0" & _
                "       Or A.ͨ�ü� = 1 And A.����id In(Select A.����id From ������Ա A Where A.��Աid = [12])" & _
                "       Or A.ͨ�ü� = 2 And A.��Աid =[12]))"
            
            strSQL = "Select A.ID,A.���,A.����,A.ͨ�ü�,Substr(Min(A.�����ı�),4) as �����ı�" & _
                " From (" & strSQL & ") A Where " & strMatch & " Group by A.ID,A.���,A.����,A.ͨ�ü� Order by A.���"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mstrInput, mintType, CStr(NVL(mrsPati!�Ա�)), CStr(NVL(mrsPati!����״��)), _
                CStr(NVL(mrsPati!סԺĿ��)), CStr(NVL(mrsPati!���˲���)), CStr(NVL(mrsPati!��Ժ��ʽ)), "", "", "", "", glngUserId)
        End If
        vsList.Redraw = flexRDNone
        vsList.Rows = vsList.FixedRows
        If rsTmp Is Nothing Then Screen.MousePointer = 0: Exit Function
        If Not rsTmp.EOF Then
            vsList.Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                vsList.RowData(i) = Val(rsTmp!ID)
                vsList.TextMatrix(i, 1) = NVL(rsTmp!���)
                vsList.TextMatrix(i, 2) = NVL(rsTmp!����)
                vsList.TextMatrix(i, 3) = NVL(rsTmp!�����ı�)
                vsList.Cell(flexcpPicture, i, 0) = imgList.ListImages(NVL(rsTmp!ͨ�ü�, 0) + 1).Picture
                rsTmp.MoveNext
            Next
            vsList.Cell(flexcpPictureAlignment, 1, 0, vsList.Rows - 1, 0) = 4
            vsList.ROW = 1: vsList.COL = 2
        End If
        vsList.Redraw = flexRDDirect
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "=" Then
        Call InitVsf(1)
        '91329:��ȡҽ������:��ҩ��ʽ�ɱ�����ִ�кͲ���ִ��"C.ִ������ IN (1,2)"
        '125170,18-07-24,CL,������ĿĿ¼��ִ�п��ұȲ���ҽ����¼��ִ������׼ȷ
        intִ�з��� = lng����id
        If tvw_s.SelectedItem.Key = "=" Then intִ�з��� = 99
        strSQL = "" & _
            " Select a.Id, a.���id, b.���� ������Ŀ, Decode(Substr('' || Nvl(a.�ܸ�����, 0), 1, 1), '.', 0, '') || a.�ܸ����� as �ܸ�����, Decode(Substr('' || Nvl(a.��������, 0), 1, 1), '.', 0, '') || a.�������� as ��������, b.���㵥λ, a.ҽ������, a.ҽ������, a.����ҽ��, a.��ʼִ��ʱ��, d.���� ��ҩ;��, d.ִ�з���, 1 As ͨ�ü�," & vbNewLine & _
            "       a.ҽ������ || Decode(Substr('' || Nvl(a.��������, 0), 1, 1), '.', 0, '') || a.�������� || b.���㵥λ || c.ҽ������ As �����ı�" & vbNewLine & _
            " From ����ҽ����¼ a, ������ĿĿ¼ b, ����ҽ����¼ c, ������ĿĿ¼ d" & vbNewLine & _
            " Where a.������� In ('5', '6', '7') And a.������Ŀid = b.Id And a.����id = [1] And a.��ҳid = [2] And a.Ӥ�� = [3] And c.������� = 'E' And" & vbNewLine & _
            "      d.ִ�п��� In (1, 2, 3, 4, 6) And Nvl(d.ִ�з���, 0) = [4] And d.Id = c.������Ŀid And a.���id = c.Id And c.�ϴ�ִ��ʱ�� Is Not Null" & vbNewLine & _
            " Order By a.��ʼִ��ʱ�� Desc"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mlng��ҳID, mintӤ��, intִ�з���)
        vsList.Redraw = flexRDNone
        vsList.Rows = vsList.FixedRows
        If rsTmp Is Nothing Then Screen.MousePointer = 0: Exit Function
        If Not rsTmp.EOF Then
            vsList.Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                vsList.RowData(i) = Val(rsTmp!ID)
                vsList.TextMatrix(i, vsList.ColIndex("���ID")) = NVL(rsTmp!���ID)
                vsList.TextMatrix(i, vsList.ColIndex("һ��")) = ""
                vsList.TextMatrix(i, vsList.ColIndex("ҽ������")) = NVL(rsTmp!ҽ������)
                vsList.TextMatrix(i, vsList.ColIndex("��������")) = NVL(rsTmp!��������) & NVL(rsTmp!���㵥λ)
                vsList.TextMatrix(i, vsList.ColIndex("��ҩ;��")) = NVL(rsTmp!��ҩ;��)
                vsList.TextMatrix(i, vsList.ColIndex("�ܸ�����")) = NVL(rsTmp!�ܸ�����)
                vsList.TextMatrix(i, vsList.ColIndex("������Ŀ")) = NVL(rsTmp!������Ŀ)
                vsList.TextMatrix(i, vsList.ColIndex("����ҽ��")) = NVL(rsTmp!����ҽ��)
                vsList.TextMatrix(i, vsList.ColIndex("��ʼʱ��")) = Format(NVL(rsTmp!��ʼִ��ʱ��), "YYYY-MM-DD HH:mm")
                vsList.TextMatrix(i, vsList.ColIndex("ҽ������")) = NVL(rsTmp!ҽ������)
                vsList.TextMatrix(i, vsList.ColIndex("�����ı�")) = NVL(rsTmp!�����ı�)
                vsList.Cell(flexcpPicture, i, 0) = imgList.ListImages(NVL(rsTmp!ͨ�ü�, 0) + 1).Picture
                
                rsTmp.MoveNext
            Next
            vsList.Cell(flexcpPictureAlignment, 1, 0, vsList.Rows - 1, 0) = 4
            vsList.ROW = 1: vsList.COL = 2
        End If
        vsList.Redraw = flexRDDirect
        For i = vsList.FixedRows To vsList.Rows - 1
            If vsList.TextMatrix(i, vsList.ColIndex("һ��")) = "" Then
                Call SetTagһ����ҩ(i)
            End If
        Next
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "��" Then
        '��ȡ������
        Call ShowLisList
    End If
    Screen.MousePointer = 0
    ShowList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cbo���_Click()
    Dim arrField As Variant, i As Long
    
    '1.��鲢���µ�ǰ��������
    '------------------------------
    If Visible Then
        If Not UpdateFormat Then Exit Sub
    End If
    '2.��ʾ���л�������������
    '------------------------------
    mintIndex = cbo���.ListIndex
    
    '��ʾ�����ֶ��б�
    cbo�ֶ�.Clear
    mrsField.Filter = "���=" & cbo���.ItemData(cbo���.ListIndex)
    arrField = Split(mrsField!�ֶ�, ",")
    For i = 0 To UBound(arrField)
        cbo�ֶ�.AddItem arrField(i)
    Next
    
    '��ʾ��ǰ���õ�ҽ������
    mrsFormat.Filter = "���=" & cbo���.ItemData(cbo���.ListIndex)
    If Not mrsFormat.EOF Then
        If Val("" & mrsFormat!�Ƿ��޸�) = 1 Then
            txtAdvice.Text = mrsFormat!�¸�ʽ & ""
        Else
            txtAdvice.Text = mrsFormat!��ʽ & ""
        End If
    Else
        txtAdvice.Text = ""
    End If
    txtAdvice.Tag = ""
End Sub

Private Function UpdateFormat() As Boolean
    Dim strMsg As String
    
    strMsg = CheckFormat(txtAdvice.Text)
    If strMsg <> "" Then
        Call zlControl.CboSetIndex(cbo���.hWnd, mintIndex)
        MsgBox strMsg, vbInformation, gstrSysName
        txtAdvice.SetFocus: Exit Function
    End If
    mrsFormat.Filter = "���=" & cbo���.ItemData(mintIndex)
    If mrsFormat.EOF Then
        If Trim(txtAdvice.Text) <> "" Then 'ԭ��û���ݵ������
            mrsFormat.AddNew
            mrsFormat!��� = cbo���.ItemData(mintIndex)
            mrsFormat!���� = cbo���.List(mintIndex)
            mrsFormat!�¸�ʽ = txtAdvice.Text
            mrsFormat!�Ƿ��޸� = 1
            mrsFormat.Update
        End If
    Else
        If mrsFormat!��ʽ & "" <> txtAdvice.Text Then
            mrsFormat!�¸�ʽ = txtAdvice.Text
            mrsFormat!���� = cbo���.List(mintIndex)
            mrsFormat!�Ƿ��޸� = 1
            mrsFormat.Update
        Else
            If Val(mrsFormat!�Ƿ��޸� & "") = 1 Then
                mrsFormat!�Ƿ��޸� = 0
                mrsFormat.Update
            End If
        End If
    End If
    txtAdvice.Tag = ""
    UpdateFormat = True
End Function

Private Function CheckFormat(ByVal strText As String) As String
'���ܣ����ҽ�������Ƿ���ȷ
'���أ�������Ϣ
'      strPreview=Ԥ��ҽ������Ч��
    Dim intLeft As Integer, intRight As Integer
    Dim strTmp As String, strPar As String
    Dim strMsg As String, i As Long
    Dim objVBA As Object, strEval As String
    Dim objScript As New clsScript
    
    If Trim(strText) = "" And strText = Trim(strText) Then Exit Function
    If zlCommFun.ActualLen(strText) > txtAdvice.MaxLength Then
        strMsg = "��������̫����ֻ���� " & txtAdvice.MaxLength & " ���ַ��� " & txtAdvice.MaxLength \ 2 & " �����֡�"
        GoTo EndLine
    End If
    
    If Not InStr(strText, "[") > 0 Then
        strMsg = "��ʽ����ȷ,����ֶ���Ŀ��"
        GoTo EndLine
    End If
        
    '���������
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) = "[" Then
            intLeft = intLeft + 1
        ElseIf Mid(strText, i, 1) = "]" Then
            intRight = intRight + 1
            If intLeft <> intRight Then
                strMsg = """[""��""]""���Ų���ԡ�"
                GoTo EndLine
            End If
        End If
    Next
    If intLeft = 0 And intRight = 0 Then Exit Function
    If intLeft <> intRight Then
        strMsg = """[""��""]""���Ų���ԡ�"
        GoTo EndLine
    End If
    
    '����ֶ�����
    strTmp = strText
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Trim(Left(strTmp, InStr(strTmp, "]") - 1))
                        
        If strPar = "" Then
            strMsg = """[]""����֮��û����д�ֶ�����"
            GoTo EndLine
        End If
        
        For i = 0 To cbo�ֶ�.ListCount - 1
            If cbo�ֶ�.List(i) = "[" & strPar & "]" Then Exit For
        Next
        If i > cbo�ֶ�.ListCount - 1 Then
            strMsg = "ʹ���˲����ڵ�""[" & strPar & "]""�ֶΡ�"
            GoTo EndLine
        End If
    Loop
    
    'ִ�в���
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    If objVBA Is Nothing Then
        strMsg = "Microsoft Script Control δ��ȷ��װ(msscript.ocx)������ִ�м�顣�����°�װ�ͻ��˳���"
        GoTo EndLine
    End If
    Err.Clear: On Error GoTo 0
    objVBA.Language = "VBScript"
    objVBA.addObject "clsScript", objScript, True
    strEval = Replace(strText, "[", """")
    strEval = Replace(strEval, "]", """")
    On Error Resume Next
    Call objVBA.Eval(strEval)
    If objVBA.Error.Number <> 0 Then
        strMsg = objVBA.Error.Description
        objVBA.Error.Clear
    End If
EndLine:
    CheckFormat = strMsg
End Function

Private Sub cbo�ֶ�_GotFocus()
    Call zlControl.TxtSelAll(cbo�ֶ�)
End Sub

Private Sub cbo�ֶ�_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    If cbo�ֶ�.Text = "" Then Exit Sub
    txtAdvice.SelText = cbo�ֶ�.Text
    cbo�ֶ�.SetFocus
End Sub

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim strMsg As String
    
    If Trim(txtAdvice.Text) <> "" Then
        strMsg = CheckFormat(txtAdvice.Text)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            txtAdvice.SetFocus
        Else
            MsgBox "���ݸ�ʽ��д��ȷ��", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Dim blnCancel As Boolean
    If Not mrsFormat Is Nothing Then
        mrsFormat.Filter = "�Ƿ��޸�=1"
        blnCancel = mrsFormat.RecordCount > 0
        If blnCancel = False Then
            mrsFormat.Filter = "���=" & cbo���.ItemData(cbo���.ListIndex)
            If Not mrsFormat.EOF Then
                blnCancel = (mrsFormat!��ʽ & "" <> txtAdvice.Text)
            Else
                blnCancel = txtAdvice.Text <> ""
            End If
        End If
    End If
    If blnCancel = True Then
        If MsgBox("����˳����ᶪʧ�����ı�����ݣ�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            '�ָ�֮ǰ���޸�
            If Not mrsFormat Is Nothing Then
                mrsFormat.Filter = "�Ƿ��޸�=1"
                Do While Not mrsFormat.EOF
                    mrsFormat!�Ƿ��޸� = 0
                    mrsFormat!�¸�ʽ = ""
                    mrsFormat.Update
                mrsFormat.MoveNext
                Loop
            End If
            txtAdvice.Text = ""
        End If
    End If
     picDef.Visible = False
     SetControlEnable True
End Sub

Private Sub cmdDef_Click()
    With picDef
        .Left = (Me.ScaleWidth - .Width) \ 2
        .Top = (Me.ScaleHeight - .Height) \ 2
        .Visible = True
        .ZOrder 0
    End With
    SetControlEnable False
End Sub

Private Sub cmdFind_Click()
'����:�ʾ����
    Dim strText As String, strMatch As String
    Dim strFind As String, strSQL As String
    Dim lngRow As Long, lngTypeID As Long
    
    On Error GoTo ErrHand
    
    If mrsFind.State = adStateOpen Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocaItem
        Exit Sub
    End If
    
    If Trim(txtFind.Text) = "" Then
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txtFind.Text, "'") > 0 Then
        MsgBox "��������ݰ����Ƿ��ַ� ' ,����!", vbInformation, gstrSysName
        If txtFind.Enabled And txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    If Not tvw_s.SelectedItem Is Nothing Then
        lngTypeID = Val(tvw_s.SelectedItem.Tag)
    Else
        lngTypeID = 0
    End If
    
    strText = mstrLike & txtFind.Text & "%"
    If zlCommFun.IsCharChinese(txtFind.Text) Then
        strFind = " And A.���� Like '" & strText & "'"
    ElseIf IsNumeric(txtFind.Text) Then
        strFind = " And A.��� Like '" & strText & "'"
    Else
        strFind = " And zlspellcode(A.����) Like '" & UCase(strText) & "'"
    End If
    
    '���������������ȡƥ��Ĵʾ�
    strMatch = " f_Sentence_Matched(A.ID,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10])=1 "
    strSQL = "   Select A.ID,A.����ID,A.���,A.���� From �����ʾ���� B, �����ʾ�ʾ�� A" & _
        "   Where A.����id = B.ID And Nvl(Substr(B.��Χ, [1], 1), '0') = '1' And " & strMatch & _
        "   And ((Nvl(A.ͨ�ü�, 0) = 0" & _
        "       Or A.ͨ�ü� = 1 And A.����id In (Select A.����id From ������Ա A, �ϻ���Ա�� B Where A.��Աid = B.��Աid And B.�û��� = User)" & _
        "       Or A.ͨ�ü� = 2 And A.��Աid In (Select ��Աid From �ϻ���Ա�� Where �û��� = User)))" & strFind & _
        "   Order by " & IIF(lngTypeID = 0, "", " DECODE(A.����ID," & lngTypeID & ",0,1),") & "A.����ID,A.���"
    Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mintType, CStr(NVL(mrsPati!�Ա�)), CStr(NVL(mrsPati!����״��)), _
        CStr(NVL(mrsPati!סԺĿ��)), CStr(NVL(mrsPati!���˲���)), CStr(NVL(mrsPati!��Ժ��ʽ)), "", "", "", "")

    Call LocaItem
        
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub LocaItem()
    Dim lngRow As Long
    
    If mrsFind.RecordCount = 0 Then
        lblInfo.Caption = "û���ҵ�������������Ϣ"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    
    If mrsFind.EOF = True Then
        lblInfo.Caption = "�Ѿ�������ж�λ����������������"
        lblInfo.ForeColor = &HFF&
        Exit Sub
    End If
    lblInfo.Caption = "���ҵ�" & mrsFind.RecordCount & "��,��ǰ�ǵ�" & mrsFind.AbsolutePosition & "��"
    lblInfo.ForeColor = &H8000000D
    
    If mrsFind.RecordCount > 0 Then
        If mrsFind.RecordCount <> mrsFind.AbsolutePosition Then
            cmdFind.Caption = "��һ��(&L)"
        Else
            cmdFind.Caption = "��λ(&L)"
            lblInfo.Caption = "�Ѿ������һ������������������"
        End If
    End If
    
    '��ʼ���ж�λ
    tvw_s.Nodes("_" & mrsFind!����id).Selected = True
    tvw_s.SelectedItem.EnsureVisible
    Call ShowList(mrsFind!����id)
    
    For lngRow = vsList.FixedRows To vsList.Rows - 1
        If Val(vsList.RowData(lngRow)) = Val(mrsFind!ID) Then
            vsList.ROW = lngRow
            vsList.TopRow = lngRow
            Exit For
        End If
    Next lngRow
End Sub

Private Sub cmdGO_Click()
    Dim blnTrans As Boolean
    Dim rsTemp As ADODB.Recordset
    If Not UpdateFormat Then
        txtAdvice.SetFocus: Exit Sub
    End If
    On Error GoTo ErrHand
    mrsFormat.Filter = 0
    gcnOracle.BeginTrans: blnTrans = True
    With mrsFormat
        Do While Not .EOF
            If Val(!�Ƿ��޸� & "") = 1 Then
                gstrSQL = "Zl_�������ݵ��붨��_Update(" & !��� & ",'" & !���� & "','" & Replace(!�¸�ʽ, "'", "''") & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        .MoveNext
        Loop
    End With
    gcnOracle.CommitTrans: blnTrans = False
    '��������ԭ�м�¼��Ϣ
    mrsFormat.Filter = 0
    Set rsTemp = zlDatabase.CopyNewRec(mrsFormat)
    rsTemp.Filter = 0
    Do While Not rsTemp.EOF
        If Val(rsTemp!�Ƿ��޸� & "") = 1 Then
            mrsFormat.Filter = "���=" & rsTemp!���
            mrsFormat!��ʽ = rsTemp!�¸�ʽ & ""
            mrsFormat!�Ƿ��޸� = 0
            mrsFormat!�¸�ʽ = ""
            mrsFormat.Update
        End If
        rsTemp.MoveNext
    Loop
    picDef.Visible = False
    SetControlEnable True
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    If rtfSentence.Text = "" Then
        MsgBox "û�п��õĴʾ����ݡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrSentence = Replace(Replace(rtfSentence.Text, "|", "�O"), "'", "")
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    ElseIf KeyCode = vbKeyF3 Then
        If cmdFind.Enabled And cmdFind.Visible Then Call cmdFind_Click
    ElseIf KeyCode = vbKeyA And Shift = vbAltMask And picDef.Visible = True Then
        Call cmdGO_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim strSQL As String, i As Long
    Dim vRect As RECT, lngMaxH As Long
    Dim rsTemp As New ADODB.Recordset
    
    mblnShow = True
    mblnOK = False
    mstrSentence = ""
    Me.rtfSentence.Text = mstrInput
    
    On Error GoTo errH
    If mobjPublicLis Is Nothing Then
        On Error Resume Next
        Set mobjPublicLis = CreateObject("zlPublicLIS.clsSampleReprot")
        Err.Clear: On Error GoTo 0
        If Not mobjPublicLis Is Nothing Then
            Call mobjPublicLis.InitSampleReprot(gcnOracle, glngSys, 1265, "")
        End If
    End If
    If mobjPublicLis Is Nothing Then
        MsgBox "LIS��������zlPublicLIS����ʧ�ܣ������ܲ鿴�����������", vbInformation, gstrSysName
    End If
    '���������Զ�������
    '��ȡ�Զ���ʽ
    gstrSQL = "Select ���,����,��ʽ from �������ݵ��붨��"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "�������ݵ��붨��")
    Set mrsFormat = zlDatabase.CopyNewRec(rsTemp, , , Array("�Ƿ��޸�", adInteger, 1, 0, "�¸�ʽ", adVarChar, 500, Empty))
    
    txtAdvice.Tag = ""
    Call Record_Init(mrsField, "���," & adInteger & ",1|�ֶ�," & adVarChar & ",2000")
    mrsField.AddNew: mrsField!��� = 1: mrsField!�ֶ� = "[���],[����],[�����ı�]" '�����ʾ�
    mrsField.AddNew: mrsField!��� = 2: mrsField!�ֶ� = "[��ʼʱ��],[����ҽ��],[ҽ������],[������Ŀ],[����],[����],[ҽ������],[��ҩ;��]" 'ҽ������
    mrsField.AddNew: mrsField!��� = 3: mrsField!�ֶ� = "[ָ�����],[ָ��������],[������],[��λ],[�����־],[����ο�]" '������Ŀ(��ͨ��Ŀ)
    mrsField.AddNew: mrsField!��� = 4: mrsField!�ֶ� = "[ϸ����],[��ҩ����],[������],[�����ؽ��],[��ҩ��],[ҩ������],[�÷�����1],[�÷�����2],[ѪҩŨ��1],[ѪҩŨ��2],[��ҩŨ��1],[��ҩŨ��2]" '������Ŀ(΢������Ŀ)
    mrsField.UpdateBatch
    With cbo���
        .Clear
        .AddItem "����ʾ�": .ItemData(.NewIndex) = 1
        .AddItem "ҽ������": .ItemData(.NewIndex) = 2
        .AddItem "������Ŀ[��ͨ��Ŀ]": .ItemData(.NewIndex) = 3
        .AddItem "������Ŀ[΢������Ŀ]": .ItemData(.NewIndex) = 4
        .ListIndex = 0
    End With
    mintIndex = cbo���.ListIndex
   
    
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    gstrSQL = "Select B.��ҳID as ����ID,NVL(B.�Ա�,A.�Ա�) �Ա�,Nvl(B.����״��,A.����״��) as ����״��," & _
        " B.סԺĿ��,B.��ǰ���� as ���˲���,B.��Ժ��ʽ" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2]"
    Set mrsPati = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mlng����ID, mlng��ҳID)
    '��ȡ�ʾ�����
    Call ShowTree
    
    '������ʾ����
    Call RestoreWinState(Me, App.ProductName, IIF(mstrInput <> "", 1, 0))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    tvw_s.Left = 0
    tvw_s.Top = 0
    tvw_s.Height = Me.ScaleHeight - picBottom.Height
    
    fraLR.Left = tvw_s.Left + tvw_s.Width
    fraLR.Top = 0
    fraLR.Height = tvw_s.Height
    
    vsList.Top = 0
    vsList.Left = fraLR.Left + fraLR.Width
    vsList.Height = Me.ScaleHeight - rtfSentence.Height - fraUD.Height - picBottom.Height
    vsList.Width = Me.ScaleWidth - fraLR.Width - tvw_s.Width
    
    fraUD.Top = vsList.Top + vsList.Height
    fraUD.Left = vsList.Left
    fraUD.Width = vsList.Width
    
    rtfSentence.Top = fraUD.Top + fraUD.Height
    rtfSentence.Left = vsList.Left
    rtfSentence.Width = vsList.Width
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    If Not mrsPati Is Nothing Then
        If mrsPati.State = adStateOpen Then mrsPati.Close
        Set mrsPati = Nothing
    End If
    If Not mrsFind Is Nothing Then
        If mrsFind.State = adStateOpen Then mrsFind.Close
        Set mrsFind = Nothing
    End If
    If Not mrsField Is Nothing Then
        If mrsField.State = adStateOpen Then mrsField.Close
        Set mrsField = Nothing
    End If
    If Not mrsFormat Is Nothing Then
        If mrsFormat.State = adStateOpen Then mrsFormat.Close
        Set mrsFormat = Nothing
    End If
    Set mobjPublicLis = Nothing
    Set mobjXML = Nothing
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Call SaveWinState(Me, App.ProductName, IIF(mstrInput <> "", 1, 0))
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 1 Then
            If Me.Width + X < 4000 Or Me.Width + X > 9600 Then Exit Sub
            Me.Width = Me.Width + X
        ElseIf Index = 2 Then
            If Me.Height + Y < rtfSentence.Height * 2 Or Me.Height + Y > 7200 Then Exit Sub
            Me.Height = Me.Height + Y
        End If
        Call Form_Resize
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsList.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvw_s.Width = tvw_s.Width + X
        
        vsList.Left = vsList.Left + X
        vsList.Width = vsList.Width - X
        
        fraUD.Left = fraUD.Left + X
        fraUD.Width = fraUD.Width - X
        
        rtfSentence.Left = rtfSentence.Left + X
        rtfSentence.Width = rtfSentence.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub fraUD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlngPreY = Y
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If vsList.Height + (Y - mlngPreY) < 1000 Or rtfSentence.Height - (Y - mlngPreY) < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + (Y - mlngPreY)
        vsList.Height = vsList.Height + (Y - mlngPreY)
        rtfSentence.Top = rtfSentence.Top + (Y - mlngPreY)
        rtfSentence.Height = rtfSentence.Height - (Y - mlngPreY)
        
        Me.Refresh
    End If
End Sub

Private Sub imgClose_Click()
    Call cmdClose_Click
End Sub

Private Sub imgTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picTip_MouseMove(Button, Shift, X, Y)
End Sub
    
Private Sub picBottom_GotFocus()
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    
    If picBottom.ScaleWidth - cmdCancel.Width * 2 < 3500 Then Exit Sub
    cmdCancel.Left = picBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
End Sub

Private Sub picTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    strInfo = "ȱʡ��ʽ����" & vbCrLf & "  ����ʾ䣺[�����ı�]" & vbCrLf & _
        "  ҽ�����ݣ�[ҽ������]+[����]+[��ҩ;��]" & vbCrLf & _
        "  ������Ŀ[��ͨ��Ŀ]��[ָ��������]+""(""+[������]+"")""" & vbCrLf & _
        "  ������Ŀ[΢������Ŀ]��[������]+""(""+[�����ؽ��]+"")"""
    Call zlCommFun.ShowTipInfo(picTip.hWnd, strInfo, True, True)
End Sub


Private Sub rtfSentence_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    End If
End Sub

Private Sub tvw_s_Expand(ByVal Node As MSComctlLib.Node)
    If Node.Children = 1 Then
        Node.Child.Expanded = True
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Val(Mid(Node.Key, 2)) <> 0 Then
        Call ShowList(Val(Mid(Node.Key, 2)))
    Else
        vsList.Rows = vsList.FixedRows
    End If
End Sub

Private Sub txtAdvice_Change()
    txtAdvice.Tag = "1"
End Sub


Private Sub txtFind_Change()
    If Trim(txtFind.Text) = "" Then
        lblInfo.Caption = "�������������"
        lblInfo.ForeColor = &H8000&
    Else
        lblInfo.Caption = "�����λ��ɴʾ����"
        lblInfo.ForeColor = &H8000000D
    End If
    
    cmdFind.Caption = "��λ(&L)"
    Set mrsFind = New ADODB.Recordset
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdFind.SetFocus
        Call cmdFind_Click
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub vsList_DblClick()
    With vsList
        If .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call LoadWords
        End If
    End With
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call vsList_DblClick
    End If
End Sub

Private Sub LoadWords()
    Dim lngStart As Long, lngStart_LAST As Long
    Dim strText As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsValue As New ADODB.Recordset
    Dim bln�Ƿ�΢���� As Boolean, arrTag() As String
    Dim strReturn As String
    On Error GoTo ErrHand
    
    If Val(vsList.RowData(vsList.ROW)) = 0 Then Exit Sub
    lngStart_LAST = rtfSentence.SelStart
    If lngStart_LAST = 0 Then lngStart_LAST = Len(rtfSentence.Text)
    rtfSentence.Tag = rtfSentence.Text
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        Err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.addObject "clsScript", mobjScript, True
        End If
    End If
    
    If Mid(tvw_s.SelectedItem.Key, 1, 1) = "_" Then
        gstrSQL = "Select ��������,�����ı�,Ҫ������,Ҫ�ص�λ From �����ʾ���� Where �ʾ�ID=[1] Order by ���д���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, Val(vsList.RowData(vsList.ROW)))
        
        rtfSentence.Text = ""
        Do While Not rsTemp.EOF
            lngStart = Len(rtfSentence.Text)
            rtfSentence.SelStart = lngStart
            rtfSentence.SelLength = 0
            Select Case rsTemp!��������
            Case 0 '��������
                strText = NVL(rsTemp!�����ı�)
                With rtfSentence
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = False
                End With
            Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                If Not IsNull(rsTemp!�����ı�) Then
                    strText = rsTemp!�����ı�
                Else
                    strText = ""
                    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],[4]) as ���� From Dual"
                    Set rsValue = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, CStr(rsTemp!Ҫ������), mlng����ID, mlng��ҳID, 2)
                    If Not rsTemp.EOF Then strText = IIF(Not IsNull(rsValue!����), rsValue!���� & NVL(rsTemp!Ҫ�ص�λ), "")
                    If strText = "" Then strText = "{" & rsTemp!Ҫ������ & "}" & NVL(rsTemp!Ҫ�ص�λ)
                End If
                With rtfSentence
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = True
                End With
            End Select
            rsTemp.MoveNext
        Loop
        strReturn = ""
        mrsFormat.Filter = "���=1"
        If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!��ʽ & ""
        If strReturn <> "" Then
            If InStr(strReturn, "[���]") > 0 Then
               strReturn = Replace(strReturn, "[���]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("���")) & """")
            End If
            If InStr(strReturn, "[����]") > 0 Then
               strReturn = Replace(strReturn, "[����]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("����")) & """")
            End If
            If InStr(strReturn, "[�����ı�]") > 0 Then
               strReturn = Replace(strReturn, "[�����ı�]", """" & rtfSentence.Text & """")
            End If
            strReturn = mobjVBA.Eval(strReturn)
            rtfSentence.Text = strReturn
        End If
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "=" Then 'ҽ��
        strReturn = ""
        mrsFormat.Filter = "���=2"
        If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!��ʽ & ""
        If strReturn = "" Then
            strReturn = vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�����ı�"))
        Else
            '"[��ʼʱ��],[����ҽ��],[ҽ������],[������Ŀ],[����],[����],[ҽ������],[��ҩ;��]"
            If InStr(strReturn, "[��ʼʱ��]") > 0 Then
               strReturn = Replace(strReturn, "[��ʼʱ��]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��ʼʱ��")) & """")
            End If
            If InStr(strReturn, "[����ҽ��]") > 0 Then
               strReturn = Replace(strReturn, "[����ҽ��]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("����ҽ��")) & """")
            End If
            If InStr(strReturn, "[ҽ������]") > 0 Then
               strReturn = Replace(strReturn, "[ҽ������]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ҽ������")) & """")
            End If
            If InStr(strReturn, "[������Ŀ]") > 0 Then
               strReturn = Replace(strReturn, "[������Ŀ]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("������Ŀ")) & """")
            End If
            If InStr(strReturn, "[����]") > 0 Then
               strReturn = Replace(strReturn, "[����]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��������")) & """")
            End If
            If InStr(strReturn, "[����]") > 0 Then
               strReturn = Replace(strReturn, "[����]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�ܸ�����")) & """")
            End If
            If InStr(strReturn, "[ҽ������]") > 0 Then
               strReturn = Replace(strReturn, "[ҽ������]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ҽ������")) & """")
            End If
            If InStr(strReturn, "[��ҩ;��]") > 0 Then
               strReturn = Replace(strReturn, "[��ҩ;��]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��ҩ;��")) & """")
            End If
            strReturn = mobjVBA.Eval(strReturn)
        End If
        rtfSentence.Text = strReturn
    ElseIf Mid(tvw_s.SelectedItem.Key, 1, 1) = "��" Then
        arrTag = Split(tvw_s.SelectedItem.Tag, "'")
        bln�Ƿ�΢���� = Val(arrTag(4)) = 1
        strReturn = ""
        If bln�Ƿ�΢���� = False Then
            mrsFormat.Filter = "���=3"
            If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!��ʽ & ""
            If strReturn = "" Then 'Ĭ�ϵ���ָ������ƺͽ��
                strReturn = vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ָ��������"))
                If vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ָ��������")) <> "" Then
                    strReturn = strReturn & ":" & "(" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("������")) & ")"
                End If
            Else
                '"[ָ�����],[ָ��������],[������],[�����־],[����ο�]"
                 If InStr(strReturn, "[ָ�����]") > 0 Then
                    strReturn = Replace(strReturn, "[ָ�����]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ָ�����")) & """")
                 End If
                 If InStr(strReturn, "[ָ��������]") > 0 Then
                    strReturn = Replace(strReturn, "[ָ��������]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ָ��������")) & """")
                 End If
                 If InStr(strReturn, "[������]") > 0 Then
                    strReturn = Replace(strReturn, "[������]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("������")) & """")
                 End If
                 If InStr(strReturn, "[��λ]") > 0 Then
                    strReturn = Replace(strReturn, "[��λ]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��λ")) & """")
                 End If
                 If InStr(strReturn, "[�����־]") > 0 Then
                    strReturn = Replace(strReturn, "[�����־]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�����־")) & """")
                 End If
                 If InStr(strReturn, "[����ο�]") > 0 Then
                    strReturn = Replace(strReturn, "[����ο�]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("����ο�")) & """")
                 End If
                 strReturn = mobjVBA.Eval(strReturn)
            End If
            rtfSentence.Text = strReturn
        Else
            If vsList.RowOutlineLevel(vsList.ROW) <= 0 Then Exit Sub '΢������ϸ���в�����ֻ�ܵ������Ŀ
            mrsFormat.Filter = "���=4"
            '"[ϸ����],[��ҩ����],[������],[�����ؽ��],[��ҩ��],[ҩ������],[�÷�����1],[�÷�����2],[ѪҩŨ��1],[ѪҩŨ��2],[��ҩŨ��1],[��ҩŨ��2]"
            If mrsFormat.RecordCount > 0 Then strReturn = mrsFormat!��ʽ & ""
            If strReturn = "" Then 'Ĭ�ϵ���ָ������ƺͽ��
                strReturn = vsList.TextMatrix(vsList.ROW, vsList.ColIndex("������"))
                If vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�����ؽ��")) <> "" Then
                    strReturn = strReturn & ":" & "(" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�����ؽ��")) & ")"
                End If
            Else
                If InStr(strReturn, "[ϸ����]") > 0 Then
                   strReturn = Replace(strReturn, "[ϸ����]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ϸ����")) & """")
                End If
                If InStr(strReturn, "[��ҩ����]") > 0 Then
                   strReturn = Replace(strReturn, "[��ҩ����]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��ҩ����")) & """")
                End If
                If InStr(strReturn, "[������]") > 0 Then
                   strReturn = Replace(strReturn, "[������]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("������")) & """")
                End If
                If InStr(strReturn, "[�����ؽ��]") > 0 Then
                   strReturn = Replace(strReturn, "[�����ؽ��]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�����ؽ��")) & """")
                End If
                If InStr(strReturn, "[��ҩ��]") > 0 Then
                   strReturn = Replace(strReturn, "[��ҩ��]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��ҩ��")) & """")
                End If
                If InStr(strReturn, "[ҩ������]") > 0 Then
                   strReturn = Replace(strReturn, "[ҩ������]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ҩ������")) & """")
                End If
                If InStr(strReturn, "[�÷�����1]") > 0 Then
                   strReturn = Replace(strReturn, "[�÷�����1]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�÷�����1")) & """")
                End If
                If InStr(strReturn, "[�÷�����2]") > 0 Then
                   strReturn = Replace(strReturn, "[�÷�����2]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("�÷�����2")) & """")
                End If
                If InStr(strReturn, "[ѪҩŨ��1]") > 0 Then
                   strReturn = Replace(strReturn, "[ѪҩŨ��1]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ѪҩŨ��1")) & """")
                End If
                If InStr(strReturn, "[ѪҩŨ��2]") > 0 Then
                   strReturn = Replace(strReturn, "[ѪҩŨ��2]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("ѪҩŨ��2")) & """")
                End If
                If InStr(strReturn, "[��ҩŨ��1]") > 0 Then
                   strReturn = Replace(strReturn, "[��ҩŨ��1]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��ҩŨ��1")) & """")
                End If
                If InStr(strReturn, "[��ҩŨ��2]") > 0 Then
                  strReturn = Replace(strReturn, "[��ҩŨ��2]", """" & vsList.TextMatrix(vsList.ROW, vsList.ColIndex("��ҩŨ��2")) & """")
                End If
                strReturn = mobjVBA.Eval(strReturn)
            End If
            rtfSentence.Text = strReturn
        End If
    End If
    
    rtfSentence.Text = Mid(rtfSentence.Tag, 1, lngStart_LAST) & "��" & rtfSentence.Text & Mid(rtfSentence.Tag, lngStart_LAST + 1) & "��"
    If Mid(rtfSentence.Text, 1, 1) = "��" Then rtfSentence.Text = Mid(rtfSentence.Text, 2)
    If Right(rtfSentence.Text, 1) = "��" Then rtfSentence.Text = Mid(rtfSentence.Text, 1, Len(rtfSentence.Text) - 1)
    If lngStart_LAST = Len(rtfSentence.Tag) Then lngStart_LAST = Len(rtfSentence.Text)
    rtfSentence.SelStart = lngStart_LAST
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ShowLisList()
'���ܣ�����ѡ��ļ�����Ŀ��չʾ�����Ϣ
    Dim lngKey As Long '���鱨��id
    Dim strXMLLIS As String '���صĽ����Ϣ
    Dim strTag As String, arrTag() As String
    Dim L_Item As LisItem
    Dim objXMLNodeList As Object, objXMLNode As Object, objChildNode As Object, objPChildNode As Object
    Dim strFirstName As String, strChildFirstName As String
    Dim lngStartRow As Long
    Dim strTmp As String, i As Integer
    '120692:��Ӽ�����Ŀ
    If mobjPublicLis Is Nothing Then Exit Sub
    If tvw_s.SelectedItem.Key = "��" Then '���ڵ�
        '���ռ�����Ŀ��ָ��ֲ�
    Else
        lngKey = Val(Mid(tvw_s.SelectedItem.Key, 2))
'        objNode.Tag = L_Item.���鱨��id & "'" & L_Item.����id & "'" & L_Item.������־ & "'" & L_Item.�걾��� & "'" & L_Item.�Ƿ�΢���� & "'" & _
'                        L_Item.������� & "'" & L_Item.������ & "'" & L_Item.����� & "'" & L_Item.���ʱ�� & "'" & L_Item.����ʱ��
        strTag = tvw_s.SelectedItem.Tag
        arrTag = Split(strTag, "'")
        L_Item.�Ƿ�΢���� = arrTag(4)
        Call InitVsf(2, L_Item.�Ƿ�΢���� = 1)
        strXMLLIS = mobjPublicLis.GetLaboratoryReportResultList(lngKey)
        If strXMLLIS <> "" Then
            If OpenXMLDocument(strXMLLIS) = True Then
                If L_Item.�Ƿ�΢���� = 0 Then
                    Set objXMLNodeList = mobjXML.selectNodes(".//��ͨ��Ŀ//ָ������").Item(0).childNodes
                    strFirstName = objXMLNodeList.Item(0).nodename
                    vsList.Redraw = flexRDNone
                    For Each objXMLNode In objXMLNodeList
                        If objXMLNode.nodename = strFirstName Then
                            vsList.Rows = vsList.Rows + 1
                            vsList.Cell(flexcpPicture, vsList.Rows - 1, 0) = imgList.ListImages(2).Picture
                        End If
                        Select Case objXMLNode.nodename
                            Case "ָ��id"
                                vsList.RowData(vsList.Rows - 1) = objXMLNode.Text
                            Case "ָ�����", "ָ��Ӣ����", "ָ��������", "������", "��λ", "�����־", "����ο�"
                                vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex(objXMLNode.nodename)) = objXMLNode.Text
                        End Select
                    Next
                    vsList.Redraw = flexRDDirect
                Else '΢������Ŀ
                    Set objXMLNodeList = mobjXML.selectNodes(".//΢������Ŀ").Item(0).childNodes
                    strFirstName = objXMLNodeList.Item(0).nodename
                    vsList.Redraw = flexRDNone
                    For Each objXMLNode In objXMLNodeList
                        If objXMLNode.nodename = strFirstName Then
                            vsList.Rows = vsList.Rows + 1
                            vsList.MergeRow(vsList.Rows - 1) = True
                            vsList.Cell(flexcpPicture, vsList.Rows - 1, 0) = imgList.ListImages(2).Picture
                            lngStartRow = vsList.Rows - 1
                            strTmp = ""
                        End If
                        Select Case objXMLNode.nodename
                            Case "ϸ��id"
                                vsList.RowData(vsList.Rows - 1) = objXMLNode.Text
                            Case "ϸ����", "����", "��ҩ����"
                                vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex(objXMLNode.nodename)) = objXMLNode.Text
                            Case "�����ؽ���б�"
                                vsList.IsSubtotal(lngStartRow) = True
                                vsList.RowOutlineLevel(lngStartRow) = 0
                                strTmp = vsList.TextMatrix(lngStartRow, vsList.ColIndex("ϸ����")) & "[" & vsList.TextMatrix(lngStartRow, vsList.ColIndex("����")) & "]"
                                vsList.TextMatrix(lngStartRow, vsList.ColIndex("ϸ������")) = strTmp
                                '����Ŀ����غͽ��
                                Set objPChildNode = objXMLNode.childNodes
                                strChildFirstName = objPChildNode.Item(0).nodename
                                For Each objChildNode In objPChildNode
                                    If objChildNode.nodename = strChildFirstName Then
                                        vsList.Rows = vsList.Rows + 1
                                        vsList.MergeRow(vsList.Rows - 1) = False
                                        vsList.RowData(vsList.Rows - 1) = vsList.RowData(lngStartRow)
                                        vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex("ϸ����")) = vsList.TextMatrix(lngStartRow, vsList.ColIndex("ϸ����"))
                                        vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex("����")) = vsList.TextMatrix(lngStartRow, vsList.ColIndex("����"))
                                        vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex("��ҩ����")) = vsList.TextMatrix(lngStartRow, vsList.ColIndex("��ҩ����"))
                                        vsList.Cell(flexcpPicture, vsList.Rows - 1, 0) = imgList.ListImages(2).Picture
                                        vsList.IsSubtotal(vsList.Rows - 1) = True
                                        vsList.RowOutlineLevel(vsList.Rows - 1) = 1
                                        vsList.IsCollapsed(vsList.Rows - 1) = flexOutlineExpanded
                                    End If
                                    Select Case objChildNode.nodename
                                        Case "������", "�����ؽ��", "��ҩ��", "ҩ������", "�÷�����1", "�÷�����2", "ѪҩŨ��1", "ѪҩŨ��2", "��ҩŨ��1", "��ҩŨ��2"
                                            vsList.TextMatrix(vsList.Rows - 1, vsList.ColIndex(objChildNode.nodename)) = objChildNode.Text
                                            vsList.TextMatrix(lngStartRow, vsList.ColIndex(objChildNode.nodename)) = strTmp
                                    End Select
                                Next
                        End Select
                    Next
                    For i = vsList.ColIndex("ϸ������") To vsList.Cols - 1
                        vsList.MergeCol(i) = True
                    Next
                    vsList.Redraw = flexRDDirect
                End If
            End If
        End If
    End If
End Sub

Private Sub InitVsf(ByVal intType As Integer, Optional bln�Ƿ�΢���� As Boolean = False)
'���ܣ���ʼ�������Ϣ
'intType:0-�ʾ�ѡ��,1-ҽ��,2-���� (intType-2ʱ��Ҫ�����Ƿ�΢����)
    With vsList
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 1
        .Cols = 0
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarSimpleLeaf
        .Editable = flexEDNone
        .MergeCells = flexMergeNever
        Select Case intType
            Case 0
                .Cols = 4
                .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                .TextMatrix(0, 1) = "���": .ColKey(1) = "���": .ColWidth(1) = 795
                .TextMatrix(0, 2) = "����": .ColKey(2) = "����": .ColWidth(2) = 1530
                .TextMatrix(0, 3) = "����": .ColKey(3) = "����": .ColWidth(3) = 2535
            Case 1
                .Cols = 12
                .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                .TextMatrix(0, 1) = "���ID": .ColKey(1) = "���ID": .ColWidth(1) = 0: .ColHidden(1) = True
                .TextMatrix(0, 2) = "": .ColKey(2) = "һ��": .ColWidth(2) = 315
                .TextMatrix(0, 3) = "ҽ������": .ColKey(3) = "ҽ������": .ColWidth(3) = 4000
                .TextMatrix(0, 4) = "��������": .ColKey(4) = "��������": .ColWidth(4) = 900
                .TextMatrix(0, 5) = "��ҩ;��": .ColKey(5) = "��ҩ;��": .ColWidth(5) = 1400
                .TextMatrix(0, 6) = "�ܸ�����": .ColKey(6) = "�ܸ�����": .ColWidth(6) = 795
                .TextMatrix(0, 7) = "������Ŀ": .ColKey(7) = "������Ŀ": .ColWidth(7) = 2000
                .TextMatrix(0, 8) = "����ҽ��": .ColKey(8) = "����ҽ��": .ColWidth(8) = 900
                .TextMatrix(0, 9) = "��ʼʱ��": .ColKey(9) = "��ʼʱ��": .ColWidth(9) = 1600
                .TextMatrix(0, 10) = "ҽ������": .ColKey(10) = "ҽ������": .ColWidth(10) = 1500
                .TextMatrix(0, 11) = "�����ı�": .ColKey(11) = "�����ı�": .ColWidth(11) = 0: .ColHidden(11) = True 'a.ҽ������ || a.�������� || b.���㵥λ || c.ҽ������
            Case 2
                If bln�Ƿ�΢���� = False Then
                    .Cols = 8
                    .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                    .TextMatrix(0, 1) = "ָ�����": .ColKey(1) = "ָ�����": .ColWidth(1) = 900
                    .TextMatrix(0, 2) = "ָ��Ӣ����": .ColKey(2) = "ָ��Ӣ����": .ColWidth(2) = 1530: .ColHidden(2) = True
                    .TextMatrix(0, 3) = "ָ��������": .ColKey(3) = "ָ��������": .ColWidth(3) = 3000
                    .TextMatrix(0, 4) = "������": .ColKey(4) = "������": .ColWidth(4) = 1200
                    .TextMatrix(0, 5) = "��λ": .ColKey(5) = "��λ": .ColWidth(5) = 900
                    .TextMatrix(0, 6) = "�����־": .ColKey(6) = "�����־": .ColWidth(6) = 900
                    .TextMatrix(0, 7) = "����ο�": .ColKey(7) = "����ο�": .ColWidth(7) = 900
                Else
                    .Cols = 15
                    .OutlineCol = 4
                    .MergeCells = flexMergeRestrictRows
                    .TextMatrix(0, 0) = "": .ColKey(0) = "Key": .ColWidth(0) = 315
                    .TextMatrix(0, 1) = "����": .ColKey(1) = "����": .ColWidth(1) = 0: .ColHidden(1) = True
                    .TextMatrix(0, 2) = "��ҩ����": .ColKey(2) = "��ҩ����": .ColWidth(2) = 0: .ColHidden(2) = True
                    .TextMatrix(0, 3) = "ϸ����": .ColKey(3) = "ϸ����": .ColWidth(3) = 0: .ColHidden(3) = True
                    .TextMatrix(0, 4) = "ϸ����": .ColKey(4) = "ϸ������": .ColWidth(4) = 900
                    .TextMatrix(0, 5) = "������": .ColKey(5) = "������": .ColWidth(5) = 3000
                    .TextMatrix(0, 6) = "�����ؽ��": .ColKey(6) = "�����ؽ��": .ColWidth(6) = 1500
                    .TextMatrix(0, 7) = "��ҩ��": .ColKey(7) = "��ҩ��": .ColWidth(7) = 900
                    .TextMatrix(0, 8) = "ҩ������": .ColKey(8) = "ҩ������": .ColWidth(8) = 1200
                    .TextMatrix(0, 9) = "�÷�����1": .ColKey(9) = "�÷�����1": .ColWidth(9) = 1200
                    .TextMatrix(0, 10) = "�÷�����2": .ColKey(10) = "�÷�����2": .ColWidth(10) = 1200
                    .TextMatrix(0, 11) = "ѪҩŨ��1": .ColKey(11) = "ѪҩŨ��1": .ColWidth(11) = 1200
                    .TextMatrix(0, 12) = "ѪҩŨ��2": .ColKey(12) = "ѪҩŨ��2": .ColWidth(12) = 1200
                    .TextMatrix(0, 13) = "��ҩŨ��1": .ColKey(13) = "��ҩŨ��1": .ColWidth(13) = 1200
                    .TextMatrix(0, 14) = "��ҩŨ��2": .ColKey(14) = "��ҩŨ��2": .ColWidth(14) = 1200
                End If
        End Select
    End With
End Sub

Private Function OpenXMLDocument(ByVal strXml As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���XML�ĵ�
    '������strXML-XML�ַ���
    '���أ��ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    On Error GoTo ErrHand
    
    mstrXmlVersion = GetXMLVersion
    
    Set mobjXML = CreateObject("MSXML2.DOMDocument" & mstrXmlVersion)
    
    OpenXMLDocument = mobjXML.loadXML(strXml)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    Set mobjXML = Nothing
    OpenXMLDocument = False
End Function

Private Function GetXMLVersion() As String
    
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo ErrHand
        
    varXMLVersion = Split(".6.0,.4.0", ",")
    
    On Error Resume Next
    If OS.IsDesinMode = True Or zlRegInfo("��Ȩ����") <> "1" Then
        For intLoop = 0 To UBound(varXMLVersion)
            Err = 0
            Set objXML = CreateObject("MSXML2.DOMDocument" & varXMLVersion(intLoop))
            If Err = 0 Then
                strXMLVer = varXMLVersion(intLoop)
                Exit For
            End If
        Next
        On Error GoTo ErrHand
        
        If strXMLVer = "" Then
            MsgBox "����MSXML2.DOMDocument����ʧ��"
            Exit Function
        End If
    Else
        strXMLVer = ""
    End If
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHand:
    MsgBox Err.Description
End Function

Private Sub SetControlEnable(ByVal blnEnable As Boolean)
'���ܣ����Զ������ý��棬�����ó��ý����ϵ������ؼ������ã�ȡ����ָ�
'blnEnable =false ��ʾ����ʾ�Զ������,True��ʾ�ر��Զ������
    Dim objControl As Object
    For Each objControl In Me.Controls
        If InStr(1, ",ImageList,Line,", "," & TypeName(objControl) & ",") = 0 Then
            If objControl.Visible = True Then
                '�ų��Զ����б���
                If InStr(1, ",picDef,picTip,imgTip,lblDefTitle,imgClose,lblPrompt,lbl���,cbo���,lbl���ݸ�ʽ,txtAdvice,lbl�ֶ���Ŀ,cbo�ֶ�,cmdAdd,cmdCheck,cmdGO,cmdClose,", "," & objControl.Name & ",") = 0 Then
                    objControl.Enabled = blnEnable
                End If
            End If
        End If
    Next
End Sub
Private Sub SetTagһ����ҩ(Optional ByVal lngRow As Long)
'���ܣ���һ����ҩ��ҽ��ǰ�ӱ�־
    Dim i As Long
    Dim lngBg As Long, lngEd As Long
    Dim j As Long
    Dim lngStart As Long, lngEnd As Long
    With vsList
        If lngRow = 0 Then
            lngStart = .FixedRows
            lngEnd = .Rows - 1
        Else
            lngStart = lngRow
            lngEnd = lngRow
        End If
        For i = lngStart To lngEnd
             lngBg = -1: lngEd = -1
             If RowInһ����ҩ(i, lngBg, lngEd) Then
                For j = lngBg To lngEd
                    If j = lngBg Then
                        .TextMatrix(j, .ColIndex("һ��")) = "��"
                    ElseIf j = lngEd Then
                        .TextMatrix(j, .ColIndex("һ��")) = "��"
                    Else
                        .TextMatrix(j, .ColIndex("һ��")) = "��"
                    End If
                Next
                If lngEd <> -1 Then
                   i = lngEd + 1
                End If
            End If
        Next
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵��:PASS �е� ��RowInһ����ҩ�� ��˷�����ͬ,�޸Ĵ˷���Ҳ��Ҫͬ���޸� PASSͬ������
    Dim i As Long, blnTmp As Boolean
    With vsList
        If Val(.TextMatrix(lngRow - 1, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("���ID"))) = Val(.TextMatrix(lngRow, .ColIndex("���ID"))) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function



