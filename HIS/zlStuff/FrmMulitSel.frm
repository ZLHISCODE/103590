VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMulitSel 
   BorderStyle     =   0  'None
   Caption         =   "����ѡ����"
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   ControlBox      =   0   'False
   Icon            =   "FrmMulitSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkContinue 
      Caption         =   "����ѡ��(&M)"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picѡ���� 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2535
      ScaleWidth      =   4815
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4440
      Width           =   4815
      Begin VB.PictureBox picUpDown01 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3600
         Picture         =   "FrmMulitSel.frx":0E42
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox picOK 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   220
         Left            =   3240
         Picture         =   "FrmMulitSel.frx":1184
         ScaleHeight     =   225
         ScaleWidth      =   270
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��"
         Top             =   0
         Width           =   270
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfѡ�� 
         Height          =   2085
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   4275
         _cx             =   7541
         _cy             =   3678
         Appearance      =   0
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   32
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmMulitSel.frx":14C6
         ScrollTrack     =   0   'False
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
         ExplorerBar     =   7
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
      End
      Begin VB.Label lblѡ�� 
         BackColor       =   &H00FFEDDD&
         Caption         =   "ѡ������"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3885
      End
   End
   Begin VB.PictureBox picSplit02_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   40
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   2535
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4200
      Width           =   2535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsColSet 
      Height          =   3210
      Left            =   3285
      TabIndex        =   2
      Top             =   735
      Visible         =   0   'False
      Width           =   2700
      _cx             =   4762
      _cy             =   5662
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmMulitSel.frx":192F
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Ellipsis        =   1
      ExplorerBar     =   2
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHead 
      Height          =   2520
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8040
      _cx             =   14182
      _cy             =   4445
      Appearance      =   0
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483637
      BackColorAlternate=   -2147483628
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483644
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmMulitSel.frx":197C
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   7
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
      Begin VB.Image imgLeft 
         Height          =   240
         Left            =   30
         Picture         =   "FrmMulitSel.frx":1A51
         Top             =   30
         Width           =   240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBatch 
      Height          =   1050
      Left            =   0
      TabIndex        =   1
      Top             =   3030
      Width           =   8025
      _cx             =   14155
      _cy             =   1852
      Appearance      =   0
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
      BackColor       =   -2147483628
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483637
      BackColorAlternate=   -2147483628
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483644
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmMulitSel.frx":1FDB
      ScrollTrack     =   0   'False
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
      ExplorerBar     =   7
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
      Begin VB.Image imgBatch 
         Height          =   240
         Left            =   30
         Picture         =   "FrmMulitSel.frx":20B0
         Top             =   30
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList imgsMain 
      Left            =   6480
      Top             =   4680
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
            Picture         =   "FrmMulitSel.frx":263A
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMulitSel.frx":298C
            Key             =   "Up"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMulitSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--�������--
Private mintEditState As Integer                 '�༭״̬(1-���;2-����)
Private mlngԴ�ⷿID As Long                     'Դ�ⷿID
Private mlngĿ�ⷿID As Long                     'Ŀ�ⷿID
Private mlngʹ�ò���ID As Long                   'ʹ�ò���ID
Private mlng��Ӧ��ID As Long                     '��Ӧ��ID
Private mstrInput As String                      '�����ִ�
Private mobjOut As Form                          'ʹ�ñ�����Ĵ��壨�����ṩһ��������¼�������Է��أ�
Private mblnSelect As Boolean                    '�Ƿ�����ѡ��

Private mblnStartUp As Boolean                   '�����ɹ�
Private mblnFirstStart As Boolean                '��һ������
Private mrsUnit As New ADODB.Recordset           '��λ
Private mstrUnit As String                       '��λ����
Private mstrUnitString As String                 'SQL�ִ�
Private mintStockCheck As Integer                '�����
Private mstrFindStyle As String                  'ƥ�䷽ʽ
Private mbln�̵㵥 As Boolean                    '�̵㵥�ݱ�־
Private mbln������ As Boolean                    '�Ƿ����ӿ����ι�����
Private mblnCheck As Boolean                     '�Ƿ����س���ԭ��(�����λ�ʱ����������)
Private mblnPrice As Boolean                     '�Ƿ�����ʱ�ۻ��������������
Private mstrCode As String                       '����
Private mblnTrackUsing As Boolean                '�������ò���

'������ʹ�ü�¼��
Private mrsData As New ADODB.Recordset           '������;����
Private mrsCard As New ADODB.Recordset           '���Ŀ�Ƭ
Private mrsStock As New ADODB.Recordset          '���Ĺ��
Private mstrTittle As String                     'ѡ��������

'���ؼ�¼��
Private mrsReturn As ADODB.Recordset             '���ؼ�¼��(������Ϣ������,����Ŀ¼������,���Ŀ��������)
Private mint�ⷿ As Integer                      '1-���Ŀ�;2-����;3-�Ƽ���
Private mint���� As Integer                      '0-������;1-�ⷿ����;2-���÷���;3-���Ŀ����÷���
Private mblnʱ�� As Boolean                      'ʱ��
Private mblnStock As Boolean
Private mstrCardSortBy As String                 '���Ŀ�Ƭ������
Private mstrPhysicSortBy As String               '���Ĺ��������
Private mlngCardRow As Long
Private mlngPhysicRow As Long
Private mlngLastSelect����ID As Long             '�ϴ�ѡ��Ĳ���ID�������Ƿ�ˢ�£�
Private mblnɢװ��λ As Boolean
Private mblnֻ��ʾ�������� As Boolean
Private mbln����ʾ������� As Boolean
Private mbln��ʾ���� As Boolean                 '�Ƿ���ʾ�����б� true-��ʾ���Σ�false-������ʾ����,��Ҫ�ǳ���ҵ�������ж��Ƿ�������ҵ�������첻��ȷ����ģʽ����ʾ�����б�
Private mlngModule As Long
Private mblnSelectSucess As Boolean
Private mbln���޴洢�ⷿ���� As Boolean
Private mblnCostView As Boolean                 '�鿴�ɱ��������Ϣ true-����鿴 false-������鿴
Private mblnProvider As Boolean                 '�鿴�ϴι�Ӧ�������Ϣ true-����鿴 false-������鿴
Private mstrPrivs As String                     '����ԱȨ��
Private mbln�Ƿ���� As Boolean

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'����get���ÿ��󣬷��صĿ���������ʵ��������ʵ�ʽ�ʵ�ʲ��
Private msin�������� As Single
Private msinʵ������ As Single
Private msinʵ�ʽ�� As Single
Private msinʵ�ʲ�� As Single
Private Const MFRM_MIN_WIDTH = 8040
Private Const MFRM_MIN_HEIGHT = 3630
Private mstr�̵�ʱ�� As String

'--����--
'Private Const strFormat As String = "'999999999990.9999'"
Private Type WinLocate
    Left As Double
    Top As Double
    lngTxtW As Long
    lngTxtH As Long
End Type
Private mWindowPosition As WinLocate           '����λ��

Public Property Get In_�༭״̬() As Integer
    In_�༭״̬ = mintEditState
End Property

Public Property Let In_�༭״̬(ByVal vNewValue As Integer)
    mintEditState = vNewValue
End Property

Public Property Get In_Դ�ⷿ() As Long
    In_Դ�ⷿ = mlngԴ�ⷿID
End Property

Public Property Let In_Դ�ⷿ(ByVal vNewValue As Long)
    mlngԴ�ⷿID = vNewValue
End Property

Public Property Get In_�ִ�() As String
    In_�ִ� = mstrInput
End Property

Public Property Let In_�ִ�(ByVal vNewValue As String)
    mstrInput = vNewValue
End Property

Public Property Get In_Ŀ�ⷿ() As Long
    In_Ŀ�ⷿ = mlngĿ�ⷿID
End Property

Public Property Let In_Ŀ�ⷿ(ByVal vNewValue As Long)
    mlngĿ�ⷿID = vNewValue
End Property

Public Property Get In_����() As Long
    In_���� = mlngʹ�ò���ID
End Property

Public Property Let In_����(ByVal vNewValue As Long)
    mlngʹ�ò���ID = vNewValue
End Property

Public Property Let In_MainFrm(ByVal vNewValue As Form)
    Set mobjOut = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntMain As Integer = 1, Optional ByVal BlnSetHeader As Boolean = False)
    Dim intCol As Integer
    
    '���ø��б�ؼ��ĸ�ʽ
    Select Case IntMain
    Case 1
        With vsHead
            For intCol = 0 To .Cols - 1
                .FixedAlignment(intCol) = flexAlignCenterCenter
                .ColKey(intCol) = UCase(.TextMatrix(0, intCol))
                If InStr(1, .ColKey(intCol), "ID") > 0 Then
                    .ColHidden(intCol) = True: .ColWidth(intCol) = 0
                    '.coldata(i):1-�̶�,-1-����ѡ,0-��ѡ
                    .ColData(intCol) = -1
                ElseIf InStr(1, .ColKey(intCol), "����") > 0 Or _
                   (InStr(1, .ColKey(intCol), "��") > 0 And .ColKey(intCol) <> "ʱ��") Then
                    .ColAlignment(intCol) = flexAlignRightCenter
                    .ColWidth(intCol) = 1000
                    '.coldata(i):1-�̶�,-1-����ѡ,0-��ѡ
                    .ColData(intCol) = 0
                Else
                    .ColData(intCol) = 0
                    Select Case .ColKey(intCol)
                    Case "��Ч��", "ɢװ", "��װ", "ϵ��", "һ���Բ���", _
                        "�޾��Բ���", "���Ч��", "���ʧЧ��", "�ⷿ����", "���÷���", "ʱ��"
                        .ColAlignment(intCol) = flexAlignCenterCenter
                    Case "����", "ͨ������"
                        '.coldata(i):1-�̶�,-1-����ѡ,0-��ѡ
                        .ColAlignment(intCol) = flexAlignLeftCenter
                        .ColData(intCol) = 1
                    Case Else
                        .ColAlignment(intCol) = flexAlignLeftCenter
                    End Select
                End If
            Next
            '�Զ������п�
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            .Row = 1
            '�ָ��п�
             
            zl_vsGrid_Para_Restore mlngModule, vsHead, mstrTittle, "�����Ϣ", False
            
            If mlngModule = 1725 Or .ColWidth(.ColIndex("�ϴι�Ӧ��")) = 0 Then .ColWidth(.ColIndex("�ϴι�Ӧ��")) = IIf(mblnProvider = True, 1300, 0)
            If mlngModule = 1725 Or .ColHidden(.ColIndex("�ϴι�Ӧ��")) = True Then .ColHidden(.ColIndex("�ϴι�Ӧ��")) = Not mblnProvider
            If .ColWidth(.ColIndex("ָ�������")) <> 0 Then .ColWidth(.ColIndex("ָ�������")) = 0
            If .ColHidden(.ColIndex("ָ�������")) = False Then .ColHidden(.ColIndex("ָ�������")) = True
            
            If mblnCostView = False Then
                If .ColHidden(.ColIndex("�����")) = False Then .ColHidden(.ColIndex("�����")) = True
                If .ColHidden(.ColIndex("���³ɱ���")) = False Then .ColHidden(.ColIndex("���³ɱ���")) = True
                If .ColHidden(.ColIndex("ָ��������")) = False Then .ColHidden(.ColIndex("ָ��������")) = True
                .ColWidth(.ColIndex("�����")) = 0
                .ColData(.ColIndex("�����")) = -1
                .ColWidth(.ColIndex("���³ɱ���")) = 0
                .ColData(.ColIndex("���³ɱ���")) = -1
                .ColWidth(.ColIndex("ָ��������")) = 0
                .ColData(.ColIndex("ָ��������")) = -1
            Else
                If .ColHidden(.ColIndex("�����")) = True Then .ColHidden(.ColIndex("�����")) = False
                If .ColHidden(.ColIndex("���³ɱ���")) = True Then .ColHidden(.ColIndex("���³ɱ���")) = False
                If .ColHidden(.ColIndex("ָ��������")) = True Then .ColHidden(.ColIndex("ָ��������")) = False
                If .ColWidth(.ColIndex("�����")) = 0 Then .ColWidth(.ColIndex("�����")) = 1000
                If .ColWidth(.ColIndex("���³ɱ���")) = 0 Then .ColWidth(.ColIndex("���³ɱ���")) = 1000
                If .ColWidth(.ColIndex("ָ��������")) = 0 Then .ColWidth(.ColIndex("ָ��������")) = 1000
            End If
        End With
    Case 0
        With vsBatch
            DoEvents
            For intCol = 0 To .Cols - 1
                .FixedAlignment(intCol) = flexAlignCenterCenter
                .ColKey(intCol) = UCase(.TextMatrix(0, intCol))
                '.coldata(i):1-�̶�,-1-����ѡ,0-��ѡ
                If InStr(1, .ColKey(intCol), "ID") > 0 Then
                    .ColHidden(intCol) = True: .ColWidth(intCol) = 0
                    .ColData(intCol) = -1
                ElseIf InStr(1, .ColKey(intCol), "����") > 0 Then
                    .ColHidden(intCol) = True: .ColWidth(intCol) = 0
                    .ColData(intCol) = 0
                ElseIf InStr(1, .ColKey(intCol), "����") > 0 Or _
                    InStr(1, .ColKey(intCol), "��") > 0 Then
                    .ColAlignment(intCol) = flexAlignRightCenter
                    .ColWidth(intCol) = 1000
                    .ColData(intCol) = 0
                Else
                    Select Case .ColKey(intCol)
                    Case "����"
                        .ColAlignment(intCol) = flexAlignCenterCenter
                        .ColData(intCol) = 1
                    Case "��������"
                        .ColAlignment(intCol) = flexAlignCenterCenter
                        .ColData(intCol) = 0
                    Case Else
                        .ColAlignment(intCol) = flexAlignLeftCenter
                        .ColData(intCol) = 0
                    End Select
                End If
            Next
            '�Զ������п�
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            If .Rows >= 2 Then .Row = 1
            '�ָ��п�
            zl_vsGrid_Para_Restore mlngModule, vsBatch, mstrTittle, "������Ϣ", False
            If mblnCostView = False Then
                .ColWidth(.ColIndex("�����")) = 0
                .ColWidth(.ColIndex("�ϴι���")) = 0
                .ColWidth(.ColIndex("�ɱ���")) = 0
                .ColData(.ColIndex("�����")) = -1
                .ColData(.ColIndex("�ϴι���")) = -1
                .ColData(.ColIndex("�ɱ���")) = -1
            End If
        End With
    End Select
End Sub

Private Sub OnCancel()
    Unload Me
    Exit Sub
End Sub

Private Sub OnSelect()
    Dim blnValid As Boolean
    
    If chkContinue.Value = 0 Then
        If In_�༭״̬ = 2 Then If CheckData = False Then Exit Sub
        '�������������������Ƿ�һ��
        If In_�༭״̬ = 2 Then
            blnValid = ���������(mlngԴ�ⷿID, mlngLastSelect����ID)
        Else
            blnValid = ���������(mlngĿ�ⷿID, mlngLastSelect����ID)
        End If
        If Not blnValid Then
            MsgBox "���ָ������ڵ�ǰ�ⷿ�еĿ���¼���ڴ��󣨿����ǻ����������ô������鵱ǰ�ⷿ�Ĳ������ʼ������ĵķ������ԣ���", vbInformation, gstrSysName
            Exit Sub
        End If
        '��װ��¼��
        If CombinateRec = False Then Exit Sub
        mblnSelectSucess = True
        Unload Me
        Exit Sub
    Else
        If CombinateRec = False Then Exit Sub
        mblnSelectSucess = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub chkContinue_Click()

    If vsfѡ��.Rows > 2 And chkContinue.Value = 0 Then
        If MsgBox("�Ѿ���ѡ�����Ĵ��ڣ�ȡ��������ѡ�񡱽������ѡ�������ģ���ȷ����" _
            , vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            vsfѡ��.Rows = 1
            vsfѡ��.Rows = 2
            lblѡ��.Caption = "ѡ��ҩƷ"
        Else
            chkContinue.Value = 1
            Exit Sub
        End If
        
    End If

    picѡ����.Visible = chkContinue.Value = 1
    picSplit02_S.Visible = chkContinue.Value = 1
    Form_Resize
    
    
    If chkContinue.Value = 0 Then
        picѡ����.Tag = "չ��"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
        picSplit02_S.MousePointer = 0
    End If
    
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then
        Unload Me
        Exit Sub
    End If

    Call ReSetWindowsFormLocal
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    mblnStartUp = False
    mblnFirstStart = False
    Call vsBatch_LostFocus:
    Call vsHead_LostFocus
    'ȡ�ۼ۵�λ
    mstrUnit = ""
    mstrFindStyle = IIf(gstrMatchMethod = "0", "%", "")
    mstrUnitString = ""
    mintStockCheck = 0
    mlngLastSelect����ID = 0
    vsBatch.Visible = (In_�༭״̬ = 2)
    
    picѡ����.Visible = False
    picSplit02_S.Visible = False
    chkContinue.Visible = mbln�Ƿ���� = False
    picѡ����.Tag = "չ��"
    
    
    On Error GoTo ErrHandle
    '��ʼ����¼��
    InitRec
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    If mlngModule = 1725 Then
        mblnProvider = zlStr.IsHavePrivs(mstrPrivs, "�鿴��Ӧ��")
    Else
        mblnProvider = True
    End If
    
    If mstrInput = "" Then Exit Sub
    If mobjOut Is Nothing Then
        MsgBox "��ָ�������壡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��λ
    With mWindowPosition
        Me.Left = .Left
        Me.Top = .Top
    End With
    
    '��ȡ��ǰ�����Ʋ���
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ���ϳ����� Where �ⷿID=[1]"
    Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID)
    
    If Not mrsUnit.EOF Then
        mintStockCheck = mrsUnit!�����
    End If
    
    '���Դ�ⷿ�Ƿ�Ϊ���Ŀ�
    If mlngԴ�ⷿID <> 0 Then
        mint�ⷿ = 3
        gstrSQL = "select ����ID from ��������˵�� where (�������� like '���ϲ���' Or �������� like '%�Ƽ���') And ����id=[1]"
        Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID)
        If mrsUnit.EOF Then
            gstrSQL = "select ����ID from ��������˵�� where �������� In ('���Ŀ�', '����ⷿ') And ����id=[1]"
            Set mrsUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID)
            If Not mrsUnit.EOF Then mint�ⷿ = 1
        Else
            mint�ⷿ = 2
        End If
    End If
    
    '������ʹ�õĵ�λ����
    If mlngʹ�ò���ID <> 0 Then
        If mblnɢװ��λ Then
            mstrUnitString = "/1"
        Else
            mstrUnitString = "/nvl(����ϵ��,1)"
        End If
    End If
    
  
    '���˺�:����С����ʽ����
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_�ɱ���, True)
        .FM_��� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_���, True)
        .FM_���ۼ� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_�ۼ�, True)
        .FM_���� = GetFmtString(IIf(mblnɢװ��λ, 0, 1), g_����, True)
    End With
    
    mblnStartUp = RefreshData
    
    On Error Resume Next
    If mrsCard.RecordCount = 1 Then
        If Not (((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And mblnPrice) Or mintEditState = 1 Then OnSelect
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    mblnFirstStart = True
    
    With vsHead
        .Top = chkContinue.Height + chkContinue.Top * 2
        .Height = IIf(vsBatch.Visible = False, Me.ScaleHeight - .Top, (Me.ScaleHeight - .Top) / 2)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
        .Width = Me.ScaleWidth
        
    End With
    With vsBatch
        .Top = vsHead.Height + vsHead.Top
        .Width = vsHead.Width
        .Height = Me.ScaleHeight - (vsHead.Height + vsHead.Top)
        .Left = ScaleLeft
    End With
    
    If picSplit02_S.Visible Then
        '���÷ֽ��ߵ�top
        If vsBatch.Visible Then '���οɼ�
            vsBatch.Height = vsBatch.Height - (lblѡ��.Height + picSplit02_S.Height)
            picSplit02_S.Top = vsBatch.Top + vsBatch.Height
        Else
            vsHead.Height = vsHead.Height - (lblѡ��.Height + picSplit02_S.Height)
            picSplit02_S.Top = vsHead.Top + vsHead.Height
        End If
        picSplit02_S.Left = 0
        picSplit02_S.Width = vsHead.Width
    End If
    
    If picѡ����.Visible Then
        picѡ����.Width = vsHead.Width
        picѡ����.Height = lblѡ��.Height
        
        picѡ����.Top = picSplit02_S.Top + picSplit02_S.Height
        
        With lblѡ��
            .Top = 0
            .Left = 0
            .Width = picѡ����.Width
        End With
        With picUpDown01
            .Left = picѡ����.Width - .Width
            .Top = 0
        End With
        
        With picOK
            .Left = picUpDown01.Left - .Width
            .Top = 0
        End With
        
        With vsfѡ��
            .Top = lblѡ��.Height
            .Left = 0
            .Width = lblѡ��.Width
        End With
        
        If picѡ����.Tag = "����" Then
            picѡ����.Tag = "չ��"
            Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
            picSplit02_S.MousePointer = 0
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrCode = ""
    SaveWinState Me
    zl_vsGrid_Para_Save mlngModule, vsHead, mstrTittle, "�����Ϣ", False
    zl_vsGrid_Para_Save mlngModule, vsBatch, mstrTittle, "������Ϣ", False
End Sub

 

Private Sub imgBatch_Click()
    Call LoadFulltoColSel(True)
End Sub


Private Sub picOK_Click()
    OnSelect
End Sub

Private Sub picSplit02_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit02_S
        If .Top + y < vsHead.Top + 500 Then Exit Sub
        If .Top + y > Me.ScaleHeight - 500 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    picѡ����.Move picѡ����.Left, picѡ����.Top + y, picѡ����.Width, picѡ����.Height - y
    vsfѡ��.Move vsfѡ��.Left, vsfѡ��.Top, picѡ����.Width, vsfѡ��.Height - y
    
End Sub

Private Sub picUpDown01_Click()
    If picѡ����.Tag = "չ��" Then
        picѡ����.Tag = "����"
        Set picUpDown01.Picture = imgsMain.ListImages(1).Picture
        picSplit02_S.MousePointer = 7

        picSplit02_S.Top = picSplit02_S.Top - vsfѡ��.Height
        picѡ����.Top = picѡ����.Top - vsfѡ��.Height
        picѡ����.Height = picѡ����.Height + vsfѡ��.Height
    Else
        picѡ����.Tag = "չ��"
        Set picUpDown01.Picture = imgsMain.ListImages(2).Picture
        picSplit02_S.MousePointer = 0
        
        Form_Resize
    End If
End Sub

Private Sub vsBatch_Click()
'    Dim StrHeader As String
'    Dim intCol As Integer
'    'ʵ��������
'    With vsBatch
'        If .MouseRow <> 0 Then Exit Sub
'        If mrsStock.EOF Then Exit Sub
'
'        StrHeader = .TextMatrix(0, .MouseCol)
'        Set .DataSource = Nothing
'        If Mid(mstrPhysicSortBy, 2) = StrHeader Then
'            mstrPhysicSortBy = IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
'            mrsStock.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(mstrPhysicSortBy, 1, 1) = "A", " Desc", " Asc")
'        Else
'            mstrPhysicSortBy = "A" & .TextMatrix(0, .MouseCol)
'            mrsStock.Sort = .TextMatrix(0, .MouseCol) & " Asc"
'        End If
'        Set .DataSource = mrsStock
'
'        For intCol = 0 To .Cols - 1
'            .ColAlignmentFixed(intCol) = 4
'        Next
'
'        Call SetFormat(0, False)
'    End With
End Sub

Private Sub vsBatch_DblClick()
    With mrsStock
        If .RecordCount <> 0 Then .MoveFirst
        If .EOF Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
    End With
    
    If mblnSelect Then
        If chkContinue.Value = 1 Then
            FillVSFѡ��
            Exit Sub
        End If
        
        OnSelect
    End If
End Sub
Private Sub vsBatch_GotFocus()
    With vsBatch
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
        .BackColorSel = &H8000000D
    End With
End Sub
Private Sub vsBatch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then vsBatch_DblClick
End Sub

Private Sub vsBatch_LostFocus()
    With vsBatch
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        .BackColorSel = &H8000000A
    End With
End Sub

  


Private Sub vsfѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfѡ��.Rows > 2 Then
            If vsfѡ��.Row <> vsfѡ��.Rows - 1 Then vsfѡ��.RemoveItem vsfѡ��.Row
            If vsfѡ��.Rows = 2 Then
                lblѡ��.Caption = "ѡ������"
            Else
                lblѡ��.Caption = "ѡ�����ģ�" & vsfѡ��.Rows - 2 & "����"
            End If
        End If
    End If
End Sub

Private Sub vsHead_DblClick()
    If mrsCard.EOF Then Exit Sub
    If mrsCard.RecordCount = 0 Then Exit Sub
    
    If mblnSelect Then '����ѡ��Ž���
        If chkContinue.Value = 1 Then
            FillVSFѡ��
            Exit Sub
        End If

        OnSelect
    Else
        MsgBox "������û�п�棬���ܼ���������", vbInformation, gstrSysName
    End If
End Sub


Private Sub FillVSFѡ��()
    Dim blnEof As Boolean         '�Ƿ�������ο��
    Dim i As Integer
    Dim blnValid    As Boolean
    
    '���ҩ�����ظ�
    If chkContinue.Value = 1 Then
        For i = 1 To vsfѡ��.Rows - 2
            If Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����ID"))) = Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("����ID"))) Then
                If vsBatch.Visible Then
                    If vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����")) = vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("����")) Then
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            End If
        Next
    End If
    
    If In_�༭״̬ = 2 Then If CheckData = False Then Exit Sub
        
    '�������������������Ƿ�һ��
    If In_�༭״̬ = 2 Then
        blnValid = ���������(mlngԴ�ⷿID, mlngLastSelect����ID)
    Else
        blnValid = ���������(mlngĿ�ⷿID, mlngLastSelect����ID)
    End If
    
    If Not blnValid Then
        ShowMsgBox "���ָ������ڵ�ǰ�ⷿ�еĿ���¼���ڴ��󣨿����ǻ���������" & vbCrLf & "�ô������鵱ǰ�ⷿ�Ĳ������ʼ������ĵķ������ԣ���"
        Exit Sub
    End If
    
    With mrsCard
        If .RecordCount <> 0 Then .MoveFirst
        .Find "����ID=" & Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("����ID")))
        If .EOF Then
            MsgBox "�����ڲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mbln��ʾ���� = True Then 'ֻ����ʾ���ε�����²���Ҫ�����²���
            If ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And In_�༭״̬ = 2 Then
                With mrsStock
                    If .RecordCount <> 0 Then .MoveFirst
                    .Find "����=" & Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("����")))
                    If .EOF Then
                        blnEof = True
                        If mblnPrice Then
                            MsgBox "�����ڲ�����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End With
            End If
        End If
    End With
    
    
    'װ����д���¼��������������ʹ��
    With vsfѡ��
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 2, .ColIndex("����ID")) = mrsCard!����ID
        .TextMatrix(.Rows - 2, .ColIndex("����id")) = mrsCard!����id
        .TextMatrix(.Rows - 2, .ColIndex("����id")) = mrsCard!����id
        
        .TextMatrix(.Rows - 2, .ColIndex("����")) = mrsCard!����
        .TextMatrix(.Rows - 2, .ColIndex("����")) = zlStr.Nvl(mrsCard!ͨ������)
        .TextMatrix(.Rows - 2, .ColIndex("��Ʒ��")) = zlStr.Nvl(mrsCard!��Ʒ��)
        .TextMatrix(.Rows - 2, .ColIndex("���")) = mrsCard!���
        .TextMatrix(.Rows - 2, .ColIndex("����")) = "" & mrsCard!����
        .TextMatrix(.Rows - 2, .ColIndex("�ۼ�")) = zlStr.Nvl(mrsCard!�ۼ�, 0)
        .TextMatrix(.Rows - 2, .ColIndex("ɢװ��λ")) = mrsCard!ɢװ
        .TextMatrix(.Rows - 2, .ColIndex("����ϵ��")) = mrsCard!ϵ��
        .TextMatrix(.Rows - 2, .ColIndex("��װ��λ")) = mrsCard!��װ
        .TextMatrix(.Rows - 2, .ColIndex("���Ч��")) = "" & mrsCard!��Ч��
        .TextMatrix(.Rows - 2, .ColIndex("���Ч��")) = "" & mrsCard!���Ч��
        .TextMatrix(.Rows - 2, .ColIndex("���ʧЧ��")) = "" & mrsCard!���ʧЧ��
        .TextMatrix(.Rows - 2, .ColIndex("һ���Բ���")) = IIf(mrsCard!һ���Բ��� = "��", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("�޾��Բ���")) = IIf(mrsCard!�޾��Բ��� = "��", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("�ⷿ����")) = IIf(mrsCard!�ⷿ���� = "��", 1, 0)
        .TextMatrix(.Rows - 2, .ColIndex("���÷���")) = IIf(mrsCard!���÷��� = "��", 1, 0)
        
        .TextMatrix(.Rows - 2, .ColIndex("ʱ��")) = IIf(mrsCard!ʱ�� = "��", 1, 0)

        '�����ҷ���
        If In_�༭״̬ = 2 And ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) Then
            If mbln��ʾ���� = True Then 'ֻ����ʾ��������²���Ҫ�����²���������������
                If vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("����")) = "����������������" Then
                    .TextMatrix(.Rows - 2, .ColIndex("����")) = -1
                Else
                    If Not blnEof Then
                        .TextMatrix(.Rows - 2, .ColIndex("����")) = Val(mrsStock!����)
                        .TextMatrix(.Rows - 2, .ColIndex("����")) = "" & mrsStock!����
                        .TextMatrix(.Rows - 2, .ColIndex("Ч��")) = "" & mrsStock!ʧЧ��
                        .TextMatrix(.Rows - 2, .ColIndex("���ʧЧ��")) = "" & mrsStock!���ʧЧ��
                        .TextMatrix(.Rows - 2, .ColIndex("����")) = "" & mrsStock!����
                        .TextMatrix(.Rows - 2, .ColIndex("��������")) = "" & mrsStock!��������
                        .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = "" & mrsStock!��׼�ĺ�
                        .TextMatrix(.Rows - 2, .ColIndex("��ҩ��λID")) = "" & mrsStock!�ϴι�Ӧ��id
                        .TextMatrix(.Rows - 2, .ColIndex("��������")) = IIf(IsNull(mrsStock!��������), 0, mrsStock!��������)
                        .TextMatrix(.Rows - 2, .ColIndex("ʵ������")) = IIf(IsNull(mrsStock!�������), 0, mrsStock!�������)
                        .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʽ��")) = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                        .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʲ��")) = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                        If Not mblnStock Then Call Get���ÿ��(.TextMatrix(.Rows - 2, .ColIndex("����ID")), .TextMatrix(.Rows - 2, .ColIndex("����")))
                    End If
                End If
            Else
                If Not mblnStock Then Call Get���ÿ��(mrsCard!����ID, 0)
            End If
        Else
        '���򲻷���
            .TextMatrix(.Rows - 2, .ColIndex("��������")) = IIf(IsNull(mrsCard!��������), 0, mrsCard!��������)
            .TextMatrix(.Rows - 2, .ColIndex("ʵ������")) = IIf(IsNull(mrsCard!�������), 0, mrsCard!�������)
            .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʽ��")) = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
            .TextMatrix(.Rows - 2, .ColIndex("ʵ�ʲ��")) = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
            If In_�༭״̬ = 1 Then
                .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = "" & mrsCard!��׼�ĺ�
            Else
                If mrsStock.RecordCount > 0 Then
                    mrsStock.MoveFirst
                    .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = zlStr.Nvl(mrsStock!��׼�ĺ�)
                Else
                    .TextMatrix(.Rows - 2, .ColIndex("��׼�ĺ�")) = ""
                End If
            End If

            If Not mblnStock Then Call Get���ÿ��(.TextMatrix(.Rows - 2, .ColIndex("����ID")), 0)
        End If
        
        '�������ʾ�Է��ⷿ�Ŀ�棬��������ȡ������
        If Not mblnStock Then
            .TextMatrix(.Rows - 2, .ColIndex("msin��������")) = msin��������
            .TextMatrix(.Rows - 2, .ColIndex("msinʵ������")) = msinʵ������
            .TextMatrix(.Rows - 2, .ColIndex("msinʵ�ʽ��")) = msinʵ�ʽ��
            .TextMatrix(.Rows - 2, .ColIndex("msinʵ�ʲ��")) = msinʵ�ʲ��
        End If
        .TextMatrix(.Rows - 2, .ColIndex("ָ��������")) = mrsCard!ָ��������
        .TextMatrix(.Rows - 2, .ColIndex("ָ�������")) = mrsCard!ָ�������
    End With
    
    lblѡ��.Caption = "ѡ�����ģ�" & vsfѡ��.Rows - 2 & "����"
End Sub



Private Sub vsHead_EnterCell()
    Dim lng�շ�ϸĿID As Long, intCol As Integer, LngSelectRow As Long
    Dim strTmp As String, recGetPrice As New ADODB.Recordset
    Dim strKc As String
    Dim i As Integer
    
   ' On Error Resume Next
    err = 0: On Error GoTo ErrHand:
    With vsHead
        '����ù�����ĵļ۸�ִ��ʱ�仹δִ��,�򴥷�
        If Not mrsCard.EOF Then
            lng�շ�ϸĿID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        End If
        If lng�շ�ϸĿID = 0 Then
            vsBatch.Clear 1
            vsBatch.Rows = 2
            mlngLastSelect����ID = 0
            Exit Sub
        End If
        
        If mlngLastSelect����ID = lng�շ�ϸĿID Then Exit Sub
        mlngLastSelect����ID = lng�շ�ϸĿID
        
        
        '����ѵ�ִ�����ڶ��۸�δִ�У�ִ�м������
        gstrSQL = " Select ID From �շѼ�Ŀ Where �շ�ϸĿID=[1]" & _
                 " And �䶯ԭ��=0" & GetPriceClassString("")
        Set recGetPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�շ�ϸĿID)
        
        With recGetPrice
            If Not .EOF Then
                If Not IsNull(!Id) Then
                    lng�շ�ϸĿID = !Id
                    gstrSQL = "zl_�����շ���¼_Adjust(" & lng�շ�ϸĿID & ")"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption & "-�����۸������¼"
                End If
            End If
        End With
    End With
    
    If In_�༭״̬ = 2 Then
        vsBatch.Visible = False
        '���������Ĺ�������е��������ο����Ϣ
        mblnʱ�� = (vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("ʱ��")) = "��")
        mint���� = 0
        If vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("�ⷿ����")) = "��" Or vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("���÷���")) = "��" Then
            If vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("�ⷿ����")) = "��" And vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("���÷���")) = "��" Then
                mint���� = 3
            ElseIf vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("�ⷿ����")) = "��" Then
                mint���� = 1
            Else
                mint���� = 2
            End If
        End If
        If Not ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) Then         '��������Ĳ�����
            vsBatch.Visible = False
            Form_Resize
        Else
            If vsBatch.Visible = False Then
                If mbln��ʾ���� = True Then '�˲��������ܲ�����ʾ�����б������첻��ȷ����ģʽ
                    vsBatch.Visible = True
                End If
            Else
                If mbln��ʾ���� = False Then '�˲��������ܲ�����ʾ�����б������첻��ȷ����ģʽ
                    vsBatch.Visible = False
                End If
            End If
        End If
        Form_Resize
        
        With mrsStock
            If .State = 1 Then .Close
            gstrSQL = ""
            If mbln������ Then
                gstrSQL = "Select " & IIf(mstr�̵�ʱ�� <> "", "/*+ Rule*/", "") & " 1 RID,���� �ⷿ,0 ����,'����������������' ����,sysdate ʧЧ��," & _
                          "to_char(0," & gOraFmt_Max.FM_���� & ") ��������,to_char(0," & gOraFmt_Max.FM_���� & ") �������,to_char(0," & gOraFmt_Max.FM_��� & ") �����,to_char(0," & gOraFmt_Max.FM_��� & ") �����,sysdate as ���ʧЧ��,to_char(0," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,'' As �ɱ���,to_char(0," & gOraFmt_Max.FM_�ɱ��� & ") �ϴι���,'' ���� , Sysdate As ��������, 0 As �ϴι�Ӧ��id,'' ��׼�ĺ�,'' ��Ʒ����,'' �ڲ����� " & _
                          " From ���ű�" & _
                          " Where ID=[1]" & _
                          " Union "
            End If
             
            gstrSQL = gstrSQL & " Select " & IIf(mstr�̵�ʱ�� <> "", "/*+ Rule*/", "") & " 2 RID,P.���� �ⷿ,K.����,K.�ϴ����� ����,K.Ч�� ʧЧ��,"
            If mblnStock Then
                If mblnɢװ��λ Then
                    strTmp = " to_char( K.��������," & gOraFmt_Max.FM_���� & ") ��������," & _
                             " to_char( K.ʵ������," & gOraFmt_Max.FM_���� & ") as �������,"
                Else
                    strTmp = " to_char( K.��������" & mstrUnitString & "," & gOraFmt_Max.FM_���� & ") ��������," & _
                             " to_char( K.ʵ������" & mstrUnitString & "," & gOraFmt_Max.FM_���� & ") �������,"
                End If
            Else
                strTmp = "to_char( ''," & gOraFmt_Max.FM_���� & ") ��������,to_char( ''," & gOraFmt_Max.FM_���� & ") �������,"
            End If
            
            
            'ȡ���
            '20060731:���˺���룬��Ҫ����̵�ʱ��Ŀ��
            strKc = "" & _
                "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����,a.�ϴι�Ӧ��ID, a.�ϴβɹ���,A.���ۼ�,ƽ���ɱ���," & _
                "           a.ʵ������,a.ʵ�ʽ��, a.ʵ�ʲ��, a.��������,a.�ϴ�����,a.�ϴβ���,a.Ч��,a.���Ч��,a.�ϴ���������,a.��׼�ĺ�,a.��Ʒ����,a.�ڲ����� " & _
                "   FROM ҩƷ��� a " & _
                "   Where a.ҩƷid = [3]" & _
                "           AND a.����=1 " & _
                "           AND a.�ⷿid+0 = "
            If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
                strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[1]", "[2]")
            End If
            
            If mstr�̵�ʱ�� <> "" Then
                strKc = strKc & _
                    "   UNION ALL " & _
                    "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����, a.��ҩ��λID �ϴι�Ӧ��ID,max(a.�ɱ���) �ϴβɹ���,max(A.���ۼ�) as ���ۼ�,0 as ƽ���ɱ���, " & _
                    "           -SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ʵ������, " & _
                    "           -SUM (DECODE (a.���ϵ��, 1, a.���۽��, -a.���۽��)) AS ʵ�ʽ��," & _
                    "           -SUM (DECODE (a.���ϵ��, 1, a.���, -a.���)) AS ʵ�ʲ��,-SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����))  AS ��������,a.����,a.���� , A.Ч��,a.���Ч��,a.��������,a.��׼�ĺ�,a.��Ʒ����,a.�ڲ����� " & _
                    "   FROM ҩƷ�շ���¼ a " & _
                    "   Where a.ҩƷid+0=[3]  " & _
                    "           AND a.�ⷿid + 0 ="
                If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
                    strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[1]", "[2]")
                End If
                strKc = strKc & " AND a.������� >[5] " & _
                    " GROUP BY A.�ⷿid, a.ҩƷid,a.��ҩ��λid, A.����, A.����, A.����, A.Ч��, A.���Ч��,a.��������,a.��׼�ĺ�,a.��Ʒ����,a.�ڲ�����"
            End If
                      
            strKc = "" & _
                "   Select �ⷿid,ҩƷid,nvl(����,0) ����,max(�ϴ�����) �ϴ�����,min(���Ч��) as ���ʧЧ��,max(�ϴι�Ӧ��ID) �ϴι�Ӧ��ID, " & _
                "       Sum(nvl(��������,0)) ��������," & _
                "       Sum(ʵ������) ʵ������," & _
                "       Sum(ʵ�ʽ��) ʵ�ʽ��," & _
                "       Sum(ʵ�ʲ��) ʵ�ʲ��," & _
                "       max(�ϴβɹ���) �ϴβɹ���,max(���ۼ�) as ���ۼ�,max(ƽ���ɱ���) as ƽ���ɱ���, " & _
                "        Min(���Ч��) ���Ч��,Min(Ч��) Ч��,max(�ϴβ���) �ϴβ��� ,max(�ϴ���������) �ϴ���������,max(��׼�ĺ�) as ��׼�ĺ�,max(��Ʒ����) as ��Ʒ����,max(�ڲ�����) as �ڲ�����,1 As ����" & _
                "   From (" & strKc & ")" & _
                "   Group by �ⷿid,ҩƷid,nvl(����,0) "
             
            gstrSQL = gstrSQL & strTmp & _
                     IIf(mblnStock, "to_char(K.ʵ�ʽ��," & gOraFmt_Max.FM_��� & ")  as �����,", "to_char(0," & gOraFmt_Max.FM_��� & ")  �����,") & _
                     IIf(mblnStock, " to_char(K.ʵ�ʲ��," & gOraFmt_Max.FM_��� & ")  as �����", "to_char(0," & gOraFmt_Max.FM_��� & ")  �����") & " ,K.���Ч�� as ���ʧЧ��," & _
                     IIf(mblnStock, "to_char(Decode(nvl(M.�Ƿ���,0),0,G.�ּ�,decode(nvl(K.���ۼ�,0),0,nvl(K.ʵ�ʽ��,0)/decode(K.ʵ������,null,1,0,1,K.ʵ������),nvl(K.���ۼ�,0)))" & IIf(mblnɢװ��λ, "", "*nvl(D.����ϵ��,1)") & "," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,", "to_char(0," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,") & _
            " to_char(k.ƽ���ɱ���," & gOraFmt_Max.FM_�ɱ��� & ") as �ɱ���, " & _
                     IIf(mblnStock, "to_char(decode(nvl(K.�ϴβɹ���,0),0,(nvl(K.ʵ�ʽ��,0)-nvl(K.ʵ�ʲ��,0))/decode(K.ʵ������,null,1,0,1,K.ʵ������),K.�ϴβɹ���)" & IIf(mblnɢװ��λ, "", "*nvl(D.����ϵ��,1)") & "," & gOraFmt_Max.FM_�ɱ��� & ") �ϴι���", "to_char(0," & gOraFmt_Max.FM_�ɱ��� & ") �ϴι���") & _
            "        ,K.�ϴβ��� ����,k.�ϴ��������� �������� ,k.�ϴι�Ӧ��ID,k.��׼�ĺ�,k.��Ʒ����,k.�ڲ����� " & _
            " From ���ű� P,�������� D," & IIf(mstr�̵�ʱ�� <> "", "(" & strKc & ")", " ҩƷ���") & " K,�շ���ĿĿ¼ M,�շѼ�Ŀ G " & _
            " Where K.�ⷿID = P.ID And D.����ID = K.ҩƷID And K.�ⷿID " & IIf(mstr�̵�ʱ�� <> "", " +0=", "=") & IIf(mlngԴ�ⷿID = 0, "[1]", "[2]") & _
            " And K.ҩƷID " & IIf(mstr�̵�ʱ�� <> "", " +0=", "=") & " [3]  And K.����=1 " & _
            " And D.����id=G.�շ�ϸĿID(+) " & _
            " And D.����ID=M.ID And (M.վ��=[7] or M.վ�� is null) " & _
            " And m.Id = g.�շ�ϸĿid And (Sysdate Between g.ִ������ And Nvl(g.��ֹ����, Sysdate)) " & _
            GetPriceClassString("G")
            
            Dim dtDate As Date
            If mstr�̵�ʱ�� <> "" Then
                dtDate = CDate(mstr�̵�ʱ��)
            Else
                dtDate = Now
            End If
                     
            If mbln�̵㵥 Then
                gstrSQL = gstrSQL & " And (K.ʵ������<>0 Or K.ʵ�ʽ��<>0 Or K.ʵ�ʲ��<>0)"
            Else
                gstrSQL = gstrSQL & " And K.ʵ������<>0 "
            End If
            
            If mstrCode <> "" Then
                gstrSQL = gstrSQL & " And (K.��Ʒ����=[6] Or K.�ڲ�����=[6]) "
            End If
             
            ' If mlng��Ӧ��ID <> 0 Then gstrSQL = gstrSQL & " And K.�ϴι�Ӧ��ID=[4]"
             
            If gSystem_Para.P156_�����㷨 = 0 Then
                gstrSQL = gstrSQL & " Order by RID,����"
            Else
                gstrSQL = gstrSQL & " Order by RID,ʧЧ��,����"
            End If
            
            Set mrsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngĿ�ⷿID, IIf(mlngԴ�ⷿID = 0, mlngĿ�ⷿID, mlngԴ�ⷿID), _
                           mlngLastSelect����ID, mlng��Ӧ��ID, dtDate, mstrCode, gstrNodeNo)
        End With
        Dim blnState As Boolean
           
        With vsBatch
            .Redraw = flexRDNone
            Set .DataSource = mrsStock
            If mrsStock.EOF Then
                .Clear 1
                .Rows = 2
            End If
            Call SetFormat(0, mrsStock.EOF)
            
            .Redraw = flexRDBuffered
            If mbln������ And mrsStock.RecordCount <> 0 Then
                If .Rows >= 3 Then .Row = 2
                If .Rows = 2 Then .Row = 1
            End If
            blnState = ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And Not mrsStock.EOF        '��������Ĳ�����
            If mbln��ʾ���� = True And blnState = True Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        Form_Resize
    End If
    
    '���ð�ť״̬
    With mrsCard
        If .RecordCount <> 0 Then .MoveFirst
        .Find "����ID=" & Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("����ID")))
        If .EOF Then
            MsgBox "�����ڲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        'mint�ⷿ:1-���Ŀ�;2-����;3-�Ƽ���
        'mint����:0-������;1-�ⷿ����;2-���÷���;3-���Ŀ����÷���
        If In_�༭״̬ = 2 And ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And mblnPrice Then
            If mbln��ʾ���� = False Then
                mblnSelect = True
            Else
                mblnSelect = blnState
            End If
        Else
            mblnSelect = True
        End If
    End With
    'Call ReSetWindowsFormLocal
    
    With vsBatch
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "��ʾ�Է����") = 0 And vsBatch.Visible = True Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("�������"))) > 0 Then
                        .TextMatrix(i, .ColIndex("�������")) = "��"
                    Else
                        .TextMatrix(i, .ColIndex("�������")) = "��"
                    End If
                    .TextMatrix(i, .ColIndex("�����")) = ""
                    .TextMatrix(i, .ColIndex("�����")) = ""
                Next
            End If
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsHead_GotFocus()
    With vsHead
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
        .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsHead_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    If KeyCode = vbKeyReturn Then vsHead_DblClick: Exit Sub
    
    With vsHead
        Select Case KeyCode
            Case vbKeyRight
                If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                    .LeftCol = .LeftCol + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyLeft
                If .LeftCol <> 0 Then
                    .LeftCol = .LeftCol - 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyHome
                If .LeftCol <> 0 Then
                    .LeftCol = 0
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                End If
            Case vbKeyEnd
                For i = .Cols - 1 To 0 Step -1
                    sngWidth = sngWidth + .ColWidth(i)
                    If sngWidth > .Width Then
                        .LeftCol = i + 1
                        .Col = .LeftCol
                        .ColSel = .Cols - 1
                        Exit For
                    End If
                Next
        End Select
    End With
    
End Sub

Private Sub vsHead_LostFocus()
    With vsHead
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
        .BackColorSel = &H8000000A
    End With
End Sub

Private Function RefreshData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������������������,����������,����ͷ����
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, StrGroupBy As String
    Dim strLike As String
    Dim strSerach As String
    Dim strKc As String
    Dim strInput As String
    Dim rsTmp As ADODB.Recordset
    Dim blnVirtualStock As Boolean
    Dim strCode As String
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim int������ȡֵ��ʽ As Integer
    Dim lng����ID As Long
    
    On Error GoTo ErrHandle

    RefreshData = False
    
    If mlngModule = 1712 Or mlngModule = 1714 Then
        int������ȡֵ��ʽ = Val(zlDatabase.GetPara(268, glngSys))
    End If
    
    '�Ȱ��������
    If gblnCode = True Then
        gstrSQL = "Select ҩƷid From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And (��Ʒ���� = [2] Or �ڲ����� = [2])"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "RefreshData", mlngԴ�ⷿID, UCase(mstrInput))
        If Not rsData.EOF Then
            mstrCode = UCase(mstrInput)
            mstrInput = UCase(mstrInput)
            lng����ID = rsData!ҩƷid
        Else
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "RefreshData", mlngԴ�ⷿID, mstrInput)
            If Not rsData.EOF Then
                mstrCode = mstrInput
                lng����ID = rsData!ҩƷid
            Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "RefreshData", mlngԴ�ⷿID, LCase(mstrInput))
                If Not rsData.EOF Then
                    mstrCode = LCase(mstrInput)
                    mstrInput = LCase(mstrInput)
                    lng����ID = rsData!ҩƷid
                End If
            End If
        End If
    End If
    
    mblnֻ��ʾ�������� = �ж�ֻ�߱����ϲ���(mlngĿ�ⷿID)
    
    '�ж�����ⷿ
    gstrSQL = "select count(*) rec from ��������˵�� where ��������='����ⷿ' and ����id=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж�����ⷿ", mlngĿ�ⷿID)
    If rsTmp!rec = 1 And (mobjOut.Name = "frmPurchaseCard" Or mobjOut.Name = "frmOtherInputCard") Then blnVirtualStock = True
    
    strLike = "" & GetMatchingSting(UCase(mstrInput), False) & ""
        
    '����ƥ�䣺ƥ����룬���룬���ƣ���Ʒ���루�̶���ƥ�䣩���ڲ����루�̶���ƥ�䣩
    strSerach = " And (A.���� Like [4] OR B.���� Like [4] OR ( B.���� LIKE [4] and B.����=[6]))"
    If IsNumeric(mstrInput) Then                         '���������,��ֻȡ����
        If Mid(gSystem_Para.Para_���뷽ʽ, 1, 1) = "1" Then strSerach = " And (A.���� Like [4] And B.����=[6])"
    ElseIf zlStr.IsCharAlpha(mstrInput) Then          '����ȫ����ĸʱֻƥ�����
        If Mid(gSystem_Para.Para_���뷽ʽ, 2, 1) = "1" Then strSerach = " And B.���� Like [4] And B.����=[6] "
    ElseIf zlStr.IsCharChinese(mstrInput) Then
        strSerach = " And B.���� Like [4] And B.����=[6] "
    End If
    
    strInput = " Select a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.�������, a.�Ƿ��� " & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
            " Where A.ID=B.�շ�ϸĿID And (A.վ��=[8] or A.վ�� is null) AND A.��� ='4' And (A.����ʱ�� is null Or A.����ʱ��>[5]) " & strSerach & _
            " Union All " & _
            " Select a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.�������, a.�Ƿ��� " & _
            " From �շ���ĿĿ¼ A, ҩƷ��� B " & _
            " Where a.Id = b.ҩƷid And ���� = 1 And �ⷿid + 0 = [1] And (��Ʒ���� = [7] Or �ڲ����� = [7]) "
    
    '������������;���ࡢ����ָ���������
    '--��Ϊ����������ƻ���룬���˷�ʽ����ָ���������--
    '����ͷ��˳��
    gstrSQL = "" & _
    " Select  D.����id,D.����id,D.����ID,D.����,D.ͨ������,D.��Ʒ��,D.���,D.����,D.��׼�ĺ�,d.ע��֤��,x.���� As �ϴι�Ӧ��,to_char(D.�ۼ�," & gOraFmt_Max.FM_���ۼ� & ") �ۼ�,to_char(d.�ɱ���," & gOraFmt_Max.FM_�ɱ��� & ") as ���³ɱ���,D.ɢװ��λ ɢװ,D.����ϵ�� ϵ��,D.��װ��λ ��װ," & _
                 IIf(mblnStock, "to_char(S.�������� " & IIf(mblnɢװ��λ, "", "/D.����ϵ��") & "," & gOraFmt_Max.FM_���� & ") ��������, " & _
                                "to_char(S.������� " & IIf(mblnɢװ��λ, "", "/D.����ϵ��") & "," & gOraFmt_Max.FM_���� & ") �������, " & _
                                "to_char(S.�����," & gOraFmt_Max.FM_��� & ") �����,to_char(S.�����," & gOraFmt_Max.FM_��� & ") �����,", _
                      "to_char(''," & gOraFmt_Max.FM_���� & ") ��������,to_char(''," & gOraFmt_Max.FM_���� & ") �������,to_char(''," & gOraFmt_Max.FM_��� & ") �����,to_char(''," & gOraFmt_Max.FM_��� & ") �����,") & _
    "           D.���Ч�� ��Ч��,D.���Ч��,S.���ʧЧ��,D.һ���Բ���,D.�޾��Բ���,D.�ⷿ����,D.���÷���,D.ʱ��,to_char(D.ָ��������," & gOraFmt_Max.FM_���ۼ� & ") ָ��������,D.ָ�������,E.�ⷿ��λ " & _
    " From "
   
    
    '������Ϣ������Ŀ¼
    If mblnֻ��ʾ�������� Then
        gstrSQL = gstrSQL & " (Select Distinct u.����id,u.����id,H.����ID,V.����,v.���� As ͨ������,B.���� As ��Ʒ��,V.���," & IIf(int������ȡֵ��ʽ = 0, "decode(u.�ϴβ���,null,v.����,u.�ϴβ���)", "decode(v.����,null,u.�ϴβ���,v.����)") & " as ����,u.��׼�ĺ�,u.ע��֤��,V.���㵥λ as ɢװ��λ,U.��װ��λ," & _
                    "                       To_Char(U.����ϵ��," & GFM_XS & " ) ����ϵ��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��," & _
                    "                       Decode(U.�ⷿ����,1,'��','��') �ⷿ����,Decode(U.���÷���,1,'��','��') ���÷���,Decode(U.һ���Բ���,1,'��','��')  һ���Բ���,Decode(U.�޾��Բ���,1,'��','��') �޾��Բ���,Decode(V.�Ƿ���,1,'��','��') ʱ��," & _
                    "                       U.ָ��������,To_Char(U.ָ�������," & GFM_CJL & " ) ָ�������,�ּ� �ۼ�,u.�ɱ���,Nvl(u.�ϴι�Ӧ��id, 0) As �ϴι�Ӧ��id " & _
                    "               From �������� U, " & _
                    "                    ( " & strInput & ") V," & _
                    "                    ������ĿĿ¼ H, " & _
                    "                   (SELECT �շ�ϸĿid, ִ�п���id FROM �շ�ִ�п��� WHERE ִ�п���ID" & IIf(mlngԴ�ⷿID <> 0, "+0=[1]", IIf(mlngĿ�ⷿID <> 0, "+0=[2]", " Is Not NULL")) & ") K," & _
                    "                   (Select �շ�ϸĿID, ִ�п���ID From �շ�ִ�п��� Where ִ�п���ID" & IIf(mlngĿ�ⷿID <> 0, "+0=[2]", IIf(mlngԴ�ⷿID <> 0, "+0=[1]", " Is Not NULL")) & " ) i," & _
                    "               �շ���Ŀ���� B, �շѼ�Ŀ P " & _
                    "               where U.����id=v.id and U.����id=H.id And V.ID = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                    "                       AND U.����id=K.�շ�ϸĿID  " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                    "                       And U.����id=i.�շ�ϸĿId " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                                            IIf(mblnֻ��ʾ��������, " And U.��������=1 ", IIf(mblnTrackUsing = True, " and  U.�������� =0 ", "")) & " And v.Id = p.�շ�ϸĿid And (Sysdate Between p.ִ������ And Nvl(p.��ֹ����, Sysdate)) " & _
                                            GetPriceClassString("P")
    Else
        gstrSQL = gstrSQL & " (Select Distinct u.����id,u.����id,H.����ID,V.����,v.���� As ͨ������,B.���� As ��Ʒ��,V.���," & IIf(int������ȡֵ��ʽ = 0, "decode(u.�ϴβ���,null,v.����,u.�ϴβ���)", "decode(v.����,null,u.�ϴβ���,v.����)") & " as ����,u.��׼�ĺ�,u.ע��֤��,V.���㵥λ as ɢװ��λ,U.��װ��λ," & _
                    "                       To_Char(U.����ϵ��," & GFM_XS & " ) ����ϵ��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��,nvl(To_Char(U.���Ч��,'9999990'),0) ���Ч��," & _
                    "                       Decode(U.�ⷿ����,1,'��','��') �ⷿ����,Decode(U.���÷���,1,'��','��') ���÷���,Decode(U.һ���Բ���,1,'��','��')  һ���Բ���,Decode(U.�޾��Բ���,1,'��','��') �޾��Բ���,Decode(V.�Ƿ���,1,'��','��') ʱ��," & _
                    "                       U.ָ�������� ,To_Char(U.ָ�������," & GFM_CJL & " ) ָ�������,�ּ� �ۼ�,u.�ɱ���,Nvl(u.�ϴι�Ӧ��id, 0) As �ϴι�Ӧ��id " & _
                    "               From �������� U," & _
                    "                    (" & strInput & ") V," & _
                    "                   ������ĿĿ¼ H,�շ���Ŀ���� B, �շѼ�Ŀ P," & _
                    "                   (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID" & IIf(mlngԴ�ⷿID <> 0, "=[1]", IIf(mlngĿ�ⷿID <> 0, "=[2]", " Is Not NULL")) & " ) K," & _
                    "                   (Select ִ�п���ID,�շ�ϸĿID From �շ�ִ�п��� Where ִ�п���ID" & IIf(mlngĿ�ⷿID <> 0, "+0=[2]", IIf(mlngԴ�ⷿID <> 0, "+0=[1]", " Is Not NULL")) & " ) i" & _
                    "               where U.����id=v.id and U.����id=H.id And V.ID = B.�շ�ϸĿid(+) And B.����(+) = 3 " & _
                    "               AND U.����id=K.�շ�ϸĿID " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                    "               And U.����id=I.�շ�ϸĿID  " & IIf(mbln���޴洢�ⷿ����, "(+)", "") & _
                                    IIf(mblnֻ��ʾ��������, " And U.��������=1 ", IIf(mblnTrackUsing = True, " and  U.�������� =0 ", "")) & " And v.Id = p.�շ�ϸĿid And (Sysdate Between p.ִ������ And Nvl(p.��ֹ����, Sysdate)) " & _
                                    GetPriceClassString("P")

    End If
    
'    'ֻ����δͣ�õĹ������
'    If mstr�̵�ʱ�� <> "" Then      '���̵�ʱ����˵������̵�ʱ��С��ͣ�õ�ʱ��ҲӦ����ʾ����
'        gstrSQL = gstrSQL & " And (V.����ʱ�� Is Null Or V.����ʱ��>[5]) "
'    Else
'        gstrSQL = gstrSQL & " And (V.����ʱ�� Is Null Or To_char(V.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
'    End If
'
    
    gstrSQL = gstrSQL & IIf(blnVirtualStock, " And nvl(u.��ֵ����,0)=1 and nvl(u.���ٲ���,0)=1 and nvl(u.��������,0)=1 and nvl(u.���÷���,0)=1", "")

    If mlngĿ�ⷿID > 0 Then
        gstrSQL = gstrSQL & " And " & _
            "     ( exists(select 1 from ��������˵�� where �������� In ('�Ƽ���', '���Ŀ�', '���ϲ���', '����ⷿ')  and ����id=[2]" & ")  " & _
            "       or v.�������=(select distinct '1' from ��������˵�� where �������� like '���ϲ���' and ����id=[2] and ������� in(1,3))" & _
            "       or v.�������=(select distinct '2' from ��������˵�� where �������� like '���ϲ���' and ����id=[2] and ������� in(2,3)))"
    End If
    
    'ֻ����ָ�����ʷ���Ĺ�����
    gstrSQL = gstrSQL & " ) D,"

    'ȡ���
    '20060731:���˺���룬��Ҫ����̵�ʱ��Ŀ��
    strKc = "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����,a.�ϴι�Ӧ��ID," & _
        "           a.ʵ������,a.ʵ�ʽ��, a.ʵ�ʲ��, a.��������,a.�ϴ�����,a.�ϴβ���,a.Ч��,a.���Ч�� " & _
        "   FROM ҩƷ��� a " & _
        "   Where a.����=1 AND a.�ⷿid = "
    If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
        strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[2]", "[1]")
    End If
    
    '�̵�ʱ�����̵�ʱ������̵�ʱ�䵽��ǰʱ��ķ�����
    If mstr�̵�ʱ�� <> "" Then
        strKc = strKc & _
            "   UNION ALL " & _
            "   SELECT a.�ⷿid, a.ҩƷid, NVL (a.����, 0) AS ����, a.��ҩ��λID �ϴι�Ӧ��ID, " & _
            "           -SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ʵ������, " & _
            "           -SUM (DECODE (a.���ϵ��, 1, a.���۽��, -a.���۽��)) AS ʵ�ʽ��," & _
            "           -SUM (DECODE (a.���ϵ��, 1, a.���, -a.���)) AS ʵ�ʲ��,-SUM (DECODE (a.���ϵ��, 1, a.ʵ������*a.����, -a.ʵ������*a.����)) AS ��������,a.����,a.���� , A.Ч��,a.���Ч��" & _
            "   FROM ҩƷ�շ���¼ a " & _
            "   Where a.�ⷿid + 0 ="
        If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
            strKc = strKc & IIf(mlngԴ�ⷿID = 0, "[2]", "[1]")
        End If
        strKc = strKc & " AND a.������� >[5] " & _
            " GROUP BY A.�ⷿid, a.ҩƷid,a.��ҩ��λid, A.����, A.����, A.����, A.Ч��, A.���Ч�� "
    End If

    If mblnStock Then
        gstrSQL = gstrSQL & " (Select ҩƷid as ����id,min(���Ч��) as ���ʧЧ�� , Sum(��������) ��������," & _
                " Sum(ʵ������) �������," & _
                " Sum(ʵ�ʽ��)  �����," & _
                " Sum(ʵ�ʲ��) �����"
    Else
        gstrSQL = gstrSQL & " (Select ҩƷid as ����id,min(���Ч��) as ���ʧЧ��, 0 ��������," & _
                " 0 �������,0 �����,0 �����"
    End If
    If mstr�̵�ʱ�� <> "" Then
         gstrSQL = gstrSQL & " From (" & strKc & ") where 1=1 "
    Else
         gstrSQL = gstrSQL & " From ҩƷ��� Where ����=1 "
    End If
    
    
    'If mlng��Ӧ��ID <> 0 Then gstrSQL = gstrSQL & " And �ϴι�Ӧ��ID=[3]"
    
    If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
        gstrSQL = gstrSQL & " And �ⷿID" & IIf(mstr�̵�ʱ�� <> "", " +0 =", "=") & IIf(mlngԴ�ⷿID = 0, "[2]", "[1]") & "  Group By ҩƷid) S"
    Else
        gstrSQL = gstrSQL & " Group By ҩƷid) S"
    End If
    gstrSQL = gstrSQL & ",(Select ����id,�ⷿID,�ⷿ��λ From ���ϴ����޶� " & _
              " Where �ⷿID=" & IIf(mintEditState = 2, "[1]", "[2]") & ") E,��Ӧ�� X"
    
    '������
    gstrSQL = gstrSQL & " Where D.����ID=S.����ID"
    
    If mbln����ʾ������� And mblnStock Then
        gstrSQL = gstrSQL & " And S.��������<>0"
    Else
        '��ϵͳ���������ĳ������顱Ϊ�����ֹʱ��������Ϊ��
        If Not (mintStockCheck = 2 And In_�༭״̬ = 2) Or mbln�̵㵥 Or Not mblnCheck Then gstrSQL = gstrSQL & "(+) "
        'If In_�༭״̬ = 2 Then gstrSQL = gstrSQL & " And S.��������<>0"
    End If
    gstrSQL = gstrSQL & " And D.����ID=E.����ID(+)  And d.�ϴι�Ӧ��id = x.Id(+) Order By D.����"
        
    Dim dtDate As Date
    If mstr�̵�ʱ�� <> "" Then
        dtDate = CDate(mstr�̵�ʱ��)
    Else
        dtDate = CDate("2999-12-31")
    End If

    Set mrsCard = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngԴ�ⷿID, mlngĿ�ⷿID, mlng��Ӧ��ID, _
                    strLike, dtDate, gSystem_Para.int���뷽ʽ + 1, mstrInput, gstrNodeNo)
    
    If lng����ID = 0 Then
        gstrSQL = "Select distinct a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.�������, a.�Ƿ���" & vbNewLine & _
                    "From �շ���ĿĿ¼ A, �շ���Ŀ���� B" & vbNewLine & _
                    "Where a.Id = b.�շ�ϸĿid And (a.վ�� =[3] Or a.վ�� Is Null) And a.��� = '4' And" & vbNewLine & _
                    "      (a.����ʱ�� Is Null Or a.����ʱ�� > To_Date('2999-12-31 00:00:00', 'YYYY-MM-DD HH24:MI:SS')) And (b.���� Like [1] or a.���� Like [1] or b.���� Like [1])"
    Else
        gstrSQL = "Select distinct a.Id, a.����, a.����, a.���, a.����, a.���㵥λ, a.�������, a.�Ƿ���" & vbNewLine & _
                    "From �շ���ĿĿ¼ A, �շ���Ŀ���� B" & vbNewLine & _
                    "Where a.Id = b.�շ�ϸĿid And (a.վ�� =[3] Or a.վ�� Is Null) And a.��� = '4' And" & vbNewLine & _
                    "      (a.����ʱ�� Is Null Or a.����ʱ�� > To_Date('2999-12-31 00:00:00', 'YYYY-MM-DD HH24:MI:SS')) And a.id = [2] "
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ�������Ĳ�ѯ", "%" & UCase(mstrInput) & "%", lng����ID, gstrNodeNo)
    
    If In_�༭״̬ = 2 Then
        '����
        If rsTmp.RecordCount = 0 Then
            MsgBox "�޴����ģ����������룡", vbInformation, gstrSysName
            Exit Function
        ElseIf rsTmp.RecordCount > 0 And mrsCard.RecordCount = 0 Then
            If blnVirtualStock = False Then
                MsgBox "�������޿�棡", vbInformation, gstrSysName
            Else
                MsgBox "�������޿���" & _
                vbCrLf & "����ⷿ��ͨ��Ҫ���ľ��и�ֵ���ϡ����ٲ��ˡ��������á����÷������ԣ����飡", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    Else
        '���
        If rsTmp.RecordCount = 0 Then
            MsgBox "�޴����ģ����������룡", vbInformation, gstrSysName
            Exit Function
        ElseIf mrsCard.RecordCount = 0 Then
            MsgBox "δ�ҵ��������������ģ�������δ���ô洢�ⷿ������", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With vsHead
        Set .DataSource = mrsCard
        If mrsCard.EOF Then
            .Rows = 2
        End If
        
        Call SetFormat(1, mrsCard.EOF)
        mblnSelect = (mrsCard.EOF <> True)
    End With
    
    With vsHead
        If glngModul = 1717 Then
            If InStr(1, gstrPrivs, "��ʾ�Է����") = 0 Then
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("�������"))) > 0 Then
                        .TextMatrix(i, .ColIndex("�������")) = "��"
                    Else
                        .TextMatrix(i, .ColIndex("�������")) = "��"
                    End If
                    .TextMatrix(i, .ColIndex("�����")) = ""
                    .TextMatrix(i, .ColIndex("�����")) = ""
                Next
            End If
        End If
    End With
    
    Call vsHead_EnterCell
    RefreshData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitRec()
    '----------------------------------------------------------------------------------------
    '����:�����������ݼ��ṹ
    '----------------------------------------------------------------------------------------
    Set mrsReturn = New ADODB.Recordset
    With mrsReturn
        If .State = 1 Then .Close
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҩ��λID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 82, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "�ۼ�", adDouble, 18, adFldIsNullable
        .Fields.Append "ɢװ��λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 11, adFldIsNullable
        .Fields.Append "��װ��λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "���Ч��", adDouble, 5, adFldIsNullable
        .Fields.Append "���Ч��", adDouble, 5, adFldIsNullable
        
        .Fields.Append "һ���Բ���", adDouble, 2, adFldIsNullable
        .Fields.Append "�޾��Բ���", adDouble, 2, adFldIsNullable
        .Fields.Append "�ⷿ����", adDouble, 2, adFldIsNullable
        .Fields.Append "���÷���", adDouble, 2, adFldIsNullable
        .Fields.Append "��׼�ĺ�", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        
        .Fields.Append "ʱ��", adDouble, 2, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "Ч��", adDate, , adFldIsNullable
        .Fields.Append "���ʧЧ��", adDate, , adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ʵ������", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ʵ�ʽ��", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ʵ�ʲ��", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ָ��������", adDouble, 11, adFldIsNullable
        .Fields.Append "ָ�������", adDouble, 11, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function

Private Function CombinateRec() As Boolean
    '��װ��¼��
    '��λ��¼��
    Dim blnEof As Boolean                   '�Ƿ���ڿ�����μ�¼
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    CombinateRec = False
    
    If chkContinue.Value = 0 Then '��װһ������
        With mrsCard
            If .RecordCount <> 0 Then .MoveFirst
            .Find "����ID=" & Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("����ID")))
            If .EOF Then
                MsgBox "�����ڲ�����", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mbln��ʾ���� = True Then 'ֻ����ʾ���ε�����²���Ҫ�����²���
                If ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) And In_�༭״̬ = 2 Then
                    With mrsStock
                        If .RecordCount <> 0 Then .MoveFirst
                        .Find "����=" & Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("����")))
                        If .EOF Then
                            blnEof = True
                            If mblnPrice Then
                                MsgBox "�����ڲ�����", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End With
                End If
            End If
        End With
        
        'װ����д���¼��������������ʹ��
        With mrsReturn
            If .EOF Then .AddNew
            !����ID = mrsCard!����ID
            !����id = mrsCard!����id
            !����id = mrsCard!����id
            
            !���� = mrsCard!����
            !���� = zlStr.Nvl(mrsCard!ͨ������)
            !��Ʒ�� = zlStr.Nvl(mrsCard!��Ʒ��)
            !��� = mrsCard!���
            !���� = mrsCard!����
            !�ۼ� = zlStr.Nvl(mrsCard!�ۼ�, 0)
            !ɢװ��λ = mrsCard!ɢװ
            !����ϵ�� = mrsCard!ϵ��
            !��װ��λ = mrsCard!��װ
            !���Ч�� = mrsCard!��Ч��
            !���Ч�� = mrsCard!���Ч��
            !һ���Բ��� = IIf(mrsCard!һ���Բ��� = "��", 1, 0)
            !�޾��Բ��� = IIf(mrsCard!�޾��Բ��� = "��", 1, 0)
            !�ⷿ���� = IIf(mrsCard!�ⷿ���� = "��", 1, 0)
            !���÷��� = IIf(mrsCard!���÷��� = "��", 1, 0)
            !���ʧЧ�� = mrsCard!���ʧЧ��
              
            !ʱ�� = IIf(mrsCard!ʱ�� = "��", 1, 0)
            
            '�����ҷ���
            If In_�༭״̬ = 2 And ((mint���� = 3 And mint�ⷿ <> 3) Or (mint���� = 1 And mint�ⷿ = 1) Or (mint���� = 2 And mint�ⷿ = 2)) Then
                If mbln��ʾ���� = True Then
                    If vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("����")) = "����������������" Then
                        !���� = -1
                    Else
                        If Not blnEof Then
                            !���� = zlStr.Nvl(mrsStock!����)
                            !�������� = mrsStock!��������
                            !��׼�ĺ� = zlStr.Nvl(mrsStock!��׼�ĺ�)
                            !��ҩ��λID = mrsStock!�ϴι�Ӧ��id
                            !���� = Val(zlStr.Nvl(mrsStock!����))
                            !���� = zlStr.Nvl(mrsStock!����)
                            !Ч�� = mrsStock!ʧЧ��
                            !���ʧЧ�� = mrsStock!���ʧЧ��
                            !�������� = IIf(IsNull(mrsStock!��������), 0, mrsStock!��������)
                            !ʵ������ = IIf(IsNull(mrsStock!�������), 0, mrsStock!�������)
                            !ʵ�ʽ�� = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                            !ʵ�ʲ�� = IIf(IsNull(mrsStock!�����), 0, mrsStock!�����)
                            
                            If Not mblnStock Then Call Get���ÿ��(!����ID, !����)
                        End If
                    End If
                Else
                    If Not mblnStock Then Call Get���ÿ��(mrsCard!����ID, 0)
                End If
            Else
                '���򲻷���
                !�������� = IIf(IsNull(mrsCard!��������), 0, mrsCard!��������)
                !ʵ������ = IIf(IsNull(mrsCard!�������), 0, mrsCard!�������)
                !ʵ�ʽ�� = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
                !ʵ�ʲ�� = IIf(IsNull(mrsCard!�����), 0, mrsCard!�����)
                If In_�༭״̬ = 1 Then
                    !��׼�ĺ� = zlStr.Nvl(mrsCard!��׼�ĺ�)
                Else
                    If mrsStock.RecordCount > 0 Then
                        mrsStock.MoveFirst
                        !��׼�ĺ� = zlStr.Nvl(mrsStock!��׼�ĺ�)
                    Else
                        !��׼�ĺ� = ""
                    End If
                End If
                
                If Not mblnStock Then Call Get���ÿ��(!����ID, 0)
            End If
            
            '�������ʾ�Է��ⷿ�Ŀ�棬��������ȡ������
            If Not mblnStock Then
                !�������� = msin��������
                !ʵ������ = msinʵ������
                !ʵ�ʽ�� = msinʵ�ʽ��
                !ʵ�ʲ�� = msinʵ�ʲ��
            End If
            
            !ָ�������� = mrsCard!ָ��������
            !ָ������� = mrsCard!ָ�������
            .Update
        End With
    Else '��װ��������
        With mrsReturn
            For i = 1 To vsfѡ��.Rows - 2
                .AddNew
                !����ID = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����ID")))
                !����id = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����id")))
                !����id = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����id")))
                
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !��Ʒ�� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��Ʒ��"))
                !��� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���"))
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !�ۼ� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("�ۼ�")))
                !ɢװ��λ = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ɢװ��λ"))
                !����ϵ�� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����ϵ��")))
                !��װ��λ = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��װ��λ"))
                !���Ч�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���Ч��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���Ч��")))
                !���Ч�� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���Ч��")))
                !���ʧЧ�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���ʧЧ��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���ʧЧ��")))
                !һ���Բ��� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("һ���Բ���")))
                !�޾��Բ��� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("�޾��Բ���"))
                !�ⷿ���� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("�ⷿ����")))
                !���÷��� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("���÷���")))
                !ʱ�� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʱ��")))
                !���� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����")))
                !���� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("����"))
                !Ч�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("Ч��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("Ч��")))
                !�������� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")))
                !��׼�ĺ� = vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��׼�ĺ�"))
                !��ҩ��λID = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��ҩ��λID")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��ҩ��λID")))
                !�������� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("��������")))
                !ʵ������ = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ������")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ������")))
                !ʵ�ʽ�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʽ��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʽ��")))
                !ʵ�ʲ�� = IIf(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʲ��")) = "", Null, vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ʵ�ʲ��")))
                !ָ�������� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ָ��������")))
                !ָ������� = Val(vsfѡ��.TextMatrix(i, vsfѡ��.ColIndex("ָ�������")))
                
                .Update
            Next
        End With
        
    End If
    
    CombinateRec = True
    Exit Function
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '��ǰ�����
    Dim intCol As Integer
    '����Ƿ�����ѡ��
    CheckData = False
    
    If mblnSelect = False Then Exit Function
    
    If mbln��ʾ���� = False Then
        CheckData = True
        Exit Function '����ǲ���ʾ����ģʽ��ֱ�Ӳ������
    End If
    
    If vsBatch.Visible Then
        'lng��Ӧ��ID��Ϊ�㣬��ʾ�˻����޿��ʱ��׼����
        If mlng��Ӧ��ID <> 0 Then
            intCol = vsBatch.ColIndex("�ϴι�Ӧ��ID")
            If intCol < 0 Then Exit Function
            If Val(vsBatch.TextMatrix(vsBatch.Row, intCol)) <> 0 And mlng��Ӧ��ID <> Val(vsBatch.TextMatrix(vsBatch.Row, intCol)) Then
                MsgBox "��ѡ����˻��̲��Ǹ��������ϵĹ�Ӧ�̣����ܼ���������", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    
        If mblnStock Then
            DblCurStock = Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("��������")))
        Else
            DblCurStock = Get���ÿ��(Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("����ID"))), Val(vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("����"))))
        End If
    Else
        If Not mrsCard.EOF Then
            If mblnStock Then
                DblCurStock = Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("��������")))
            Else
                DblCurStock = Get���ÿ��(Val(vsHead.TextMatrix(vsHead.Row, vsHead.ColIndex("����ID"))))
            End If
        End If
    End If
    
    If DblCurStock > 0 Then
        CheckData = True
        Exit Function
    End If
    
    '���Դ�ⷿ��Ŀ�ⷿΪ�գ������������Ŀ¼�Լ��ڽ��г������ã����ж�
    If (mlngԴ�ⷿID = 0 And mlngĿ�ⷿID = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '������̵㵥��������ѡ�����������жϣ�ֱ���˳�
    If mbln�̵㵥 Then
        CheckData = True
        Exit Function
    End If
    If vsBatch.Visible Or mblnʱ�� Then
        If (DblCurStock <> 0) Or Not mblnPrice Or vsBatch.TextMatrix(vsBatch.Row, vsBatch.ColIndex("����")) = "����������������" Then CheckData = True: Exit Function
        MsgBox "��" & IIf(mblnʱ��, "ʱ��", "����") & "�����Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    Else
        If mblnCheck = False Then
           CheckData = True
           Exit Function
        End If
    End If
    
    'mlng��Ӧ��ID��Ϊ�㣬��ʾ�˻����޿��ʱ��׼����
    If mlng��Ӧ��ID <> 0 Then
        MsgBox "�������Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    Select Case mintStockCheck
    Case 1
        If MsgBox("�������Ѿ�û�п�棬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Case 2
        MsgBox "�������Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
        Exit Function
    End Select
    CheckData = True
End Function

Public Function ShowSelect(ByVal frmMain As Form, ByVal �༭ģʽ As Integer, Optional ByVal Դ�ⷿ As Long = 0, _
                    Optional ByVal Ŀ�ⷿ As Long = 0, Optional ByVal ʹ�ò��� As Long = 0, Optional ByVal ��ѯ�� As String = "", _
                    Optional ByVal WinLeft As Double = 0, Optional ByVal WinTop As Double = 0, _
                    Optional ByVal lngWidth As Long = 0, Optional ByVal lngTxtHeight As Long = 0, Optional ByVal Bln����� As Boolean = True, _
                    Optional ByVal bln������λ�ʱ�� As Boolean = True, Optional ByVal mbln�̵㵥�� As Boolean = False, Optional ByVal bln���ӿ����� As Boolean = False, _
                    Optional ByVal bln��ʾ��� As Boolean = True, Optional ByVal lng��Ӧ�� As Long = 0, Optional ByVal blnɢװ��λ As Boolean = True, _
                    Optional ByVal str�̵�ʱ�� As String = "", _
                    Optional ByVal bln����ʾ������� As Boolean = False, _
                    Optional ByVal lngModule As Long = 0, _
                    Optional ByVal bln���޴洢�ⷿ���� As Boolean = False, _
                    Optional ByVal strPrivs As String = "", _
                    Optional ByVal bln��ʾ���� As Boolean = True, Optional ByVal bln�Ƿ���� As Boolean = True) As ADODB.Recordset
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ѡ��
    '����:
    '
    '   bln�����:�����������ļ�ʱ���������治׼����ԭ�򣬿�ǿ������not (���� or ʱ��) ���ĳ���
    '   bln������λ�ʱ��:����������������ļ�ʱ�����ĳ���
    '   mlng��Ӧ��ID:��Ϊ���ʾ�˻�
    '   str�̵�ʱ��:���̵���Ч,��Ҫ�Ǽ����̵�ʱ�õĿ����
    '����:��ѡ������ĵļ�¼��
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    mblnɢװ��λ = blnɢװ��λ
    
    mlngModule = lngModule
    If mlngModule = 1717 Then   '1717:��������
        mblnTrackUsing = IIf(Val(zlDatabase.GetPara("��������", glngSys, mlngModule, "0")) = 1, True, False)
    Else
        mblnTrackUsing = False
    End If
    
    '�޸�:���˺�   Bug:12972    ����:2008-05-08 15:28:14
    mbln����ʾ������� = bln����ʾ������� ': mlngModule = 0  '��ʱδ����ģ���,�Ժ���ݲ���������
    mblnSelectSucess = False
    If frmMain Is Nothing Then
        mstrTittle = "����ѡ����"
    Else
        mstrTittle = frmMain.Caption
    End If
    With mWindowPosition
        .Left = WinLeft
        .Top = WinTop
        .lngTxtH = lngTxtHeight
        .lngTxtW = lngWidth
    End With
    With Me
        .In_�༭״̬ = �༭ģʽ
        .In_Դ�ⷿ = Դ�ⷿ
        .In_Ŀ�ⷿ = Ŀ�ⷿ
        .In_���� = ʹ�ò���
        .In_�ִ� = Trim(��ѯ��)
        .In_MainFrm = frmMain
        mbln�̵㵥 = mbln�̵㵥��
        mbln������ = bln���ӿ�����
        mblnCheck = Bln�����
        mblnPrice = bln������λ�ʱ��
        mblnStock = bln��ʾ���
        mlng��Ӧ��ID = lng��Ӧ��
        mstr�̵�ʱ�� = str�̵�ʱ��
        mbln���޴洢�ⷿ���� = bln���޴洢�ⷿ����
        mstrPrivs = strPrivs
        mbln��ʾ���� = bln��ʾ����
        mbln�Ƿ���� = bln�Ƿ����
        Me.Caption = mstrTittle
        If mblnSelectSucess Then GoTo GoOk:
        .Show 1, frmMain
    End With
GoOk:
    Set ShowSelect = mrsReturn.Clone
End Function

Public Function Get���ÿ��(ByVal lng����ID As Long, Optional ByVal lng���� As Long = 0) As Single
    Dim rsTemp As New ADODB.Recordset
     
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        " Select Sum(A.��������" & mstrUnitString & ") ��������,Sum(A.ʵ������" & mstrUnitString & ") ʵ������,sum(A.ʵ�ʽ��) ʵ�ʽ��,sum(A.ʵ�ʲ��) ʵ�ʲ�� " & _
              " From ҩƷ��� A,�������� B " & _
              " Where A.ҩƷID=B.����ID and A.����=1 And A.ҩƷID=[1]" & IIf(lng���� = 0, "", " And Nvl(A.����,0)=[2]")
    
    If mlngԴ�ⷿID <> 0 Or mlngĿ�ⷿID <> 0 Then
        gstrSQL = gstrSQL & " And A.�ⷿID=" & IIf(mlngԴ�ⷿID = 0, "[4]", "[3]")
    End If
    
    
    gstrSQL = gstrSQL & " Group By A.ҩƷid"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, lng����, mlngԴ�ⷿID, mlngĿ�ⷿID)
    
    msin�������� = 0
    msinʵ�ʲ�� = 0
    msinʵ�ʽ�� = 0
    msinʵ������ = 0
    If Not rsTemp.EOF Then
        msin�������� = IIf(IsNull(rsTemp!��������), 0, rsTemp!��������)
        msinʵ�ʲ�� = IIf(IsNull(rsTemp!ʵ�ʲ��), 0, rsTemp!ʵ�ʲ��)
        msinʵ�ʽ�� = IIf(IsNull(rsTemp!ʵ�ʽ��), 0, rsTemp!ʵ�ʽ��)
        msinʵ������ = IIf(IsNull(rsTemp!ʵ������), 0, rsTemp!ʵ������)
    End If
    Get���ÿ�� = msin��������
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub imgLeft_Click()
    Call LoadFulltoColSel(False)
End Sub

Public Sub ReSetWindowsFormLocal()
    '����:�������ô��ڵĴ�С��λ��
    Dim dblColsWidth As Double, dblMinRowheight As Double, lngScrW As Long
    Dim lngTaskHeight As Long
    Dim dblRowsHeight As Double
    Dim dblRowBatchHeight As Double
    Dim dblTemp As Double
    Dim i As Long
    '��λ
    With mWindowPosition
        Me.Left = .Left + 15
        Me.Top = .Top
    End With
    
    dblColsWidth = 0
    For i = 0 To vsHead.Cols - 1
        If Not vsHead.ColHidden(i) Then
            dblColsWidth = dblColsWidth + vsHead.ColWidth(i) + 15
        End If
    Next
    dblMinRowheight = vsBatch.RowHeightMin
    lngTaskHeight = GetTaskbarHeight
    dblColsWidth = dblColsWidth + 300
    lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75

    dblRowsHeight = dblMinRowheight * vsHead.Rows + 30
    dblRowBatchHeight = (dblMinRowheight) * 6 'Ŀǰ�̶�����  'IIf(vsBatch.Visible, 1, 0) *   'IIf(vsBatch.Rows <= 4, 4, vsBatch.Rows)
    
    dblColsWidth = IIf(dblColsWidth < MFRM_MIN_WIDTH, MFRM_MIN_WIDTH, dblColsWidth)
    
    If Me.Top + dblRowsHeight + dblRowBatchHeight <= Screen.Height Then
        '���嶥��+���и߶�+С�ڵ�����Ļ�߶ȡ�
        '���Ƿ����С�߶Ȼ�С,�����С,������С�ȸ�Ϊ׼
        Me.vsBatch.Height = dblRowBatchHeight
        If dblRowBatchHeight + dblRowsHeight < MFRM_MIN_HEIGHT Then
            Me.Height = MFRM_MIN_HEIGHT
        Else
            Me.Height = dblRowBatchHeight + dblRowsHeight
        End If
        
        Me.vsBatch.Height = dblRowBatchHeight
     '   If Me.ScaleHeight < Me.vsBatch.Top + Me.vsBatch.Height Then Me.vsBatch.Height = Me.ScaleHeight - Me.vsBatch.Top
        
    Else
        '���嶥��+�������߶�+���ŵ��ܸ߶ȴ�����Ļ�߶�,��Ҫ��һ�¼��
        '1.���ϰ���Ļ�߶��Ƿ���°����߶�Ҫ�ߣ���������ϰ����ĸ߶�Ϊ׼���������°���Ϊ׼.
        
        If Screen.Height - Me.Top > Me.Top - mWindowPosition.lngTxtH - 15 Then
            '�°���Ҫ��
            Me.Height = Screen.Height - Me.Top - lngTaskHeight
            '������ȫװ��,ֻ�ܸ���������������б�������б�ĸ߶�
            dblTemp = Me.ScaleHeight - dblRowsHeight
            If dblTemp > 6 * dblMinRowheight + 30 Then
               '���µĸ߶�Ҫ����4�и߶�,�������εĸ߶Ⱦ�Ϊ���µĸ߶�
               vsBatch.Height = dblTemp
            Else
                '���µĸ߶Ȳ���4�еĸ߶�,����4�и߶�Ϊ׼
                vsBatch.Height = 6 * dblMinRowheight + 30
            End If
        Else
            dblTemp = Me.Top - mWindowPosition.lngTxtH - 15
            Me.Top = Me.Top - mWindowPosition.lngTxtH - 15
            '�ϰ���Ҫ��
            If dblTemp - dblRowBatchHeight - dblRowsHeight > 0 Then
                '�ϰ�������ȫ��װ��
                Me.vsBatch.Height = dblRowBatchHeight
                Me.Height = dblRowBatchHeight + dblRowsHeight
                If Me.Height < MFRM_MIN_HEIGHT Then Me.Height = MFRM_MIN_HEIGHT
            Else
                Me.Height = dblTemp
                '������ȫװ��,ֻ�ܸ���������������б�������б�ĸ߶�
                dblTemp = Me.ScaleHeight - dblRowsHeight
                If dblTemp > 4 * dblMinRowheight Then
                   '���µĸ߶�Ҫ����4�и߶�,�������εĸ߶Ⱦ�Ϊ���µĸ߶�
                   vsBatch.Height = dblTemp
                Else
                    '���µĸ߶Ȳ���4�еĸ߶�,����4�и߶�Ϊ׼
                    vsBatch.Height = 4 * dblMinRowheight
                End If
            End If
            Me.Top = Me.Top - Me.Height
        End If
    End If
    '�����ȶ�λ
    '����п�����С�ڵ��ڵ�ǰ����Ŀ��,����������Ϊ׼
    If dblColsWidth + Me.Left < Screen.Width Then
        '���еĿ����ȫ����ʾ
        Me.Width = dblColsWidth
    Else
        '����Ƿ������Ļ�����ұ���Ļ��
        If Screen.Width - Me.Left >= Me.Left Then
            '�ұ���Ļ��
            Me.Width = Screen.Width - Me.Left
        Else
            Me.Left = Me.Left + mWindowPosition.lngTxtW
            '�����Ļ��
            If dblColsWidth < Me.Left Then
                Me.Width = dblColsWidth
            Else
                Me.Width = Me.Left
            End If
            Me.Left = Me.Left - Me.Width
        End If
    End If
 
    vsBatch.Top = Me.ScaleHeight - vsBatch.Height
    With vsHead
        .Height = IIf(vsBatch.Visible = False, Me.ScaleHeight - .Top, vsBatch.Top - .Top)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
        .Width = Me.ScaleWidth
        
    End With
    With vsBatch
        .Width = vsHead.Width
        .Left = ScaleLeft
    End With
End Sub

Private Function LoadFulltoColSel(ByVal blnBatch As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-09 16:46:43
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    Dim sngFrmHeight As Single, sngSelSumHeight As Single
    
    If blnBatch Then
        Set vsGrid = vsBatch
        vsColSet.Tag = "Batch"
    Else
        Set vsGrid = vsHead
        vsColSet.Tag = "Head"
    End If
    vsColSet.Clear 1
    vsColSet.Rows = 2
    With vsGrid
        lngRow = 1
        For i = 0 To .Cols - 1
            '.coldata(i):1-�̶�,-1-����ѡ,0-��ѡ
            If Trim(.ColKey(i)) <> "" And (.ColData(i) = 1 Or .ColData(i) = 0) Then
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("����")) = .ColKey(i)
                vsColSet.TextMatrix(lngRow, vsColSet.ColIndex("ѡ��")) = IIf(.ColWidth(i) = 0 Or .ColHidden(i), False, True)
                vsColSet.RowData(lngRow) = .ColData(i)
                If .ColData(i) = 1 Then
                    vsColSet.Cell(flexcpForeColor, lngRow, 0, lngRow, vsColSet.Cols - 1) = vbBlue
                End If
                vsColSet.Rows = vsColSet.Rows + 1
                lngRow = lngRow + 1
            End If
        Next
    End With
    If vsColSet.Rows > 2 Then vsColSet.Rows = vsColSet.Rows - 1
    sngFrmHeight = Me.ScaleHeight
    With vsColSet
        sngSelSumHeight = (.RowHeight(0) + 60) * (.Rows) + 60
        .Cell(flexcpBackColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000001
        .Cell(flexcpForeColor, 0, 0, 0, vsColSet.Cols - 1) = &H80000005
        .BackColorSel = &H8000000D
        .Row = 1
        .Visible = True
        .Editable = flexEDKbdMouse
        .ZOrder 0
        .Left = vsGrid.Left + .Cell(flexcpWidth, 0, 0, 0, 0) + 30
        If blnBatch Then
            .Height = IIf(vsGrid.Top > sngSelSumHeight, sngSelSumHeight, vsGrid.Top)
            .Top = vsBatch.Top - .Height
        Else
            .Top = vsGrid.Top + vsGrid.RowHeight(0) + 15
            sngFrmHeight = sngFrmHeight - .Top
            If sngFrmHeight > sngSelSumHeight Then
                .Height = sngSelSumHeight
            Else
                .Height = IIf(sngFrmHeight < 0, 0, sngFrmHeight)
            End If
        End If
        .SetFocus
    End With
End Function
Private Function SetVsGridCol(ByVal strColKey As String, ByVal blnShow As Boolean, ByVal blnBatch As Boolean) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������ʾ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-09 17:31:22
    '-----------------------------------------------------------------------------------------------------------
    Dim vsGrid As VSFlexGrid, i As Long, lngRow As Long
    If blnBatch Then
        Set vsGrid = vsBatch
    Else
        Set vsGrid = vsHead
    End If
    With vsGrid
        .ColHidden(.ColIndex(strColKey)) = Not blnShow
        If .ColWidth(.ColIndex(strColKey)) = 0 Then .ColWidth(.ColIndex(strColKey)) = 1000
    End With
    If blnBatch Then
        zl_vsGrid_Para_Save mlngModule, vsBatch, mstrTittle, "������Ϣ", False
    End If
End Function
Private Sub vsColSet_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�޸ĺ�
    Dim strColKey As String, blnShow As Boolean
    With vsColSet
        Select Case Col
        Case .ColIndex("ѡ��")
            blnShow = GetVsGridBoolColVal(vsColSet, Row, .ColIndex("ѡ��"))
            Call SetVsGridCol(.TextMatrix(Row, .ColIndex("����")), blnShow, IIf(.Tag = "Head", False, True))
        Case Else
        End Select
    End With
End Sub

Private Sub vsColSet_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsColSet
        Select Case Col
        Case .ColIndex("ѡ��")
            'rowdata(i):1-�̶�,-1-����ѡ,0-��ѡ
            If (.TextMatrix(Row, 1) = "�����" Or .TextMatrix(Row, 1) = "�ɱ���" Or .TextMatrix(Row, 1) = "�ϴβɹ���" Or .TextMatrix(Row, 1) = "�ϴι���") And mblnCostView = False Then
                Cancel = True
            End If
            If .RowData(Row) = 1 Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Sub vsColSet_LostFocus()
    vsColSet.Visible = False
End Sub

