VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFilesUpgradeManage 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   6612
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6612
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRepairFiles 
      Caption         =   "�������ļ����(&T)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7755
      TabIndex        =   4
      ToolTipText     =   "�����������ϴ����ļ����Զ��޸�"
      Top             =   180
      Width           =   1875
   End
   Begin VB.CommandButton cmdRepairList 
      Caption         =   "�����ļ��嵥����(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5760
      TabIndex        =   3
      ToolTipText     =   "������ԭ�����ļ��嵥������Ϊ��׼�嵥"
      Top             =   180
      Width           =   1875
   End
   Begin VB.CommandButton cmdUpLoad 
      Caption         =   "�����ļ��ϴ�(&R)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3765
      TabIndex        =   2
      ToolTipText     =   "�ϴ��ļ��������úÿ����ӵķ�����"
      Top             =   180
      Width           =   1600
   End
   Begin VB.PictureBox picBtn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1305
      ScaleHeight     =   348
      ScaleWidth      =   6780
      TabIndex        =   16
      Top             =   6060
      Width           =   6780
      Begin VB.CheckBox chkFilter 
         Caption         =   "ֻ��ʾ����������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4185
         TabIndex        =   14
         Top             =   30
         Width           =   2055
      End
      Begin VB.CommandButton cmdExpired 
         Caption         =   "����(&Q)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3015
         TabIndex        =   13
         ToolTipText     =   "���������ļ��嵥�е�ĳ���������ļ������������ļ��嵥"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2010
         TabIndex        =   12
         ToolTipText     =   "ɾ�������ļ��嵥�е�ĳ���������ļ���Ϣ"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "�޸�(&E)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1005
         TabIndex        =   11
         ToolTipText     =   "�޸������ļ��嵥�еĵ������ļ���Ϣ"
         Top             =   0
         Width           =   900
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "�����������ļ��������ļ��嵥"
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   570
      ScaleHeight     =   264
      ScaleWidth      =   2676
      TabIndex        =   15
      Top             =   195
      Width           =   2700
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   45
         TabIndex        =   1
         Top             =   30
         Width           =   2580
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFileList 
      Height          =   1005
      Left            =   600
      TabIndex        =   7
      Top             =   1905
      Width           =   3870
      _cx             =   6826
      _cy             =   1764
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483638
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   2
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFilesUpgradeManage.frx":0000
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
      Ellipsis        =   1
      ExplorerBar     =   7
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfExpFileList 
      Height          =   1005
      Left            =   600
      TabIndex        =   8
      Top             =   2910
      Visible         =   0   'False
      Width           =   3870
      _cx             =   6826
      _cy             =   1764
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483638
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   2
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFilesUpgradeManage.frx":0161
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
      Ellipsis        =   1
      ExplorerBar     =   7
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   6120
      Width           =   900
   End
   Begin VB.Label LblBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�����ļ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   5415
      TabIndex        =   6
      ToolTipText     =   "�Ѿ����õ��ļ�"
      Top             =   1100
      Width           =   855
   End
   Begin VB.Label LblBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�����ļ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   4575
      MousePointer    =   4  'Icon
      TabIndex        =   5
      ToolTipText     =   "����ʹ�õ��ļ�"
      Top             =   1100
      Width           =   855
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   0
      Top             =   240
      Width           =   360
   End
End
Attribute VB_Name = "frmFilesUpgradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Const StopColor = vbRed '����ʱ����ɫ
'Const StartColor = &H80000008 '����ʱ����ɫ
'Const SelectColor = &H8000000A 'ѡ��ʱ������ɫ
'Const MoveColor = &H80000004 '�ƶ�ʱ������ɫ
'Const noSelectColor = &HFFFFFF 'δѡ��ʱ������ɫ

'Const SelectColor = &HFFFFFF
'Const MoveColor = &H8000000A
'Const noSelectColor = &H80000004

Const COLOR_SELECT = &H80000004 'ѡ��ʱ������ɫ
Const COLOR_MOVE = &HFFFFFF '����ƶ�����ťʱ������ɫ
Const COLOR_NOT_SELECT = &H8000000A 'δѡ��ʱ������ɫ

'Private mobjFindKey             As CommandBarPopup      '��ѯ
Private mstrFindKey             As String               '��ѯ��
Private m_strCurTypeName        As String               '��ǰѡ�еķ�ʽ
Private m_strCurFileName        As String               '��ǰѡ�е�����
Private m_strCurVision          As String               '��ǰѡ�еİ汾
Private m_strCurEditDate        As String               '��ǰѡ�е��޸�����
Private m_strCurSysNum          As String               '��ǰѡ�е�ϵͳ
Private m_strCurSetupPath       As String               '��ǰѡ�еİ�װ·��
Private m_strCurSetupPathADD       As String         '��ǰѡ�еĸ��Ӱ�װ·��
Private m_strCurSysOption       As String               '��ǰѡ�е�ϵͳ����
Private m_strCurFileExplanation As String               '��ǰѡ�е��ļ�˵��
Private m_strCurSellFile        As String               '��ǰѡ�е������ļ�
Private m_blnCurReg             As Boolean              '��ǰѡ�е��ļ��Ƿ�ע��
Private m_blnCurUpData          As Boolean              '��ǰѡ�е��ļ��Ƿ�ǿ�Ƹ���
Private mintfgMainTag           As Integer              '��ǰ�����ʾ 0-�����ļ� 1-�����ļ�
Private mrsTemp      As New ADODB.Recordset
Private mstrLocationFileName As String
Public blnRefreshData As Boolean '�����л�ˢ���жϱ�־

Public Enum RegFileType
    RFT_NotReg = 0                  '��ע��Ķ���
    RFT_NormalReg = 1               '����ע�ᣬ�Զ�ʶ��.NET������.NET����ͨ��Regasmע�ᣬ����ͨ������DLLRegServerע��
    RFT_NETGAC = 2                  'NET����ע�ᣬͨ��gacutilע�ᵽȫ�ֳ��򼯻���
    RFT_NETServer = 3               'NET����ע�ᣬͨ��installUtil���а�װж�ء�
    RFT_NETComReg = 4               '.NET Com����ע�ᣬͨ������Regasm���
    RFT_VBComReg = 5                'ͨ����дע���ע��
    RFT_DelphiComReg = 6            'DelphiComע�ᣬͨ��DLLRegServerע��
    RFT_PBComReg = 7                'PBComע�ᣬͨ��DLLRegServerע��
End Enum

Public Enum FileListCol
    FC_��� = 0
    FC_�ļ����� = 1
    FC_�ļ��� = 2
    FC_�汾�� = 3
    FC_�޸����� = 4
    FC_����ϵͳ = 5
    FC_ҵ�񲿼� = 6
    FC_��ʵ·�� = 7
    FC_����ID = 8
    FC_��װ·�� = 9
    FC_ϵͳ���� = 10
    FC_�Զ�ע�� = 11
    FC_�ļ�˵�� = 12
    FC_ǿ�Ƹ��� = 13
    FC_���Ӱ�װ·�� = 14
    FC_���� = 15
End Enum

Public Enum ExpFileListCol
    EFC_��� = 0
    EFC_�ļ��� = 1
    EFC_ϵͳ�汾 = 2
    EFC_��װ·�� = 3
    EFC_ϵͳ��� = 4
    EFC_�ļ�˵�� = 5
    EFC_���� = 6
End Enum

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = IIf(mintfgMainTag = 0, "�����ļ��嵥", "�����ļ��嵥")

    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡʱ�䣺" & Format(date, "yyyy��MM��dd��")
    Set objPrint.Body = IIf(mintfgMainTag = 0, Me.vsfFileList, Me.vsfExpFileList)
    objPrint.BelowAppRows.Add objRow
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub chkFilter_Click()
    RefreshData
End Sub

Private Sub cmdExpired_Click()
    Call StandardAbandon
End Sub

Private Sub cmdAdd_Click()
    '����
    Call StandardAdd
End Sub

Private Sub cmdDel_Click()
    'ɾ��
    Call StandardDel
End Sub

Private Sub cmdEdit_Click()
    '�޸�
     Call StandardEdit
End Sub

Private Sub cmdRepairFiles_Click()
    Dim frmRepair As New frmFilesRepair
    
    frmRepair.ShowMe
End Sub

Private Sub cmdRepairList_Click()
    Dim frmMsgbox As New frmMessageBox
    If frmMsgbox.ShowMe(0, gstrSysName) Then
        Call ExecuteProcedure("zlFilesUpgrade_Repair", Me.Caption)
        Me.RefreshData
    End If
'    If MsgBox("��������Ҫ�����ϴ������ļ������������Ƿ���Ҫ�޵�ǰ����ʹ�õ��ļ��嵥��", vbQuestion + vbOKCancel, gstrSysName) <> vbCancel Then
'        Call ExecuteProcedure("zlFilesUpgrade_Repair", Me.Caption)
'   End If
End Sub

Private Sub cmdUpload_Click()
    Call StandardUpLoad
End Sub

Private Sub vsfExpFileList_AfterSort(ByVal Col As Long, Order As Integer)
    vsfExpFileList.Row = vsfExpFileList.FindRow(mstrLocationFileName, , 2)
    If vsfExpFileList.Row > 0 Then vsfExpFileList.ShowCell vsfExpFileList.Row, 0
End Sub

Private Sub vsfExpFileList_RowColChange()
    mstrLocationFileName = vsfExpFileList.TextMatrix(vsfExpFileList.Row, 2)
End Sub

Private Sub vsfFileList_AfterSort(ByVal Col As Long, Order As Integer)
    vsfFileList.Row = vsfFileList.FindRow(mstrLocationFileName, , 2)
    If vsfFileList.Row > 0 Then vsfFileList.ShowCell vsfFileList.Row, 0
End Sub

Private Sub vsfFileList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(0).BackColor = COLOR_MOVE Then LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
    If LblBtn.Item(1).BackColor = COLOR_MOVE Then LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
End Sub

Private Sub vsfExpFileList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(0).BackColor = COLOR_MOVE Then LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
    If LblBtn.Item(1).BackColor = COLOR_MOVE Then LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        'ˢ��
    End If
    If KeyCode = vbKeyDelete Then
        If cmdDel.Enabled Then
            cmdDel_Click
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(0).BackColor = COLOR_MOVE Then LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
    If LblBtn.Item(1).BackColor = COLOR_MOVE Then LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    vsfFileList.Move 50, 950, Me.ScaleWidth - 120, Me.ScaleHeight - 950 - 600
    vsfExpFileList.Move 50, 950, Me.ScaleWidth - 120, Me.ScaleHeight - 950 - 600
    
'    LblBtn.Item(0).Move Me.ScaleWidth / 2 - LblBtn.Item(0).Width - 250, 1100
'    LblBtn.Item(1).Move Me.ScaleWidth / 2 - 265, 1100
    
    LblBtn.Item(0).Move 50, 700
    LblBtn.Item(1).Move LblBtn.Item(0).Width - 15, 700

'    LblBtn.Item(0).Move 50, 1100, vsfFileList.Width / 2
'    LblBtn.Item(1).Move vsfFileList.Width / 2 + 30, 1100, vsfFileList.Width / 2 + 15
    lblItem.Item(0).Move 100, vsfFileList.Top + vsfFileList.Height + 200
    picBtn.Move lblItem.Item(0).Left + 1100, lblItem.Item(0).Top - 50

End Sub


'==============================================================================
'=���ܣ� ���ڳ�ʼ��
'==============================================================================
Private Sub Form_Load()
  On Error GoTo errH
    KeyPreview = True
    '���ҿ��ʼ��
    txtFind.Tag = "�������ļ����Ʋ���"
    txtFind.Text = txtFind.Tag
    txtFind.ForeColor = vbGrayText
    mintfgMainTag = 0
'    LblBtn.Item(0).Move 50, 1100
'    LblBtn.Item(1).Move LblBtn.Item(0).Width + 30, 1055, LblBtn.Item(1).Width, LblBtn.Item(1).Height + 45
    '���Combo
'    Call InitComBo
'    Call InitVsfMain

'    LoadFilesList
'    LoadExpFilesList
'    LblBtn_Click 0
'    Call SetMenu
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub


'==============================================================================
'=���ܣ� ����fgMain������ˢ��״̬��Ϣ
'==============================================================================
Private Sub vsfFileList_Click()
    On Error GoTo errH
    vsfFileList_SelChange
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=���ܣ� �������б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub vsfFileList_RowColChange()
    On Error GoTo errH
    Call vsfFileList_SelChange
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=���ܣ� ����ѡ�����б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub vsfFileList_SelChange()
    Dim lngID       As Long
    On Error GoTo errH

    If vsfFileList.Row = 0 Then Exit Sub
    m_strCurTypeName = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 1)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 1))   '��ȡID
    m_strCurFileName = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 2)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 2))     '�ļ���
    m_strCurVision = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 3)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 3))
    m_strCurEditDate = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 4)) = 0, 0, vsfFileList.Cell(flexcpText, vsfFileList.Row, 4))
    m_strCurSysNum = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 5)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 5))
    m_strCurSellFile = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 6)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 6))
    m_strCurSetupPath = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 7)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 7))
    m_strCurSetupPathADD = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, FC_���Ӱ�װ·��)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, FC_���Ӱ�װ·��))
    m_strCurSysOption = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 10)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 10))
    m_blnCurReg = IIf(vsfFileList.Cell(flexcpText, vsfFileList.Row, 11) = "��", True, False) '�Զ�ע��
    m_strCurFileExplanation = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 12)) = 0, "", vsfFileList.Cell(flexcpText, vsfFileList.Row, 12)) '�ļ�˵��
    m_blnCurUpData = IIf(Len(vsfFileList.Cell(flexcpText, vsfFileList.Row, 13)) = 0, False, vsfFileList.Cell(flexcpText, vsfFileList.Row, 13)) 'ǿ�Ƹ���

    If m_strCurTypeName = "��������" Then
        cmdEdit.Enabled = True
        cmdExpired.Enabled = True
        cmdDel.Enabled = True
    Else
        cmdEdit.Enabled = False
        cmdExpired.Enabled = False
        cmdDel.Enabled = False
    End If
    mstrLocationFileName = m_strCurFileName
'    Call SetMenu
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub vsfFileList_DblClick()
    If vsfFileList.MouseRow <> vsfFileList.Row Then Exit Sub '�̶���˫����Ч
    If m_strCurTypeName = "��������" Then
        Call StandardEdit
    End If
End Sub

Public Sub LoadFilesList(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp      As String
    Dim arrSys As Variant
    On Error GoTo errH
    
    If strFilter = "" Then strFilter = "0,1,2,3,4,5"
    
    With vsfFileList
        .Redraw = flexRDNone
        .Tag = ""
        .Rows = 1
        .Clear
        .Cols = FC_����
        
        .Cell(flexcpText, 0, FC_���) = "���"
        .Cell(flexcpAlignment, 0, FC_���) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, FC_�ļ�����) = "�ļ�����"
        .Cell(flexcpAlignment, 0, FC_�ļ�����) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, FC_�ļ���) = "�ļ���"
        .Cell(flexcpAlignment, 0, FC_�ļ���) = flexAlignCenterCenter
        .ColWidth(FC_�ļ���) = 2200
        
        .Cell(flexcpText, 0, FC_�汾��) = "�汾��"
        .Cell(flexcpAlignment, 0, FC_�汾��) = flexAlignCenterCenter
        .ColWidth(FC_�汾��) = 1200
        
        .Cell(flexcpText, 0, FC_�޸�����) = "�޸�����"
        .Cell(flexcpAlignment, 0, FC_�޸�����) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, FC_����ϵͳ) = "����ϵͳ"
        .Cell(flexcpAlignment, 0, FC_����ϵͳ) = flexAlignCenterCenter
        .ColWidth(FC_����ϵͳ) = 1800
        
        .Cell(flexcpText, 0, FC_ҵ�񲿼�) = "ҵ�񲿼�"
        .Cell(flexcpAlignment, 0, FC_ҵ�񲿼�) = flexAlignCenterCenter
        .ColWidth(FC_ҵ�񲿼�) = 3000

        .Cell(flexcpText, 0, FC_��ʵ·��) = "��װ·��"
        .Cell(flexcpAlignment, 0, FC_��ʵ·��) = flexAlignCenterCenter
        .ColHidden(FC_��ʵ·��) = True
        
        .Cell(flexcpText, 0, FC_����ID) = "����ID"
        .Cell(flexcpAlignment, 0, FC_����ID) = flexAlignCenterCenter
        .ColHidden(FC_����ID) = True

        .Cell(flexcpText, 0, FC_��װ·��) = "��װ·��"
        .Cell(flexcpAlignment, 0, FC_��װ·��) = flexAlignCenterCenter
        .ColWidth(FC_��װ·��) = 2000

        .Cell(flexcpText, 0, FC_ϵͳ����) = "ϵͳ����"
        .Cell(flexcpAlignment, 0, FC_ϵͳ����) = flexAlignCenterCenter
        .ColHidden(FC_ϵͳ����) = True
        
        .Cell(flexcpText, 0, FC_�Զ�ע��) = "�Զ�ע��"
        .Cell(flexcpAlignment, 0, FC_�Զ�ע��) = flexAlignCenterCenter
        .ColWidth(FC_�Զ�ע��) = 1000

        .Cell(flexcpText, 0, FC_�ļ�˵��) = "�ļ�˵��"
        .Cell(flexcpAlignment, 0, FC_�ļ�˵��) = flexAlignCenterCenter
        .ColWidth(FC_�ļ�˵��) = 5000

        .Cell(flexcpText, 0, FC_ǿ�Ƹ���) = "ǿ�Ƹ���"
        .ColHidden(FC_ǿ�Ƹ���) = True
        
        .Cell(flexcpText, 0, FC_���Ӱ�װ·��) = "���Ӱ�װ·��"
        .ColHidden(FC_���Ӱ�װ·��) = True
        
        If CheckAndAdjustMustTable("zlFilesUpgrade", , True) = False Then
            Exit Sub
        End If
        
        strSQL = "Select a.���,a.�ļ����� As ����ID,Decode(a.�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', 'δ֪����') As �ļ�����, a.�ļ���, a.�ļ��汾�� �汾��, a.�޸�����," & vbNewLine & _
                        "       a.����ϵͳ, a.ҵ�񲿼�,a.��װ·��,a.�ļ�˵��,a.�Զ�ע��,a.���Ӱ�װ·��" & vbNewLine & _
                        "From zlFilesUpgrade A" & vbNewLine & _
                        "Where a.�ļ����� In (" & strFilter & ") order by lpad(a.���,5,'0')"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '��������
        .Rows = mrsTemp.RecordCount + 1

        i = 1
        Do Until mrsTemp.EOF
            .Cell(flexcpText, i, FC_���) = Nvl(mrsTemp.Fields("���"), 0) 'mrsTemp.AbsolutePosition
            .Cell(flexcpAlignment, i, FC_���) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_�ļ�����) = Nvl(mrsTemp.Fields("�ļ�����"))
            .Cell(flexcpAlignment, i, FC_�ļ�����) = flexAlignCenterCenter

            .Cell(flexcpText, i, FC_�ļ���) = Nvl(mrsTemp.Fields("�ļ���"))
            .Cell(flexcpAlignment, i, FC_�ļ���) = flexAlignLeftCenter

            strTemp = Nvl(mrsTemp.Fields("�汾��"))
'            strTemp = GetFileVision(strTemp)

            .Cell(flexcpText, i, FC_�汾��) = strTemp
            .Cell(flexcpAlignment, i, FC_�汾��) = flexAlignLeftCenter

            If Nvl(mrsTemp.Fields("�޸�����")) <> "" Then
                strTemp = Format(Nvl(mrsTemp.Fields("�޸�����")), "yyyy-mm-dd hh:mm:ss")
            Else
                strTemp = ""
            End If

            .Cell(flexcpText, i, FC_�޸�����) = strTemp
            .Cell(flexcpAlignment, i, FC_�޸�����) = flexAlignCenterCenter

            strTemp = Nvl(mrsTemp.Fields("����ϵͳ"))

            If Trim(strTemp) <> "" Then
                arrSys = Split(Trim(strTemp), ",")
                strTemp = ""
                For j = 0 To UBound(arrSys)
                    If GetSystemName(arrSys(j)) <> "" Then strTemp = strTemp & "��" & GetSystemName(arrSys(j))
                Next
                strTemp = Mid(strTemp, 2)
            End If

            .Cell(flexcpText, i, FC_����ϵͳ) = strTemp
            .Cell(flexcpData, i, FC_����ϵͳ) = Nvl(mrsTemp.Fields("����ϵͳ"))
            .Cell(flexcpAlignment, i, FC_����ϵͳ) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_ҵ�񲿼�) = Nvl(mrsTemp.Fields("ҵ�񲿼�"))
            .Cell(flexcpAlignment, i, FC_ҵ�񲿼�) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_��ʵ·��) = Nvl(mrsTemp.Fields("��װ·��"))
            .Cell(flexcpAlignment, i, FC_��ʵ·��) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_����ID) = Nvl(mrsTemp.Fields("����ID"))
            .Cell(flexcpAlignment, i, FC_����ID) = flexAlignLeftTop

            .Cell(flexcpText, i, FC_��װ·��) = Nvl(mrsTemp.Fields("��װ·��"))
            .Cell(flexcpAlignment, i, FC_��װ·��) = flexAlignLeftCenter

'            .Cell(flexcpText, i, FC_����ϵͳ) = Nvl(mrsTemp.Fields("����ϵͳ")) 'NVL(mrsTemp.Fields("ϵͳ����"))
'            .Cell(flexcpAlignment, i, FC_����ϵͳ) = flexAlignCenterCenter

            .Cell(flexcpText, i, FC_�Զ�ע��) = IIf(Nvl(mrsTemp.Fields("�Զ�ע��"), "") = "0", "��", "��")
            .Cell(flexcpAlignment, i, FC_�Զ�ע��) = flexAlignCenterCenter

            .Cell(flexcpText, i, FC_�ļ�˵��) = Nvl(mrsTemp.Fields("�ļ�˵��"), "")
            .Cell(flexcpAlignment, i, FC_�ļ�˵��) = flexAlignLeftCenter

            .Cell(flexcpText, i, FC_ǿ�Ƹ���) = ""
       
            .Cell(flexcpText, i, FC_���Ӱ�װ·��) = Nvl(mrsTemp.Fields("���Ӱ�װ·��"), "")

            mrsTemp.MoveNext
            i = i + 1
        Loop
        
        'ѡ�п���
        .FocusRect = flexFocusSolid
        '���һ���Զ��п�
        .ExtendLastCol = True
        '�����������
        .ScrollTrack = True
        '�Զ�����
        .WordWrap = True
        '�и�����
        .RowHeightMin = 300
        .RowHeightMax = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
'        .AllowSelection = False
        
        .Redraw = flexRDBuffered
        
        'ˢ�¶�λ
        If strLocationName <> "" Then
            strLocationName = UCase(strLocationName)
            For j = 0 To .Rows - 1
                If UCase(.TextMatrix(j, 2)) = strLocationName Then .Row = j: Call .ShowCell(j, 2): Exit For
            Next
        Else
            If .Rows > 1 Then .Row = 1
        End If
        'ˢ���޸ġ�ɾ����ť״̬
        vsfFileList_SelChange

        If .Visible = True Then .SetFocus
'         Call SetMenu
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub

Public Sub LoadExpFilesList(Optional ByVal strFilter As String, Optional ByVal strLocationName As String)
'�ϳ����
    Dim i, j As Long
    Dim strSQL       As String
    Dim strTemp      As String
    On Error GoTo errH

    With vsfExpFileList
        .Redraw = flexRDNone
        .Tag = ""
        .Rows = 1
        .Clear
        .Cols = EFC_����
        
        .Cell(flexcpText, 0, EFC_���) = "���"
        .Cell(flexcpAlignment, 0, EFC_���) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, EFC_�ļ���) = "�ļ���"
        .Cell(flexcpAlignment, 0, EFC_�ļ���) = flexAlignCenterCenter
        .ColWidth(EFC_�ļ���) = 1800
        
        .Cell(flexcpText, 0, EFC_ϵͳ�汾) = "ϵͳ�汾"
        .Cell(flexcpAlignment, 0, EFC_ϵͳ�汾) = flexAlignCenterCenter
        .ColWidth(EFC_ϵͳ�汾) = 1000
        
        .Cell(flexcpText, 0, EFC_��װ·��) = "��װ·��"
        .Cell(flexcpAlignment, 0, EFC_��װ·��) = flexAlignCenterCenter
        .ColWidth(EFC_��װ·��) = 3000
        
        .Cell(flexcpText, 0, EFC_ϵͳ���) = "ϵͳ���"
        .Cell(flexcpAlignment, 0, EFC_ϵͳ���) = flexAlignCenterCenter
        .ColWidth(EFC_ϵͳ���) = 1000
        
        .Cell(flexcpText, 0, EFC_�ļ�˵��) = "�ļ�˵��"
        .Cell(flexcpAlignment, 0, EFC_�ļ�˵��) = flexAlignCenterCenter
        .ColWidth(EFC_�ļ�˵��) = 5000

        If CheckAndAdjustMustTable("zlFilesUpgrade", , True) = False Then
            Exit Sub
        End If
        
        strSQL = "select * from zlfilesexpired"
        Call OpenRecordset(mrsTemp, strSQL, Me.Caption)

        '��������
        .Rows = mrsTemp.RecordCount + 1

        i = 1
        Do Until mrsTemp.EOF
            .Cell(flexcpText, i, EFC_���) = i 'Nvl(mrsTemp.Fields("���"), 0) 'mrsTemp.AbsolutePosition
            .Cell(flexcpAlignment, i, EFC_���) = flexAlignCenterCenter

            .Cell(flexcpText, i, EFC_�ļ���) = Nvl(mrsTemp.Fields("�ļ���"))
            .Cell(flexcpAlignment, i, EFC_�ļ���) = flexAlignLeftCenter

            .Cell(flexcpText, i, EFC_ϵͳ�汾) = Nvl(mrsTemp.Fields("ϵͳ�汾"))
            .Cell(flexcpAlignment, i, EFC_ϵͳ�汾) = flexAlignLeftCenter

            .Cell(flexcpText, i, EFC_��װ·��) = Nvl(mrsTemp.Fields("��װ·��"))
            .Cell(flexcpAlignment, i, EFC_��װ·��) = flexAlignLeftCenter

            .Cell(flexcpText, i, EFC_ϵͳ���) = Nvl(mrsTemp.Fields("ϵͳ���"))
            .Cell(flexcpAlignment, i, EFC_ϵͳ���) = flexAlignCenterCenter

            .Cell(flexcpText, i, EFC_�ļ�˵��) = Nvl(mrsTemp.Fields("˵��"), "")
            .Cell(flexcpAlignment, i, EFC_�ļ�˵��) = flexAlignLeftCenter

            mrsTemp.MoveNext
            i = i + 1
        Loop
        
        'ѡ�п���
        .FocusRect = flexFocusSolid
        '���һ���Զ��п�
        .ExtendLastCol = True
        '�����������
        .ScrollTrack = True
        '�Զ�����
        .WordWrap = True
        '�и�����
        .RowHeightMin = 300
        .RowHeightMax = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .AllowUserResizing = flexResizeColumns
'        .AllowSelection = False
        
        .Redraw = flexRDBuffered
        
        'ˢ�¶�λ
        If strLocationName <> "" Then
            strLocationName = UCase(strLocationName)
            For j = 0 To .Rows - 1
                If UCase(.TextMatrix(j, 2)) = strLocationName Then .Row = j: Call .ShowCell(j, 2): Exit For
            Next
        Else
            If .Rows > 1 Then .Row = 1
        End If
'        'ˢ���޸ġ�ɾ����ť״̬
'        vsfFileList_SelChange

        If .Visible = True Then .SetFocus
'         Call SetMenu
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Sub
'==============================================================================
'=���ܣ� ��ʾ��¼����Ϣ
'==============================================================================
Public Sub SetMenu()
    If mintfgMainTag = 0 Then
        frmMDIMain.stbThis.Panels(2).Text = "�б��й���ʾ��" & vsfFileList.Rows - 1 & "�����ݡ�"
    Else
        frmMDIMain.stbThis.Panels(2).Text = "�б��й���ʾ��" & vsfExpFileList.Rows - 1 & "�����ݡ�"
    End If
End Sub

'==============================================================================
'=���ܣ� �����Ƿ����±���߱��Ƿ����
'==============================================================================
Private Function CheckTable() As Boolean
    On Error GoTo errH
    Dim strSQL As String
    Dim i As Integer
    Dim blnUse As Boolean
    strSQL = "select * from zlFilesUpgrade where rownum =1"

    Call OpenRecordset(mrsTemp, strSQL, Me.Caption)
    If mrsTemp.RecordCount >= 0 Then
        For i = 1 To mrsTemp.Fields.Count
            If mrsTemp.Fields.Item(i - 1).name = "����ϵͳ" Then
                blnUse = True
                Exit For
            End If
        Next

        If blnUse Then
            CheckTable = True
        Else
            MsgBox "��zlFilesUpgrade����,û���ҵ���Ӧ���ֶ�!" & vbCrLf & "�����ṹ�Ƿ�Ϊ����!", vbInformation
            CheckTable = False
        End If
    End If
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

'��ȡ�汾��ֱ����ʾֵ
Private Function GetFileVision(ByVal strVision As String) As String
    Dim lng�汾�� As Variant
    Dim str�汾�� As String
    If Len(strVision) > 0 Then
        lng�汾�� = strVision
        str�汾�� = Int(lng�汾�� / 10 ^ 8)
        If Len(lng�汾��) > 9 Then
            lng�汾�� = Right(lng�汾��, 9) Mod (10 ^ 8)
        Else
            lng�汾�� = lng�汾�� Mod (10 ^ 8)
        End If

        str�汾�� = str�汾�� & "." & Int(lng�汾�� / 10 ^ 4)
        lng�汾�� = lng�汾�� Mod 10 ^ 4
        str�汾�� = str�汾�� & "." & lng�汾��
        GetFileVision = str�汾��
    End If
End Function

Private Function GetCommpentVersion(ByVal strFile As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡָ���ؼ��İ汾��
    '���:
    '����:
    '����:�ɹ�,���ذ汾��,���򷵻ؿ�
    '����:���˺�
    '����:2009-01-16 16:59:34
    '-----------------------------------------------------------------------------------------------------------
    Dim objFile As New FileSystemObject
    Dim strVer As String, varVersion As Variant

    err = 0: On Error Resume Next
    '��ȡ�ļ��汾��
    strVer = objFile.GetFileVersion(strFile)
    If err <> 0 Then
        err.Clear: err = 0
        GetCommpentVersion = ""
        Exit Function
    End If
    If Trim(strVer) <> "" Then
        varVersion = Split(strVer, ".")
        If UBound(varVersion) > 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(3)
        ElseIf UBound(varVersion) = 2 Then
            strVer = varVersion(0) & "." & varVersion(1) & "." & varVersion(2)
        End If
    End If
    GetCommpentVersion = strVer
End Function

Private Function GetSystemName(ByVal strNum As String) As String
'����ϵͳ��ţ���ö�Ӧϵͳ���ƣ���δ�ҵ�
On err GoTo errH
    Select Case strNum
        Case "1", "100"
            GetSystemName = "ҽԺϵͳ��׼��"
        Case "2", "200"
            GetSystemName = "���¹���ϵͳ"
        Case "3", "300"
            GetSystemName = "��������ϵͳ"
        Case "4", "400"
            GetSystemName = "���ʹ�Ӧϵͳ"
        Case "5", "500"
            GetSystemName = "�������ϵͳ"
        Case "6", "600"
            GetSystemName = "�豸����ϵͳ"
        Case "7", "700"
            GetSystemName = "�ɱ�Ч�����ϵͳ"
        Case "21", "2100"
            GetSystemName = "������ϵͳ"
        Case "22", "2200"
            GetSystemName = "Ѫ�����ϵͳ"
        Case "23", "2300"
            GetSystemName = "Ժ�й���ϵͳ"
        Case "24", "2400"
            GetSystemName = "�������ϵͳ"
        Case "25", "2500"
            GetSystemName = "�ٴ��������ϵͳ"
        Case "26", "2600"
            GetSystemName = "������������ϵͳ"
    End Select
    Exit Function

errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If False Then
        Resume
    End If
End Function

'==============================================================================
'=�޸��ļ�
'==============================================================================
Private Sub StandardEdit()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    On Error GoTo errH
    Dim strLocationName As String
    strSysNum = 100
    
    vsfFileList.Row = vsfFileList.FindRow(m_strCurFileName, , 2)
    Call f.ShowForm("�޸�", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0", m_strCurSetupPathADD)
    Call RefreshData(m_strCurFileName)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub


'==============================================================================
'=�����ļ�
'==============================================================================
Private Sub StandardAdd()
    Dim f As New frmScriptEdit
    Dim strSysNum As String
    Dim strLocationName As String

    On Error GoTo errH
    strSysNum = 1
    
    strLocationName = f.ShowForm("����", m_strCurTypeName, m_strCurFileName, m_strCurSysNum, strSysNum, m_strCurVision, m_strCurSetupPath, m_strCurEditDate, m_strCurSysOption, m_strCurFileExplanation, m_strCurSellFile, m_blnCurReg, m_blnCurUpData, "0", m_strCurSetupPathADD)
    If strLocationName = "" Then
        Call RefreshData(m_strCurFileName)
    Else
        Call RefreshData(strLocationName)
    End If
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=ɾ���ļ�
'==============================================================================
Private Sub StandardDel()
    Dim i         As Long
    Dim strName   As String
    Dim rs        As ADODB.Recordset
    Dim strSQL    As String
    Dim strSys    As String
    Dim strSysNum As String
    Dim lngRow    As Long
    Dim lngCurRow As Long
    Dim frmMsgbox As New frmMessageBox
    On Error GoTo errH

    Select Case mintfgMainTag
        Case 0 '���ò���ɾ��
            If vsfFileList.SelectedRows = 0 Then Exit Sub
            If m_strCurTypeName <> "��������" Then Exit Sub
            
            If vsfFileList.SelectedRows = 1 Then
'                If MsgBox("��ȷ��Ҫɾ��" & m_strCurFileName & "������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
                If frmMsgbox.ShowMe(2, gstrSysName, "��ȷ��Ҫɾ��" & m_strCurFileName & "������") = False Then Exit Sub
            Else
'                If MsgBox("��ȷ��Ҫɾ��ѡ���" & vsfFileList.SelectedRows & "��������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
                If frmMsgbox.ShowMe(2, gstrSysName, "��ȷ��Ҫɾ��ѡ���" & vsfFileList.SelectedRows & "��������") = False Then Exit Sub
            End If
'            gcnOracle.BeginTrans
            lngRow = vsfFileList.FindRow(CStr(m_strCurFileName), , 2)
            For i = 0 To vsfFileList.SelectedRows
                If vsfFileList.SelectedRow(i) > 0 Then
                    lngCurRow = vsfFileList.SelectedRow(i)
                    If vsfFileList.TextMatrix(lngCurRow, FC_�ļ�����) = "��������" Then
                        strName = UCase(IIf(Len(vsfFileList.Cell(flexcpText, lngCurRow, 2)) = 0, 0, vsfFileList.Cell(flexcpText, lngCurRow, 2)))
                        
                        gstrSQL = "delete zlFilesUpgrade where upper(�ļ���)= upper('" & strName & "')"
                        gcnOracle.Execute gstrSQL
                    End If
                End If
            Next
'            gcnOracle.CommitTrans

            If lngRow <> -1 Then
                If lngRow >= 2 And vsfFileList.Rows > 2 Then
                  vsfFileList.Select lngRow - 1, 2
                  vsfFileList.ShowCell lngRow - 1, 2
                End If
            End If
            
        Case 1 '���ò���ɾ��
         If vsfExpFileList.SelectedRows = 1 Then
                If MsgBox("��ȷ��Ҫɾ��" & vsfExpFileList.Cell(flexcpText, vsfExpFileList.Row, 1) & "���ò�����", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
            Else
                If MsgBox("��ȷ��Ҫɾ��ѡ���" & vsfExpFileList.SelectedRows & "���ø�������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
            End If
        '    gcnOracle.BeginTrans
'            lngRow = vsfExpFileList.FindRow(CStr(vsfExpFileList.Cell(flexcpText, lngCurRow, 1)), , 2)
            For i = 0 To vsfExpFileList.SelectedRows
                If vsfExpFileList.SelectedRow(i) Then
                    lngCurRow = vsfExpFileList.SelectedRow(i)
                    If lngCurRow <> -1 Then
                        gstrSQL = "delete zlfilesexpired where upper(�ļ���)= upper('" & Trim(vsfExpFileList.Cell(flexcpText, lngCurRow, 1)) & "')"
                        gcnOracle.Execute gstrSQL
                    End If
                End If
            Next
        '    gcnOracle.CommitTrans
        End Select
        Call RefreshData
    Exit Sub
errH:
'    gcnOracle.RollbackTrans
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

'==============================================================================
'=�����ļ�
'==============================================================================
Private Sub StandardAbandon()
    Dim strName   As String
    Dim lngRow    As Long
    Dim i As Long
    Dim lngCurRow As Long
    Dim frmMsgbox As New frmMessageBox
    
    lngRow = vsfFileList.FindRow(CStr(m_strCurFileName), , 2)
    
    If vsfFileList.SelectedRows > 1 Then
        If frmMsgbox.ShowMe(1, gstrSysName, "��ȷ��Ҫ����ѡ��� " & vsfFileList.SelectedRows & " ��������") = False Then Exit Sub
'        If MsgBox("��ȷ��Ҫ����ѡ��� " & vsfFileList.SelectedRows & " ��������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
    Else
        If frmMsgbox.ShowMe(1, gstrSysName, "��ȷ��Ҫ���� " & vsfFileList.TextMatrix(lngRow, 2) & " ������") = False Then Exit Sub
'        If MsgBox("��ȷ��Ҫ���� " & vsfFileList.TextMatrix(lngRow, 2) & " ��������", vbQuestion + vbOKCancel, gstrSysName) = vbCancel Then Exit Sub
    End If
    
    For i = 0 To vsfFileList.SelectedRows
        If vsfFileList.SelectedRow(i) > 0 Then
            lngCurRow = vsfFileList.SelectedRow(i)
            If vsfFileList.TextMatrix(lngCurRow, FC_�ļ�����) = "��������" Then
                strName = IIf(Len(vsfFileList.Cell(flexcpText, lngCurRow, 2)) = 0, 0, vsfFileList.Cell(flexcpText, lngCurRow, 2))
                strName = UCase(strName)
                gstrSQL = "  Insert Into Zlfilesexpired (�ļ���,��װ·��,ϵͳ���,ϵͳ�汾,˵��) select " & _
                                "'" & vsfFileList.Cell(flexcpText, lngCurRow, 2) & "','" & vsfFileList.Cell(flexcpText, lngCurRow, 7) & "','" & vsfFileList.Cell(flexcpData, lngCurRow, 5) & "','" & vsfFileList.Cell(flexcpText, lngCurRow, 3) & "','" & vsfFileList.Cell(flexcpText, lngCurRow, 10) & "' " & _
                                " from dual Where Not Exists " & _
                                " (Select 1 From Zlfilesexpired A Where A.�ļ��� = '" & vsfFileList.Cell(flexcpText, lngCurRow, 2) & "')"
                gcnOracle.Execute gstrSQL
                gstrSQL = "delete zlFilesUpgrade where upper(�ļ���)= upper('" & strName & "')"
                gcnOracle.Execute gstrSQL
            End If
        End If
    Next
    
    Call RefreshData
    If lngRow <> -1 Then
        If lngRow >= 2 And vsfFileList.Rows > 2 Then
          vsfFileList.Select lngRow - 1, 2
          vsfFileList.ShowCell lngRow - 1, 2
        End If
    End If
    Exit Sub
End Sub

'==============================================================================
'=�ϴ��ļ�
'==============================================================================
Private Sub StandardUpLoad()
    Dim frmUpload As New frmFilesUpload
    
    frmUpload.ShowMe

End Sub

Private Sub LblBtn_Click(Index As Integer)
    Select Case Index
         Case 0 '���ò���
            If mintfgMainTag = 1 Then
                LblBtn.Item(0).BackColor = COLOR_SELECT
                LblBtn.Item(1).BackColor = COLOR_NOT_SELECT
                vsfFileList.Visible = True
                vsfExpFileList.Visible = False
    
                cmdAdd.Enabled = True
                chkFilter.Visible = True
                
                mintfgMainTag = 0
                
                RefreshData

    '            LblBtn.Item(0).Move LblBtn.Item(0).Left, LblBtn.Item(0).Top - 45, LblBtn.Item(0).Width, LblBtn.Item(0).Height + 45
    '            LblBtn.Item(1).Move LblBtn.Item(1).Left, LblBtn.Item(1).Top + 45, LblBtn.Item(1).Width, LblBtn.Item(1).Height - 45
            End If
         Case 1 '���ò���
            If mintfgMainTag = 0 Then
                LblBtn.Item(1).BackColor = COLOR_SELECT
                LblBtn.Item(0).BackColor = COLOR_NOT_SELECT
                vsfFileList.Visible = False
                vsfExpFileList.Visible = True
                
                cmdAdd.Enabled = False
                cmdEdit.Enabled = False
                cmdExpired.Enabled = False
                cmdDel.Enabled = True
                chkFilter.Visible = False
                
                mintfgMainTag = 1

                RefreshData

                If vsfExpFileList.Rows <= vsfExpFileList.FixedRows Then cmdDel.Enabled = False
    '            LblBtn.Item(0).Move LblBtn.Item(0).Left, LblBtn.Item(0).Top + 45, LblBtn.Item(0).Width, LblBtn.Item(0).Height - 45
    '            LblBtn.Item(1).Move LblBtn.Item(1).Left, LblBtn.Item(1).Top - 45, LblBtn.Item(1).Width, LblBtn.Item(1).Height + 45
            End If
    End Select
End Sub

Private Sub LblBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If LblBtn.Item(Index).BackColor = COLOR_NOT_SELECT Then
        LblBtn.Item(Index).BackColor = COLOR_MOVE
    End If
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = txtFind.Tag Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    End If
End Sub

'==============================================================================
'=���ٶ�λ
'==============================================================================
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long

    On Error GoTo errH

    lngRow = 0
    If txtFind.Locked Then Exit Sub
    If mstrFindKey = "����" Then mstrFindKey = "�ļ�����"
    If KeyAscii = vbKeyReturn Then
        '��ȡ���ڵ�ǰ�еļ�¼����
        For lngLoop = vsfFileList.Row + 1 To vsfFileList.Rows - 1
            If InStr(UCase(vsfFileList.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '��ȡС�ڵ�ǰ�еļ�¼����
        If lngRow = 0 Then
            For lngLoop = 0 To vsfFileList.Row
                If InStr(UCase(vsfFileList.TextMatrix(lngLoop, 2)), UCase(txtFind.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If vsfFileList.Rows > 1 And lngRow >= 1 Then
            vsfFileList.Row = lngRow
            vsfFileList.ShowCell lngRow, 2
        End If
        'Call LocationObj(txtFind)
    End If
    If mstrFindKey = "�ļ�����" Then mstrFindKey = "����"

    Exit Sub
errH:
    mstrFindKey = "����"
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = txtFind.Tag
        txtFind.ForeColor = vbGrayText
    End If
End Sub

Private Sub InitVsfMain()
With vsfExpFileList
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 50
        .Clear
        .Cols = 14
'        Exit Sub
        .Cell(flexcpText, 0, 0) = "���"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "�ļ�����"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "�ļ���"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .ColWidth(2) = 2200
        .Cell(flexcpText, 0, 3) = "�汾��"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .ColWidth(3) = 1200
        .Cell(flexcpText, 0, 4) = "�޸�����"
        .Cell(flexcpAlignment, 0, 4) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 5) = "����ϵͳ"
        .Cell(flexcpAlignment, 0, 5) = flexAlignCenterCenter
        .ColWidth(5) = 1800
        .Cell(flexcpText, 0, 6) = "ҵ�񲿼�"
        .Cell(flexcpAlignment, 0, 6) = flexAlignCenterCenter
        .ColWidth(6) = 3000

        .Cell(flexcpText, 0, 7) = "��װ·��"
        .Cell(flexcpAlignment, 0, 7) = flexAlignCenterCenter
        .ColWidth(7) = 0

        .Cell(flexcpText, 0, 8) = "����ID"
        .Cell(flexcpAlignment, 0, 8) = flexAlignCenterCenter
        .ColWidth(8) = 0

        .Cell(flexcpText, 0, 9) = "��װ·��"
        .Cell(flexcpAlignment, 0, 9) = flexAlignCenterCenter
        .ColWidth(9) = 2000

        .Cell(flexcpText, 0, 10) = "ϵͳ����"
        .Cell(flexcpAlignment, 0, 10) = flexAlignCenterCenter
        .ColWidth(10) = 0
        .Cell(flexcpText, 0, 11) = "�Զ�ע��"
        .Cell(flexcpAlignment, 0, 11) = flexAlignCenterCenter
        .ColWidth(11) = 1000

        .Cell(flexcpText, 0, 12) = "�ļ�˵��"
        .Cell(flexcpAlignment, 0, 12) = flexAlignCenterCenter
        .ColWidth(12) = 5000

        .Cell(flexcpText, 0, 13) = "ǿ�Ƹ���"
        .Cell(flexcpAlignment, 0, 13) = flexAlignCenterCenter
        .ColWidth(13) = 0
        
        .ExtendLastCol = True
'        .ScrollTips = True
'        .FocusRect = flexFocusSolid

        .RowHeightMin = 300
        .RowHeightMax = 300
        '���������
        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("�ͻ�������")
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False
    End With
End Sub

Public Sub RefreshData(Optional strLocationFileName As String = "")
    Select Case mintfgMainTag
        Case 0
            If chkFilter.value = 1 Then
                Call LoadFilesList("4", strLocationFileName)
            Else
                Call LoadFilesList(, strLocationFileName)
            End If
        Case 1
            LoadExpFilesList
    End Select
    SetMenu
End Sub
