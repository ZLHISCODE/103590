VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBalanceBat 
   Caption         =   "������;����"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBalanceBat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11820
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   1320
      Left            =   -15
      ScaleHeight     =   1320
      ScaleWidth      =   11730
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6540
      Width           =   11730
      Begin VB.CommandButton cmdOK 
         Caption         =   "����(&O)"
         Default         =   -1  'True
         Height          =   400
         Left            =   8685
         TabIndex        =   13
         Top             =   825
         Width           =   1400
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&C)"
         Height          =   400
         Left            =   10200
         TabIndex        =   14
         Top             =   825
         Width           =   1400
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   360
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9675
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   1905
      End
      Begin MSMask.MaskEdBox txtDateEnd 
         Height          =   360
         Left            =   375
         TabIndex        =   7
         Top             =   0
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-mm-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   435
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   825
         Visible         =   0   'False
         Width           =   8460
         _cx             =   14922
         _cy             =   767
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBalanceBat.frx":617A
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         ExplorerBar     =   3
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
      Begin VB.Label lblDateEnd 
         Caption         =   "��                     ֮ǰ�ķ��ý���"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   60
         Width           =   4440
      End
      Begin VB.Label lbl���㷽ʽ 
         Caption         =   "���㷽ʽ"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
         Height          =   240
         Left            =   8880
         TabIndex        =   10
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Caption         =   "�����n�����˽���"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   480
         Width           =   8295
      End
   End
   Begin VB.Frame fra 
      Height          =   645
      Left            =   90
      TabIndex        =   15
      Top             =   0
      Width           =   11685
      Begin VB.ComboBox cboInsure 
         Height          =   360
         Left            =   4095
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   195
         Width           =   3885
      End
      Begin VB.ComboBox cboʹ����� 
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label lblInsure 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   240
         Left            =   3135
         TabIndex        =   17
         Top             =   270
         Width           =   960
      End
      Begin VB.Label lblRpt 
         AutoSize        =   -1  'True
         Caption         =   "sss"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   8085
         TabIndex        =   2
         Top             =   300
         Width           =   405
      End
      Begin VB.Label lblʹ����� 
         Caption         =   "ʹ�����"
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   960
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   4860
      Left            =   2160
      TabIndex        =   4
      Top             =   675
      Width           =   2460
      _cx             =   4339
      _cy             =   8572
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
      BackColorSel    =   13627390
      ForeColorSel    =   0
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":6245
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   4860
      Left            =   4680
      TabIndex        =   5
      Top             =   690
      Width           =   7065
      _cx             =   12462
      _cy             =   8572
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
      BackColorSel    =   12640511
      ForeColorSel    =   0
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
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":628D
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsFeeType 
      Height          =   4875
      Left            =   120
      TabIndex        =   3
      Top             =   675
      Width           =   1980
      _cx             =   3492
      _cy             =   8599
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
      BackColorSel    =   15790320
      ForeColorSel    =   0
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceBat.frx":63DF
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
      ExplorerBar     =   1
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
End
Attribute VB_Name = "frmBalanceBat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPatis As String '���ڼ�¼ѡ��Ŀ����±��Ϊ�����ʵĲ���ID
Private mlng����ID As Long
Private mrsRptFormat As ADODB.Recordset
Private mobjInvoice As clsInvoice
Private mobjFact As clsFactProperty
Private mblnNotClick As Boolean
Private mlngPreInsure As Long
Private mlngModul As Long
Private mstrPrivs As String 'Ȩ�޴�
'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    �ֱҴ��� As Boolean
    δ�����Ժ As Boolean
    ����ʹ�ø����ʻ� As Boolean
    ��Ժ��������Ժ As Boolean
    ��Ժ���˽������� As Boolean
    ��;������������ϴ����� As Boolean
    �������ú���ýӿ� As Boolean
    �������Ϻ��ӡ�ص� As Boolean
    סԺ�������� As Boolean
    ������;���� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

'3.3 ģ���������
Private Type Ty_ModulePara
     blnZero  As Boolean '����ʱ�Ƿ��������
     int����ʱ�� As Integer '0-���Ǽ�ʱ��,1-������ʱ��
End Type
Private mty_ModulePara As Ty_ModulePara
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������(������;����)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-07 09:52:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mblnOK = False
    mstrPrivs = strPrivs
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    ShowMe = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mty_ModulePara
         .blnZero = zlDatabase.GetPara("���������", glngSys, mlngModul) = "1"
         .int����ʱ�� = IIf(zlDatabase.GetPara("���ʷ���ʱ��", glngSys, mlngModul) = "1", 1, 0)
    End With
End Sub
Private Sub LoadInsureType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ر������
    '����:���˺�
    '����:2015-03-25 14:26:17
    '����:81661
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSelect As String, strSql As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(mstrPrivs, "ҽ��������;����") = False Then
        cboInsure.Visible = False: lblInsure.Visible = False
        lblRpt.Left = lblInsure.Left
        Exit Sub
    End If
    If Not gclsInsure Is Nothing Then
        strSelect = gclsInsure.GetAvailabilityInsures
    End If
    mblnNotClick = True
    cboInsure.Clear
    cboInsure.AddItem ""
    cboInsure.ItemData(cboInsure.NewIndex) = 0
    cboInsure.ListIndex = 0
    If InStr(strSelect, ",") = 0 And Val(strSelect) = 0 Then Exit Sub
    strSql = "" & _
    "   Select A.���,A.����,A.˵��,Nvl(A.���,0) AS ���" & _
    "   From ������� A " & _
    "   Where Nvl(�Ƿ��ֹ,0)=0 " & _
    "       And A.��� in (Select Column_value From Table(f_Num2List([1])))" & _
    "   Order By A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strSelect)
    With rsTemp
        Do While Not .EOF
            cboInsure.AddItem "" & rsTemp!��� & "-" & rsTemp!����
            cboInsure.ItemData(cboInsure.NewIndex) = Val(NVL(rsTemp!���))
            .MoveNext
        Loop
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotClick = False
End Sub
Private Sub cboInsure_Click()
    Dim lngDeptID As Long
    Dim lngInsure As Long
    If mblnNotClick Then Exit Sub
    lngInsure = cboInsure.ItemData(cboInsure.ListIndex)
    If mlngPreInsure = lngInsure Then Exit Sub  '��ͬʱ,���ı�
    mlngPreInsure = lngInsure
    If vsDept.Row > 0 Then lngDeptID = Val(vsDept.RowData(vsDept.Row))
    vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
    mstrPatis = ""
    Call LoadPati(lngDeptID)
End Sub

Private Sub cboʹ�����_Click()
    Dim lngDeptID As Long
    If mblnNotClick Then Exit Sub
    
    lblRpt.Caption = ""
    If mrsRptFormat Is Nothing Then Exit Sub
    mrsRptFormat.Filter = "���=" & cboʹ�����.ItemData(cboʹ�����.ListIndex)
    If Not mrsRptFormat.EOF Then
        lblRpt.Caption = NVL(mrsRptFormat!˵��)
    End If
    mlng����ID = 0
    Call InitFact(cboʹ�����.Text)
    Call RefreshFact
    If vsDept.Row > 0 Then lngDeptID = Val(vsDept.RowData(vsDept.Row))
    vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
    mstrPatis = ""
    Call LoadPati(lngDeptID)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, m As Long, blnPrint As Boolean
    Dim rsPati As ADODB.Recordset
    
    For i = 1 To vsDept.Rows - 1
        If vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked Then
            m = m + 1
        End If
    Next
    If m = vsDept.Rows - 1 Then
        MsgBox "������ѡ��һ������.", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set rsPati = GetPatiSet
    If rsPati.RecordCount = 0 Then
        MsgBox "������ѡ��һ������.", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not IsDate(txtDateEnd.Text) Then
        MsgBox "���ý�ֹʱ���ʽ����ȷ.", vbInformation, gstrSysName
        txtDateEnd.SetFocus
        Exit Sub
    End If
    
    blnPrint = mobjFact.��ӡ��ʽ <> 0
    If mobjFact.��ӡ��ʽ = 2 Then
        If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
    End If
    
    If blnPrint Then
        If mobjFact.�ϸ���� Then    '�ϸ�Ʊ�ݹ���
            If Trim(txtInvoice.Text) = "" Then
                MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
            If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.�û���, mobjFact.Ʊ��, mobjFact.ʹ�����, mlng����ID, mobjFact.��������ID, mlng����ID, 1, Trim(txtInvoice.Text)) = False Then Exit Sub
            If mlng����ID <= 0 Then
                Select Case mlng����ID
                    Case 0 '����ʧ��
                    Case -1
                        MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -2
                        MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Case -3
                        MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,����������", vbInformation, gstrSysName
                        txtInvoice.SetFocus
                End Select
                Exit Sub
            End If
        Else
            If Len(txtInvoice.Text) <> mobjFact.Ʊ�ų��� And txtInvoice.Text <> "" Then
                MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Sub
            End If
        End If
    End If
    
    If MsgBox("��ѡ����" & rsPati.RecordCount & "λ����,�������ν�����;����!" & _
        vbCrLf & "��׼���ú�ȷ��.", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
        Exit Sub
    End If
    
    cmdOK.Enabled = False
    Screen.MousePointer = 11
    Call SaveBalance(blnPrint, rsPati)
    Call LoadPati(Val(vsDept.RowData(vsDept.Row)))
    Screen.MousePointer = 0
    cmdOK.Enabled = True
    mblnOK = True
End Sub

Private Sub GetMaxMinDate(ByVal lngPatiID As Long, ByVal DatEnd As Date, ByRef DatMax As Date, ByRef DatMin As Date)
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim strTable As String, strDateMode As String
    
    'Ҫ�͹���Zl_���ʷ��ü�¼_Patient�еĴ�������α�һ��,�������û�н��ʷ��õĽ��ʵ�.
    '�����:������������ʵ,��SQL����������,�����������ķ��ý��н���,��ʵ����Ӧ��ֻ���סԺ����,���,���β��ֻ�滻��סԺ���ü�¼
    strDateMode = IIf(mty_ModulePara.int����ʱ�� = 1, "����ʱ��", "�Ǽ�ʱ��")
    
    strSql = "" & _
    " Select Max(Maxʱ��) DatMax, Min(Minʱ��) DatMin" & vbNewLine & _
    " From ( Select Max(" & strDateMode & ") Maxʱ��, Min(" & strDateMode & ") Minʱ��" & vbNewLine & _
    "        From סԺ���ü�¼ A" & vbNewLine & _
    "        Where A.����id = [1] And A.����id Is Null And A.��¼״̬ <> 0 And Mod(��¼����, 10) In (2, 3) And" & vbNewLine & _
    "             " & strDateMode & " < [2] " & vbCrLf & _
    "             And Not Exists ( Select 1" & vbNewLine & _
    "                              From סԺ���ü�¼ B" & vbNewLine & _
    "                              Where B.NO = A.NO And B.��¼���� = A.��¼���� And B.��� = A.���" & vbNewLine & _
    "                              Group By B.NO, B.��¼����, B.���" & vbNewLine & _
    "                              Having Nvl(Sum(B.ʵ�ս��), 0) = Decode(" & IIf(gblnZero, 1, 0) & ", 1, 1 + Nvl(Sum(B.ʵ�ս��), 0), 0))" & vbNewLine & _
    "       Union All" & vbNewLine & _
    "       Select Max(" & strDateMode & ") Maxʱ��, Min(" & strDateMode & ") Minʱ��" & vbNewLine & _
    "       From " & zlGetFullFieldsTable("סԺ���ü�¼") & vbNewLine & _
    "       Where A.����id = [1] And A.����id Is Not Null And Mod(��¼����, 10) In (2, 3) And Nvl(A.ʵ�ս��, 0) <> Nvl(A.���ʽ��, 0) And" & vbNewLine & _
    "             " & strDateMode & " < [2]" & vbNewLine & _
    "       Group By A.NO, A.���, Mod(A.��¼����, 10), A.��¼״̬, A.ִ��״̬" & vbNewLine & _
    "       Having Nvl(Sum(A.ʵ�ս��), 0) - Nvl(Sum(A.���ʽ��), 0) <> 0)"


    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatiID, DatEnd)
    DatMax = NVL(rsTmp!DatMax, CDate(0))
    DatMin = NVL(rsTmp!DatMin, CDate(0))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDateStr(DatTmp As Date) As String
    GetDateStr = "To_Date('" & Format(DatTmp, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Function GetBalanceSum(ByVal Dat�տ�ʱ�� As Date) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������������Ľ�����Ϣ
    '����:���˺�
    '����:2015-07-07 16:26:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    strSql = "" & _
    " Select Decode(Mod(B.��¼����,10),1,0,C.����)  as ���," & _
    "       DECODE(mod(B.��¼����,10),1,'��Ԥ��', B.���㷽ʽ) as ���㷽ʽ," & _
    "       min(Decode(Mod(B.��¼����,10),1,0,C.����)) as ��������," & _
    "       Sum(B.��Ԥ��) ������" & vbNewLine & _
    " From ���˽��ʼ�¼ A, ����Ԥ����¼ B,���㷽ʽ C" & vbNewLine & _
    " Where A.�շ�ʱ�� = [1] And A.����Ա���� = [2] and B.���㷽ʽ=C.����(+) And A.ID = B.����id" & vbNewLine & _
    " Group By decode(Mod(B.��¼����,10),1,0,C.����),DECODE(mod(B.��¼����,10),1,'��Ԥ��', B.���㷽ʽ)" & _
    " order by ���"
    On Error GoTo errH
    Set GetBalanceSum = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Dat�տ�ʱ��, UserInfo.����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SaveBalance(ByRef blnPrint As Boolean, ByRef rsPati As ADODB.Recordset)
    Dim strNO As String, i As Long, j As Long
    Dim lng��ҳID As Long, lng����ID As Long, lng����ID As Long
    Dim dtEndDate As Date, dtStartDate As Date, dtBalanceDate As Date
    Dim intCol As Integer
    Dim arrSQL As Variant, lngNum As Long, blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strPreBalance As String 'Ԥ������Ϣ:
    Dim intInsure As Integer, cllPro As Collection
    Dim strSql As String
    Dim str������� As String
    Dim intScreenMouse As Integer
    
    vsBalance.Visible = False
    intScreenMouse = Screen.MousePointer
    Err = 0: On Error GoTo ErrHand:
    dtBalanceDate = zlDatabase.Currentdate '��¼Ϊͳһ�Ľ���ʱ��

    Set cllPro = New Collection
    
    For i = 1 To rsPati.RecordCount
        arrSQL = Array()
        lng����ID = rsPati!����ID
        lng��ҳID = Val(NVL(rsPati!��ҳID))
        str������� = NVL(rsPati!��������)
        Call GetMaxMinDate(lng����ID, CDate(txtDateEnd.Text), dtEndDate, dtStartDate)
        
        If Not (dtEndDate = dtStartDate And dtEndDate = CDate(0)) Then 'û�д�����ò�����
            lblInfo.Caption = "��ǰ����:��" & rsPati.RecordCount & "λ,���ڽ��е�" & i & "λ," & rsPati!���� & ":" & rsPati!����
            Me.Refresh
            
            intInsure = Val(NVL(rsPati!����))
            
            If zlStr.IsHavePrivs(mstrPrivs, "ҽ��������;����") = False Then intInsure = 0
            
            MCPAR.������;���� = False
            If intInsure <> 0 Then
                '��ʼ������
                Call InitInsurePara(lng����ID, intInsure)
            End If
            
            If MCPAR.������;���� = False Then intInsure = 0
            If intInsure = 0 Then str������� = ""
            
            
            'ҽ��Ԥ����
            strPreBalance = "" '������ʽ|���||....
            If InsureBudgeting(lng����ID, lng��ҳID, intInsure, dtStartDate, dtEndDate, strPreBalance) = False Then GoTo GoNextPati:
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            strNO = zlDatabase.GetNextNo(15)
            
            '1.���˽��ʼ�¼
            '58758
            'Zl_���˽��ʼ�¼_Insert
            strSql = "Zl_���˽��ʼ�¼_Insert("
            '  Id_In           ���˽��ʼ�¼.Id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ���ݺ�_In       ���˽��ʼ�¼.No%Type,
            strSql = strSql & "'" & strNO & "',"
            '  ����id_In       ���˽��ʼ�¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  �շ�ʱ��_In     ���˽��ʼ�¼.�շ�ʱ��%Type,
            strSql = strSql & "" & GetDateStr(dtBalanceDate) & ","
            '  ��ʼ����_In     ���˽��ʼ�¼.��ʼ����%Type,
            strSql = strSql & "" & GetDateStr(dtStartDate) & ","
            '  ��������_In     ���˽��ʼ�¼.��������%Type,
            strSql = strSql & "" & GetDateStr(dtEndDate) & ","
            '  ��;����_In     ���˽��ʼ�¼.��;����%Type := 0,
            strSql = strSql & "1,"
            '  �ಡ�˽���_In   Number := 0,
            strSql = strSql & "0,"
            '  �����ʴ���_In Number := 0,
            strSql = strSql & "" & lng��ҳID & ","
            '  ��ע_In         ���˽��ʼ�¼.��ע%Type := Null,
            strSql = strSql & "NULL,"
            '  ��Դ_In         Number := 1,
            strSql = strSql & "2,"
            '  ԭ��_In         ���˽��ʼ�¼.ԭ��%Type := Null,
            strSql = strSql & "NULL,"
            '  ��������_In     ���˽��ʼ�¼.��������%Type := 2,
            strSql = strSql & "2,"
            '  ����״̬_In     ���˽��ʼ�¼.����״̬%Type := 0,
            strSql = strSql & "1,"
            '  סԺ����_In     ���˽��ʼ�¼.סԺ����%Type := Null,  'סԺ���������ʽ����Zl_���ʷ��ü�¼_Patient�����д���
            strSql = strSql & "" & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ","
            '  ���ʽ��_In     ���˽��ʼ�¼.���ʽ��%Type := Null
            strSql = strSql & 0 & ")"
            zlAddArray cllPro, strSql
            
            '3.סԺ���ü�¼
            'Zl_���ʷ��ü�¼_Patient
            strSql = "Zl_���ʷ��ü�¼_Patient("
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ����id_In     סԺ���ü�¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ��ֹʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
            strSql = strSql & "" & GetDateStr(CDate(txtDateEnd.Text)) & ","
            '  ʱ��ģʽ_In Number, --1: ����ʱ�� , 0: �Ǽ�ʱ��
            strSql = strSql & "" & mty_ModulePara.int����ʱ�� & ","
            '  ����ý���_In Number
            strSql = strSql & "" & IIf(mty_ModulePara.blnZero, 1, 0) & ")"
            zlAddArray cllPro, strSql
            
            '4.������ʵĽ�����Ϣ
            'Zl_�������ʽ���_Update
            strSql = "Zl_�������ʽ���_Update("
            '  ����id_In     ������ü�¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSql = strSql & "" & lng����ID & ","
            '  ���ս���_In   Varchar2,
            strSql = strSql & "" & IIf(strPreBalance = "", "null", "'" & strPreBalance & "'") & ","
            '  �������_In   �������.����%Type,
            strSql = strSql & "" & IIf(str������� = "", "null", "'" & str������� & "'") & ","
            '  ֧����ʽ_In   ���㷽ʽ.����%Type,
            strSql = strSql & "'" & cbo���㷽ʽ.Text & "',"
            '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
            strSql = strSql & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
            strSql = strSql & "'" & UserInfo.���� & "',"
            '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
            strSql = strSql & "" & GetDateStr(dtBalanceDate) & ","
            '  ��ɽ���_In Number:=0
            strSql = strSql & "" & IIf(intInsure <> 0, 0, 1) & ")"
            zlAddArray cllPro, strSql
            
            '4.��ʼƱ�ݺ�
            If blnPrint And Trim(txtInvoice.Text) <> "" Then
                strSql = "Zl_Ʊ����ʼ��_Update('" & strNO & "','" & Trim(txtInvoice.Text) & "',3)"
                zlAddArray cllPro, strSql
            End If
            
            On Error GoTo errH
            blnTrans = True
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            'ҽ������
            If InsureBalance(lng����ID, lng����ID, intInsure, strPreBalance, str�������, dtBalanceDate) Then
                gcnOracle.CommitTrans: blnTrans = False
                lngNum = lngNum + 1 '��¼ʵ�ʽ�������
                'Ʊ�ݴ�ӡ
                If blnPrint Then
                    mobjFact.LastUseID = mlng����ID
                    Call frmPrint.ReportPrint(1, strNO, lng����ID, mobjFact, txtInvoice.Text, dtBalanceDate, "", "", lng����ID, mobjFact.��ӡ��ʽ)
                    Call RefreshFact
                End If
            End If
            Set cllPro = New Collection
            blnTrans = False
        End If
GoNextPati:
        rsPati.MoveNext
    Next
    If lngNum = 0 Then
        lblInfo.Caption = "ѡ����" & rsPati.RecordCount & "λ����,����ָ���Ľ�ֹʱ��ǰ��������δ�����!"
        vsBalance.Visible = False
    Else
        lblInfo.Caption = "��" & rsPati.RecordCount & "λ������,����δ����õ�" & lngNum & "λ�������;����."
        vsBalance.Visible = True
        Set rsTmp = GetBalanceSum(dtBalanceDate)
        With vsBalance
            intCol = 0: .Cols = rsTmp.RecordCount * 2
            .Rows = 1
            Do While Not rsTmp.EOF
                .TextMatrix(0, intCol) = NVL(rsTmp!���㷽ʽ)
                .Cell(flexcpFontBold, 0, intCol) = True
                .TextMatrix(0, intCol + 1) = zlStr.FormatEx(Val(NVL(rsTmp!������)), 6)
                intCol = intCol + 2
                rsTmp.MoveNext
            Loop
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        End With
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Screen.MousePointer = intScreenMouse
        Resume
    End If
    
    Call SaveErrLog
    If lngNum > 0 Then
        lblInfo.Caption = "ѡ����" & rsPati.RecordCount & "λ����,ʵ�ʶ�" & lngNum & "λ�����������;����."
    End If
    Exit Sub
ErrHand:
     If ErrCenter = 1 Then
        Resume
     End If
End Sub
Private Function InsureBudgeting(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal intInsure As Integer, ByVal dtStartDate As Date, _
    ByVal dtEndDate As Date, ByRef strPreBalance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��Ԥ����
    '���: intInsure-����
    '����:strBalance-����Ԥ������Ϣ:������ʽ|���||....
    '����:Ԥ��ɹ�(����ͨ����δ��ҽ���������),����true,���򷵻�False
    '����:���˺�
    '����:2015-01-06 16:48:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln����ʱ�� As Boolean, strҽ���� As String
    Dim strBalance As String, varData As Variant, varTemp As Variant
    Dim strNotBalance As String '�����ڵĽ��㷽ʽ
    Dim lngRow As Long, blnOk As Boolean
    Dim cur�����ʻ� As Currency, curͳ��֧�� As Currency
    Dim curMoney As Currency
    Dim rsDetail As ADODB.Recordset
    
    Dim i As Long, byt״̬ As Byte, bytEditSta As Byte
    On Error GoTo errHandle
    strPreBalance = ""
    If intInsure = 0 Then InsureBudgeting = True: Exit Function
    
    bln����ʱ�� = mty_ModulePara.int����ʱ�� = 1 '0-���Ǽ�ʱ��,1-������ʱ��
    'ҽ��Ԥ����
    Set rsDetail = GetZYBalance_Insure(intInsure, lng����ID, _
         IIf(lng��ҳID = 0, "", CStr(lng��ҳID)), dtStartDate, dtEndDate, _
        MCPAR.��;������������ϴ�����, False, 0, "", "", "", "", bln����ʱ��)
    
    '���㷽ʽ;���;�Ƿ������޸�|...
    strBalance = gclsInsure.WipeoffMoney(rsDetail, lng����ID, strҽ����, "1", intInsure, "|1")
    varData = Split(strBalance, "|")
    
    strPreBalance = ""
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i) & ";;;;", ";")
        If varTemp(0) <> "" Then
            strPreBalance = strPreBalance & "||" & varTemp(0) & "|" & varTemp(1)
        End If
    Next
    If strPreBalance <> "" Then strPreBalance = Mid(strPreBalance, 3)
    InsureBudgeting = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function InsureBalance(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal intInsure As Integer, ByVal strԤ���� As String, ByVal str������� As String, ByVal dtBalanceDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������ӿ�
    '���:intInsure-����
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-07-06 15:30:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTruns As Boolean
    Dim cllPro As New Collection
    On Error GoTo errHandle
    
    'סԺҽ������
    If intInsure = 0 Then InsureBalance = True: Exit Function
    
    If Not gclsInsure.SettleSwap(lng����ID, intInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
    End If
    If strAdvance <> "" Then
        If zlInsure_Check(strԤ����, strAdvance) Then
            blnTruns = True
            Call ҽ�����ݸ���(lng����ID, lng����ID, strAdvance, str�������, dtBalanceDate, cllPro)
            zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
        End If
    End If
    InsureBalance = True
    Exit Function
errHandle:
    Call gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function ҽ�����ݸ���(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal strҽ������ As String, ByVal str������� As String, ByVal dtBalanceDate As Date, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������У�Ը���
    '����:У�Գɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-12 17:45:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    '4.������ʵĽ�����Ϣ
    'Zl_�������ʽ���_Update
    strSql = "Zl_�������ʽ���_Update("
    '  ����id_In     ������ü�¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSql = strSql & "" & lng����ID & ","
    '  ���ս���_In   Varchar2,
    strSql = strSql & "" & IIf(strҽ������ = "", "null", "'" & strҽ������ & "'") & ","
    '  �������_In   �������.����%Type,
    strSql = strSql & "" & IIf(str������� = "", "null", "'" & str������� & "'") & ","
    '  ֧����ʽ_In   ���㷽ʽ.����%Type,
    strSql = strSql & "'" & cbo���㷽ʽ.Text & "',"
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    strSql = strSql & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSql = strSql & "'" & UserInfo.���� & "',"
    '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    strSql = strSql & "" & GetDateStr(dtBalanceDate) & ","
    '  ��ɽ���_In Number:=0
    strSql = strSql & "1)"
    zlAddArray cllPro, strSql
    ҽ�����ݸ��� = True
End Function
Public Function zlInsure_Check(ByVal str���ս��� As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:str���ս���-���ս���
    '       strAdvance-ҽ�����صĽ���
    '����:
    '����:��Ҫ�϶�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    
    On Error GoTo errHandle
    If Not (strAdvance <> "" And str���ս��� <> strAdvance) Then Exit Function
    '��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    blnMedicareCheck = True
    varData = Split(str���ս���, "||"): varData1 = Split(strAdvance, "||")
    
    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsure_Check = blnMedicareCheck
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitInsurePara(ByVal lng����ID As Long, ByVal intInsure As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2015-03-25 17:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, intInsure)
    MCPAR.δ�����Ժ = gclsInsure.GetCapability(supportδ�����Ժ, lng����ID, intInsure)
    MCPAR.����ʹ�ø����ʻ� = gclsInsure.GetCapability(support����ʹ�ø����ʻ�, lng����ID, intInsure)
    MCPAR.��Ժ��������Ժ = gclsInsure.GetCapability(support��Ժ��������Ժ, lng����ID, intInsure)
    MCPAR.��;������������ϴ����� = gclsInsure.GetCapability(support��;������������ϴ�����, lng����ID, intInsure)
    MCPAR.�������ú���ýӿ� = gclsInsure.GetCapability(support����_�������ú���ýӿ�, lng����ID, intInsure)
    MCPAR.סԺ�������� = gclsInsure.GetCapability(supportסԺ��������, lng����ID, intInsure)
    MCPAR.������;���� = gclsInsure.GetCapability(support������;����, lng����ID, intInsure)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFact(ByVal strʹ����� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ʊ��Ϣ
    '����:���˺�
    '����:2015-02-05 11:26:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytInvoiceKind     As Byte, intFormat As Integer
    Dim intPrintMode As Integer, lngShareUseID As Long
    On Error GoTo errHandle
 
    
    bytInvoiceKind = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, 1137, "0"))

    Set mobjInvoice = New clsInvoice: Set mobjFact = New clsFactProperty
    mobjInvoice.zlInitCommon glngSys, gcnOracle, gstrDBUser
    
    Call mobjInvoice.zlGetInvoicePreperty(1137, IIf(bytInvoiceKind = 0, 3, 1), 0, 0, 0, mobjFact, , , 2)
    
    mobjFact.ʹ����� = strʹ�����
    mobjFact.Ʊ�� = IIf(bytInvoiceKind = 0, 3, 1)
    Call mobjInvoice.zlGetInvoicePrintFormat(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, intFormat, 2)
    mobjFact.��ӡ��ʽ = intFormat
    If mobjInvoice.zlGetInvoicePrintMode(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, intPrintMode) = False Then Exit Sub
    mobjFact.��ӡ��ʽ = intPrintMode
    If mobjInvoice.zlGetInvoiceShareID(1137, mobjFact.Ʊ��, mobjFact.ʹ�����, lngShareUseID) = False Then Exit Sub
    mobjFact.��������ID = lngShareUseID
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ���շ�Ʊ�ݺ�
    '����:���˺�
    '����:2015-02-05 11:40:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
       
    On Error GoTo errHandle
        
    If mobjFact.��ӡ��ʽ = 0 Then Exit Sub
    If Not mobjFact.�ϸ���� Then
        '���ϸ������
        '��ɢ��ȡ��һ������
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    If zlGetInvoiceGroupUseID(mlng����ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '�ϸ�ȡ��һ������
    If mobjInvoice.zlGetNextBill(1137, mlng����ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
    '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
    '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    txtInvoice.SelStart = Len(txtInvoice.Text)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetPatiSet() As ADODB.Recordset
    Dim strSql As String, str�ѱ� As String, strDeptIDs As String, i As Long
    Dim intInsure As Integer
    Dim strWhere As String
    
    str�ѱ� = Get�ѱ�ѡ��
    If str�ѱ� <> "" Then
        If UBound(Split(str�ѱ�, ",")) + 1 < vsFeeType.Rows - 1 Then
            str�ѱ� = "," & str�ѱ� & ","
            strWhere = " And Instr([2],','||A.�ѱ�||',') > 0"
        End If
    End If
    
    For i = 1 To vsDept.Rows - 1
        If Not (vsDept.Cell(flexcpChecked, i, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, i, 0) = flexTSUnchecked) Then
            strDeptIDs = strDeptIDs & "," & vsDept.RowData(i)
        End If
    Next
    strDeptIDs = Mid(strDeptIDs, 2)
    If UBound(Split(strDeptIDs, ",")) + 1 = vsDept.Rows - 1 Then strDeptIDs = ""
    
    If strDeptIDs <> "" Then
        strWhere = strWhere & " And B.����ID In(" & strDeptIDs & ")"
    End If
    
    If mstrPatis <> "" Then
        mstrPatis = "," & mstrPatis & ","
        strWhere = strWhere & " And Instr([1],','||B.����ID||',') = 0"
    End If
    
    intInsure = 0
    If cboInsure.ListIndex >= 0 Then intInsure = cboInsure.ItemData(cboInsure.ListIndex)
    strWhere = strWhere & IIf(intInsure = 0, "", " And nvl(A.����,0) =[4]")
    
    strSql = "" & _
    "Select Distinct C.���� as ����,A.����,A.����ID,A.סԺ����,A.��ҳID," & _
    "   A.סԺ��,nvl(A.����,0) as ����,J.���� as �������� " & vbNewLine & _
    "From ������Ϣ A, ��λ״����¼ B, ���ű� C,������ҳ M,������� J" & vbNewLine & _
    "Where A.����id = B.����ID   " & _
    "       And B.����ID = C.ID And A.����=J.���(+)  " & _
    "       And Zl_Billclass(A.����ID,A.��ҳID,A.����)=[3]  " & strWhere & vbNewLine & _
    "Order by ����,סԺ��"

    On Error GoTo errH
    Set GetPatiSet = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrPatis, str�ѱ�, Trim(cboʹ�����.Text), intInsure)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Load()
    Set mrsRptFormat = Nothing
    lblInfo.Caption = ""
    mlngModul = 1137
    
    Call zlInitModulePara
    Call LoadUseType    '����ʹ�����
    Call LoadInsureType '������Ч������
    Call InitFact(cboʹ�����.Text)
    txtInvoice.Locked = Not zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�") And gblnStrictCtrl '89302
    
    mblnNotClick = True
    If Not InitData Then Unload Me: mblnNotClick = False: Exit Sub
    If vsDept.Rows > 1 Then
        vsDept.Row = 1
    Else
        cmdOK.Enabled = False
    End If
    If vsDept.Row > 0 Then
        Call LoadPati(Val(vsDept.RowData(vsDept.Row)))
        lblRpt.Caption = ""
        If mrsRptFormat Is Nothing Then Exit Sub
        mrsRptFormat.Filter = "���=" & cboʹ�����.ItemData(cboʹ�����.ListIndex)
        If Not mrsRptFormat.EOF Then
            lblRpt.Caption = NVL(mrsRptFormat!˵��)
        End If
    End If
    mblnNotClick = False
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset, i As Long

    Set rsTmp = Get�ѱ�
    If rsTmp.RecordCount = 0 Then
        MsgBox "�ѱ�δ����,����ʹ�ô˹���!", vbInformation, gstrSysName
        Exit Function
    Else
        vsFeeType.Rows = rsTmp.RecordCount + 1
        vsFeeType.ColDataType(0) = flexDTBoolean
        vsFeeType.Cell(flexcpChecked, 1, 0, vsFeeType.Rows - 1, 0) = flexChecked
        vsFeeType.Row = 1: vsFeeType.Col = 1: vsFeeType.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsFeeType.TextMatrix(i, 1) = NVL(rsTmp!����)
        rsTmp.MoveNext
    Next
    Call LoadDept
    
    txtDateEnd.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    Set rsTmp = Get���㷽ʽ("����", 2)
    If rsTmp.RecordCount = 0 Then
        MsgBox "û���������ڽ��ʳ��ϵķ��ֽ���㷽ʽ,����ʹ�ô˹���!", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 1 To rsTmp.RecordCount
        cbo���㷽ʽ.AddItem rsTmp!����
        rsTmp.MoveNext
    Next
    cbo���㷽ʽ.ListIndex = 0
    
    Call RefreshFact
    
    InitData = True
End Function

Private Function Get�ѱ�() As ADODB.Recordset
    Dim strSql As String
 
    strSql = "Select ����,���� From �ѱ� Where ������� In (2, 3) And ���� = 1 Order by ����"
    On Error GoTo errH
    Set Get�ѱ� = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadDept()
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "" & _
    "Select A.ID, A.����" & vbNewLine & _
    "From ���ű� A, ��������˵�� B" & vbNewLine & _
    "Where A.ID = B.����id And B.������� In (2, 3) And B.�������� = '�ٴ�'" & vbNewLine & _
    "   And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & vbNewLine & _
    "   And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
    "   And Exists(Select 1 From ��λ״����¼ C Where C.����id Is Not Null And C.����id = A.ID) " & _
    " Order by ����"
    On Error GoTo errH
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    vsDept.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        vsDept.Cell(flexcpChecked, 1, 0, vsDept.Rows - 1, 0) = flexChecked
        vsDept.Row = 1: vsDept.Col = 1: vsDept.Col = 0
    End If
    For i = 1 To rsTmp.RecordCount
        vsDept.TextMatrix(i, 1) = rsTmp!����
        vsDept.RowData(i) = Val(rsTmp!ID)
        rsTmp.MoveNext
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get�ѱ�ѡ��() As String
    Dim i As Long, strTmp As String
    
    For i = 1 To vsFeeType.Rows - 1
        If vsFeeType.Cell(flexcpChecked, i, 0) = flexChecked Then strTmp = strTmp & "," & vsFeeType.TextMatrix(i, 1)
    Next
    Get�ѱ�ѡ�� = Mid(strTmp, 2)
End Function

Private Sub LoadPati(ByVal lngDeptID As Long)
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long, str�ѱ� As String
    Dim intInsure As Integer
    Dim strWhere As String
    
    str�ѱ� = Get�ѱ�ѡ��
    If str�ѱ� <> "" Then
        If UBound(Split(str�ѱ�, ",")) + 1 < vsFeeType.Rows - 1 Then
            str�ѱ� = "," & str�ѱ� & ","
            strSql = " And Instr([2],','||A.�ѱ�||',')>0"
        End If
    End If
    intInsure = 0
    If cboInsure.ListIndex >= 0 Then intInsure = cboInsure.ItemData(cboInsure.ListIndex)
    strWhere = IIf(intInsure = 0, "", " And nvl(A.����,0) =[4] ")
    
    On Error GoTo errH
    strSql = "" & _
    "   Select Distinct A.����ID,A.סԺ��, Nvl(D.����,A.����) as ����, Nvl(D.�Ա�,A.�Ա�) as �Ա�, " & _
    "           Nvl(D.����,A.����) as ����, B.������� δ�����, Ԥ����� ����Ԥ��, A.�ѱ�," & _
    "           nvl(A.����,0) as ����,M.���� as ��������" & vbNewLine & _
    "   From ������Ϣ A, ������� B,��λ״����¼ C,������ҳ D,������� M" & vbNewLine & _
    "   Where C.����id = [1]  " & strWhere & _
    "         And A.����id=D.����ID(+) And A.��ҳid = D.��ҳid(+) " & _
    "         And A.����id=C.����ID  And A.����id = B.����id(+) " & _
    "         And B.����(+) = 1  And B.����(+)=2 and A.����=M.���(+) " & _
    "         And Zl_Billclass(A.����ID, A.��ҳID, A.����)=[3] " & strSql & vbNewLine & _
    "   Order by A.סԺ��"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngDeptID, str�ѱ�, Trim(cboʹ�����.Text), intInsure)
    
    vsPati.Rows = 1 '�������,��������б�ͷ
    vsPati.Rows = rsTmp.RecordCount + 1
    If rsTmp.RecordCount > 0 Then
        If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
        Else
            vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
        End If
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    End If
    
    With vsPati
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, .ColIndex("סԺ��")) = "" & rsTmp!סԺ��
            .TextMatrix(i, .ColIndex("����")) = "" & rsTmp!����
            .TextMatrix(i, .ColIndex("�Ա�")) = "" & rsTmp!�Ա�
            .TextMatrix(i, .ColIndex("����")) = "" & rsTmp!����
            .TextMatrix(i, .ColIndex("δ�����")) = Format(Val(NVL(rsTmp!δ�����)), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("����Ԥ��")) = Format(Val(NVL(rsTmp!����Ԥ��)), "###0.00;-###0.00;;")
            .TextMatrix(i, .ColIndex("�ѱ�")) = "" & rsTmp!�ѱ�
            .TextMatrix(i, .ColIndex("����")) = Val(NVL(rsTmp!����))
            .TextMatrix(i, .ColIndex("�������")) = NVL(rsTmp!��������)
            .RowData(i) = Val(rsTmp!����ID)
            If Len(mstrPatis) > 0 Then
                If InStr("," & mstrPatis & ",", "," & rsTmp!����ID & ",") > 0 Then
                    .Cell(flexcpChecked, i, 0) = flexUnchecked
                End If
            End If
            If Val(NVL(rsTmp!����)) <> 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
            End If
            rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then .Row = 1: .Col = 1: .Col = 0
    End With
    Exit Sub
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.Width < 12060 Then Me.Width = 12060
    If Me.Height < 7635 Then Me.Height = 7635
    With fra
        .Width = ScaleWidth - .Left * 2
    End With
    With picDown
        .Width = ScaleWidth
        .Top = ScaleHeight - .Height - 100
    End With
     With vsFeeType
        .Height = picDown.Top - .Top - 50
        vsDept.Height = .Height
        vsPati.Height = .Height
        vsPati.Width = ScaleWidth - vsPati.Left - 50
     End With
     vsBalance.Width = cmdOK.Left - vsBalance.Left - 100
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRptFormat = Nothing
    mstrPatis = ""
    mlng����ID = 0
    Set mobjInvoice = Nothing
    Set mobjFact = Nothing

End Sub

Private Sub picDown_Resize()
  Err = 0: On Error Resume Next
    With cmdCancel
        .Left = picDown.ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = .Left - cmdOK.Width - 50
    End With
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    'If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed Then vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked    '�ֶ����ʱ��Ϊ��ɫ����Ϊѡ��
    
    If Row <> vsDept.Row Then vsDept.Row = Row
    If vsPati.Rows < 2 Then Exit Sub
    
    If vsDept.Cell(flexcpChecked, Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, Row, 0) = flexTSUnchecked Then
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexUnchecked
    Else
        vsPati.Cell(flexcpChecked, 1, 0, vsPati.Rows - 1, 0) = flexChecked
    End If
    Call SetPatiLists
End Sub
Private Sub vsdept_DblClick()
    If vsDept.MouseCol = 0 And vsDept.MouseRow = 0 Then
        Call SetVSAll(vsDept)
        Call vsDept_AfterEdit(vsDept.Row, vsDept.Col)
        mstrPatis = ""
    End If
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow And NewRow <> 0 Then Call LoadPati(Val(vsDept.RowData(NewRow)))
End Sub



Private Sub vsFeeType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    i = vsDept.Row
    vsDept.Row = 0
    vsDept.Row = i
    
End Sub

Private Sub vsPati_DblClick()
    If vsPati.MouseCol = 0 And vsPati.MouseRow = 0 Then
        If vsPati.Rows < 2 Then Exit Sub
        
        Call SetVSAll(vsPati)
        Call SetDeptState
        Call SetPatiLists
    End If
End Sub

Private Sub vsFeeType_DblClick()
    Dim i As Long
    If vsFeeType.MouseCol = 0 And vsFeeType.MouseRow = 0 Then
        Call SetVSAll(vsFeeType)
        i = vsDept.Row
        vsDept.Row = 0
        vsDept.Row = i
    End If
End Sub

Private Sub SetVSAll(ByRef vsf As VSFlexGrid)
    If vsf.Rows < 2 Then Exit Sub
    vsf.Cell(flexcpChecked, 1, 0, vsf.Rows - 1, 0) = IIf(Val(vsf.Tag) = 1, flexChecked, flexUnchecked)
    vsf.Tag = IIf(Val(vsf.Tag) = 0, 1, 0)
End Sub


Private Sub vsPati_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked Or vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSUnchecked Then
        SetPatiLists
    Else
        Call SetPatistr(Row)
    End If
    Call SetDeptState
End Sub

Private Sub SetPatistr(ByVal lngRow As Long)
'���ܣ���¼û��ѡ��Ĳ��ˣɣ�
    If vsPati.Cell(flexcpChecked, lngRow, 0) = flexUnchecked Then
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") = 0 Then
            If mstrPatis = "" Then
                mstrPatis = vsPati.RowData(lngRow)
            Else
                mstrPatis = mstrPatis & "," & vsPati.RowData(lngRow)
            End If
        End If
    Else
        If InStr("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",") > 0 Then
            mstrPatis = Replace("," & mstrPatis & ",", "," & vsPati.RowData(lngRow) & ",", ",")
            mstrPatis = Mid(mstrPatis, 2)   'ȥ��ǰ���
            If mstrPatis <> "" Then mstrPatis = Mid(mstrPatis, 1, Len(mstrPatis) - 1)
        End If
    End If
    If mstrPatis = "," Then mstrPatis = ""
End Sub

Private Sub SetPatiLists()
'����:��鵱ǰ�����б���û��ѡ��ļ��뵽�����У���ѡ��ģ��ӱ�����ɾ��
    Dim i As Long
    
    If vsPati.Rows < 2 Then Exit Sub
    
    For i = 1 To vsPati.Rows - 1
        Call SetPatistr(i)
    Next
End Sub

Private Function SetDeptState() As Boolean
'���ܣ����ÿ���ѡ��״̬
    Dim i As Long, m As Long
    
    For i = 1 To vsPati.Rows - 1
        If vsPati.Cell(flexcpChecked, i, 0) = flexChecked Then m = m + 1
    Next
    If m = vsPati.Rows - 1 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexChecked
    ElseIf m = 0 Then
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexUnchecked
    Else
        vsDept.Cell(flexcpChecked, vsDept.Row, 0) = flexTSGrayed
    End If
End Function

Private Sub vspati_EnterCell()
    If vsPati.Col = 0 Then
        vsPati.Editable = flexEDKbdMouse
    Else
        vsPati.Editable = flexEDNone
    End If
End Sub
Private Sub vsfeetype_EnterCell()
    If vsFeeType.Col = 0 Then
        vsFeeType.Editable = flexEDKbdMouse
    Else
        vsFeeType.Editable = flexEDNone
    End If
End Sub
Private Sub vsDept_EnterCell()
    If vsDept.Col = 0 Then
        vsDept.Editable = flexEDKbdMouse
    Else
        vsDept.Editable = flexEDNone
    End If
End Sub
Private Sub LoadUseType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʹ�����
    '����:���˺�
    '����:2011-04-28 15:09:10
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim i As Long, strSql As String
    Dim varData As Variant, varTemp As Variant
    Dim strRptName As String
    Dim strShareInvoice As String
    
    On Error GoTo errHandle
    
    strShareInvoice = zlDatabase.GetPara("���ʷ�Ʊ��ʽ", glngSys, 1137)
    varData = Split(strShareInvoice, "|")
    
    strRptName = IIf(gbytInvoiceKind = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    
    'Ʊ�ݸ�ʽ����
    strSql = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set mrsRptFormat = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    
    mblnNotClick = True
    strSql = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With cboʹ�����
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!����)
            .ItemData(.NewIndex) = 0
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!����)) Then
                    .ItemData(.NewIndex) = Val(varTemp(1))
                    Exit For
                End If
            Next
            rsTemp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjInvoice.zlGetInvoiceGroupID(1137, UserInfo.����, mobjFact.Ʊ��, _
        mobjFact.ʹ�����, lng����ID, mobjFact.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng����ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng����ID
        Case 0 '����ʧ��
        Case -1
            If Trim(mobjFact.ʹ�����) = "" Then
                MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "��û�����ú͹��õġ�" & mobjFact.ʹ����� & "������Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFact.ʹ�����) = "" Then
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "���صĹ���Ʊ�ݵġ�" & mobjFact.ʹ����� & "������Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

