VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFeeRefundmentMain 
   BorderStyle     =   0  'None
   Caption         =   "frmFeeRefundmengMain"
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   960
      ScaleHeight     =   1725
      ScaleWidth      =   4290
      TabIndex        =   7
      Top             =   3390
      Width           =   4290
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   780
         ScaleHeight     =   375
         ScaleWidth      =   3405
         TabIndex        =   8
         Top             =   930
         Width           =   3405
         Begin VB.TextBox txtSum 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   1185
         End
         Begin VB.ComboBox cboStyle 
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
            Height          =   360
            Left            =   615
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label lblBack 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�˿�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   30
            TabIndex        =   11
            Top             =   60
            Width           =   480
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   735
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   11160
         _cx             =   19685
         _cy             =   1296
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
         GridColor       =   -2147483633
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
         Rows            =   2
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundmengMain.frx":0000
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
   End
   Begin VB.PictureBox picBalanceStyle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   -60
      ScaleHeight     =   1950
      ScaleWidth      =   3045
      TabIndex        =   4
      Top             =   975
      Width           =   3045
      Begin VSFlex8Ctl.VSFlexGrid vsBalanceStyle 
         Height          =   1290
         Left            =   0
         TabIndex        =   5
         Top             =   135
         Width           =   2565
         _cx             =   4524
         _cy             =   2275
         Appearance      =   3
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeRefundmengMain.frx":00CB
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         ExplorerBar     =   0
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "��ǰת���ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   1605
         Width           =   1665
      End
   End
   Begin VB.PictureBox picInvoice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   5340
      ScaleHeight     =   3165
      ScaleWidth      =   2535
      TabIndex        =   2
      Top             =   2490
      Width           =   2535
      Begin VSFlex8Ctl.VSFlexGrid vsfInvoice 
         Height          =   1605
         Left            =   0
         TabIndex        =   3
         Top             =   45
         Width           =   2055
         _cx             =   3625
         _cy             =   2831
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeRefundmengMain.frx":0139
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   101
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
         ExplorerBar     =   0
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
   Begin VB.PictureBox picFee 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   3870
      ScaleHeight     =   2145
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   315
      Width           =   2010
      Begin VSFlex8Ctl.VSFlexGrid vsFee 
         Height          =   1470
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5490
         _cx             =   9684
         _cy             =   2593
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
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
         FormatString    =   ""
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
         ExplorerBar     =   0
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
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1500
      Top             =   630
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFeeRefundmentMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmFeeDetail As frmFeeDetail
Private mstrStyle As String, mlngModule As Long, mstrPrivs As String
Private mrsFeeList As ADODB.Recordset, mrsInfo As ADODB.Recordset, mrsBalance As ADODB.Recordset
Private mintType As Integer, mbln�������� As Boolean, mbln����תסԺ����� As Boolean
Private mblnSel As Boolean, mblnҩ����λ As Boolean, mint�շ��嵥 As Integer
Private mstrFindFpNo As String, mstrFindNO As String, mlng����ID As Long
Private mlngShareUseID As Long, mrsBalanceDup As ADODB.Recordset
Private mobjSquare As Object
Private mstrThreeSwapBalance As String
Private mstrThreeSwapCardType As String
Private mstrThreeSwapMoney As String
Private Enum mObjPancel
    Pan_BalanceInfo = 1
    Pan_Bill = 2
    Pan_List = 3
    Pan_Balance = 4
    Pan_Invoice = 5
End Enum

Public Sub InitMe(ByVal lngModule As Long, ByVal strPrivs As String, ByVal intTYPE As Integer)
    '-------------------------------------------------------------------------------------------------
    '����:�������,��ʼ��
    '���:
    '       lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '����:������
    '����:2014-06-18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mintType = intTYPE
    If mobjSquare Is Nothing Then Set mobjSquare = gobjSquare.objSquareCard
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
    mbln����תסԺ����� = IIf(Val(zlDatabase.GetPara("����תסԺ�����", glngSys, 1143, 0)) = 1, True, False)
    mint�շ��嵥 = 0: mblnҩ����λ = False
    If mintType = 1 Then
        mint�շ��嵥 = Val(zlDatabase.GetPara("�շ��嵥��ӡ��ʽ", glngSys, 1121))   '�����շ�
        mblnҩ����λ = zlDatabase.GetPara("ҩƷ��λ", glngSys, 1121) = "1"
        mlngShareUseID = Val(zlDatabase.GetPara("�����շ�Ʊ������", glngSys, mlngModule, "0"))
    Else
        mlngShareUseID = 0
    End If
    mblnSel = False
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_BalanceInfo
        Item.Handle = picBalanceStyle.hWnd
    Case Pan_Bill
        Item.Handle = picFee.hWnd
    Case Pan_List
        Item.Handle = mfrmFeeDetail.hWnd
    Case Pan_Balance
        Item.Handle = picBalance.hWnd
    Case Pan_Invoice
        Item.Handle = picInvoice.hWnd
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next

    zl_vsGrid_Para_Save mlngModule, vsFee, Me.Caption, IIf(mintType = 1, "�˷��б�", "�����б�"), True
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Caption, IIf(mintType = 1, "��ʷ�˷��б�", "��ʷ�����б�"), True
    zl_vsGrid_Para_Save mlngModule, vsBalanceStyle, Me.Caption, IIf(mintType = 1, "�˷ѽ�����Ϣ", "���ʽ�����Ϣ"), , True
    zl_vsGrid_Para_Save mlngModule, vsfInvoice, Me.Caption, IIf(mintType = 1, "�˷ѷ�Ʊ�б�", "���ʷ�Ʊ�б�"), True
    
    Unload mfrmFeeDetail
    Set mfrmFeeDetail = Nothing
    Set mrsFeeList = Nothing
    Set mrsInfo = Nothing
    Set mrsBalance = Nothing
End Sub

Private Sub Form_Load()
    Call InitPanel
    Call LoadStyle
    Call SetHeader
End Sub

Private Function InitBlanceData(ByVal strBalance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '���:strBalance-ָ���Ľ������,�Զ��ŷ���:'0001,0002
    '����:
    '����:
    '����:������
    '����:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Err = 0: On Error GoTo errHandle
    If mintType = 2 Then
        InitBlanceData = True
        Exit Function
    End If
    If strBalance = "" Then InitBlanceData = True: Exit Function
    
    strSql = _
    "Select a.���㷽ʽ, Nvl(b.����, 1) As ����, b.Ӧ����, a.���" & vbNewLine & _
    "From (Select Decode(a.��¼����, 3, a.���㷽ʽ, Null) As ���㷽ʽ, Sum(a.��Ԥ��) As ���" & vbNewLine & _
    "       From ����Ԥ����¼ A" & vbNewLine & _
    "       Where a.����id In (Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.����id" & vbNewLine & _
    "                        From ������ü�¼ C, ������ü�¼ D, (Select Distinct ����ID From ����Ԥ����¼ I,Table(f_Str2list([1])) J Where I.�������=J.Column_Value) E" & vbNewLine & _
    "                        Where c.����id = e.����id And c.No = d.No And Mod(d.��¼����, 10) = 1) And a.��¼���� In (1, 11, 3) And a.����id=[2] And" & vbNewLine & _
    "             Nvl(a.��Ԥ��, 0) <> 0" & vbNewLine & _
    "" & vbNewLine & _
    "       Group By Decode(a.��¼����, 3, a.���㷽ʽ, Null)) A, ���㷽ʽ B" & vbNewLine & _
    "Where a.���㷽ʽ = b.����(+)"
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Replace(strBalance, "'", ""), mlng����ID)
    Set mrsBalanceDup = mrsBalance
    InitBlanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SetPicBack(ByVal strBalance As String) As Boolean
    'vsBalance.Width = picBalance.Width - 4000
    'picBack.Left = vsBalance.Width + vsBalance.Left + 30
    picBack.Visible = True
    SetPicBack = True
End Function

Private Sub SetBlanceShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���㷽ʽ
    '���:blnAllSel-ѡ�����еĵ���
    '����:���˺�
    '����:2011-02-23 14:54:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, lngRow As Long, i As Long, str���� As String
    Dim blnȫѡ As Boolean, blnδѡ As Boolean, intCol As Integer
    Dim strFilter As String, bln�˿� As Boolean, rsTmp As ADODB.Recordset
    Dim strSelNos As String, strNO As String, strSql As String
    If mintType = 2 Then Exit Sub
    With vsFee
        blnȫѡ = True: blnδѡ = True
        For lngRow = 1 To .Rows - 1
            strBalance = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
            If .TextMatrix(lngRow, .ColIndex("ѡ��")) = "��" And Val(strBalance) <> 0 Then
                If InStr(1, strSelNos & ",", "," & strBalance & ",") = 0 Then
                    strSelNos = strSelNos & "," & strBalance
                    blnδѡ = False
                End If
            End If
             If InStr(1, strSelNos & ",", "," & strBalance & ",") = 0 Then blnȫѡ = False
        Next
    End With
    If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    bln�˿� = False
    
    '��ʾ����ѡ��ĵ��ݵĽ��㷽ʽ֮��
    If Not mrsBalance Is Nothing Then
        If blnȫѡ Or blnδѡ Then
            mrsBalance.Filter = ""
            If blnȫѡ Then bln�˿� = True
        Else
            strFilter = Replace(strSelNos, ",", "' Or �������='")
            strFilter = " �������=" & strFilter & ""
            'mrsBalance.Filter = strFilter
            bln�˿� = True
        End If
        If SetPicBack(strSelNos) = True Then
            txtSum.Text = InitPatialBalance(strSelNos)
        Else
            Call InitBlanceData(strSelNos)
        End If
        mrsBalance.Sort = "����,Ӧ����,���㷽ʽ"
        mrsBalanceDup.Sort = "����,Ӧ����,���㷽ʽ"
        vsBalance.Redraw = flexRDNone
        vsBalance.Clear 1
        vsBalance.Cols = 1
        If Not mrsBalanceDup.EOF Then
            For i = 1 To mrsBalanceDup.RecordCount
                If Val(NVL(mrsBalanceDup!���)) <> 0 Then
                    If NVL(mrsBalanceDup!���㷽ʽ, "��Ԥ��") <> strBalance Then
                        strBalance = NVL(mrsBalanceDup!���㷽ʽ, "��Ԥ��")
                        vsBalance.Cols = vsBalance.Cols + 2
                        vsBalance.ColAlignment(vsBalance.Cols - 2) = 7
                        vsBalance.ColAlignment(vsBalance.Cols - 1) = 1
                    End If
                    If mrsBalanceDup!���� <> 1 Then
                        vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '����
                        vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue
                    ElseIf bln�˿� Then
                        vsBalance.Cell(flexcpFontBold, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = True  '����
                        vsBalance.Cell(flexcpForeColor, 0, vsBalance.Cols - 1, 0, vsBalance.Cols - 2) = vbBlue  '��ɫ:�˿�
                    End If
                    vsBalance.TextMatrix(0, vsBalance.Cols - 2) = strBalance & ":"
                    vsBalance.TextMatrix(0, vsBalance.Cols - 1) = _
                        Val(vsBalance.TextMatrix(0, vsBalance.Cols - 1)) + NVL(mrsBalanceDup!���, 0)
                End If
                mrsBalanceDup.MoveNext
            Next
        End If
        intCol = 0
        strBalance = ""
        If Not mrsBalance.EOF Then
            For i = 1 To mrsBalance.RecordCount
                If Val(NVL(mrsBalance!���)) <> 0 And Val(NVL(mrsBalance!����)) <> 9 Then
                    If NVL(mrsBalance!���㷽ʽ, "��Ԥ��") <> strBalance Then
                        strBalance = NVL(mrsBalance!���㷽ʽ, "��Ԥ��")
                        intCol = intCol + 2
                        vsBalance.ColAlignment(intCol - 1) = 7
                        vsBalance.ColAlignment(intCol) = 1
                    End If
                    If mrsBalance!���� <> 1 Then
                        vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '����
                        vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = vbRed '��ɫ
                    ElseIf bln�˿� Then
                        vsBalance.Cell(flexcpFontBold, 1, intCol, 1, intCol - 1) = True '����
                        vsBalance.Cell(flexcpForeColor, 1, intCol, 1, intCol - 1) = vbRed '��ɫ:�˿�
                    End If
                    vsBalance.TextMatrix(1, intCol - 1) = strBalance & ":"
                    vsBalance.TextMatrix(1, intCol) = _
                    Val(vsBalance.TextMatrix(1, intCol)) + NVL(mrsBalance!���, 0)
                End If
                mrsBalance.MoveNext
            Next
        End If
        If strSelNos = "" Then
            For i = 1 To vsBalance.Cols - 1
                vsBalance.TextMatrix(1, i) = ""
            Next i
        End If
        
        Call vsBalance.AutoSize(0, vsBalance.Cols - 1)
        vsBalance.Row = vsBalance.FixedRows
        If vsBalance.Cols <> 1 Then vsBalance.Col = vsBalance.FixedCols
        'vsBalance.TextMatrix(0, 0) = "�տ����"
        vsBalance.Redraw = flexRDDirect
    End If
End Sub

Private Function InitPatialBalance(ByVal strBalance As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������˷ѵĽ�������
    '���:strBalance-ָ���Ľ������,�Զ��ŷ���:'0001,0002
    '����:
    '����:
    '����:������
    '����:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, dblSum As Double, i As Integer
    Err = 0: On Error GoTo errHandle
    If mintType = 2 Then
        InitPatialBalance = 0
        Exit Function
    End If
    If strBalance = "" Then InitPatialBalance = 0: Exit Function
    
    Call InitBlanceData(strBalance)
    Do While Not mrsBalance.EOF
        dblSum = dblSum + Val(NVL(mrsBalance!���))
        mrsBalance.MoveNext
    Loop
    'ȫ�˼�¼(Ԥ����)
    strSql = _
    "Select /*+ RULE*/" & vbNewLine & _
    " a.���㷽ʽ, Nvl(b.����, 1) As ����, b.Ӧ����, a.���" & vbNewLine & _
    "From (Select Decode(a.��¼����, 3, a.���㷽ʽ, Null) As ���㷽ʽ, Sum(a.��Ԥ��) As ���" & vbNewLine & _
    "       From ����Ԥ����¼ A," & vbNewLine & _
    "            (Select /*+ rule */" & vbNewLine & _
    "              Distinct d.����id" & vbNewLine & _
    "              From ������ü�¼ C, ������ü�¼ D," & vbNewLine & _
    "                   (Select Distinct ����id" & vbNewLine & _
    "                     From ����Ԥ����¼ I, Table(f_Str2list([1])) J" & vbNewLine & _
    "                     Where i.������� = j.Column_Value) E" & vbNewLine & _
    "              Where c.����id = e.����id And c.No = d.No And Mod(d.��¼����, 10) = 1 And Not Exists" & vbNewLine & _
    "               (Select 1" & vbNewLine & _
    "                     From ������ü�¼" & vbNewLine & _
    "                     Where ����id In (Select Max(����id)" & vbNewLine & _
    "                                    From ������ü�¼" & vbNewLine & _
    "                                    Where NO In ((Select Distinct k.No" & vbNewLine & _
    "                                                 From ������ü�¼ K, ����Ԥ����¼ L" & vbNewLine & _
    "                                                 Where l.������� In (Select Column_Value From Table(f_Str2list([1]))) And" & vbNewLine & _
    "                                                       k.����id = l.����id)) And Mod(��¼����, 10) = 1) And Mod(��¼����, 10) = 1 And" & vbNewLine & _
    "                           ��¼״̬ = 2)) K" & vbNewLine & _
    "       Where a.����id = k.����id And a.��¼���� In (1, 11) And a.����id = [2] And Nvl(a.��Ԥ��, 0) <> 0" & vbNewLine & _
    "       Group By Decode(a.��¼����, 3, a.���㷽ʽ, Null)) A, ���㷽ʽ B" & vbNewLine & _
    "Where a.���㷽ʽ = b.����(+) "

    'ȫ�˼�¼(���ѿ�)
    strSql = strSql & " Union " & _
    "Select a.���㷽ʽ, Nvl(b.����, 1) As ����, b.Ӧ����, a.���" & vbNewLine & _
    "From (Select Decode(a.��¼����, 3, a.���㷽ʽ, Null) As ���㷽ʽ, Sum(a.��Ԥ��) As ���" & vbNewLine & _
    "       From ����Ԥ����¼ A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.����id" & vbNewLine & _
    "                        From ������ü�¼ C, ������ü�¼ D, (Select Distinct ����ID From ����Ԥ����¼ I,Table(f_Str2list([1])) J Where I.�������=J.Column_Value) E" & vbNewLine & _
    "                        Where c.����id = e.����id And c.No = d.No And Mod(d.��¼����, 10) = 1) K" & _
    "       Where a.����id=K.����id  And a.��¼���� = 3 And a.����id=[2] And Nvl(a.��Ԥ��, 0) <> 0" & vbNewLine & _
    "       Group By Decode(a.��¼����, 3, a.���㷽ʽ, Null)) A, ���㷽ʽ B" & vbNewLine & _
    "Where a.���㷽ʽ = b.���� And B.���� = 8"
    'ȫ�˼�¼(�������ֵ������˻�)
    strSql = strSql & " Union " & _
    "Select a.���㷽ʽ, Nvl(b.����, 1) As ����, b.Ӧ����, a.���" & vbNewLine & _
    "From (Select Decode(a.��¼����, 3, a.���㷽ʽ, Null) As ���㷽ʽ, Sum(a.��Ԥ��) As ���" & vbNewLine & _
    "       From ����Ԥ����¼ A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.����id" & vbNewLine & _
    "                        From ������ü�¼ C, ������ü�¼ D, (Select Distinct ����ID From ����Ԥ����¼ I,Table(f_Str2list([1])) J Where I.�������=J.Column_Value) E" & vbNewLine & _
    "                        Where c.����id = e.����id And c.No = d.No And Mod(d.��¼����, 10) = 1) K" & _
    "       Where a.����id=K.����id  And a.��¼���� = 3 And a.����id=[2] And Nvl(a.��Ԥ��, 0) <> 0" & vbNewLine & _
    "         And Exists (Select 1 From ҽ�ƿ���� Where ID=A.�����ID And �Ƿ�����=0)" & _
    "       Group By Decode(a.��¼����, 3, a.���㷽ʽ, Null)) A, ���㷽ʽ B" & vbNewLine & _
    "Where a.���㷽ʽ = b.���� And B.���� = 7"
    'ҽ������
    strSql = strSql & " Union " & _
    "Select a.���㷽ʽ, Nvl(b.����, 1) As ����, b.Ӧ����, a.���" & vbNewLine & _
    "From (Select Decode(a.��¼����, 3, a.���㷽ʽ, Null) As ���㷽ʽ, Sum(a.��Ԥ��) As ���" & vbNewLine & _
    "       From ����Ԥ����¼ A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.����id" & vbNewLine & _
    "                        From ������ü�¼ C, ������ü�¼ D, (Select Distinct ����ID From ����Ԥ����¼ I,Table(f_Str2list([1])) J Where I.�������=J.Column_Value) E" & vbNewLine & _
    "                        Where c.����id = e.����id And c.No = d.No And Mod(d.��¼����, 10) = 1) K" & _
    "       Where a.����id =K.����id  And a.��¼���� In (1, 11, 3) And a.����id=[2] And" & vbNewLine & _
    "             Nvl(a.��Ԥ��, 0) <> 0" & vbNewLine & _
    "" & vbNewLine & _
    "       Group By Decode(a.��¼����, 3, a.���㷽ʽ, Null)) A, ���㷽ʽ B" & vbNewLine & _
    "Where a.���㷽ʽ = b.���� And B.���� In (3,4)"
    '����
    strSql = strSql & " Union " & _
    "Select a.���㷽ʽ, Nvl(b.����, 1) As ����, b.Ӧ����, a.���" & vbNewLine & _
    "From (Select Decode(a.��¼����, 3, a.���㷽ʽ, Null) As ���㷽ʽ, Sum(a.��Ԥ��) As ���" & vbNewLine & _
    "       From ����Ԥ����¼ A,(Select /*+ rule */" & vbNewLine & _
    "                        Distinct d.����id" & vbNewLine & _
    "                        From ������ü�¼ C, ������ü�¼ D, (Select Distinct ����ID From ����Ԥ����¼ I,Table(f_Str2list([1])) J Where I.�������=J.Column_Value) E" & vbNewLine & _
    "                        Where c.����id = e.����id And c.No = d.No And Mod(d.��¼����, 10) = 1) K" & _
    "       Where a.����id =K.����id  And a.��¼���� = 3 And a.����id=[2] And" & vbNewLine & _
    "             Nvl(a.��Ԥ��, 0) <> 0" & vbNewLine & _
    "" & vbNewLine & _
    "       Group By Decode(a.��¼����, 3, a.���㷽ʽ, Null)) A, ���㷽ʽ B" & vbNewLine & _
    "Where a.���㷽ʽ = b.���� And B.���� = 9"
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Replace(strBalance, "'", ""), mlng����ID)
    Do While Not mrsBalance.EOF
        dblSum = dblSum - Val(NVL(mrsBalance!���))
        mrsBalance.MoveNext
    Loop
    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    
    InitPatialBalance = Format(dblSum, "0.00")
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetFpToBIllNOs(ByVal strFpNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ķ�Ʊ��,�ҳ���Ӧ�ĵ��ݺ�
    '����:���ض�Ӧ�ĵ��ݺ�,�ö��ŷָ�
    '����:���˺�
    '����:2011-02-25 10:50:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, strNos As String
    
    On Error GoTo errHandle
    
    strSql = "" & _
    "   Select distinct NO From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B " & _
    "   Where A.��������=1 and A.ID=B.��ӡID and B.Ʊ��=1 And B.����=[1]  " & _
    "   Order by NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strFpNo)
    strNos = ""
    With rsTemp
        Do While Not .EOF
            strNos = strNos & "," & NVL(rsTemp!NO)
            .MoveNext
        Loop
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    zlGetFpToBIllNOs = strNos
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub CalcSUMMony()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:״̬����Ϣ����
    '����:������
    '����:2014-6-20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cur��� As Currency
    With vsFee
        cur��� = 0
        For i = .FixedRows To .Rows - 1
            If vsFee.TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                cur��� = cur��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
            End If
        Next
        lblSum.Caption = "��ǰת���ϼ�:" & Format(cur���, "###0.00;-###0.00;0.00;0.00")
    End With
End Sub

Private Sub vsFee_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsFee
        Select Case Col
        Case .ColIndex("ѡ��")
            Call SetBlanceShow
            Call CalcSUMMony
        Case Else
        End Select
    End With
End Sub

Private Sub vsFee_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strBalance As String, bytType As Byte, blnOld As Boolean
    
    If NewRow = OldRow Or NewRow < 1 Then Exit Sub
    With vsFee
        If mintType = 1 Then
            strBalance = Trim(.TextMatrix(NewRow, .ColIndex("�������")))
            If Val(strBalance) > 0 Then
                strBalance = Trim(.TextMatrix(NewRow, .ColIndex("����ID")))
                blnOld = True
            End If
        Else
            strBalance = Trim(.TextMatrix(NewRow, .ColIndex("���ŵ���")))
        End If
        If NewRow = 0 Or strBalance = "" Then
            mfrmFeeDetail.zlRefresh mintType, 0
        Else
            mfrmFeeDetail.zlRefresh mintType, strBalance, blnOld
        End If
        .ForeColorSel = vsFee.CellForeColor
    End With
    LoadInvoice mintType, NewRow
    LoadBalance mintType, NewRow
End Sub

Private Sub LoadBalance(ByVal intTYPE As Integer, ByVal NewRow As Integer)
'-----------------------------------------------------------------------------------------------------------------------
'����:��ȡ��ǰѡ���¼�Ľ�����Ϣ
'����:������
'����:2014-6-20
'��ע:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset, strBalance As String
    vsBalanceStyle.Clear 1
    vsBalanceStyle.Rows = 2
    If intTYPE = 1 Then
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("�������")))
        strSql = "" & _
            "Select ���㷽ʽ, Sum(��Ԥ��) As ������" & vbNewLine & _
            "From (Select a.���㷽ʽ, a.��Ԥ��" & vbNewLine & _
            "       From ����Ԥ����¼ A," & vbNewLine & _
            "            (Select Distinct ����id" & vbNewLine & _
            "              From ������ü�¼" & vbNewLine & _
            "              Where Mod(��¼����, 10) = 1 And ��¼״̬ <> 0 And" & vbNewLine & _
            "                    NO In (Select Distinct NO" & vbNewLine & _
            "                           From ������ü�¼ C, (Select Distinct ����id From ����Ԥ����¼ Where ������� = [1]) D" & vbNewLine & _
            "                           Where Mod(c.��¼����, 10) = 1 And c.��¼״̬ <> 0 And c.����id = d.����id)) B" & vbNewLine & _
            "       Where a.��¼���� = 3 And a.����id = b.����id)" & vbNewLine & _
            "Group By ���㷽ʽ" & vbNewLine & _
            "Union" & vbNewLine & _
            "Select ���㷽ʽ, Sum(��Ԥ��) As ������" & vbNewLine & _
            "From (Select 'Ԥ����' As ���㷽ʽ, a.��Ԥ��" & vbNewLine & _
            "       From ����Ԥ����¼ A," & vbNewLine & _
            "            (Select Distinct ����id" & vbNewLine & _
            "              From ������ü�¼" & vbNewLine & _
            "              Where Mod(��¼����, 10) = 1 And ��¼״̬ <> 0 And" & vbNewLine & _
            "                    NO In (Select Distinct NO" & vbNewLine & _
            "                           From ������ü�¼ C, (Select Distinct ����id From ����Ԥ����¼ Where ������� = [1]) D" & vbNewLine & _
            "                           Where Mod(c.��¼����, 10) = 1 And c.��¼״̬ <> 0 And c.����id = d.����id)) B" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And a.����id = b.����id)" & vbNewLine & _
            "Group By ���㷽ʽ"

        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(strBalance))
        With vsBalanceStyle
            Do While Not rsTmp.EOF
                If Val(NVL(rsTmp!������)) <> 0 Then
                    .TextMatrix(.Rows - 1, 0) = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("���ŵ���")))
                    .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!���㷽ʽ)
                    .TextMatrix(.Rows - 1, 2) = Format(NVL(rsTmp!������), "0.00")
                    .Rows = .Rows + 1
                End If
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    Else
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("���ŵ���")))
        strSql = "Select a.���㷽ʽ, Sum(a.��Ԥ��) As ������" & vbNewLine & _
                "From ����Ԥ����¼ A, ������ü�¼ B" & vbNewLine & _
                "Where Mod(a.��¼����, 10) = 2 And a.����id = b.����id And Mod(b.��¼����, 10) = 2 And b.��¼״̬ <> 0 And b.No = [1]" & vbNewLine & _
                "Group By ���㷽ʽ" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select 'Ԥ����' As ���㷽ʽ, Sum(a.��Ԥ��) As ������" & vbNewLine & _
                "From ����Ԥ����¼ A, ������ü�¼ B" & vbNewLine & _
                "Where Mod(a.��¼����, 10) = 2 And a.����id = b.����id And Mod(b.��¼����, 10) = 2 And b.��¼״̬ <> 0 And b.No = [1]"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalance)
        With vsBalanceStyle
            Do While Not rsTmp.EOF
                If Val(NVL(rsTmp!������)) <> 0 Then
                    .TextMatrix(.Rows - 1, 0) = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("���ŵ���")))
                    .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!���㷽ʽ)
                    .TextMatrix(.Rows - 1, 2) = Format(NVL(rsTmp!������), "0.00")
                    .Rows = .Rows + 1
                End If
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    End If
End Sub

Private Sub LoadInvoice(ByVal bytType As Byte, ByVal NewRow As Long)
'-----------------------------------------------------------------------------------------------------------------------
'����:��ȡ��ǰѡ���¼��Ʊ����Ϣ
'����:������
'����:2014-6-20
'��ע:
'-----------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset, strBalance As String
    vsfInvoice.Clear 1
    vsfInvoice.Rows = 2
    If bytType = 1 Then
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("�������")))
        strSql = "Select Distinct D.����" & vbNewLine & _
                " From Ʊ�ݴ�ӡ���� C,Ʊ��ʹ����ϸ D," & _
                " (Select Distinct A.NO From ������ü�¼ A,����Ԥ����¼ B Where A.����ID=B.����ID And Mod(A.��¼����, 10) = 1 And B.�������= [1]) E" & vbNewLine & _
                " Where E.No=C.No(+) And C.��������(+)=1 And C.ID = D.��ӡID(+) " & _
                " And Not Exists (Select 1 From Ʊ��ʹ����ϸ Where Ʊ�� = d.Ʊ�� And ���� = d.���� And ���� = 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(strBalance))
        With vsfInvoice
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, 0) = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("���ŵ���")))
                .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!����)
                .Rows = .Rows + 1
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    Else
        strBalance = Trim(vsFee.TextMatrix(NewRow, vsFee.ColIndex("���ŵ���")))
        strSql = " Select Distinct E.����,C.NO" & vbNewLine & _
                 " From ������ü�¼ A, ������ü�¼ B,���˽��ʼ�¼ C,Ʊ�ݴ�ӡ���� D,Ʊ��ʹ����ϸ E" & vbNewLine & _
                 " Where Mod(a.��¼����, 10) = 2 And a.No = [1] And b.����id = a.����ID And C.ID=B.����ID And C.No=D.No(+) And D.��������(+)=3 And D.ID=E.��ӡID(+) " & _
                 " And Not Exists (Select 1 From Ʊ��ʹ����ϸ Where Ʊ�� = e.Ʊ�� And ���� = e.���� And ���� = 2)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strBalance)
        With vsfInvoice
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, 0) = NVL(rsTmp!NO)
                .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!����)
                .Rows = .Rows + 1
                rsTmp.MoveNext
            Loop
            .Rows = .Rows - 1
            If .Rows = 1 Then .Rows = 2
        End With
    End If
End Sub

Private Function LoadStyle() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    cboStyle.Clear
    On Error GoTo errH
    Set rsTmp = Get���㷽ʽ("�շ�", "1,2")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,", "," & rsTmp!���� & ",") > 0 And Val(NVL(rsTmp!Ӧ����)) = 0 Then
            cboStyle.AddItem rsTmp!����
            cboStyle.ItemData(cboStyle.NewIndex) = rsTmp!����
            If rsTmp!ȱʡ = 1 And cboStyle.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboStyle.hWnd, cboStyle.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboStyle.ListIndex = -1 And cboStyle.ListCount > 0 Then Call zlControl.CboSetIndex(cboStyle.hWnd, 0)
    txtSum.ForeColor = vbRed
    strSql = "" & _
            " Select B.����,B.����,Nvl(B.ȱʡ��־,0) as ȱʡ,Nvl(B.����,1) as ����,Nvl(B.Ӧ����,0) as Ӧ����" & _
            " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
            " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ " & _
            " And B.����<>8 " & _
            " Order by ����,lpad(����,3,' ')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "�շ�")
    For i = 1 To rsTmp.RecordCount
        If InStr(",1,2,7,", "," & rsTmp!���� & ",") > 0 Then
            mstrStyle = mstrStyle & rsTmp!���� & ":"
        End If
        rsTmp.MoveNext
    Next
    LoadStyle = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SelAllNO()
    Dim i As Long
    With vsFee
        If .Rows = 2 And .TextMatrix(1, .ColIndex("�������")) = "" Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ŵ���")) <> "" Then
                .TextMatrix(i, .ColIndex("ѡ��")) = "��"
            End If
        Next
        Call CheckInsure
        Call SetBlanceShow
        Call CalcSUMMony
        mblnSel = True
    End With
End Sub

Private Sub CheckInsure()
    Dim i As Integer, intInsure As Integer, blnSelect As Boolean
    With vsFee
        For i = 1 To .Rows - 1
            intInsure = Val(.TextMatrix(i, .ColIndex("����")))
            blnSelect = .TextMatrix(i, .ColIndex("ѡ��")) <> ""
            If intInsure > 0 And blnSelect Then
                If gclsInsure.GetCapability(support�����������, mlng����ID, intInsure) = False Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = ""
                End If
            End If
        Next i
    End With
End Sub

Private Sub picBalanceStyle_Resize()
    On Error Resume Next
    With vsBalanceStyle
        .Top = 0
        .Left = 0
        .Width = picBalanceStyle.Width
        .Height = picBalanceStyle.Height - lblSum.Height - 60
    End With
    With lblSum
        .Top = vsBalanceStyle.Height
        .Left = 15
    End With
End Sub

Private Sub picBalance_Resize()
    On Error Resume Next
    With vsBalance
        .Top = 0
        .Left = 0
        .Width = picBalance.Width
        .Height = picBalance.Height - picBack.Height - 30
    End With
    
    With picBack
        .Left = picBalance.Width - 3500
        .Top = vsBalance.Top + vsBalance.Height
    End With
End Sub

Private Sub picFee_Resize()
    With vsFee
        .Top = 0
        .Left = 0
        .Width = picFee.Width
        .Height = picFee.Height
    End With
End Sub

Private Sub picInvoice_Resize()
    With vsfInvoice
        .Top = 0
        .Left = 0
        .Width = picInvoice.Width
        .Height = picInvoice.Height
    End With
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    With vsFee
        If .DataSource Is Nothing Then
            strHead = "ѡ��,4,500|���,4,850|����,4,800|ҽ��,4,500|���ŵ���,4,850|���ŷ�Ʊ,4,1100|������,4,800|Ӧ�ս��,7,850|ʵ�ս��,7,850|����ʱ��,4,1850|�������,4,0|����,4,0"
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
            Next
            .Rows = 2
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        'ѡ��,4,500|���,4,850|ҽ��,4,500|���ݺ�,4,850|Ʊ�ݺ�,4,1100|������,4,800|Ӧ�ս��,7,850|ʵ�ս��,7,850|����ʱ��,4,1850|����ID,4,0|����,4,0
        For i = 0 To .Cols - 1
             .FixedAlignment(i) = flexAlignCenterCenter
             .ColAlignment(i) = flexAlignLeftCenter
             .ColKey(i) = Trim(.TextMatrix(0, i))
             Select Case .ColKey(i)
             Case "ѡ��", "���", "����", "ҽ��", "���ŵ���", "���ŷ�Ʊ"
                .ColAlignment(i) = flexAlignCenterCenter
             Case "Ӧ�ս��", "ʵ�ս��"
                .ColAlignment(i) = flexAlignRightCenter
             End Select
             If .ColKey(i) Like "*ID" Or .ColKey(i) = "����" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
             End If
        Next
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, "����תסԺ�б�", True
        .RowHeight(0) = 320
        .Row = 1
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ�е�ѡ��״̬
    '       ����Ƕ��ŵ����е�һ��,����ͬʱ���ö����е���������
    '����:���˺�
    '����:2011-02-21 16:10:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNO As String, i As Long, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant
    Dim str���� As String
    
    With vsFee
        If .TextMatrix(lngRow, .ColIndex("ѡ��")) <> IIf(blnSelect, "��", "") Then
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("���")))
            If intInsure > 0 And blnSelect And str���� = "�շ�" Then
                strNO = .TextMatrix(lngRow, .ColIndex("���ŵ���"))
                If Not gclsInsure.GetCapability(support�����������, mlng����ID, intInsure) Then
                    frmFeeRefundment.stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧�������������,���в�����ѡ��ת��!"
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                    Exit Function
                Else
                    '���жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                    'strTmp = GetBalanceType(strNO)
                    If strTmp <> "" Then
                        arrBalanceType = Split(strTmp, ",")
                        For i = 0 To UBound(arrBalanceType)
                            strBalanceType = arrBalanceType(i)
                            If Not gclsInsure.GetCapability(support�����������, mlng����ID, intInsure, strBalanceType) Then
                                frmFeeRefundment.stbThis.Panels(2).Text = "����[" & strNO & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                                Exit Function
                            End If
                        Next
                    End If
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
        End If
    End With
    SetRowSelected = True
End Function

Public Sub ClsAllNO()
   Dim i As Long
    With vsFee
        If .Rows = 2 And .TextMatrix(1, .ColIndex("�������")) = "" Then Exit Sub
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ŵ���")) <> "" Then
                .TextMatrix(i, .ColIndex("ѡ��")) = ""
            End If
        Next
        Call SetBlanceShow
        mblnSel = False
        Call CalcSUMMony
    End With
End Sub

Private Sub InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:������
    '����:2014-06-19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panTop As Pane, panBottom As Pane, panRight As Pane
    If mfrmFeeDetail Is Nothing Then Set mfrmFeeDetail = New frmFeeDetail
    Call mfrmFeeDetail.ShowMe(lblBack.Font, mlngModule, mstrPrivs, 1, 0)
    Load mfrmFeeDetail
    
    Set panThis = dkpMain.CreatePane(mObjPancel.Pan_Bill, 250, 580, DockTopOf, Nothing)
    panThis.Title = "����תסԺ�б�"
    panThis.Tag = mObjPancel.Pan_Bill
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picFee.hWnd
    
    Set panRight = dkpMain.CreatePane(mObjPancel.Pan_Invoice, 1500 / Screen.TwipsPerPixelX, 300, DockRightOf, panThis)
    panRight.Title = "��Ʊ��Ϣ"
    panRight.Tag = mObjPancel.Pan_Invoice
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picInvoice.hWnd
    
    Set panRight = dkpMain.CreatePane(mObjPancel.Pan_BalanceInfo, 1500 / Screen.TwipsPerPixelX, 580, DockBottomOf, panRight)
    panRight.Title = "�տ����"
    panRight.Tag = mObjPancel.Pan_Balance
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picBalanceStyle.hWnd
    
    Set panThis = dkpMain.CreatePane(mObjPancel.Pan_List, 250, 580, DockBottomOf, panThis)
    panThis.Title = "������ϸ�б�"
    panThis.Tag = mObjPancel.Pan_List
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = mfrmFeeDetail.hWnd
    
    
    Set panThis = dkpMain.CreatePane(mObjPancel.Pan_Balance, 250, 580, DockBottomOf, Nothing)
    panThis.Title = "������Ϣ"
    panThis.Tag = mObjPancel.Pan_BalanceInfo
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Handle = picBalance.hWnd
    panThis.MaxTrackSize.Height = 75
    panThis.MinTrackSize.Height = 75
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    Set dkpMain.PaintManager.CaptionFont = vsFee.Font
    
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub

Public Function ReadListData(ByVal strFindNo As String, ByVal strFindFpNo As String, _
                            rsInfo As ADODB.Recordset, Optional blnFilter As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ���ʵ���ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�Flase
    '����:���˺�
    '����:2011-01-25 17:10:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim lngCol As Long, strSql As String, lngRow As Long
    Dim strFilter As String, strNos As String
    Dim strWhere As String, strTable1 As String
    Dim strALLNOs As String
    mstrFindNO = strFindNo
    mstrFindFpNo = strFindFpNo
    Set mrsInfo = rsInfo
    If mrsInfo Is Nothing Then
        lng����ID = 0
    ElseIf mrsInfo.State <> 1 Then
        lng����ID = 0
    Else
        lng����ID = Val(NVL(mrsInfo!����ID))
    End If
    
    If mstrFindNO <> "" Then
        If mintType = 1 Then
            strNos = Replace(GetMultiNOs(mstrFindNO), "'", "")
            strSql = "Select ����ID From ������ü�¼ Where MOD(��¼����,10)=1 And NO=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNos)
            lng����ID = Val(NVL(rsTemp!����ID))
            strWhere = "  And A.����ID=[1]"
        Else
            strNos = mstrFindNO
            strTable1 = ",Table( f_Str2list([2])) J "
            strWhere = "  And A.NO=J.Column_Value"
        End If
    ElseIf mstrFindFpNo <> "" And mintType = 1 Then
        strNos = zlGetFpToBIllNOs(mstrFindFpNo)
        If strNos = "" Then
            MsgBox "δ�ҵ���Ӧ��Ʊ�ŵĵ���,����!"
            Exit Function
        End If
        strSql = "Select ����ID From ������ü�¼ Where MOD(��¼����,10)=1 And NO=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNos)
        lng����ID = Val(NVL(rsTemp!����ID))
        strWhere = "  And A.����ID=[1]"
    Else
        strTable1 = ""
        strWhere = "  And A.����ID=[1]"
    End If
    mblnSel = False
    On Error GoTo errHandle
    If blnFilter = False Then zlCommFun.ShowFlash "���ڶ�ȡ��������,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    If mintType = 1 Then
        strSql = "" & _
            "Select /*+ rule */" & vbNewLine & _
            " '��' As ѡ��, '�շ�' As ���, Decode(Max(b.����), Null, '', '��') As ҽ��, Min(a.No) As ���ŵ���, Min(a.ʵ��Ʊ��) As ���ŷ�Ʊ, a.�ѱ�," & vbNewLine & _
            " LTrim(To_Char(Sum(a.Ӧ�ս��), '9999999990.00')) As Ӧ�ս��, LTrim(To_Char(Sum(a.ʵ�ս��), '9999999990.00')) As ʵ�ս��," & vbNewLine & _
            " a.����Ա���� As ����Ա, Null As ������, Max(b.����) As ����, Nvl(e.�������, e.����id) As �������, Min(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, Null as ����id " & vbNewLine & _
            "From ������ü�¼ A, ���ս����¼ B," & vbNewLine & _
            "     (Select Distinct Nvl(d.����id, c.����id) As ����id, d.������� As �������" & vbNewLine & _
            "       From ����Ԥ����¼ C, ����Ԥ����¼ D" & vbNewLine & _
            "       Where d.����id = [1] And c.����id = [1] And c.������� = d.�������(+) And d.�������(+) < 0 ) E" & vbNewLine & _
            "Where a.����id = b.��¼id(+) And b.����(+) = 1 And Nvl(b.���(+),1)=1 And Mod(a.��¼����, 10) = 1 And a.����id = [1] And Exists" & vbNewLine & _
            " (Select 1 From ������ü�¼ Where ����id = e.����id And ��¼״̬ In (1, 3)) And" & vbNewLine & _
            "      a.No In (Select Distinct x.No" & vbNewLine & _
            "               From ������ü�¼ X, ������ü�¼ Y" & vbNewLine & _
            "               Where y.����id = e.����id And x.No = y.No And Mod(x.��¼����, 10) = 1" & vbNewLine & _
            "               Group By x.No, x.���" & vbNewLine & _
            "               Having Sum(Nvl(x.����, 1) * x.����) <> 0) And Not Exists" & vbNewLine & _
            " (Select 1 From ������ü�¼ Where NO = a.No And Mod(��¼����, 10) = 1 And Nvl(����״̬, 0) = 1) And Exists" & vbNewLine & _
            " (Select 1 From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ In (1, 3) And ����id = e.����id) And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From ������˼�¼ E, ������ü�¼ F" & vbNewLine & _
            "       Where e.��¼״̬ = 1 And f.Id = e.����id And f.No = a.No And Mod(f.��¼����, 10) = 1)" & vbNewLine & _
            "Group By a.�ѱ�, Nvl(e.�������, e.����id), a.����Ա����"

        strSql = strSql & " Union " & _
            " Select  '��' As ѡ��, '�շ�' As ���,Decode(a.����, Null, '', '��') As ҽ��,a.No as ���ŵ���,a.ʵ��Ʊ�� as ���ŷ�Ʊ,a.�ѱ�, " & _
            "   LTrim(To_Char(a.Ӧ�ս��, '9999999990.00')) As Ӧ�ս��, LTrim(To_Char(a.ʵ�ս��, '9999999990.00')) As ʵ�ս��, " & _
            "   a.����Ա���� As ����Ա, Null As ������, a.���� As ����, Nvl(c.�������,c.����id) As �������,a.�Ǽ�ʱ�� As �Ǽ�ʱ��, c.����id " & vbNewLine & _
            " From (Select Max(����) as ����, Decode(Max(����), 0, '', '��') As ҽ��, Min(Decode(�۸񸸺�, Null, ID, 0)) As ID, " & vbNewLine & _
            "           NO, ʵ��Ʊ��, Avg(Nvl(����, 1)) As ����, Sum(����) ����, Sum(Ӧ�ս��) As Ӧ�ս��, " & vbNewLine & _
            "           Sum(ʵ�ս��) As ʵ�ս��, ������, Min(�Ǽ�ʱ��) As �Ǽ�ʱ��, " & vbNewLine & _
            "           Min(����Ա����) As ����Ա����, �ѱ� " & vbNewLine & _
            "       From (Select Row_Number() Over(Partition By a.ID Order By m.���) As Rn,a.Id,Nvl(M.����,0) as ����, " & _
            "               A.�۸񸸺�, A.NO, A.ʵ��Ʊ��, A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��, A.������, A.�Ǽ�ʱ��, " & vbNewLine & _
            "               a.����Ա����, a.�ѱ� " & vbNewLine & _
            "             From ������ü�¼ A, ���ս����¼ M, ������˼�¼ Q " & vbNewLine & _
            "             Where A.��¼���� = 1 And A.����ID= [1] " & _
            "                   And A.��¼״̬ <> 0 And A.����id = M.��¼id(+) " & vbNewLine & _
            "                   And  M.����(+) = 1 And A.ID = Q.����id(+) And Nvl(a.���ӱ�־,0) <> 9 " & vbNewLine & _
            "                   And a.Id In (Select b.Id " & vbNewLine & _
            "                        From ������ü�¼ B, ������ü�¼ C, ������˼�¼ D" & vbNewLine & _
            "                        Where c.Id = d.����id And d.��¼״̬ = 1 And b.No = c.No))" & vbNewLine & _
            "       Where Rn < 2" & _
            "    Group By NO, ʵ��Ʊ��, ������, �ѱ� " & _
            "    Having Sum(����) <> 0) A, ������ü�¼ B, ����Ԥ����¼ C " & _
            " Where a.Id = b.Id And b.����ID=c.����ID And Nvl(C.�������,1) > 0"
    Else
        '���ʵ�
        strSql = "" & _
            " Select /*+ rule */   '��' as ѡ��,'����' as ���,Decode(NULL,Null,'','��') as ҽ��, A.NO As ���ŵ���, A.ʵ��Ʊ�� As ���ŷ�Ʊ, 0 As �������,A.�ѱ�," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��,A.����Ա���� As ����Ա,A.������," & vbNewLine & _
            "       Max(A.�Ǽ�ʱ��) As �Ǽ�ʱ��,0 AS ���� " & vbNewLine & _
            " From ������ü�¼ A" & vbNewLine & _
            " Where A.��¼���� =2 And A.��¼״̬ <> 0 " & strWhere & vbNewLine & _
            "           And Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��¼����=A.��¼���� And K.���ӱ�־ <> 9 Group By K.��� Having Sum(K.����) <> 0) " & vbNewLine & _
            "      And Exists (Select 1 From ������˼�¼ E,������ü�¼ F Where E.��¼״̬=1 And F.ID=E.����ID And F.NO=A.NO And MOD(F.��¼����,10)=2)  " & _
            "Group By A.NO, A.ʵ��Ʊ��, A.������, A.����Ա����,A.�ѱ� " & vbNewLine
            
    End If
    
    strSql = strSql & " Order By ���,���, ���ŷ�Ʊ Desc, ���ŵ��� Desc"
    If mrsFeeList Is Nothing Or blnFilter = False Then
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, strNos)
    Else
        mrsFeeList.Filter = 0
    End If
    mlng����ID = lng����ID
    vsFee.Redraw = flexRDNone
    vsFee.Clear: vsFee.Cols = 0
    Set vsFee.DataSource = mrsFeeList
    If vsFee.Rows <= 1 Then vsFee.Rows = 2
    With vsFee
        For lngCol = 0 To .Cols - 1
             .ColAlignment(lngCol) = flexAlignLeftCenter
             .FixedAlignment(lngCol) = flexAlignCenterCenter
              .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
              If .ColKey(lngCol) Like "*ID" Or InStr(1, ",����,����,���,��������,ת����־,�շ����,�������,����ID,", "," & .ColKey(lngCol) & ",") > 0 Then
                    .ColHidden(lngCol) = True
              ElseIf .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��*" Or .ColKey(lngCol) Like "*��" Then
                    .ColAlignment(lngCol) = flexAlignRightCenter
              ElseIf .ColKey(lngCol) Like "ѡ��*" Then
                    .ColAlignment(lngCol) = flexAlignCenterCenter
              End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsFee, Me.Caption, IIf(mintType = 1, "�˷��б�", "�����б�"), True
        '����
        Dim strNO As String, str���� As String
        strALLNOs = ""
        For lngRow = 1 To .Rows - 1
            If strNO <> Trim(.TextMatrix(lngRow, .ColIndex("�������"))) _
                 And strNO <> "" Then
                '�����ָ���
                .Select lngRow, .FixedCols, lngRow, .Cols - 1
                .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
            End If
            .Cell(flexcpData, lngRow, .ColIndex("�������")) = .TextMatrix(lngRow, .ColIndex("�������"))
            strNO = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("���")))
            strALLNOs = strALLNOs & "," & strNO
        Next
        .Editable = flexEDNone
    End With
    
    If strALLNOs <> "" Then strALLNOs = Mid(strALLNOs, 2)
    If blnFilter = False Then zlCommFun.StopFlash
    vsFee.Redraw = flexRDBuffered
    '���ؽ��㷽ʽ
    Call CheckInsure
    Call InitBlanceData(strALLNOs)
    Call SetBlanceShow
    Call CalcSUMMony
    Call frmFeeRefundment.StatusShowBillSum
    Call vsFee_AfterRowColChange(0, 0, 1, 0)
    Call picBalance_Resize
    Screen.MousePointer = 0
    ReadListData = True
    Exit Function
errHandle:
    vsFee.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
   If blnFilter = False Then zlCommFun.StopFlash
End Function

Public Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʻ��˷�
    '����:�˷ѻ����ʳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-23 11:21:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, lng���� As Long, lng����ID As Long
    Dim strOutNos As String, strTemp As String, strDelDate As String
    Dim m As Long, i As Long, blnHaveData As Boolean, blnPrintList As Boolean '�Ƿ��ӡ�嵥
    Dim cllDelNO As Collection, strDelNOs As String, lngRow As Long, strNO As String
    Dim lng����ID As Long, rsTmp As ADODB.Recordset, blnOld As Boolean, strFirstNo As String
    
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    blnPrintList = False
    If InStr(mstrPrivs, ";��ӡ�嵥;") > 0 And mintType = 1 Then
        Select Case mint�շ��嵥    '0-����ӡ,1-Ҫ��ӡ,2-ѡ���Ƿ��ӡ
        Case 2
             If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintList = True
             End If
        Case 1
            blnPrintList = True
        End Select
    End If
    
    With vsFee
        If .Rows <= 1 Then Exit Function
        If .Cols <= 1 Then Exit Function
        Set cllDelNO = New Collection
        strTemp = ""
        For lngRow = 1 To .Rows - 1
            '���ʵ���
            If mintType = 1 Then
                strNO = Trim(.TextMatrix(lngRow, .ColIndex("�������")))
                If CheckBillExistReplenishData(0, Val(strNO)) Then
                    MsgBox "ѡ����˷ѵ��ݴ��ڲ�������¼���޷������˷ѣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                strNO = Trim(.TextMatrix(lngRow, .ColIndex("���ŵ���")))
            End If
            If .TextMatrix(lngRow, .ColIndex("ѡ��")) <> "" _
                And strNO <> "" And InStr(1, "," & strTemp & ",", "," & strNO & ",") = 0 Then
                lng���� = Val(.TextMatrix(lngRow, .ColIndex("����")))
                If mintType = 1 Then
                    If Val(strNO) > 0 Then
                        blnOld = True
                        strOutNos = strNO
                        strFirstNo = NVL(.TextMatrix(lngRow, .ColIndex("���ŵ���")))
                        lng����ID = Val(.TextMatrix(lngRow, .ColIndex("����ID")))
                        cllDelNO.Add Array(strFirstNo, strFirstNo, lng����, lng����ID, True, strFirstNo)
                        strTemp = strTemp & "," & strNO & "," & strOutNos
                    Else
                        lng����ID = Val(.TextMatrix(lngRow, .ColIndex("�������")))
                        strFirstNo = NVL(.TextMatrix(lngRow, .ColIndex("���ŵ���")))
                        strOutNos = strNO
                        
                        If strOutNos <> "" Then
                            '�����ŵ��Ƿ����
                            blnHaveData = False
                            For i = 1 To cllDelNO.Count
                                If cllDelNO(i)(0) = strNO Then
                                    blnHaveData = True: Exit For
                                End If
                                If InStr(1, "," & cllDelNO(i)(1) & ",", "," & strNO & ",") > 0 Then
                                    blnHaveData = True: Exit For
                                End If
                            Next
                            If blnHaveData = False Then
                                '�������ʵ���
                                cllDelNO.Add Array(strNO, strOutNos, lng����, lng����ID, False, strFirstNo)
                            End If
                            strTemp = strTemp & "," & strNO & "," & strOutNos
                        End If
                    End If
                Else
                    lng����ID = Val(.TextMatrix(lngRow, .ColIndex("�������")))
                    strOutNos = strNO
                    
                    If strOutNos <> "" Then
                        '�����ŵ��Ƿ����
                        blnHaveData = False
                        For i = 1 To cllDelNO.Count
                            If cllDelNO(i)(0) = strNO Then
                                blnHaveData = True: Exit For
                            End If
                            If InStr(1, "," & cllDelNO(i)(1) & ",", "," & strNO & ",") > 0 Then
                                blnHaveData = True: Exit For
                            End If
                        Next
                        If blnHaveData = False Then
                            '�������ʵ���
                            cllDelNO.Add Array(strNO, strOutNos, lng����, lng����ID)
                        End If
                        strTemp = strTemp & "," & strNO & "," & strOutNos



                    End If

                End If
            End If
        Next
    End With
    'ִ�о������ʻ��˷Ѳ���
    If cllDelNO.Count = 0 Then
        MsgBox "ע��:" & vbCrLf & "    û��ѡ��һ����Ҫ�����˷ѻ����ʵĵ���,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '�˷�
    strDelNOs = ""
    If mintType = 2 Then
        If ExecuteWirteOff(strDelDate, cllDelNO) = False Then Exit Function
    Else
        For i = 1 To cllDelNO.Count
            If ExecuteDelBill(strDelDate, IIf(cllDelNO(i)(1) <> "", cllDelNO(i)(1), cllDelNO(i)(0)), Val(cllDelNO(i)(2)), Val(cllDelNO(i)(2)), cllDelNO(i)(4), cllDelNO(i)(5)) = False Then
                    Exit Function
            End If
            strDelNOs = strDelNOs & "," & cllDelNO(i)(5)
        Next
    End If
    If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
    '��ӡ�����嵥
    If blnPrintList And mintType = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & "'" & Replace(strDelNOs, ",", "','") & "'", "ҩƷ��λ=" & IIf(mblnҩ����λ, 1, 0), 2)
    End If
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteWirteOff(strDELDae As String, ByVal cllDel As Variant) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�������������
    '����:���˺�
    '����:2011-02-25 10:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strSql As String
    Dim cllPro As Collection
    Set cllPro = New Collection
    For i = 1 To cllDel.Count
        'Zl_����תסԺ_����ת��
        strSql = "Zl_����תסԺ_����ת��("
        '  No_In         סԺ���ü�¼.NO%Type,
        strSql = strSql & "'" & cllDel(i)(0) & "',"
        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
        strSql = strSql & "'" & UserInfo.��� & "',"
        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
        strSql = strSql & "'" & UserInfo.���� & "',"
        '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type
        strSql = strSql & "To_Date('" & strDELDae & "','yyyy-mm-dd hh24:mi:ss'),"
        '   ��������_In   Number := 0
        '   --��������_In:0-����תסԺ��������;1-��������˷�ģʽ
        strSql = strSql & "1)"
        zlAddArray cllPro, strSql
    Next
    On Error GoTo errHandle
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ExecuteWirteOff = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceSet() As ADODB.Recordset
'���ܣ�����һ�������¼������
    Dim rsTmp As New ADODB.Recordset
       
    rsTmp.Fields.Append "�������", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "���㷽ʽ", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "������", adCurrency, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set GetBalanceSet = rsTmp
End Function

Private Function ExecuteDelBill(ByVal strDelDate As String, ByVal strNos As String, intInsure As Integer, _
                                ByVal lng����ID As Long, Optional ByVal blnOld As Boolean, Optional ByVal strFirstNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ������˷Ѳ���
    '���:strNos-���ݺ�:�����Ƕ൥��
    '       lngInsure-����
    '����:ִ�гɹ�������true,���򷵻�False
    '����:���˺�
    '����:2011-02-24 15:35:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, k As Long, varTemp  As Variant, strAllBalance      As String, strBalance As String
    Dim blnҽ���ӿڴ�ӡƱ�� As Boolean, bln�൥��һ�ν��� As Boolean, blnYB�������� As Boolean, bln�˷Ѻ��ӡ�ص� As Boolean
    Dim lng����ID As Long, cllPro As Collection, blnTrans As Boolean, lng����ID As Long, str������ˮ�� As String, str����˵�� As String
    Dim lng����ID1 As Long, varBalance As Variant, strAdvance As String, strInvoice As String
    Dim strSql As String, j As Long, blnTransMedicare As Boolean, rsTmp As ADODB.Recordset
    Dim str���㷽ʽ As String, cur������ As Currency, cur�ɷ���� As Currency, cur����� As Currency, cur��� As Currency, cur�˿�ϼ� As Currency
    Dim strDelNOs As String, lng����ID As Long, blnExecuteThreeSwap As Boolean
    
    If intInsure <> 0 Then
        blnҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, , intInsure, CStr(lng����ID))
        bln�൥��һ�ν��� = gclsInsure.GetCapability(support�൥��һ�ν���, , intInsure)
        blnYB�������� = gclsInsure.GetCapability(support�����������, , intInsure)
        If blnYB�������� = False Then
            MsgBox "ע��:" & vbCrLf & "   ���ݺ�Ϊ" & strNos & "�ĵ���,��֧��ҽ����������,����"
            Exit Function
        End If
        bln�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, , intInsure)
    End If
    
    If intInsure <> 0 And blnҽ���ӿڴ�ӡƱ�� Then
        Dim strUserType As String
        Dim lngShareUseID As Long
        If mrsInfo Is Nothing Then
            lng����ID = mlng����ID
        ElseIf mrsInfo.State <> 1 Then
            lng����ID = mlng����ID
        Else
            lng����ID = Val(NVL(mrsInfo!����ID))
        End If
        strUserType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
        lngShareUseID = zl_GetInvoiceShareID(1121, strUserType)
         
        lng����ID = GetInvoiceGroupID(1, 1, lng����ID, lngShareUseID)
        Select Case lng����ID
            Case -1
                MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Exit Function
            Case -2
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Exit Function
        End Select
        strInvoice = GetNextBill(lng����ID)
    End If
    
    '��ȡ����ID
    Err = 0: On Error GoTo errHandle
    Set cllPro = New Collection
    varTemp = Split(strNos, ",")
    For i = 0 To UBound(varTemp)
            'Zl_����תסԺ_�շ�ת��
            strSql = "Zl_����תסԺ_�շ�ת��("
            '     �������_In   ����Ԥ����¼.�������%Type,
            strSql = strSql & IIf(blnOld, "Null,", "'" & varTemp(i) & "',")
            '     NO_In         ������ü�¼.NO%Type,
            strSql = strSql & IIf(blnOld, "'" & varTemp(i) & "',", "Null,")
            '     ����Ա���_In סԺ���ü�¼.����Ա���%Type,
            strSql = strSql & "'" & UserInfo.��� & "',"
            '     ����Ա����_In סԺ���ü�¼.����Ա����%Type,
            strSql = strSql & "'" & UserInfo.���� & "',"
            '     �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
            strSql = strSql & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '     �����˷�_In   Number := 0(�����˷�_In:0-����תסԺ��������;1-�����˷�ģʽ:Ϊ1ʱ:��Ժ����id_In����ҳID_IN���Բ���)
            strSql = strSql & "1,"
            '     ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
            strSql = strSql & "Null,"
            '     ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null
            strSql = strSql & "Null,"
            '     ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null
            strSql = strSql & IIf(picBack.Visible, "'" & cboStyle.Text & "'", "Null") & ","
           
           strAllBalance = strAllBalance & "," & lng����ID
           cllPro.Add Array(strSql, lng����ID, varTemp(i), CStr(varTemp(i)))
    Next
    mstrThreeSwapBalance = ""
    mstrThreeSwapCardType = ""
    mstrThreeSwapMoney = ""
    
     If intInsure <> 0 And bln�൥��һ�ν��� Then
        On Error GoTo errH: blnTrans = True
        
        gcnOracle.BeginTrans
            '�����һ�ſ�ʼ��
        For i = cllPro.Count To 1 Step -1
            blnExecuteThreeSwap = False
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & ")", Me.Caption)
            
            If ExecuteThreeSwap(Val(cllPro(i)(1)), lng����ID, str������ˮ��, str����˵��) = True Then
                blnExecuteThreeSwap = True
            End If
            
            'Zl_����תסԺ_����������
            strSql = "Zl_����תסԺ_����������("
            '  �������_In   ����Ԥ����¼.�������%Type,
            strSql = strSql & IIf(blnOld, "Null,", "'" & cllPro(i)(2) & "',")
            '  No_In         סԺ���ü�¼.NO%Type,
            strSql = strSql & IIf(blnOld, "'" & cllPro(i)(2) & "',", "Null,")
            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
            strSql = strSql & "'" & UserInfo.��� & "',"
            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
            strSql = strSql & "'" & UserInfo.���� & "',"
            '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
            strSql = strSql & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  �����˷�_In   Number := 0,
            strSql = strSql & "" & 1 & ","
            '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
            strSql = strSql & "Null,"
            '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
            strSql = strSql & "Null,"
            '  �����˷�_In   Number := 0,
            strSql = strSql & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
            '  ����ID_In     סԺ���ü�¼.����id%Type)
            strSql = strSql & "" & lng����ID & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "����������")
        Next
        
        '�Ȳ���Ʊ�ݣ�ҽ���ӿڲ���ȡ��
        If blnҽ���ӿڴ�ӡƱ�� Then
            strSql = "zl_�����շѼ�¼_RePrint('" & strFirstNo & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                "To_Date('" & Format(strDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        End If
        strAdvance = strAllBalance
        If Not gclsInsure.ClinicDelSwap(Val(cllPro(cllPro.Count)(1)), , intInsure, strAdvance) Then
            GoTo errH
        Else
            blnTransMedicare = True
        End If
        
        If Not (strAdvance = strAllBalance Or strAdvance = "") Then
            '���ݷ��صĽ�����Ϣ������Ԥ����¼��strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2:���...
            '�ȷ�̯��ÿ�ŵ�����
            Set rsTmp = GetBalanceSet
            varBalance = Split(strAdvance, "||")
            For i = 0 To UBound(varBalance)
                str���㷽ʽ = Split(varBalance(i), "|")(0)
                cur������ = -1 * Val(Split(varBalance(i), "|")(1))
                For k = 0 To UBound(varTemp)
                    cur�ɷ���� = Getʵ�ս��(varTemp(k))
                    rsTmp.Filter = "�������=" & k
                    For j = 1 To rsTmp.RecordCount
                        cur�ɷ���� = cur�ɷ���� - rsTmp!������
                        rsTmp.MoveNext
                    Next
                    If cur�ɷ���� > 0 Then
                        If cur�ɷ���� <= cur������ Then
                            cur������ = cur������ - cur�ɷ����
                        Else
                            cur�ɷ���� = cur������
                            cur������ = 0
                        End If
                        rsTmp.AddNew
                        rsTmp!������� = k
                        rsTmp!���㷽ʽ = str���㷽ʽ
                        rsTmp!������ = cur�ɷ����
                        rsTmp.Update
                        
                        If cur������ = 0 Then Exit For
                    End If
                Next
            Next
            
            For k = 0 To UBound(varTemp)
                strBalance = ""
                cur����� = 0
                cur��� = Getʵ�ս��(varTemp(k))
                
                rsTmp.Filter = "�������=" & k
                For i = 1 To rsTmp.RecordCount
                    strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!���㷽ʽ & "|" & -1 * rsTmp!������
                    cur��� = cur��� - rsTmp!������
                    rsTmp.MoveNext
                Next

                '��Ϊָ���Ľ��㷽ʽ��������ֽ𣬿��ܲ����µ������
                'If cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1 Then
                    cur������ = Format(CentMoney(cur���), "0.00")
                    cur����� = cur������ - cur���
'                Else
'                    cur������ = cur���
'                End If
                cur�˿�ϼ� = cur�˿�ϼ� + cur������
                lng����ID = GetDelBalanceID(varTemp(k))
                strSql = "zl_�����շѽ���_Update(" & lng����ID & ",'" & "�ֽ�" & "|" & -1 * cur������ & "| ',0,'" & strBalance & "'," & -1 * cur����� & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
     Else
         '�����һ�ſ�ʼ��
        For i = cllPro.Count To 1 Step -1
            gcnOracle.BeginTrans: On Error GoTo errH: blnTrans = True
            mstrThreeSwapBalance = ""
            mstrThreeSwapCardType = ""
            mstrThreeSwapMoney = ""
            blnExecuteThreeSwap = False
            
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            Call zlDatabase.ExecuteProcedure(CStr(cllPro(i)(0)) & lng����ID & ")", Me.Caption)
            
            blnTransMedicare = False
            If intInsure <> 0 Then                    '����ҽ���ӿ�
                  If blnYB�������� Then
                        strAdvance = cllPro.Count & "|" & i
                        If Not gclsInsure.ClinicDelSwap(CStr(cllPro(i)(1)), True, intInsure, strAdvance) Then
                            GoTo errH
                        Else
                            blnTransMedicare = True
                        End If
                    End If
            End If
            gcnOracle.CommitTrans: blnTrans = False
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
            
            If ExecuteThreeSwap(Val(cllPro(i)(2)), lng����ID, str������ˮ��, str����˵��) = True Then
                blnExecuteThreeSwap = True
            End If
            
            'Zl_����תסԺ_����������
            strSql = "Zl_����תסԺ_����������("
            '  �������_In   ����Ԥ����¼.�������%Type,
            strSql = strSql & IIf(blnOld, "Null,", "'" & cllPro(i)(2) & "',")
            '  No_In         סԺ���ü�¼.NO%Type,
            strSql = strSql & IIf(blnOld, "'" & cllPro(i)(2) & "',", "Null,")
            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
            strSql = strSql & "'" & UserInfo.��� & "',"
            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
            strSql = strSql & "'" & UserInfo.���� & "',"
            '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
            strSql = strSql & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
            '  �����˷�_In   Number := 0,
            strSql = strSql & "" & 1 & ","
            '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
            strSql = strSql & "Null,"
            '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
            strSql = strSql & "Null,"
            '  �����˷�_In   Number := 0,
            strSql = strSql & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
            '  ����ID_In     סԺ���ü�¼.����id%Type)
            strSql = strSql & "" & lng����ID & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "����������")
            
            strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & cllPro(i)(0)
        Next
     End If
    
    If intInsure <> 0 And bln�˷Ѻ��ӡ�ص� And InStr(1, mstrPrivs, ";ҽ���˷ѻص�;") > 0 Then
        '����:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO='" & strFirstNo & "'", 2)
    End If
    ExecuteDelBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
    
    If Err.Number <> 0 Then Call SaveErrLog
    
    '�ж���ʾ,����ӡ�������˷Ѻ��ٴ�ӡ���Լ�ѡ���ش�
    If strDelNOs <> "" Then
        MsgBox "����[" & strNos & "]�˷�ʧ�ܡ����ǣ�����[" & strDelNOs & "]�ѳɹ��˷ѡ�" & vbCrLf & _
            "����δ��ӡ�����ִ��ʧ�ܵĵ��������˷ѣ�", vbInformation, gstrSysName
    End If
    Exit Function
End Function

Private Function ExecuteThreeSwap(lngBalance As Long, lng����ID As Long, Optional ByRef str������ˮ�� As String, Optional ByRef str����˵�� As String) As Boolean
    Dim strSql As String, rsTemp As ADODB.Recordset, strBalanceIDs As String, rsTotal As ADODB.Recordset
    Dim dblMoney As Double
    If mobjSquare Is Nothing Then Exit Function
    strSql = _
        " Select a.�����id, a.����, Min(a.����id) As ����id, Sum(a.��Ԥ��) As ��Ԥ��, Min(a.������ˮ��) As ������ˮ��, Min(a.����˵��) As ����˵��" & vbNewLine & _
        " From ����Ԥ����¼ A, ���㷽ʽ B," & vbNewLine & _
        "      (Select Distinct k.����id" & vbNewLine & _
        "        From ����Ԥ����¼ I, ������ü�¼ J, ������ü�¼ K" & vbNewLine & _
        "        Where i.������� = [1] And i.��¼���� = 3 And i.����id = j.����id And k.No = j.No And Mod(k.��¼����, 10) = 1) C, ҽ�ƿ���� D" & vbNewLine & _
        " Where a.���㷽ʽ = b.���� And b.���� = 7 And a.����id = c.����id And a.�����id = d.id And d.�Ƿ����� = 0 And a.У�Ա�־ <> 1" & vbNewLine & _
        " Group By a.�����id, a.����" & vbNewLine & _
        " Having Sum(a.��Ԥ��) <> 0"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalance)
    
'    strSQL = _
'        " Select  Sum(a.��Ԥ��) As ��Ԥ�� " & vbNewLine & _
'        " From ����Ԥ����¼ A, ���㷽ʽ B," & vbNewLine & _
'        "      (Select Distinct k.����id" & vbNewLine & _
'        "        From ����Ԥ����¼ I, ������ü�¼ J, ������ü�¼ K" & vbNewLine & _
'        "        Where i.������� = [1] And i.��¼���� = 3 And i.����id = j.����id And k.No = j.No And Mod(k.��¼����, 10) = 1) C" & vbNewLine & _
'        " Where a.���㷽ʽ = b.���� And a.����id = c.����id" & vbNewLine & _
'        " Having Sum(a.��Ԥ��) <> 0"
'
'    Set rsTotal = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBalance)
    
    If rsTemp.RecordCount = 0 Then Exit Function
'    If rsTotal.RecordCount = 0 Then Exit Function
    
'    If Val(NVL(rsTemp!��Ԥ��)) > Val(NVL(rsTotal!��Ԥ��)) Then
'        dblMoney = Val(NVL(rsTotal!��Ԥ��))
'    Else
    dblMoney = Val(NVL(rsTemp!��Ԥ��))
'    End If
    
    Do While Not rsTemp.EOF
        strBalanceIDs = "3|" & Val(NVL(rsTemp!����ID))
        If mobjSquare.zlReturnCheck(Me, mlngModule, Val(NVL(rsTemp!�����ID)), False, NVL(rsTemp!����), _
            strBalanceIDs, dblMoney, str������ˮ��, str����˵��, "3|" & lng����ID) = False Then Exit Function
        If mobjSquare.zlReturnMoney(Me, mlngModule, Val(NVL(rsTemp!�����ID)), False, NVL(rsTemp!����), _
            strBalanceIDs, dblMoney, str������ˮ��, str����˵��, "3|" & lng����ID) = False Then Exit Function
        mstrThreeSwapBalance = mstrThreeSwapBalance & "|" & lngBalance
        mstrThreeSwapCardType = mstrThreeSwapCardType & "|" & Val(NVL(rsTemp!�����ID))
        mstrThreeSwapMoney = mstrThreeSwapMoney & "|" & dblMoney
        rsTemp.MoveNext
    Loop
    ExecuteThreeSwap = True
End Function

Public Function Getʵ�ս��(ByVal strNO As String) As Currency
    Dim i As Long, cur��� As Currency
    With vsFee
        cur��� = 0
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("�������")) = strNO Then
                cur��� = cur��� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
            End If
        Next
        Getʵ�ս�� = cur���
    End With
End Function

Public Sub ClearData()
    Dim i As Integer
    vsFee.Clear 1
    vsFee.Rows = 2
    vsfInvoice.Clear 1
    vsfInvoice.Rows = 2
    vsBalanceStyle.Clear 1
    vsBalanceStyle.Rows = 2
    For i = 1 To vsBalance.Cols - 1
        vsBalance.TextMatrix(0, i) = ""
        vsBalance.TextMatrix(1, i) = ""
    Next i
    txtSum.Text = 0
    picBack.Visible = False
End Sub

Private Sub vsFee_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsFee_DblClick()
    With vsFee
        If .TextMatrix(.Row, .ColIndex("���")) = "" Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("ѡ��")) = "" Then
            Call SetRowSelected(.Row, True)
        Else
            Call SetRowSelected(.Row, False)
        End If
    End With
    Call SetBlanceShow
    Call CalcSUMMony
End Sub

Private Sub vsFee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        With vsFee
            If .TextMatrix(.Row, .ColIndex("ѡ��")) = "" Then
                Call SetRowSelected(.Row, True)
            Else
                Call SetRowSelected(.Row, False)
            End If
        End With
        Call SetBlanceShow
        Call CalcSUMMony
    End If
End Sub
