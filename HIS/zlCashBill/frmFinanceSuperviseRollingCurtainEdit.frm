VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFinanceSuperviseRollingCurtainEdit 
   Caption         =   "�����տ"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinanceSuperviseRollingCurtainEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11775
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   660
      ScaleHeight     =   1905
      ScaleWidth      =   2685
      TabIndex        =   6
      Top             =   3195
      Width           =   2685
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   870
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   1860
         _cx             =   3281
         _cy             =   1535
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinanceSuperviseRollingCurtainEdit.frx":6852
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
   Begin VB.PictureBox picDown 
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
      Height          =   2340
      Left            =   120
      ScaleHeight     =   2340
      ScaleWidth      =   11490
      TabIndex        =   23
      Top             =   5160
      Width           =   11490
      Begin VB.CommandButton cmdCashMoney 
         Caption         =   "�㳮(&D)"
         Height          =   350
         Left            =   -15
         TabIndex        =   32
         Top             =   1875
         Width           =   1100
      End
      Begin VB.TextBox txtActual 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4800
         MaxLength       =   16
         TabIndex        =   18
         Top             =   810
         Width           =   2625
      End
      Begin VB.TextBox txtLendTotal 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   8640
         TabIndex        =   15
         Top             =   405
         Width           =   2625
      End
      Begin VB.TextBox txtBorrowTotal 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   4800
         TabIndex        =   13
         Top             =   405
         Width           =   2625
      End
      Begin VB.TextBox txtPrepay 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   975
         TabIndex        =   11
         Top             =   390
         Width           =   2625
      End
      Begin VB.TextBox txtMemo 
         Height          =   350
         Left            =   975
         MaxLength       =   500
         TabIndex        =   9
         Top             =   -15
         Width           =   10305
      End
      Begin VB.TextBox txtTime 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   975
         TabIndex        =   24
         Top             =   1245
         Width           =   2625
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9975
         TabIndex        =   21
         Top             =   1875
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   8775
         TabIndex        =   20
         Top             =   1875
         Width           =   1100
      End
      Begin VB.Label lblRemainMoney 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   8655
         TabIndex        =   31
         Top             =   810
         Width           =   2625
      End
      Begin VB.Label lblSupposeMoney 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   975
         TabIndex        =   30
         Top             =   810
         Width           =   2625
      End
      Begin VB.Label lblLendTotal 
         AutoSize        =   -1  'True
         Caption         =   "����ϼ�"
         Height          =   210
         Left            =   7800
         TabIndex        =   14
         Top             =   450
         Width           =   840
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "�տ�ʱ��"
         Height          =   210
         Left            =   45
         TabIndex        =   25
         Top             =   1275
         Width           =   840
      End
      Begin VB.Label lblRemain 
         AutoSize        =   -1  'True
         Caption         =   "�����ݴ�"
         Height          =   210
         Left            =   7815
         TabIndex        =   19
         Top             =   885
         Width           =   840
      End
      Begin VB.Label lblActual 
         AutoSize        =   -1  'True
         Caption         =   "�ֽ�ʵ��"
         Height          =   210
         Left            =   3960
         TabIndex        =   17
         Top             =   885
         Width           =   840
      End
      Begin VB.Label lblSuppose 
         AutoSize        =   -1  'True
         Caption         =   "�ֽ�Ӧ��"
         Height          =   210
         Left            =   60
         TabIndex        =   16
         Top             =   885
         Width           =   840
      End
      Begin VB.Label lblBorrowTotal 
         AutoSize        =   -1  'True
         Caption         =   "���ϼ�"
         Height          =   210
         Left            =   3960
         TabIndex        =   12
         Top             =   450
         Width           =   840
      End
      Begin VB.Label lblPrepay 
         AutoSize        =   -1  'True
         Caption         =   "��Ԥ��"
         Height          =   210
         Left            =   255
         TabIndex        =   10
         Top             =   450
         Width           =   630
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "ժҪ"
         Height          =   210
         Left            =   465
         TabIndex        =   8
         Top             =   45
         Width           =   420
      End
      Begin VB.Line linMain 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   10440
         Y1              =   1650
         Y2              =   1650
      End
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   75
      ScaleHeight     =   825
      ScaleWidth      =   11175
      TabIndex        =   22
      Top             =   510
      Width           =   11175
      Begin VB.TextBox txtGroups 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   3735
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   450
         Width           =   2490
      End
      Begin VB.ComboBox cboNO 
         Height          =   330
         Left            =   8925
         TabIndex        =   26
         Top             =   75
         Width           =   2040
      End
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   473
         Width           =   1785
      End
      Begin VB.Label lblGroups 
         AutoSize        =   -1  'True
         Caption         =   "��Ա����"
         Height          =   240
         Left            =   2775
         TabIndex        =   29
         Top             =   510
         Width           =   960
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         Height          =   210
         Left            =   8565
         TabIndex        =   27
         Top             =   135
         Width           =   210
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "�ɿ��"
         Height          =   210
         Left            =   2700
         TabIndex        =   2
         Top             =   525
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "�ɿ���"
         Height          =   210
         Left            =   135
         TabIndex        =   0
         Top             =   525
         Width           =   630
      End
   End
   Begin VB.PictureBox picRollingCurtain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   2640
      ScaleHeight     =   1740
      ScaleWidth      =   8370
      TabIndex        =   4
      Top             =   1650
      Width           =   8370
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   33
         Top             =   60
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinanceSuperviseRollingCurtainEdit.frx":68CC
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   930
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   1640
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinanceSuperviseRollingCurtainEdit.frx":6E1A
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFinanceSuperviseRollingCurtainEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mstr����IDs As String, mlngGroupID As Long
Private mlng�ɿ���ID As Long, mstr�ɿ��� As String
Private Enum mPaneIndex
    EM_PN_��ͷ��Ϣ = 1
    EM_PN_������Ϣ = 2
    EM_PN_������Ϣ = 3
    EM_PN_��β��Ϣ = 4
End Enum
Private mblnOK As Boolean
Private mblnNotBrush As Boolean
Private mrsBalance As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnChange  As Boolean
Private Sub LoadBalance(ByVal blnReOpenRecord As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ�����Ϣ
    '���:blnReOpenRecord-���´򿪼�¼��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-10 14:50:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim strSQL As String, bytType As Byte, i As Long
   Dim lng����ID As Long, blnSel As Boolean, str����IDs As String
   Dim str���㷽ʽ As String, lngRow As Long, blnFind As Boolean
   Dim dblTotal(0 To 3) As Double
   
    On Error GoTo errHandle
    
    If mrsBalance Is Nothing Or blnReOpenRecord Then
        strSQL = "" & _
        "   Select /*+ rule */ decode(nvl(M.����,0),1,1,2,2,3,10,4,11,4) as ���,A.ID as �ս�ID,  " & _
        "           b.���㷽ʽ,b.���,b.�����,b.���,b.�����," & _
        "           a.��Ԥ���� as ��Ԥ��,A.����ϼ� as ���ϼ�,A.����ϼ� " & _
        "   From ��Ա�սɼ�¼ A, ��Ա�ս���ϸ B,���㷽ʽ M, Table(f_Num2list([2])) J" & _
        "   Where a.Id = b.�ս�id  And A.��¼����=[1]  And A.ID=J.Column_Value and B.���㷽ʽ=M.����(+) and nvl(���,0)<>0 " & _
        "   Order by ���,���㷽ʽ"
        bytType = IIf(mlngGroupID <> 0, 3, 1)
        Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytType, mstr����IDs)
    End If
    For i = 0 To 3
        dblTotal(i) = 0
    Next
    With vsRollingCurtain
        For i = 1 To .Rows - 1
            lng����ID = Val(.TextMatrix(i, .ColIndex("ID")))
            blnSel = Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0
            If blnSel And lng����ID <> 0 Then
                '��Ҫ����
                str����IDs = str����IDs & "," & lng����ID
                dblTotal(0) = dblTotal(0) + Val(.TextMatrix(i, .ColIndex("��Ԥ����")))
                dblTotal(1) = dblTotal(1) + Val(.TextMatrix(i, .ColIndex("����ϼ�")))
                dblTotal(2) = dblTotal(2) + Val(.TextMatrix(i, .ColIndex("����ϼ�")))
            End If
        Next
    End With
 
    With vsBalance
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If str����IDs = "" Then GoTo goEnd:
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            str���㷽ʽ = NVL(mrsBalance!���㷽ʽ)
            If Val(NVL(mrsBalance!���)) <> 0 _
                And InStr(str����IDs & ",", "," & NVL(mrsBalance!�ս�ID) & ",") > 0 Then
                blnFind = False
                For i = 1 To .Rows - 1
                    If str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))) Then
                        blnFind = True: lngRow = i: Exit For
                    End If
                Next
                If blnFind = False Then
                    If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = "" Then
                         lngRow = .Rows - 1
                    Else
                        .Rows = .Rows + 1: lngRow = .Rows - 1
                    End If
                End If
                .TextMatrix(lngRow, .ColIndex("���")) = NVL(mrsBalance!���)
                .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = str���㷽ʽ
                .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(.TextMatrix(lngRow, .ColIndex("���"))) + Val(NVL(mrsBalance!���)), gstrDec)
                
                If InStr(Mid(str����IDs, 2), ",") > 0 Then
                    'һ����ȡ���ʱ����Ҫ����¼��
                    .TextMatrix(lngRow, .ColIndex("�������")) = ""
                Else
                    'ֻ���һ�����ʼ�¼��ȡʱ����ȡԭ�������
                    .TextMatrix(lngRow, .ColIndex("�������")) = NVL(mrsBalance!�����)
                End If
                If Val(NVL(mrsBalance!���)) = 1 Then
                    '�ֽ�ϼ�
                    dblTotal(3) = dblTotal(3) + Val(NVL(mrsBalance!���))
                End If
            End If
            mrsBalance.MoveNext
        Loop
goEnd:
        .Cell(flexcpSort, 1, .ColIndex("���"), .Rows - 1, .ColIndex("���")) = flexSortNumericAscending
        .Redraw = flexRDBuffered
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .ColWidth(.ColIndex("�������")) = 3000
    End With
    '�ָ�������
    'zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False
    '���غϼ�����
    txtPrepay.Text = Format(dblTotal(0), "##0.00;-##0.00;;")
    txtBorrowTotal.Text = Format(dblTotal(1), "##0.00;-##0.00;;")
    txtLendTotal.Text = Format(dblTotal(2), "##0.00;-##0.00;;")
    lblSupposeMoney.Caption = Format(dblTotal(3), "##0.00;-##0.00;;")
    txtActual.Text = Format(dblTotal(3), "##0.00;-##0.00;;")
    lblRemainMoney.Caption = Format(0, "##0.00;-##0.00;0.00;0.00")
    txtActual.Enabled = dblTotal(3) <> 0 And mlngGroupID = 0
    txtActual.BackColor = IIf(txtActual.Enabled, &H80000005, txtLendTotal.BackColor)
    
    Exit Sub
errHandle:
    vsBalance.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function LoadGroup() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�������Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-24 12:16:41
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If mlngGroupID = 0 Then LoadGroup = True: Exit Function
    '��ȡ������
    strSQL = " " & _
    "   Select a.Id As ����, a.������, a.����, b.���� As �鸺����,A.˵��" & _
    "   From ����ɿ���� A, ��Ա�� B " & _
    "   Where a.������id = b.Id and A.ID=[1] " & _
    "   Order By a.������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������Ϣ", mlngGroupID)
    If rsTemp.RecordCount <> 0 Then
        txtGroups.Text = NVL(rsTemp!������)
    End If
    LoadGroup = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Public Function zlShowMe(ByVal frmMain As Object, _
    ByVal lngModule As String, ByVal strPrivs As String, _
    ByVal str�ɿ��� As String, ByVal lng�ɿ���ID As Long, ByVal str����IDs As String, _
    Optional ByVal lngGroupID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '       lngModule-ģ���
    '       strPrivs-Ȩ�޴�
    '       str����IDs-����Ҫ�տ������IDS����
    '       lngGroupID>0:��Բ������տ�(��������ID)
    '����:�տ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-10 14:08:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mlngModule = lngModule: mstrPrivs = strPrivs
    mstr����IDs = str����IDs: mlngGroupID = lngGroupID
    mstr�ɿ��� = str�ɿ���: mlng�ɿ���ID = lng�ɿ���ID
    
    Call ClearData
    Call SetCtrlEnable
    txtName.Text = mstr�ɿ���
    'If LoadDept = False Then Unload Me: Exit Function
    If LoadGroup = False Then Unload Me: Exit Function
    If LoadCollectData = False Then Unload Me: Exit Function
    mblnChange = False
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowMe = mblnOK
End Function

Private Function LoadDept() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽɿ��˲�����Ϣ
    '����:���˺�
    '����:2013-09-11 14:05:08
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
        
    strSQL = "" & _
    "   Select Distinct a.Id, a.����, a.����,b.ȱʡ" & vbNewLine & _
    "   From ���ű� a, ������Ա b" & vbNewLine & _
    "   Where a.Id = b.����id And b.��ԱID=[1] " & vbNewLine & _
     "              And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "               And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
    "   Order By a.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ɿ���ID)
    With cboDept
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!����) & "-" & rsTemp!����
            .ItemData(.NewIndex) = Val(NVL(rsTemp!ID))
            If Val(NVL(rsTemp!ȱʡ)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 And .ListCount <> 0 Then .ListIndex = 0
    End With


    LoadDept = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '����:���˺�
    '����:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
      Dim i As Long, strHead As String, varData As Variant
   Dim lngWidth As Long
    strHead = "����,ѡ��,ID,���ʵ���,��ʼʱ��,��ֹʱ��,������,����ʱ��,�տ�Ա,�տ��,��Ԥ����,����ϼ�,����ϼ�,С���տ���,С���տ�ʱ��,����˵��"
    varData = Split(strHead, ",")
    With vsRollingCurtain
        .Clear: .Rows = 2: .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = varData(i)
            If .ColKey(i) = "����" Then .TextMatrix(0, i) = ""
            If .ColKey(i) = "����" Or .ColKey(i) = "ѡ��" Or .ColKey(i) = "ID" Or .ColKey(i) = "�տ��" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "������" Or .ColKey(i) = "��Ԥ����" Or .ColKey(i) = "����ϼ�" Or .ColKey(i) = "����ϼ�" Then .ColHidden(i) = True
            If .ColKey(i) = "���ʵ���" Or .ColKey(i) = "��ʼʱ��" Or .ColKey(i) = "��ֹʱ��" Or .ColKey(i) = "����ʱ��" Then .ColData(i) = "1|0"
            If .ColKey(i) = "�տ��" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "�տ�Ա" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ʵ���" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "ѡ��" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColDataType(i) = flexDTBoolean
        
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        lngWidth = .ColWidth(.ColIndex("ѡ��"))
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False
        .ColWidth(.ColIndex("ѡ��")) = lngWidth
        .Editable = flexEDKbdMouse
    End With
    
    With vsBalance
           Set .Font = Me.Font
           .Clear 1
           .Cols = 4: .Rows = 2
           .FixedRows = 1
           .TextMatrix(0, 0) = "���"
           .TextMatrix(0, 1) = "���㷽ʽ"
           .TextMatrix(0, 2) = "���"
           .TextMatrix(0, 3) = "�������"
           For i = 0 To .Cols - 1
               .ColKey(i) = .TextMatrix(0, i)
               If i = .ColIndex("���") Then
                   .ColAlignment(i) = flexAlignRightCenter
               Else
                   .ColAlignment(i) = flexAlignLeftCenter
               End If
               .FixedAlignment(i) = flexAlignCenterCenter
           Next
           .ColHidden(.ColIndex("���")) = True
           .AutoSizeMode = flexAutoSizeColWidth
           .AutoResize = True
           Call .AutoSize(0, .Cols - 1)
           
           .ColWidth(.ColIndex("�������")) = .ColWidth(.ColIndex("�������")) * 3
           .ExtendLastCol = False
           'zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False
           .Editable = flexEDKbdMouse
       End With
End Sub
Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2013-10-10 14:21:53
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitGrid
    txtMemo.Text = ""
    txtPrepay.Text = ""
    txtBorrowTotal.Text = ""
    txtLendTotal.Text = ""
    lblSupposeMoney.Caption = ""
    txtActual.Text = ""
    lblRemainMoney.Caption = ""
End Sub
Public Function LoadCollectData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����տ�����
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-26 11:38:15
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng����ID As Long, bytType As Byte, i As Long, lngWidth As Long
 
    On Error GoTo errHandle
    txtTime.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If mstr����IDs = "" Then
        MsgBox "�㵱δѡ���κ����ʼ�¼�����ܽ����տ����Ա!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
     
    strSQL = "" & _
    "   Select /*+ rule */-1 as ѡ��,a.Id,a.No As ���ʵ���, a.��ʼʱ��, a.��ֹʱ��, a.�Ǽ��� As ������, a.�Ǽ�ʱ�� As ����ʱ��,  " & _
    "         a.�տ�Ա ,b.���� As �տ��, " & _
    "         ltrim(to_char(a.��Ԥ����,'9999999999990.00')) as ��Ԥ����, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�, " & _
    "         ltrim(to_char(a.����ϼ�,'9999999999990.00')) as ����ϼ�," & _
    "         a.С���տ���, To_Char(a.С���տ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As С���տ�ʱ��, " & _
    "         a.ժҪ As ����˵��" & _
    "  From ��Ա�սɼ�¼ A, ���ű� B, Table(f_Num2list([2])) J " & _
    "  Where a.�տ��id = b.Id(+) And A.ID=J.Column_Value And a.��¼���� = [1] " & _
    "               And A.����ʱ�� is Null and A.�����տ�ID is null  " & _
    "  Order by �Ǽ�ʱ�� desc,���ʵ��� desc,С���տ�ʱ�� desc"

    bytType = IIf(mlngGroupID <> 0, 3, 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytType, mstr����IDs)
    With vsRollingCurtain
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .FixedRows = 1
        .Redraw = flexRDNone
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = -1
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("���ʵ���")) = NVL(rsTemp!���ʵ���)
            .TextMatrix(.Rows - 1, .ColIndex("��ʼʱ��")) = NVL(rsTemp!��ʼʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("��ֹʱ��")) = NVL(rsTemp!��ֹʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = NVL(rsTemp!����ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("�տ�Ա")) = NVL(rsTemp!�տ�Ա)
            .TextMatrix(.Rows - 1, .ColIndex("��Ԥ����")) = NVL(rsTemp!��Ԥ����)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("����ϼ�")) = NVL(rsTemp!����ϼ�)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ���")) = NVL(rsTemp!С���տ���)
            .TextMatrix(.Rows - 1, .ColIndex("С���տ�ʱ��")) = NVL(rsTemp!С���տ�ʱ��)
            .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = NVL(rsTemp!����˵��)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "�տ��" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "�տ�Ա" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ʱ��" Or .ColKey(i) = "���ʵ���" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*�ϼ�" Or .ColKey(i) = "��Ԥ����" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "ѡ��" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColDataType(i) = flexDTBoolean
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        If .Rows > 2 Then .Rows = .Rows - 1
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        lngWidth = .ColWidth(.ColIndex("ѡ��"))
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False
         .ColWidth(.ColIndex("ѡ��")) = lngWidth
        If .Enabled And .Visible Then .SetFocus
        .Redraw = flexRDBuffered
    End With
    '���ؽ�������
    Call LoadBalance(True)
    LoadCollectData = True
    mblnNotBrush = False
    Exit Function
errHandle:
    vsBalance.Redraw = flexRDBuffered
    vsRollingCurtain.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Function
Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2013-10-10 11:51:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngHeight As Long, lngTemp As Long
    lngHeight = 1740 / Screen.TwipsPerPixelY
    With dkpMan
        lngTemp = 825 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(EM_PN_��ͷ��Ϣ, 100, lngTemp, DockLeftOf, Nothing)
        objPane.Title = "��ͷ": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngTemp: objPane.MaxTrackSize.Height = lngTemp
        objPane.Handle = picTop.hWnd
        
        Set objPane = .CreatePane(EM_PN_������Ϣ, 100, lngHeight, DockBottomOf, objPane)
        objPane.Title = "������Ϣ": objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngHeight
        objPane.Handle = picRollingCurtain.hWnd
        Set objPane = .CreatePane(EM_PN_������Ϣ, 400, 400, DockBottomOf, objPane)
        objPane.Title = "������ϸ"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
        objPane.MinTrackSize.Height = 1000 / Screen.TwipsPerPixelY
        
        lngTemp = 2340 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(EM_PN_��β��Ϣ, 100, lngTemp, DockBottomOf, objPane)
        objPane.Title = "��β": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngTemp: objPane.MaxTrackSize.Height = lngTemp
        objPane.Handle = picDown.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        .VisualTheme = ThemeOffice2003
    End With
End Function
 

'Private Sub cboDept_Click()
'    mblnChange = True
'End Sub

'Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then vsRollingCurtain.SetFocus
'End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdCashMoney_Click()
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ֽ�㳮
    '����:���˺�
    '����:2013-09-13 16:08:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim objCash As New clsChargeBill
    objCash.CheckCash Me, dblMoney
    Set objCash = Nothing
End Sub

Private Sub cmdOK_Click()
    Dim str����IDs As String, strδѡ����IDs As String
    Dim strNO As String
    str����IDs = GetSelRollingCurtainIds(strδѡ����IDs)
    If isValied(str����IDs) = False Then Exit Sub
    If SaveData(str����IDs, strNO) = False Then Exit Sub
    mblnOK = True
    '��ӡ�վ�
    cboNO.AddItem strNO
    Call BillPrint(strNO)
    mblnChange = False
    If strδѡ����IDs = "" Then Unload Me: Exit Sub
     '���¼�������
     mstr����IDs = strδѡ����IDs
    Call LoadCollectData
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
'    If cboDept.ListCount = 1 And cboDept.Enabled And cboDept.Visible Then
'        cboDept.SetFocus
'    End If
End Sub

Private Sub Form_Load()
    Call InitPanel
    RestoreWinState Me, App.ProductName
    mblnFirst = True
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Width < 12015 * 0.8 Then Width = 12015 * 0.8
    If Height < 9075 * 0.8 Then Height = 9075 * 0.8
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
    
    Err = 0: On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Set mrsBalance = Nothing
End Sub
Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Top = .ScaleTop
        vsBalance.Left = .ScaleLeft
        vsBalance.Width = .ScaleWidth
        vsBalance.Height = .ScaleHeight
    End With
End Sub
Private Sub picDown_Resize()
    Dim lngSplit As Long
    Err = 0: On Error Resume Next
    lngSplit = (picDown.ScaleWidth - 600) \ 3
    With picDown
        txtMemo.Width = .ScaleWidth - txtMemo.Left - 50
        linMain.X2 = .ScaleWidth
        txtPrepay.Width = lngSplit - txtPrepay.Left
        lblSupposeMoney.Width = txtPrepay.Width
        txtTime.Width = txtPrepay.Width
        
        lblBorrowTotal.Left = lngSplit + 300
        txtBorrowTotal.Left = lblBorrowTotal.Left + lblBorrowTotal.Width
        txtBorrowTotal.Width = lngSplit * 2 + 300 - txtBorrowTotal.Left
        lblActual.Left = lblBorrowTotal.Left
        txtActual.Left = txtBorrowTotal.Left
        txtActual.Width = txtBorrowTotal.Width
        
        lblLendTotal.Left = lngSplit * 2 + 600
        txtLendTotal.Left = lblLendTotal.Left + lblLendTotal.Width
        txtLendTotal.Width = .ScaleWidth - txtLendTotal.Left - 50
        lblRemain.Left = lblLendTotal.Left
        lblRemainMoney.Left = txtLendTotal.Left
        lblRemainMoney.Width = txtLendTotal.Width
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub
Private Sub picRollingCurtain_Resize()
    Err = 0: On Error Resume Next
    With picRollingCurtain
        vsRollingCurtain.Top = .ScaleTop
        vsRollingCurtain.Left = .ScaleLeft
        vsRollingCurtain.Width = .ScaleWidth
        vsRollingCurtain.Height = .ScaleHeight
    End With
End Sub

 

Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    cboNO.Left = picTop.ScaleWidth - cboNO.Width - 50
    lblNO.Left = cboNO.Left - lblNO.Width - 10
End Sub

Private Sub txtActual_Change()
    lblRemainMoney.Caption = Format(Val(lblSupposeMoney.Caption) - Val(txtActual.Text), "0.00")
    If Val(lblRemainMoney.Caption) <> 0 Then
        lblRemainMoney.ForeColor = vbRed
    Else
        lblRemainMoney.ForeColor = txtActual.ForeColor
    End If
    mblnChange = True
End Sub

Private Sub txtActual_GotFocus()
    zlCommFun.OpenIme False
    zlControl.TxtSelAll txtActual
End Sub

Private Sub txtActual_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtMemo_Change()
    mblnChange = True
End Sub

Private Sub txtMemo_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtMemo
End Sub
Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtMemo_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsRollingCurtain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("ѡ��")
            Call LoadBalance(False)
            mblnChange = True
        End Select
    End With
End Sub
Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("ѡ��")
            Cancel = True
        Case Else
            Exit Sub
        End Select
    End With
End Sub
Private Sub vsRollingCurtain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng����ID As Long
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("ѡ��")
            lng����ID = Val(.TextMatrix(Row, .ColIndex("ID")))
           ' If lng����ID = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub
Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
End Sub
Private Sub vsRollingCurtain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: vsBalance.SetFocus
End Sub

Private Sub vsRollingCurtain_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsRollingCurtain)
    vsRollingCurtain.Tag = "0"
End Sub
Private Sub vsRollingCurtain_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsRollingCurtain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsRollingCurtain, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsRollingCurtain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    'zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False, zlCheckPrivs(mstrPrivs, "��������")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    'zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False, zlCheckPrivs(mstrPrivs, "��������")
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBalance
        Select Case Col
        Case .ColIndex("�������")
            If .TextMatrix(Row, .ColIndex("���㷽ʽ")) Like "*��Ԥ��*" _
                Or .TextMatrix(Row, Col) Like "*���ϼ�*" _
                Or .TextMatrix(Row, Col) Like "*���ϼ�*" Then
                Cancel = True: Exit Sub
            End If
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub
Private Sub vsBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsBalance
        If .Col = .Cols - 1 And .Row = .Rows - 1 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("���㷽ʽ"), vsBalance.Cols - 1, False)
End Sub
Private Sub vsBalance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("���㷽ʽ"), vsBalance.Cols - 1, False)
End Sub
Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub
Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsBalance
        If Row <= 1 Then Exit Sub
            VsFlxGridCheckKeyPress vsBalance, Row, Col, KeyAscii, m�ı�ʽ
            If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0
    End With
End Sub
Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '������֤
    With vsBalance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("�������")
            If zlCommFun.ActualLen(strKey) > 10 Then
                MsgBox "������볬��,���ֻ������10���ַ���5������", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If InStr(1, strKey, "'") > 0 Or InStr(1, strKey, "|") > 0 Or InStr(1, strKey, ",") > 0 Then
                MsgBox "��������в��ܰ��������ַ�:',| ", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        Case Else
        End Select
    End With
End Sub
Private Sub SetCtrlEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�Enabled����
    '����:���˺�
    '����:2013-10-10 17:08:06
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtName.Enabled = False
    txtGroups.Visible = mlngGroupID <> 0
    lblGroups.Visible = mlngGroupID <> 0
 End Sub
Private Function isValied(ByVal strSel����IDs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ǰ�����ݺϷ��Լ��
    '���:strSel����IDs-ѡ�е�����IDs(����ö��ŷ���)
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-10 17:31:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim str����IDs As String, lng����ID As Long, dblMoney As Double, blnFind As Boolean
    
    On Error GoTo errHandle
    If strSel����IDs = "" Then
        MsgBox "�տ�ʱ,��������ѡ��һ�����ʼ�¼,���ܽ����տ�!", vbInformation, gstrSysName
        If vsRollingCurtain.Visible And vsRollingCurtain.Enabled Then vsRollingCurtain.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(strSel����IDs) > 4000 Then
        MsgBox "�տ�ʱ,ѡ������ʼ�¼����,��ȡ������������ʼ�¼���տ�!", vbInformation, gstrSysName
        If vsRollingCurtain.Visible And vsRollingCurtain.Enabled Then vsRollingCurtain.SetFocus
        Exit Function
    End If
    '�����:110281,����,2017/08/15,������˵�������޴�50���ַ�����Ϊ500���ַ�
    If zlCommFun.ActualLen(txtMemo.Text) > 500 Then
        MsgBox "ժҪ����,���ֻ������250�����ӻ�500���ַ�", vbInformation, gstrSysName
        If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
        Exit Function
    End If
    If InStr(1, txtMemo.Text, "'") > 0 Then
        MsgBox "ժҪ�в��ܰ���������!", vbInformation, gstrSysName
        If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
        Exit Function
    End If
'    If cboDept.ListIndex < 0 Then
'        MsgBox "δѡ��ɿ��!", vbInformation, gstrSysName
'        If cboDept.Visible And cboDept.Enabled Then cboDept.SetFocus
'        Exit Function
'    End If
    If Val(txtActual.Text) > Val(lblSupposeMoney.Caption) Then
        MsgBox "�ֽ�ʵ�ս��ܴ����ֽ�Ӧ�ս��!", vbInformation, gstrSysName
        If txtActual.Visible And txtActual.Enabled Then txtActual.SetFocus
        Exit Function
    End If
    
    With vsBalance
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("�������"))
            If zlCommFun.ActualLen(strTemp) > 10 Then
                MsgBox "������볬��,���ֻ������10���ַ���5������", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("�������")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") > 0 Or InStr(1, strTemp, "|") > 0 Or InStr(1, strTemp, ",") > 0 Then
                MsgBox "��������в��ܰ��������ַ�:',| ", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("�������")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
        Next
    End With
    '�ܽ����
    strSQL = "" & _
       "   Select  b.���㷽ʽ,sum(b.���) as ��� " & _
       "   From  ��Ա�ս���ϸ B, Table(f_Num2list([1])) J" & _
       "   Where  B.�ս�ID=J.Column_Value " & _
       "   Group by  B.���㷽ʽ"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr����IDs)
    With vsBalance
        For i = 1 To .Rows - 1
            strTemp = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            rsTemp.Filter = "���㷽ʽ='" & strTemp & "'"
'            If rsTemp.EOF And Val(txtPrepay.Text) = 0 And Val(txtBorrowTotal.Text) = 0 And Val(txtLendTotal.Text) = 0 Then
'               If MsgBox(" �ڽ�����ϸ�б���" & strTemp & "�Ľ��㷽ʽ " & vbCrLf & _
'                "��ѡ�е����ʼ�¼�в�����,��������Ϊ����ԭ����ɵ�," & vbCrLf & _
'                "Ϊ�˱�֤���ݵ�һ����,����Ҫ������ȡ����," & vbCrLf & _
'                "���Ƿ�Ҫ������ȡ���ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
'                Call LoadCollectData
'               End If
'                If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
'                Exit Function
'            End If
            If Not rsTemp.EOF Then
                dblMoney = Val(NVL(rsTemp!���))
                If dblMoney <> Val(.TextMatrix(i, .ColIndex("���"))) Then
                   If MsgBox(" �ڽ�����ϸ�б���" & strTemp & "�ĺϼ����� " & vbCrLf & _
                    "ѡ�е����ʼ�¼�ĺϼ�����һ��,��������Ϊ����ԭ����ɵ�," & vbCrLf & _
                    "Ϊ�˱�֤���ݵ�һ����,����Ҫ������ȡ����," & vbCrLf & _
                    "���Ƿ�Ҫ������ȡ���ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    Call LoadCollectData
                   End If
                    If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
                    Exit Function
                End If
            End If
        Next
        rsTemp.Filter = 0
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            strTemp = NVL(rsTemp!���㷽ʽ)
            dblMoney = Val(NVL(rsTemp!���))
            If dblMoney <> 0 Then
                blnFind = False
                For i = 1 To .Rows - 1
                    If strTemp = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))) Then
                        blnFind = True: Exit For
                    End If
                Next
                If Not blnFind Then
                    If MsgBox(" �ڽ�����ϸ�б��в�����" & strTemp & "�Ľ��㷽ʽ " & vbCrLf & _
                     "��������Ϊ����ԭ����ɵ�, Ϊ�˱�֤���ݵ�һ����,����Ҫ������ȡ����," & vbCrLf & _
                     "���Ƿ�Ҫ������ȡ���ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                     Call LoadCollectData
                    End If
                     If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
                     Exit Function
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData(ByVal str����IDs As String, ByRef strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:str����IDs-ѡ�е�����IDs(����ö��ŷָ�)
    '����:strNO-����ɹ���,���ص��տ�ݺ�
    '����:����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-10 18:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngID As String, str������Ϣ As String, str���㷽ʽ As String
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBalance
        For i = 1 To .Rows - 1
            str���㷽ʽ = .TextMatrix(i, .ColIndex("���㷽ʽ"))
            If str���㷽ʽ <> "" And Trim(.TextMatrix(i, .ColIndex("�������"))) <> "" Then
                str������Ϣ = str������Ϣ & "|" & str���㷽ʽ & "," & Trim(.TextMatrix(i, .ColIndex("�������")))
            End If
        Next
    End With
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 2)
    If zlCommFun.ActualLen(str������Ϣ) > 4000 Then
        MsgBox "�ڽ�����ϸ��Ϣ������Ľ����������㷽ʽ������,���ֻ��Ϊ4000���ַ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lngID = zlDatabase.GetNextId("��Ա�սɼ�¼")
    strNO = zlDatabase.GetNextNo(140)
    'Zl_�����տ��¼_Insert
    strSQL = "Zl_�����տ��¼_Insert("
    '  Id_In         In ��Ա�սɼ�¼.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In         In ��Ա�սɼ�¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  �տ�Ա_In     In ��Ա�սɼ�¼.�տ�Ա%Type,
    strSQL = strSQL & "'" & mstr�ɿ��� & "',"
    '  �տ��id_In In ��Ա�սɼ�¼.�տ��id%Type,
    strSQL = strSQL & "Null,"
'    strSQL = strSQL & cboDept.ItemData(cboDept.ListIndex) & ","
    '  �ɿ���id_In   In ��Ա�սɼ�¼.�ɿ���id%Type,
    strSQL = strSQL & "" & IIf(mlngGroupID = 0, "NULL", mlngGroupID) & ","
    '  �ݴ���_In   In ��Ա�սɼ�¼.��Ԥ����%Type,
    strSQL = strSQL & "" & IIf(mlngGroupID <> 0, "0", Val(Replace(lblRemainMoney.Caption, ",", ""))) & ","
    '  ժҪ_In       In ��Ա�սɼ�¼.ժҪ%Type,
    strSQL = strSQL & "" & IIf(Trim(txtMemo.Text) = "", "NULL", "'" & Trim(txtMemo.Text) & "'") & ","
    '  �Ǽ���_In     In ��Ա�սɼ�¼.�Ǽ���%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ǽ�ʱ��_In   In ��Ա�սɼ�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "sysdate,"
    '  ����ids_In    In Varchar2,����IDs_In:����ID1,����ID2,...
    strSQL = strSQL & "'" & str����IDs & "',"
    '  �������_In   In Varchar2:���㷽ʽ,�������|���㷽ʽ,�������,..
    strSQL = strSQL & "'" & str������Ϣ & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSelRollingCurtainIds(Optional ByRef strNotSelRollingCurtainIDs As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰѡ�е�����ID
    '����:strNotSelRollingCurtainIDs-δѡ�е�����IDs(�ö��ŷ���
    '����:ѡ�е�����ID
    '����:���˺�
    '����:2013-10-10 17:40:05
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim str����IDs As String, lng����ID As Long, i As Long
    On Error GoTo errHandle
   strNotSelRollingCurtainIDs = ""
    With vsRollingCurtain
        For i = 1 To .Rows - 1
            lng����ID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lng����ID <> 0 And Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Then
                str����IDs = str����IDs & "," & lng����ID
            ElseIf lng����ID <> 0 Then
                strNotSelRollingCurtainIDs = strNotSelRollingCurtainIDs & "," & lng����ID
            End If
        Next
    End With
    If strNotSelRollingCurtainIDs <> "" Then strNotSelRollingCurtainIDs = Mid(strNotSelRollingCurtainIDs, 2)
    If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    GetSelRollingCurtainIds = str����IDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub BillPrint(ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�տ��վݴ�ӡ
    '����:���˺�
    '����:2013-09-11 11:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    blnPrint = False
    If Not zlStr.IsHavePrivs(mstrPrivs, "�տ��վݴ�ӡ") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("�տ��վݴ�ӡ��ʽ", glngSys, mlngModule))     'ʹ��ҽ��վ����ز���
    Case 0    '����ӡ
        Exit Sub
    Case 1    '��������ӡ
        blnPrint = True
    Case 2    'ѡ���ӡ
        If MsgBox("���Ƿ�Ҫ��ӡ�ɿ��վݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            blnPrint = True
        End If
    End Select
    If blnPrint = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500", Me, "NO=" & strNO, "��¼����=4", 2)
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "������Ϣ�б�", False, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
