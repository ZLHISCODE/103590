VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRegistHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���Һ��嵥"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9030
   Icon            =   "frmRegistHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdQuery 
      Cancel          =   -1  'True
      Caption         =   "��ѯ(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7830
      TabIndex        =   10
      Top             =   480
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   2
      Top             =   885
      Width           =   10785
   End
   Begin VB.Frame Frame2 
      Height          =   90
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   10785
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7680
      TabIndex        =   0
      Top             =   5970
      Width           =   1230
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRegList 
      Height          =   4755
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   8805
      _cx             =   1964915883
      _cy             =   1964908739
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRegistHistory.frx":0442
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
      OutlineBar      =   4
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
   Begin MSComCtl2.DTPicker dtpTimes 
      Height          =   375
      Index           =   0
      Left            =   1170
      TabIndex        =   5
      Top             =   510
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   141688835
      CurrentDate     =   40722
   End
   Begin MSComCtl2.DTPicker dtpTimes 
      Height          =   375
      Index           =   1
      Left            =   5130
      TabIndex        =   7
      Top             =   510
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   141688835
      CurrentDate     =   40722
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1170
      TabIndex        =   12
      Top             =   150
      Width           =   150
   End
   Begin VB.Label lblDoctor 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��"
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
      Left            =   660
      TabIndex        =   11
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblSum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   9
      Top             =   6030
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "~"
      Height          =   45
      Left            =   3960
      TabIndex        =   8
      Top             =   690
      Width           =   135
   End
   Begin VB.Label lblTimes 
      Caption         =   "����ʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4140
      TabIndex        =   6
      Top             =   570
      Width           =   1005
   End
   Begin VB.Label lblTimes 
      Caption         =   "��ʼʱ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   570
      Width           =   1005
   End
End
Attribute VB_Name = "frmRegistHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrҽ������ As String
Private mdatRegDate As Date
Private mblnAppointment As Boolean
Private mlngModul As Long
Private mstrTittle As String
'-----------------------------------------------------------------------------------

Public Function ShowRegist(ByVal frmMain As Form, ByVal lngModul As Long, ByVal blnAppointment As Boolean, _
                            ByVal strҽ������ As String, ByVal datRegDate As Date) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ѯҽ���ĹҺ�lblDoctorԼ������Ϣ
    '��Σ�strҽ����lblDoctor- ҽ������lblDoctor
    '      datRegDate - ��ǰ�Һ�/ԤԼʱ�䣬��Ϊ��ѯ��ȱʡʱ��
    '���Σ�
    '���أ��ɹ�����true,���򷵻�False
    '���ƣ����ϴ�
    '���ڣ�2019/3/6 15:55
    '------------------------------------------------------------------------------------------------------------------------
    mlngModul = lngModul
    mstrҽ������ = strҽ������: mdatRegDate = datRegDate
    mblnAppointment = blnAppointment
    
    Me.Caption = IIf(blnAppointment, "ԤԼ�嵥", IIf(gSysPara.bln��Һ�ģʽ, "�����嵥", "�Һ��嵥")) & "��ѯ"
    lblName.Caption = strҽ������
    mstrTittle = "lblDoctorվ�Һ��嵥"
    Me.Show 1, frmMain
    ShowRegist = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    Dim strSql As String, rsList As ADODB.Recordset
    
    On Error GoTo errH
    If dtpTimes(0).Value > dtpTimes(1).Value Then
        MsgBox "����ʱ��С���˿�ʼʱ�䡣", vbInformation, gstrSysName
        If dtpTimes(1).Visible And dtpTimes(1).Enabled Then dtpTimes(1).SetFocus
        Exit Sub
    End If
    
    If DateDiff("d", dtpTimes(0).Value, dtpTimes(1).Value) > 30 Then
        If MsgBox("ѡ���ʱ�䷶Χ������30�죬�Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    
    Call gobjCommFun.ShowFlash("���ڼ��ز���ҽ�ƿ���Ϣ,���Ե�...", Me)
    
    strSql = "Select a.�ű�, b.���� As ����, a.No As ���ݺ�, a.����, c.���� As ��Ŀ, a.����, a.����ʱ�� as �Һ�ʱ��, a.�Ǽ�ʱ��" & vbNewLine & _
                "From ���˹Һż�¼ a, ���ű� b, �շ���ĿĿ¼ c" & vbNewLine & _
                "Where a.��¼���� = [1] And a.��¼״̬ = 1 And a.ִ���� = [2] And a.����ʱ�� Between [3] And [4] And" & vbNewLine & _
                "      a.ִ�в���id = b.Id And a.�Һ���Ŀid = c.Id And Nvl(a.��¼��־,0) <> -1" & vbNewLine & _
                " Order by a.�ű�, ����, a.����ʱ�� Desc"
    Set rsList = gobjDatabase.OpenSQLRecord(strSql, "�Һ���Ϣ", IIf(mblnAppointment, 2, 1), mstrҽ������, dtpTimes(0).Value, dtpTimes(1).Value)
    Set vsRegList.DataSource = rsList
    Call SetFeeListHead
    Call gobjCommFun.StopFlash
    lblSum.Caption = "�� " & rsList.RecordCount & "����¼"
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dtpTimes_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Dim datMinDate As Date
    
    Call SetFeeListHead(True)
    If mblnAppointment Then
        datMinDate = gobjDatabase.Currentdate
        dtpTimes(0).minDate = Format(datMinDate, "YYYY-MM-DD")
        dtpTimes(1).minDate = Format(datMinDate, "YYYY-MM-DD")
    End If
    dtpTimes(0).Value = Format(mdatRegDate, "YYYY-MM-DD")
    dtpTimes(1).Value = Format(mdatRegDate, "YYYY-MM-DD 23:59:59")
    
    Call cmdQuery_Click
End Sub

Private Sub vsRegList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsRegList, mstrTittle, "�Һ��嵥"
End Sub

Private Sub vsRegList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRegList
        If Col = .ColIndex("��־") Then Cancel = True
    End With
End Sub

Private Sub vsRegList_GotFocus()
    vsRegList.BackColorSel = &H8000000D
End Sub

Private Sub vsRegList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub vsRegList_LostFocus()
    vsRegList.BackColorSel = GRD_LOSTFOCUS_COLORSEL
End Sub
Private Sub vsRegList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsRegList, mstrTittle, "�Һ��嵥"
End Sub

Private Sub SetFeeListHead(Optional blnInitHead As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷�����Ϣ��ͷ
    '���:blnInitHead-�Ƿ��ʼ����
    '����:���˺�
    '����:2014-09-10 17:01:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strHead As String, i As Long, varData As Variant
    
    On Error GoTo errHandle
    With vsRegList
        .Redraw = flexRDNone
        
        If blnInitHead Then
            strHead = "�ű�|����|���ݺ�|����|�Һ���ĿID|���|�Һ�ʱ��|�Ǽ�ʱ��"
            
            .Clear
            .Rows = 2
            varData = Split(strHead, "|")
            .Cols = UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, i) = varData(i)
            Next
        ElseIf .Rows <= 1 Then
            .Clear 1
            .Rows = 2
        End If
        
        For i = 0 To .Cols - 1
            .ColKey(i) = UCase(Trim(.TextMatrix(0, i)))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            If .ColKey(i) = "�Һ���ĿID" Or .ColKey(i) = "���" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) Like "*" Then
                .ColAlignment(i) = flexAlignCenterCenter
            End If
        Next
        
        If mblnAppointment Then
            .TextMatrix(0, .ColIndex("�Һ�ʱ��")) = "ԤԼʱ��"
        ElseIf gSysPara.bln��Һ�ģʽ Then
            .TextMatrix(0, .ColIndex("�Һ�ʱ��")) = "����ʱ��"
        End If
        
        .FrozenCols = 2 '������
        .HighLight = flexHighlightWithFocus
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore mlngModul, vsRegList, mstrTittle, "�Һ��嵥"
        
        .RowHeight(0) = 350
        .ColWidth(.ColIndex("���ݺ�")) = 1000
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        If .TextMatrix(1, .ColIndex("�ű�")) <> "" Then Call SplitGroupToRegList
        
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    vsRegList.Redraw = flexRDBuffered
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SplitGroupToRegList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ϸ���ݷ�����ʾ
    '����:���˺�
    '����:2014-09-10 17:12:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    With vsRegList
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("�ű�"), , , &H8000000F, , True, "%s", , True
        '.Subtotal flexSTSum, .ColIndex("����"), , , &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("���ݺ�")
        .OutlineCol = .ColIndex("���ݺ�")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                .RowHeight(i) = 350
                
                .TextMatrix(i, .ColIndex("�ű�")) = Trim(.Cell(flexcpTextDisplay, i + 1, .ColIndex("�ű�")))
                 strTemp = "�ű�:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("�ű�"))
                 strTemp = strTemp & Space(2) & "����:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("����"))
                 
                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("���ݺ�"), i, .ColIndex("�Ǽ�ʱ��")) = 1
                 For j = 0 To .Cols - 1
                    If j > .ColIndex("���ݺ�") Then
                        .Cell(flexcpText, i, j) = strTemp
                        .Cell(flexcpFontBold, i, j) = True
                    End If
                 Next
            End If
        Next
        
        For j = 0 To .Cols - 1
            If j < .ColIndex("���ݺ�") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        .ColHidden(.ColIndex("�ű�")) = True
        .ColHidden(.ColIndex("����")) = True
    End With
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub
