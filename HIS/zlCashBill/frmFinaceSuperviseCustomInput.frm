VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFinaceSuperviseCustomInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ɿ�Ǽǿ�"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8625
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinaceSuperviseCustomInput.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtTotal 
      BackColor       =   &H80000004&
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
      Height          =   350
      Left            =   810
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6000
      Width           =   7665
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����(&S)"
      Height          =   350
      Left            =   1260
      TabIndex        =   21
      Top             =   7590
      Width           =   1590
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   20
      Top             =   7590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6015
      TabIndex        =   15
      Top             =   7590
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7215
      TabIndex        =   18
      Top             =   7590
      Width           =   1100
   End
   Begin VB.TextBox txtInputPerson 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Left            =   810
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6885
      Width           =   1785
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   330
      Left            =   5835
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6885
      Width           =   2625
   End
   Begin VB.TextBox txtMemo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   810
      MaxLength       =   500
      TabIndex        =   10
      Top             =   6450
      Width           =   7665
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
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
      Height          =   315
      Left            =   6540
      TabIndex        =   5
      Top             =   1230
      Width           =   1935
   End
   Begin VB.ComboBox cboDept 
      Height          =   330
      Left            =   3435
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1222
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.ComboBox cboNO 
      Height          =   330
      Left            =   6540
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBalance 
      Height          =   4305
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   8355
      _cx             =   14737
      _cy             =   7594
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
      FormatString    =   $"frmFinaceSuperviseCustomInput.frx":6852
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   975
      TabIndex        =   1
      Top             =   1230
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   194641923
      CurrentDate     =   41520
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "��  ��"
      Height          =   210
      Left            =   165
      TabIndex        =   7
      Top             =   6075
      Width           =   630
   End
   Begin VB.Label lblTittle 
      Alignment       =   2  'Center
      Caption         =   "����ɿ�Ǽǿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      TabIndex        =   19
      Top             =   195
      Width           =   8310
   End
   Begin VB.Line linMain 
      BorderColor     =   &H8000000C&
      X1              =   -30
      X2              =   10410
      Y1              =   7305
      Y2              =   7305
   End
   Begin VB.Label lblInputPerson 
      AutoSize        =   -1  'True
      Caption         =   "�Ǽ���"
      Height          =   210
      Left            =   165
      TabIndex        =   11
      Top             =   6945
      Width           =   630
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "�Ǽ�ʱ��"
      Height          =   210
      Left            =   4920
      TabIndex        =   13
      Top             =   6945
      Width           =   840
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "�ɿ�ʱ��"
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
      Left            =   180
      TabIndex        =   0
      Top             =   1297
      Width           =   720
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "ժ  Ҫ"
      Height          =   210
      Left            =   165
      TabIndex        =   9
      Top             =   6480
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "�ɿ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5910
      TabIndex        =   4
      Top             =   1275
      Width           =   630
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "�ɿ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2655
      TabIndex        =   2
      Top             =   1282
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   210
      Left            =   6270
      TabIndex        =   17
      Top             =   765
      Width           =   210
   End
End
Attribute VB_Name = "frmFinaceSuperviseCustomInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mblnOtherPerson As Boolean
Private mstr�ɿ��� As String, mlng�ɿ���ID  As Long
Private mrsBalance As ADODB.Recordset
Private mblnChange As Boolean '�Ƿ��û�������
Private mblnSuccess As Boolean
Private mblnFirst  As Boolean
Public Function EditCard(ByVal frmMain As Object, _
    ByVal str�ɿ��� As String, ByVal lng�ɿ���ID As Long, _
    ByVal lngModule As Long, ByVal strPrivs As String, Optional ByVal blnOtherPerson As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����༭���(�ֹ��ɿ�)
    '���:str�ɿ���-�ɿ���
    '       lng�ɿ���ID-�ɿ���ID
    '       blnOtherPerson-trueʱΪ������Ա�տ�;����Ϊ�շ�Ա���
    '����:
    '����:����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-11 18:08:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstr�ɿ��� = str�ɿ���: mlng�ɿ���ID = lng�ɿ���ID
    mlngModule = lngModule: mstrPrivs = strPrivs
    mblnOtherPerson = blnOtherPerson
    Call InitFace
    If LoadCollectData = False Then mblnChange = False: Unload Me: Exit Function
    mblnChange = False: mblnSuccess = False
    If frmMain Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmMain
    End If
    EditCard = mblnSuccess
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function
Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:���˺�
    '����:2013-10-11 16:00:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim datCurrnet  As Date
    txtName.Text = mstr�ɿ���: txtInputPerson.Text = UserInfo.����
    
    datCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    dtpDate.Value = datCurrnet
    dtpDate.MaxDate = datCurrnet
    
    Call InitGrid
    'Call LoadDept
End Sub
Private Function LoadCollectData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����տ�����
    '����:���سɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-11 18:14:31
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long
    On Error GoTo errHandle
        
    strSQL = "" & _
    "   Select decode(nvl(M.����,0),1,1,2,2,3,10,4,11,4) as ���, A.���㷽ʽ,A.��� " & _
    "   From ��Ա�ɿ���� A,���㷽ʽ M" & _
    "   Where A.���㷽ʽ=M.����(+)  and A.����=1 and nvl(A.���,0)<>0  " & _
    "           And  A.�տ�Ա =[1] " & _
    "   Order by ���,���㷽ʽ"
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���)
    If mrsBalance.RecordCount = 0 Then
        MsgBox "�շ�Ա��" & mstr�ɿ��� & "��û���ݴ��������нɿ������", vbExclamation, gstrSysName
        Exit Function
    End If
    
    With vsBalance
        .Clear 1: .Rows = IIf(mrsBalance.RecordCount = 0, 1, mrsBalance.RecordCount) + 1
        i = 1
        Do While Not mrsBalance.EOF
             .TextMatrix(i, .ColIndex("���")) = NVL(mrsBalance!���)
             .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(mrsBalance!���㷽ʽ)
             .Cell(flexcpData, i, .ColIndex("���㷽ʽ")) = Trim(NVL(mrsBalance!���㷽ʽ))
             .TextMatrix(i, .ColIndex("���")) = Format(Val(NVL(mrsBalance!���)), "###0.00;-###0.00;0.00;0.00")
             .TextMatrix(i, .ColIndex("�������")) = ""
            i = i + 1
            mrsBalance.MoveNext
        Loop
        .ColComboList(.ColIndex("���㷽ʽ")) = .BuildComboList(mrsBalance, "���㷽ʽ,���", "���㷽ʽ")
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoResize = True
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False
    End With
    Call CalcTotal
    LoadCollectData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2013-10-11 15:59:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsBalance
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
           .ExtendLastCol = True
           .AutoSizeMode = flexAutoSizeColWidth
           .AutoResize = True
           Call .AutoSize(0, .Cols - 1)
           zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False
           .Editable = flexEDKbdMouse
    End With
End Sub
Private Sub CalcTotal()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ɿ��ܽ��
    '����:���˺�
    '����:2013-10-11 16:14:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double, i As Integer
    With vsBalance
        For i = 1 To .Rows - 1
             dblTemp = dblTemp + Val(.TextMatrix(i, .ColIndex("���")))
        Next
    End With
    txtTotal.Text = Format(dblTemp, "###0.00;-###0.00;0;") & "Ԫ" & IIf(dblTemp = 0, "", " ��" & zlCommFun.UppeMoney(dblTemp) & "��")
End Sub

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

'Private Sub cboDept_Click()
'    mblnChange = True
'End Sub

'Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
'
'End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_Load()
    mblnFirst = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNO As String
    On Error GoTo errHandle
    If isValied() = False Then Exit Sub
    If SaveData(strNO) = False Then Exit Sub
    mblnChange = False: mblnSuccess = True
    Call BillPrint(strNO)
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdPrintSet_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me
End Sub
Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txtMemo_Change()
    mblnChange = True
End Sub
Private Sub txtMemo_GotFocus()
    zlControl.TxtSelAll txtMemo
    zlCommFun.OpenIme True
End Sub
Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txtMemo_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblMoney As Double
    With vsBalance
        mblnChange = True
        Select Case Col
        Case .ColIndex("���")
            Call CalcTotal '���¼����ܽ��
        Case .ColIndex("�������")
        Case .ColIndex("���㷽ʽ")
            dblMoney = 0
            If Not mrsBalance Is Nothing Then
                mrsBalance.Filter = "���㷽ʽ='" & .TextMatrix(Row, Col) & "'"
                If Not mrsBalance.EOF Then
                    dblMoney = Val(NVL(mrsBalance!���))
                End If
            End If
            .TextMatrix(Row, .ColIndex("���")) = Format(dblMoney, "##0.00;-##0.00;0.00;")
            Call CalcTotal '���¼����ܽ��
        End Select
    End With
End Sub
Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
    zlCommFun.OpenIme False
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "���㷽ʽ�б�", False, zlStr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBalance
        Select Case Col
        Case .ColIndex("���"), .ColIndex("�������"), .ColIndex("���㷽ʽ")
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub
Private Sub vsBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsBalance
        If .Col = .Cols - 1 And .Row = .Rows - 1 _
            And Trim(.TextMatrix(.Row, .ColIndex("���㷽ʽ"))) = "" Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("���㷽ʽ"), vsBalance.Cols - 1, True)
End Sub
Private Sub vsBalance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("���㷽ʽ"), vsBalance.Cols - 1, True)
End Sub
Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub
Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsBalance
        If Row <= 1 Then Exit Sub
        Select Case Col
        Case .ColIndex("�������")
            VsFlxGridCheckKeyPress vsBalance, Row, Col, KeyAscii, m�ı�ʽ
            If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0: Exit Sub
        Case .ColIndex("���")
            VsFlxGridCheckKeyPress vsBalance, Row, Col, KeyAscii, m�����ʽ
        End Select
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
        Case .ColIndex("���")
            If Not IsNumeric(strKey) Then
                MsgBox "��������������,�������������ַ���", vbInformation, gstrSysName
                Cancel = True: Exit Sub
             End If
             If Val(strKey) > 999999999 Then
                MsgBox "����������,���ֻ������999999999��", vbInformation, gstrSysName
                Cancel = True: Exit Sub
             End If
             If Val(strKey) < -999999999 Then
                MsgBox "��������С,���ֻ������-999999999��", vbInformation, gstrSysName
                Cancel = True: Exit Sub
                Exit Sub
             End If
             .EditText = Format(strKey, "###0.00;-###0.00;0.00;0.00")
        End Select
    End With
End Sub



Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݺϷ��Լ��
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-11 16:35:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblMoney As Double, strTemp As String, j As Long
    Dim str���㷽ʽ As String
    On Error GoTo errHandle
    isValied = False
    '�����:110281,����,2017/08/15,������˵�������޴�50���ַ�����Ϊ500���ַ�
    If zlCommFun.ActualLen(txtMemo.Text) > 500 Then
        MsgBox "ժҪ�ĳ��Ȳ��ܳ���250�����ֻ�500���ַ���", vbInformation, gstrSysName
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    If InStr(txtMemo.Text, "'") > 0 Then
        MsgBox "ժҪ���зǷ��ַ���'����", vbInformation, gstrSysName
        zlControl.TxtSelAll txtMemo
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
'    If cboDept.ListIndex < 0 Then
'        MsgBox "δѡ��ɿ��!", vbInformation, gstrSysName
'        If cboDept.Visible And cboDept.Enabled Then cboDept.SetFocus
'        Exit Function
'    End If
    With vsBalance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
           If str���㷽ʽ <> "" Then
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
                strTemp = Trim(.TextMatrix(i, .ColIndex("���")))
                If Not IsNumeric(strTemp) Then
                   MsgBox "��������������,�������������ַ���", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("���")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                If Val(strTemp) > 999999999 Then
                   MsgBox "����������,���ֻ������999999999��", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("���")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                If Val(strTemp) < -999999999 Then
                   MsgBox "��������С,���ֻ������-999999999��", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("���")
                    If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                        .TopRow = .Row: .LeftCol = .Col
                    End If
                    If .Visible And .Enabled Then .SetFocus
                    Exit Function
                End If
                mrsBalance.Filter = "���㷽ʽ='" & str���㷽ʽ & "'"
                If mrsBalance.EOF Then
                    If MsgBox("�ɿ��˲�����" & str���㷽ʽ & "���ݴ��,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .Row = i: .Col = .ColIndex("���")
                        If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                            .TopRow = .Row: .LeftCol = .Col
                        End If
                        If .Visible And .Enabled Then .SetFocus
                        Exit Function
                    End If
                Else
                    '������Ƿ���ȷ
                    If Val(strTemp) > Val(NVL(mrsBalance!���)) Then
                        If MsgBox(str���㷽ʽ & "�Ľɿ���(" & Format(Val(strTemp), "0.00") & ")�����ݴ���(" & Format(Val(NVL(mrsBalance!���)), "0.00") & ")���Ƿ������", vbYesNo Or vbQuestion Or vbDefaultButton2, Me.Caption) = vbNo Then
                            .Row = i: .Col = .ColIndex("���")
                            If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                                .TopRow = .Row: .LeftCol = .Col
                            End If
                            If .Visible And .Enabled Then .SetFocus
                            Exit Function
                        End If
                    End If
                End If
                
                '�����㷽ʽ�Ƿ��ظ�
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("���㷽ʽ"))) = Trim(.TextMatrix(j, .ColIndex("���㷽ʽ"))) And i <> j Then
                        MsgBox "��" & i & "�����" & j & "�еĽ��㷽ʽ��ͬ,��ϲ���", vbInformation, gstrSysName
                         .Row = i: .Col = .ColIndex("���")
                         If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                             .TopRow = .Row: .LeftCol = .Col
                         End If
                         If .Visible And .Enabled Then .SetFocus
                         Exit Function
                    End If
                Next
            End If
        Next
    End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData(ByRef strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '����:strNo-���ݱ���ɹ���,���سɹ��ĵ��ݺ�
    '����:����ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-10-11 16:59:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngID As Long, strTemp As String, i As Long
    Dim str���㷽ʽ As String, str������ As String, str������� As String
 
    On Error GoTo errHandle
    With vsBalance
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("���㷽ʽ"))
            If strTemp <> "" Then
                str���㷽ʽ = str���㷽ʽ & "," & strTemp
                str������ = str������ & "," & Val(.TextMatrix(i, .ColIndex("���")))
                str������� = str������� & "," & Trim(.TextMatrix(i, .ColIndex("�������")))
            End If
        Next
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    If str������ <> "" Then str������ = Mid(str������, 2)
    If str������� <> "" Then str������� = Mid(str�������, 2)
    
    If str���㷽ʽ = "" Then
        MsgBox "�����ڽɿ�����,���������ɿ�����,���ܽ��������տ�", vbInformation + vbOKOnly, gstrSysName
        If vsBalance.Enabled And vsBalance.Visible Then vsBalance.SetFocus
        Exit Function
    End If
    
    If zlCommFun.ActualLen(str���㷽ʽ) > 4000 Then
        MsgBox "�ڽ�����ϸ��Ϣ������Ľ��㷽ʽ����,���ܽ����տ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If zlCommFun.ActualLen(str������) > 4000 Then
        MsgBox "�ڽ�����ϸ��Ϣ������Ľ��㷽ʽ����Ӧ�Ľ��������,���ܽ����տ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If zlCommFun.ActualLen(str�������) > 4000 Then
        MsgBox "�ڽ�����ϸ��Ϣ������Ľ��㷽ʽ����Ӧ�Ľ���������,���ܽ����տ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If str������� = "" Then str������� = " "
    lngID = zlDatabase.GetNextId("��Ա�սɼ�¼")
    strNO = zlDatabase.GetNextNo(140)
    'Zl_�ֹ��տ��¼_Insert
    strSQL = "Zl_�ֹ��տ��¼_Insert("
    '  Id_In         In ��Ա�սɼ�¼.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In         In ��Ա�սɼ�¼.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  �տ�Ա_In     In ��Ա�սɼ�¼.�տ�Ա%Type,
    strSQL = strSQL & "'" & mstr�ɿ��� & "',"
    '  �տ��id_In In ��Ա�սɼ�¼.�տ��id%Type,
    strSQL = strSQL & "" & "Null,"
    'strSQL = strSQL & "" & cboDept.ItemData(cboDept.ListIndex) & ","
    '  �տ�ʱ��_In   In ��Ա�սɼ�¼.��ʼʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(dtpDate.Value, "yyyy-mm-dd") & "','yyyy-mm-dd'),"
    '  ժҪ_In       In ��Ա�սɼ�¼.ժҪ%Type,
    strSQL = strSQL & IIf(Trim(txtMemo.Text) = "", "NULL", "'" & txtMemo.Text & "'") & ","
    '  �Ǽ���_In     In ��Ա�սɼ�¼.�Ǽ���%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ǽ�ʱ��_In   In ��Ա�սɼ�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "Sysdate,"
    ' �սɱ�־_In   In ��Ա�սɼ�¼.�սɱ�־%Type,
    strSQL = strSQL & IIf(mblnOtherPerson, 1, "NULL") & ","
    '  ���㷽ʽ_In   Varchar2,���㷽ʽ_IN:������,���ʱ,�ö��ŷ���,����:�ֽ�,֧Ʊ,...
    '       ���㷽ʽ_In,������_In,�������_IN ����������ֵ�ĸ���Ҫһһ��Ӧ:����:���㷽ʽ_IN (�ֽ�,֧Ʊ...),������(100,0...),�������_IN(A001,A002,...)
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  ������_In   Varchar2,������_IN:������,���ʱ,�ö��ŷ���,����㷽ʽ_IN һһ��Ӧ.
    strSQL = strSQL & "'" & str������ & "',"
    '  �������_In   In Varchar2,�������_In:������,���ʱ,�ö��ŷ���,����㷽ʽ_IN һһ��Ӧ
    strSQL = strSQL & "'" & str������� & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
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
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500", Me, "NO=" & strNO, "��¼����=5", 2)
End Sub
