VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm������ҩ 
   Caption         =   "������ҩ���Ĵ���"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   Icon            =   "frm������ҩ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   7725
   StartUpPosition =   1  '����������
   Begin VB.ComboBox Cbo���� 
      Height          =   300
      Left            =   1095
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1290
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtInput 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   165
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "####"
      Top             =   1320
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picFunc 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   75
      ScaleHeight     =   600
      ScaleWidth      =   6765
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4140
      Width           =   6765
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   2670
         TabIndex        =   11
         ToolTipText     =   "ɾ����ǰѡ����"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "��ҩ(&O)"
         Height          =   350
         Left            =   4275
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5490
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "ȫ��(&E)"
         Height          =   350
         Left            =   1440
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   165
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.TextBox TxtNo 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   165
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Bill 
      Height          =   2985
      Left            =   90
      TabIndex        =   3
      Top             =   1020
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5265
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "������ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   135
      TabIndex        =   2
      Top             =   600
      Width           =   7395
   End
   Begin VB.Label LblNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   255
      TabIndex        =   0
      Top             =   225
      Width           =   540
   End
End
Attribute VB_Name = "frm������ҩ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strUnit As String
Private strPrivs As String
Private mblnRefresh As Boolean
Private mlngҩ��ID As Long
Private mbln������ҩ As Boolean
Private mstr�۸�ʧЧ��ʾ As String
Private mint����λ�� As Integer                      '���ý���λ��
Private Enum ����
    NO = 0
    ����
    ���
    �ⷿ
    Id
    ҩƷID
    ����
    ����
    ����
    ����
    ҩƷ����
    ��Ʒ��
    ���
    ����
    Ч��
    ����
    ����
    ����
    ������
    ׼����
    ��ҩ��
    ��λ
    ����
    ʵ������
    ���
    ��¼����
    �����־
End Enum
Private Const int����Last = 27

Private rs��� As New ADODB.Recordset
Private mrs��ҩ As New ADODB.Recordset
Private mobjPlugIn As Object             '��ҽӿڶ���

Public Property Get In_Ȩ��() As String
    In_Ȩ�� = strPrivs
End Property

Public Property Let In_Ȩ��(ByVal vNewValue As String)
    strPrivs = vNewValue
End Property

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property

Public Function ShowEditor(ByVal frmParent As Object, ByVal lngҩ��ID As Long, _
    Optional ByVal bln������ҩ As Boolean = True, _
    Optional ByVal int����λ�� As Integer = 2, Optional ByVal strPriv As String) As Boolean
    mblnRefresh = False
    mlngҩ��ID = lngҩ��ID
    mbln������ҩ = bln������ҩ
    mint����λ�� = int����λ��
    strPrivs = strPriv
    Me.Show 1, frmParent
    ShowEditor = mblnRefresh
End Function

Private Sub Bill_DblClick()
    '��ʾ��ҩ���ı���ȱʡΪ��ǰ��λ�����ݣ������û��޸ġ�
    '�������ֵ�Ƿ����㡢�ո񡢷Ƿ���������ȫ��������������ȱʡΪȫ��
    With Bill
        .Col = ����.��ҩ��
        If Val(.TextMatrix(Bill.Row, ����.Id)) = 0 Then Exit Sub
        TxtInput.Tag = Val(.TextMatrix(Bill.Row, ����.׼����))
        TxtInput.Text = Format(Val(Bill.TextMatrix(Bill.Row, ����.��ҩ��)), "#####0.00000;-#####0.00000; ;")
        Call ShowTxt
    End With
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Call Bill_DblClick
End Sub

Private Sub Bill_Scroll()
    Dim blnCancel As Boolean
    Call TxtInput_Validate(blnCancel)
    Bill.Row = Bill.TopRow
    Call Bill_EnterCell
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    Dim lngRow As Long, lngRows As Long
    '����ҩ����Ϊ��
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        Bill.TextMatrix(lngRow, ����.��ҩ��) = ""
    Next
End Sub

Private Sub cmdDelete_Click()
    Dim lngCol As Long, lngCols As Long
    If Bill.Row = Bill.rows - 1 And Bill.Row = 1 Then
        lngCols = Bill.Cols - 1
        For lngCol = 0 To lngCols
            Bill.TextMatrix(1, lngCol) = ""
        Next
    Else
        Bill.RemoveItem Bill.Row
        Call Bill_EnterCell
    End If
End Sub

Private Sub cmdOk_Click()
    Dim blnInput As Boolean
    Dim dbl��ҩ�� As Double
    Dim StrDate As String
    Dim strShow As String, strReturn As String, strSubSql As String
    Dim str���� As String, strЧ�� As String, str���� As String
    Dim lng���� As Long, lng���� As Long, lngRow As Long, lngRows As Long
    Dim RecRecord As New ADODB.Recordset
    Dim bln�Ƿ�����ҩ As Boolean
    Dim strҩƷid As String
    Dim blnIsReturn As Boolean
    Dim arrSql As Variant
    Dim i As Long
    Dim blnBeginTrans As Boolean
    Dim int���� As Integer
    Dim Int��ҩ As Integer
    Dim strReturnInfo As String
    Dim strReserve As String
    
    arrSql = Array()
    
    On Error GoTo ErrHand
    
    'ִ��Ԥ����
    Call setNOtExcetePrice
    
    '����Ƿ��������
    lngRow = Bill.rows - 1 - IIf(Val(Bill.TextMatrix(Bill.rows - 1, ����.Id)) = 0, 1, 0)
    If Val(Bill.TextMatrix(lngRow, ����.ҩƷID)) = 0 Then Exit Sub
    Call BuildRecord
    If Not CheckCorrelation Then Exit Sub
    If Not CheckBillOperate Then Exit Sub
    
    '��ʾ
    If MsgBox("��ȷ��Ҫ��ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    StrDate = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(mlngҩ��ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(mlngҩ��ID, gint����ҩ��)
    Else
        strUnit = GetSpecUnit(mlngҩ��ID, gintסԺҩ��)
    End If
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "*1"
    Case "���ﵥλ"
        strSubSql = "*Decode(�����װ,Null,1,0,1,�����װ)"
    Case "סԺ��λ"
        strSubSql = "*Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
    Case "ҩ�ⵥλ"
        strSubSql = "*Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
    End Select
    
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        If Val(Bill.TextMatrix(lngRow, ����.��ҩ��)) <> 0 Then
            lng���� = Val(Bill.TextMatrix(lngRow, ����.����))
            lng���� = Val(Bill.TextMatrix(lngRow, ����.����))
            '���ԭ�������������ڷ���
            If lng���� = 0 And lng���� = 1 Then
                '������Ż�Ч��Ϊ�գ�����ȡ���û�����
                blnInput = (Trim(Bill.TextMatrix(lngRow, ����.����)) = "")
                If blnInput Then
                    strShow = Bill.TextMatrix(lngRow, ����.����) & "||" & _
                    Bill.TextMatrix(lngRow, ����.����) & "|" & Bill.TextMatrix(lngRow, ����.ҩƷ����) & "|" & _
                    Val(Bill.TextMatrix(lngRow, ����.ҩƷID))
                    strReturn = Frm��ҩ����.ShowMe(Me, strShow)
                    If strReturn = "" Then Exit Sub
                    '�������š�Ч�ڼ�����
                    Bill.TextMatrix(lngRow, ����.����) = Split(strReturn, "|")(0)
                    Bill.TextMatrix(lngRow, ����.Ч��) = Split(strReturn, "|")(1)
                    Bill.TextMatrix(lngRow, ����.����) = Split(strReturn, "|")(2)
                End If
            End If
        End If
    Next
    
    bln�Ƿ�����ҩ = False

    Call BuildRecordReturn
    If mrs��ҩ.RecordCount <> 0 Then mrs��ҩ.MoveFirst
    mrs��ҩ.Sort = "ҩƷID"
    Do While Not mrs��ҩ.EOF
        dbl��ҩ�� = Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.��ҩ��))
        If dbl��ҩ�� <> 0 Then
            If CheckBill(2, Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.Id)), Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.��¼����)), Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.�����־))) <> 0 Then
                Exit Sub
            End If

            gstrSQL = " Select round(" & dbl��ҩ�� & strSubSql & ",5) ���� From ҩƷĿ¼" & _
                         " Where ҩƷID=[1]"
            Set RecRecord = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.ҩƷID)))
            
            With RecRecord
                dbl��ҩ�� = zlStr.Nvl(!����, 0)
            End With
            If Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.��ҩ��)) = Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.׼����)) Then
                dbl��ҩ�� = Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.ʵ������))
            End If
            str���� = Bill.TextMatrix(mrs��ҩ!�к�, ����.����)
            strЧ�� = Bill.TextMatrix(mrs��ҩ!�к�, ����.Ч��)
            str���� = Bill.TextMatrix(mrs��ҩ!�к�, ����.����)
            
            blnIsReturn = False
            If CheckPrice(Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.Id)), mstr�۸�ʧЧ��ʾ) = False Then
                If MsgBox("ҩƷ[" & Bill.TextMatrix(mrs��ҩ!�к�, ����.ҩƷ����) & "]" & mstr�۸�ʧЧ��ʾ, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnIsReturn = True
                End If
            Else
                blnIsReturn = True
            End If
            
            If blnIsReturn = True Then
                
                If Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.��¼����)) = 1 Or (Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.��¼����)) = 2 And (Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.�����־))) = 1 Or (Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.�����־))) = 4) Then
                    int���� = 1
                Else
                    int���� = 2
                End If
                
                gstrSQL = "zl_ҩƷ�շ���¼_������ҩ("
                '�շ�ID
                gstrSQL = gstrSQL & Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.Id))
                '�����
                gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
                '�������
                gstrSQL = gstrSQL & ",To_Date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')"
                '����
                gstrSQL = gstrSQL & "," & IIf(str���� = "", "NULL", IIf(Mid(str����, 1, 1) = "(", "NULL", "'" & Mid(str����, 1, 8) & "'"))
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(strЧ�� = "", "NULL", "To_Date('" & Format(strЧ��, "yyyy-MM-dd") & "','yyyy-MM-dd')")
                '����
                gstrSQL = gstrSQL & "," & IIf(str���� = "", "NULL", "'" & str���� & "'")
                '��ҩ��
                gstrSQL = gstrSQL & "," & dbl��ҩ��
                '��ҩ�ⷿ
                gstrSQL = gstrSQL & "," & mlngҩ��ID
                '��ҩ��
                gstrSQL = gstrSQL & ",NULL"
                '����λ��
                gstrSQL = gstrSQL & "," & mint����λ��
                '�����־
                gstrSQL = gstrSQL & "," & int����
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL

                bln�Ƿ�����ҩ = True
                
                If InStr("," & strҩƷid & ",", "," & Bill.TextMatrix(mrs��ҩ!�к�, ����.ҩƷID) & ",") = 0 Then
                    strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & Bill.TextMatrix(mrs��ҩ!�к�, ����.ҩƷID)
                End If
                
                strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & Val(Bill.TextMatrix(mrs��ҩ!�к�, ����.Id)) & "," & dbl��ҩ��
            End If
        End If
        
        mrs��ҩ.MoveNext
    Loop
    
    '��ʾͣ��ҩƷ
    If strҩƷid <> "" Then
        Int��ҩ = 1
        Call CheckStopMedi(strҩƷid, Int��ҩ)
        If Int��ҩ = 2 Then Exit Sub
    End If
    
    If bln�Ƿ�����ҩ = True Then
        '���д�����ҩ����
        gcnOracle.BeginTrans
        blnBeginTrans = True
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-ҩƷ��ҩ")
        Next
        gcnOracle.CommitTrans
        blnBeginTrans = False
    
        If MsgBox("����Ҫ��ӡ��ҩ�嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If mbln������ҩ Then
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_5", "ZL8_BILL_1341_5"), Me, "��ҩʱ��=" & StrDate, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
            Else
                Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1342_1", "ZL8_BILL_1342_1"), Me, "��ҩʱ��=" & StrDate, "��װϵ��=" & IIf(strUnit = "���ﵥλ", "C.�����װ", "C.סԺ��װ"), 2)
            End If
        End If
    Else
        MsgBox "����û����ҩ��"
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ�����ҩ Then
        On Error Resume Next
        mobjPlugIn.DrugReturnByID mlngҩ��ID, strReturnInfo, CDate(StrDate), strReserve
        err.Clear: On Error GoTo 0
    End If
    
    'ˢ��
    mblnRefresh = True
    Call SetFormat
    TxtNo.SetFocus
    Exit Sub
ErrHand:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    
    Call SaveErrLog
End Sub

Private Sub BuildRecordReturn()
    '��ҩ���ݼ�
    Dim intRow As Integer, intRows As Integer
        
    Call InitRecReturn
    '���ݴ���ҩ�嵥��������ҩ���ݼ�
    intRows = Bill.rows - 1
    For intRow = 1 To intRows
        If Val(Bill.TextMatrix(intRow, ����.Id)) <> 0 Then
           If Val(Bill.TextMatrix(intRow, ����.��ҩ��)) <> 0 Then
                mrs��ҩ.AddNew
                mrs��ҩ!�к� = intRow
                mrs��ҩ!ҩƷID = Val(Bill.TextMatrix(intRow, ����.ҩƷID))
                mrs��ҩ.Update
            End If
        End If
    Next
End Sub

Private Sub InitRecReturn()
    '��ҩ���ݼ���������������ҩʱ��ҩƷID����
    Set mrs��ҩ = New ADODB.Recordset
    With mrs��ҩ
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷid", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
Private Sub cmdSelAll_Click()
    Dim lngRow As Long, lngRows As Long
    '����ҩ����Ϊ׼����
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        Bill.TextMatrix(lngRow, ����.��ҩ��) = Bill.TextMatrix(lngRow, ����.׼����)
    Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call SetFormat
End Sub

Private Function CheckBillOperate() As Boolean
    Dim n As Integer
    
    For n = 1 To Bill.rows - 1
        If Bill.TextMatrix(n, 1) <> "" Then
            If CheckBillControl(4, Val(Bill.TextMatrix(n, ����.����)), Bill.TextMatrix(n, ����.NO), Val(Bill.TextMatrix(n, ����.���))) = False Then
                Exit Function
            End If
        End If
    Next
    
    CheckBillOperate = True
End Function
Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With lblTitle
        .Top = TxtNo.Top + TxtNo.Height + 80
        .Width = Me.ScaleWidth - .Left - 100
    End With
    
    With picFunc
        .Left = lblTitle.Left
        .Width = lblTitle.Width
        .Top = Me.ScaleHeight - .Height
    End With
    
    With Bill
        .Left = lblTitle.Left
        .Top = lblTitle.Top + lblTitle.Height
        .Width = lblTitle.Width
        .Height = Me.ScaleHeight - picFunc.Height - .Top
    End With
    
    With cmdCancel
        .Left = picFunc.Width - .Width - 100
    End With
    With cmdOK
        .Left = cmdCancel.Left - .Width - 100
    End With
    Call Bill_Scroll
End Sub

Private Sub TxtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long, lngNewRow As Long
    Dim blnCancel As Boolean
    lngRow = Bill.Row
    lngNewRow = lngRow                  'ȱʡΪ��ǰ��
    
    Select Case KeyCode
    Case vbKeyUp
        If Bill.Row > 1 Then lngNewRow = Bill.Row - 1
    Case vbKeyDown, vbKeyReturn
        If Bill.Row = Bill.rows - 1 Then
            Call TxtInput_Validate(blnCancel)
            cmdDelete.SetFocus
        ElseIf Bill.Row < Bill.rows - 1 Then
            lngNewRow = Bill.Row + 1
        End If
    Case Else
        Exit Sub
    End Select
    
    KeyCode = 0
    If lngRow <> lngNewRow Then
        Call TxtInput_Validate(blnCancel)
        Bill.Row = lngNewRow
        Call Bill_EnterCell
    End If
End Sub
Private Sub TxtInput_Validate(Cancel As Boolean)
    Dim blnUnValid As Boolean, dblCount As Double
    Dim rstemp As New ADODB.Recordset
    On Error Resume Next
    
    If Not TxtInput.Visible Then Exit Sub
    blnUnValid = False
    TxtInput = Trim(TxtInput)
    
    blnUnValid = (TxtInput = "")
    If Not blnUnValid Then blnUnValid = Not IsNumeric(TxtInput)
    If Not blnUnValid Then blnUnValid = Not ((Abs(TxtInput) <= Abs(TxtInput.Tag)) And ((Val(TxtInput) >= 0 And Val(TxtInput.Tag) >= 0) Or (Val(TxtInput) <= 0 And Val(TxtInput.Tag) <= 0)))
    If blnUnValid Then
        If TxtInput = "" Then
            TxtInput = 0
        Else
            TxtInput = Val(TxtInput.Tag)
        End If
    End If
    
    Bill.TextMatrix(Bill.Row, ����.��ҩ��) = Format(Val(TxtInput.Text), "#####0.00000;-#####0.00000; ;")
    TxtInput.Visible = False
End Sub

Private Sub TxtNo_GotFocus()
    Call zlControl.TxtSelAll(TxtNo)
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String, str��λ As String, str��װ As String
    Dim intYear As Integer, strYear As String
    Dim rsBill As New ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    '���ݸ�Ʊ�ݺŶ�����������ҩ
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtNo.Text) = "" Then Exit Sub

    '--���������λ,�򰴹������--
    Me.TxtNo.Text = Trim(UCase(TxtNo.Text))
    Me.TxtNo.Text = GetFullNO(Me.TxtNo.Text, 13)
    strInput = Me.TxtNo.Text
    
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("ҩ������", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(mlngҩ��ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(mlngҩ��ID, gint����ҩ��)
    Else
        strUnit = GetSpecUnit(mlngҩ��ID, gintסԺҩ��)
    End If
    Select Case strUnit
    Case "�ۼ۵�λ"
        str��λ = "X.���㵥λ"
        str��װ = "1"
    Case "���ﵥλ"
        str��λ = "D.���ﵥλ"
        str��װ = "D.�����װ"
    Case "סԺ��λ"
        str��λ = "D.סԺ��λ"
        str��װ = "D.סԺ��װ"
    Case "ҩ�ⵥλ"
        str��λ = "D.ҩ�ⵥλ"
        str��װ = "D.ҩ���װ"
    End Select
    
    '�������Ŷ�ȡ��Ӧ���ѷ�ҩ��¼
    gstrSQL = "" & _
            " SELECT DISTINCT S.ID,S.����,S.ҩƷID,S.NO,S.���,S.����,F.���� AS �ⷿ,P.���� ����,C.��¼����,C.�����־,'' ����,C.����,'['||X.����||']' As ҩƷ����,X.���� As ͨ����,A.���� As ��Ʒ��," & _
            " NVL(D.ҩ������,0) ����,DECODE(X.���,NULL,S.����,DECODE(S.����,NULL,X.���,X.���||'|'||S.����)) ���," & str��λ & " ��λ," & str��װ & " ��װ," & _
            " S.���� ��,S.ʵ������ ����,S.��������,S.�ѷ����� ׼����,DECODE(S.����,NULL,'',S.����)||DECODE(S.����,NULL,'',0,'','('||S.����||')') ����," & _
            " NVL(S.����,0) ����,S.Ч��, S.���ۼ� ����,S.���۽�� ���,S.����,S.Ƶ��,S.�÷�,S.ժҪ ˵��,S.�����,TO_CHAR(S.�������,'YYYY-MM-DD HH24:MI:SS') ��ҩʱ��,1 �ɲ���" & _
            " FROM" & _
            "     (SELECT A.ID,A.NO,A.����,A.���,A.ҩƷID,A.����ID,A.����,A.����,A.����,A.Ч��,NVL(A.����,0) ����," & _
            "     NVL(A.����,1) ����,A.ʵ������ ʵ������,NVL(A.����,1)*A.ʵ������-B.�ѷ����� ��������,B.�ѷ�����,A.��¼״̬," & _
            "     A.���ۼ� , A.���۽��, A.����, A.Ƶ��, A.�÷�, A.ժҪ, A.�����, A.�������, A.�Է�����ID, A.�ⷿID" & _
            "     FROM ҩƷ�շ���¼ A," & _
            "         (SELECT A.NO,A.����,A.ҩƷID,A.���,SUM(NVL(A.����,1)*A.ʵ������) �ѷ�����" & _
            "         FROM ҩƷ�շ���¼ A" & _
            "         WHERE A.����� IS NOT NULL And A.�ⷿID+0<>[2] " & _
            "         AND NO =[1] And ���� In (8,9)" & _
            "         GROUP BY A.NO,A.����,A.ҩƷID,A.���) B" & _
            "     WHERE A.NO = B.NO AND A.���� = B.���� AND A.ҩƷID+0 = B.ҩƷID AND A.��� = B.��� AND B.�ѷ�����<>0 And A.����� IS NOT NULL AND (A.��¼״̬=1 OR MOD(A.��¼״̬,3)=0)) S," & _
            "     ������ü�¼ C,���ű� P,ҩƷ��� D,�շ���ĿĿ¼ X,�շ���Ŀ���� A,���ű� F"
    gstrSQL = gstrSQL & _
            " WHERE S.ҩƷID=D.ҩƷID AND D.ҩƷID=X.ID AND S.�Է�����ID+0=P.ID AND S.����ID=C.ID And S.�ⷿID+0=F.ID" & _
            " AND D.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 " & _
            " AND (S.��¼״̬=1 OR MOD(S.��¼״̬,3)=0) AND S.����� IS NOT NULL AND S.�ⷿID+0<>[2] " & _
            " AND S.ʵ������*S.����>S.�������� "
            
    strsql = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    strsql = Replace(strsql, "'' ����", "C.����")
    gstrSQL = gstrSQL & " Union All " & strsql
    gstrSQL = gstrSQL & " ORDER BY ����,NO"
    
    Set rsBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ�ô����Ŷ�Ӧ���ѷ�ҩ��¼]", strInput, mlngҩ��ID)
            
    If rsBill.RecordCount = 0 Then
        MsgBox "�ô����Ŷ�Ӧ�Ĵ�����δ��ҩ����ȫ����ҩ����ת���������ݿ⣡", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(TxtNo)
        Exit Sub
    End If
    
    Call WriteBill(rsBill)
    Call zlControl.TxtSelAll(TxtNo)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetBillID() As String
    Dim lngRow As Long, lngRows As Long
    Dim strReturn As String
    '�����Ѵ��ڵĴ�����ϸID���Ա���飬���������ͬ�ģ��򲻼���
    lngRows = Bill.rows - 1
    For lngRow = 1 To lngRows
        If Val(Bill.TextMatrix(lngRow, ����.Id)) <> 0 Then
            strReturn = strReturn & "," & Bill.TextMatrix(lngRow, ����.Id)
        End If
    Next
    If strReturn = "" Then Exit Function
    GetBillID = strReturn & ","
End Function

Private Sub WriteBill(ByVal rsBill As ADODB.Recordset)
    Dim lngRow As Long
    Dim strID As String
    '�����ʼ��
    lngRow = Bill.rows - 1 + IIf(Val(Bill.TextMatrix(Bill.rows - 1, ����.Id)) = 0, 0, 1)
    If Bill.rows - 1 < lngRow Then Bill.rows = Bill.rows + 1
    
    '�����Ѵ��ڵĴ�����ϸID���Ա���飬���������ͬ�ģ��򲻼���
    strID = GetBillID
    
    '��ҩƷ��ϸд���ѷ�ҩ�嵥
    With rsBill
        Do While Not .EOF
            '��ǰû����ļ�¼��д���ѷ�ҩ�嵥�У����û���ҩ
            If InStr(1, strID, "," & !Id & ",") = 0 Then
                Bill.TextMatrix(lngRow, ����.�ⷿ) = !�ⷿ
                Bill.TextMatrix(lngRow, ����.NO) = !NO
                Bill.TextMatrix(lngRow, ����.����) = !����
                Bill.TextMatrix(lngRow, ����.���) = !���
                Bill.TextMatrix(lngRow, ����.Id) = !Id
                Bill.TextMatrix(lngRow, ����.ҩƷID) = !ҩƷID
                Bill.TextMatrix(lngRow, ����.����) = !����
                Bill.TextMatrix(lngRow, ����.����) = !����
                Bill.TextMatrix(lngRow, ����.����) = !����
                Bill.TextMatrix(lngRow, ����.����) = !����
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    Bill.TextMatrix(lngRow, ����.ҩƷ����) = !ҩƷ���� & !ͨ����
                Else
                    Bill.TextMatrix(lngRow, ����.ҩƷ����) = !ҩƷ���� & IIf(IsNull(!��Ʒ��), !ͨ����, !��Ʒ��)
                End If
                
                Bill.TextMatrix(lngRow, ����.��Ʒ��) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
                Bill.TextMatrix(lngRow, ����.���) = IIf(IsNull(!���), "", !���)
                Bill.TextMatrix(lngRow, ����.����) = IIf(IsNull(!����), "", !����)
                Bill.TextMatrix(lngRow, ����.Ч��) = IIf(IsNull(!Ч��), "", !Ч��)
                Bill.TextMatrix(lngRow, ����.����) = ""
                Bill.TextMatrix(lngRow, ����.����) = Format(!��, "#####0;-#####0; ;")
                Bill.TextMatrix(lngRow, ����.����) = Format(!���� / !��װ, "#####0.00000;-#####0.00000; ;")
                Bill.TextMatrix(lngRow, ����.������) = Format(!�������� / !��װ, "#####0.00000;-#####0.00000; ;")
                Bill.TextMatrix(lngRow, ����.׼����) = Format(!׼���� / !��װ, "#####0.00000;-#####0.00000; ;")
                Bill.TextMatrix(lngRow, ����.��ҩ��) = Format(!׼���� / !��װ, "#####0.00000;-#####0.00000; ;")
                Bill.TextMatrix(lngRow, ����.��λ) = IIf(IsNull(!��λ), "", !��λ)
                Bill.TextMatrix(lngRow, ����.����) = Format(!���� * !��װ, "#####0.00000;-#####0.00000; ;")
                Bill.TextMatrix(lngRow, ����.ʵ������) = !׼����
                Bill.TextMatrix(lngRow, ����.���) = !���
                Bill.TextMatrix(lngRow, ����.��¼����) = !��¼����
                Bill.TextMatrix(lngRow, ����.�����־) = !�����־
                If lngRow >= Bill.rows - 1 Then
                    lngRow = lngRow + 1
                    Bill.rows = Bill.rows + 1
                End If
            End If
            .MoveNext
        Loop
        
        'ɾ�����Ŀհ���
        If Val(Bill.TextMatrix(Bill.rows - 1, ����.Id)) = 0 Then
            Bill.rows = Bill.rows - 1
        End If
    End With
End Sub

Private Sub Bill_EnterCell()
    Dim blnCancel As Boolean
    If TxtInput.Visible Then
        Call TxtInput_Validate(blnCancel)
        TxtInput.Visible = False
    End If
End Sub

Private Sub Bill_GotFocus()
    Bill_EnterCell
End Sub

Private Sub ShowTxt(Optional ByVal ���뷽ʽ As Integer = 1)
    '0-�����;1-�Ҷ���;2-���ж���
    On Error Resume Next
    With TxtInput
        .Alignment = ���뷽ʽ
        .Left = Bill.Left + Bill.CellLeft
        .Top = Bill.Top + Bill.CellTop
        .Width = Bill.CellWidth - 20
        .Visible = True
        .ZOrder 0
        .SetFocus
    End With
    Call zlControl.TxtSelAll(TxtInput)
End Sub

Private Function CheckBill(ByVal IntOper As Integer, ByVal LngID As Long, ByVal int��¼���� As Integer, ByVal int�����־ As Integer) As Integer
    Dim RecCheck As New ADODB.Recordset
    
    '--���ݽ�Ҫִ�еĲ������ж��Ƿ�����--
    '0-�ܷ�;1-��ҩ;2-��ҩ
    '����:
    '0-�������
    '1-�ѷ�ҩ
    '2-��ɾ��
    '3-δ��ҩ
    On Error GoTo errHandle
    gstrSQL = " Select A.�����,Decode(Nvl(A.ժҪ,'С��'),'�ܷ�',3,B.ִ��״̬) ִ��״̬ From ҩƷ�շ���¼ A,������ü�¼ B " & _
             " Where A.����ID=B.ID And A.ID=[1]"
    If IntOper = 2 Then
        gstrSQL = gstrSQL & " And ����� IS Not Null"
    Else
        gstrSQL = gstrSQL & " And ����� IS Null"
    End If
    
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngID)
    
    With RecCheck
        If .EOF Then CheckBill = 2: MsgBox "δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!�����) Then
            If IntOper <> 2 Then CheckBill = 1: MsgBox "�ô����ѱ���������Ա��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
        Else
            If IntOper = 2 Then CheckBill = 3: MsgBox "�ô�����δ��ҩ������������ֹ��", vbInformation, gstrSysName: Exit Function
        End If
        If IntOper = 1 And !ִ��״̬ = 3 Then CheckBill = 2: MsgBox "�ô����Ѿܷ�������������ֹ��", vbInformation, gstrSysName: Exit Function
    End With
    
    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetFormat()
Dim intCol As Integer
    '���ñ��
    With Bill
        .rows = 2
        .Cols = int����Last
        .Clear
        
        .TextMatrix(0, ����.�ⷿ) = "�ⷿ"
        .TextMatrix(0, ����.NO) = "NO"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.���) = "���"
        .TextMatrix(0, ����.Id) = "ID"
        .TextMatrix(0, ����.ҩƷID) = "ҩƷID"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, ����.��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, ����.���) = "���"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.Ч��) = "Ч��"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.����) = "��"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.������) = "������"
        .TextMatrix(0, ����.׼����) = "׼����"
        .TextMatrix(0, ����.��ҩ��) = "��ҩ��"
        .TextMatrix(0, ����.��λ) = "��λ"
        .TextMatrix(0, ����.����) = "����"
        .TextMatrix(0, ����.ʵ������) = "ʵ������"
        .TextMatrix(0, ����.���) = "���"
        .TextMatrix(0, ����.��¼����) = "��¼����"
        .TextMatrix(0, ����.�����־) = "�����־"
        
        .ColWidth(����.�ⷿ) = 1200
        .ColWidth(����.����) = 0
        .ColWidth(����.���) = 0
        .ColWidth(����.NO) = 900
        .ColWidth(����.Id) = 0
        .ColWidth(����.ҩƷID) = 0
        .ColWidth(����.����) = 0
        .ColWidth(����.����) = 0
        .ColWidth(����.����) = 0
        .ColWidth(����.����) = 0
        .ColWidth(����.ҩƷ����) = 2000
        .ColWidth(����.��Ʒ��) = IIf(gintҩƷ������ʾ = 2, 2000, 0)
        .ColWidth(����.���) = 1500
        .ColWidth(����.����) = 1500
        .ColWidth(����.Ч��) = 0
        .ColWidth(����.����) = 0
        .ColWidth(����.����) = 300
        .ColWidth(����.����) = 1000
        .ColWidth(����.������) = 1000
        .ColWidth(����.׼����) = 1000
        .ColWidth(����.��ҩ��) = 1000
        .ColWidth(����.��λ) = 600
        .ColWidth(����.����) = 1000
        .ColWidth(����.ʵ������) = 0
        .ColWidth(����.���) = 0
        .ColWidth(����.��¼����) = 0
        .ColWidth(����.�����־) = 0
    
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        .ColAlignment(����.���) = 1
        .ColAlignment(����.����) = 1
        .ColAlignment(����.������) = 7
        .ColAlignment(����.׼����) = 7
        .ColAlignment(����.��ҩ��) = 7
    End With
End Sub

Private Sub BuildRecord()
    Dim intRow As Integer, intRows As Integer
    Dim strNo As String, lng���� As Long, str��� As String
    
    Call InitRec
    '���ݴ���ҩ�嵥�������ݻ�ȡ��ϸ���
    intRows = Bill.rows - 1
    For intRow = 1 To intRows
        If Val(Bill.TextMatrix(intRow, ����.Id)) <> 0 Then
            strNo = Bill.TextMatrix(intRow, ����.NO)
            lng���� = Val(Bill.TextMatrix(intRow, ����.����))
            If Val(Bill.TextMatrix(intRow, ����.��ҩ��)) <> 0 Then
                If rs���.RecordCount <> 0 Then rs���.MoveFirst
                rs���.Find "���ݱ�ʶ='" & strNo & "|" & lng���� & "'"
                If rs���.EOF Then rs���.AddNew
                rs���!���ݱ�ʶ = strNo & "|" & lng����
                rs���!��¼���� = Val(Bill.TextMatrix(intRow, ����.��¼����))
                rs���!�����־ = Val(Bill.TextMatrix(intRow, ����.�����־))
                str��� = zlStr.Nvl(rs���!���)
                If InStr(1, "," & str��� & ",", "," & Val(Bill.TextMatrix(intRow, ����.���)) & ",") = 0 Then
                    If str��� = "" Then
                        str��� = Val(Bill.TextMatrix(intRow, ����.���))
                    Else
                        str��� = str��� & "," & Val(Bill.TextMatrix(intRow, ����.���))
                    End If
                    rs���!��� = str���
                End If
                rs���.Update
            End If
        End If
    Next
End Sub

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng���� As Long, str��� As String
    '��鴦���Ƿ��ѽ��ʡ����ò����Ƿ��ѳ�Ժ������Ȩ�޽��м��
    With rs���
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !���ݱ�ʶ
            lng���� = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            str��� = zlStr.Nvl(!���)
            If Not IsReceiptBalance_Charge(1, strPrivs, lng����, strNo, str���, Val(!��¼����), Val(!�����־)) Then Exit Function
            If Not IsOutPatient(strPrivs, lng����, strNo, Val(!��¼����), Val(!�����־)) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set rs��� = New ADODB.Recordset
    With rs���
        If .State = 1 Then .Close
        .Fields.Append "���ݱ�ʶ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
