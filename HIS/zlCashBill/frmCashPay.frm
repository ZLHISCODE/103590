VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmCashPay 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ɿ��¼"
   ClientHeight    =   6180
   ClientLeft      =   435
   ClientTop       =   720
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashPay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cbo�ɿ�� 
      Height          =   360
      Left            =   2220
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4800
      Width           =   2010
   End
   Begin VB.TextBox txtPay 
      Enabled         =   0   'False
      Height          =   360
      Left            =   930
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1245
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   5070
      TabIndex        =   14
      Top             =   5520
      Width           =   1200
   End
   Begin ZL9BillEdit.BillEdit msh 
      Height          =   2205
      Left            =   930
      TabIndex        =   1
      Top             =   1410
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   3889
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   930
      TabIndex        =   18
      Top             =   915
      Width           =   1665
      _ExtentX        =   2937
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
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   190906371
      CurrentDate     =   36904
   End
   Begin VB.TextBox txtSum 
      Enabled         =   0   'False
      Height          =   360
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3780
      Width           =   5325
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -210
      TabIndex        =   11
      Top             =   5250
      Width           =   7125
   End
   Begin VB.TextBox txtDigest 
      Height          =   360
      Left            =   930
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4290
      Width           =   5325
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ����(&S)"
      Height          =   420
      Left            =   210
      TabIndex        =   15
      Top             =   5520
      Width           =   1530
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   3720
      TabIndex        =   13
      Top             =   5520
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   2370
      TabIndex        =   12
      Top             =   5520
      Width           =   1200
   End
   Begin VB.TextBox txtHandle 
      Enabled         =   0   'False
      Height          =   360
      Left            =   5130
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1125
   End
   Begin VB.Label lblSum 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼ�"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   390
      TabIndex        =   2
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label lblDigest 
      Caption         =   "���"
      Height          =   240
      Left            =   390
      TabIndex        =   0
      Top             =   1470
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "ժҪ"
      Height          =   240
      Left            =   390
      TabIndex        =   4
      Top             =   4350
      Width           =   480
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   240
      Left            =   390
      TabIndex        =   17
      Top             =   975
      Width           =   480
   End
   Begin VB.Label lblHandle 
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   240
      Left            =   4320
      TabIndex        =   9
      Top             =   4860
      Width           =   720
   End
   Begin VB.Label lblPay 
      BackStyle       =   0  'Transparent
      Caption         =   "�ɿ���"
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   4860
      Width           =   720
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ɿ�Ǽǿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2190
      TabIndex        =   16
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "frmCashPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim colMax As Collection         '����ý��㷽ʽ������ݴ���

Dim mblnChange As Boolean     'Ϊ��ʱ��ʾ�Ѹı���
Dim mdatCurrnet As Date
Dim mblnSuccess As Boolean

Private Sub InitTable()
    dtpDate.Value = mdatCurrnet
    dtpDate.MaxDate = mdatCurrnet
    
    With msh
        .Font.Size = 12
        .CboFont.Size = 12
        .TxtEditFont.Size = 12
        .Cols = 3
        .TextMatrix(0, 0) = "���㷽ʽ"
        .TextMatrix(0, 1) = "���"
        .TextMatrix(0, 2) = "�����"
        .ColWidth(0) = 1350
        .ColWidth(1) = 2475
        .ColWidth(2) = 1350
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        
        .ColData(0) = 3
        .ColData(1) = 4
        .ColData(2) = 4
        .PrimaryCol = 0
        .Active = True
    End With
    
    '��ʼ��Ʊ�ݴ�ӡ
    'On Error Resume Next
    'BillInit gcnOracle
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
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
    If ValidateContent() = False Then Exit Sub
    If Save() = False Then Exit Sub
    mblnChange = False
    mblnSuccess = True
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then msh.SetFocus
End Sub

Private Sub msh_EnterCell(Row As Long, Col As Long)
    Call ShowSum
End Sub

Private Sub msh_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub msh_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If msh.TxtVisible = False Then
        If msh.Col = 1 Then
            If msh.TextMatrix(msh.Row, 1) = "" And msh.TextMatrix(msh.Row, 2) = "" Then
                txtDigest.SetFocus
            End If
        ElseIf msh.Col = 2 Then
            If msh.TextMatrix(msh.Row, 2) = "" Then msh.TextMatrix(msh.Row, 2) = " "
        End If
        Exit Sub
    End If
    '������ֵ������֤
    msh.Text = Trim(msh.Text)
    If msh.Col = 1 Then
        If Not IsNumeric(msh.Text) Then
            MsgBox "����������", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If Val(msh.Text) > 99999999 Then
            MsgBox "������", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If Val(msh.Text) < -999999999 Then
            MsgBox "����С��", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        msh.Text = Format(Val(msh.Text), "###########0.00;-###########0.00;0.00;0.00")
    Else
        If LenB(StrConv(msh.Text, vbFromUnicode)) > 10 Then
            MsgBox "����ŵĳ��Ȳ��ܳ���10λ��", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If InStr(msh.Text, "'") > 0 Then
            MsgBox "����ź��зǷ��ַ���", vbInformation, gstrSysName
            msh.TxtSetFocus
            Cancel = True
            Exit Sub
        End If
        If msh.Text = "" Then msh.Text = " "
    End If
    mblnChange = True
End Sub

Private Sub txtDigest_Change()
    mblnChange = True
End Sub

Private Sub txtDigest_GotFocus()
    zlControl.TxtSelAll txtDigest
    zlCommFun.OpenIme True
End Sub

Private Sub txtDigest_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Function ValidateContent() As Boolean
'����:����������ݵ��Ƿ���Ч
'����:��Ч�򷵻�True,���򷵻�False

    Dim intTemp As Integer
    Dim intTempSub As Integer
    Dim douSum As Double
    
    Dim strJudged As String
    
    ValidateContent = False
    If LenB(StrConv(txtDigest.Text, vbFromUnicode)) > 50 Then
        MsgBox "ժҪ�ĳ��Ȳ��ܳ���25�����ֻ�50����ĸ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtDigest
        txtDigest.SetFocus
        Exit Function
    End If
    If InStr(txtDigest.Text, "'") > 0 Then
        MsgBox "ժҪ���зǷ��ַ���'����", vbInformation, gstrSysName
        zlControl.TxtSelAll txtDigest
        txtDigest.SetFocus
        Exit Function
    End If
    For intTemp = 1 To msh.Rows - 1
        douSum = 0
        If msh.TextMatrix(intTemp, 0) <> "" And msh.TextMatrix(intTemp, 1) <> "" Then
            '���ֽ��㷽ʽ��ǰ���ж���û��
            If InStr(strJudged, msh.TextMatrix(intTemp, 0) & ",") = 0 Then
                'ͳ�Ƴ����ֽ��㷽ʽ���ܽ��
                For intTempSub = intTemp To msh.Rows - 1
                    If msh.TextMatrix(intTempSub, 0) = msh.TextMatrix(intTemp, 0) Then douSum = douSum + Val(msh.TextMatrix(intTempSub, 1))
                Next
                
                If douSum > colMax(msh.TextMatrix(intTemp, 0)) Then     '����������ݴ���
                     If MsgBox(msh.TextMatrix(intTemp, 0) & "�Ľɿ�������ݴ���Ƿ������", vbYesNo Or vbQuestion Or vbDefaultButton2, Me.Caption) = vbNo Then
                        msh.Row = intTemp
                        msh.Col = 1
                        msh.SetFocus
                        msh.TxtSetFocus
                        Exit Function
                     End If
                End If
            strJudged = strJudged & msh.TextMatrix(intTemp, 0) & ","
            End If
        End If
    Next
    ValidateContent = True
End Function

Private Function Save() As Boolean
'����:����༭������
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim intTemp As Integer
    Dim strTemp As String
    Dim lngID As Long, lng���� As Long
    
    On Error GoTo errHandle
    Save = False
    gcnOracle.BeginTrans
    With msh
        lng���� = zlDatabase.GetNextId("��Ա�ɿ��¼")
        For intTemp = 1 To .Rows - 1
            If lngID = 0 Then
                lngID = lng����
            Else
                lngID = zlDatabase.GetNextId("��Ա�ɿ��¼")
            End If
            
            If .TextMatrix(intTemp, 0) <> "" And .TextMatrix(intTemp, 1) <> "" Then
                gstrSQL = "zl_��Ա�ɿ��¼_insert(" & lngID & "," & lng���� & _
                    ",to_date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),'" & txtPay.Text & "','" & txtHandle.Text & _
                    "','" & .TextMatrix(intTemp, 0) & "'," & .TextMatrix(intTemp, 1) & ",'" & .TextMatrix(intTemp, 2) & _
                    "','" & txtDigest.Text & "',Null," & cbo�ɿ��.ItemData(cbo�ɿ��.ListIndex) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    gcnOracle.CommitTrans
    
    '��ӡƱ��
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me, "����ID=" & lng����, 2)  '2��ʾֱ�Ӵ�ӡ
    
    Save = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowSum()
'����:����ɿ��ܶ�
    
    Dim dblTemp As Double
    Dim intTemp As Integer
    
    For intTemp = 1 To msh.Rows - 1
         dblTemp = dblTemp + Val(msh.TextMatrix(intTemp, 1))
    Next
    txtSum.Text = Format(dblTemp, "######0.00;-######0.00;0;") & "Ԫ" & IIf(dblTemp = 0, "", " ��" & zlCommFun.UppeMoney(dblTemp) & "��")
End Sub

Public Function �༭�ɿ��¼(ByVal str�ɿ��� As String, ByVal lng�ɿ���ID As Long) As Boolean
'����:��������õĲ����ش��ڽ���ͨѶ�ĳ���,�������ӽɿ��¼
'����:str�ɿ���     �ɿ��˵�����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intRow As Integer
    
    On Error GoTo errHandle
    
    mblnSuccess = False
    If UserInfo.���� = "" Then
        MsgBox "��ǰ��¼�û�δָ����Ӧ����Ա������ʹ�ñ����ܡ�", vbExclamation, gstrSysName
        Set frmCashPay = Nothing
        Exit Function
    End If
    
    mdatCurrnet = Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    
    txtPay.Text = str�ɿ���
    txtHandle.Text = UserInfo.����
    Set rsTmp = GetPersonnelDept(lng�ɿ���ID)
    Call zlControl.CboAddData(cbo�ɿ��, rsTmp, True)
    If cbo�ɿ��.ListCount > 0 Then cbo�ɿ��.ListIndex = 0
    
    gstrSQL = "Select ���㷽ʽ,��� " & _
            " From ��Ա�ɿ���� Where �տ�Ա =[1] and ����=1 and ���<>0 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�ɿ���)
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "�շ�Ա" & str�ɿ��� & "û���ݴ��������нɿ������", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Call InitTable
    msh.Clear
    msh.Rows = rsTmp.RecordCount + 1
    intRow = 1
    Set colMax = New Collection
    Do Until rsTmp.EOF
        msh.TextMatrix(intRow, 0) = rsTmp("���㷽ʽ")
        msh.TextMatrix(intRow, 1) = Format(rsTmp("���"), "###########0.00;-###########0.00;0.00;0.00")
        msh.TextMatrix(intRow, 2) = " "
        '�������ֵ
        colMax.Add CDbl(rsTmp("���")), CStr(rsTmp("���㷽ʽ"))
        msh.AddItem rsTmp("���㷽ʽ")
        intRow = intRow + 1
        rsTmp.MoveNext
    Loop
    
    mblnChange = False
    frmCashPay.Show vbModal, frmCashSupervise
    �༭�ɿ��¼ = mblnSuccess
    Exit Function
errHandle:
    MsgBox "���ݶ���ʧ�ܡ�", vbExclamation, gstrSysName
    �༭�ɿ��¼ = False
End Function

Private Sub txtDigest_LostFocus()
    zlCommFun.OpenIme False
End Sub
