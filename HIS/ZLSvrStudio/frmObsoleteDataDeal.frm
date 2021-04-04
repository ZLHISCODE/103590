VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmObsoleteDataDeal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "���ݴ���"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin XtremeSyntaxEdit.SyntaxEdit steProcedure 
      Height          =   7995
      Left            =   75
      TabIndex        =   5
      Top             =   405
      Width           =   9090
      _Version        =   983043
      _ExtentX        =   16034
      _ExtentY        =   14102
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   -1  'True
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�(&C)"
      Height          =   350
      Left            =   8070
      TabIndex        =   4
      Top             =   8580
      Width           =   1100
   End
   Begin VB.TextBox txtDays 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   945
      MaxLength       =   6
      TabIndex        =   3
      Top             =   8685
      Width           =   465
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "ִ��(&E)"
      Height          =   350
      Left            =   6960
      TabIndex        =   0
      Top             =   8580
      Width           =   1100
   End
   Begin VB.Line linNumber 
      X1              =   930
      X2              =   1410
      Y1              =   8895
      Y2              =   8895
   End
   Begin VB.Label lblDays 
      AutoSize        =   -1  'True
      Caption         =   "���ݱ���     ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   2
      Top             =   8670
      Width           =   1575
   End
   Begin VB.Label lblProcedure 
      AutoSize        =   -1  'True
      Caption         =   "�洢���̣�"
      Height          =   180
      Left            =   75
      TabIndex        =   1
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmObsoleteDataDeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrBusinessName As String
Private mstrProcedureName As String
Private mstrUserName As String
Private mDateNow As Date
Private mlngDays As Long
Private mblnOK As Boolean

Public Function ShowMe(ByVal strBusinessName As String, ByVal strProcedureName As String, ByRef lngDays As Long, ByRef strUsername As String, ByRef dateNow As Date) As Boolean
    mstrBusinessName = strBusinessName
    mstrProcedureName = strProcedureName
    mstrUserName = strUsername
    mlngDays = lngDays
    Me.Show vbModal, frmMDIMain
    lngDays = mlngDays
    dateNow = mDateNow
    strUsername = mstrUserName
    ShowMe = mblnOK
End Function

Private Sub FillProcedure()
'���洢����
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strProcedure As String
    
    On Error GoTo errH
    strSQL = "Select Text From Sys.User_Source Where Name = [1] Order By Line"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "��ȡ�洢����", UCase(mstrProcedureName))
    With rsTemp
        Do While Not .EOF
            strProcedure = strProcedure & !Text
            .MoveNext
        Loop
        If strProcedure = "" Then
            strProcedure = "�ô洢���̲����ڣ����飡"
        End If
        steProcedure.Text = strProcedure
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExec_Click()
'����������Ϣ��ִ�д洢����
    Dim strProcedure As String
    
    On Error GoTo errH
    
    If Val(txtDays.Text) < 1 Then
        MsgBox "������������Ϊ1�죬�����µ�����", vbInformation, gstrSysName
        txtDays.Text = mlngDays
        txtDays.SetFocus
        Exit Sub
    End If
    
    If MsgBox("�����������ݿ��ܻ�Ӱ������ҵ���ʹ�ã�ȷ������������", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    mDateNow = CurrentDate()
    strProcedure = mstrProcedureName
    Call ExecuteProcedure(strProcedure, "ִ�д洢����")
    
    Call ExecuteProcedure("Zltools.Zl_Zlobsoletedatadeal_Update('" & mstrBusinessName & "',Null," & _
                        Val(txtDays.Text) & ",'" & _
                        gstrLoginUserName & "','" & _
                        mDateNow & "')", "�޸ı�����������Աʱ����Ϣ")
    '������Ҫ������־
    Call SaveAuditLog(2, "���ݴ���", "ִ�й��̡�" & mstrProcedureName & "����ҵ��" & mstrBusinessName & _
                        "�������ݽ��д����������ݱ���������" & mlngDays & "���޸�Ϊ" & Val(txtDays.Text) & "��")

    mlngDays = Val(txtDays.Text)
    mstrUserName = gstrLoginUserName
    mblnOK = True
    MsgBox "����ִ�гɹ���", vbInformation, gstrSysName
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Load()
    lblProcedure.Caption = "�洢���̣�" & mstrProcedureName
    txtDays.Text = mlngDays
    mblnOK = False
    Call FillProcedure
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
'������ƣ�Ҫ��ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub
