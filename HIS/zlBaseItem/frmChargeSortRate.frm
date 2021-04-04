VERSION 5.00
Begin VB.Form frmChargeSortRate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͳһʵ�ձ���"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmChargeSortRate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   1725
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   3195
      Begin VB.TextBox txtPercentage 
         Height          =   300
         Left            =   1440
         TabIndex        =   2
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ʵ�ձ���(&P)"
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Top             =   1110
         Width           =   150
      End
      Begin VB.Label lbl���� 
         Caption         =   "    ���е�������Ŀ����ͳһ��ʵ�ձ��ʡ���ǰ����Ҫ�����ǰ�ĸ�������Ŀ��δ�����ֶΡ�"
         Height          =   585
         Left            =   450
         TabIndex        =   1
         Top             =   330
         Width           =   2595
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3570
      TabIndex        =   4
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   3570
      TabIndex        =   6
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   5
      Top             =   690
      Width           =   1100
   End
End
Attribute VB_Name = "frmChargeSortRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOk As Boolean       '�Ƿ�ɹ�
Dim mstr�ѱ� As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandle
    Dim str����  As String
    
    str���� = Trim(txtPercentage.Text)
    If str���� = "" Then
        MsgBox "ʵ�ձ��ʲ���Ϊ�ա�", vbExclamation, gstrSysName
        txtPercentage.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(str����) Then
        MsgBox "ʵ�ձ���Ӧ����һ����ֵ��", vbExclamation, gstrSysName
        txtPercentage.SetFocus
        zlControl.TxtSelAll txtPercentage
        Exit Sub
    End If
    If Val(str����) < 0 Or Val(str����) > 500 Then
        MsgBox "ʵ�ձ���ֻ���� 0��500֮�䡣", vbExclamation, gstrSysName
        txtPercentage.SetFocus
        zlControl.TxtSelAll txtPercentage
        Exit Sub
    End If
    
    gstrSQL = "zl_�ѱ�_Unify('" & mstr�ѱ� & "'," & txtPercentage.Text & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    mblnOk = True
    Unload Me
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function UnifyPercentage(ByVal str�ѱ� As String, ByVal lng���� As Long) As Boolean
'����:��������õ�������Ŀ�����ڽ���ͨѶ�ĳ���
'����:str�ѱ�     Ҫ���õķѱ�
'     lng����     �ٷֱ�
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select B.���� from �ѱ���ϸ A,������Ŀ B " & _
           " where a.������ĿID=B.ID and A.�ѱ�=[1] " & _
           " group by B.ID,B.����  having count(B.ID)>1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�ѱ�)
        
    If rsTemp.RecordCount > 0 Then
        Do Until rsTemp.EOF
            strTemp = strTemp & rsTemp("����") & ","
            rsTemp.MoveNext
        Loop
        strTemp = "    " & Mid(strTemp, 1, Len(strTemp) - 1)
        MsgBox "����������Ŀ�Ѿ��ֶΣ�" & vbCrLf & strTemp & vbCrLf & "���������ܼ�����", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mstr�ѱ� = str�ѱ�
    Frame1.Caption = str�ѱ�
    txtPercentage.Text = Format(lng����, "###0.00;-##0.00;0.00;0.00")
    
    mblnOk = False
    frmChargeSortRate.Show vbModal, frmChargeSortGrade
    UnifyPercentage = mblnOk
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPercentage_GotFocus()
    zlControl.TxtSelAll txtPercentage
    OS.OpenIme False
End Sub
