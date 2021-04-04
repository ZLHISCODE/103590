VERSION 5.00
Begin VB.Form frmReview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Ϣ"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmReview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5655
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboDiagnosisType 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancelReview 
      Caption         =   "ȡ�����"
      Height          =   350
      Left            =   3240
      TabIndex        =   4
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   4440
      TabIndex        =   3
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   3120
      Width           =   1100
   End
   Begin VB.TextBox txtReview 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "��Ϸ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2565
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngOrderID As Long     'ҽ��ID
Private mlngSendNo As Long      '���ͺ�
Private mstrReview As String    '�������
Private mModifyReview As Boolean '�Ƿ��޸��������

Public Function ShowMe(lngOrderID As Long, lngSendNO As Long, frmParent As Object, _
    strDeptName As String, strReview As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    mlngOrderID = lngOrderID
    mlngSendNo = lngSendNO
    
    Me.cboDiagnosisType.Clear
    strSQL = "select ���� from Ӱ����Ϸ��� Where ��������= [1] order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ѡ����Ϸ���", strDeptName)
    While Not rsTemp.EOF
        Me.cboDiagnosisType.AddItem rsTemp!����
        rsTemp.MoveNext
    Wend
    
    strSQL = "Select �������,��Ϸ��� From Ӱ�����¼ Where ҽ��id=[1] And ���ͺ� = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngOrderID, mlngSendNo)
    
    If Not rsTemp.EOF Then
        Me.txtReview.Text = Nvl(rsTemp!�������)
        Me.cboDiagnosisType.Text = Nvl(rsTemp!��Ϸ���)
    Else
        Me.txtReview.Text = ""
        Me.cboDiagnosisType.Text = ""
    End If
    
    Me.Show 1, frmParent
    
    strReview = mstrReview
    ShowMe = mModifyReview
End Function

Private Sub cmdCancel_Click()
    mModifyReview = False
    Unload Me
End Sub

Private Sub cmdCancelReview_Click()
    Dim strSQL As String
    If MsgBoxD(Me, "�Ƿ������ü�¼��", vbOKCancel) = vbOK Then
        strSQL = "Zl_Ӱ�����_Update(" & mlngOrderID & ",'')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        mstrReview = ""
        mModifyReview = True
        
        Unload Me
   End If
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    
    arrSQL = Array()
    
    On Error GoTo errHandle
    
    strSQL = "Zl_Ӱ����Ϸ���_Update(" & mlngOrderID & ",'" & cboDiagnosisType.Text & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    
    strSQL = "Zl_Ӱ�����_Update(" & mlngOrderID & ",'" & txtReview.Text & "')"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
        
    gcnOracle.BeginTrans        '----------������Ϸ�������
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������Ϸ�������")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    mstrReview = Me.txtReview.Text
    mModifyReview = True
    
    Unload Me
    
    Exit Sub
errHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
        Call SaveErrLog
End Sub

