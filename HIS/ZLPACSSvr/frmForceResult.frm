VERSION 5.00
Begin VB.Form frmBuildResult 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����Զ�����"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   Icon            =   "frmForceResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdVerify 
      Caption         =   "��֤(&V)"
      Height          =   350
      Left            =   1080
      TabIndex        =   6
      Top             =   4080
      Width           =   1100
   End
   Begin VB.TextBox txtBuildResult 
      Height          =   3135
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.ListBox lstDBItem 
      Height          =   3120
      ItemData        =   "frmForceResult.frx":000C
      Left            =   120
      List            =   "frmForceResult.frx":000E
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   1100
   End
   Begin VB.Label Label2 
      Caption         =   "�Զ�������"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "��ѡ�����ݿ���Ŀ��"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmBuildResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strReturnString As String

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If funVerifyResult(Me.txtBuildResult.Text) = 0 Then
        strReturnString = Me.txtBuildResult.Text
        Unload Me
    End If
End Sub

Private Sub cmdVerify_Click()
    funVerifyResult Me.txtBuildResult.Text
End Sub

Public Function funVerifyResult(strString As String) As Integer
    '����ֵ��1-��ʽ����[]��ƥ�䣻2��[]�в������ݿ��ֶΡ�
    Dim strTemp As String
    Dim strField As String
    Dim i As Integer, iPoint As Integer
    
    funVerifyResult = 0
    strTemp = strString
    
    '���ַ����У����������'��
    If InStr(strTemp, "'") > 0 Then
        funVerifyResult = 1
        MsgBox "�Զ�������ʽ���ַǷ��ַ������޸ĺ�������֤��"
        Exit Function
    End If
    
    iPoint = 1
    Do While iPoint <= Len(strTemp)
        iPoint = InStr(iPoint, strTemp, "[")
        If iPoint = 0 Then Exit Do
        
        If InStr(iPoint, strTemp, "]") = 0 Then
            funVerifyResult = 1
            Exit Do
        End If
        
        strField = Mid(strTemp, iPoint + 1, InStr(iPoint, strTemp, "]") - iPoint - 1)
        For i = 0 To Me.lstDBItem.ListCount - 1
            If strField = Me.lstDBItem.list(i) Then Exit For
        Next i
        If i >= Me.lstDBItem.ListCount Then
            funVerifyResult = 2
            Exit Do
        End If
'        strTemp = Right(strTemp, Len(strTemp) - InStr(strTemp, "]"))
        iPoint = iPoint + 1
    Loop
    
    '������
    If funVerifyResult = 1 Then
        MsgBox "�Զ�������ʽ�д��󣬡�[���͡�]��������ƥ�䣬���޸ĺ�������֤��"
    ElseIf funVerifyResult = 2 Then
        MsgBox "�Զ��������д��󣬡�[���͡�]���а��������ֲ���ϵͳ�ṩ�����ݿ��ֶΣ����޸ĺ�������֤��"
    End If
End Function

Private Sub Form_Load()
    Me.lstDBItem.Clear
    Me.lstDBItem.AddItem "CallingAET"
    Me.lstDBItem.AddItem "�״�����"
    Me.lstDBItem.AddItem "�״�ʱ��"
    Me.lstDBItem.AddItem "Ӱ�����"
    Me.lstDBItem.AddItem "ִ�м�"
    Me.lstDBItem.AddItem "ִ�й���"
    Me.lstDBItem.AddItem "ҽ��ID"
    Me.lstDBItem.AddItem "���ͺ�"
    Me.lstDBItem.AddItem "����"
    Me.lstDBItem.AddItem "��ʶ��"
    Me.lstDBItem.AddItem "Ӣ����"
    Me.lstDBItem.AddItem "�Ա�"
    Me.lstDBItem.AddItem "����"
    Me.lstDBItem.AddItem "��������"
    Me.lstDBItem.AddItem "������"
    Me.lstDBItem.AddItem "����豸"
    strReturnString = ""
End Sub

Private Sub lstDBItem_DblClick()
Dim intStart As Integer
    Dim strTemp As String
    intStart = Me.txtBuildResult.SelStart
    strTemp = Me.txtBuildResult.Text
    Me.txtBuildResult.Text = Left(strTemp, intStart) & "[" & Me.lstDBItem.list(Me.lstDBItem.ListIndex) _
                            & "]" & Right(strTemp, Len(strTemp) - intStart)
    Me.txtBuildResult.SelStart = intStart + Len(Me.lstDBItem.list(Me.lstDBItem.ListIndex)) + 2
    Me.txtBuildResult.SetFocus
End Sub
