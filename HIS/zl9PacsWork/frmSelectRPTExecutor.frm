VERSION 5.00
Begin VB.Form frmSelectRPTExecutor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ�񱨸�ҽ��"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3465
   Icon            =   "frmSelectRPTExecutor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ListBox lstRPTExecutor 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSelectRPTExecutor.frx":000C
      TabIndex        =   1
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSelectRPTExecutor.frx":136E
      TabIndex        =   0
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1100
   End
End
Attribute VB_Name = "frmSelectRPTExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrResult As String

Public Function GetRPTExecutor(ByVal lngCurDeptId As Long, ByVal objParent As Object, Optional strRPTExecutor As String = "") As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "Select /*+ RULE*/" & vbNewLine & _
                "Distinct b.id,b.����, Upper(b.����) As ����" & vbNewLine & _
                " From ������Ա a, ��Ա�� b " & vbNewLine & _
                " Where a.��Աid = b.Id And " & vbNewLine & _
                "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and a.����id = [1] " & vbNewLine & _
                " Order By ���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCurDeptId)
    
    lstRPTExecutor.Clear
    Do Until rsTemp.EOF
        lstRPTExecutor.AddItem rsTemp!���� & "-" & rsTemp!����
        If rsTemp!ID = UserInfo.ID Then lstRPTExecutor.ListIndex = lstRPTExecutor.NewIndex
        If rsTemp!���� = strRPTExecutor Then lstRPTExecutor.ListIndex = lstRPTExecutor.NewIndex
        rsTemp.MoveNext
    Loop
    
    If lstRPTExecutor.ListCount > 0 Then If lstRPTExecutor.ListIndex < 0 Then lstRPTExecutor.ListIndex = 0
    
    Me.Show 1, objParent
    
    GetRPTExecutor = mstrResult
End Function


Private Sub CmdOK_Click()
    If lstRPTExecutor.ListCount > 0 Then
        If lstRPTExecutor.ListIndex >= 0 Then
            mstrResult = Split(lstRPTExecutor.list(lstRPTExecutor.ListIndex), "-")(1)
        End If
    End If
    Unload Me
End Sub

Private Sub Command1_Click()
    mstrResult = ""
    Unload Me
End Sub
