VERSION 5.00
Begin VB.Form frmDockEx 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label lblTest 
      AutoSize        =   -1  'True
      Caption         =   "��չ�������Կ�Ƭҳǩ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   1020
      Width           =   3300
   End
End
Attribute VB_Name = "frmDockEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function GetInSideFunc() As String
'���ܣ���ȡ���幦����
'���ŵ�ͼ���Ӧ���� "����,3001|�޸�,3003|����,3091|ɾ��,3004|��,100|����,3565|�ر�,3021|��ӡ,103|Ԥ��,102|�˳�,2613|���,225|����,901|����,731|����,721|����,181|ˢ��,791|ǩ��,804"
    GetInSideFunc = "����,3001|�޸�,3003|����,3091|ɾ��,3004|��,100|����,3565|�ر�,3021|��ӡ,103|Ԥ��,102|�˳�,2613|���,225|����,901|����,731|����,721|����,181|ˢ��,791|ǩ��,804"
End Function

Public Function ExecuteFunc(ByVal strName As String) As Boolean
'���ܣ�ִ�в˵��ϵĹ��ܡ��������������
'������strName ��������
    If Not Me.Visible Then Exit Function
    Debug.Print "zlPlugIn/frmDockEx/ExecuteFunc " & strName & "���ܱ�ִ�У�����"
    ExecuteFunc = True
End Function

Public Sub RefreshInSide()
'���ܣ�����ˢ�¡��������������
    Debug.Print "zlPlugIn/frmDockEx/RefreshInSide ��ǰ��Ƭ��ˢ�£�����"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Debug.Print "ж���ˣ���������"
End Sub
