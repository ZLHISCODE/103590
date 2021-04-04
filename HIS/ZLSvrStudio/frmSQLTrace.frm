VERSION 5.00
Begin VB.Form frmSQLTrace 
   BackColor       =   &H80000005&
   Caption         =   "���ٹ���"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmSQLTrace.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   6315
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEnter 
      Caption         =   "���ڽ�����ٹ���(&E)�� "
      Height          =   350
      Left            =   840
      TabIndex        =   0
      Top             =   3600
      Width           =   2190
   End
   Begin VB.Image imgMain 
      Height          =   720
      Left            =   360
      Picture         =   "frmSQLTrace.frx":803A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Height          =   3330
      Left            =   870
      TabIndex        =   2
      Top             =   615
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL����"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmSQLTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFilePath As String

Private Sub cmdEnter_Click()
    On Error GoTo errh
    
    Call ShowFlash("���ڼ���SQL���ٹ���...")
    'Shell "E:\vb project\zlSvrTools\SQLTrace.exe zlUserName=ZLHISzlPassword=HISzlServer=QZYY"
    Shell mstrFilePath & "\ZLSQLTrace.exe zlUserName=" & gstrUserName & "zlPassword=" & gstrPassword & "zlServer=" & gstrServer
    Call ShowFlash("")
    Exit Sub
errh:
    Call ShowFlash("")
    MsgBox "����" & mstrFilePath & "\ZLSQLTrace.exe  �Ƿ���ڡ�"
End Sub

Private Sub Form_Load()
    
    lblMain.Caption = "������ͨ��Oracle��SQLTrace���������ٺͷ���SQL�������⡣" & _
    vbCrLf & vbCrLf & "֧�ֶ�ָ�����û��Ự����SQL���٣��ӷ�������ȡSQLTrace�ļ����ͻ��ˣ��Լ�����SQLTrace�ļ���" & _
    vbCrLf & vbCrLf & "֧�ֶԶ��SQLTrace�ļ����жԱȣ��Լ����ٹ��˳��������������SQL��䣬����ز鿴����ִ�мƻ���"
    
    '���û�ȡZLSQLTrace.EXE��·��
    mstrFilePath = GetSetting("ZLSOFT", "����ȫ��", "����·��", App.Path)
    If mstrFilePath = App.Path Then
        mstrFilePath = App.Path
    Else
        'C:\APPSOFT\ZLHIS+.exe
        mstrFilePath = Mid(mstrFilePath, 1, InStrRev(mstrFilePath, "\") - 1)
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    
    With lblMain
        .Top = imgMain.Top
        .Height = Me.ScaleHeight - .Top * 2
        .Left = imgMain.Left * 2 + imgMain.Width
        .Width = Me.ScaleWidth - .Left - imgMain.Left
    End With

    Dim intCount As Integer, intRows As Integer, aryRow() As String
    intRows = 1
    aryRow() = Split(lblMain.Caption, vbCrLf)
    For intCount = 0 To UBound(aryRow)
        intRows = intRows + TextWidth(aryRow(intCount)) \ (lblMain.Width - 90) + 1
    Next
    If intRows * TextHeight("A") < lblMain.Height + TextHeight("A") Then
        cmdEnter.Top = lblMain.Top + intRows * TextHeight("A")
    Else
        cmdEnter.Top = lblMain.Top + lblMain.Height + TextHeight("A")
    End If
    cmdEnter.Left = lblMain.Left
    
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

