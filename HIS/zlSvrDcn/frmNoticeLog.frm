VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmNoticeLog 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin RichTextLib.RichTextBox rtxLog 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmNoticeLog.frx":0000
   End
End
Attribute VB_Name = "frmNoticeLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mdtLogTim As Date
Private mblnLogFileExist As Boolean
Private mstrLogFile As String
Private mobjStream As TextStream

Public Sub ClearLog()
    rtxLog.Text = ""
End Sub

Public Sub WriteLog(ByVal strLog As String, ByVal intType As Integer)
    '����:������ʾ��־
    'strLog = ��־��Ϣ,  intType = ��־���� ,1-��־���� 2-������־

    If (Len(rtxLog.Text) - Len(Replace(rtxLog.Text, vbNewLine, ""))) > 102 Then '��ౣ��100��
        rtxLog.Text = ""
    End If
    rtxLog.Text = rtxLog.Text & IIf(rtxLog.Text = "", "", vbNewLine) & strLog
    
    '��־���汾��
    If intType = 1 Then
        WriteTraceFile Time & "  " & strLog
    Else
        WriteTraceFile strLog
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With rtxLog
        .Width = Me.ScaleWidth - .Left - 60
        .Height = Me.ScaleHeight - .Top - 120
    End With
    
End Sub

Private Sub WriteTraceFile(ByVal strLog As String)
    '����: ��д��־�ļ�
    Dim strLogFile As String

    If gintLog <> 1 Then Exit Sub
    
    If mblnLogFileExist = False Or mdtLogTim <> Date Then     '�жϵ�����־�ļ��Ƿ����,��������ھʹ���
        strLogFile = GetLogPath & "\zl_Notice" & Replace(Date, "/", "") & ".log"
        
        If Not gobjFile.FileExists(strLogFile) Then
            gobjFile.CreateTextFile strLogFile
        End If
        
        mstrLogFile = strLogFile
        mblnLogFileExist = True
        mdtLogTim = Date
        Set mobjStream = Nothing
    End If
    
    If mobjStream Is Nothing Then
        Set mobjStream = gobjFile.OpenTextFile(mstrLogFile, ForAppending)
    End If
    
    mobjStream.Write strLog & vbNewLine
End Sub
