VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmDocRead 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "����������Ϣ"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin SHDocVwCtl.WebBrowser WBMethod 
      Height          =   6690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      ExtentX         =   13150
      ExtentY         =   11800
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmDocRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrUrl As String
Private mstrPatiID As String
 

'�ô���ΪǶ�״���
Public Function zlRefresh(ByVal strPatiID As String, ByVal str��ҳID As String) As Boolean
    Dim strErrMsg As String
    If strPatiID = "" Or gstrIntergrateIP = "" Then Exit Function
    If mstrPatiID = strPatiID Then Exit Function
    If gobjComlib.OS.IsDesinMode = False Then  '��IDEģʽ��ȡִ�г���exe(���Ի���ʼ����vb6.exe),�Լ����webBrowser�ؼ�Ĭ��ʹ��IE11�����(ֻ�������л����¸����ò�����)
        If SetWBIEVerSion(gstrExeName, strErrMsg) = False Then
            MsgBox strErrMsg, vbInformation, gstrSysName
        End If
    End If
    WBMethod.Navigate mstrUrl & "PatientID=" & strPatiID & "&HomePageID=" & str��ҳID & "&Doctorname=" & UserInfo.���� & "&Username=" & UserInfo.�û��� & "&RYID=" & UserInfo.id

    SetProcessWorkingSetSize GetCurrentProcess, -1, -1  '���õ�ǰ
    mstrPatiID = strPatiID
    zlRefresh = True
End Function



Private Sub Form_Load()
    mstrUrl = "http://" & gstrIntergrateIP & "/ascore?"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With WBMethod
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub


Private Sub SetIERelog()
'���ܣ�����ע���
      '��������Registry   Access   Functions
      Dim mREG As New REGTool5.Registry
      Dim ret As Boolean
      ret = mREG.UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet   Explorer\Main ", "Disable   Script   Debugger ", "yes")
      ret = mREG.UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet   Explorer\Main ", "DisableScriptDebuggerIE ", "yes")
End Sub


Private Sub WBMethod_DownloadBegin()
    WBMethod.Silent = True
End Sub

Private Sub WBMethod_DownloadComplete()
    If WBMethod.LocationURL = "http:///" Then Exit Sub
End Sub
