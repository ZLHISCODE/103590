VERSION 5.00
Object = "{764C1FE7-DC41-4928-A8DC-B9939F37244B}#1.0#0"; "SBrowser_G.ocx"
Begin VB.Form frmBrowser 
   BorderStyle     =   0  'None
   Caption         =   "11"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin SBrowser_G.SBrowser SBrowser 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrUrl As String
Private mapi As New MiniblinkAPI

Private Sub SBrowser_LoadUrlBegin(ByVal title As String, ByVal url As String, ByVal job As Long)
        
    '打开url protocol中指定的程序（解决miniblink控件中无法弹出打印产程图的问题），并中断连接跳转
    If InStr(url, "openchrome") > 0 Then
        Dim strURL As String, strPath As String
        Dim varTemp As Variant

        strPath = "C:\Program Files (x86)\Google\Chrome\Application"
        strPath = GetSetting(appName:="ZLSOFT", Section:="公共", Key:="WEB浏览器路径", Default:=strPath)
        varTemp = Split(url, ":")
        strURL = Replace(url, varTemp(0) & ":", "")
        Shell """" & strPath & "\chrome.exe" & """ " & strURL
        
        mapi.wkeNetCancelRequest job
    End If
    
End Sub

Public Sub InitLoad()
'功能：登录并加载窗体，但不显示出来
    SBrowser.LoadURL gstrURLLogin
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    SBrowser.Move Me.ScaleLeft, Me.ScaleTop, Me.Width, Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrUrl = ""
End Sub


Public Sub RefreshForm(lngPatiID As Long, lngPageID As Long)
'功能：显示指定病人的助产信息
    Dim strSToken  As String, strURL As String
        
    If lngPatiID = glngPatiID And lngPageID = glngPageID Then Exit Sub
    If mstrUrl <> "" Then
       
        strSToken = Split(Split(mstrUrl, "?")(1), ":")(2)
        strURL = Replace(gstrURL, "[SESSION_TOKEN]", strSToken) & lngPatiID & "," & lngPageID
                
        SBrowser.LoadURL strURL
    End If
    glngPatiID = lngPatiID
    glngPageID = lngPageID
End Sub

Private Sub SBrowser_DocumentReady(ByVal url As String)
    mstrUrl = url
End Sub

