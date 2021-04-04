VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMethod 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser WBMethod 
      Height          =   6690
      Left            =   180
      TabIndex        =   0
      Top             =   285
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
Attribute VB_Name = "frmMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrUrl As String
Private mstrParam As String
 
Private WithEvents M_Dom As HTMLDocument
Attribute M_Dom.VB_VarHelpID = -1

'该窗体为嵌套窗体
Public Function zlRefresh(ByVal strUrl As String, ByVal strParam As String) As Boolean
    Dim strErrMsg As String
    If strUrl = "" Then Exit Function
    mstrUrl = strUrl & "?From=his"
    mstrParam = strParam
    If gobjComlib.OS.IsDesinMode = False Then  '非IDE模式获取执行程序exe(调试环境始终是vb6.exe),以及设计webBrowser控件默认使用IE11浏览器(只有在运行环境下该设置才有用)
        If SetWBIEVerSion(gstrExeName, strErrMsg) = False Then
            MsgBox strErrMsg, vbInformation, gstrSysName
        End If
    End If
    WBMethod.Navigate mstrUrl
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1  '设置当前
    zlRefresh = True
End Function

Private Sub Form_Resize()
    On Error Resume Next
    With WBMethod
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub ShowDialog(ByVal strParam As String)
'监听事件参数：
'{
'  type: "CloseDialog" || "ShowDialog" || "LOG", // CloseDialog 关闭弹窗  ShowDialog 打开弹窗  LOG写入日志数据
'  moduleUrl: "/shiftReport", //功能Url
'  title: "交班报告",
'  width: "100" || null,
'  height: "100" || null,
'  minimal: true,  //最大化
'  max: false,     //最小化
'  isRefresh: true  //是否刷新父窗体
'  data: "xxxxxxxxxxxx"  //打开弹窗是需要带上的参数
'}
'说明：主窗体程序只能只处理打开子窗体，调用该窗体不能关闭(由于是嵌套窗体)
    Dim objPopup As New clsPopup
    Dim objFrom As Object

    If strParam = "" Then Exit Sub
    Call WriteBusinessLOG("监听事件", mstrUrl, strParam)
    If AnalysisJavaScriptEvent(strParam, objPopup) = False Then Exit Sub
    If UCase(objPopup.PopupType) = UCase("ShowDialog") Then 'ShowDialog则认为是打开子窗体
        If frmChild.ShowMe(Me, objPopup) = True Then
            Call WriteBusinessLOG("主窗体刷新", mstrUrl, mstrParam)
            Call Me.WBMethod.Refresh
        End If
    ElseIf UCase(objPopup.PopupType) = UCase("LOG") Then
        Call WriteBusinessLOG("监听事件LOG", objPopup.PopupParentUrl, objPopup.PopupData)
    End If
End Sub

Private Function M_Dom_onclick() As Boolean
'点击网页任何地方都会触发此事件，只识别格式为：<span id="sendToZlhis" style="display:none;">参数写在这里面</span>
    Dim strPane As String
    Dim objXML As New DOMDocument
    Dim objNodeList As IXMLDOMNodeList
    Dim strPram As String
    
    strPane = M_Dom.parentWindow.event.srcElement.outerHTML
    'MsgBox strPane
    If strPane Like "<span id=""sendToZlhis""*" Then
        If objXML.loadXML(strPane) = False Then
            MsgBox "不是有效的XML格式：" & strPane, vbInformation, gstrSysName
            Exit Function
        End If
        Set objNodeList = objXML.selectNodes(".//span")
        strPram = decodeURIComponent(objNodeList.Item(0).Text)
        Call ShowDialog(strPram)
    End If
    M_Dom_onclick = True
End Function

Private Function M_Dom_oncontextmenu() As Boolean
    M_Dom_oncontextmenu = True
End Function

Private Sub WBMethod_DownloadBegin()
    WBMethod.Silent = True
End Sub

Private Sub WBMethod_DownloadComplete()
    Set M_Dom = WBMethod.Document
    If WBMethod.LocationURL = "http:///" Then Exit Sub
    On Error Resume Next
    M_Dom.selection.createRange.pasteHTML "<span id=""sendToWeb"" style=""display: none;"">" & encodeURIComponent(mstrParam) & "</span>"
    If Err <> 0 Then Err.Clear
End Sub

Private Sub SetIERelog()
'功能：设置注册表
      '首先引用Registry   Access   Functions
      Dim mREG As New REGTool5.Registry
      Dim ret As Boolean
      ret = mREG.UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet   Explorer\Main ", "Disable   Script   Debugger ", "yes")
      ret = mREG.UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet   Explorer\Main ", "DisableScriptDebuggerIE ", "yes")
End Sub
