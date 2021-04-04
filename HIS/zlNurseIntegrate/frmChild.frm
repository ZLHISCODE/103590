VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmChild 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9180
   Icon            =   "frmChild.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin SHDocVwCtl.WebBrowser WBMethod 
      Height          =   5220
      Left            =   765
      TabIndex        =   0
      Top             =   795
      Width           =   6780
      ExtentX         =   11959
      ExtentY         =   9208
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
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents M_Dom As HTMLDocument
Attribute M_Dom.VB_VarHelpID = -1

Private mblnRefresh As Boolean
Private mstrUrl As String
Private mstrParam As String
Private mblnActive As Boolean
Private mstrParentUrl As String
Private mstrParentParams As String

Public Function ShowMe(ByVal objParent As Object, objPopup As clsPopup) As Boolean
'功能：弹出整体护理编辑子窗体
'参数：objParent：调用父窗体
'          strParam：返回的子窗体属性字符串
'          strLessionID：整体护理病区ID (通过RelatedIDToGUID接口获取)
'          strUserID：整体用户ID (通过RelatedIDToGUID接口获取)
'返回是否需要刷新主要窗体
    Dim strUrl As String, strParams As String

    mblnRefresh = False
    
    Me.Width = (objPopup.PopupWidth / 100) * Screen.Width
    Me.Height = (objPopup.PopupHeight / 100) * Screen.Height
    Me.Caption = objPopup.PopupTitle

    '弹窗的URL参数
    'LessionID; 病区ID(通过RelatedIDToGUID接口获取)
    'UserID; 用户ID(通过RelatedIDToGUID接口获取)
    'From; his(固定参数不变, 代表从zlhis打开页面)
    'Data; 该参数来自ocx监听中所得到的数据的Data字段
    strParams = objPopup.PopupParams '"LessionID=" & objPopup.PopupUnitID & "&UserID=" & objPopup.PopupUserID & "&From=his&Data=" & objPopup.PopupData & "&PatientID=" & objPopup.PopupPatientID
    strUrl = "http://" & gstrIntergrateIP & objPopup.PopupModuleUrl
    mstrUrl = strUrl & "?From=his"
    mstrParam = strParams
    mstrParentUrl = objPopup.PopupParentUrl
    mstrParentParams = objPopup.PopupParentParam

    If Not objParent Is Nothing Then
        Me.Show 1, objParent
    Else
        Me.Show 1
    End If
    ShowMe = mblnRefresh
End Function
'
Private Sub Form_Activate()
    If mblnActive = True Then Exit Sub
    mblnActive = True
    Call WriteBusinessLOG(Me.Caption, mstrUrl, mstrParam)
    WBMethod.Navigate mstrUrl
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1  '设置当前
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Call WBMethod.Refresh
    End If
End Sub
'
Private Sub Form_Load()
    mblnActive = False
End Sub
'
Private Sub Form_Resize()
    On Error Resume Next
    With WBMethod
        .Left = 0
        .Top = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
End Sub

Private Sub ShowDialog(ByVal strParam As String)
'监听事件参数：
'{
'  type: "CloseDialog" || "ShowDialog""LOG", // CloseDialog 关闭弹窗  ShowDialog 打开弹窗  LOG写入日志数据
'  moduleUrl: "/shiftReport", //功能Url
'  title: "交班报告",
'  width: "100" || null,
'  height: "100" || null,
'  minimal: true,  //最大化
'  max: false,     //最小化
'  isRefresh: true  //是否刷新父窗体
'  data: "xxxxxxxxxxxx"  //打开弹窗是需要带上的参数
'}
    Dim objPopup As New clsPopup
    Dim objFrom As Object

    If strParam = "" Then Exit Sub
    Call WriteBusinessLOG("监听事件", mstrUrl, strParam)
    If AnalysisJavaScriptEvent(strParam, objPopup) = False Then Exit Sub
    If UCase(objPopup.PopupType) = UCase("ShowDialog") Then 'ShowDialog则认为是打开子窗体
        Set objFrom = New frmChild
        If objFrom.ShowMe(Me, objPopup) Then
            Me.WBMethod.Refresh
        End If
    ElseIf UCase(objPopup.PopupType) = UCase("CloseDialog") Then 'CloseDialog则认为是关闭子窗体
        mblnRefresh = objPopup.PopupIsRefresh
        Unload Me
    ElseIf UCase(objPopup.PopupType) = UCase("LOG") Then
        Call WriteBusinessLOG("监听事件LOG", objPopup.PopupParentUrl, objPopup.PopupData)
    ElseIf UCase(objPopup.PopupType) = UCase("CloseRefresh") Then '127510：打印刷新主列表数据
        mblnRefresh = True
    Else
        mblnRefresh = objPopup.PopupIsRefresh
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


