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
   StartUpPosition =   3  '����ȱʡ
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

'�ô���ΪǶ�״���
Public Function zlRefresh(ByVal strUrl As String, ByVal strParam As String) As Boolean
    Dim strErrMsg As String
    If strUrl = "" Then Exit Function
    mstrUrl = strUrl & "?From=his"
    mstrParam = strParam
    If gobjComlib.OS.IsDesinMode = False Then  '��IDEģʽ��ȡִ�г���exe(���Ի���ʼ����vb6.exe),�Լ����webBrowser�ؼ�Ĭ��ʹ��IE11�����(ֻ�������л����¸����ò�����)
        If SetWBIEVerSion(gstrExeName, strErrMsg) = False Then
            MsgBox strErrMsg, vbInformation, gstrSysName
        End If
    End If
    WBMethod.Navigate mstrUrl
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1  '���õ�ǰ
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
'�����¼�������
'{
'  type: "CloseDialog" || "ShowDialog" || "LOG", // CloseDialog �رյ���  ShowDialog �򿪵���  LOGд����־����
'  moduleUrl: "/shiftReport", //����Url
'  title: "���౨��",
'  width: "100" || null,
'  height: "100" || null,
'  minimal: true,  //���
'  max: false,     //��С��
'  isRefresh: true  //�Ƿ�ˢ�¸�����
'  data: "xxxxxxxxxxxx"  //�򿪵�������Ҫ���ϵĲ���
'}
'˵�������������ֻ��ֻ������Ӵ��壬���øô��岻�ܹر�(������Ƕ�״���)
    Dim objPopup As New clsPopup
    Dim objFrom As Object

    If strParam = "" Then Exit Sub
    Call WriteBusinessLOG("�����¼�", mstrUrl, strParam)
    If AnalysisJavaScriptEvent(strParam, objPopup) = False Then Exit Sub
    If UCase(objPopup.PopupType) = UCase("ShowDialog") Then 'ShowDialog����Ϊ�Ǵ��Ӵ���
        If frmChild.ShowMe(Me, objPopup) = True Then
            Call WriteBusinessLOG("������ˢ��", mstrUrl, mstrParam)
            Call Me.WBMethod.Refresh
        End If
    ElseIf UCase(objPopup.PopupType) = UCase("LOG") Then
        Call WriteBusinessLOG("�����¼�LOG", objPopup.PopupParentUrl, objPopup.PopupData)
    End If
End Sub

Private Function M_Dom_onclick() As Boolean
'�����ҳ�κεط����ᴥ�����¼���ֻʶ���ʽΪ��<span id="sendToZlhis" style="display:none;">����д��������</span>
    Dim strPane As String
    Dim objXML As New DOMDocument
    Dim objNodeList As IXMLDOMNodeList
    Dim strPram As String
    
    strPane = M_Dom.parentWindow.event.srcElement.outerHTML
    'MsgBox strPane
    If strPane Like "<span id=""sendToZlhis""*" Then
        If objXML.loadXML(strPane) = False Then
            MsgBox "������Ч��XML��ʽ��" & strPane, vbInformation, gstrSysName
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
'���ܣ�����ע���
      '��������Registry   Access   Functions
      Dim mREG As New REGTool5.Registry
      Dim ret As Boolean
      ret = mREG.UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet   Explorer\Main ", "Disable   Script   Debugger ", "yes")
      ret = mREG.UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet   Explorer\Main ", "DisableScriptDebuggerIE ", "yes")
End Sub
