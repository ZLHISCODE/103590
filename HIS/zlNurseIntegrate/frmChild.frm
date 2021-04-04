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
   StartUpPosition =   2  '��Ļ����
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
'���ܣ��������廤��༭�Ӵ���
'������objParent�����ø�����
'          strParam�����ص��Ӵ��������ַ���
'          strLessionID�����廤����ID (ͨ��RelatedIDToGUID�ӿڻ�ȡ)
'          strUserID�������û�ID (ͨ��RelatedIDToGUID�ӿڻ�ȡ)
'�����Ƿ���Ҫˢ����Ҫ����
    Dim strUrl As String, strParams As String

    mblnRefresh = False
    
    Me.Width = (objPopup.PopupWidth / 100) * Screen.Width
    Me.Height = (objPopup.PopupHeight / 100) * Screen.Height
    Me.Caption = objPopup.PopupTitle

    '������URL����
    'LessionID; ����ID(ͨ��RelatedIDToGUID�ӿڻ�ȡ)
    'UserID; �û�ID(ͨ��RelatedIDToGUID�ӿڻ�ȡ)
    'From; his(�̶���������, �����zlhis��ҳ��)
    'Data; �ò�������ocx���������õ������ݵ�Data�ֶ�
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
    SetProcessWorkingSetSize GetCurrentProcess, -1, -1  '���õ�ǰ
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
'�����¼�������
'{
'  type: "CloseDialog" || "ShowDialog""LOG", // CloseDialog �رյ���  ShowDialog �򿪵���  LOGд����־����
'  moduleUrl: "/shiftReport", //����Url
'  title: "���౨��",
'  width: "100" || null,
'  height: "100" || null,
'  minimal: true,  //���
'  max: false,     //��С��
'  isRefresh: true  //�Ƿ�ˢ�¸�����
'  data: "xxxxxxxxxxxx"  //�򿪵�������Ҫ���ϵĲ���
'}
    Dim objPopup As New clsPopup
    Dim objFrom As Object

    If strParam = "" Then Exit Sub
    Call WriteBusinessLOG("�����¼�", mstrUrl, strParam)
    If AnalysisJavaScriptEvent(strParam, objPopup) = False Then Exit Sub
    If UCase(objPopup.PopupType) = UCase("ShowDialog") Then 'ShowDialog����Ϊ�Ǵ��Ӵ���
        Set objFrom = New frmChild
        If objFrom.ShowMe(Me, objPopup) Then
            Me.WBMethod.Refresh
        End If
    ElseIf UCase(objPopup.PopupType) = UCase("CloseDialog") Then 'CloseDialog����Ϊ�ǹر��Ӵ���
        mblnRefresh = objPopup.PopupIsRefresh
        Unload Me
    ElseIf UCase(objPopup.PopupType) = UCase("LOG") Then
        Call WriteBusinessLOG("�����¼�LOG", objPopup.PopupParentUrl, objPopup.PopupData)
    ElseIf UCase(objPopup.PopupType) = UCase("CloseRefresh") Then '127510����ӡˢ�����б�����
        mblnRefresh = True
    Else
        mblnRefresh = objPopup.PopupIsRefresh
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


