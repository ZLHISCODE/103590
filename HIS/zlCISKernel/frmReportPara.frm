VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReportPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "执行单打印参数设置"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6405
   Icon            =   "frmReportPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmrClose 
      Left            =   480
      Top             =   600
   End
   Begin VB.Timer tmrOpen 
      Left            =   0
      Top             =   600
   End
   Begin MSComctlLib.ListView lvwReport 
      Height          =   5640
      Left            =   1245
      TabIndex        =   0
      Top             =   600
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   9948
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "报表"
         Object.Width           =   6615
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行单种类"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   900
   End
End
Attribute VB_Name = "frmReportPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
'Const LWA_COLORKEY = &H1
Private lngAlpha As Integer
Private j As Integer
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As PointAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" _
    (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const ALTERNATE = 1
Private mblnCancle As Boolean
Private mlngX As Long
Private mlngY As Long
Private mlng病区ID As Long
Private mblnOK As Boolean

Public Function ShowMe(objParent As Object, ByVal lng病区ID As Long) As Boolean
    mlng病区ID = lng病区ID
    Me.Show 1, objParent
    ShowMe = mblnOK
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnOK = True Then SetPara
    If mblnCancle = False Then
        Cancel = 1
        lngAlpha = 250
        tmrClose.Enabled = True
    End If
End Sub

Private Sub lvwReport_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    mblnOK = True
End Sub

Private Sub tmrOpen_Timer()
    lngAlpha = lngAlpha + 10
     SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
     SetLayeredWindowAttributes Me.hwnd, 0, lngAlpha, LWA_ALPHA  '150为透明度(0-255)
     If lngAlpha >= 255 Then tmrOpen.Enabled = False
End Sub

Private Sub tmrClose_Timer()
    lngAlpha = lngAlpha - 10
     SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
     SetLayeredWindowAttributes Me.hwnd, 0, lngAlpha, LWA_ALPHA  '150为透明度(0-255)
     If lngAlpha <= 5 Then tmrClose.Enabled = False: mblnCancle = True: Unload Me
End Sub


Private Sub Form_Load()
    '设置淡出
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 0, LWA_ALPHA  '150为透明度(0-255)
    tmrClose.Interval = 10
    tmrOpen.Interval = 10
    tmrOpen.Enabled = True
    tmrClose.Enabled = False
    lngAlpha = 5
    mblnCancle = False
    mblnOK = False
    
    Call InitReports '读取报表
End Sub

Private Function InitReports() As Boolean
'功能：读取可用报表
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem
    Dim strReports As String
    
    On Error GoTo errH
    
    strReports = zlDatabase.GetPara("执行单可用报表", glngSys, p住院医嘱发送, , Array(Label8, lvwReport), InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱选项设置;") > 0, , mlng病区ID)
    strSQL = "Select A.ID,A.编号,A.名称,NVL(A.功能,B.功能) AS 功能,a.系统" & vbNewLine & _
            "From zlReports A, zlRPTPuts B" & vbNewLine & _
            "Where a.Id = b.报表id(+)  And" & vbNewLine & _
            "      (b.程序id = 1254 Or" & vbNewLine & _
            "       a.系统 = [1] And a.编号 In ('ZL1_INSIDE_1254_4', 'ZL1_INSIDE_1254_5', 'ZL1_INSIDE_1254_6', 'ZL1_INSIDE_1254_7', 'ZL1_INSIDE_1254_8'," & vbNewLine & _
            "                'ZL1_INSIDE_1254_9', 'ZL1_INSIDE_1254_10', 'ZL1_INSIDE_1254_11', 'ZL1_INSIDE_1254_12'," & vbNewLine & _
            "                'ZL1_INSIDE_1254_13', 'ZL1_INSIDE_1254_14', 'ZL1_INSIDE_1254_15', 'ZL1_INSIDE_1254_16'))" & vbNewLine & _
            "Order By a.Id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys)
    Do While Not rsTmp.EOF
        If InStr(GetInsidePrivs(p住院医嘱发送), ";" & rsTmp!功能 & ";") > 0 Then
            Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!编号, rsTmp!名称, , 1)
            objItem.Tag = Val(rsTmp!ID)
            If strReports = "" Or InStr("|" & strReports & "|", "|" & rsTmp!编号 & "|") > 0 Then
                objItem.Checked = True
            End If
        End If
        rsTmp.MoveNext
    Loop
    InitReports = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub SetPara()
    Dim strPara As String
    Dim i As Long, j As Long
    Dim strReports As String
    
    For i = 1 To lvwReport.ListItems.Count
        If lvwReport.ListItems.Item(i).Checked = True Then
           strReports = strReports & "|" & Mid(lvwReport.ListItems.Item(i).Key, 2)
        End If
    Next
    strReports = Mid(strReports, 2)
    Call zlDatabase.SetPara("执行单可用报表", strReports, glngSys, p住院医嘱发送, , mlng病区ID)
End Sub
