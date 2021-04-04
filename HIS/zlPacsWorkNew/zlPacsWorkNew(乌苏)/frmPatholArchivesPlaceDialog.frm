VERSION 5.00
Begin VB.Form frmPatholArchivesPlaceDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "存放位置选择"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   Icon            =   "frmPatholArchivesPlaceDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消&C)"
      Height          =   400
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定&S)"
      Height          =   400
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cbxRoom 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox cbxBox 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   720
         Width           =   3375
      End
      Begin VB.ComboBox cbxDrawer 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "所属抽屉："
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1275
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "所属柜号："
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "所属房间："
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   280
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholArchivesPlaceDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsRoom As ADODB.Recordset
Private mrsBox As ADODB.Recordset
Private mrsDrawer As ADODB.Recordset

Public Room As String
Public Box As String
Public Drawer As String
Public IsOk As Boolean



Public Sub ShowPlaceDialog(ByVal strRoom As String, ByVal strBox As String, strDrawer As String, owner As Object)
On Error GoTo errHandle
    Room = ""
    Box = ""
    Drawer = ""
    IsOk = False
    
    
    Call LoadPlaceFilterData
    Call ConfigFilterInput(True, True)

    cbxRoom.Text = strRoom
    cbxBox.Text = strBox
    cbxDrawer.Text = strDrawer
        
    
    Me.Show 1, owner
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadPlaceFilterData()
'载入位置过滤数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset


    '读取已经存在的房间
    strSql = "select /*+ Rule*/ distinct 所属房间 from 病理档案信息 where 创建日期 between sysdate - 365 and sysdate order by 所属房间 "
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set mrsRoom = zlDatabase.CopyNewRec(rsData)


    '读取已经存在的柜号
    strSql = "select /*+ Rule*/ distinct 所属房间,所属柜号 from 病理档案信息 where 创建日期 between sysdate - 365 and sysdate order by 所属房间,所属柜号"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set mrsBox = zlDatabase.CopyNewRec(rsData)


    '读取已经存在的抽屉
    strSql = "select /*+ Rule*/ distinct 所属房间,所属柜号,所属抽屉 from 病理档案信息 where 创建日期 between sysdate - 365 and sysdate order by 所属房间,所属柜号,所属抽屉"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set mrsDrawer = zlDatabase.CopyNewRec(rsData)

End Sub


Private Sub ConfigFilterInput(ByVal blnRefreshRoom As Boolean, ByVal blnRefreshBox As Boolean)
'配置过滤的录入
    Dim strFilterRoom As String
    Dim strFilterBox As String
    Dim strFilterDrawer As String

    strFilterRoom = ""
    strFilterBox = ""
    strFilterDrawer = ""

    If cbxRoom.Text <> "" Then
        strFilterRoom = " 所属房间='" & cbxRoom.Text & "'"
        strFilterBox = " 所属房间='" & cbxRoom.Text & "'"
        strFilterDrawer = " 所属房间='" & cbxRoom.Text & "'"
    End If

    If cbxBox.Text <> "" And Not blnRefreshBox Then
        If strFilterBox <> "" Then
            strFilterBox = strFilterBox & " and 所属柜号='" & cbxBox.Text & "'"
        Else
            strFilterBox = " 所属柜号='" & cbxBox.Text & "'"
        End If


        If strFilterDrawer <> "" Then
            strFilterDrawer = strFilterDrawer & " and 所属柜号='" & cbxBox.Text & "'"
        Else
            strFilterDrawer = " 所属柜号='" & cbxBox.Text & "'"
        End If
    End If


    mrsBox.Filter = strFilterRoom
    mrsDrawer.Filter = strFilterBox


    If blnRefreshRoom Then Call ConfigRoomInput(mrsRoom)
    If blnRefreshBox Then Call ConfigBoxInput(mrsBox)
    
    Call ConfigDrawerInput(mrsDrawer)

End Sub



Private Sub ConfigRoomInput(rsData As ADODB.Recordset)
'所属房间
    Dim strTemp As String

    cbxRoom.Clear

    If rsData.RecordCount <= 0 Then Exit Sub

    Call cbxRoom.AddItem("")

    strTemp = "|"

    rsData.MoveFirst
    While Not rsData.EOF
        If Nvl(rsData!所属房间) <> "" Then
            If InStr(UCase(strTemp), "|" & UCase(Nvl(rsData!所属房间))) <= 0 Then

                If strTemp <> "|" Then strTemp = strTemp & "|"
                Call cbxRoom.AddItem(Nvl(rsData!所属房间))

            End If

        End If
        rsData.MoveNext
    Wend
End Sub


Private Sub ConfigBoxInput(rsData As ADODB.Recordset)
'所属柜号
    Dim strTemp As String

    cbxBox.Clear

    If rsData.RecordCount <= 0 Then Exit Sub

    Call cbxBox.AddItem("")

    strTemp = "|"

    rsData.MoveFirst
    While Not rsData.EOF
        If Nvl(rsData!所属柜号) <> "" Then
            If InStr(UCase(strTemp), "|" & UCase(Nvl(rsData!所属柜号))) <= 0 Then
                If strTemp <> "|" Then strTemp = strTemp & "|"
                
                strTemp = strTemp & Nvl(rsData!所属柜号)
                Call cbxBox.AddItem(Nvl(rsData!所属柜号))
            End If

        End If
        rsData.MoveNext
    Wend
End Sub



Private Sub ConfigDrawerInput(rsData As ADODB.Recordset)
'所属抽屉
    Dim strTemp As String

    cbxDrawer.Clear

    If rsData.RecordCount <= 0 Then Exit Sub

    Call cbxDrawer.AddItem("")

    strTemp = "|"

    rsData.MoveFirst
    While Not rsData.EOF
        If Nvl(rsData!所属抽屉) <> "" Then
            If InStr(UCase(strTemp), "|" & UCase(Nvl(rsData!所属抽屉))) <= 0 Then
                If strTemp <> "|" Then strTemp = strTemp & "|"
                
                strTemp = strTemp & Nvl(rsData!所属抽屉)
                Call cbxDrawer.AddItem(Nvl(rsData!所属抽屉))
            End If

        End If
        rsData.MoveNext
    Wend
End Sub




Private Sub cbxBox_Click()
On Error GoTo errHandle
    If Not cbxBox.Visible Then Exit Sub
    
    Call ConfigFilterInput(False, False)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbxRoom_Click()
On Error GoTo errHandle
    If Not cbxRoom.Visible Then Exit Sub
    
    Call ConfigFilterInput(False, True)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Room = ""
    Box = ""
    Drawer = ""
    
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    If cbxDrawer.Text = "" Then
        Call MsgBoxD(Me, "请选择档案的存放位置。", vbOKOnly, Me.Caption)
        Call cbxDrawer.SetFocus
        
        Exit Sub
    End If
    
    Room = cbxRoom.Text
    Box = cbxBox.Text
    Drawer = cbxDrawer.Text
    IsOk = True
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
err.Clear
End Sub
