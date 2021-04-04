VERSION 5.00
Begin VB.Form frmPatholLoseEnreg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "材料遗失"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3165
   Icon            =   "frmPatholLoseEnreg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2895
      Begin VB.TextBox txtCount 
         Height          =   300
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "1"
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label labCount 
         Caption         =   "遗失数量："
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholLoseEnreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsFind As Boolean
Private mlngMaterialArchivesId As Long
Private mufgParentGrid As ucFlexGrid

Public blnIsOk As Boolean


Public Sub ShowLoseWindow(ufgMaterialGrid As ucFlexGrid, owner As Object)
'显示材料遗失窗口
    Me.Caption = "材料遗失"
    labCount.Caption = "遗失数量："
    
    blnIsOk = False
    mblnIsFind = False
    Set mufgParentGrid = ufgMaterialGrid
    
    If Not ufgMaterialGrid.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行遗失处理的材料记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Me.Show(1, owner)
End Sub


Public Sub ShowFindWindow(ufgMaterialGrid As ucFlexGrid, owner As Object)
'显示材料找回窗口
    Me.Caption = "材料找回"
    labCount.Caption = "找回数量："
    
    blnIsOk = False
    mblnIsFind = True
    Set mufgParentGrid = ufgMaterialGrid
    
    If Not ufgMaterialGrid.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行找回处理的材料记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Me.Show(1, owner)
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    blnIsOk = False
    
    Call Me.Hide
err.Clear
End Sub

Private Sub MaterialFind()
'材料找回
    Dim lngMaterialArchivesId As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnFind As Boolean
    Dim strValue As String
    Dim chkState As CheckState
    
    lngMaterialArchivesId = Val(mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow))
    
    strSql = "select ZL_病理材料_材料找回([1],[2],[3]) as 返回值  from dual"
                                        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMaterialArchivesId, _
                                                                Val(txtCount.Text), _
                                                                UserInfo.姓名)
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "MaterialFind", "不能取得找回处理后的材料状态，操作失败。")
        Exit Sub
    End If

'    Call mufgParentGrid.GetFieldDisplayText(gstrPatholCol_存放状态, Nvl(rsData!返回值), blnFind, chkState, strValue)
'    Call mufgParentGrid.SetText(mufgParentGrid.SelectionRow, gstrPatholCol_存放状态, strValue, True)
    
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_存放状态, Nvl(rsData!返回值), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_在档数量, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_在档数量)) + Val(txtCount.Text), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_遗失数量, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_遗失数量)) - Val(txtCount.Text), True)
End Sub

Private Sub MaterialLose()
'材料遗失
    Dim lngMaterialArchivesId As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    
    lngMaterialArchivesId = Val(mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow))
    
    strSql = "select ZL_病理材料_材料遗失([1],[2],[3], [4]) as 返回值  from dual"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngMaterialArchivesId, _
                                                            Val(txtCount.Text), _
                                                            CDate(zlDatabase.Currentdate), _
                                                            UserInfo.姓名)
                                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "MaterialFind", "不能取得遗失处理后的材料状态，操作失败。")
        Exit Sub
    End If

'    Call mufgParentGrid.GetFieldConvertValue(gstrPatholCol_存放状态, Nvl(rsData!返回值), blnFind, chkState, strValue)
'    Call mufgParentGrid.SetText(mufgParentGrid.SelectRowIndex, gstrPatholCol_存放状态, strValue, True)
    
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_存放状态, Nvl(rsData!返回值), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_在档数量, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_在档数量)) - Val(txtCount.Text), True)
    Call mufgParentGrid.SyncData(mufgParentGrid.SelectionRow, gstrPatholCol_遗失数量, _
                                Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_遗失数量)) + Val(txtCount.Text), True)
    
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    If mblnIsFind Then
'        If mufgParentGrid.Text(mufgParentGrid.SelectRowIndex, gstrPatholCol_存放状态) = "存档中" Then
'            Call MsgBoxD(Me, "该材料没有遗失，不能进行找回处理。", vbOKOnly, Me.Caption)
'            Exit Sub
'        End If

        If Val(txtCount.Text) > Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_遗失数量)) Then
            Call MsgBoxD(Me, "材料找回数量不能大于遗失数量。", vbOKOnly, Me.Caption)
            Exit Sub
        End If
        
        
        '材料找回处理
        Call MaterialFind
    Else
'        If mufgParentGrid.Text(mufgParentGrid.SelectRowIndex, gstrPatholCol_存放状态) = "存档中" Then
'            Call MsgBoxD(Me, "该材料已经遗失，不能进行遗失处理。", vbOKOnly, Me.Caption)
'            Exit Sub
'        End If
        
        If Val(txtCount.Text) > Val(mufgParentGrid.Text(mufgParentGrid.SelectionRow, gstrPatholCol_在档数量)) Then
            Call MsgBoxD(Me, "材料遗失数量不能大于在档数量。", vbOKOnly, Me.Caption)
            Exit Sub
        End If
        
        
        '材料遗失处理
        Call MaterialLose
    End If
    
    blnIsOk = True
    
    Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call RestoreWinState(Me, App.ProductName)
    
err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
err.Clear
End Sub
