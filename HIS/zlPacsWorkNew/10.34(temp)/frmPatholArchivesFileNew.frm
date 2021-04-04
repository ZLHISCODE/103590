VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholArchivesFileNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增档案"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "frmPatholArchivesFileNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   6975
      TabIndex        =   28
      Top             =   4200
      Width           =   6975
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   120
         Width           =   6735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   6975
      End
   End
   Begin VB.CheckBox chkContinue 
      Caption         =   "确定后继续执行当前操作"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "确 定&S)"
      Height          =   400
      Left            =   4560
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   5880
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtPlace 
         Height          =   300
         Left            =   4440
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cbxDrawer 
         Height          =   300
         Left            =   960
         TabIndex        =   27
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cbxBox 
         Height          =   300
         Left            =   4440
         TabIndex        =   26
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox cbxRoom 
         Height          =   300
         Left            =   960
         TabIndex        =   25
         Top             =   1680
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpArchivesCreate 
         Height          =   300
         Left            =   4440
         TabIndex        =   21
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   154796035
         CurrentDate     =   40865
      End
      Begin VB.TextBox txtCreateUser 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   960
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesDescription 
         Height          =   300
         Left            =   960
         TabIndex        =   17
         Top             =   2640
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpArchivesEnd 
         Height          =   300
         Left            =   4440
         TabIndex        =   14
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   154796035
         CurrentDate     =   40864
      End
      Begin MSComCtl2.DTPicker dtpArchivesStart 
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   154796035
         CurrentDate     =   40864
      End
      Begin VB.ComboBox cbxArchivesClass 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesStudyArea 
         Height          =   300
         Left            =   960
         TabIndex        =   8
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesCode 
         Height          =   300
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtArchivesName 
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "详细地址："
         Height          =   255
         Left            =   3600
         TabIndex        =   33
         Top             =   2205
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "所属抽屉："
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2200
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "所属柜号："
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6760
         TabIndex        =   24
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label10 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3280
         TabIndex        =   22
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "创建日期："
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "创 建 人："
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "档案说明："
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2670
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "所属房间："
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "结束日期："
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "开始日期："
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "档案分类："
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "检查范围："
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   975
      End
      Begin VB.Label labArchivesCode 
         Caption         =   "档案编号："
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "档案名称："
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholArchivesFileNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentGrid As ucFlexGrid

Private mblnIsSucceed As Boolean
Private mblnIsUpdate As Boolean

Private mrsArchivesClass As ADODB.Recordset




Public Function ShowAddArchivesFileWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'显示新增文件档案窗口
    Dim curDate As Date
    
    ShowAddArchivesFileWindow = False
    
    Set mufgParentGrid = ufgParentGrid
    
    Me.Caption = "新增档案"
    mblnIsUpdate = False
    mblnIsSucceed = False
    
    curDate = zlDatabase.Currentdate
    
    dtpArchivesStart.value = curDate
    dtpArchivesEnd.value = curDate
    dtpArchivesCreate.value = curDate
    txtCreateUser.Text = UserInfo.姓名
    
    
    Call CloseProcessHint
    
    chkContinue.value = False
    chkContinue.Visible = True
    
    Call Me.Show(1, owner)
    
    ShowAddArchivesFileWindow = mblnIsSucceed

End Function



Public Function ShowUpdateArchivesFileWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'显示抗体更新窗口
    ShowUpdateArchivesFileWindow = False
    
    Set mufgParentGrid = ufgParentGrid
        
    Me.Caption = "更新档案"
    mblnIsUpdate = True
    mblnIsSucceed = False
        
    Call CloseProcessHint
    
    Call ConfigUpdateFace
    
    chkContinue.value = False
    chkContinue.Visible = False

    
    Call Me.Show(1, owner)
    
    ShowUpdateArchivesFileWindow = mblnIsSucceed
End Function



Public Sub ConfigUpdateFace()
On Error Resume Next
    Dim strPlace As String
    
    With mufgParentGrid
        txtArchivesName.Text = .Text(.SelectionRow, gstrPatholCol_档案名称)
        txtArchivesCode.Text = .Text(.SelectionRow, gstrPatholCol_档案编号)
        txtArchivesStudyArea.Text = .Text(.SelectionRow, gstrPatholCol_检查范围)
        dtpArchivesStart.value = .Text(.SelectionRow, gstrPatholCol_开始日期)
        dtpArchivesEnd.value = .Text(.SelectionRow, gstrPatholCol_结束日期)
        dtpArchivesCreate.value = .Text(.SelectionRow, gstrPatholCol_创建日期)
        txtCreateUser.Text = .Text(.SelectionRow, gstrPatholCol_创建人)
        txtArchivesDescription.Text = .Text(.SelectionRow, gstrPatholCol_档案说明)
        
        '读取档案分类
        cbxArchivesClass.Text = .Text(.SelectionRow, gstrPatholCol_档案分类)


        '读取存放位置
        cbxRoom.Text = .Text(.SelectionRow, gstrPatholCol_所属房间)
        cbxBox.Text = .Text(.SelectionRow, gstrPatholCol_所属柜号)
        cbxDrawer.Text = .Text(.SelectionRow, gstrPatholCol_所属抽屉)
        txtPlace.Text = .Text(.SelectionRow, gstrPatholCol_详细地址)
    End With
    
err.Clear
    
End Sub



Private Sub ShowProcessHint(ByVal strHint As String)
'显示处理信息
    txtShow.Text = strHint
End Sub


Private Sub CloseProcessHint()
'关闭处理提示
    txtShow.Text = ""
End Sub





Private Sub cmdCancel_Click()
On Error GoTo errHandle
    Call Unload(Me)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function CheckArchivesFileDataIsValid() As String
    CheckArchivesFileDataIsValid = ""
    
    '检查档案名称是否为空
    If Trim(txtArchivesName.Text) = "" Then
        CheckArchivesFileDataIsValid = "档案名称不能为空。"
        
        Call txtArchivesName.SetFocus
        Exit Function
    End If
    
    
    '检查档案分类是否为空
    If Trim(cbxArchivesClass.Text) = "" Then
        CheckArchivesFileDataIsValid = "档案分类不能为空。"
        
        Call cbxArchivesClass.SetFocus
        Exit Function
    End If
    
    
    
    '检查档案名称是否重复
    Dim i As Integer
    For i = 1 To mufgParentGrid.GridRows - 1
        If Not mufgParentGrid.RowHidden(i) Then
            If Not mblnIsUpdate Then
                If mufgParentGrid.Text(i, gstrPatholCol_档案名称) = txtArchivesName.Text Then
                    CheckArchivesFileDataIsValid = "档案名称重复。"
    
                    Call txtArchivesName.SetFocus
                    Exit Function
                End If
            Else
                If Not mufgParentGrid.SelectionRow = i Then
                    If mufgParentGrid.Text(i, gstrPatholCol_档案名称) = txtArchivesName.Text Then
                        CheckArchivesFileDataIsValid = "档案名称重复。"
    
                        Call txtArchivesName.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function



Private Sub NewArchivesInf()
'在数据库中新增档案记录
'返回档案ID
    Dim strSql As String
    Dim rsReture As ADODB.Recordset
    Dim lngNewRecordIndex As Long
    Dim lngNewArchivesId As Long
    
    

    strSql = "select Zl_病理档案_新增文件档案([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13]) as 返回值 from dual"
                                
    Set rsReture = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                txtArchivesName.Text, _
                                txtArchivesCode.Text, _
                                txtArchivesStudyArea.Text, _
                                Val(cbxArchivesClass.ItemData(cbxArchivesClass.ListIndex)), _
                                CDate(dtpArchivesStart.value), _
                                CDate(dtpArchivesEnd.value), _
                                CDate(dtpArchivesCreate.value), _
                                UserInfo.姓名, _
                                cbxRoom.Text, _
                                cbxBox.Text, _
                                cbxDrawer.Text, _
                                txtPlace.Text, _
                                txtArchivesDescription.Text)
                                
    If rsReture.RecordCount <= 0 Then
        Call err.Raise(0, "NewArchivesFile", "未成功获取新增后的档案ID,本次操作失败。")
        Exit Sub
    End If
    
    
    With mufgParentGrid
        lngNewRecordIndex = .NewRow
        
        .Text(lngNewRecordIndex, gstrPatholCol_ID) = Nvl(rsReture!返回值)
        .Text(lngNewRecordIndex, gstrPatholCol_档案名称) = txtArchivesName.Text
        .Text(lngNewRecordIndex, gstrPatholCol_档案编号) = txtArchivesCode.Text
        .Text(lngNewRecordIndex, gstrPatholCol_检查范围) = txtArchivesStudyArea.Text
        .Text(lngNewRecordIndex, gstrPatholCol_开始日期) = Format(dtpArchivesStart.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrPatholCol_结束日期) = Format(dtpArchivesEnd.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrPatholCol_创建日期) = Format(dtpArchivesCreate.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrPatholCol_档案说明) = txtArchivesDescription.Text
        .Text(lngNewRecordIndex, gstrPatholCol_创建人) = txtCreateUser.Text
        .Text(lngNewRecordIndex, gstrPatholCol_档案状态) = "未归档"
        .Text(lngNewRecordIndex, gstrPatholCol_所属房间) = cbxRoom.Text
        .Text(lngNewRecordIndex, gstrPatholCol_所属柜号) = cbxBox.Text
        .Text(lngNewRecordIndex, gstrPatholCol_所属抽屉) = cbxDrawer.Text
        .Text(lngNewRecordIndex, gstrPatholCol_详细地址) = txtPlace.Text
        
        .Text(lngNewRecordIndex, gstrPatholCol_档案分类) = cbxArchivesClass.Text
        
        mrsArchivesClass.Filter = "分类名称='" & cbxArchivesClass.Text & "'"
        If mrsArchivesClass.RecordCount > 0 Then
            .Text(lngNewRecordIndex, gstrPatholCol_材料类型) = Val(Nvl(mrsArchivesClass!材料类型))
            .Text(lngNewRecordIndex, gstrPatholCol_报表名称) = Nvl(mrsArchivesClass!报表名称)
        End If
        
        Call .LocateRow(lngNewRecordIndex)
        
    End With
End Sub




Private Sub UpdateArchivesInf()
'更新数据库中的档案信息
    Dim strSql As String
    Dim lngCurArchivesId As Long
    Dim lngUpdateRecordIndex As Long
    


    lngCurArchivesId = mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow)
    
    strSql = "Zl_病理档案_更新文件档案(" & lngCurArchivesId & ",'" & txtArchivesName.Text & "','" & txtArchivesCode.Text & "','" & txtArchivesStudyArea.Text & "'," & _
                                cbxArchivesClass.ItemData(cbxArchivesClass.ListIndex) & "," & To_Date(dtpArchivesStart.value) & "," & _
                                To_Date(dtpArchivesEnd.value) & ",'" & cbxRoom.Text & "','" & cbxBox.Text & "','" & cbxDrawer.Text & "','" & txtPlace.Text & "','" & txtArchivesDescription.Text & "')"
                                
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    
    lngUpdateRecordIndex = mufgParentGrid.SelectionRow
    
    With mufgParentGrid
        .Text(lngUpdateRecordIndex, gstrPatholCol_档案名称) = txtArchivesName.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_档案编号) = txtArchivesCode.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_检查范围) = txtArchivesStudyArea.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_开始日期) = Format(dtpArchivesStart.value, gstrDateFormat)
        .Text(lngUpdateRecordIndex, gstrPatholCol_结束日期) = Format(dtpArchivesEnd.value, gstrDateFormat)
        .Text(lngUpdateRecordIndex, gstrPatholCol_档案说明) = txtArchivesDescription.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_所属房间) = cbxRoom.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_所属柜号) = cbxBox.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_所属抽屉) = cbxDrawer.Text
        .Text(lngUpdateRecordIndex, gstrPatholCol_详细地址) = txtPlace.Text
        
        .Text(lngUpdateRecordIndex, gstrPatholCol_档案分类) = cbxArchivesClass.Text
        
        mrsArchivesClass.Filter = "分类名称='" & cbxArchivesClass.Text & "'"
        If mrsArchivesClass.RecordCount > 0 Then
            .Text(lngUpdateRecordIndex, gstrPatholCol_材料类型) = Val(Nvl(mrsArchivesClass!材料类型))
            .Text(lngUpdateRecordIndex, gstrPatholCol_报表名称) = Nvl(mrsArchivesClass!报表名称)
        End If
        
    End With
End Sub




Private Sub cmdSure_Click()
On Error GoTo errHandle
    Dim strErr As String
    Dim strNewArchivesId As String

    '检查是否录入有效数据
    strErr = CheckArchivesFileDataIsValid()
    If Trim(strErr) <> "" Then
        Call ShowProcessHint(strErr)
        Exit Sub
    End If
    
    
    If Not mblnIsUpdate Then
        '新增档案
        Call NewArchivesInf
        
        Call mufgParentGrid.LocateRow(mufgParentGrid.GridRows - 1)
    Else
        '更新档案
        Call UpdateArchivesInf
    End If
    
    mblnIsSucceed = True
    
    If Not CBool(chkContinue.value) Then
        Call Unload(Me)
    End If
    
    Call CloseProcessHint
Exit Sub
errHandle:
    Call ShowProcessHint(err.Description)
    err.Clear
End Sub

Private Sub Form_Initialize()
    mblnIsSucceed = False
    mblnIsUpdate = False
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    Call LoadArchivesClassData
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadArchivesClassData()
'加载档案分类数据
    Dim strSql As String
    
    strSql = "select ID,分类名称,材料类型,报表名称 from 病理档案分类"
    
    Set mrsArchivesClass = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    Call cbxArchivesClass.Clear
    If mrsArchivesClass.RecordCount <= 0 Then Exit Sub
    
    While Not mrsArchivesClass.EOF
        Call cbxArchivesClass.AddItem(Nvl(mrsArchivesClass!分类名称))
        
        cbxArchivesClass.ItemData(cbxArchivesClass.ListCount - 1) = Nvl(mrsArchivesClass!ID)
        
        mrsArchivesClass.MoveNext
    Wend
    
    If cbxArchivesClass.ListCount > 0 Then cbxArchivesClass.ListIndex = 0
End Sub






Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
err.Clear
End Sub

