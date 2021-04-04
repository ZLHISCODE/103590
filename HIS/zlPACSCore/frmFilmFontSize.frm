VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFilmFontSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "胶片字体设置"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8070
   Icon            =   "frmFilmFontSize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdAdd 
      Caption         =   "增加(&A)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "修改(&M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1470
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "删除(&D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2790
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&Q)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6840
      TabIndex        =   7
      Top             =   4680
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2820
      Width           =   7875
      Begin VB.CheckBox chkFontTransparent 
         Caption         =   "字体透明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   16
         Top             =   233
         Width           =   1575
      End
      Begin VB.CheckBox chkFontShadow 
         Caption         =   "字体阴影"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   15
         Top             =   593
         Width           =   2055
      End
      Begin VB.CheckBox chkFontInverse 
         Caption         =   "字体反色"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   998
         Width           =   1455
      End
      Begin VB.TextBox txtPostureFontSize 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2070
         TabIndex        =   9
         Top             =   968
         Width           =   1665
      End
      Begin VB.CheckBox chkPostureAutoZoom 
         Caption         =   "体位标注随图像缩放"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CheckBox ChkAutoZoom 
         Caption         =   "信息随图像缩放"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxtFontSize 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2070
         TabIndex        =   5
         Top             =   578
         Width           =   1665
      End
      Begin VB.TextBox txtImageType 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         Top             =   218
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "体位标注字体大小:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   1020
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "信息字体大小:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   4
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "影像类型:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   945
      End
   End
   Begin MSComctlLib.ListView LivMain 
      Height          =   2715
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmFilmFontSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CmdAdd_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    '检查输入有效性
    If Len(Trim(Me.txtImageType)) < 1 Then
        MsgBox "请输入影像类别", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.TxtFontSize)) < 1 Then
        MsgBox "请输入打印的信息字体大小", vbInformation, gstrSysName
        Me.TxtFontSize.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.txtPostureFontSize)) < 1 Then
        MsgBox "请输入打印的体位标注字体大小", vbInformation, gstrSysName
        Me.txtPostureFontSize.SetFocus
        Exit Sub
    End If
    On Error GoTo errh
    
    '查询确认影像类别是否已经存在
    If blLocalRun = True Then
        strSQL = "select count(*) as 总计 from DICOM胶片打印字体 where 影像类别 = """ & Me.txtImageType & """"
        Set rsTmp = cnAccess.Execute(strSQL)
    Else
        strSQL = "select count(*) as 总计 from 影像胶片打印字体 where 影像类别 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(UCase(Me.txtImageType)))
    End If
    If rsTmp("总计") > 0 Then
        MsgBox "您增加的影像类别已存在！请重新输入！", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    
    '插入新的字体设置
    If blLocalRun = True Then
        strSQL = "insert into DICOM胶片打印字体 (影像类别,字体大小,是否随图像缩放,体位标注字体大小,体位标注随图像缩放) values (""" & _
                 UCase(Me.txtImageType) & """,""" & Me.TxtFontSize & """," & IIf(Me.ChkAutoZoom.Value = 1, True, False) & _
                 ",""" & Me.txtPostureFontSize & """," & IIf(Me.chkPostureAutoZoom.Value = 1, True, False) & ")"
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_影像胶片打印字体_INSERT('" & UCase(Me.txtImageType) & "'," & Me.TxtFontSize & _
                 "," & IIf(Me.ChkAutoZoom.Value = 1, 1, 0) & ",'" & Me.txtPostureFontSize & "'," & _
                 IIf(Me.chkPostureAutoZoom.Value = 1, 1, 0) & "," & IIf(Me.chkFontInverse.Value = 1, 1, 0) & _
                 "," & IIf(Me.chkFontShadow.Value = 1, 1, 0) & "," & IIf(Me.chkFontTransparent.Value = 1, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    '清空录入信息，设置默认值
    Me.txtImageType.Text = ""
    Me.TxtFontSize.Text = ""
    Me.ChkAutoZoom.Value = 0
    Me.txtPostureFontSize.Text = ""
    Me.chkPostureAutoZoom.Value = 1
    Me.chkFontInverse.Value = 0
    Me.chkFontShadow = 0
    Me.chkFontTransparent.Value = 1
    
    '列表重新显示
    LoadDate
    Me.LivMain.ListItems(Me.LivMain.ListItems.Count).Selected = True
    
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub CmdDelete_Click()
    Dim strSQL As String
    Dim i As Integer
    If Me.LivMain.ListItems.Count < 1 Then Exit Sub
    If Len(Trim(Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text)) < 1 Then
        MsgBox "请选中一个要删除的影像类别！", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errh
    i = Me.LivMain.SelectedItem.Index
    
    If blLocalRun = True Then
        strSQL = "delete from DICOM胶片打印字体 where 影像类别 = '" & Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & "'"
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_影像胶片打印字体_DELETE('" & Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    Me.txtImageType = ""
    Me.TxtFontSize = ""
    Me.ChkAutoZoom.Value = 0
    Me.txtPostureFontSize.Text = ""
    Me.chkPostureAutoZoom.Value = 1
    Me.chkFontInverse.Value = 0
    Me.chkFontShadow = 0
    Me.chkFontTransparent.Value = 1
    
    LoadDate
    If i > 1 Then
        Me.LivMain.ListItems(i - 1).Selected = True
    End If
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdModify_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    '检查输入有效性
    If Me.LivMain.ListItems.Count < 1 Then Exit Sub
    If Len(Trim(Me.txtImageType)) < 1 Then
        MsgBox "请选择一个影像类别", vbInformation, gstrSysName
        Me.txtImageType.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.TxtFontSize)) < 1 Then
        MsgBox "请输入打印的信息字体大小", vbInformation, gstrSysName
        Me.TxtFontSize.SetFocus
        Exit Sub
    End If
    If Len(Trim(Me.txtPostureFontSize)) < 1 Then
        MsgBox "请输入打印的体位标注字体大小", vbInformation, gstrSysName
        Me.txtPostureFontSize.SetFocus
        Exit Sub
    End If
    On Error GoTo errh
    
    '修改胶片字体信息
    If blLocalRun = True Then
        strSQL = "update DICOM胶片打印字体 set 影像类别 = '" & UCase(Me.txtImageType) & "',字体大小 ='" & TxtFontSize & _
                 "',是否随图像缩放 = " & IIf(Me.ChkAutoZoom.Value = 1, True, False) & ",体位标注字体大小 = '" & txtPostureFontSize & _
                 "',体位标注随图像缩放 = " & IIf(Me.chkPostureAutoZoom.Value = 1, True, False) & " where 影像类别 = '" & _
                 Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & "'"
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_影像胶片打印字体_UPDATE('" & Me.LivMain.ListItems(Me.LivMain.SelectedItem.Index).Text & _
                 "','" & UCase(Me.txtImageType) & "'," & Me.TxtFontSize & "," & IIf(Me.ChkAutoZoom.Value = 1, 1, 0) & _
                 ",'" & Me.txtPostureFontSize & "'," & IIf(Me.chkPostureAutoZoom.Value = 1, 1, 0) & "," & _
                  IIf(Me.chkFontInverse.Value = 1, 1, 0) & "," & IIf(Me.chkFontShadow.Value = 1, 1, 0) & "," & _
                  IIf(Me.chkFontTransparent.Value = 1, 1, 0) & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    i = Me.LivMain.SelectedItem.Index
    LoadDate
    Me.LivMain.ListItems(i).Selected = True
    Exit Sub
errh:
    If blLocalRun = True Then
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub Form_Load()
    InitLivHead
    LoadDate
End Sub
Sub InitLivHead()
    Dim chColHeader As ColumnHeader
    '初使化列表头
    With Me.LivMain
        Set chColHeader = .ColumnHeaders.Add(, "A", "影像类别")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "B", "信息字体")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "C", "信息随图缩放")
        chColHeader.width = 1200
        Set chColHeader = .ColumnHeaders.Add(, "D", "体位字体")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "E", "体位随图缩放")
        chColHeader.width = 1200
        Set chColHeader = .ColumnHeaders.Add(, "F", "字体透明")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "G", "字体阴影")
        chColHeader.width = 900
        Set chColHeader = .ColumnHeaders.Add(, "H", "字体反色")
        chColHeader.width = 900
    End With
End Sub
Sub LoadDate()
    '读入数据
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objItem As ListItem
    
    Me.LivMain.ListItems.Clear
    
    If blLocalRun = True Then
        strSQL = "select 影像类别 , 字体大小, 是否随图像缩放,体位标注字体大小,体位标注随图像缩放 from DICOM胶片打印字体"
        Set rsTmp = cnAccess.Execute(strSQL)
    Else
        strSQL = "select 影像类别 , 字体大小, 是否随图像缩放,体位标注字体大小,体位标注随图像缩放,字体反色,字体阴影,字体背景透明 from 影像胶片打印字体"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    Do Until rsTmp.EOF
        With Me.LivMain
            Set objItem = .ListItems.Add(, "A" & rsTmp("影像类别"), rsTmp("影像类别"))
            objItem.SubItems(1) = Nvl(rsTmp("字体大小"))
            objItem.SubItems(2) = IIf(Nvl(rsTmp("是否随图像缩放"), 0), "√", "")
            objItem.SubItems(3) = Nvl(rsTmp("体位标注字体大小"))
            objItem.SubItems(4) = IIf(Nvl(rsTmp("体位标注随图像缩放"), 0), "√", "")
            objItem.SubItems(5) = IIf(Nvl(rsTmp("字体背景透明"), 1), "√", "")
            objItem.SubItems(6) = IIf(Nvl(rsTmp("字体阴影"), 0), "√", "")
            objItem.SubItems(7) = IIf(Nvl(rsTmp("字体反色"), 0), "√", "")
        End With
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    If LivMain.ListItems.Count >= 1 Then
        Call LivMain_ItemClick(LivMain.ListItems(1))
    End If
End Sub

Private Sub LivMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtImageType = Item.Text
    Me.TxtFontSize = Item.SubItems(1)
    Me.ChkAutoZoom.Value = IIf(Item.SubItems(2) <> "", 1, 0)
    Me.txtPostureFontSize = Item.SubItems(3)
    Me.chkPostureAutoZoom.Value = IIf(Item.SubItems(4) <> "", 1, 0)
    Me.chkFontTransparent.Value = IIf(Item.SubItems(5) <> "", 1, 0)
    Me.chkFontShadow.Value = IIf(Item.SubItems(6) <> "", 1, 0)
    Me.chkFontInverse.Value = IIf(Item.SubItems(7) <> "", 1, 0)
End Sub

Private Sub TxtFontSize_GotFocus()
    Me.TxtFontSize.SelStart = 0
    Me.TxtFontSize.SelLength = Len(Me.TxtFontSize)
End Sub

Private Sub TxtFontSize_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtImageType_GotFocus()
    Me.txtImageType.SelStart = 0
    Me.txtImageType.SelLength = Len(Me.txtImageType)
End Sub

Private Sub txtPostureFontSize_GotFocus()
    Me.txtPostureFontSize.SelStart = 0
    Me.txtPostureFontSize.SelLength = Len(Me.txtPostureFontSize)
End Sub

Private Sub txtPostureFontSize_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
