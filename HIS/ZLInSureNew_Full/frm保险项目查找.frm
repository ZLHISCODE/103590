VERSION 5.00
Begin VB.Form frm保险项目查找 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保项目查找"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "frm保险项目查找.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra匹配 
      Caption         =   "匹配方式"
      Height          =   1005
      Left            =   2820
      TabIndex        =   15
      Top             =   1530
      Width           =   1845
      Begin VB.OptionButton optMatch 
         Caption         =   "从左匹配"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "任意匹配"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   660
         Width           =   1365
      End
   End
   Begin VB.Frame fra范围 
      Caption         =   "查找范围"
      Height          =   1005
      Left            =   120
      TabIndex        =   12
      Top             =   1530
      Width           =   2235
      Begin VB.OptionButton optClass 
         Caption         =   "所有收费类别"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   660
         Width           =   1875
      End
      Begin VB.OptionButton optClass 
         Caption         =   "当前类别"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   330
         Value           =   -1  'True
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   4950
      TabIndex        =   11
      Top             =   1620
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   4950
      TabIndex        =   10
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   4950
      TabIndex        =   9
      Top             =   210
      Width           =   1100
   End
   Begin VB.Frame fra条件 
      Caption         =   "查找条件"
      Height          =   1305
      Left            =   90
      TabIndex        =   18
      Top             =   120
      Width           =   4560
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3180
         MaxLength       =   255
         TabIndex        =   7
         Top             =   750
         Width           =   1185
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   870
         MaxLength       =   255
         TabIndex        =   1
         Top             =   330
         Width           =   1035
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   3180
         MaxLength       =   255
         TabIndex        =   3
         Top             =   330
         Width           =   1185
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   870
         MaxLength       =   255
         TabIndex        =   5
         Top             =   750
         Width           =   1035
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "医保编码(&Y)"
         Height          =   180
         Index           =   3
         Left            =   2130
         TabIndex        =   6
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "编码(&C)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   810
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   2490
         TabIndex        =   2
         Top             =   390
         Width           =   630
      End
   End
   Begin VB.Label lbl结果 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 请输入查找条件"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2670
      Width           =   5955
   End
End
Attribute VB_Name = "frm保险项目查找"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsFind As New ADODB.Recordset
Dim mint险类 As Integer
Dim mblnHIS10 As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim nod As Node
    '取出当前类别
    mblnHIS10 = IsZLHIS10
    Set nod = frm保险项目.tvwMain_S.SelectedItem
    With frm保险项目.cmb险类
        mint险类 = .ItemData(.ListIndex)
    End With
    
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    optClass(0).Caption = nod.Text
    optClass(0).Tag = Mid(nod.Key, 2, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    Set mrsFind = Nothing
End Sub

Private Sub cmdFind_Click()
    Dim str医保支付项目 As String
    
    If mrsFind.State = 1 Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocateItem
        Exit Sub
    End If
    If IsValid = False Then Exit Sub
    gstrSQL = ""
    If txtEdit(0).Text <> "" Then
        gstrSQL = "upper(A.编码) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(0).Text) & "%' or "
    End If
    If txtEdit(1).Text <> "" Then
        gstrSQL = gstrSQL & "upper(B.名称) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(1).Text) & "%' or "
    End If
    If txtEdit(2).Text <> "" Then
        gstrSQL = gstrSQL & "upper(B.简码) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(2).Text) & "%' or "
    End If
    If txtEdit(3).Text <> "" Then
        str医保支付项目 = "(select 收费细目ID,险类 from 保险支付项目 where upper(项目编码) like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(3).Text) & "%')"
    Else
        str医保支付项目 = "保险支付项目"
    End If
    If gstrSQL <> "" Or txtEdit(3).Text <> "" Then
        If gstrSQL <> "" Then
            gstrSQL = " and (" & Mid(gstrSQL, 1, Len(gstrSQL) - 4) & ") "
        End If
    Else
        MsgBox "请输入查找条件。", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        Exit Sub
    End If
    
    gstrSQL = "select distinct A.类别,A.上级ID,A.ID,A.名称 " & _
               " from 收费细目 A,收费别名 B," & str医保支付项目 & " C " & _
               " where A.ID =B.收费细目ID and A.末级=1 " & gstrSQL & _
               IIf(optClass(1).Value = True, "", "and A.类别='" & optClass(0).Tag & "'") & _
                " and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','yyyy-mm-dd')) " & _
                " and A.ID=C.收费细目ID" & IIf(txtEdit(3).Text <> "", "", "(+)") & " and C.险类(+)=" & mint险类
    Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Call LocateItem
End Sub

Private Sub LocateItem()
    Dim rsTemp As New ADODB.Recordset
    Dim lngID As Long
    Dim lngCount As Long
    Dim str材质 As String
    
    If mrsFind.RecordCount = 0 Then
        lbl结果.Caption = " 没有找到合适的收费细目"
        Beep
        Exit Sub
    End If
    If mrsFind.EOF = True Then
        lbl结果.Caption = " 已经定位完所有找的收费细目，请重新输入条件"
        Beep
        Exit Sub
    End If
    lbl结果.Caption = "           找到" & mrsFind.RecordCount & "条符合要求的收费细目。" & vbCrLf & "当前是第" & mrsFind.AbsolutePosition & "条，名称为：" & mrsFind("名称")
    
    With frm保险项目.tvwMain_S
        lngID = mrsFind("ID")
        If mrsFind!类别 = "4" And mblnHIS10 Then
            gstrSQL = "Select B.分类ID " & _
                      " From 收费项目目录 A, 诊疗项目目录 B, 材料特性 C " & _
                      " Where A.ID = C.材料ID " & _
                      " And B.ID = C.诊疗ID and A.ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            If rsTemp.EOF Then
                Exit Sub
            End If
            
            If IsNull(rsTemp!分类ID) Then
                .Nodes("G4").Selected = True
            Else
                .Nodes("G4" & rsTemp("分类ID")).Selected = True
            End If
        ElseIf mrsFind!类别 = "K" And mblnHIS10 Then
            gstrSQL = "Select B.分类ID " & _
                      " From 诊疗项目目录 B, 血液规格 C " & _
                      " Where C.品种ID=B.ID and C.规格ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            If rsTemp.EOF Then
                Exit Sub
            End If
            
            If IsNull(rsTemp!分类ID) Then
                .Nodes("X8").Selected = True
            Else
                .Nodes("X8" & rsTemp("分类ID")).Selected = True
            End If
        ElseIf mrsFind("类别") = "5" Or mrsFind("类别") = "6" Or mrsFind("类别") = "7" Then
            '这条细目是药品,其定位要复杂一些
            gstrSQL = "select B.材质分类,B.用途分类ID from 药品目录 A,药品信息 B " & _
                      " Where A.药名ID = B.药名ID And A.药品ID =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
            
            If rsTemp.EOF Then
                Exit Sub
            End If
            
            Select Case rsTemp("材质分类")
                Case "中成药"
                    str材质 = "E6"
                Case "中草药"
                    str材质 = "F7"
                Case Else
                    str材质 = "D5"
            End Select
                    
            If IsNull(rsTemp("用途分类ID")) Then
                .Nodes(str材质).Selected = True
            Else
                .Nodes(str材质 & rsTemp("用途分类ID")).Selected = True
            End If
            
        Else
            If mblnHIS10 Then
                gstrSQL = " Select ID,分类ID From 收费项目目录 Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
                
                If rsTemp.EOF Then Exit Sub
                
                .Nodes("CA" & rsTemp("分类ID")).Selected = True
            Else
                If IsNull(mrsFind("上级ID")) Then
                    .Nodes("R" & mrsFind("类别")).Selected = True
                Else
                    .Nodes("C" & mrsFind("类别") & mrsFind("上级ID")).Selected = True
                End If
            End If
        End If
        .SelectedItem.EnsureVisible
    End With
    frm保险项目.FillSum
        
    lngID = mrsFind("ID")
    With frm保险项目.mshSum_S
        For lngCount = 1 To .Rows - 1
            If .RowData(lngCount) = lngID Then
                .Row = lngCount
                .msfObj.TopRow = lngCount
                Exit Sub
            End If
        Next
    End With
    MsgBox "收费细目“" & mrsFind("名称") & "”的价格还未设置，或已经过期了。", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
'功能:分析输入有关费别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = txtEdit.LBound To txtEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If InStr(strTemp, "'") > 0 Then
            MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            zlControl.TxtSelAll txtEdit(i)
            Exit Function
        End If
    Next
    IsValid = True
End Function

Private Sub optClass_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optClass_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub optMatch_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optMatch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    lbl结果.Caption = "  条件已改变，请重新定位"
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 1 Then
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub
