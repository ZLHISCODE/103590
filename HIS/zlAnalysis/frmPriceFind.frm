VERSION 5.00
Begin VB.Form frmPriceFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "价目查找"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   930
   ClientWidth     =   6225
   Icon            =   "frmPriceFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra匹配 
      Caption         =   "匹配方式"
      Height          =   630
      Left            =   105
      TabIndex        =   10
      Top             =   1650
      Width           =   4560
      Begin VB.OptionButton optMatch 
         Caption         =   "从左匹配"
         Height          =   180
         Index           =   0
         Left            =   825
         TabIndex        =   11
         Top             =   270
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "任意匹配"
         Height          =   180
         Index           =   1
         Left            =   2355
         TabIndex        =   12
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   4950
      TabIndex        =   4
      Top             =   1230
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   4950
      TabIndex        =   3
      Top             =   495
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "定位(&L)"
      Height          =   350
      Left            =   4950
      TabIndex        =   2
      Top             =   120
      Width           =   1100
   End
   Begin VB.Frame fra条件 
      Caption         =   "查找条件"
      Height          =   1455
      Left            =   90
      TabIndex        =   13
      Top             =   120
      Width           =   4560
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   885
         MaxLength       =   255
         TabIndex        =   6
         Top             =   240
         Width           =   3525
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   870
         MaxLength       =   255
         TabIndex        =   8
         Top             =   630
         Width           =   3525
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   870
         MaxLength       =   255
         TabIndex        =   1
         Top             =   1020
         Width           =   3525
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "编码(&C)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   0
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   690
         Width           =   630
      End
   End
   Begin VB.Label lbl结果 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 请输入查找条件"
      ForeColor       =   &H8000000D&
      Height          =   510
      Left            =   120
      TabIndex        =   9
      Top             =   2355
      Width           =   5925
   End
End
Attribute VB_Name = "frmPriceFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsFind As New ADODB.Recordset


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    Set mrsFind = Nothing
End Sub

Private Sub cmdFind_Click()
    If mrsFind.State = 1 Then
        If Not mrsFind.EOF Then mrsFind.MoveNext
        Call LocateItem
        Exit Sub
    End If
    If IsValid = False Then Exit Sub
    gstrSQL = ""
    If txtEdit(0).Text <> "" Then
        gstrSQL = "A.编码 like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(0).Text) & "%' or "
    End If
    If txtEdit(1).Text <> "" Then
        gstrSQL = gstrSQL & "B.名称 like '" & IIf(optMatch(1).Value = True, "%", "") & txtEdit(1).Text & "%' or "
    End If
    If txtEdit(2).Text <> "" Then
        gstrSQL = gstrSQL & "B.简码 like '" & IIf(optMatch(1).Value = True, "%", "") & UCase(txtEdit(2).Text) & "%' or "
    End If
    If gstrSQL <> "" Then
        gstrSQL = Mid(gstrSQL, 1, Len(gstrSQL) - 4)
    Else
        MsgBox "请输入查找条件。", vbExclamation, gstrSysName
        txtEdit(0).SetFocus
        Exit Sub
    End If
    
    Select Case Mid(frmPriceQuery.tvwMain_S.SelectedItem.Key, 2, 1)
    Case 0
        gstrSQL = "(" & gstrSQL & ") And A.类别 not in ('4','5','6','7')"
    Case 1
        gstrSQL = "(" & gstrSQL & ") And A.类别='5'"
    Case 2
        gstrSQL = "(" & gstrSQL & ") And A.类别='6'"
    Case 3
        gstrSQL = "(" & gstrSQL & ") And A.类别='7'"
    Case 7
        gstrSQL = "(" & gstrSQL & ") And A.类别='4'"
    End Select
    
    gstrSQL = "select distinct A.分类ID,A.ID,A.名称 " & _
            " from 收费项目目录 A,收费项目别名 B  " & _
            " where A.ID =B.收费细目ID(+) And " & gstrSQL & _
            " and (A.撤档时间 is null or A.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
            IIf(frmPriceQuery.mnuViewShowDynamic.Checked, "", " and A.是否变价=0")
    Call OpenRecordset(mrsFind, Me.Caption)
    
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
    lbl结果.Caption = "找到" & mrsFind.RecordCount & "条符合要求的收费细目。" & vbCrLf & "当前是第" & mrsFind.AbsolutePosition & "条，名称为：" & mrsFind("名称")
    
    With frmPriceQuery.tvwMain_S
        lngID = mrsFind("ID")
        Select Case Mid(.SelectedItem.Key, 2, 1)
        Case 0
            .Nodes(Mid(.SelectedItem.Key, 1, 2) & mrsFind("分类ID")).Selected = True
        Case 1, 2, 3
            gstrSQL = "Select Z.分类id From 药品规格 T, 诊疗项目目录 Z Where T.药名id = Z.Id and T.药品ID =" & lngID
            Call OpenRecordset(rsTemp, Me.Caption)
            If rsTemp.EOF Then
                Exit Sub
            End If
            .Nodes(Mid(.SelectedItem.Key, 1, 2) & rsTemp("分类ID")).Selected = True
        Case 4
            '等卫生材料落实后编制
            Exit Sub
        End Select
        .SelectedItem.EnsureVisible
    End With
    frmPriceQuery.FillSum
        
    lngID = mrsFind("ID")
    With frmPriceQuery.mshSum
        For lngCount = 1 To .Rows - 1
            If .RowData(lngCount) = lngID Then
                .Row = lngCount
                .TopRow = lngCount
                Exit Sub
            End If
        Next
    End With
    MsgBox "“" & mrsFind("名称") & "”的价格还未设置，或已经过期了。", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
'功能:分析输入有关费别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 2
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
          SendKeys "{TAB}"
    End If
End Sub

Private Sub optMatch_Click(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
End Sub

Private Sub optMatch_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mrsFind.State = 1 Then mrsFind.Close
    lbl结果.Caption = "  条件已改变，请重新定位"
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    '10.35.70，部分用户输入法会导致错误退出窗口，原因未知，暂时屏蔽
'    If Index = 1 Then
'        zlCommFun.OpenIme True
'    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
'    zlCommFun.OpenIme False
End Sub
