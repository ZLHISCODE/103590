VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQualityDataQueyRule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "质控规则"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5985
   Icon            =   "FrmQualityDataQueyRule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4740
      TabIndex        =   1
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3390
      TabIndex        =   0
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   8
      Top             =   3780
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3645
      Left            =   30
      TabIndex        =   9
      Top             =   30
      Width           =   5895
      Begin VB.CommandButton CmdMoveOneToRight 
         Caption         =   ">"
         Height          =   405
         Left            =   2640
         TabIndex        =   3
         Top             =   780
         Width           =   585
      End
      Begin VB.CommandButton CmdMoveOneToLeft 
         Caption         =   "<"
         Height          =   405
         Left            =   2640
         TabIndex        =   4
         Top             =   1365
         Width           =   585
      End
      Begin VB.CommandButton CmdMoveAllRight 
         Caption         =   ">>"
         Height          =   405
         Left            =   2640
         TabIndex        =   5
         Top             =   1950
         Width           =   585
      End
      Begin VB.CommandButton CmdMoveAllLeft 
         Caption         =   "<<"
         Height          =   405
         Left            =   2640
         TabIndex        =   6
         Top             =   2550
         Width           =   585
      End
      Begin MSComctlLib.ListView LivAll 
         Height          =   3135
         Left            =   60
         TabIndex        =   2
         Top             =   390
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "编码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "规则类型"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "X"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "M"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LivSelect 
         Height          =   3135
         Left            =   3300
         TabIndex        =   7
         Top             =   390
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "编码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "规则类型"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "X"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "M"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "未选质控规则"
         Height          =   180
         Left            =   90
         TabIndex        =   11
         Top             =   180
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "已选质控规则"
         Height          =   180
         Left            =   3330
         TabIndex        =   10
         Top             =   180
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmQualityDataQueyRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DefaultID As Long                               '得到默认的ID

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub CmdHelp_Click()
    '显示帮助
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub CmdMoveAllLeft_Click()
    Dim SelectIndex As Integer
    Dim ItmX As ListItem
    Dim i As Integer
    '取消全部
    If Me.LivSelect.ListItems.Count > 0 Then
        
        For i = 1 To Me.LivSelect.ListItems.Count
            
            With Me.LivSelect.ListItems(i)
                Set ItmX = Me.LivAll.ListItems.Add(, .Key, .Text)
                ItmX.SubItems(1) = .SubItems(1)
                ItmX.SubItems(2) = .SubItems(2)
                ItmX.SubItems(3) = .SubItems(3)
                ItmX.SubItems(4) = .SubItems(4)
            End With
            
        Next
        
        Me.LivSelect.ListItems.Clear
        DefaultID = 0
        
    End If
    
End Sub

Private Sub CmdMoveAllRight_Click()
    Dim SelectIndex As Integer
    Dim ItmX As ListItem, i As Integer
    '选中全部
    If Me.LivAll.ListItems.Count > 0 Then
        For i = 1 To Me.LivAll.ListItems.Count
            With Me.LivAll.ListItems(i)
                Set ItmX = Me.LivSelect.ListItems.Add(, .Key, .Text)
                ItmX.SubItems(1) = .SubItems(1)
                ItmX.SubItems(2) = .SubItems(2)
                ItmX.SubItems(3) = .SubItems(3)
                ItmX.SubItems(4) = .SubItems(4)
            End With
        Next
        Me.LivAll.ListItems.Clear
    End If
End Sub

Private Sub CmdMoveOneToLeft_Click()
    Dim SelectIndex As Integer
    Dim ImtX As ListItem
    
    '取消选中一条
    If Me.LivSelect.ListItems.Count <= 0 Then Exit Sub
    If Me.LivSelect.SelectedItem.Index = 0 Then Exit Sub
    
    SelectIndex = Me.LivSelect.SelectedItem.Index
    
    With Me.LivSelect.ListItems(SelectIndex)

        Set ImtX = Me.LivAll.ListItems.Add(, .Key, .Text)
        
        ImtX.SubItems(1) = .SubItems(1)
        ImtX.SubItems(2) = .SubItems(2)
        ImtX.SubItems(3) = .SubItems(3)
        ImtX.SubItems(4) = .SubItems(4)
        
        If .SubItems(2) = "√" Then
            DefaultID = 0
        End If
        
    End With
    
    Me.LivSelect.ListItems.Remove (SelectIndex)
    Me.LivSelect.SetFocus
    
End Sub

Private Sub CmdMoveOneToRight_Click()
    Dim SelectIndex As Integer
    Dim ItmX As ListItem
    
    '选中一条
    If Me.LivAll.ListItems.Count <= 0 Then Exit Sub
    If Me.LivAll.SelectedItem.Index = 0 Then Exit Sub
    
    SelectIndex = Me.LivAll.SelectedItem.Index
    
    With Me.LivAll.ListItems(SelectIndex)

        Set ItmX = Me.LivSelect.ListItems.Add(, .Key, .Text)
        ItmX.SubItems(1) = .SubItems(1)
        ItmX.SubItems(2) = .SubItems(2)
        ItmX.SubItems(3) = .SubItems(3)
        ItmX.SubItems(4) = .SubItems(4)
        
    End With
    
    Me.LivAll.ListItems.Remove (SelectIndex)
    Me.LivAll.SetFocus
    
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    FrmQualityDataQuery.QualityRule = ""
    For i = 1 To Me.LivSelect.ListItems.Count
        With FrmQualityDataQuery
            If Len(.QualityRule) = 0 Then
                .QualityRule = .QualityRule & Mid(Me.LivSelect.ListItems(i).Key, 2)
            Else
                .QualityRule = .QualityRule & "," & Mid(Me.LivSelect.ListItems(i).Key, 2)
            End If
        End With
    Next
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "选用质控规则", FrmQualityDataQuery.QualityRule)
    
    Unload Me
End Sub

Private Sub Form_Load()
            
    '读入质控规则
    LoadQualityRule

        
End Sub
Sub LoadQualityRule()
    '''''''''''''''''''''''''''''''''
    '功能           读入质控规则
    '    参数
    '    Default    =1读入默认为真的记录;=0读入默认值为假的记录
    '''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    Dim strTmp As String
    Dim intTmp As Integer
    Dim strsql As String
    Dim SelectItem As Boolean
    Dim strQualityRule As String
    Dim i As Integer
    '清空列表
    
    Me.LivSelect.ListItems.Clear
    Me.LivAll.ListItems.Clear
    
    strQualityRule = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "选用质控规则", "")
    If strQualityRule = "" Then
        '已选中
        gstrSql = "select * from 检验质控规则 where 缺省规则 = " & 1
    Else
        gstrSql = "select * from 检验质控规则 "
        strTmp = strQualityRule
        Do Until Len(strTmp) = 0
            intTmp = InStr(strTmp, ",")
            If strsql = "" Then
                If intTmp = 0 Then
                    strsql = strsql & " where id = " & strTmp
                Else
                    strsql = strsql & " where id = " & Mid(strTmp, 1, intTmp - 1)
                End If
            Else
                If intTmp = 0 Then
                    strsql = strsql & " or id = " & strTmp
                Else
                
                    strsql = strsql & " or id = " & Mid(strTmp, 1, intTmp - 1)
                End If
            End If
            If intTmp = 0 Then
                strTmp = ""
            Else
                strTmp = Mid(strTmp, intTmp + 1)
            End If
        Loop
    End If
    
    gstrSql = gstrSql & strsql
    
    OpenRecord rsTmp, gstrSql, Me.Caption

    Do Until rsTmp.EOF
        
        Set ItmX = Me.LivSelect.ListItems.Add(, "A" & rsTmp("ID"), rsTmp("编码"))
        
        ItmX.SubItems(1) = rsTmp("规则名称")
        ItmX.SubItems(3) = IIf(rsTmp("N") = 0, "", rsTmp("N"))
        ItmX.SubItems(4) = IIf(rsTmp("X") = 0, "", rsTmp("X"))
        ItmX.SubItems(5) = IIf(rsTmp("M") = 0, "", rsTmp("M"))
        
        Select Case rsTmp("规则类型")
            Case 0
                ItmX.SubItems(2) = "N-XS"
            Case 1
                ItmX.SubItems(2) = "R-Xs"
            Case 2
                ItmX.SubItems(2) = "N-T"
            Case 3
                ItmX.SubItems(2) = "N-X"
            Case 4
                ItmX.SubItems(2) = "(M of N)XS"
        End Select
        
        rsTmp.MoveNext
    Loop
    
    rsTmp.Close
    
    If strQualityRule = "" Then
        '已选中
        gstrSql = "select * from 检验质控规则 where 缺省规则 = " & 0
    Else
        gstrSql = "select * from 检验质控规则 "
    End If
        
    OpenRecord rsTmp, gstrSql, Me.Caption
    
    Do Until rsTmp.EOF
        
        For i = 1 To Me.LivSelect.ListItems.Count
            If Mid(Me.LivSelect.ListItems(i).Key, 2) = rsTmp("ID") Then
                SelectItem = True
                Exit For
            End If
        Next
        
        If SelectItem = False Then
            Set ItmX = Me.LivAll.ListItems.Add(, "A" & rsTmp("ID"), rsTmp("编码"))
            
            ItmX.SubItems(1) = rsTmp("规则名称")
            ItmX.SubItems(3) = IIf(rsTmp("N") = 0, "", rsTmp("N"))
            ItmX.SubItems(4) = IIf(rsTmp("X") = 0, "", rsTmp("X"))
            ItmX.SubItems(5) = IIf(rsTmp("M") = 0, "", rsTmp("M"))
            
            Select Case rsTmp("规则类型")
                Case 0
                    ItmX.SubItems(2) = "N-XS"
                Case 1
                    ItmX.SubItems(2) = "R-Xs"
                Case 2
                    ItmX.SubItems(2) = "N-T"
                Case 3
                    ItmX.SubItems(2) = "N-X"
                Case 4
                    ItmX.SubItems(2) = "(M of N)XS"
            End Select
        End If
        SelectItem = False
        rsTmp.MoveNext
    Loop
    
    rsTmp.Close
End Sub

