VERSION 5.00
Begin VB.Form frmSetWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "发药窗口设置"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   Icon            =   "frmSetWindow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -120
      TabIndex        =   6
      Top             =   1600
      Width           =   5025
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   1100
   End
   Begin VB.ComboBox cbo药房 
      ForeColor       =   &H80000012&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2280
   End
   Begin VB.ListBox lst发药窗口 
      Columns         =   1
      ForeColor       =   &H80000012&
      Height          =   900
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmSetWindow.frx":000C
      Left            =   1155
      List            =   "frmSetWindow.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   645
      Width           =   2280
   End
   Begin VB.Label Lbl药房 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "药房"
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
   Begin VB.Label Lbl发药窗口 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "发药窗口"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   690
      Width           =   720
   End
End
Attribute VB_Name = "frmSetWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbo药房_Click()
    Dim intDO As Integer
    Dim bln门诊 As Boolean, bln住院 As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    
    '不可能，如果没有设置药房，主界面都进不了
    If Me.cbo药房.ListCount = 0 Then Exit Sub
    If Val(Me.cbo药房.Tag) = Me.cbo药房.ListIndex Then
        Exit Sub
    Else
        Me.cbo药房.Tag = Me.cbo药房.ListIndex
    End If
    
    '根据药房显示单位
    strTmp = " Select 名称 From 发药窗口 Where 药房ID=" & Me.cbo药房.ItemData(Me.cbo药房.ListIndex)
    rsTmp.Open strTmp, gcnOracle
    
    With rsTmp
        Me.lst发药窗口.Clear
        lst发药窗口.Columns = 2
        Do While Not .EOF
            lst发药窗口.AddItem !名称
            .MoveNext
        Loop
        .Close
    End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub cmdOK_Click()
    Dim i As Integer
    Dim strFormNO As String
    
    If Me.cbo药房.ListCount = 0 Then Exit Sub
    
    For i = 0 To Me.lst发药窗口.ListCount - 1
        If Me.lst发药窗口.Selected(i) Then
            strFormNO = Me.lst发药窗口.List(i)
            Exit For
        End If
    Next
    
    SaveSetting "ZLSOFT", "未发药病人显示", "药房", cbo药房.ItemData(cbo药房.ListIndex)
    frmUnSendDrug.Entry cbo药房.ItemData(cbo药房.ListIndex), strFormNO
    Unload Me

End Sub

Private Sub Form_Load()
    Dim strTmp As String
    Dim lngStockID As Long, i As Long
    Dim rsTmp As New ADODB.Recordset
    strTmp = "Select Distinct p.Id, p.名称" & vbNewLine & _
            "From 部门表 P" & vbNewLine & _
            "Where p.Id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') And" & vbNewLine & _
            "      (p.撤档时间 Is Null Or p.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By p.名称"
    With rsTmp
        Me.cbo药房.Clear
        .Open strTmp, gcnOracle
        Do While Not .EOF
            Me.cbo药房.AddItem !名称
            Me.cbo药房.ItemData(Me.cbo药房.NewIndex) = !ID
            .MoveNext
        Loop
        .Close
        If Me.cbo药房.ListCount > 0 Then
            lngStockID = Val(GetSetting(appName:="ZLSOFT", Section:="未发药病人显示", Key:="药房", Default:=""))
            If lngStockID > 0 Then
                For i = 0 To cbo药房.ListCount - 1
                    If cbo药房.ItemData(i) = lngStockID Then
                        cbo药房.ListIndex = i
                        Exit For
                    End If
                Next
            Else
                cbo药房.ListIndex = 0
            End If
        End If
    End With
    Call cbo药房_Click
End Sub

Private Sub lst发药窗口_ItemCheck(Item As Integer)
    Dim i As Integer
    On Error Resume Next
    For i = 0 To lst发药窗口.ListCount - 1
        If i <> Item Then
            lst发药窗口.Selected(i) = False
        End If
    Next
End Sub

