VERSION 5.00
Begin VB.Form frm医保类别中心 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保中心编辑"
   ClientHeight    =   2250
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   5595
   Icon            =   "frm医保类别中心.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4260
      TabIndex        =   9
      Top             =   1710
      Width           =   1100
   End
   Begin VB.Frame fra中心 
      Caption         =   "医保中心"
      Height          =   1905
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   3855
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   4
         Top             =   780
         Width           =   1755
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1935
         MaxLength       =   5
         TabIndex        =   2
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1935
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1215
         Width           =   1740
      End
      Begin VB.Image img装饰 
         Height          =   240
         Left            =   240
         Picture         =   "frm医保类别中心.frx":000C
         Top             =   540
         Width           =   240
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "中心编码(&D)"
         Height          =   180
         Index           =   1
         Left            =   915
         TabIndex        =   3
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "中心序号(&S)"
         Height          =   180
         Index           =   0
         Left            =   900
         TabIndex        =   1
         Top             =   420
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "中心名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   930
         TabIndex        =   5
         Top             =   1275
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4260
      TabIndex        =   8
      Top             =   750
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4260
      TabIndex        =   7
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frm医保类别中心"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum编辑
    Text序号 = 0
    text编码 = 1
    Text名称 = 2
End Enum

Dim mlng险类 As Long           '当前编辑的保险中心险类
Dim mstr序号 As String         '当前编辑的保险中心序号
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    
    MousePointer = vbHourglass
    If Save保险中心() = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    MousePointer = vbDefault
    
    mblnOK = True
    mblnChange = False
    
    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能:分析输入有关保险中心的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim lngIndex As Integer
    Dim strTemp As String
    For lngIndex = Text序号 To Text名称
        If zlCommFun.StrIsValid(Trim(TxtEdit(lngIndex).Text), TxtEdit(lngIndex).MaxLength) = False Then
            TxtEdit(lngIndex).SetFocus
            zlControl.TxtSelAll TxtEdit(lngIndex)
            Exit Function
        End If
        
        If Len(Trim(TxtEdit(lngIndex).Text)) = 0 Then
            TxtEdit(lngIndex).Text = ""
            MsgBox "序号或名称都不能为空。", vbExclamation, gstrSysName
            TxtEdit(lngIndex).SetFocus
            Exit Function
        End If
    Next
    
    If TxtEdit(Text序号).Enabled = True Then
        If IsNumeric(TxtEdit(Text序号)) = False Or Val(TxtEdit(Text序号).Text) <= 0 Then
            MsgBox "序号只能是大于900的整数。", vbExclamation, gstrSysName
            zlControl.TxtSelAll TxtEdit(Text序号)
            TxtEdit(Text序号).SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Function Save保险中心() As Boolean
'功能:保存编辑的内容到保险中心表中
'参数:
'返回值:成功返回True,否则为False
    Dim lng序号 As Long
    Dim lst As ListItem
    
    On Error GoTo errHandle
    
    If mstr序号 = "" Then     '新增一条记录
        lng序号 = TxtEdit(Text序号).Text
        gstrSQL = "zl_保险中心目录_Insert(" & mlng险类 & "," & lng序号 & ",'" & TxtEdit(text编码).Text & "','" & TxtEdit(Text名称).Text & "')"
    Else                      '修改
        gstrSQL = "zl_保险中心目录_Update(" & mlng险类 & "," & mstr序号 & ",'" & TxtEdit(text编码).Text & "','" & TxtEdit(Text名称).Text & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '在主界面上做相应的调整
    With frm医保类别.cmb中心
        If mstr序号 <> "" Then
            .RemoveItem .ListIndex
            lng序号 = Val(mstr序号)
        End If
        '新增
        .AddItem TxtEdit(text编码).Text & "." & TxtEdit(Text名称).Text
        .ItemData(.NewIndex) = lng序号
        .ListIndex = .NewIndex
    End With
    
    Save保险中心 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 编辑保险中心(ByVal lng险类 As Long, ByVal str序号 As String) As Boolean
'功能:用来与调用的保险中心管理窗口进行通讯的程序
'参数::lng险类           当前编辑的保险中心的的险类
'      str序号           当前编辑的保险中心的的序号
'返回值:编辑成功返回True,否则为False
    Dim rs保险中心 As New ADODB.Recordset
    Dim lng序号 As Long
    
    mlng险类 = lng险类
    mstr序号 = str序号
    mblnOK = False
    
    rs保险中心.CursorLocation = adUseClient
    
    If str序号 <> "" Then
        gstrSQL = "Select 名称,编码 From 保险中心目录  Where 序号=[1] and 险类=[2]"
        Set rs保险中心 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(str序号), mlng险类)
        
        TxtEdit(Text序号).Text = str序号
        TxtEdit(text编码).Text = rs保险中心("编码")
        TxtEdit(Text名称).Text = rs保险中心("名称")
        
        lblEdit(Text序号).Enabled = False
        TxtEdit(Text序号).Enabled = False
    Else
        lng序号 = Val(zlDatabase.GetMax("保险中心目录", "序号", 5, " where 险类=" & mlng险类))
        If lng序号 < 1 Then lng序号 = 0
        TxtEdit(Text序号).Text = lng序号
    End If
    
    mblnChange = False
    frm医保类别中心.Show vbModal
    编辑保险中心 = mblnOK
End Function

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    Select Case Index
        Case Text名称
          zlCommFun.OpenIme True
        Case Text序号, text编码
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 '使之不响
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = Text序号 Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub
