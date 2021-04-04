VERSION 5.00
Begin VB.Form frmReportElement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择要素,右键确定"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   Icon            =   "frmReportElement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   4515
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkCheckOrderInsert 
      Caption         =   "按照勾选顺序加入元素"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   5160
      Width           =   3135
   End
   Begin VB.PictureBox picElement 
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   3915
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin VB.VScrollBar vScroll 
         Height          =   2295
         LargeChange     =   50
         Left            =   3480
         SmallChange     =   10
         TabIndex        =   6
         Top             =   120
         Width           =   350
      End
      Begin VB.Frame frmElement 
         Height          =   4215
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3375
         Begin VB.OptionButton optItem 
            Caption         =   "Option1"
            Height          =   400
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   3415
         End
         Begin VB.CheckBox chkItem 
            Caption         =   "Check1"
            Height          =   400
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Visible         =   0   'False
            Width           =   2775
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   2400
      TabIndex        =   1
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   350
      Left            =   720
      TabIndex        =   0
      Top             =   6240
      Width           =   1100
   End
End
Attribute VB_Name = "frmReportElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strReturnElement As String
Private mintSelType As Integer '要素类型 0--单选；1--复选
Private mlngFrameDelta As Long  'frame在垂直方向的位移修正
Private mblnCheckOrderInsert As Boolean     '按照勾选顺序加入元素
Private mstrElement As String               '按照勾选顺序保存元素内容

'本窗体的事件
Public Event ReturnElement(strElement As String)

Public Sub ShowElement(strElements As String, iType As Integer)
    'iType 要素类型 0--单选；1--复选
    Dim strItems() As String
    Dim iItemCount As Integer
    Dim strTemp As String
    Dim i As Integer
    
    strReturnElement = ""
    mintSelType = iType
    
    strTemp = Left(strElements, Len(strElements) - 2)
    strTemp = Right(strTemp, Len(strTemp) - 2)
    strItems = Split(strTemp, ";")
    '清除原有控件
    For i = 1 To optItem.Count - 1
        Unload optItem(i)
    Next i
    For i = 1 To chkItem.Count - 1
        Unload chkItem(i)
    Next i

    '一页最多显示20个选项，超过20个选项，就显示滚动条
    If mintSelType = 0 Then
        For i = 0 To UBound(strItems)
            Load optItem(i + 1)
            If i + 1 = 1 Then
                optItem(i + 1).Top = 200
            Else
                optItem(i + 1).Top = optItem(i).Top + 400
            End If
            optItem(i + 1).Left = 80
            optItem(i + 1).Visible = True
            optItem(i + 1).value = True
            
            optItem(i + 1).Caption = strItems(i)
        Next i
        '设置frmEelement的高度
        frmElement.Height = optItem(optItem.Count - 1).Top + optItem(optItem.Count - 1).Height + 200
    ElseIf mintSelType = 1 Then
        For i = 0 To UBound(strItems)
            Load chkItem(i + 1)
            If i + 1 = 1 Then
                chkItem(i + 1).Top = 200
            Else
                chkItem(i + 1).Top = chkItem(i).Top + 400
            End If
            chkItem(i + 1).Left = 80
            chkItem(i + 1).Visible = True
            
            chkItem(i + 1).Caption = strItems(i)
        Next i
        '设置frmEelement的高度
        frmElement.Height = chkItem(chkItem.Count - 1).Top + chkItem(chkItem.Count - 1).Height + 200
    End If
    
    vScroll.Visible = False
    '调整窗口大小
    If Me.frmElement.Height > 7000 Then
        Me.picElement.Height = 7000
        vScroll.Visible = True
    Else
        Me.picElement.Height = Me.frmElement.Height + 50 - mlngFrameDelta
    End If
    
    If mintSelType = 1 Then     '多选
        chkCheckOrderInsert.Visible = True
        Me.Height = Me.picElement.Height + 1600
    Else
        chkCheckOrderInsert.Visible = False
        Me.Height = Me.picElement.Height + 1200
    End If
    
    '把窗口放到当前鼠标的位置
    Dim vPos As POINTAPI
    
    GetCursorPos vPos
    Me.Left = vPos.X * Screen.TwipsPerPixelX - Me.Width / 6
    Me.Top = vPos.Y * Screen.TwipsPerPixelY - Me.Height / 2
    
    '如果窗口不能完全显示，则调整窗口位置
    If Me.Top < 0 Then
        Me.Top = 10
    End If
    
    If Me.Top + frmElement.Height + 1600 > Screen.Height Then
        Me.Top = IIf(Screen.Height - Me.Height - 1100 > 0, Screen.Height - Me.Height - 1100, 10)
    End If
    
    If Me.Left < 0 Then
        Me.Left = 10
    End If
    
    If Me.Left + Me.Width > Screen.Width Then
        Me.Left = IIf(Screen.Width - Me.Width - 500 > 0, Screen.Width - Me.Width - 500, 10)
    End If
    
    Me.Show 1
End Sub

Private Sub chkCheckOrderInsert_Click()
    mblnCheckOrderInsert = (chkCheckOrderInsert.value = 1)
End Sub

Private Sub chkItem_Click(Index As Integer)
    If chkItem(Index).value = 1 Then
        mstrElement = mstrElement & chkItem(Index).Caption & "，"
    Else
        '删除被选中的元素
        If InStr(mstrElement, "，" & chkItem(Index).Caption & "，") > 0 Then
            mstrElement = Left(mstrElement, InStr(mstrElement, "，" & chkItem(Index).Caption & "，")) _
                & Mid(mstrElement, InStr(mstrElement, "，" & chkItem(Index).Caption & "，") + Len("，" & chkItem(Index).Caption & "，"))
        End If
    End If
End Sub

Private Sub chkItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call CmdOK_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim strRElement As String
    Dim i As Integer
    
    If mintSelType = 0 Then
        For i = 1 To optItem.Count - 1
            If optItem(i).value = True Then
                strRElement = optItem(i).Caption
                Exit For
            End If
        Next i
    ElseIf mintSelType = 1 Then
        If mblnCheckOrderInsert = True Then
            If Len(mstrElement) > 2 Then
                strRElement = Mid(mstrElement, 2, Len(mstrElement) - 2)
            End If
        Else
            For i = 1 To chkItem.Count - 1
                If chkItem(i).value = 1 Then
                    If strRElement = "" Then
                        strRElement = chkItem(i).Caption
                    Else
                        strRElement = strRElement & "，" & chkItem(i).Caption
                    End If
                End If
            Next i
        End If
    End If
    
    RaiseEvent ReturnElement(strRElement)
    strReturnElement = strRElement
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strTemp As String
    
    On Error GoTo err
    
    ''''''''''''''''''''''''''处理鼠标滚轮'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set mReport.fReportElement = Me
    glngEelmentWinProc = ElementHook(Me.hWnd)
    ''''''''''''''''''''''''''处理鼠标滚轮'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '设置默认值
    Me.vScroll.SmallChange = 50
    Me.vScroll.LargeChange = 200
    mlngFrameDelta = 100
    mstrElement = "，"
    
    strTemp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmReportElement", "按照勾选顺序加入元素", "")
    If strTemp = "1" Then
        mblnCheckOrderInsert = True
    Else
        mblnCheckOrderInsert = False
    End If
    chkCheckOrderInsert.value = IIf(mblnCheckOrderInsert = True, 1, 0)
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Sub Form_Resize()
    Call subResizeForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ''''''''''''''''''''''解除鼠标滚轮'''''''''''''''''''''''''''''''''
     '    卸载hook
    Call ElementUnhook(Me.hWnd, glngEelmentWinProc)
    ''''''''''''''''''''''解除鼠标滚轮'''''''''''''''''''''''''''''''''
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmReportElement", "按照勾选顺序加入元素", IIf(mblnCheckOrderInsert = True, "1", "0"))
    
End Sub

Private Sub frmElement_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call CmdOK_Click
    End If
End Sub

Private Sub optItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call CmdOK_Click
    End If
End Sub

Public Sub subMouseWheel(intDirection As Integer)
'------------------------------------------------
'功能：处理鼠标滚轮的消息
'参数：intDirection--滚轮滚动方向 1-鼠标上滚；2-鼠标下滚
'返回：无
'------------------------------------------------
    
    '发生错误，不做任何提示
    On Error Resume Next
    
    If intDirection = 1 Then
        '鼠标上滚
        vScroll.value = IIf(vScroll.value - vScroll.LargeChange < vScroll.Min, vScroll.Min, vScroll.value - vScroll.LargeChange)
    Else
        '鼠标下滚
        vScroll.value = IIf(vScroll.value + vScroll.LargeChange > vScroll.Max, vScroll.Max, vScroll.value + vScroll.LargeChange)
    End If
    
End Sub

Private Sub vScroll_Change()
    Me.frmElement.Top = -(vScroll.value + mlngFrameDelta)
End Sub

Private Sub subResizeForm()
'------------------------------------------------
'功能：重新调整窗体内容
'参数：无
'返回：无
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    '调整控件位置
    Me.picElement.Left = 20
    Me.picElement.Top = 20
    Me.picElement.Width = Me.Width - 100
    Me.frmElement.Top = 0 - mlngFrameDelta
    Me.frmElement.Left = 0
    
    '是否显示滚动条
    If Me.vScroll.Visible = True Then
        Me.vScroll.Left = Abs(Me.picElement.Width - Me.vScroll.Width - 50)
        Me.vScroll.Top = 20
        Me.vScroll.Height = Me.picElement.Height - 60
        Me.vScroll.Max = Abs(Me.frmElement.Height - Me.picElement.Height - mlngFrameDelta)
        Me.vScroll.value = 0
        Me.frmElement.Width = vScroll.Left
    Else
        Me.frmElement.Width = Abs(Me.picElement.Width - 50)
    End If

    If Me.chkCheckOrderInsert.Visible = True Then
        Me.chkCheckOrderInsert.Top = Me.picElement.Height + 100
        Me.chkCheckOrderInsert.Left = 150
        Me.cmdOK.Top = Me.chkCheckOrderInsert.Top + Me.chkCheckOrderInsert.Height + 100
    Else
        Me.cmdOK.Top = Me.picElement.Height + 200
    End If
    
    Me.cmdCancel.Top = Me.cmdOK.Top
    Me.cmdOK.Left = Abs(Me.Width - Me.cmdOK.Width * 2 - 500) / 2
    Me.cmdCancel.Left = Me.cmdOK.Left + Me.cmdOK.Width + 500
    
    For i = 1 To optItem.Count - 1
        optItem(i).Width = Abs(Me.frmElement.Width - 200)
    Next i
    
    For i = 1 To chkItem.Count - 1
        chkItem(i).Width = Abs(Me.frmElement.Width - 200)
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub
