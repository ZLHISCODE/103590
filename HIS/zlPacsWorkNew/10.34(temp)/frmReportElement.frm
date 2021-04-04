VERSION 5.00
Begin VB.Form frmReportElement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择要素,右键确定"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   Icon            =   "frmReportElement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   350
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1100
   End
   Begin VB.Frame frmElement 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
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
      Begin VB.OptionButton optItem 
         Caption         =   "Option1"
         Height          =   400
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   3415
      End
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
    Me.Height = 2000
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
            
            '调整窗口大小
            If optItem(i + 1).Top + optItem(i + 1).Height > frmElement.Height Then Me.Height = optItem(i + 1).Top + optItem(i + 1).Height + 1100
        Next i
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
            
            '调整窗口大小
            If chkItem(i + 1).Top + chkItem(i + 1).Height > frmElement.Height Then Me.Height = chkItem(i + 1).Top + chkItem(i + 1).Height + 1100
        Next i
    End If
    
    
    '把窗口放到当前鼠标的位置
    Dim vPos As PointAPI
    
    GetCursorPos vPos
    Me.Left = vPos.X * Screen.TwipsPerPixelX - Me.Width / 6
    Me.Top = vPos.Y * Screen.TwipsPerPixelY - Me.Height / 2
    
    Me.Show 1
End Sub


Private Sub chkItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call cmdOK_Click
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
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
        For i = 1 To chkItem.Count - 1
            If chkItem(i).value = 1 Then
                If strRElement = "" Then
                    strRElement = chkItem(i).Caption
                Else
                    strRElement = strRElement & "," & chkItem(i).Caption
                End If
            End If
        Next i
    End If
    
    RaiseEvent ReturnElement(strRElement)
    strReturnElement = strRElement
    Unload Me
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    '调整控件位置
    Me.frmElement.Left = 50
    Me.frmElement.Top = 20
    Me.frmElement.Width = Abs(Me.Width - 200)
    Me.frmElement.Height = Abs(Me.Height - 1000)
    
    Me.cmdOK.Top = Me.frmElement.Top + Me.frmElement.Height + 50
    Me.cmdCancel.Top = Me.cmdOK.Top
    
    For i = 1 To optItem.Count - 1
        optItem(i).Width = Abs(Me.frmElement.Width - 200)
    Next i
    
    For i = 1 To chkItem.Count - 1
        chkItem(i).Width = Abs(Me.frmElement.Width - 200)
    Next i
    
End Sub

Private Sub frmElement_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call cmdOK_Click
    End If
End Sub

Private Sub optItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call cmdOK_Click
    End If
End Sub
