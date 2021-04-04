VERSION 5.00
Begin VB.Form frmDataType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "病人类型说明"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicType 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00E2E2E2&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   2940
      TabIndex        =   4
      Top             =   300
      Width           =   2970
   End
   Begin VB.PictureBox PicTitle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   2940
      TabIndex        =   0
      Top             =   0
      Width           =   2970
      Begin VB.Label LabClose 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2580
         TabIndex        =   1
         ToolTipText     =   "关闭窗口"
         Top             =   30
         Width           =   345
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "表示颜色"
         Height          =   195
         Index           =   0
         Left            =   1830
         TabIndex        =   2
         Top             =   45
         Width           =   1095
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "数据类型"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.Line Line3 
      X1              =   3555
      X2              =   3555
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   345
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3555
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmDataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MP As POINTAPI
Dim blnClick As Boolean, mfrmParent As Object
Private Sub PicType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        GetCursorPos MP
        blnClick = True
    End If
End Sub

Private Sub PicType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TmpMp As POINTAPI
    If blnClick = True Then
        GetCursorPos TmpMp
        Me.Top = Me.Top + (TmpMp.Y - MP.Y) * 15
        Me.Left = Me.Left + (TmpMp.X - MP.X) * 15
        GetCursorPos MP
    End If
End Sub

Private Sub PicType_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        blnClick = False
    End If
End Sub

Private Sub PicTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        GetCursorPos MP
        blnClick = True
    End If
End Sub

Private Sub PicTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TmpMp As POINTAPI
    If blnClick = True Then
        GetCursorPos TmpMp
        Me.Top = Me.Top + (TmpMp.Y - MP.Y) * 15
        Me.Left = Me.Left + (TmpMp.X - MP.X) * 15
        GetCursorPos MP
    End If
End Sub

Private Sub PicTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        blnClick = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub LabClose_Click()
    Unload Me
End Sub
Public Sub ShowPatiType(frmParent As Object, ByVal strContent As String)
'功能:在frmParent窗口右下角显示一窗体，内容为各种数据修改类型的颜色说明
'strContent 格式为 名称-颜色||名称-颜色
Dim IndexTmp As Integer
Dim ArrNote() As String, ArrCode() As String
Dim strTmp As String
Dim i As Integer
    
    On Error GoTo errH
    
    If strContent = "" Then Exit Sub
    
    ArrNote = Split(strContent, "||")
    
    Set mfrmParent = frmParent
    If Me.Visible Then Unload Me
  
    For i = 0 To UBound(ArrNote)
        strTmp = ArrNote(i)
        If strTmp = "" Then strTmp = "未知-0"
        ArrCode = Split(strTmp, "-")
        IndexTmp = lblType.ubound + 1
        Load lblType(IndexTmp)
        Load lblColor(IndexTmp)
        lblType(IndexTmp).AutoSize = True
        lblType(IndexTmp).Height = 200
        lblColor(IndexTmp).Height = 200
        
        Set lblType(IndexTmp).Container = PicType
        Set lblColor(IndexTmp).Container = PicType
        lblType(IndexTmp).Top = IIf(IndexTmp = 1, 100, (lblType.ubound - 1) * 300 + 100)
        lblType(IndexTmp).Left = 105
        lblColor(IndexTmp).Top = lblType(IndexTmp).Top
        lblColor(IndexTmp).Left = 1830
        lblType(IndexTmp).Caption = ArrCode(0): If lblType(IndexTmp).Width > 1600 Then lblType(IndexTmp).Width = 1600
        lblType(IndexTmp).BackColor = PicType.BackColor
        
        lblColor(IndexTmp).Caption = ""
        lblColor(IndexTmp).BackColor = ArrCode(1)
        lblType(IndexTmp).Visible = True
        lblColor(IndexTmp).Visible = True
    Next i
    
    PicType.Height = lblType.ubound * 300 + 100
    Me.Height = PicTitle.Height + PicType.Height
    On Error Resume Next
    
    If Me.Top < 0 Or Me.Left < 0 Then
        Me.Top = 0: Me.Left = 0
    End If

    Dim objBar As Object, objPoint As RECT
    On Error Resume Next
    For Each objBar In mfrmParent
        If UCase(TypeName(objBar)) = "STATUSBAR" Then Exit For
    Next
    Call GetWindowRect(objBar.hWnd, objPoint)
    
    Me.Top = objPoint.Top * Screen.TwipsPerPixelY - Me.Height: If Me.Top < 0 Then Me.Top = 0
    Me.Left = objPoint.Right * Screen.TwipsPerPixelX - Me.Width - 200: If Me.Left < 0 Then Me.Left = 0
    
    Me.Show 0, frmParent
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
