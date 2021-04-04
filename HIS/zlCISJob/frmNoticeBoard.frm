VERSION 5.00
Begin VB.Form frmNoticeBoard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   4395
   ControlBox      =   0   'False
   Icon            =   "frmNoticeBoard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmNoticeBoard"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrFresh 
      Interval        =   60
      Left            =   3285
      Top             =   660
   End
   Begin VB.PictureBox picBak 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   165
      Picture         =   "frmNoticeBoard.frx":000C
      ScaleHeight     =   4665
      ScaleWidth      =   4065
      TabIndex        =   7
      Top             =   885
      Width           =   4095
      Begin VB.PictureBox PicElementCT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   870
         ScaleHeight     =   240
         ScaleWidth      =   1020
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   1024
         Begin VB.Label lblElementCT 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "要素内容"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Visible         =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   7
         Left            =   -30
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   6
         Left            =   -30
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   5
         Left            =   330
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   4
         Left            =   720
         Top             =   270
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   3
         Left            =   720
         Top             =   120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   2
         Left            =   720
         Top             =   -30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   1
         Left            =   330
         Top             =   -30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Shape shpCircle 
         BorderColor     =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   105
         Index           =   0
         Left            =   -30
         Top             =   -30
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblElementName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "要素名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   75
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   450
      ScaleHeight     =   240
      ScaleWidth      =   1890
      TabIndex        =   0
      Top             =   390
      Width           =   1890
      Begin VB.Label lblClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1665
         TabIndex        =   2
         Top             =   30
         Width           =   210
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病区公告栏"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   75
         TabIndex        =   1
         Top             =   30
         Width           =   975
      End
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   45
      Index           =   3
      Left            =   315
      TabIndex        =   6
      Top             =   5925
      Width           =   2000
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   45
      Index           =   2
      Left            =   330
      TabIndex        =   5
      Top             =   105
      Width           =   2000
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   5835
      Index           =   1
      Left            =   4275
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   135
      Width           =   45
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   5835
      Index           =   0
      Left            =   60
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   75
      Width           =   45
   End
End
Attribute VB_Name = "frmNoticeBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event ItemClick(ByVal strBeds As String)

Private mrsBoard As New ADODB.Recordset
Private mfrmParent As Object
Public mblnShow As Boolean
Private mlng病区ID  As Long
Private Const mlngMinH = 5120
Private Const mlngMinW = 4250
Private Const mlngMaxW = 8500
Private mblnFresh As Boolean

Public Sub ShowMe(frmParent As Object, ByVal lng病区ID As Long)
    Dim blnShow As Boolean
    
    Set mfrmParent = frmParent
    mlng病区ID = lng病区ID
    blnShow = Not mblnShow
    
    If blnShow Then
        mblnShow = True
        Me.Show , frmParent
    End If
    mblnFresh = True
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("`") Then
        KeyAscii = 0
        If mfrmParent.Visible Then
            mfrmParent.SetFocus
        End If
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strPos As String, lngH As Long, lngW As Long

    '-------------------------------------------------------------------
    On Error GoTo errH
    
    Call zlControl.FormSetCaption(Me, False, False)
    
    '恢复窗体尺寸,要先恢复尺寸(原本的宽度与设置的宽度存在差异）
    strPos = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "NoticeBoardSize", "4250,5120")
    If InStr(1, strPos, ",") <> 0 Then
        lngW = Val(Split(strPos, ",")(0))
        lngH = mlngMinH
    Else
        lngW = mlngMinW
        lngH = mlngMinH
    End If
    If lngW < mlngMinW Then lngW = mlngMinW
    If lngW > mlngMaxW Then lngW = mlngMaxW
    
    Me.Height = lngH: Me.Width = lngW
    strPos = (mfrmParent.Height - lngH) & ",-100"
    strPos = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "NoticeBoardPostion", strPos)
    Me.Top = mfrmParent.Top + Val(Split(strPos, ",")(0))
    If Me.Top > Screen.Height - lngH Then Me.Top = Screen.Height - lngH
    If Me.Top < 100 Then Me.Top = 100
    Me.Left = mfrmParent.Left + mfrmParent.Width + Val(Split(strPos, ",")(1)) - Me.Width
    If Me.Left > Screen.Width - lngW Then Me.Left = Screen.Width - lngW
    If Me.Left < 100 Then Me.Left = 100
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshData()
    Dim intCount As Integer, intDel As Integer
    Dim strSQL As String
    On Error GoTo ErrHand
    
    '首先更新公告栏数据
    strSQL = "Zl_病区公告栏样式_Updatedata(" & mlng病区ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "公告栏数据更新")
    
    '先删除所有控件
    intCount = lblElementName.Count - 1
    For intDel = 1 To intCount
        Unload lblElementName(intDel)
        Unload lblElementCT(intDel)
        Unload PicElementCT(intDel)
    Next
    picBak.Tag = ""
    
    strSQL = " Select ID,名称,别名,行号,位置,是否固定,是否隐藏,内容" & _
              " From 病区公告栏样式 " & _
              " Where 病区ID=[1] " & _
              " Order by 行号,位置"
    Set mrsBoard = zlDatabase.OpenSQLRecord(strSQL, "提取病区公告", mlng病区ID)
    
    '依次加载控件
    With mrsBoard
        Do While Not .EOF
            Load lblElementName(.AbsolutePosition)
            lblElementName(.AbsolutePosition).Tag = !ID
            lblElementName(.AbsolutePosition).Caption = !别名
            lblElementName(.AbsolutePosition).Top = lblElementName(0).Top + (!行号 - 1) * 360
            lblElementName(.AbsolutePosition).Left = IIf(!位置 = 1, 60, picBak.Width - lblElementName(.AbsolutePosition).Width - 1000)
            lblElementName(.AbsolutePosition).BackStyle = 0
            lblElementName(.AbsolutePosition).Visible = True
            
            Load PicElementCT(.AbsolutePosition)
            PicElementCT(.AbsolutePosition).Top = lblElementName(.AbsolutePosition).Top
            PicElementCT(.AbsolutePosition).Left = lblElementName(.AbsolutePosition).Left + lblElementName(.AbsolutePosition).Width + 60
            PicElementCT(.AbsolutePosition).Height = 240
            PicElementCT(.AbsolutePosition).Visible = True
            PicElementCT(.AbsolutePosition).ZOrder 0
            PicElementCT(.AbsolutePosition).Tag = !行号 & "," & !位置
            
            Load lblElementCT(.AbsolutePosition)
            Set lblElementCT(.AbsolutePosition).Container = PicElementCT(.AbsolutePosition)
            lblElementCT(.AbsolutePosition).Caption = IIf(IsNull(!内容), "", !内容)
            lblElementCT(.AbsolutePosition).Top = 0
            lblElementCT(.AbsolutePosition).Left = 0
            lblElementCT(.AbsolutePosition).AutoSize = False
            lblElementCT(.AbsolutePosition).WordWrap = False
            lblElementCT(.AbsolutePosition).Height = 240
            lblElementCT(.AbsolutePosition).Visible = True
            PicElementCT(.AbsolutePosition).Width = lblElementCT(.AbsolutePosition).Width
            
            
            If Val("" & !是否隐藏) = 1 And Trim(lblElementCT(.AbsolutePosition).Caption) = "" Then
                lblElementName(.AbsolutePosition).Visible = False
                lblElementCT(.AbsolutePosition).Visible = False
                PicElementCT(.AbsolutePosition).Visible = False
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    '重新整理要素位置
    Call picBak_Resize
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetShape(Optional ByVal intIndex As Integer = 0)
    Dim blnShow As Boolean
    blnShow = (intIndex > 0)
    
    If blnShow Then
        shpCircle(0).Left = lblElementName(intIndex).Left - shpCircle(0).Width
        shpCircle(0).Top = lblElementName(intIndex).Top - shpCircle(0).Height
        shpCircle(1).Left = lblElementName(intIndex).Left + (lblElementName(intIndex).Width - shpCircle(0).Width) / 2
        shpCircle(1).Top = shpCircle(0).Top
        shpCircle(2).Left = lblElementName(intIndex).Left + lblElementName(intIndex).Width
        shpCircle(2).Top = shpCircle(0).Top
        shpCircle(3).Left = shpCircle(2).Left
        shpCircle(3).Top = lblElementName(intIndex).Top + (lblElementName(intIndex).Height - shpCircle(3).Height) / 2
        shpCircle(4).Left = shpCircle(2).Left
        shpCircle(4).Top = lblElementName(intIndex).Top + lblElementName(intIndex).Height
        shpCircle(5).Left = shpCircle(1).Left
        shpCircle(5).Top = shpCircle(4).Top
        shpCircle(6).Left = shpCircle(0).Left
        shpCircle(6).Top = shpCircle(4).Top
        shpCircle(7).Left = shpCircle(0).Left
        shpCircle(7).Top = shpCircle(3).Top
    End If
    
    shpCircle(0).Visible = blnShow
    shpCircle(1).Visible = blnShow
    shpCircle(2).Visible = blnShow
    shpCircle(3).Visible = blnShow
    shpCircle(4).Visible = blnShow
    shpCircle(5).Visible = blnShow
    shpCircle(6).Visible = blnShow
    shpCircle(7).Visible = blnShow
End Sub

Private Sub SetCloseButton(ByVal intState As Integer, Optional ByVal blnSize As Boolean)
'参数：intState=0-正常,1-弹起,2-按下
    If intState = 0 Then
        lblClose.BackColor = picTitle.BackColor
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 0
    ElseIf intState = 1 Then
        lblClose.BackColor = &HD2BDB6
        lblClose.ForeColor = vbBlack
        lblClose.BorderStyle = 1
    ElseIf intState = 2 Then
        lblClose.BackColor = 11899525
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 1
    End If
    
    If blnSize Then
        lblClose.Width = 210
        lblClose.Height = 195
        lblClose.Left = picTitle.Width - lblClose.Width - 15
        lblClose.Top = (picTitle.Height - lblClose.Height) / 2
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.hwnd)
    End If
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetCloseButton(0)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lblBdr(0).Left = 0: lblBdr(0).Top = 0: lblBdr(0).Height = Me.ScaleHeight
    lblBdr(1).Left = Me.ScaleWidth - lblBdr(1).Width: lblBdr(1).Top = 0: lblBdr(1).Height = Me.ScaleHeight
    lblBdr(2).Left = 0: lblBdr(2).Top = 0: lblBdr(2).Width = Me.ScaleWidth
    lblBdr(3).Left = 0: lblBdr(3).Top = Me.ScaleHeight - lblBdr(3).Height: lblBdr(3).Width = Me.ScaleWidth
    
    picTitle.Left = lblBdr(0).Width + 30
    picTitle.Top = lblBdr(2).Height + 30
    picTitle.Width = Me.Width - picTitle.Left * 2
    picBak.Left = picTitle.Left
    picBak.Top = picTitle.Top + picTitle.Height + 30
    picBak.Width = picTitle.Width
        
    Call SetCloseButton(0, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngTop As Long, lngRight As Long
    
    If mfrmParent.WindowState <> 1 Then
        lngTop = Me.Top - mfrmParent.Top
        lngRight = Me.Left + Me.Width - (mfrmParent.Left + mfrmParent.Width)
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "NoticeBoardPostion", lngTop & "," & lngRight
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & mfrmParent.Name, "NoticeBoardSize", Me.Width & "," & Me.Height
    End If
    
    Set mrsBoard = Nothing
    mblnShow = False
    mblnFresh = False
End Sub

Private Sub lblBdr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngMaxW As Long
    If Button = 1 Then
        On Error Resume Next
        If mlngMaxW <= mfrmParent.Width / 2 Then
            lngMaxW = mlngMaxW
        Else
            lngMaxW = mfrmParent.Width / 2
        End If
        If Index = 0 Then
            If Me.Width - X > lngMaxW Then Exit Sub
            If Me.Width - X < mlngMinW Then Exit Sub
            Me.Left = Me.Left + X
            Me.Width = Me.Width - X
        ElseIf Index = 1 Then
            If Me.Width + X > lngMaxW Then Exit Sub
            If Me.Width + X < mlngMinW Then Exit Sub
            Me.Width = Me.Width + X
        ElseIf Index = 2 Then
'            If Me.Height - Y < mlngMinH Then Exit Sub
'            Me.Top = Me.Top + Y
'            Me.Height = Me.Height - Y
        ElseIf Index = 3 Then
'            If Me.Height + Y < mlngMinH Then Exit Sub
'            Me.Height = Me.Height + Y
        End If
        If Index = 0 Or Index = 1 Then Call Form_Resize
    End If
End Sub

Private Sub lblBdr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call SetCloseButton(2)
    End If
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And Y >= 0 And X <= lblClose.Width And Y <= lblClose.Height Then
        If Button = 1 Then
            Call SetCloseButton(2)
        Else
            Call SetCloseButton(1)
        End If
    Else
        Call SetCloseButton(1)
    End If
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And Y >= 0 And X <= lblClose.Width And Y <= lblClose.Height Then
        Call SetCloseButton(0)
        If mfrmParent.Visible Then
            mfrmParent.SetFocus
        End If
        Unload Me
    End If
End Sub

Private Sub lblElementCT_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Dim strTmp As String, strText As String
    Dim lngWidth As Long
    Dim arrText
    Dim i As Long
    
    strText = lblElementCT(Index).Caption
    If PicElementCT(Index).TextWidth(strText) <= 4000 Then
        strInfo = strText
    Else
        i = 1
        Do While True
            lngWidth = lngWidth + PicElementCT(Index).TextWidth(Mid(strText, i, 1))
            If lngWidth <= 4000 Then
                strTmp = strTmp & Mid(strText, i, 1)
            Else
                If strInfo = "" Then
                    strInfo = strInfo & strTmp
                Else
                    strInfo = strInfo & vbCrLf & strTmp
                End If
                strTmp = Mid(strText, i, 1)
                lngWidth = PicElementCT(Index).TextWidth(strTmp)
            End If
            If i = Len(strText) Then
                If strInfo = "" Then
                    strInfo = strInfo & strTmp
                Else
                    strInfo = strInfo & vbCrLf & strTmp
                End If
                Exit Do
            End If
            i = i + 1
        Loop
    End If
    Call zlCommFun.ShowTipInfo(PicElementCT(Index).hwnd, strInfo, True)
End Sub

Private Sub lblElementName_Click(Index As Integer)
    Dim intDo As Integer, intCount As Integer
    
    intCount = lblElementName.Count - 1
    For intDo = 1 To intCount
        lblElementName(intDo).BackStyle = 0
    Next
    picBak.Tag = lblElementName(Index).Tag
    lblElementName(Index).BackStyle = 1
    Call SetShape
End Sub

Private Sub lblElementName_DblClick(Index As Integer)
    Dim strBeds As String
    mrsBoard.Filter = "ID=" & Val(picBak.Tag)
    If mrsBoard.RecordCount = 0 Then Exit Sub
    strBeds = "" & mrsBoard!内容
    RaiseEvent ItemClick(strBeds)
    
    If mfrmParent.Visible And mfrmParent.Enabled Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.hwnd)
    End If
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub picBak_Resize()
    Dim intCount As Integer, intDel As Integer
    Dim arrLeft, arrRight
    Dim i As Integer, j As Integer
    On Error Resume Next
    arrLeft = Array()
    arrRight = Array()
    intCount = lblElementName.Count - 1
    For intDel = 1 To intCount
        If InStr(1, PicElementCT(intDel).Tag, ",") <> 0 Then
            If Val(Split(PicElementCT(intDel).Tag, ",")(1)) = 1 Then
                ReDim Preserve arrLeft(UBound(arrLeft) + 1)
                arrLeft(UBound(arrLeft)) = intDel & "," & Val(Split(PicElementCT(intDel).Tag, ",")(0))
            Else
                ReDim Preserve arrRight(UBound(arrRight) + 1)
                arrRight(UBound(arrRight)) = intDel & "," & Val(Split(PicElementCT(intDel).Tag, ",")(0))
                lblElementName(intDel).Left = picBak.Width - lblElementName(intDel).Width - 1000
                PicElementCT(intDel).Left = lblElementName(intDel).Left + lblElementName(intDel).Width + 60
            End If
        End If
    Next
    '重新整理要素位置
    For i = 0 To UBound(arrLeft)
        For j = 0 To UBound(arrRight)
            If Split(arrLeft(i), ",")(1) = Split(arrRight(j), ",")(1) Then
                lblElementCT(Val(Split(arrLeft(i), ",")(0))).Width = lblElementName(Val(Split(arrRight(j), ",")(0))).Left - PicElementCT(Val(Split(arrLeft(i), ",")(0))).Left - 60
                PicElementCT(Val(Split(arrLeft(i), ",")(0))).Width = lblElementCT(Val(Split(arrLeft(i), ",")(0))).Width
                Exit For
            End If
        Next j
    Next i
    picBak.PaintPicture picBak.Picture, 0, 0, picBak.Width
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.hwnd)
        If mfrmParent.Visible Then
            mfrmParent.SetFocus
        End If
    End If
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetCloseButton(0)
End Sub

Private Sub tmrFresh_Timer()
    If mblnFresh = False Then Exit Sub
    mblnFresh = False
    RefreshData
    Call SetShape
End Sub
