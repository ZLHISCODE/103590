VERSION 5.00
Begin VB.Form frmCardSel 
   Caption         =   "��Ƭѡ����"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9795
   Icon            =   "frmCardSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9795
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPatiList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   1800
      ScaleHeight     =   3000
      ScaleWidth      =   2370
      TabIndex        =   1
      Top             =   1995
      Width           =   2370
      Begin VB.PictureBox picPati 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   2190
         Index           =   0
         Left            =   240
         Picture         =   "frmCardSel.frx":6852
         ScaleHeight     =   2190
         ScaleWidth      =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   1800
         Begin VB.Label lblSource 
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   10
            Top             =   120
            Width           =   855
         End
         Begin VB.Image imgMark 
            Height          =   300
            Index           =   0
            Left            =   130
            Picture         =   "frmCardSel.frx":A219
            Stretch         =   -1  'True
            Top             =   110
            Width           =   300
         End
         Begin VB.Label lblName 
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   345
            TabIndex        =   8
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblSex 
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   7
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblAge 
            BackStyle       =   0  'Transparent
            Caption         =   "25��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   6
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblDisease 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "��˲�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   1440
            Width           =   1400
         End
         Begin VB.Label lblTime 
            BackStyle       =   0  'Transparent
            Caption         =   "2015/01/01 00:00"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   4
            Top             =   1815
            Width           =   1575
         End
      End
      Begin VB.VScrollBar HScr 
         Height          =   5295
         LargeChange     =   10
         Left            =   14280
         Max             =   100
         SmallChange     =   5
         TabIndex        =   2
         Top             =   -120
         Width           =   255
      End
   End
   Begin VB.Frame fraHead 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   300
      TabIndex        =   0
      Top             =   285
      Width           =   16935
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "��˫��ѡ��һ����Ŀ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   2100
      End
   End
   Begin VB.Image imgState 
      Height          =   300
      Index           =   0
      Left            =   6000
      Picture         =   "frmCardSel.frx":3EEF1
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgState 
      Height          =   300
      Index           =   1
      Left            =   6480
      Picture         =   "frmCardSel.frx":73BC9
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   1
      Left            =   3360
      Picture         =   "frmCardSel.frx":A88A1
      Top             =   3600
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   0
      Left            =   1200
      Picture         =   "frmCardSel.frx":AC268
      Top             =   3600
      Visible         =   0   'False
      Width           =   1800
   End
End
Attribute VB_Name = "frmCardSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdblScaleHeight  As Double
Private mlngSelIndex As Long        'ѡ��ķ�����
Private mlngID As Long              'ѡ��ķ�����
Private mlngCount As Long           '������������
Private mlngPageCount As Long       'һҳ�Ŀ�Ƭ����
Private mlngCardCount As Long       '��Ƭ����
Private mlngReportCount As Long     'ʵ����ʾ�ķ���������
Private mrsData As ADODB.Recordset

Public Function ShowMe(ByVal rsData As ADODB.Recordset, ByRef frmParent As Object) As Long
    
    Set mrsData = zlDatabase.CopyNewRec(rsData)
    mlngCount = mrsData.RecordCount
    Call AdjustCardsPosition
    Me.Show 1, frmParent
    ShowMe = mlngID
End Function

Private Sub LoadPatiCard(ByVal intIndex As Integer)
    If intIndex = 0 Then
        Call SetPicVisible(0, True)
        Exit Sub
    End If
    
    Load picPati(intIndex)
    With picPati(intIndex)
        .Width = picPati(0).Width
        .Height = picPati(0).Height
        .Picture = Nothing
        .Visible = True
        .ZOrder 1
    End With

    Load lblName(intIndex)
    Set lblName(intIndex).Container = picPati(intIndex)
    lblName(intIndex).Visible = True
    lblName(intIndex).FontSize = lblName(0).FontSize
    lblName(intIndex).Left = lblName(0).Left
    lblName(intIndex).Top = lblName(0).Top
    lblName(intIndex).Width = lblName(0).Width
    lblName(intIndex).Height = lblName(0).Height
    lblName(intIndex).Caption = ""
    lblName(intIndex).ZOrder 0
    
    Load lblAge(intIndex)
    Set lblAge(intIndex).Container = picPati(intIndex)
    lblAge(intIndex).Visible = True
    lblAge(intIndex).FontSize = lblAge(0).FontSize
    lblAge(intIndex).Left = lblAge(0).Left
    lblAge(intIndex).Top = lblAge(0).Top
    lblAge(intIndex).Width = lblAge(0).Width
    lblAge(intIndex).Height = lblAge(0).Height
    lblAge(intIndex).Caption = ""
    lblAge(intIndex).ZOrder 0
    
    Load lblSex(intIndex)
    Set lblSex(intIndex).Container = picPati(intIndex)
    lblSex(intIndex).Visible = True
    lblSex(intIndex).FontSize = lblSex(0).FontSize
    lblSex(intIndex).Left = lblSex(0).Left
    lblSex(intIndex).Top = lblSex(0).Top
    lblSex(intIndex).Width = lblSex(0).Width
    lblSex(intIndex).Height = lblSex(0).Height
    lblSex(intIndex).Caption = ""
    lblSex(intIndex).ZOrder 0
    
    Load lblDisease(intIndex)
    Set lblDisease(intIndex).Container = picPati(intIndex)
    lblDisease(intIndex).Visible = True
    lblDisease(intIndex).FontSize = lblDisease(0).FontSize
    lblDisease(intIndex).Left = lblDisease(0).Left
    lblDisease(intIndex).Top = lblDisease(0).Top
    lblDisease(intIndex).Width = lblDisease(0).Width
    lblDisease(intIndex).Height = lblDisease(0).Height
    lblDisease(intIndex).Caption = ""
    lblDisease(intIndex).ZOrder 0
    
    Load lblTime(intIndex)
    Set lblTime(intIndex).Container = picPati(intIndex)
    lblTime(intIndex).Visible = True
    lblTime(intIndex).FontSize = lblTime(0).FontSize
    lblTime(intIndex).Left = lblTime(0).Left
    lblTime(intIndex).Top = lblTime(0).Top
    lblTime(intIndex).Width = lblTime(0).Width
    lblTime(intIndex).Height = lblTime(0).Height
    lblTime(intIndex).Caption = ""
    lblTime(intIndex).ZOrder 0
    
    Load lblSource(intIndex)
    Set lblSource(intIndex).Container = picPati(intIndex)
    lblSource(intIndex).Visible = True
    lblSource(intIndex).FontSize = lblSource(0).FontSize
    lblSource(intIndex).Left = lblSource(0).Left
    lblSource(intIndex).Top = lblSource(0).Top
    lblSource(intIndex).Width = lblSource(0).Width
    lblSource(intIndex).Height = lblSource(0).Height
    lblSource(intIndex).Caption = ""
    lblSource(intIndex).ZOrder 0
    
    Load imgMark(intIndex)
    Set imgMark(intIndex).Container = picPati(intIndex)
    imgMark(intIndex).Visible = True
    imgMark(intIndex).Left = imgMark(0).Left
    imgMark(intIndex).Top = imgMark(0).Top
    imgMark(intIndex).Width = imgMark(0).Width
    imgMark(intIndex).Height = imgMark(0).Height
    imgMark(intIndex).ZOrder 0
End Sub

Private Sub SetPicVisible(ByVal Index As Long, ByVal blnVisible As Boolean)
    lblName(Index).Visible = blnVisible
    lblAge(Index).Visible = blnVisible
    lblSex(Index).Visible = blnVisible
    lblDisease(Index).Visible = blnVisible
    lblTime(Index).Visible = blnVisible
    lblSource(Index).Visible = blnVisible
    picPati(Index).Visible = blnVisible
End Sub

Private Sub AdjustCardsPosition(Optional ByVal lngY As Long = 0)
    Dim lngRowCount As Long
    Dim lngColCount As Long
    Dim lngX As Long, lngStart As Long
    Dim lngShowed As Long
    Dim blnAdjust As Boolean
    Dim i As Long
   
    blnAdjust = (lngY = 0)
    lngX = 50
    'ÿһ���ж��ٸ�
    lngRowCount = Abs((picPatiList.Width - HScr.Width - 50) / (picPati(0).Width + 15) - 0.5)
    If lngRowCount < 1 Then lngRowCount = 1
    lngColCount = Abs(picPatiList.Height / picPati(0).Height + 1)
    mlngPageCount = lngColCount * lngRowCount
    
    Call zlControl.FormLock(Me.hwnd)
    '���ؿ�Ƭ
    If mlngCardCount < mlngPageCount Then
        For i = mlngCardCount + 1 To mlngPageCount
             Call LoadPatiCard(i)
        Next
        mlngCardCount = mlngPageCount
    End If
    '����������֮������λ��
    
    If lngY <> 0 Then
        lngStart = CLng((-1 * lngY) / picPati(0).Height - 0.5) * lngRowCount
        If lngStart < 0 Then lngStart = 0
        lngY = lngY + CLng((-1 * lngY) / picPati(0).Height - 0.5) * picPati(0).Height
    End If
    
    '���ؿ�Ƭ�������Ϣ
    Call LoadPatiCardInfo(lngStart)
    
    '���ÿ�Ƭ�Ŀɼ���
    For i = 0 To mlngReportCount - 1
        Call SetPicVisible(i, True)
    Next
    If mlngReportCount - 1 < mlngCardCount Then
        For i = mlngReportCount To mlngCardCount
            Call SetPicVisible(i, False)
        Next
    End If
    
    '����ÿ�ſ�Ƭ��λ��
    If mlngPageCount > 0 Then
        For i = 0 To mlngPageCount
            If i Mod (lngRowCount) = 0 And i <> 0 Then
                lngX = 50
                lngY = lngY + picPati(0).Height
            End If
            picPati(i).Left = lngX
            picPati(i).Top = lngY
            lngX = lngX + picPati(0).Width
        Next
    End If
    mdblScaleHeight = picPati(0).Height * IIF(mlngCount / lngRowCount > CLng(mlngCount / lngRowCount), CLng(mlngCount / lngRowCount + 0.5), CLng(mlngCount / lngRowCount))
    If blnAdjust Then
        HScr.value = 0
        HScr.Visible = mdblScaleHeight > picPatiList.Height
    End If
    
    Call zlControl.FormLock(0)
End Sub


Private Sub UnloadControls(ByVal blnUnload As Boolean)
    Dim j As Long
    For j = picPati.Count - 1 To 1 Step -1
        If blnUnload Then
            Unload lblName(j)
            Unload lblAge(j)
            Unload lblSex(j)
            Unload lblDisease(j)
            Unload lblTime(j)
            Unload lblSource(j)
            Unload imgMark(j)
            Unload picPati(j)
        Else
            Call SetPicVisible(j, False)
        End If
    Next
    Call SetPicVisible(0, False)
    mlngSelIndex = -1
End Sub

Private Sub LoadPatiCardInfo(ByVal lngStart As Long)
    Dim i As Long
    If mrsData.RecordCount > 0 Then
        Call mrsData.Move(lngStart, adBookmarkFirst)
        Do While Not mrsData.EOF
            If i >= mlngPageCount Then Exit Do
            picPati(i).Tag = mrsData!ID
            picPati(i).Picture = imgCardBack(1).Picture
            
            If Val(mrsData!��¼״̬ & "") = 0 Then
                imgMark(i).Visible = False
                imgMark(i).Tag = "δ��"
            Else
                imgMark(i).Visible = True
                imgMark(i).Picture = imgState(1).Picture
                imgMark(i).Tag = "����"
            End If
        
            lblName(i).Caption = mrsData!���� & ""
            lblName(i).Tag = mrsData!������ & ""
            lblAge(i).Caption = mrsData!���� & ""
            lblSex(i).Caption = mrsData!�Ա� & ""
            lblDisease(i).Caption = mrsData!���� & ""
            If IsDate(mrsData!��¼ʱ�� & "") Then
                 lblTime(i).Caption = Format(mrsData!��¼ʱ�� & "", "yyyy-mm-dd HH:MM")
            End If
            lblSource(i).Caption = mrsData!ִ�п��� & ""
            mrsData.MoveNext
            i = i + 1
        Loop
    End If
    mlngReportCount = i
End Sub
    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = HScr.value
    lngMin = HScr.Min
    lngMax = HScr.Max
    
    If KeyCode = vbKeyPageDown Then '��
        If Between(lngCur + (lngMax - lngMin) / 100, lngMin, lngMax) Then
            HScr.value = lngCur + (lngMax - lngMin) / 100
        Else
            HScr.value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '��
        If Between(lngCur - (lngMax - lngMin) / 100, lngMin, lngMax) Then
            HScr.value = lngCur - (lngMax - lngMin) / 100
        Else
            HScr.value = lngMin
        End If
    End If
End Sub

Private Sub Form_Activate()
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_Load()
    mlngID = 0
    mlngReportCount = 0
    mlngCardCount = 0
    Me.BorderStyle = 3
    Me.Caption = "�����ļ�ѡ��"
    lblnote.Visible = True
    If mlngCount = 2 Then
        Me.Width = 4030
    Else
        Me.Width = 5850
    End If
    Me.Height = 3200
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraHead.Move 0, 0, Me.ScaleWidth, 400
    picPatiList.Move 0, 400, Me.ScaleWidth, Me.ScaleHeight - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnloadControls(True)
End Sub

Private Sub HScr_Change()
    Dim lngMove As Long
    Dim lngY As Long
    If Not HScr.Visible Then Exit Sub
    '���㵥������
    lngMove = CLng((mdblScaleHeight - picPatiList.Height) / 100)

    If lngMove < 0 Then lngMove = 0
    lngY = -1 * HScr.value * lngMove
    If lngY >= 0 And lngY < 100 Then lngY = 0
    Call AdjustCardsPosition(lngY)
End Sub

Private Sub lblAge_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblAge(Index).Left + X, lblAge(Index).Top + Y)
End Sub

Private Sub lblDisease_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(picPati(Index).hwnd, "�ļ�����" & lblDisease(Index).Caption)
End Sub

Private Sub lblDisease_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblDisease(Index).Left + X, lblDisease(Index).Top + Y)
End Sub

Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblName(Index).Left + X, lblName(Index).Top + Y)
End Sub

Private Sub lblName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(picPati(Index).hwnd, "������" & lblName(Index).Caption)
End Sub

Private Sub lblSource_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call zlCommFun.ShowTipInfo(picPati(Index).hwnd, "ִ�п��ң�" & lblSource(Index).Caption)
End Sub

Private Sub picPati_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(picPati(Index).hwnd, "״̬��" & imgMark(Index).Tag)
End Sub

Private Sub lblSex_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblSex(Index).Left + X, lblSex(Index).Top + Y)
End Sub

Private Sub lblSource_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblSource(Index).Left + X, lblSource(Index).Top + Y)
End Sub

Private Sub lblTime_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblTime(Index).Left + X, lblTime(Index).Top + Y)
End Sub

Private Sub lblAge_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblDisease_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblName_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblSex_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblSource_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblTime_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(picPati(Index).hwnd, "����ʱ�䣺" & lblTime(Index).Caption)
End Sub

Private Sub picPati_DblClick(Index As Integer)
    If mlngSelIndex < 0 Then Exit Sub
    mlngID = CLng(picPati(mlngSelIndex).Tag)
    Unload Me
End Sub

Private Sub picPati_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mlngSelIndex Then
        If mlngSelIndex >= 0 Then
            If lblName(mlngSelIndex).Tag <> "" Then
                picPati(mlngSelIndex).Picture = imgCardBack(1).Picture
            Else
                picPati(mlngSelIndex).Picture = imgCardBack(1).Picture
            End If
        End If
        mlngSelIndex = Index
        If lblName(mlngSelIndex).Tag <> "" Then
            picPati(mlngSelIndex).Picture = imgCardBack(0).Picture
        Else
            picPati(mlngSelIndex).Picture = imgCardBack(0).Picture
        End If
    End If
End Sub

Private Sub picPatiList_Resize()
    On Error Resume Next
    HScr.Move picPatiList.ScaleWidth - HScr.Width, 0, HScr.Width, picPatiList.ScaleHeight
    If Me.Visible Then Call AdjustCardsPosition
End Sub
