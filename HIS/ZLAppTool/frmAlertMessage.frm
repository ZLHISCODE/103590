VERSION 5.00
Begin VB.Form frmAlertMessage 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCaption 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   390
      ScaleHeight     =   285
      ScaleWidth      =   2940
      TabIndex        =   4
      Top             =   915
      Width           =   2940
      Begin VB.Image Image1 
         Height          =   195
         Left            =   45
         MouseIcon       =   "frmAlertMessage.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmAlertMessage.frx":030A
         Top             =   45
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒消息"
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   60
         Width           =   720
      End
      Begin VB.Image imgClose 
         Height          =   195
         Left            =   2700
         MouseIcon       =   "frmAlertMessage.frx":0590
         MousePointer    =   99  'Custom
         Picture         =   "frmAlertMessage.frx":089A
         Top             =   45
         Width           =   195
      End
   End
   Begin VB.PictureBox picBackground 
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   390
      ScaleHeight     =   1515
      ScaleWidth      =   2940
      TabIndex        =   0
      Top             =   1290
      Width           =   2940
      Begin VB.Label lblReport 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品付款单据"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   450
         MouseIcon       =   "frmAlertMessage.frx":0B20
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lblState 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "第1条 共3条"
         Height          =   180
         Left            =   435
         TabIndex        =   2
         Top             =   1245
         Width           =   990
      End
      Begin VB.Image imgDown 
         Height          =   240
         Left            =   2490
         MouseIcon       =   "frmAlertMessage.frx":0E2A
         MousePointer    =   99  'Custom
         Picture         =   "frmAlertMessage.frx":1134
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgUp 
         Height          =   240
         Left            =   2085
         MouseIcon       =   "frmAlertMessage.frx":14BE
         MousePointer    =   99  'Custom
         Picture         =   "frmAlertMessage.frx":17C8
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgMessage 
         Height          =   240
         Left            =   75
         Picture         =   "frmAlertMessage.frx":1B52
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lblMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "查询在今后一段时间内将失效的药品的当前库存，以便及时处理。"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   435
         TabIndex        =   1
         Top             =   105
         Width           =   2370
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5475
      Top             =   570
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   1095
   End
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5430
      Top             =   105
   End
   Begin VB.Shape shp 
      BorderColor     =   &H80000003&
      FillColor       =   &H00404040&
      Height          =   1845
      Left            =   0
      Top             =   0
      Width           =   2595
   End
End
Attribute VB_Name = "frmAlertMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' API Declarations
Private Declare Function GetSystemMetrics& Lib "User32" (ByVal nIndex As Long)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Constants
Private Const SM_CXFULLSCREEN = 16   ' Width of window client area
Private Const SM_CYFULLSCREEN = 17   ' Height of window client area
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10

Private mblnStartUp As Boolean
Private mrsMessage As New ADODB.Recordset
Private mfrmMain As Object
Private clsGradient As New clsGradient
Private mlngMaxHeight As Long
Private AlertIndex As Long
Private mstrLastFile As String

Public Function PlayWave(lngKey As Long) As String
    '功能:将资源文件中的指定资源生成磁盘文件
    '参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
    '返回:生成文件名

    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255
    Dim strR As String

    On Error Resume Next

    arrData = LoadResData(lngKey, "WAVE")
    intFile = FreeFile

    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & "zlTempSoundFile" & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile

    '直接从字节数组进行播放声音
    Call sndPlaySound(strR, SND_NODEFAULT Or SND_ASYNC)
    If Dir(mstrLastFile) <> "" Then
        Kill mstrLastFile
    End If
    Kill strR
    mstrLastFile = strR
End Function

Private Sub ShowMessage(ByVal rs As ADODB.Recordset)

    If IsNull(rs("提醒声音").Value) = False Then

        Call PlayWave(rs("提醒声音").Value)

    End If

    lblMessage.Caption = rs("提醒内容").Value

    If lblMessage.Height > 705 Then lblMessage.Height = 705

    lblReport.Caption = ""
    lblReport.Tag = ""

    If IsNull(rs("提醒报表").Value) = False Then
        lblReport.Caption = rs("提醒报表").Value
        lblReport.Tag = rs("报表系统").Value & ";" & rs("模块").Value

    End If
    lblReport.Visible = (lblReport.Caption <> "")

    lblReport.Top = lblMessage.Top + lblMessage.Height + lblReport.Height

    lblState.Visible = (rs.RecordCount > 1)
    imgDown.Visible = lblState.Visible
    imgUp.Visible = lblState.Visible

    lblState.Caption = "第" & rs.AbsolutePosition & "条 共" & rs.RecordCount & "条"
    
    'Call picBackground_Paint
End Sub

Private Sub Form_Activate()

    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False

    tmrAlert.Interval = Val(zlDatabase.GetPara("自动消息停留时间")) * 1000
    
'    With clsGradient
'        .Angle = 130
'        .Color1 = RGB(255, 255, 255)
'        .Color2 = RGB(128, 230, 255)
'        '.Color2 = 14737632
'        .Color2 = &HFED7BC
'        .Draw picBackground
'    End With

'    picBackground.Refresh

    '显示第一条信息
    If mrsMessage.BOF = False Then Call ShowMessage(mrsMessage)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Dir(mstrLastFile) <> "" Then
        Kill mstrLastFile
    End If
End Sub

Private Sub imgClose_Click()
    tmrClose.Interval = 1
    tmrClose.Enabled = True
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        imgClose.Left = imgClose.Left + 15
        imgClose.Top = imgClose.Top + 15
    End If
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        imgClose.Left = imgClose.Left - 15
        imgClose.Top = imgClose.Top - 15
    End If
End Sub

Private Sub imgDown_Click()

    On Error GoTo errHand

    mrsMessage.MoveNext

    If mrsMessage.EOF = False Then

        Call ShowMessage(mrsMessage)

    End If

errHand:

End Sub

Private Sub imgDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub

Private Sub imgMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub

Private Sub imgUp_Click()
    On Error GoTo errHand
    
    mrsMessage.MovePrevious
    If mrsMessage.BOF = False Then

        Call ShowMessage(mrsMessage)

    End If
errHand:
    
End Sub

Private Sub imgUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
'        imgUp.Left = imgUp.Left + 15
'        imgUp.Top = imgUp.Top + 15
    End If
End Sub

Private Sub imgUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub

Private Sub imgUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        imgUp.Left = imgUp.Left - 15
'        imgUp.Top = imgUp.Top - 15
'    End If
End Sub

Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
'        imgDown.Left = imgDown.Left + 15
'        imgDown.Top = imgDown.Top + 15
    End If
End Sub

Private Sub imgDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
'        imgDown.Left = imgDown.Left - 15
'        imgDown.Top = imgDown.Top - 15
    End If
End Sub


Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub

Private Sub lblReport_Click()

    Call mfrmMain.RunModual(Val(Split(lblReport.Tag, ";")(0)), Val(Split(lblReport.Tag, ";")(1)), "")

End Sub

Private Sub lblReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub

Private Sub lblState_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub


Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrAlert.Enabled = False
    tmrAlert.Enabled = True
End Sub

Private Sub picBackground_Paint()
    Call DrawColorToColor(picBackground, &HFFFFFF, &HFFC0C0)
    zlControl.PicShowFlat picBackground, 1
End Sub

Private Sub picCaption_Paint()
    Call DrawColorToColor(picCaption, &HFFFFFF, &HFFC0C0)
    zlControl.PicShowFlat picCaption, 1
End Sub

Private Sub tmrAlert_Timer()

    '显示信息记录

    Call imgDown_Click

    If mrsMessage.EOF Then

        '已经显示完，则退出

        tmrAlert.Enabled = False

        tmrClose.Interval = 1
        tmrClose.Enabled = True

    End If
End Sub

Private Sub tmrClose_Timer()

    Dim lngHeight As Long
    Dim lngTop As Long

    lngHeight = Me.Height

    If lngHeight > 90 Then

        lngHeight = lngHeight - 30
        lngTop = Me.Top + 30

        On Error Resume Next

        MoveWindow Me.hWnd, Me.Left / 15, lngTop / 15, Me.Width / 15, lngHeight / 15, 1
        SetWindowPos Me.hWnd, -1, Me.Left / 15, lngTop / 15, Me.Width / 15, lngHeight / 15, &H10 Or &H1

    Else
        Unload Me
    End If
End Sub

Private Sub tmrOpen_Timer()

    Dim lngHeight As Long
    Dim lngNewHeight As Long
    Dim lngScaleY As Long

    Dim lngH As Long
    Dim lngTop As Long

    lngScaleY = Me.Height - Me.ScaleHeight
    lngHeight = Me.Height

    If lngHeight < mlngMaxHeight + lngScaleY Then
        lngNewHeight = lngHeight + 30

        If lngNewHeight > mlngMaxHeight + lngScaleY Then lngNewHeight = mlngMaxHeight + lngScaleY

        lngH = Me.Height + (lngNewHeight - lngHeight)
        lngTop = Me.Top - (lngNewHeight - lngHeight)

        On Error Resume Next

        MoveWindow Me.hWnd, Me.Left / 15, lngTop / 15, Me.Width / 15, lngH / 15, 1
        SetWindowPos Me.hWnd, -1, Me.Left / 15, lngTop / 15, Me.Width / 15, lngH / 15, &H10 Or &H1

    Else

        tmrOpen.Enabled = False
        tmrAlert.Enabled = True

    End If


End Sub

Public Function ShowAlert(ByVal strMessage As String, ByVal frmMain As Object) As Boolean
    '---------------------------------------------------------------------------------------
    '功能： 弹出提示窗口
    '参数： strMessage-信息集
    '       lngDelay-每条信息停止的时间
    '---------------------------------------------------------------------------------------
    Dim lngScreenX As Long
    Dim lngScreenY As Long
    Dim lngScaleX As Long
    Dim lngScaleY As Long

    mblnStartUp = True
    
    Set mfrmMain = frmMain
    
    Set mrsMessage = New ADODB.Recordset
    With mrsMessage
        .Fields.Append "提醒内容", adVarChar, 1000
        .Fields.Append "提醒声音", adBigInt
        .Fields.Append "提醒报表", adVarChar, 255
        .Fields.Append "报表系统", adBigInt
        .Fields.Append "系统", adBigInt
        .Fields.Append "模块", adBigInt
        .Fields.Append "提醒窗口", adTinyInt
        .Open
    End With
    
    '分析字串
    '格式:第一条;0;;100;0|第二条;0;;100;0
    
    Dim varTmp As Variant
    Dim varTmp2 As Variant
    Dim lngCount As Long
    
    varTmp = Split(strMessage, "[INFOITEM-BEGIN']")
    
    For lngCount = 0 To UBound(varTmp)
    
        mrsMessage.AddNew
        
        varTmp2 = Split(varTmp(lngCount), "[''']")
        
        mrsMessage("提醒内容").Value = varTmp2(0)
        mrsMessage("提醒声音").Value = Val(varTmp2(1))
        mrsMessage("提醒报表").Value = varTmp2(2)
        mrsMessage("系统").Value = Val(varTmp2(3))
        mrsMessage("模块").Value = Val(varTmp2(4))
        mrsMessage("提醒窗口").Value = Val(varTmp2(6))
        
        If UBound(varTmp2) >= 7 Then
            mrsMessage("报表系统").Value = Val(varTmp2(7))
        Else
            mrsMessage("报表系统").Value = Val(varTmp2(3))
        End If
        
    Next
    
    If mrsMessage.RecordCount > 0 Then
        mrsMessage.Filter = "提醒窗口=1"
        If mrsMessage.RecordCount > 0 Then
            mrsMessage.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
                
    lngScreenX = GetSystemMetrics(SM_CXFULLSCREEN)
    lngScreenY = GetSystemMetrics(SM_CYFULLSCREEN)

    lngScaleX = Me.Width - Me.ScaleWidth
    lngScaleY = Me.Height - Me.ScaleHeight
    
    mlngMaxHeight = picBackground.Height + picCaption.Height + 15
    
    shp.Top = 0
    shp.Left = 0
    shp.Width = picBackground.Width + lngScaleX + 15
    shp.Height = mlngMaxHeight
    
    With picCaption
        .Left = 15
        .Top = 15
        .Width = picBackground.Width
    End With
    
    With picBackground
        .Left = 15
        .Top = picCaption.Top + picCaption.Height
    End With
    
    Me.Height = 90
    Me.Width = picBackground.Width + lngScaleX + 15
    Me.Left = lngScreenX * Screen.TwipsPerPixelX - Me.Width - 15
    Me.Top = (lngScreenY * Screen.TwipsPerPixelY) + 160

    On Error Resume Next
    Call Form_Activate

    ShowWindow Me.hWnd, 4
    SetWindowPos Me.hWnd, -1, Me.Left / 15, Me.Top / 15, Me.Width / 15, Me.Height / 15, &H10 Or &H1

    tmrOpen.Enabled = True

End Function

