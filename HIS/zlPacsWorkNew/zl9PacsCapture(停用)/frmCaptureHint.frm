VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmCaptureHint 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "当前图像"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   Icon            =   "frmCaptureHint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1080
      Top             =   600
   End
   Begin DicomObjects.DicomViewer dcmViewer 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _Version        =   262147
      _ExtentX        =   5741
      _ExtentY        =   3625
      _StockProps     =   35
   End
End
Attribute VB_Name = "frmCaptureHint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Enum THintPos
    hpCus = 0
    hpLT = 1
    hpRT = 2
    hpRB = 3
    hpLB = 4
    hpCen = 5
End Enum

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1


Private mlngTransparentValue As Long
Private mlngState As Long
Private mblnMouseDown As Boolean
Private mhpPos As THintPos
Private mobjParent As Object
Private mstrImgFile As String
Private mblnBeep As Boolean
Private mlngLoadState As Long



Public Sub ShowCaptureHint(ByVal strImgFile As String, ByVal blnBeep As Boolean, ByVal hpPos As THintPos, owner As Object)
On Error Resume Next
    mlngTransparentValue = IIf(mlngLoadState <> 1, 20, 25)
    mlngState = 0
    mhpPos = hpPos
    mstrImgFile = strImgFile
    mblnBeep = blnBeep

    'If strImgFile = "" And blnBeep Then Call Beep(2500, 500)
    If blnBeep Then Call Beep(2500, 500)
    
    If strImgFile = "" Then Exit Sub
    
    Set mobjParent = owner
    
    If mlngLoadState <> 1 Then
        Load Me
    Else
        '如果状态不为1，说明窗体已经被加载，需要手动刷新提示图像
        Timer1.Enabled = False
        
        Call dcmViewer.Images.Clear
        Call dcmViewer.Images.ReadFile(mstrImgFile)

        Timer1.Enabled = True
    End If
    
    err.Clear
End Sub


Private Sub dcmViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    mblnMouseDown = True
    
   
    err.Clear
End Sub

Private Sub dcmViewer_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    mblnMouseDown = False
    
    err.Clear
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim sty As Long
    
    Dim lngMonitorIndex As Long
    Dim lngCurScreenRight As Long
    Dim lngCurScreenBottom As Long
    Dim lngCurScreenLeft As Long
    Dim lngCurScreenTop As Long
    
    mlngLoadState = 1
    
    mblnMouseDown = False
    lngMonitorIndex = GetMonitorIndex(mobjParent.hWnd)
    
    lngCurScreenRight = ScaleX(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Right, vbPixels, vbTwips)
    lngCurScreenBottom = ScaleY(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Bottom, vbPixels, vbTwips)
    lngCurScreenLeft = ScaleX(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Left, vbPixels, vbTwips)
    lngCurScreenTop = ScaleY(gmonitors(lngMonitorIndex + 1).monitorInf.rcMonitor.Top, vbPixels, vbTwips)
    
    Select Case mhpPos
        Case THintPos.hpCus
        Case THintPos.hpCen
        Case THintPos.hpLT
        Case THintPos.hpRT
        Case THintPos.hpRB
            Me.Left = lngCurScreenRight - Me.Width
            Me.Top = lngCurScreenBottom - Me.Height
        Case THintPos.hpLB
            Me.Left = lngCurScreenLeft
            Me.Top = lngCurScreenBottom - Me.Height
            
    End Select
    
    Call dcmViewer.Images.Clear
    Call dcmViewer.Images.ReadFile(mstrImgFile)
    
    

    sty = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    sty = sty Or WS_EX_LAYERED
    
    SetWindowLong Me.hWnd, GWL_EXSTYLE, sty
'    SetLayeredWindowAttributes Me.hWnd, 0, mlngTransparentValue, LWA_ALPHA '该处不需设置透明度，否则会造成闪烁
    
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, &H1 Or &H2 Or &H10 Or &H40 '将窗口置顶
    
    Timer1.Enabled = True
    
    err.Clear
End Sub


Private Sub Form_Resize()
On Error Resume Next
    dcmViewer.Left = 0
    dcmViewer.Top = 0
    dcmViewer.Height = Me.ScaleHeight
    dcmViewer.Width = Me.ScaleWidth
    
    err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    mlngLoadState = 2
    
    err.Clear
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If mblnMouseDown Then Exit Sub
    
    If mlngState = 0 And mlngTransparentValue > 255 Then
        mlngState = 1
        mlngTransparentValue = 1
    ElseIf mlngState = 1 And mlngTransparentValue > 50 Then
        mlngState = 2
        mlngTransparentValue = 255
    ElseIf mlngState = 2 And mlngTransparentValue < 20 Then
        Timer1.Enabled = False
        Unload Me
        Exit Sub
    End If
    
    If mlngState = 0 Then
        'If mblnBeep And mlngTransparentValue = 20 Then Call Beep(2500, 500)
        
        mlngTransparentValue = mlngTransparentValue + 5
    ElseIf mlngState = 1 Then
        mlngTransparentValue = mlngTransparentValue + 1
    Else
        mlngTransparentValue = mlngTransparentValue - 5
    End If

    If mlngState <> 1 Then SetLayeredWindowAttributes Me.hWnd, 0, mlngTransparentValue, LWA_ALPHA

    err.Clear
End Sub


