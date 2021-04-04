VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmCaptureHint 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "当前图像"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   Icon            =   "frmCaptureHint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMedia 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   1080
      Picture         =   "frmCaptureHint.frx":0E42
      ScaleHeight     =   2805
      ScaleWidth      =   2610
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   2040
      Top             =   600
   End
   Begin DicomObjects.DicomViewer dcmViewer 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      _Version        =   262147
      _ExtentX        =   8705
      _ExtentY        =   6376
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
Private Const WS_EX_NOACTIVATE = &H8000000

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Private Const SWP_NOACTIVATE = &H10

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long



Private mlngTransparentValue As Long
Private mlngState As Long
Private mblnMouseDown As Boolean
Private mhpPos As THintPos
Private mlngParentHwnd As Long
Private mstrImgFile As String
Private mblnBeep As Boolean
Private mblnBgCap As Boolean
Private mstrDes As String
Private mlngLoadState As Long
Private mblnIsPlay As Boolean


Public Sub ShowCaptureHint(ByVal strImgFile As String, ByVal blnBeep As Boolean, ByVal blnBgCap As Boolean, _
    ByVal hpPos As THintPos, ByVal lngParentHwnd As Long, Optional ByVal strDes As String = "")
On Error Resume Next
    mlngTransparentValue = IIf(mlngLoadState <> 1, 20, 25)
    
    mlngState = 0
    mhpPos = hpPos
    mstrImgFile = strImgFile
    mblnBeep = blnBeep
    mblnBgCap = blnBgCap
    mstrDes = strDes
    
    mlngParentHwnd = lngParentHwnd
    mblnIsPlay = False
    
    If mlngLoadState <> 1 Then
        Load Me
    Else
        '如果状态不为1，说明窗体已经被加载，需要手动刷新提示图像
        Timer1.Enabled = False
        
        If UCase(strImgFile) = "AVI" Or UCase(strImgFile) = "WAV" Then
            dcmViewer.Visible = False
            picMedia.Visible = True
        Else
            picMedia.Visible = False
            dcmViewer.Visible = True
            
            Call dcmViewer.Images.Clear
            If InStr(mstrImgFile, Dir(mstrImgFile, 7)) Then Call ReadViewImage(mstrImgFile, dcmViewer)
        End If
        
        Timer1.Enabled = True
    End If
    
    Err.Clear
End Sub

Public Function ReadViewImage(ByVal strFile As String, Optional ByRef dcmViewer As DicomViewer = Nothing) As DicomImage
On Error GoTo errHandle
    Dim dImgs As DicomImages
    
    '如果包含_copy_vdat_，说明是临时文件
    If InStr(strFile, "_copy_vdat_") > 0 Then
        Set ReadViewImage = Nothing
        Call Kill(strFile)
        
        Exit Function
    End If
    
    If dcmViewer Is Nothing Then
        Set dImgs = New DicomImages
    Else
        Set dImgs = dcmViewer.Images
    End If
    
    Set ReadViewImage = ReadDicomFile(strFile, dImgs)
 
Exit Function
errHandle:
    Set ReadViewImage = Nothing
End Function


Private Function ReadDicomFile(ByVal strFile As String, dcmImgs As DicomImages) As DicomImage
On Error Resume Next
    Dim curImage As DicomImage
    Dim blnUseUrl As Boolean
    Dim strFileTime As String
    
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    
    If blnUseUrl Then
        'readurl不支持空格
        Set curImage = dcmImgs.ReadURL(strFile)
    Else
        Set curImage = dcmImgs.ReadFile(strFile)
    End If

    
    If Err.Number = 0 Then
        Set ReadDicomFile = curImage
        Exit Function
    End If

    
    '2098错误一种是文件不是dicom文件，另一种是存在共享访问错误
    If InStr(Err.Description, "sharing violation") > 0 Then
        Err.Clear
        
        strFileTime = Format(Now, "YYMMDD") & GetTickCount
        Call FileCopy(strFile, strFile & "_copy_vdat_" & strFileTime)
    
        If blnUseUrl Then
            'readurl不支持空格
            Set curImage = dcmImgs.ReadURL(strFile & "_copy_vdat_" & strFileTime)
        Else
            Set curImage = dcmImgs.ReadFile(strFile & "_copy_vdat_" & strFileTime)
        End If
        
        If Err.Number = 0 Then
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
            Err.Clear
        Else
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
        End If
    Else
        Err.Clear
        
        Set curImage = dcmImgs.AddNew
        Call curImage.FileImport(strFile, "JPG")
        
        If Err.Number <> 0 Then
            Err.Clear
            'not a JPG file
            Call curImage.FileImport(strFile, "BMP")
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            'not a BMP file
            Call curImage.FileImport(strFile, "AVI")
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            'not a AVI file
            Call curImage.FileImport(strFile, "MPG")
        End If
        
        If Err.Number <> 0 Then
            Call dcmImgs.Remove(dcmImgs.Count)
        End If
    End If
    
    If Err.Number = 0 Then
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    Set ReadDicomFile = Nothing
    
Err.Clear
End Function


Private Sub dcmViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next
    mblnMouseDown = True
    
   
    Err.Clear
End Sub

Private Sub dcmViewer_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error Resume Next
    mblnMouseDown = False
    
    Err.Clear
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
    
    If InStr(mstrImgFile, Dir(mstrImgFile, 7)) > 0 Then
        If mblnBgCap Then Me.Caption = "缓存图像"
        If Len(mstrDes) > 0 Then Me.Caption = Me.Caption & " - " & mstrDes
    
        mblnMouseDown = False
        lngMonitorIndex = GetMonitorIndex(mlngParentHwnd)
        
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
        
        If UCase(mstrImgFile) = "AVI" Or UCase(mstrImgFile) = "WAV" Then
            dcmViewer.Visible = False
            picMedia.Visible = True
        Else
            picMedia.Visible = False
            dcmViewer.Visible = True
            
            Call dcmViewer.Images.Clear
            Call ReadViewImage(mstrImgFile, dcmViewer)
        End If
        
        picMedia.Left = (Me.ScaleWidth - picMedia.Width) / 2
        picMedia.Top = (Me.ScaleHeight - picMedia.Height) / 2
        
        dcmViewer.Left = 0
        dcmViewer.Top = 0
        dcmViewer.Height = Me.ScaleHeight
        dcmViewer.Width = Me.ScaleWidth
        
    
        sty = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
        sty = sty Or WS_EX_LAYERED 'WS_EX_NOACTIVATE Or
        
        SetWindowLong Me.hwnd, GWL_EXSTYLE, sty
        
        SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, &H1 Or &H2 Or SWP_NOACTIVATE Or &H40 '将窗口置顶
    End If
    
    Timer1.Enabled = True
    
    Err.Clear
End Sub


Private Sub Form_Resize()
On Error Resume Next
    picMedia.Left = (Me.ScaleWidth - picMedia.Width) / 2
    picMedia.Top = (Me.ScaleHeight - picMedia.Height) / 2
        
    dcmViewer.Left = 0
    dcmViewer.Top = 0
    dcmViewer.Height = Me.ScaleHeight
    dcmViewer.Width = Me.ScaleWidth
    
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    mlngLoadState = 2
    
    Err.Clear
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If mblnIsPlay = False Then
        If mblnBeep Then
            If mblnBgCap Then
                Call MessageBeep(-1)
                Call Beep(2000, 250)
                Call Beep(2500, 250)
            Else
                Call MessageBeep(-1)
                Call Beep(2500, 500)
            End If
            
        End If
        mblnIsPlay = True
    End If
    
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

    If mlngState <> 1 Then SetLayeredWindowAttributes Me.hwnd, 0, mlngTransparentValue, LWA_ALPHA

    Err.Clear
End Sub


