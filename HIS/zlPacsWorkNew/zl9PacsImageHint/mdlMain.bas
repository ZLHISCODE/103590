Attribute VB_Name = "mdlMain"
Option Explicit


Public Sub Main()
    Dim strDcmFile As String
    Dim lngParentHwnd As Long
    Dim blnIsSound As Boolean
    Dim blnIsBgCap As Boolean
    Dim strDes As String
    
    Dim aryPars() As String
    
    Dim objHint As New frmCaptureHint
    
    
    aryPars = Split(Command & ",0,1,0,,", ",")   '0-文件名,1-句柄,2-是否声音提示,3-是否后台采集,4-说明
    
    strDcmFile = aryPars(0)
    lngParentHwnd = Val(aryPars(1))
    
    blnIsSound = IIf(Val(aryPars(2)) = 0, False, True)
    blnIsBgCap = IIf(Val(aryPars(3)) = 0, False, True)
    strDes = aryPars(4)
    
    If UCase(strDcmFile) <> "AVI" And UCase(strDcmFile) <> "WAV" Then
        If Dir(strDcmFile, 7) = "" Then Exit Sub
    End If
     
  
    objHint.ShowCaptureHint strDcmFile, blnIsSound, blnIsBgCap, hpRB, lngParentHwnd, strDes
    
End Sub
