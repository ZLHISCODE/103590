VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' cDIBFilter:
' Filter processing
'
' All filters by Manuel Augusto Santos, except:
'   - Colorize, from VB Accelerator.
'   - Coded from Despeckle filter article (http://www.dai.ed.ac.uk/HIPR2/crimmins.htm#2)
'   - Equalize, Shift and ReplaceXXX filters, from Le Premier Soleil [http://hp.vector.co.jp/authors/VA019249/index.html]
'   - SaltAndPepperRemoval, by Robert Rayment.
'   - ShiftRGB filter, by gonchucki.
'   - Outline filter (based on Soften + Contour filters).
'   - Texturize filter.
'
'   also:
'   - Added ConnectMargins option for Emboss/Engrave filters (-> Texturize tiling)

Option Explicit

'-- API:

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

'-- Private Variables:
Private t As Long
Private X As Long, xIn As Long
Private Y As Long, yIn As Long
Private W As Long
Private H As Long

'-- Public events:
Public Event ProgressStart()
Public Event Progress(ByVal p As Long)
Public Event ProgressEnd()

'========================================================================================
' Methods
'========================================================================================

Public Sub Blur(DIB As cDIB, Optional ByVal Level As Long = 100)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim Lev     As Long
  Dim Wgt     As Long
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadBlt(DIB.hdc)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits()), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits()), VarPtr(dtSA), 4)
            
        W = DIB.Width - 2
        H = DIB.Height - 2
        
        Lev = 101 - Level
        Wgt = Lev + 8

        For Y = 1 To H
            For X = 1 To W
                With dBits(X, Y)
                    .B = (Lev * CLng(sBits(X, Y).B) + sBits(X + 1, Y).B + sBits(X - 1, Y).B + sBits(X, Y + 1).B + sBits(X, Y - 1).B + sBits(X + 1, Y + 1).B + sBits(X + 1, Y - 1).B + sBits(X - 1, Y + 1).B + sBits(X - 1, Y - 1).B) \ Wgt
                    .G = (Lev * CLng(sBits(X, Y).G) + sBits(X + 1, Y).G + sBits(X - 1, Y).G + sBits(X, Y + 1).G + sBits(X, Y - 1).G + sBits(X + 1, Y + 1).G + sBits(X + 1, Y - 1).G + sBits(X - 1, Y + 1).G + sBits(X - 1, Y - 1).G) \ Wgt
                    .R = (Lev * CLng(sBits(X, Y).R) + sBits(X + 1, Y).R + sBits(X - 1, Y).R + sBits(X, Y + 1).R + sBits(X, Y - 1).R + sBits(X + 1, Y + 1).R + sBits(X + 1, Y - 1).R + sBits(X - 1, Y + 1).R + sBits(X - 1, Y - 1).R) \ Wgt
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Contour(DIB As cDIB)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim v As Long, vMax As Long
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadBlt(DIB.hdc)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits()), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits()), VarPtr(dtSA), 4)
         
        W = DIB.Width - 2
        H = DIB.Height - 2
        
        For Y = 1 To H
            For X = 1 To W
        
                vMax = 0
                v = sBits(X - 1, Y - 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y - 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y - 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y + 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y + 1).B
                If (v > vMax) Then vMax = v
                
                dBits(X, Y).B = 255 Xor (vMax - sBits(X, Y).B)
      
                vMax = 0
                v = sBits(X - 1, Y - 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y - 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y - 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y + 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y + 1).G
                If (v > vMax) Then vMax = v
                
                dBits(X, Y).G = 255 Xor (vMax - sBits(X, Y).G)
                
                vMax = 0
                v = sBits(X - 1, Y - 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y - 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y - 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y + 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y + 1).R
                If (v > vMax) Then vMax = v
            
                dBits(X, Y).R = 255 Xor (vMax - sBits(X, Y).R)
                
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Contrast(DIB As cDIB, Optional ByVal Level As Long = 0)

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim yRel As Single
  
  Dim Spd(765) As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim mCol As Long
  
  Dim Lev As Single
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
         
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        For X = 0 To 765
            Spd(X) = X \ 3
        Next X
        
        mCol = 0
        For Y = 0 To H
            For X = 0 To W
                B = Bits(X, Y).B
                G = Bits(X, Y).G
                R = Bits(X, Y).R
                mCol = mCol + Spd(R + G + B)
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
    
        mCol = mCol \ (DIB.Width * DIB.Height)
        Lev = (Level + 100) / 100
        For X = 0 To 255
            Spd(X) = (X - mCol) * Lev + mCol
        Next X
    
        For Y = 0 To H
            For X = 0 To W
            
                B = Spd(Bits(X, Y).B)
                G = Spd(Bits(X, Y).G)
                R = Spd(Bits(X, Y).R)
                
                Do While (B < 0 Or B > 255 Or G < 0 Or G > 255 Or R < 0 Or R > 255)
                    If (B <= 0 And G <= 0 And R <= 0) Then
                        B = 0
                        G = 0
                        R = 0
                    End If
                    If (B >= 255 And G >= 255 And R >= 255) Then
                        B = 255
                        G = 255
                        R = 255
                    End If
                    If (B < 0) Then
                        G = G + B \ 2
                        R = R + B \ 2
                        B = 0
                    End If
                    If (B > 255) Then
                        G = G + (B - 255) \ 2
                        R = R + (B - 255) \ 2
                        B = 255
                    End If
                    If (G < 0) Then
                        B = B + G \ 2
                        R = R + G \ 2
                        G = 0
                    End If
                    If (G > 255) Then
                        B = B + (G - 255) \ 2
                        R = R + (G - 255) \ 2
                        G = 255
                    End If
                    If (R < 0) Then
                        G = G + R \ 2
                        B = B + R \ 2
                        R = 0
                    End If
                    If (R > 255) Then
                        G = G + (R - 255) \ 2
                        B = B + (R - 255) \ 2
                        R = 255
                    End If
                Loop
                With Bits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Colorize(DIB As cDIB, ByVal Hue As Single, Optional ByVal Saturation As Single = 0.5, Optional ByVal Luminosity As Single = 1)
' Needs mHSL module

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim R As Byte, G As Byte, B As Byte
  Dim L As Single

    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
    
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
    
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        For Y = 0 To H
            For X = 0 To W
                L = (CLng(Bits(X, Y).B) + Bits(X, Y).G + Bits(X, Y).R) / 765
                HSLtoRGB Hue, Saturation, L * Luminosity, R, G, B
                With Bits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Despeckle(DIB As cDIB)
'-- Crimmins Speckle Removal
'   (Simplified)

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim yRel As Single
  
  Dim pX As Byte             ' Current pixel
  Dim pA As Byte, pB As Byte ' Neighbors [N-S, W-S, NW-SE, SW-NE]
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits), VarPtr(tSA), 4)
        
        W = DIB.Width - 2
        H = DIB.Height - 2
        
        '== N-S direction ==
        
        For Y = 1 To H
            For X = 1 To W
                
                '-- Blue channel
                    pX = Bits(X, Y).B
                    pA = Bits(X, Y - 1).B
                    pB = Bits(X, Y + 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                
                '-- Green channel
                    pX = Bits(X, Y).G
                    pA = Bits(X, Y - 1).G
                    pB = Bits(X, Y + 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX
                    
                '-- Red channel
                    pX = Bits(X, Y).R
                    pA = Bits(X, Y - 1).R
                    pB = Bits(X, Y + 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX
                    
            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y
        
        '== W-E direction ==
        
        For Y = 1 To H
            For X = 1 To W
                
                '-- Blue channel
                    pX = Bits(X, Y).B
                    pA = Bits(X - 1, Y).B
                    pB = Bits(X + 1, Y).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                    
                '-- Green channel
                    pX = Bits(X, Y).G
                    pA = Bits(X - 1, Y).G
                    pB = Bits(X + 1, Y).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX

                '-- Red channel
                    pX = Bits(X, Y).R
                    pA = Bits(X - 1, Y).R
                    pB = Bits(X + 1, Y).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX
                    
            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y
        
        '== NW-SE direction ==
        
        For Y = 1 To H
            For X = 1 To W
                    
                '-- Blue channel:
                    pX = Bits(X, Y).B
                    pA = Bits(X - 1, Y - 1).B
                    pB = Bits(X + 1, Y + 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                    
                '-- Green channel:
                    pX = Bits(X, Y).G
                    pA = Bits(X - 1, Y - 1).G
                    pB = Bits(X + 1, Y + 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX
                    
                '-- Red channel:
                    pX = Bits(X, Y).R
                    pA = Bits(X - 1, Y - 1).R
                    pB = Bits(X + 1, Y + 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX
                    
            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y
        
        '== SW-NE direction ==
        
        For Y = 1 To H
            For X = 1 To W
                    
                '-- Blue channel:
                    pX = Bits(X, Y).B
                    pA = Bits(X - 1, Y + 1).B
                    pB = Bits(X + 1, Y - 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                    
                '-- Green channel:
                    pX = Bits(X, Y).G
                    pA = Bits(X - 1, Y + 1).G
                    pB = Bits(X + 1, Y - 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX

                '-- Red channel:
                    pX = Bits(X, Y).R
                    pA = Bits(X - 1, Y + 1).R
                    pB = Bits(X + 1, Y - 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX

            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub DespeckleMore(DIB As cDIB)
'-- Crimmins Speckle Removal

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim yRel As Single
  
  Dim pX As Byte             ' Current pixel
  Dim pA As Byte, pB As Byte ' Neighbors [N-S, W-S, NW-SE, SW-NE]
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits), VarPtr(tSA), 4)
        
        W = DIB.Width - 2
        H = DIB.Height - 2
        
        '== N-S direction ==
        
        For Y = 1 To H
            For X = 1 To W
                
                '-- Blue channel
                    pX = Bits(X, Y).B
                    pA = Bits(X, Y - 1).B
                    pB = Bits(X, Y + 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                
                '-- Green channel
                    pX = Bits(X, Y).G
                    pA = Bits(X, Y - 1).G
                    pB = Bits(X, Y + 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX
                    
                '-- Red channel
                    pX = Bits(X, Y).R
                    pA = Bits(X, Y - 1).R
                    pB = Bits(X, Y + 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX
                    
            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y
        
        '== W-E direction ==
        
        For Y = 1 To H
            For X = 1 To W
                
                '-- Blue channel
                    pX = Bits(X, Y).B
                    pA = Bits(X - 1, Y).B
                    pB = Bits(X + 1, Y).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                    
                '-- Green channel
                    pX = Bits(X, Y).G
                    pA = Bits(X - 1, Y).G
                    pB = Bits(X + 1, Y).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX

                '-- Red channel
                    pX = Bits(X, Y).R
                    pA = Bits(X - 1, Y).R
                    pB = Bits(X + 1, Y).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX
                    
            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y
        
        '== NW-SE direction ==
        
        For Y = 1 To H
            For X = 1 To W
                    
                '-- Blue channel:
                    pX = Bits(X, Y).B
                    pA = Bits(X - 1, Y - 1).B
                    pB = Bits(X + 1, Y + 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                    
                '-- Green channel:
                    pX = Bits(X, Y).G
                    pA = Bits(X - 1, Y - 1).G
                    pB = Bits(X + 1, Y + 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX
                    
                '-- Red channel:
                    pX = Bits(X, Y).R
                    pA = Bits(X - 1, Y - 1).R
                    pB = Bits(X + 1, Y + 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX
                    
            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y
        
        '== SW-NE direction ==
        
        For Y = 1 To H
            For X = 1 To W
                    
                '-- Blue channel:
                    pX = Bits(X, Y).B
                    pA = Bits(X - 1, Y + 1).B
                    pB = Bits(X + 1, Y - 1).B
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).B = pX
                    
                '-- Green channel:
                    pX = Bits(X, Y).G
                    pA = Bits(X - 1, Y + 1).G
                    pB = Bits(X + 1, Y - 1).G
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).G = pX

                '-- Red channel:
                    pX = Bits(X, Y).R
                    pA = Bits(X - 1, Y + 1).R
                    pB = Bits(X + 1, Y - 1).R
                '   Light pixel:
                    If (pA < pX - 1) Then If (pX > 0) Then pX = pX - 1
                    If (pA < pX And pX >= pB) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX And pX >= pA) Then If (pX > 0) Then pX = pX - 1
                    If (pB < pX - 1) Then If (pX > 0) Then pX = pX - 1
                '   Dark pixel:
                    If (pA > pX + 1) Then If (pX < 255) Then pX = pX + 1
                    If (pA > pX And pX <= pB) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX And pX <= pA) Then If (pX < 255) Then pX = pX + 1
                    If (pB > pX + 1) Then If (pX < 255) Then pX = pX + 1
                '   Output:
                    Bits(X, Y).R = pX

            Next X
            yRel = yRel + 0.25
            RaiseEvent Progress(yRel)
        Next Y

        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Diffuse(DIB As cDIB, Optional ByVal xRange As Long = 1, Optional ByVal yRange As Long = 1, Optional ByVal DiffuseStep As Long = 1)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim xRange2 As Long
  Dim yRange2 As Long
  
    If (DIB.hDIB <> 0) Then
        
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
            
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        xRange2 = 2 * xRange
        yRange2 = 2 * yRange
        
        If (DiffuseStep < 1) Then DiffuseStep = 1
        
        For Y = 0 To H Step DiffuseStep
            For X = 0 To W Step DiffuseStep
                xIn = X + Rnd * xRange2 - xRange
                yIn = Y + Rnd * yRange2 - yRange
                If (xIn < 1) Then xIn = 0 Else If (xIn > W) Then xIn = W
                If (yIn < 1) Then yIn = 0 Else If (yIn > H) Then yIn = H
                With dBits(X, Y)
                    .B = sBits(xIn, yIn).B
                    .G = sBits(xIn, yIn).G
                    .R = sBits(xIn, yIn).R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Emboss(DIB As cDIB, Optional ByVal Level As Long = 100, Optional ByVal ConnectMargins As Boolean = -1)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim xOff As Long
  Dim yOff As Long
  Dim yRel As Single
  
  Dim Lev As Single
  Dim L   As Byte
  
    If (DIB.hDIB <> 0) Then
        
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
         
        W = DIB.Width - IIf(ConnectMargins, 1, 2)
        H = DIB.Height - IIf(ConnectMargins, 1, 2)
        
        Lev = Level / 100
        
        For Y = 0 To H
            If (Y = H) Then yIn = 0 Else yIn = Y + 1
            For X = 0 To W
                If (X = 0) Then xIn = 0 Else xIn = X - 1
                B = Abs((CLng(sBits(X, Y).B) - sBits(xIn, yIn).B) * Lev + 128)
                G = Abs((CLng(sBits(X, Y).G) - sBits(xIn, yIn).G) * Lev + 128)
                R = Abs((CLng(sBits(X, Y).R) - sBits(xIn, yIn).R) * Lev + 128)
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                With dBits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
        
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        For Y = 0 To H
            For X = 0 To W
                With dBits(X, Y)
                     L = 0.222 * .R + 0.707 * .G + 0.071 * .B
                    'L = 0.299 * .R + 0.587 * .G + 0.114 * .B
                    .B = L
                    .G = L
                    .R = L
                End With
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Engrave(DIB As cDIB, Optional ByVal Level As Long = 100, Optional ByVal ConnectMargins As Boolean = 0)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim xOff As Long
  Dim yOff As Long
  Dim yRel As Single
  
  Dim Lev As Single
  Dim L   As Byte
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
         
        W = DIB.Width - IIf(ConnectMargins, 1, 2)
        H = DIB.Height - IIf(ConnectMargins, 1, 2)
        
        Lev = Level / 100
        
        For Y = 0 To H
            If (Y = H) Then yIn = 0 Else yIn = Y + 1
            For X = 0 To W
                If (X = 0) Then xIn = 0 Else xIn = X - 1
                B = Abs((CLng(-sBits(X, Y).B) + sBits(xIn, yIn).B) * Lev + 128)
                G = Abs((CLng(-sBits(X, Y).G) + sBits(xIn, yIn).G) * Lev + 128)
                R = Abs((CLng(-sBits(X, Y).R) + sBits(xIn, yIn).R) * Lev + 128)
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                With dBits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
        
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        For Y = 0 To H
            For X = 0 To W
                With dBits(X, Y)
                     L = 0.222 * .R + 0.707 * .G + 0.071 * .B
                    'L = 0.299 * .R + 0.587 * .G + 0.114 * .B
                    .B = L
                    .G = L
                    .R = L
                End With
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

'Public Sub Equalize(DIB As cDIB)
''-- L equalization
'
'  Dim Bits() As RGBQUAD
'  Dim tSA    As SAFEARRAY2D
'
'  Dim Hst(255) As Long, L As Long
'  Dim Avg As Long
'  Dim Sum As Long
'  Dim Cum As Long
'  Dim Uf As Single, Lf As Single
'
'    If (DIB.hDIB <> 0) Then
'
'        Call pvBuildSA(tSA, DIB)
'        Call CopyMemory (ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
'
'        W = DIB.Width - 1
'        H = DIB.Height - 1
'
'        '-- Get Histogram
'        For y = 0 To H
'            For x = 0 To W
'                L = 0.114 * Bits(x, y).B + 0.587 * Bits(x, y).G + 0.299 * Bits(x, y).R
'                Hst(L) = Hst(L) + 1
'            Next x
'        Next y
'
'        '-- Find Avg., range factors
'        Sum = 0
'        Cum = 0.5 * (DIB.Width * DIB.Height)
'        For x = 0 To 255
'            Sum = Sum + Hst(x)
'            If (Sum > Cum) Then Exit For
'        Next x
'        If (Hst(x) \ 2 > Sum - Cum) Then
'            Avg = x
'          Else
'            Avg = x - 1
'        End If
'        Uf = 127 / (255 - Avg)
'        Lf = 127 / (Avg - 1)
'
'        '-- Equalize
'        For y = 0 To H
'            For x = 0 To W
'                With Bits(x, y)
'                    If (.B < Avg) Then .B = .B * Lf Else .B = (.B - Avg) * Uf + 128
'                    If (.G < Avg) Then .G = .G * Lf Else .G = (.G - Avg) * Uf + 128
'                    If (.R < Avg) Then .R = .R * Lf Else .R = (.R - Avg) * Uf + 128
'                End With
'            Next x
'            RaiseEvent Progress(y)
'        Next y
'        Call CopyMemory ByVal(VarPtrArray(Bits), 0&, 4)
'        RaiseEvent ProgressEnd
'    End If
'End Sub

'Public Sub GammaCorrection(DIB As cDIB, ByVal Level As Long)
'
'  Dim Bits() As RGBQUAD
'  Dim tSA    As SAFEARRAY2D
'
'  Dim R As Long
'  Dim G As Long
'  Dim B As Long
'
'  Dim dB As Double
'  Dim dG As Double
'  Dim dR As Double
'
'  Dim Lev As Single
'
'    If (DIB.hDIB <> 0) Then
'
'        Call pvBuildSA(tSA, DIB)
'        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
'
'        W = DIB.Width - 1
'        H = DIB.Height - 1
'
'        Lev = (Level + 100) / 100
'
'        For y = 0 To H
'            For x = 0 To W
'                With Bits(x, y)
'                    B = .B
'                    G = .G
'                    R = .R
'                End With
'                If (Lev > 0) Then
'                    dB = B / 255
'                    dG = G / 255
'                    dR = R / 255
'                    dB = dB ^ (1 / Lev)
'                    dG = dG ^ (1 / Lev)
'                    dR = dR ^ (1 / Lev)
'                    B = dB * 255
'                    G = dG * 255
'                    R = dR * 255
'                  Else
'                    B = 0
'                    G = 0
'                    R = 0
'                End If
'                With Bits(x, y)
'                    .B = B
'                    .G = G
'                    .R = R
'                End With
'            Next x
'            RaiseEvent Progress(y)
'        Next y
'        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
'        RaiseEvent ProgressEnd
'    End If
'End Sub

Public Sub Greys(DIB As cDIB)
    
  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim L As Byte
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        pvBuildSA tSA, DIB
        CopyMemory ByVal VarPtrArray(Bits()), VarPtr(tSA), 4
        
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        For Y = 0 To H
            For X = 0 To W
                With Bits(X, Y)
                    L = 0.114 * .B + 0.587 * .G + 0.299 * .R
                    .B = L
                    .G = L
                    .R = L
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Negative(DIB As cDIB)

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        W = DIB.Width - 1
        H = DIB.Height - 1

        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)

        For Y = 0 To H
            For X = 0 To W
                With Bits(X, Y)
                    .B = 255 Xor .B
                    .G = 255 Xor .G
                    .R = 255 Xor .R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Noise(DIB As cDIB, Optional ByVal Level As Long = 5, Optional ByVal xyStep As Long = 1)
    
  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim Lev As Long, DblLev As Long
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
    
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
        
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        DblLev = 2 * Level
        
        For Y = 0 To H Step xyStep
            For X = 0 To W Step xyStep
                Lev = Rnd * DblLev - Level
                B = CLng(Bits(X, Y).B) + Lev
                G = CLng(Bits(X, Y).G) + Lev
                R = CLng(Bits(X, Y).R) + Lev
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                With Bits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Outline(DIB As cDIB, Optional ByVal Level As Long = -50, Optional ByVal Shift As Long = 5)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim v As Long, vMax As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim Lev  As Single
  
  Dim yRel As Single
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
         
        W = DIB.Width - 2
        H = DIB.Height - 2
        
        Lev = -Level / 50
        
        For Y = 1 To H
            For X = 1 To W
                With sBits(X, Y)
                    .B = (CLng(.B) + sBits(X + 1, Y).B + sBits(X - 1, Y).B + sBits(X, Y + 1).B + sBits(X, Y - 1).B) \ 5
                    .G = (CLng(.G) + sBits(X + 1, Y).G + sBits(X - 1, Y).G + sBits(X, Y + 1).G + sBits(X, Y - 1).G) \ 5
                    .R = (CLng(.R) + sBits(X + 1, Y).R + sBits(X - 1, Y).R + sBits(X, Y + 1).R + sBits(X, Y - 1).R) \ 5
                End With
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
        
        For Y = 1 To H
            For X = 1 To W
        
                vMax = 0
                v = sBits(X - 1, Y - 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y - 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y - 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y + 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y + 1).B
                If (v > vMax) Then vMax = v
                
                v = dBits(X, Y).B - Lev * (vMax - sBits(X, Y).B) + Shift
                If (v < 0) Then v = 0 Else If (v > 255) Then v = 255
                dBits(X, Y).B = v
                
                vMax = 0
                v = sBits(X - 1, Y - 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y - 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y - 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y + 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y + 1).G
                If (v > vMax) Then vMax = v
                
                v = dBits(X, Y).G - Lev * (vMax - sBits(X, Y).G) + Shift
                If (v < 0) Then v = 0 Else If (v > 255) Then v = 255
                dBits(X, Y).G = v
                
                vMax = 0
                v = sBits(X - 1, Y - 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y - 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y - 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y + 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y + 1).R
                If (v > vMax) Then vMax = v
            
                v = dBits(X, Y).R - Lev * (vMax - sBits(X, Y).R) + Shift
                If (v < 0) Then v = 0 Else If (v > 255) Then v = 255
                dBits(X, Y).R = v
                
            Next X
            yRel = yRel + 0.5
            RaiseEvent Progress(yRel)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Pixelize(DIB As cDIB, Optional ByVal PixelSize As Long = 2)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim PixelSize1 As Long
  Dim pX As Long, pY As Long
  Dim sx As Long, sy As Long
  Dim mC As Long
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
    
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
        
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        PixelSize1 = PixelSize - 1
        
        For Y = 0 To H
            sy = (Y \ PixelSize) * PixelSize
            For X = 0 To W
                sx = (X \ PixelSize) * PixelSize
                If (X Mod PixelSize) = 0 Then
                    B = 0
                    G = 0
                    R = 0
                    mC = 0
                    For pX = sx To sx + PixelSize1
                        For pY = sy To sy + PixelSize1
                            If (pX <= W And pY <= H) Then
                                B = B + sBits(pX, pY).B
                                G = G + sBits(pX, pY).G
                                R = R + sBits(pX, pY).R
                                mC = mC + 1
                            End If
                        Next pY
                    Next pX
                    If (mC > 0) Then
                        B = B \ mC
                        G = G \ mC
                        R = R \ mC
                    End If
                End If
                With dBits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub RankFilterMaximum(DIB As cDIB)
'-- Dilate (4N pixels)

  Dim sDIB    As New cDIB

  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D

  Dim v As Long, vMax As Long

    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
        
        W = DIB.Width - 2
        H = DIB.Height - 2

        For Y = 1 To H
            For X = 1 To W

                vMax = 0
                v = sBits(X, Y - 1).B
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).B
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).B
                If (v > vMax) Then vMax = v

                dBits(X, Y).B = vMax

                vMax = 0
                v = sBits(X, Y - 1).G
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).G
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).G
                If (v > vMax) Then vMax = v

                dBits(X, Y).G = vMax

                vMax = 0
                v = sBits(X, Y - 1).R
                If (v > vMax) Then vMax = v
                v = sBits(X - 1, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X + 1, Y).R
                If (v > vMax) Then vMax = v
                v = sBits(X, Y + 1).R
                If (v > vMax) Then vMax = v
                
                dBits(X, Y).R = vMax

            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub RankFilterMinimum(DIB As cDIB)
'-- Erode (4N pixels)

  Dim sDIB    As New cDIB

  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D

  Dim v As Long, vMin As Long

    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
        
        W = DIB.Width - 2
        H = DIB.Height - 2

        For Y = 1 To H
            For X = 1 To W

                vMin = 255
                v = sBits(X, Y - 1).B
                If (v < vMin) Then vMin = v
                v = sBits(X - 1, Y).B
                If (v < vMin) Then vMin = v
                v = sBits(X, Y).B
                If (v < vMin) Then vMin = v
                v = sBits(X + 1, Y).B
                If (v < vMin) Then vMin = v
                v = sBits(X, Y + 1).B
                If (v < vMin) Then vMin = v

                dBits(X, Y).B = vMin

                vMin = 255
                v = sBits(X, Y - 1).G
                If (v < vMin) Then vMin = v
                v = sBits(X - 1, Y).G
                If (v < vMin) Then vMin = v
                v = sBits(X, Y).G
                If (v < vMin) Then vMin = v
                v = sBits(X + 1, Y).G
                If (v < vMin) Then vMin = v
                v = sBits(X, Y + 1).G
                If (v < vMin) Then vMin = v

                dBits(X, Y).G = vMin

                vMin = 255
                v = sBits(X, Y - 1).R
                If (v < vMin) Then vMin = v
                v = sBits(X - 1, Y).R
                If (v < vMin) Then vMin = v
                v = sBits(X, Y).R
                If (v < vMin) Then vMin = v
                v = sBits(X + 1, Y).R
                If (v < vMin) Then vMin = v
                v = sBits(X, Y + 1).R
                If (v < vMin) Then vMin = v

                dBits(X, Y).R = vMin

            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Relieve(DIB As cDIB, Optional ByVal Level As Long = 75)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim Lev As Single
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
         
        W = DIB.Width - 2
        H = DIB.Height - 2
        
        Lev = (101 - Level) / 10
        
        For Y = 1 To H
            For X = 1 To W
                B = (CLng(Lev * sBits(X, Y).B) + sBits(X + 1, Y - 1).B + sBits(X + 1, Y).B + sBits(X, Y - 1).B - sBits(X - 1, Y + 1).B - sBits(X - 1, Y).B - sBits(X, Y + 1).B) / Lev
                G = (CLng(Lev * sBits(X, Y).G) + sBits(X + 1, Y - 1).G + sBits(X + 1, Y).G + sBits(X, Y - 1).G - sBits(X - 1, Y + 1).G - sBits(X - 1, Y).G - sBits(X, Y + 1).G) / Lev
                R = (CLng(Lev * sBits(X, Y).R) + sBits(X + 1, Y - 1).R + sBits(X + 1, Y).R + sBits(X, Y - 1).R - sBits(X - 1, Y + 1).R - sBits(X - 1, Y).R - sBits(X, Y + 1).R) / Lev
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                With dBits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub ReplaceHS(DIB As cDIB, ByVal ColorFrom As Long, ByVal ColorTo As Long, ByVal Tolerance As Long, ByVal HTolerance As Long, ByVal STolerance As Long)
' Needs mHSL module
  
  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D

  Dim Rf As Long, Gf As Long, Bf As Long
  Dim Rt As Long, Gt As Long, Bt As Long
  
  Dim Hf As Single, sF As Single, Lf As Single
  Dim Ht As Single, St As Single, Lt As Single
  
  Dim lTol As Long, sTolH As Single, sTolS As Single
  Dim D As Long, ratio As Single

    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
            
        Rf = (ColorFrom And &HFF&)
        Gf = (ColorFrom And &HFF00&) \ 256
        Bf = (ColorFrom And &HFF0000) \ 65536
        Rt = (ColorTo And &HFF&)
        Gt = (ColorTo And &HFF00&) \ 256
        Bt = (ColorTo And &HFF0000) \ 65536
        
        RGBtoHSL Rf, Gf, Bf, Hf, sF, Lf
        
        lTol = (127.5 * Tolerance / 100) ^ 2 + 1
        sTolH = HTolerance / 20
        sTolS = STolerance / 100
        
        W = DIB.Width - 1
        H = DIB.Height - 1

        For Y = 0 To H
            For X = 0 To W
                D = (Bits(X, Y).B - Bf) * (Bits(X, Y).B - Bf) + (Bits(X, Y).G - Gf) * (Bits(X, Y).G - Gf) + (Bits(X, Y).R - Rf) * (Bits(X, Y).R - Rf)
                If (D <= lTol) Then
                    RGBtoHSL Bits(X, Y).R, Bits(X, Y).G, Bits(X, Y).B, Ht, St, Lt
                    If (Abs(Ht - Hf) <= sTolH And Abs(St - sF) <= sTolS) Then
                        ratio = 1 - D / lTol
                        With Bits(X, Y)
                            .B = .B + (Bt - .B) * ratio
                            .G = .G + (Gt - .G) * ratio
                            .R = .R + (Rt - .R) * ratio
                        End With
                    End If
                End If
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub ReplaceL(DIB As cDIB, ByVal ColorFrom As Long, ByVal ColorTo As Long, ByVal Tolerance As Long, ByVal LTolerance As Long)
  
  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D

  Dim Rf As Long, Gf As Long, Bf As Long
  Dim Rt As Long, Gt As Long, Bt As Long
  
  Dim lTol As Long, sTolL As Single
  Dim D As Long, ratio As Single

    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
            
        Rf = (ColorFrom And &HFF&)
        Gf = (ColorFrom And &HFF00&) \ 256
        Bf = (ColorFrom And &HFF0000) \ 65536
        Rt = (ColorTo And &HFF&)
        Gt = (ColorTo And &HFF00&) \ 256
        Bt = (ColorTo And &HFF0000) \ 65536
        
        lTol = (127.5 * Tolerance / 100) ^ 2 + 1
        sTolL = LTolerance / 100
        
        W = DIB.Width - 1
        H = DIB.Height - 1

        For Y = 0 To H
            For X = 0 To W
                D = (Bits(X, Y).B - Bf) * (Bits(X, Y).B - Bf) + (Bits(X, Y).G - Gf) * (Bits(X, Y).G - Gf) + (Bits(X, Y).R - Rf) * (Bits(X, Y).R - Rf)
                If (D <= lTol) Then
                    If (Abs((CLng(Bits(X, Y).B) + Bits(X, Y).G + Bits(X, Y).R - Bf - Gf - Rf) / 765) <= sTolL) Then
                        ratio = 1 - D / lTol
                        With Bits(X, Y)
                            .B = .B + (Bt - .B) * ratio
                            .G = .G + (Gt - .G) * ratio
                            .R = .R + (Rt - .R) * ratio
                        End With
                    End If
                End If
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub SaltAndPepperRemoval(DIB As cDIB, Optional ByVal Threshold As Long = 64)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim i As Long, j As Long
  
  Dim Ravg As Long, Rdiff As Long
  Dim Gavg As Long, Gdiff As Long
  Dim Bavg As Long, Bdiff As Long
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
         
        W = DIB.Width - 3
        H = DIB.Height - 3
        
        For Y = 1 To H
            For X = 1 To W
            
                '-- Sum 12 pixels around block of 4
                Bavg = 0
                Gavg = 0
                Ravg = 0
                For j = -1 To 2
                For i = -1 To 2
                    If (i = -1 Or i = 2 Or j = -1 Or j = 2) Then
                        With sBits(X + i, Y + j)
                            Bavg = Bavg + .B
                            Gavg = Gavg + .G
                            Ravg = Ravg + .R
                        End With
                    End If
                Next i
                Next j
                Bavg = Bavg \ 12
                Gavg = Gavg \ 12
                Ravg = Ravg \ 12
                
                For j = 0 To 1
                For i = 0 To 1
                    With sBits(X + i, Y + j)
                        Bdiff = Abs(Bavg - .B)
                        Gdiff = Abs(Gavg - .G)
                        Rdiff = Abs(Ravg - .R)
                    End With
                    If (Bdiff > Threshold Or Gdiff > Threshold Or Rdiff > Threshold) Then
                        With dBits(X + i, Y + j)
                            .B = Bavg
                            .G = Gavg
                            .R = Ravg
                        End With
                        With sBits(X + i, Y + j)
                            .B = Bavg
                            .G = Gavg
                            .R = Ravg
                        End With
                    End If
                Next i
                Next j
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Saturation(DIB As cDIB, Optional ByVal Level As Long = 0)

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim Spd(765) As Long
  Dim R As Long
  Dim G As Long
  Dim B As Long
  Dim v As Long
  
  Dim Lev As Single
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
         
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        For X = 0 To 765
            Spd(X) = X \ 3
        Next X
    
        Lev = Level / 100
        
        For Y = 0 To H
            For X = 0 To W
                B = Bits(X, Y).B
                G = Bits(X, Y).G
                R = Bits(X, Y).R
                v = Spd(B + G + R)
                B = B + Lev * (B - v)
                G = G + Lev * (G - v)
                R = R + Lev * (R - v)
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                With Bits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Scanlines(DIB As cDIB, Optional ByVal BlackLevel As Long = 20, Optional ByVal WhiteLevel As Long = 20)

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim cFInc As Single
  Dim cFDec As Single
  Dim SpdInc(255) As Byte
  Dim SpdDec(255) As Byte
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
        
        cFInc = 1 + WhiteLevel / 100
        cFDec = 1 - BlackLevel / 100
        
        For X = 0 To 255
            If (X * cFInc > 255) Then SpdInc(X) = 255 Else SpdInc(X) = X * cFInc
            If (X * cFDec < 0) Then SpdDec(X) = 0 Else SpdDec(X) = X * cFDec
        Next X
        
        W = DIB.Width - 1
        H = DIB.Height - 2
                
        For Y = 0 To H Step 2
            For X = 0 To W
                With Bits(X, Y)
                    .B = SpdInc(.B)
                    .G = SpdInc(.G)
                    .R = SpdInc(.R)
                End With
                With Bits(X, Y + 1)
                    .B = SpdDec(.B)
                    .G = SpdDec(.G)
                    .R = SpdDec(.R)
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Sharpen(DIB As cDIB, Optional ByVal Level As Long = 85)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
  Dim Lev As Long
  Dim Wgt As Long
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
            
        W = DIB.Width - 2
        H = DIB.Height - 2
    
        Lev = 109 - Level
        Wgt = Lev - 8
        
        For Y = 1 To H
            For X = 1 To W
                B = (Lev * CLng(sBits(X, Y).B) - sBits(X + 1, Y).B - sBits(X - 1, Y).B - sBits(X, Y + 1).B - sBits(X, Y - 1).B - sBits(X + 1, Y + 1).B - sBits(X + 1, Y - 1).B - sBits(X - 1, Y + 1).B - sBits(X - 1, Y - 1).B) \ Wgt
                G = (Lev * CLng(sBits(X, Y).G) - sBits(X + 1, Y).G - sBits(X - 1, Y).G - sBits(X, Y + 1).G - sBits(X, Y - 1).G - sBits(X + 1, Y + 1).G - sBits(X + 1, Y - 1).G - sBits(X - 1, Y + 1).G - sBits(X - 1, Y - 1).G) \ Wgt
                R = (Lev * CLng(sBits(X, Y).R) - sBits(X + 1, Y).R - sBits(X - 1, Y).R - sBits(X, Y + 1).R - sBits(X, Y - 1).R - sBits(X + 1, Y + 1).R - sBits(X + 1, Y - 1).R - sBits(X - 1, Y + 1).R - sBits(X - 1, Y - 1).R) \ Wgt
                If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                With dBits(X, Y)
                    .B = B
                    .G = G
                    .R = R
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Shift(DIB As cDIB, ByVal ColorFrom As Long, ByVal Inc As Long, ByVal Tolerance As Long)

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D

  Dim Spd(255) As Integer
  Dim Rf As Long, Gf As Long, Bf As Long
  Dim lTol As Long
  Dim D As Long, ratio As Single
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)

        Rf = (ColorFrom And &HFF&)
        Gf = (ColorFrom And &HFF00&) \ 256
        Bf = (ColorFrom And &HFF0000) \ 65536
        lTol = (127.5 * Tolerance / 100) ^ 2 + 1
    
        For X = 0 To 255
            Spd(X) = X + Inc
            If (Spd(X) < 0) Then Spd(X) = 0 Else If (Spd(X) > 255) Then Spd(X) = 255
        Next X
        
        W = DIB.Width - 1
        H = DIB.Height - 1
    
        For Y = 0 To H
            For X = 0 To W
                D = (Bits(X, Y).B - Bf) * (Bits(X, Y).B - Bf) + (Bits(X, Y).G - Gf) * (Bits(X, Y).G - Gf) + (Bits(X, Y).R - Rf) * (Bits(X, Y).R - Rf)
                If (D <= lTol) Then
                    ratio = 1 - D / lTol
                    With Bits(X, Y)
                        .B = .B + (Spd(.B) - .B) * ratio
                        .G = .G + (Spd(.G) - .G) * ratio
                        .R = .R + (Spd(.R) - .R) * ratio
                    End With
                End If
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub ShiftRGB(DIB As cDIB, ByVal RInc As Long, ByVal GInc As Long, ByVal BInc As Long)

  Dim Bits() As RGBQUAD
  Dim tSA    As SAFEARRAY2D
  
  Dim SpdR(255) As Long
  Dim SpdG(255) As Long
  Dim SpdB(255) As Long

    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        For X = 0 To 255
            SpdB(X) = X + BInc
            SpdG(X) = X + GInc
            SpdR(X) = X + RInc
            If (SpdB(X) < 0) Then SpdB(X) = 0 Else If (SpdB(X) > 255) Then SpdB(X) = 255
            If (SpdG(X) < 0) Then SpdG(X) = 0 Else If (SpdG(X) > 255) Then SpdG(X) = 255
            If (SpdR(X) < 0) Then SpdR(X) = 0 Else If (SpdR(X) > 255) Then SpdR(X) = 255
        Next X
    
        Call pvBuildSA(tSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(Bits()), VarPtr(tSA), 4)
         
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        For Y = 0 To H
            For X = 0 To W
                With Bits(X, Y)
                    .B = SpdB(.B)
                    .G = SpdG(.G)
                    .R = SpdR(.R)
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(Bits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Soften(DIB As cDIB, Optional ByVal Level As Long = 95)

  Dim sDIB    As New cDIB
  
  Dim sBits() As RGBQUAD
  Dim dBits() As RGBQUAD
  Dim stSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim Lev As Long
  Dim Wgt As Long
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call sDIB.Create(DIB.Width, DIB.Height)
        Call sDIB.LoadDIBBlt(DIB)
        
        Call pvBuildSA(stSA, sDIB)
        Call CopyMemory(ByVal VarPtrArray(sBits), VarPtr(stSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits), VarPtr(dtSA), 4)
            
        W = DIB.Width - 2
        H = DIB.Height - 2
    
        Lev = 101 - Level
        Wgt = Lev + 4
        
        For Y = 1 To H
            For X = 1 To W
                With dBits(X, Y)
                    .B = (Lev * CLng(sBits(X, Y).B) + sBits(X + 1, Y).B + sBits(X - 1, Y).B + sBits(X, Y + 1).B + sBits(X, Y - 1).B) \ Wgt
                    .G = (Lev * CLng(sBits(X, Y).G) + sBits(X + 1, Y).G + sBits(X - 1, Y).G + sBits(X, Y + 1).G + sBits(X, Y - 1).G) \ Wgt
                    .R = (Lev * CLng(sBits(X, Y).R) + sBits(X + 1, Y).R + sBits(X - 1, Y).R + sBits(X, Y + 1).R + sBits(X, Y - 1).R) \ Wgt
                End With
            Next X
            RaiseEvent Progress(Y)
        Next Y
        Call CopyMemory(ByVal VarPtrArray(sBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

Public Sub Texturize(DIB As cDIB, TextureDIB As cDIB)

  Dim tBits() As RGBQUAD 'Texture
  Dim dBits() As RGBQUAD
  Dim ttSA    As SAFEARRAY2D
  Dim dtSA    As SAFEARRAY2D
  
  Dim tW As Long, tx As Long, txStp As Long
  Dim tH As Long, ty As Long, tyStp As Long
  Dim tF As Single, tS As Long
  Dim lTiles As Long
  Dim lcTile As Long
  
  Dim R As Long
  Dim G As Long
  Dim B As Long
  
    If (DIB.hDIB <> 0) Then
        RaiseEvent ProgressStart
        
        Call pvBuildSA(ttSA, TextureDIB)
        Call CopyMemory(ByVal VarPtrArray(tBits()), VarPtr(ttSA), 4)
        Call pvBuildSA(dtSA, DIB)
        Call CopyMemory(ByVal VarPtrArray(dBits()), VarPtr(dtSA), 4)
            
        W = DIB.Width - 1
        H = DIB.Height - 1
        
        tW = TextureDIB.Width - 1
        tH = TextureDIB.Height - 1
        txStp = tW + 1
        tyStp = tH + 1
        
        lTiles = -Int(-(W + 1) \ (tW + 1)) * -Int(-(H + 1) \ (tH + 1)) + 1
        
        For Y = 0 To H Step tyStp
            For X = 0 To W Step txStp
                For ty = 0 To tH
                    yIn = ty + Y
                    For tx = 0 To tW
                        xIn = tx + X
                        If (xIn <= W And yIn <= H) Then
                            '-- R = B = G (based on embossed/engraved 'map')
                            '   128 = med. L
                            tF = tBits(tx, ty).B / 127.5
                            tS = tBits(tx, ty).B - 128
                            B = tF * dBits(xIn, yIn).B + tS
                            G = tF * dBits(xIn, yIn).G + tS
                            R = tF * dBits(xIn, yIn).R + tS
                            If (B < 0) Then B = 0 Else If (B > 255) Then B = 255
                            If (G < 0) Then G = 0 Else If (G > 255) Then G = 255
                            If (R < 0) Then R = 0 Else If (R > 255) Then R = 255
                            With dBits(xIn, yIn)
                                .B = B
                                .G = G
                                .R = R
                            End With
                        End If
                    Next tx
                Next ty
                lcTile = lcTile + 1
                RaiseEvent Progress((lcTile / lTiles) * H)
            Next X
        Next Y
        Call CopyMemory(ByVal VarPtrArray(tBits), 0&, 4)
        Call CopyMemory(ByVal VarPtrArray(dBits), 0&, 4)
        RaiseEvent ProgressEnd
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvBuildSA(tSA As SAFEARRAY2D, DIB As cDIB)
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.Width
        .pvData = DIB.lpBits
    End With
End Sub
