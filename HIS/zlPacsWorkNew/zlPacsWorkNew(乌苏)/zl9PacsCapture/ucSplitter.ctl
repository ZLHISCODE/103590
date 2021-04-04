VERSION 5.00
Begin VB.UserControl ucSplitter 
   Alignable       =   -1  'True
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   180
   FillColor       =   &H00404040&
   MousePointer    =   9  'Size W E
   ScaleHeight     =   4110
   ScaleWidth      =   180
   ToolboxBitmap   =   "ucSplitter.ctx":0000
End
Attribute VB_Name = "ucSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'分割条类型
Public Enum enmSplitType
    stHorizontal = 0
    stVertical = 1
End Enum

'保存鼠标位置
Private Type POINT
    X As Long
    Y As Long
End Type


'鼠标双击类型
Public Enum enmDBClickType
    dtNone = 0
    dtHideControl1 = 1
    dtHideControl2 = 2
End Enum

'边框样式
Public Enum enmBorderStyle
    bsNont = 0
    bsFixedSingle = 1
End Enum

'控件样式
Public Enum enmAppearance
    atFlat = 0
    at3D = 1
End Enum


'指针样式
Public Enum enmCursorType
    ctDefault = 0
    ctArrow = 1
    ctCross = 2
    ctIBeam = 3
    ctIcon = 4
    ctSize = 5
    ctSizeNESW = 6
    ctSizeNS = 7
    ctSizeNWSE = 8
    ctSizeWE = 9
    ctUpArrow = 10
    ctHourGlass = 11
    ctNoDrop = 12
    ctArrowAndHourGlass = 13
    ctArrowAndQuestion = 14
    ctSizeAll = 15
    ctCustom = 99
End Enum

'分割控件的层次级别
Public Enum enmSplitLevel
    slNone = 0
    slFirst = 1
    slEnd = 2
    slFirstAndEnd = 3
End Enum


''画布区域
'Private Type TRECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'
''横向对齐方式
'Public Enum enmHorizontalAlignment
'    alLeft = 0
'    alCenter = 1
'    alRight = 2
'End Enum
'
''纵向对齐方式
'Public Enum enmVerticalAlignment
'    alTop = 0
'    alMidlle = 1
'    alBottom = 2
'End Enum


Private mobjControl1 As Object, mobjControl2 As Object
Private mstSplitType As enmSplitType
Private mlngSplitWidth As Long
Private mdtDBClickType As enmDBClickType
Private mslSplitLevel As enmSplitLevel
Private mblnAllowMove As Boolean
Private mblnSyncParentHeight As Boolean
Private mblnSyncParentWidth As Boolean
Private mblnAllowPaintOtherSpliter As Boolean
Private mlngCon1MinSize As Long
Private mlngCon2MinSize As Long
Private mlngStartDistance As Long
Private mlngOldPostion As Long



Private mblnMouseDown As Boolean
Private mblnLayOutState As Boolean
Private mblnDbClickState As Boolean
Private mstrControl1Name As String
Private mstrControl2Name As String

Private mControl2RightBottom As POINT
Private mDBClickSourcePostion As POINT



'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As TRECT, ByVal wFormat As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Event OnMoveStart(ByRef blnCancel As Boolean)
Public Event OnMoveEnd()



Property Get Con1MinSize() As Long
    Con1MinSize = mlngCon1MinSize
End Property

Property Let Con1MinSize(value As Long)
    mlngCon1MinSize = value
End Property





Property Get Con2MinSize() As Long
    Con2MinSize = mlngCon2MinSize
End Property

Property Let Con2MinSize(value As Long)
    mlngCon2MinSize = value
End Property



Property Get StartDistance() As Long
    StartDistance = mlngStartDistance
End Property

Property Let StartDistance(value As Long)
    mlngStartDistance = value
End Property




Property Get AllowMove() As Boolean
    AllowMove = mblnAllowMove
End Property

Property Let AllowMove(value As Boolean)
    mblnAllowMove = value
    
    If Not mblnAllowMove Then
        CursorType = 0
    Else
        CursorType = IIf(mstSplitType = stHorizontal, 7, 9)
    End If
End Property



Property Get AllowPaintOtherSpliter() As Boolean
    AllowPaintOtherSpliter = mblnAllowPaintOtherSpliter
End Property

Property Let AllowPaintOtherSpliter(value As Boolean)
    mblnAllowPaintOtherSpliter = value
End Property




Property Get SyncParentHeight() As Boolean
    SyncParentHeight = mblnSyncParentHeight
End Property

Property Let SyncParentHeight(value As Boolean)
    mblnSyncParentHeight = value
End Property




Property Get SyncParentWidth() As Boolean
    SyncParentWidth = mblnSyncParentWidth
End Property

Property Let SyncParentWidth(value As Boolean)
    mblnSyncParentWidth = value
End Property




Property Get SplitType() As enmSplitType
    SplitType = mstSplitType
End Property

Property Let SplitType(value As enmSplitType)
    mstSplitType = value
    
    '设置鼠标指针类型
    CursorType = IIf(mstSplitType = stHorizontal, 7, 9)
    
    If mobjControl1 Is Nothing And mstSplitType = stHorizontal Then
        Extender.Width = 2000
    ElseIf mobjControl1 Is Nothing And mstSplitType = stVertical Then
        Extender.Height = 2000
    End If
    
    Call AdjustLayOut
End Property



'鼠标双击类型
Property Get DBClickType() As enmDBClickType
    DBClickType = mdtDBClickType
End Property

Property Let DBClickType(value As enmDBClickType)
    mdtDBClickType = value
End Property




'分割组件层次级别
Property Get SplitLevel() As enmSplitLevel
    SplitLevel = mslSplitLevel
End Property

Property Let SplitLevel(value As enmSplitLevel)
    mslSplitLevel = value
    
    Call AdjustLayOut
End Property





Property Get SplitWidth() As Long
    SplitWidth = mlngSplitWidth
End Property


Property Let SplitWidth(value As Long)
    mlngSplitWidth = value
    
    Call AdjustLayOut
End Property




'关联控件名称1
Property Get Control1Name() As String
On Error GoTo errHandle
    If mobjControl1 Is Nothing Then
        Control1Name = mstrControl1Name
    Else
        Control1Name = mobjControl1.Name
    End If
    
    Exit Sub
errHandle:
    Set mobjControl1 = Nothing
    Control1Name = mstrControl1Name
End Property


Property Let Control1Name(value As String)
    mstrControl1Name = value
    
    Set mobjControl1 = FindControl(mstrControl1Name)
    
    Call AdjustLayOut
End Property



'关联控件名称2
Property Get Control2Name() As String
On Error GoTo errHandle
    If mobjControl2 Is Nothing Then
        Control2Name = mstrControl2Name
    Else
        Control2Name = mobjControl2.Name
    End If
    
    Exit Sub
errHandle:
    Set mobjControl2 = Nothing
    Control2Name = mstrControl2Name
End Property


Property Let Control2Name(value As String)
    mstrControl2Name = value
    
    Set mobjControl2 = FindControl(mstrControl2Name)
    
    Call AdjustLayOut
End Property



'背景颜色
Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Property Let BackColor(value As OLE_COLOR)
    UserControl.BackColor = value
End Property


'边框样式
Property Get BorderStyle() As enmBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

Property Let BorderStyle(value As enmBorderStyle)
    UserControl.BorderStyle = value
End Property


'鼠标指针类型
Property Get CursorType() As enmCursorType
    CursorType = UserControl.MousePointer
End Property

Property Let CursorType(value As enmCursorType)
    UserControl.MousePointer = value
End Property

'鼠标指针
Property Get CursorIcon() As IPictureDisp
    Set CursorIcon = UserControl.MouseIcon
End Property

Property Set CursorIcon(value As IPictureDisp)
On Error GoTo errHandle
    Set UserControl.MouseIcon = value
    Exit Property
errHandle:
    MsgboxEx hwnd, err.Description, vbOKOnly, G_STR_HINT_TITLE
    err.Clear
End Property


'3D样式
Property Get Appearance() As enmAppearance
    Appearance = UserControl.Appearance
End Property


Property Let Appearance(value As enmAppearance)
    UserControl.Appearance = value
End Property


'Private Function RECT(ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngRight As Long, ByVal lngBottom As Long) As TRECT
'    RECT.Left = lngLeft
'    RECT.Top = lngTop
'
'    RECT.Right = IIf(ScaleMode <> vbPixels, ScaleX(lngRight, ScaleMode, vbPixels), lngRight)
'    RECT.Bottom = IIf(ScaleMode <> vbPixels, ScaleY(lngBottom, ScaleMode, vbPixels), lngBottom)
'End Function




Private Function FindControl(ByVal strControlName As String) As Object
'获取控件自身
    Dim i As Long

    Set FindControl = Nothing

    If UserControl.Parent Is Nothing Then Exit Function
    
'    For Each FindControl In Extender.Container
'        If UCase(FindControl.Name) = UCase(strControlName) Then Exit Function
'    Next
    
    For i = 0 To UserControl.ParentControls.Count - 1
        If UCase(UserControl.ParentControls.Item(i).Name) = UCase(strControlName) Then
            If UserControl.ParentControls.Item(i).Container Is Extender.Container Then
                If Not (TypeOf UserControl.ParentControls.Item(i) Is ucSplitter) Then
                    Set FindControl = UserControl.ParentControls.Item(i)
                    Exit Function
                End If
            End If
        End If
    Next i
End Function


Private Sub AdjustLayOut(Optional ByVal blnAllowRefreshContainer As Boolean = True)
On Error Resume Next
'调整布局
    mblnLayOutState = True
    
    If mstSplitType = stHorizontal Then

        If mobjControl1.Height >= Extender.Container.Height Then
            '档Control1的高度大于所在容器控件的高度时，调整到容器控件的范围内
            mobjControl1.Height = Extender.Container.Height - Fix(Extender.Container.Height / 5)
        End If

        
        Extender.Height = mlngSplitWidth
        
        
        If (mslSplitLevel = slFirst Or mslSplitLevel = slFirstAndEnd) And Not (mobjControl1 Is Nothing) Then
            mobjControl1.Top = mlngStartDistance
            
            If mblnSyncParentWidth Then mobjControl1.Left = 0
            If mblnSyncParentWidth Then mobjControl1.Width = Extender.Container.ScaleWidth
        End If
        
        
        If Not (mobjControl1 Is Nothing) Then
            
            Extender.Left = mobjControl1.Left
            Extender.Top = mobjControl1.Top + mobjControl1.Height
            Extender.Width = mobjControl1.Width
        End If
        
        '调整control2的位置大小
        If Not (mobjControl2 Is Nothing) Then
            mobjControl2.Left = Extender.Left
            mobjControl2.Top = Extender.Top + Extender.Height
            
            If mobjControl1 Is Nothing Then
                Extender.Width = mobjControl2.Width
            Else
                mobjControl2.Width = Extender.Width
            End If
        End If
                
        
        If (mslSplitLevel = slEnd Or mslSplitLevel = slFirstAndEnd) And Not (mobjControl2 Is Nothing) Then
            If mblnSyncParentWidth Then mobjControl2.Left = 0
            If mblnSyncParentWidth Then mobjControl2.Width = Extender.Container.ScaleWidth
            mobjControl2.Height = Extender.Container.ScaleHeight - mobjControl2.Top
        End If
        
        
'        '修正control2位置及大小
'        If Extender.Top + mlngCon2MinSize >= Extender.Container.ScaleHeight And mlngOldPostion = 0 Then
'            mlngOldPostion = Extender.Container.ScaleHeight
'        End If
'
'        If Extender.Container.ScaleHeight > mlngOldPostion Then mlngOldPostion = 0
'
'        If mlngOldPostion <> 0 Then
'            Extender.Top = Extender.Container.ScaleHeight - mlngCon2MinSize
'
'            mobjControl1.Height = Extender.Top
'
'            mobjControl2.Top = Extender.Top + mlngSplitWidth
'            mobjControl2.Hegith = Extender.Container.ScaleHeight - Extender.Top - mlngSplitWidth
'        End If
        
    Else

        If mobjControl1.Width > Extender.Container.Width Then
            '档Control1的宽度大于所在容器控件的宽度时，调整到容器控件的范围内
            mobjControl1.Width = Extender.Container.Width - Fix(Extender.Container.Width / 5)
        End If

        
        Extender.Width = mlngSplitWidth
        
        If (mslSplitLevel = slFirst Or mslSplitLevel = slFirstAndEnd) And Not (mobjControl1 Is Nothing) Then
            mobjControl1.Left = mlngStartDistance
            If mblnSyncParentHeight Then mobjControl1.Top = 0
            If mblnSyncParentHeight Then mobjControl1.Height = Extender.Container.ScaleHeight
        End If
        
        If Not (mobjControl1 Is Nothing) Then
            
            Extender.Left = mobjControl1.Left + mobjControl1.Width
            Extender.Top = mobjControl1.Top
            Extender.Height = mobjControl1.Height
        End If
        
        '调整control2的位置大小
        If Not (mobjControl2 Is Nothing) Then
            mobjControl2.Left = Extender.Left + Extender.Width
            mobjControl2.Top = Extender.Top
            
            If mobjControl1 Is Nothing Then
                Extender.Height = mobjControl2.Height
            Else
                mobjControl2.Height = Extender.Height
            End If
        End If
        
        
        If (mslSplitLevel = slEnd Or mslSplitLevel = slFirstAndEnd) And Not (mobjControl2 Is Nothing) Then
            If mblnSyncParentHeight Then mobjControl2.Top = 0
            If mblnSyncParentHeight Then mobjControl2.Height = Extender.Container.ScaleHeight
            mobjControl2.Width = Extender.Container.ScaleWidth - mobjControl2.Left
        End If

    End If
    
    If blnAllowRefreshContainer Then Extender.Container.Refresh
    
    mblnLayOutState = False
    
    err.Clear
End Sub


Private Sub UserControl_DblClick()
On Error Resume Next
    Dim lngNum As Long
    
    Select Case mdtDBClickType
        Case dtNone
            Exit Sub
        Case dtHideControl1
            If mobjControl1 Is Nothing Then Exit Sub
            
            If Not mblnDbClickState Then
                '保存组件位置
                mDBClickSourcePostion.X = mobjControl1.Width
                mDBClickSourcePostion.Y = mobjControl1.Height
                
                mblnDbClickState = True
    
                '隐藏控件1
                If mstSplitType = stHorizontal Then
                    mobjControl1.Height = 0
                    Extender.Top = mobjControl1.Top + mobjControl1.Height
                    
                    If Not (mobjControl2 Is Nothing) Then
                        mobjControl2.Height = mobjControl2.Height + (mobjControl2.Top - (Extender.Top + mlngSplitWidth))
                        mobjControl2.Top = Extender.Top + mlngSplitWidth
                    End If
                Else
                    mobjControl1.Width = 0
                    Extender.Left = mobjControl1.Left + mobjControl1.Width
                    
                    If Not (mobjControl2 Is Nothing) Then
                        mobjControl2.Width = mobjControl2.Width + (mobjControl2.Left - (Extender.Left + mlngSplitWidth))
                        mobjControl2.Left = Extender.Left + mlngSplitWidth
                    End If
                End If
            Else
                mblnDbClickState = False
                
                '显示控件1
                If mstSplitType = stHorizontal Then
                    mobjControl1.Height = mDBClickSourcePostion.Y
                    Extender.Top = mobjControl1.Top + mobjControl1.Height
                    
                    If Not (mobjControl2 Is Nothing) Then
                        mobjControl2.Height = mobjControl2.Height - mDBClickSourcePostion.Y
                        mobjControl2.Top = Extender.Top + mlngSplitWidth
                    End If
                Else
                    mobjControl1.Width = mDBClickSourcePostion.X
                    Extender.Left = mobjControl1.Left + mobjControl1.Width
                    
                    If Not (mobjControl2 Is Nothing) Then
                        mobjControl2.Width = mobjControl2.Width - mDBClickSourcePostion.X
                        mobjControl2.Left = Extender.Left + mlngSplitWidth
                    End If
                End If
            End If
            
            
        Case dtHideControl2
            If mobjControl2 Is Nothing Then Exit Sub
            
            If Not mblnDbClickState Then
                '保存组件位置
                mDBClickSourcePostion.X = mobjControl2.Width
                mDBClickSourcePostion.Y = mobjControl2.Height
                
                mblnDbClickState = True
                    
                '隐藏控件2
                If mstSplitType = stHorizontal Then
                    mobjControl2.Top = mobjControl2.Top + mobjControl2.Height
                    mobjControl2.Height = 0
                    
                    '重新计算控件2的位置（某些控件的高度不能小于某个值）
                    mobjControl2.Top = mobjControl2.Top - mobjControl2.Height
                    
                    Extender.Top = mobjControl2.Top - mlngSplitWidth
                    
                    If Not (mobjControl1 Is Nothing) Then
                        mobjControl1.Height = mobjControl1.Height + (Extender.Top - (mobjControl1.Top + mobjControl1.Height))
                    End If
                Else
                    mobjControl2.Left = mobjControl2.Left + mobjControl2.Width
                    mobjControl2.Width = 0
                    
                    '重新计算控件2的位置（某些控件的宽度不能小于某个值）
                    mobjControl2.Left = mobjControl2.Left - mobjControl2.Width
                    
                    Extender.Left = mobjControl2.Left - mlngSplitWidth
                    
                    If Not (mobjControl1 Is Nothing) Then
                        mobjControl1.Width = mobjControl1.Width + (Extender.Left - (mobjControl1.Left + mobjControl1.Width))
                    End If
                End If
            Else
                mblnDbClickState = False
                
                
                '显示控件2
                If mstSplitType = stHorizontal Then
                    lngNum = mobjControl2.Height    '获取控件的最小高度，某些控件的高度不能为0
                    
                    mobjControl2.Top = mobjControl2.Top - mDBClickSourcePostion.Y + lngNum
                    mobjControl2.Height = mDBClickSourcePostion.Y
                        
                    
                    Extender.Top = mobjControl2.Top - mlngSplitWidth
                    
                    If Not (mobjControl1 Is Nothing) Then
                        mobjControl1.Height = mobjControl1.Height - mDBClickSourcePostion.Y + lngNum
                    End If
                Else
                    lngNum = mobjControl2.Width '获取控件的最小宽度，某些控件的宽度不能为0
                    
                    mobjControl2.Left = mobjControl2.Left - mDBClickSourcePostion.X + lngNum
                    mobjControl2.Width = mDBClickSourcePostion.X
                    
                    Extender.Left = mobjControl2.Left - mlngSplitWidth
                    
                    If Not (mobjControl2 Is Nothing) Then
                        mobjControl1.Width = mobjControl1.Width - mDBClickSourcePostion.X + lngNum
                    End If
                End If
            End If
    End Select
    
    Extender.Container.Refresh
    
    Call PaintOtherSplitter
    
    err.Clear
End Sub

Private Sub UserControl_Initialize()
    mblnMouseDown = False
    mblnLayOutState = False
    mblnDbClickState = False
End Sub

Private Sub UserControl_InitProperties()
    mstSplitType = stVertical
    mdtDBClickType = dtNone
    
    UserControl.Appearance = 1
    UserControl.BorderStyle = 0
    UserControl.BackColor = &HE0E0E0
    UserControl.MousePointer = 9
    
    
    mlngSplitWidth = 135
    mstrControl1Name = ""
    mstrControl2Name = ""
    mblnAllowMove = True
    mblnSyncParentHeight = True
    mblnSyncParentWidth = True
    mblnAllowPaintOtherSpliter = False
    mlngCon1MinSize = 0
    mlngCon2MinSize = 0
    mlngStartDistance = 0
    
    Set mobjControl1 = Nothing
    Set mobjControl2 = Nothing
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim blnCancel As Boolean

    blnCancel = False
    RaiseEvent OnMoveStart(blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    mblnMouseDown = True And mblnAllowMove
    
    mControl2RightBottom.X = 0
    mControl2RightBottom.Y = 0
    
    If Not (mobjControl2 Is Nothing) Then
        '保存第二个组件的右下角坐标
        mControl2RightBottom.X = mobjControl2.Left + mobjControl2.Width
        mControl2RightBottom.Y = mobjControl2.Top + mobjControl2.Height
    End If
            
    '锁定鼠标到当前控件
    SetCapture hwnd
    
    err.Clear
End Sub


Private Function IsSameContainer(obj As Object) As Boolean
'递归判断当前对象是否处于相同的容器组件中
    IsSameContainer = False
    
    If obj Is Nothing Then Exit Function
    
    If obj Is Extender.Container Then
        IsSameContainer = True
        Exit Function
    End If
    
On Error GoTo errHandle
    IsSameContainer = IsSameContainer(obj.Container)
    Exit Function
errHandle:
    IsSameContainer = False
    err.Clear
End Function


Private Sub PaintOtherSplitter()
On Error Resume Next
'绘制其他分割控件
    Dim i As Long
    Dim objControl As Object

    If Not mblnAllowPaintOtherSpliter Then Exit Sub
    If UserControl.Parent Is Nothing Then Exit Sub
    
    For i = 0 To UserControl.ParentControls.Count - 1
        If TypeName(UserControl.ParentControls.Item(i)) = "ucSplitter" Then
            Set objControl = UserControl.ParentControls.Item(i)
            If IsSameContainer(objControl) Then objControl.RePaint
        End If
    Next i
    err.Clear
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errHandle
    Dim lngDistance As Long
    
    If Not mblnMouseDown Then Exit Sub
    
    mblnDbClickState = False
    mblnLayOutState = True
    
'    Extender.Parent.Refresh
    
    
    If mstSplitType = stHorizontal Then
        '纵向调节
        lngDistance = Y
        
        '如果超出容器控件的范围，则不允许进行调整
        If Extender.Top + lngDistance < 0 Then Exit Sub
        If Extender.Top + lngDistance + mlngSplitWidth > Extender.Container.ScaleHeight Then Exit Sub
        
        '调整组件的上下位置
        If Not (mobjControl1 Is Nothing) Then
            If mobjControl1.Height + lngDistance <= 0 Then Exit Sub
            
            
            If Not (mobjControl2 Is Nothing) Then
                If Extender.Top + lngDistance + mlngSplitWidth >= mControl2RightBottom.Y Or _
                  Extender.Top + lngDistance + mlngSplitWidth >= mControl2RightBottom.Y - mlngCon2MinSize And lngDistance >= 0 Then Exit Sub
            End If
            
            '当移动距离小于指定范围时，则结束移动(lngDistance小于零表示向上移动)
            If Extender.Top + lngDistance < mobjControl1.Top + mlngCon1MinSize And lngDistance <= 0 Then Exit Sub
            mobjControl1.Height = mobjControl1.Height + lngDistance
            
            '判断是否允许继续调整
            If Extender.Top + lngDistance < mobjControl1.Top + mobjControl1.Height Then Exit Sub
        End If


        If Not (mobjControl2 Is Nothing) Then
        
            If mobjControl2.Height - lngDistance <= 0 Then Exit Sub
        
            '当移动距离小于指定范围时，则结束移动
            If Extender.Top + lngDistance + mlngSplitWidth >= mControl2RightBottom.Y Or _
                Extender.Top + lngDistance + mlngSplitWidth > mControl2RightBottom.Y - mlngCon2MinSize And lngDistance >= 0 Then
                '恢复control1的调整
                If Not (mobjControl1 Is Nothing) Then mobjControl1.Height = Extender.Top - mobjControl1.Top
                Exit Sub
            End If
            
            
            mobjControl2.Top = mobjControl2.Top + lngDistance
            mobjControl2.Height = mControl2RightBottom.Y - mobjControl2.Top
            
            '如果当前调整的范围，大于控件2的右下角坐标，则不允许调整
            If (Extender.Top + lngDistance + mlngSplitWidth) + mobjControl2.Height > mControl2RightBottom.Y Or _
                (Extender.Top + lngDistance + mlngSplitWidth) + mobjControl2.Height > mControl2RightBottom.Y - mlngCon2MinSize And lngDistance >= 0 Then

                mobjControl2.Top = mControl2RightBottom.Y - mobjControl2.Height
                Extender.Top = mobjControl2.Top - mlngSplitWidth
                
                If Not (mobjControl1 Is Nothing) Then mobjControl1.Height = Extender.Top - mobjControl1.Top
                
                Exit Sub
            End If
            
'            If Extender.Top + lngDistance + mlngSplitWidth >= mControl2RightBottom.Y Then Exit Sub
        End If
        
        Extender.Top = Extender.Top + lngDistance
    Else
        '横向调节
        lngDistance = X
        
        '如果超出容器控件的范围，则不允许进行调整
        If Extender.Left + lngDistance < 0 Then Exit Sub
        If Extender.Left + lngDistance + mlngSplitWidth > Extender.Container.ScaleWidth And lngDistance <= 0 Then Exit Sub
        
        '调整组件的左右位置
        If Not (mobjControl1 Is Nothing) Then
            If mobjControl1.Width + lngDistance <= 0 Then Exit Sub
            
            If Not (mobjControl2 Is Nothing) Then
                If Extender.Left + lngDistance + mlngSplitWidth + 1 >= mControl2RightBottom.X Or _
                   Extender.Left + lngDistance + mlngSplitWidth + 1 >= mControl2RightBottom.X - mlngCon2MinSize And lngDistance >= 0 Then Exit Sub
            End If
            
            'lngDistance小于零表示向左移动
            If Extender.Left + lngDistance <= mobjControl1.Left + mlngCon1MinSize And lngDistance <= 0 Then Exit Sub
            mobjControl1.Width = mobjControl1.Width + lngDistance
            
            '判断是否允许继续调整
            If Extender.Left + lngDistance < mobjControl1.Left + mobjControl1.Width Then Exit Sub
        End If
        
        If Not (mobjControl2 Is Nothing) Then
            If mobjControl2.Width - lngDistance <= 0 Then Exit Sub
            
            If Extender.Left + lngDistance + mlngSplitWidth + 1 >= mControl2RightBottom.X Or _
                Extender.Left + lngDistance + mlngSplitWidth + 1 >= mControl2RightBottom.X - mlngCon2MinSize And lngDistance >= 0 Then
                '恢复control1的调整
                If Not (mobjControl1 Is Nothing) Then mobjControl1.Width = Extender.Left - mobjControl1.Left
                Exit Sub
            End If
            
            mobjControl2.Left = mobjControl2.Left + lngDistance
            mobjControl2.Width = mControl2RightBottom.X - mobjControl2.Left
            
            '如果当前调整的范围，大于控件2的右下角坐标，则不允许调整
            If (Extender.Left + lngDistance + mlngSplitWidth) + mobjControl2.Width > mControl2RightBottom.X Or _
                (Extender.Left + lngDistance + mlngSplitWidth) + mobjControl2.Width > mControl2RightBottom.X - mlngCon2MinSize And lngDistance >= 0 Then

                mobjControl2.Left = mControl2RightBottom.X - mobjControl2.Width
                Extender.Left = mobjControl2.Left - mlngSplitWidth
                
                If Not (mobjControl1 Is Nothing) Then mobjControl1.Width = Extender.Left - mobjControl1.Left
                
                Exit Sub
            End If
            
'            If Extender.Left + lngDistance + mlngSplitWidth + 1 >= mControl2RightBottom.X Then Exit Sub
        End If
        
        Extender.Left = Extender.Left + lngDistance
    End If
    
    'Extender.Container.Refresh 'MODIFY:2013-09-10
    
    mblnLayOutState = False
    
    Exit Sub
errHandle:
    err.Clear
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If mblnMouseDown Then Call PaintOtherSplitter
    
    mblnMouseDown = False
    
    '释放鼠标
    Call ReleaseCapture
    
    RaiseEvent OnMoveEnd
    
    err.Clear
End Sub

Public Sub RePaint(Optional ByVal blnAllowRefreshContainer As Boolean = True)
On Error Resume Next
    Call Extender.ZOrder(0)
    
    If mobjControl1 Is Nothing Then Set mobjControl1 = FindControl(mstrControl1Name)
    If mobjControl2 Is Nothing Then Set mobjControl2 = FindControl(mstrControl2Name)
    
    Call AdjustLayOut(blnAllowRefreshContainer)
    
    err.Clear
End Sub

Public Sub RedrawSelf()
'刷新
    Call Refresh
End Sub
 
Private Sub UserControl_Paint()
On Error Resume Next
    If mblnLayOutState Then Exit Sub
    
    Call RePaint(False)
    err.Clear
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'读取控件属性
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HE0E0E0)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 9)
    UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    
    mlngSplitWidth = PropBag.ReadProperty("SplitWidth", 135)
    mstSplitType = PropBag.ReadProperty("SplitType", stVertical)
    mslSplitLevel = PropBag.ReadProperty("SplitLevel", slNone)
    mdtDBClickType = PropBag.ReadProperty("DBClickType", dtNone)
    mblnAllowMove = PropBag.ReadProperty("AllowMove", True)
    mblnSyncParentWidth = PropBag.ReadProperty("SyncParentWidth", True)
    mblnSyncParentHeight = PropBag.ReadProperty("SyncParentHeight", True)
    mblnAllowPaintOtherSpliter = PropBag.ReadProperty("AllowPaintOtherSpliter", False)
    mlngCon1MinSize = PropBag.ReadProperty("Con1MinSize", 0)
    mlngCon2MinSize = PropBag.ReadProperty("Con2MinSize", 0)
    mlngStartDistance = PropBag.ReadProperty("StartDistance", 0)
    
    
    mstrControl1Name = PropBag.ReadProperty("Control1Name", "")
    mstrControl2Name = PropBag.ReadProperty("Control2Name", "")
    
    Set mobjControl1 = FindControl(mstrControl1Name)
    Set mobjControl2 = FindControl(mstrControl2Name)
End Sub



Private Sub UserControl_Resize()
On Error Resume Next
    If mblnLayOutState Then Exit Sub

    Call AdjustLayOut
    
    err.Clear
End Sub

Private Sub UserControl_Show()
'On Error Resume Next
'    Call Extender.ZOrder(0)
'
'    If mobjControl1 Is Nothing Then Set mobjControl1 = FindControl(mstrControl1Name)
'    If mobjControl2 Is Nothing Then Set mobjControl2 = FindControl(mstrControl2Name)
'
'    Call AdjustLayOut
'
'    Err.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'保存控件属性
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 9)
    Call PropBag.WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)
    
    Call PropBag.WriteProperty("SplitWidth", mlngSplitWidth, 135)
    Call PropBag.WriteProperty("SplitType", mstSplitType, stVertical)
    Call PropBag.WriteProperty("DBClickType", mdtDBClickType, dtNone)
    Call PropBag.WriteProperty("SplitLevel", mslSplitLevel, slNone)
    Call PropBag.WriteProperty("AllowMove", mblnAllowMove, True)
    Call PropBag.WriteProperty("SyncParentWidth", mblnSyncParentWidth, True)
    Call PropBag.WriteProperty("SyncParentHeight", mblnSyncParentHeight, True)
    Call PropBag.WriteProperty("AllowPaintOtherSpliter", mblnAllowPaintOtherSpliter, False)
    Call PropBag.WriteProperty("Con1MinSize", mlngCon1MinSize, 0)
    Call PropBag.WriteProperty("Con2MinSize", mlngCon2MinSize, 0)
    Call PropBag.WriteProperty("StartDistance", mlngStartDistance, 0)
    
    
On Error GoTo errControl1
    If mobjControl1 Is Nothing Then
        Call PropBag.WriteProperty("Control1Name", mstrControl1Name, "")
    Else
        Call PropBag.WriteProperty("Control1Name", mobjControl1.Name, "")
    End If
    
    GoTo continue
errControl1:
    Call PropBag.WriteProperty("Control1Name", mstrControl1Name, "")
    err.Clear
    
continue:
On Error GoTo errControl2
    
    If mobjControl2 Is Nothing Then
        Call PropBag.WriteProperty("Control2Name", mstrControl2Name, "")
    Else
        Call PropBag.WriteProperty("Control2Name", mobjControl2.Name, "")
    End If
    
    Exit Sub
errControl2:
    Call PropBag.WriteProperty("Control2Name", mstrControl2Name, "")
End Sub
