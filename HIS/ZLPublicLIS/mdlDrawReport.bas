Attribute VB_Name = "mdlDrawReport"
Option Explicit

'定义
'######################################################################################################################
'
'######################################################################################################################
Public Function AppendPrintData(ByVal str类别 As String, _
                                ByVal str对象 As String, _
                                Optional ByVal bytHAlignment As Byte = 1, _
                                Optional ByVal blnWordWarp As Boolean, _
                                Optional ByVal str内容 As String, _
                                Optional ByVal bytVAlignment As Byte = 2, _
                                Optional ByVal blnMuliLine As Boolean, _
                                Optional ByVal intRows As Integer = 1, _
                                Optional ByVal blnAutoFit As Boolean = False, _
                                Optional ByVal blnDebug As Boolean = False, _
                                Optional ByVal strPrex As String = "A", _
                                Optional ByVal bytAngle As Integer = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngRows     As Long
    Dim strTmp      As String
    Dim lngLoop     As Long
    Dim strLineText As String
    Dim lngDiff     As Long
    Dim rsLine      As ADODB.Recordset
    Dim intLoop     As Integer
    
On Error GoTo errHand

    lngRows = 1
    strTmp = str内容
    
    Select Case str类别
    '------------------------------------------------------------------------------------------------------------------
    Case "封面", "页眉", "页脚"
    
        Select Case str对象
        '--------------------------------------------------------------------------------------------------------------
        Case "文本", "页码"
    
            gobjDraw.FontName = gobjFont.Name
            gobjDraw.FontSize = gobjFont.Size

            If gobjRect.Y1 = 0 Then gobjRect.Y1 = gobjRect.Y0 + gobjDraw.TextHeight("高")
        End Select
        
        Call InsertData(str类别, str对象, bytHAlignment, strTmp, bytVAlignment, blnWordWarp, intRows, , blnAutoFit, blnDebug, strPrex, bytAngle)
        
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        
        Select Case str对象
        '--------------------------------------------------------------------------------------------------------------
        Case "文本", "续页"

            gobjDraw.FontName = gobjFont.Name
            gobjDraw.FontSize = gobjFont.Size
            
            lngRows = 1
            If blnMuliLine Then
                Set rsLine = GetLineText(gobjDraw, str内容, gobjRect.X1 - gobjRect.X0)
                lngRows = rsLine.RecordCount
                If lngRows > 1 Then
                    rsLine.MoveFirst
                    Do While Not rsLine.EOF
                        strLineText = rsLine("内容").Value
                        intLoop = intLoop + 1
                        If intLoop > 1 Then
                            gobjRect.Y0 = gobjRect.Y0 + (gobjDraw.TextHeight("高") + gobjRect.R0)
                        Else
                            gobjRect.Y0 = gobjRect.Y0 + gobjRect.R0
                        End If
                        
                        gobjRect.Y1 = gobjRect.Y0 + gobjDraw.TextHeight("高")
    
                        If str对象 <> "续页" Then
                            If Val(gobjRect.Y1) > Val(gobjPaper.Height - gobjPaper.BorderBottom - gobjPaper.PageFoot - gobjPaper.SpaceBottom) Then
                                Call NewPage
                            End If
                        End If
    
                        Call InsertData(str类别, str对象, bytHAlignment, strLineText, bytVAlignment, blnWordWarp, lngRows, , blnAutoFit, blnDebug, strPrex, bytAngle)
    
                        rsLine.MoveNext
                    Loop
                End If
            End If
            
            If lngRows <= 1 Then
                If blnMuliLine Then
                    gobjRect.Y1 = gobjRect.Y0 + gobjDraw.TextHeight("高")
                Else
                    If gobjRect.Y1 = 0 Then gobjRect.Y1 = gobjRect.Y0 + gobjDraw.TextHeight("高")
                End If
                                    
                If str对象 <> "续页" Then
                    If Val(gobjRect.Y1) > Val(gobjPaper.Height - gobjPaper.BorderBottom - gobjPaper.PageFoot - gobjPaper.SpaceBottom) Then
                        Call NewPage
                    End If
                End If
                
                Call InsertData(str类别, str对象, bytHAlignment, strTmp, bytVAlignment, blnWordWarp, intRows, , blnAutoFit, blnDebug, strPrex, bytAngle)
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            
            If Val(gobjRect.Y1) > Val(gobjPaper.Height - gobjPaper.BorderBottom - gobjPaper.PageFoot - gobjPaper.SpaceBottom) Then
                Call NewPage
            End If
            
            Call InsertData(str类别, str对象, bytHAlignment, strTmp, bytVAlignment, blnWordWarp, intRows, , blnAutoFit, blnDebug, strPrex, bytAngle)
        
        End Select
        
    End Select

    AppendPrintData = True
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function NewPage(Optional ByVal bytDoeal As Byte = 0, Optional ByVal blnDrawLine As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim lngDiff As Long

    On Error GoTo errHand
    If blnDrawLine Then Call BeforeNewPage(bytDoeal)

    gobjRect.Page = gobjRect.Page + 1
    Call InsertPage(gobjRect.Page)

    If blnDrawLine Then Call AfterNewPage(bytDoeal)

    lngDiff = gobjRect.Y0 - (gobjPaper.BorderTop + gobjPaper.PageHead + gclsLisReportLib.GetTwipsX(0.5))
    gobjRect.Y0 = gobjRect.Y0 - lngDiff
    If gobjRect.Y1 <> 0 Then gobjRect.Y1 = gobjRect.Y1 - lngDiff

    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function InsertPage(ByVal intPage As Long, Optional ByVal bytCalc As Byte = 1, Optional ByVal strShow As String, Optional ByVal blnShowPageHead As Boolean = True, Optional ByVal blnShowPageFoot As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    On Error GoTo errHand
    
    If bytCalc = 0 Then glngVirtualPages = glngVirtualPages + 1
    
    grsPage.AddNew
    grsPage("页号").Value = intPage
    grsPage("虚拟页号").Value = 0
    grsPage("虚拟总页").Value = 0
    grsPage("总页").Value = 0
    grsPage("页码计算").Value = bytCalc
    grsPage("显示页眉").Value = IIf(blnShowPageHead, 1, 0)
    grsPage("显示页脚").Value = IIf(blnShowPageFoot, 1, 0)
    grsPage("显示内容").Value = strShow
    
    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function BeforeNewPage(ByVal bytDoeal As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objSvrRect As USERRECT
    Dim objSvrFont As USERFONT

    On Error GoTo errHand

    If bytDoeal = 1 Then

        Call SaveRect(gobjRect, objSvrRect)
        Call SaveFont(gobjFont, objSvrFont)

        '补线
        gobjFont.ForeColor = USERCOLOR.表格线色
        gobjRect.X0 = gobjPaper.BorderLeft
        gobjRect.Y0 = gobjRect.Y1
        gobjRect.X1 = gobjPaper.Width - gobjPaper.BorderRight
        gobjRect.Y1 = gobjRect.Y0
        Call AppendPrintData("项目", "线条")

        Call SaveRect(objSvrRect, gobjRect)
        Call SaveFont(objSvrFont, gobjFont)

    End If

    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function AfterNewPage(ByVal bytDoeal As Byte) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objSvrRect As USERRECT
    Dim objSvrFont As USERFONT

    On Error GoTo errHand

    If bytDoeal = 1 Then

        Call SaveRect(gobjRect, objSvrRect)
        Call SaveFont(gobjFont, objSvrFont)

        '补线
        gobjFont.ForeColor = USERCOLOR.表格线色
        gobjRect.X0 = gobjPaper.BorderLeft
        gobjRect.Y0 = gobjPaper.BorderTop + gobjPaper.PageHead + gobjPaper.SpaceTop
        gobjRect.X1 = gobjPaper.Width - gobjPaper.BorderRight
        gobjRect.Y1 = gobjRect.Y0
        Call AppendPrintData("项目", "线条")

        Call SaveRect(objSvrRect, gobjRect)
        Call SaveFont(objSvrFont, gobjFont)

    End If

    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

'Public Function GetTwipsX(ByVal sglNumber As Single) As Single
'    GetTwipsX = gobjDraw.ScaleX(sglNumber, vbCentimeters, vbTwips)
'End Function
'
'Public Function GetTwipsY(ByVal sglNumber As Single) As Single
'    GetTwipsY = gobjDraw.ScaleY(sglNumber, vbCentimeters, vbTwips)
'End Function
'
'Public Function GetCentimetersX(ByVal sglNumber As Single) As Single
'    GetCentimetersX = gobjDraw.ScaleX(sglNumber, vbTwips, vbCentimeters)
'End Function
'
'Public Function GetCentimetersY(ByVal sglNumber As Single) As Single
'    GetCentimetersY = gobjDraw.ScaleY(sglNumber, vbTwips, vbCentimeters)
'End Function

Public Function SetRect(ByRef objRect As USERRECT, ByVal X0 As Long, ByVal Y0 As Long, ByVal X1 As Long, ByVal Y1 As Long, Optional ByVal B0 As Long, Optional ByVal R0 As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    objRect.X0 = X0
    objRect.Y0 = Y0
    objRect.X1 = X1
    objRect.Y1 = Y1
    objRect.B0 = B0
    objRect.R0 = R0
End Function

Public Function SaveRect(ByRef objFromRect As USERRECT, ByRef objToRect As USERRECT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    objToRect.X0 = objFromRect.X0
    objToRect.Y0 = objFromRect.Y0
    objToRect.X1 = objFromRect.X1
    objToRect.Y1 = objFromRect.Y1
    objToRect.B0 = objFromRect.B0
    objToRect.R0 = objFromRect.R0
    
End Function

Public Function SaveFont(ByRef objFromFont As USERFONT, ByRef objToFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    objToFont.Name = objFromFont.Name
    objToFont.Size = objFromFont.Size
    objToFont.Bold = objFromFont.Bold
    objToFont.Italic = objFromFont.Italic
    objToFont.Underline = objFromFont.Underline
    objToFont.BackColor = objFromFont.BackColor
    objToFont.ForeColor = objFromFont.ForeColor
    objToFont.LineWidth = objFromFont.LineWidth
    
End Function

Public Function SetDraw(ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    gobjDraw.FontName = objFont.Name
    gobjDraw.FontSize = objFont.Size
    gobjDraw.FontBold = objFont.Bold
    gobjDraw.FontItalic = objFont.Italic

End Function

Public Function GetFullFilePath(ByVal strText As String) As String
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    intPos1 = InStr(strText, "<Image>")
    intPos2 = InStr(strText, "</Image>")
    If intPos1 > 0 And intPos2 > 0 And intPos1 < intPos2 Then
        strText = Mid(strText, intPos1 + 7, intPos2 - intPos1 - 7)
        If Dir(strText) <> "" And strText <> "" Then
            GetFullFilePath = strText
        End If
    End If
    
End Function
