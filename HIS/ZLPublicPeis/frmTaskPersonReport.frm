VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTaskPersonReport 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '????ȱʡ
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5850
      Index           =   1
      Left            =   4380
      ScaleHeight     =   5850
      ScaleWidth      =   8250
      TabIndex        =   2
      Top             =   1185
      Width           =   8250
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   5010
         Left            =   390
         TabIndex        =   3
         Top             =   345
         Width           =   7800
         _cx             =   13758
         _cy             =   8837
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5850
      Index           =   0
      Left            =   120
      ScaleHeight     =   5850
      ScaleWidth      =   3570
      TabIndex        =   0
      Top             =   930
      Width           =   3570
      Begin MSComctlLib.TreeView tvw 
         Height          =   3465
         Left            =   45
         TabIndex        =   1
         Top             =   225
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   6112
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   6675
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTaskPersonReport.frx":0000
            Key             =   "????"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTaskPersonReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsVsf As clsVsf

'######################################################################################################################
'******************************************************************************************************************
'???ܣ???ʼ??????
'??????
'???أ?
'******************************************************************************************************************
Public Sub InitData()

    Call InitVsf
     
     Set tvw.ImageList = ils16
End Sub

'******************************************************************************************************************
'???ܣ?????????
'??????
'???أ?
'******************************************************************************************************************
Public Function LoadData(ByVal lngTaskKey As Long, ByVal lngPersonKey As Long) As Boolean
    
    LoadData = LoadTaskReport(lngTaskKey, lngPersonKey)
End Function

'******************************************************************************************************************
'???ܣ????????񱨸?
'??????
'???أ?
'******************************************************************************************************************
Private Function LoadTaskReport(ByVal lngTaskKey As Long, ByVal lngPersonKey As Long) As Boolean
    
    On Error GoTo errHand
    
    Call LoadItem(lngTaskKey, lngPersonKey)
    Call LoadResult(lngTaskKey, lngPersonKey)
    
    LoadTaskReport = True
    
    Exit Function
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

'******************************************************************************************************************
'???ܣ???????Ŀ
'??????
'???أ?
'******************************************************************************************************************
Private Sub LoadItem(ByVal lngTaskKey As Long, ByVal lngPersonKey)

    Dim rsData As ADODB.Recordset
    Dim objNode As Node
    
    On Error GoTo errHand
    
    tvw.Nodes.Clear
    tvw.Style = tvwPlusPictureText
    
    Set rsData = gclsPackage.Get_PeisPersonItem(lngTaskKey, lngPersonKey)
    
    Do While Not rsData.EOF
         Set objNode = tvw.Nodes.Add(, , "K" & NVL(rsData("?嵥ID").Value), NVL(rsData("??Ŀ").Value), "????", "????")
         
         rsData.MoveNext
    Loop
    
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Sub


'******************************************************************************************************************
'???ܣ????ؽ???
'??????
'???أ?
'******************************************************************************************************************
Private Sub LoadResult(ByVal lngTaskKey As Long, ByVal lngPersonKey As Long)
    Dim rsConclusion As ADODB.Recordset
    Dim rsItem As ADODB.Recordset
    Dim rsResult As ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo errHand
    
    mclsVsf.ClearGrid
    
    
     With vsf
        
        
        '??ȡ?ܼ?????
        Set rsConclusion = gclsPackage.Get_PeisPersonConclusion(2, lngTaskKey, lngPersonKey)
        If rsConclusion.BOF = False Then
            .Row = .Rows - 1
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
            .MergeRow(.Row) = True
            .TextMatrix(.Row, .ColIndex("ID")) = NVL(rsConclusion("ID").Value)
            .TextMatrix(.Row, .ColIndex("??Ŀ")) = "?ܼ?????"
            .Cell(flexcpData, .Row, .ColIndex("??Ŀ"), .Row, .Cols - 1) = "?ܼ?????"
            .Cell(flexcpText, .Row, .ColIndex("??Ŀ"), .Row, .Cols - 1) = "?ܼ?????"
            
            Do While Not rsConclusion.EOF
                If Trim(.TextMatrix(.Rows - 1, .ColIndex("??Ŀ"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("ָ??"))) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                If rsConclusion.AbsolutePosition = 1 Then .TextMatrix(.Row, .ColIndex("????")) = 1
                .TextMatrix(.Row, .ColIndex("ָ??")) = NVL(rsConclusion("????????").Value)
                rsConclusion.MoveNext
            Loop
            
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("ָ??"))) <> "" Then .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, .ColIndex("??Ŀ")) = " "
        End If
        
        '??ȡ??????Ŀ
        Set rsItem = gclsPackage.Get_PeisPersonItem(lngTaskKey, lngPersonKey)
        Do While Not rsItem.EOF
           
           If .TextMatrix(.Rows - 1, .ColIndex("??Ŀ")) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("ָ??"))) <> "" Then .Rows = .Rows + 1
           .Row = .Rows - 1
           
            .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
            .MergeRow(.Row) = True
           .TextMatrix(.Row, .ColIndex("ID")) = NVL(rsItem("?嵥ID").Value)
           
            .TextMatrix(.Row, .ColIndex("??Ŀ")) = NVL(rsItem("??Ŀ").Value)
            .Cell(flexcpData, .Row, .ColIndex("??Ŀ"), .Row, .Cols - 1) = NVL(rsItem("??Ŀ").Value)
            .Cell(flexcpText, .Row, .ColIndex("??Ŀ"), .Row, .Cols - 1) = NVL(rsItem("??Ŀ").Value)
           
           '??ȡָ??????
           Set rsResult = gclsPackage.get_PeisPersonResult(lngTaskKey, lngPersonKey, Val(NVL(rsItem("?嵥ID").Value)))
           If rsResult.BOF = False Then
                
                If .TextMatrix(.Rows - 1, .ColIndex("??Ŀ")) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = Color.ǳ??ɫ
                
                .TextMatrix(.Row, .ColIndex("????")) = 1
                .TextMatrix(.Row, .ColIndex("ָ??")) = "ָ??????"
                .TextMatrix(.Row, .ColIndex("????")) = "ָ??????"
                .TextMatrix(.Row, .ColIndex("??ʾ")) = "??ʾ"
                .TextMatrix(.Row, .ColIndex("?ο?")) = "?ο???Χ"
                
                Do While Not rsResult.EOF
                    
                    If Trim(.TextMatrix(.Rows - 1, .ColIndex("ָ??"))) <> "" Then .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .TextMatrix(.Row, .ColIndex("ָ??")) = NVL(rsResult("ָ??").Value)
                    .TextMatrix(.Row, .ColIndex("????")) = NVL(rsResult("????").Value)
                    .TextMatrix(.Row, .ColIndex("??ʾ")) = NVL(rsResult("??ʾ").Value)
                    .TextMatrix(.Row, .ColIndex("?ο?")) = NVL(rsResult("?ο?").Value)
                     Call ApplyResultColor(.Row, NVL(rsResult("??ʾ").Value))
                    rsResult.MoveNext
                Loop
                
           End If
           '??????Ŀ?????Ӽ??鱸ע???걾??̬
           If Val(NVL(rsItem("?ɼ???ʽid").Value)) > 0 Then
                 If Trim(.TextMatrix(.Rows - 1, .ColIndex("??Ŀ"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("ָ??"))) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .MergeRow(.Row) = True
                .TextMatrix(.Row, .ColIndex("ָ??")) = "???鱸ע"
                .TextMatrix(.Row, .ColIndex("????")) = NVL(rsItem("??ע˵??").Value)
                .Cell(flexcpData, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsItem("??ע˵??").Value)
                .Cell(flexcpText, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsItem("??ע˵??").Value)
                
                .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .MergeRow(.Row) = True
                .TextMatrix(.Row, .ColIndex("ָ??")) = "?걾??̬"
                .TextMatrix(.Row, .ColIndex("????")) = NVL(rsItem("?걾??̬").Value)
                .Cell(flexcpData, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsItem("?걾??̬").Value)
                .Cell(flexcpText, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsItem("?걾??̬").Value)
                
           End If
           
           '??ȡ??ĿС??
           Set rsConclusion = gclsPackage.Get_PeisPersonConclusion(1, lngTaskKey, lngPersonKey, Val(NVL(rsItem("?嵥id").Value)))
           
           Do While Not rsConclusion.EOF
                If Trim(.TextMatrix(.Rows - 1, .ColIndex("??Ŀ"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("ָ??"))) <> "" Then .Rows = .Rows + 1
                .Row = .Rows - 1
                
                .MergeRow(.Row) = True
                If rsConclusion.AbsolutePosition = 1 Then
                        
                    .TextMatrix(.Row, .ColIndex("ָ??")) = "??С?᡿"
                    .TextMatrix(.Row, .ColIndex("????")) = NVL(rsConclusion("????????").Value)
                    .Cell(flexcpData, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsConclusion("????????").Value)
                    .Cell(flexcpText, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsConclusion("????????").Value)
                Else
                    
                    .TextMatrix(.Row, .ColIndex("ָ??")) = ""
                    .TextMatrix(.Row, .ColIndex("????")) = NVL(rsConclusion("????????").Value)
                    .Cell(flexcpData, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsConclusion("????????").Value)
                    .Cell(flexcpText, .Row, .ColIndex("????"), .Row, .Cols - 1) = NVL(rsConclusion("????????").Value)
                End If
                rsConclusion.MoveNext
           Loop
           
            If Trim(.TextMatrix(.Rows - 1, .ColIndex("??Ŀ"))) <> "" Or Trim(.TextMatrix(.Rows - 1, .ColIndex("ָ??"))) <> "" Then .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, .ColIndex("??Ŀ")) = " "
            
           rsItem.MoveNext
        Loop
        .AutoSize 0, .ColIndex("????")
     End With
    Exit Sub
errHand:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Sub


'******************************************************************************************************************
'???ܣ?
'??????
'???أ?
'******************************************************************************************************************
Private Function ApplyResultColor(ByVal lngRow As Long, ByVal str???? As String) As Boolean
    Dim lngColor As Long
    Dim lngForeColor As Long
    Dim lngCol As Long
    Dim strSign As String
    
    If lngRow = 0 Then Exit Function
    
    strSign = str????
    Select Case str????
    Case "ƫ??"
        lngColor = Color.?ͱ걳??ɫ
        lngForeColor = Color.????ǰ??ɫ
        strSign = "??"
    Case "ƫ??"
        lngColor = Color.???걳??ɫ
        lngForeColor = Color.????ǰ??ɫ
        strSign = "??"
    Case "?쳣"
        lngColor = Color.???걳??ɫ
        lngForeColor = Color.????ǰ??ɫ
    Case "????????"
        lngColor = Color.????ƫ?߱???ɫ
        lngForeColor = Color.????ǰ??ɫ
    Case "????????"
        lngColor = Color.????ƫ?ͱ???ɫ
        lngForeColor = Color.????ǰ??ɫ
    Case "????????"
        lngColor = Color.????ƫ?߱???ɫ
        lngForeColor = Color.????ǰ??ɫ
    Case "????????"
        lngColor = Color.????ƫ?ͱ???ɫ
        lngForeColor = Color.????ǰ??ɫ
    Case Else
        lngColor = &H80000005
        lngForeColor = Color.Ĭ??ǰ??ɫ
    End Select
    
    lngCol = vsf.ColIndex("????")
    vsf.Cell(flexcpBackColor, lngRow, lngCol, lngRow, lngCol) = lngColor
    vsf.Cell(flexcpForeColor, lngRow, lngCol, lngRow, lngCol) = lngForeColor
    vsf.TextMatrix(lngRow, vsf.ColIndex("??ʾ")) = strSign
    
    ApplyResultColor = True
    
    
End Function


'******************************************************************************************************************
'???ܣ???ʼ??????
'??????
'???أ?
'******************************************************************************************************************
Private Sub InitVsf()
    
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsf, True, False)
        Call .ClearColumn
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, False, False, True)
        Call .AppendColumn("??Ŀ", 255, flexAlignLeftTop, flexDTString, , , True)
        Call .AppendColumn("ָ??", 2100, flexAlignLeftTop, flexDTString, , , True)
        Call .AppendColumn("????", 4030, flexAlignLeftTop, flexDTString, , , True)
        Call .AppendColumn("??ʾ", 450, flexAlignLeftTop, flexDTString, "", , True)
        Call .AppendColumn("????", 0, flexAlignLeftTop, flexDTString, "", , True, , , True)
        Call .AppendColumn("????", 0, flexAlignLeftTop, flexDTString, "", , True, , , True)
        Call .AppendColumn("?ο?", 900, flexAlignLeftTop, flexDTString, , , False)
        
        
        .AppendRows = False
        .AutoRowHeight = True
    End With
    vsf.RowHidden(0) = True

End Sub



'######################################################################################################################

Private Sub Form_Resize()

    On Error Resume Next
    
    picPane(0).Move 15, 15, picPane(0).Width, Me.ScaleHeight - 30
    
    picPane(1).Move picPane(0).Left + picPane(0).Width + 15, 15, Me.ScaleWidth - (picPane(0).Left + picPane(0).Width) - 30, Me.ScaleHeight - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsf = Nothing
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim lngRow As Long
    
    lngRow = mclsVsf.FindRow(Mid(Node.Key, 2), vsf.ColIndex("ID"))

    If lngRow > 0 Then
        vsf.TopRow = lngRow
    End If
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngSvrBkColor As Long
    Dim rc As RECT
    Dim rc1 As RECT
    Dim r1%, g1%, b1%
    Dim r2%, g2%, b2%
    Dim rg%, gg%, bg%
    Dim lngLoop As Long
    
    On Error Resume Next
    
    With vsf
        
        If Val(.TextMatrix(Row, .ColIndex("????"))) <> 1 Then Exit Sub

'        'flexODOver
'        '--------------------------------------------------------------------------------------------------------------
        rc.Left = Left
        rc.Top = Top
        rc.Right = Right
        rc.Bottom = Top + 1


        'Draw Frame
        '--------------------------------------------------------------------------------------------------------------
        lngSvrBkColor = SetBkColor(hDC, 0)

        Call ExtTextOut(hDC, rc.Left, rc.Top, ETO_OPAQUE, rc, " ", 1, lngLoop)
        Call InflateRect(rc, -1, -1)

'        Call SetBkColor(hDC, RGB(255, 255, 255))
        Call ExtTextOut(hDC, rc.Left, rc.Top, ETO_OPAQUE, rc, " ", 1, lngLoop)
        
    End With
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
        Case 0
            tvw.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        Case 1
            vsf.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

