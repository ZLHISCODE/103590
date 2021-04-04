VERSION 5.00
Object = "{7D52C334-5021-43A4-8EB4-86CC21862ABF}#1.2#0"; "zlTable.ocx"
Begin VB.Form frmTablePicCreator 
   BorderStyle     =   0  'None
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin zlRichEPR.ucPacsImgCanvas ucPacsImgCanvas1 
      Height          =   870
      Left            =   1845
      TabIndex        =   2
      Top             =   1035
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1535
   End
   Begin VB.PictureBox picBuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1620
      ScaleHeight     =   600
      ScaleWidth      =   690
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   690
   End
   Begin zlTable.Table Table1 
      Height          =   1095
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1931
      SingleLine      =   0   'False
   End
End
Attribute VB_Name = "frmTablePicCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ZoomPicture As Boolean           '是否进行图片缩放

'################################################################################################################
'## 功能：  获取表格替代图片
'##
'## 参数：  oTable              :当前表格对象
'##
'## 返回：  返回表格替代图片 stdPicture
'################################################################################################################
Public Function GetFinalPic(Optional ByVal oTable As cEPRTable, Optional ByVal BorderLine As Boolean) As StdPicture
    If oTable.TableType = tte_报告图片组 Then
        '缩放表格
        Dim ZoomFactor As Double
        If ZoomPicture Then
            ZoomFactor = oTable.GetFitZoomFactor
            ucPacsImgCanvas1.ZoomFactor = ZoomFactor
            oTable.Width = oTable.Width * ZoomFactor
            oTable.Height = oTable.Height * ZoomFactor
        End If
        
        ucPacsImgCanvas1.ReadPicturesFromTable oTable

        Set GetFinalPic = ucPacsImgCanvas1.FinalPic(BorderLine)
        
        '恢复尺寸
        If ZoomPicture Then
            oTable.Width = oTable.Width / ZoomFactor
            oTable.Height = oTable.Height / ZoomFactor
        End If
        Exit Function
    End If
    On Error Resume Next
    Dim i As Long, lngRows As Long, lngCols As Long, lngW As Long, lngH As Long
    Dim LL As Long, lT As Long, bHidden As Integer
    Dim lOffsetX As Long, lOffsetY As Long
    
    If Not oTable Is Nothing Then Me.ReadTable oTable, Me.Table1
    picBuff.Cls
    picBuff.Width = Table1.Width
    picBuff.Height = Table1.Height
    Table1.DrawToDC picBuff.hdc

    picBuff.Picture = picBuff.Image
    picBuff.Refresh
    Set GetFinalPic = picBuff.Picture
End Function

Public Sub ReadTable(oTable As cEPRTable, theTable As Table)
    On Error Resume Next
    '读取 cEPRTable 内容到 Table 控件中（包括Pictures和Elements）
    Dim i As Long, j As Long, strMerge As String, R1 As Long, R2 As Long, C1 As Long, C2 As Long
    Dim T As Variant, strColWidth As String, lKey As Long
    
    theTable.Tag = "Loading"
    theTable.Redraw = False
    theTable.SingleClickEdit = False
    theTable.HighlightMode = HMFilledRectAlpha
    theTable.BorderWidth = oTable.BorderWidth
    theTable.AutoHeight = oTable.AutoHeight
    theTable.Init oTable.Rows, oTable.Cols
    theTable.ExtendTag = oTable.ExtendTag
    strColWidth = oTable.ColWidthString
    T = Split(strColWidth, "|")
    On Error Resume Next
    If UBound(T) = -1 Then
        If oTable.Rows > 0 Then
            For i = 1 To oTable.Cols
                theTable.ColWidth(i) = oTable.Cell(1, i).Width
            Next
        End If
    Else
        For i = 0 To UBound(T)
            theTable.ColWidth(i + 1) = Val(T(i))
        Next
    End If
    
    For i = 1 To oTable.Rows
        theTable.ROWHEIGHT(i) = oTable.Cell(i, 1).Height
        For j = 1 To oTable.Cols
            lKey = theTable.CellKey(i, j)
            With oTable.Cell(i, j)
                theTable.Cells("K" & lKey).Text = oTable.Cell(i, j).内容文本
                theTable.Cells("K" & lKey).Margin = .Margin
                theTable.Cells("K" & lKey).Width = .Width
                theTable.Cells("K" & lKey).Height = .Height
'                theTable.Cells("K" & lkey).MergeInfo = .MergeNo
                theTable.Cells("K" & lKey).SingleLine = .SingleLine
                theTable.Cells("K" & lKey).ForeColor = .ForeColor
                theTable.Cells("K" & lKey).BackColor = .BackColor
                theTable.Cells("K" & lKey).GridLineColor = .GridLineColor
                theTable.Cells("K" & lKey).GridLineWidth = .GridLineWidth
                theTable.Cells("K" & lKey).FixedWidth = .FixedWidth
                theTable.Cells("K" & lKey).AutoHeight = .AutoHeight
                theTable.Cells("K" & lKey).FontName = .FontName
                theTable.Cells("K" & lKey).FontSize = .FontSize
                theTable.Cells("K" & lKey).FontBold = .FontBold
                theTable.Cells("K" & lKey).FontItalic = .FontItalic
                theTable.Cells("K" & lKey).FontStrikeout = .FontStrikeout
                theTable.Cells("K" & lKey).FontUnderline = .FontUnderline
                theTable.Cells("K" & lKey).FontWeight = .FontWeight
                theTable.Cells("K" & lKey).FormatString = .FormatString
                theTable.Cells("K" & lKey).Indent = .Indent
                theTable.Cells("K" & lKey).HAlignment = .HAlignment
                theTable.Cells("K" & lKey).VAlignment = .VAlignment
                theTable.Cells("K" & lKey).Protected = .Protected
                If oTable.Cell(i, j).ElementKey > 0 Then
                    theTable.Cells("K" & lKey).ToolTipText = oTable.Elements("K" & oTable.Cell(i, j).ElementKey).要素名称
                    theTable.Cells("K" & lKey).Tag = oTable.Cell(i, j).ElementKey
                End If
                If .PictureKey > 0 Then
                    oTable.Pictures("K" & .PictureKey).Row = i
                    oTable.Pictures("K" & .PictureKey).Col = j
                    Set theTable.Cells("K" & lKey).Picture = oTable.Pictures("K" & .PictureKey).DrawFinalPic(oTable)
                    theTable.Cells("K" & lKey).Tag = oTable.Cell(i, j).PictureKey
                End If
            End With
        Next
    Next
    
    For i = 1 To oTable.Cells.Count
        strMerge = oTable.Cells(i).MergeNo              '恢复单元格的合并
        If strMerge <> "" Then
            R1 = Val(Left(strMerge, 4))
            C1 = Val(Mid(strMerge, 5, 4))
            R2 = Val(Mid(strMerge, 9, 4))
            C2 = Val(Mid(strMerge, 13))
            theTable.MergeCells R1, C1, R2, C2, False
        End If
    Next
    
    theTable.ShowToolTipText = True
    theTable.MinRowHeight = 300
    theTable.Redraw = True
    theTable.Refresh
    theTable.Tag = ""

    theTable.FixCellsWidth
    If (Not theTable.AutoHeight) Then theTable.Height = oTable.Height
End Sub
