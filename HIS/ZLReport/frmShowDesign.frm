VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowDesign 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   FillColor       =   &H80000012&
   Icon            =   "frmShowDesign.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmShowDesign.frx":000C
   ScaleHeight     =   4890
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   11
      Top             =   4560
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar scrVsc 
      Height          =   3765
      Left            =   5505
      TabIndex        =   10
      Top             =   450
      Width           =   250
      _ExtentX        =   450
      _ExtentY        =   6641
      _Version        =   393216
      LargeChange     =   20
      Max             =   100
      Orientation     =   1179648
   End
   Begin MSComCtl2.FlatScrollBar scrHsc 
      Height          =   250
      Left            =   30
      TabIndex        =   9
      Top             =   4230
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      LargeChange     =   20
      Max             =   100
      Orientation     =   1179649
   End
   Begin VB.PictureBox picFormat 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   5820
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "X����"
      Top             =   0
      Width           =   5820
      Begin MSComctlLib.ImageCombo cboFormat 
         Height          =   330
         Left            =   930
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "��������޸ĸ�ʽ����"
         Top             =   45
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "img16"
      End
      Begin MSComctlLib.Toolbar tbrScale 
         Height          =   660
         Left            =   5145
         TabIndex        =   12
         Top             =   45
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   1164
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "img16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Scale"
               Object.ToolTipText     =   "������ʾ�ı���"
               ImageKey        =   "Scale"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   11
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ԭʼ��С"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�ʺϿ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�ʺϸ߶�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ȫ����ʾ"
                     Text            =   "ȫ����ʾ"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "50%"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "75%"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "100%"
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "125%"
                  EndProperty
                  BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "150%"
                  EndProperty
                  BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "200%"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.Label LblFormat 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʽ"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   135
         TabIndex        =   8
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3780
      Left            =   30
      MouseIcon       =   "frmShowDesign.frx":015E
      ScaleHeight     =   3780
      ScaleWidth      =   5475
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   5475
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         ForeColor       =   &H00FF0000&
         Height          =   3315
         Left            =   180
         MouseIcon       =   "frmShowDesign.frx":02B0
         ScaleHeight     =   3315
         ScaleWidth      =   5055
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   165
         Width           =   5055
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
            Height          =   585
            Index           =   0
            Left            =   -8888
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   1032
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   15724527
            ForeColorFixed  =   0
            ForeColorSel    =   16777215
            BackColorBkg    =   16777215
            BackColorUnpopulated=   16777215
            GridColor       =   0
            GridColorFixed  =   0
            GridColorUnpopulated=   16777215
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   0
            HighLight       =   0
            GridLinesFixed  =   1
            GridLinesUnpopulated=   1
            ScrollBars      =   0
            MergeCells      =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Shape Shp 
            Height          =   1575
            Index           =   0
            Left            =   -60000
            Top             =   0
            Width           =   2040
         End
         Begin VB.Label lblshp 
            BackColor       =   &H8000000E&
            Height          =   735
            Index           =   0
            Left            =   -60000
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.Image Img 
            Appearance      =   0  'Flat
            Height          =   735
            Index           =   0
            Left            =   -8888
            Stretch         =   -1  'True
            Top             =   390
            Width           =   555
         End
         Begin VB.Label lbl 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFEFEF&
            Caption         =   "��ǩ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   -2205
            TabIndex        =   6
            Top             =   255
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label lblLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   0
            Left            =   -2235
            TabIndex        =   5
            Top             =   75
            Visible         =   0   'False
            Width           =   1410
         End
      End
      Begin VB.PictureBox picShadow 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3360
         Left            =   240
         ScaleHeight     =   3360
         ScaleWidth      =   5070
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   5070
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1965
      Top             =   1155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowDesign.frx":0402
            Key             =   "Report"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowDesign.frx":589C
            Key             =   "Scale"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmShowDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngRPTID As Long '�룺Ҫ��Ƶı���ID

Private Const M_Shadow_W = 75
Private mintCurID As Integer '��ǰѡ��ؼ�����(��1��ʼ)
Private mobjReport As Report 'Ҫ��Ƶı������

Private msngScale As Single
Private mbytCurrFmt As Byte 'ѡ��ı����ʽ

Private Sub CboFormat_Click()
    If Trim(cboFormat.Text) = "" Then Exit Sub
    If mbytCurrFmt <> Mid(cboFormat.SelectedItem.Key, 2) Then
        mbytCurrFmt = Mid(cboFormat.SelectedItem.Key, 2)
        Call ReFlashReport
    End If
End Sub

Private Sub Form_Load()
    msngScale = 1: mintCurID = 0: mbytCurrFmt = 0
    Set mobjReport = ReadReport(mlngRPTID)
    Call LoadReportFormat
    Call ReFlashReport
    Call Form_Resize
End Sub

Private Sub ReFlashReport()
'���ܣ�����ˢ����ʾ��������
'������blnReLoad=�Ƿ����´����ݿ��м�������
    Dim objTmp As Object, tmpReport As Report, intPreMax As Long
    
    For Each objTmp In lblLine
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In lbl
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In msh
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    For Each objTmp In img
        If objTmp.Index <> 0 Then Unload objTmp
    Next
    mintCurID = 0
    
    Call ShowSize
    Call ShowScroll
    Call ShowItems
    
    Refresh
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tbrScale.Left = Me.ScaleWidth - tbrScale.Width - 150
    cboFormat.Width = Me.ScaleWidth - cboFormat.Left - tbrScale.Width - 200
    
    picBack.Left = 0
    picBack.Top = picFormat.Height
    picBack.Width = Me.ScaleWidth - scrVsc.Width
    picBack.Height = Me.ScaleHeight - picFormat.Height - sta.Height - scrHsc.Height
    
    scrHsc.Left = 0
    scrHsc.Top = picBack.Top + picBack.Height
    scrHsc.Width = picBack.Width
    
    scrVsc.Top = picBack.Top
    scrVsc.Left = picBack.Left + picBack.Width
    scrVsc.Height = picBack.Height
    
    Call ShowSize
    Call ShowScroll
    
    Me.Refresh
End Sub

Private Sub scrhsc_Change()
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrHsc.Value = 0 Then Call ShowScroll(1)
End Sub

Private Sub scrhsc_Scroll()
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrHsc.Value = 0 Then Call ShowScroll(1)
End Sub

Private Sub scrVsc_Change()
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrVsc.Value = 0 Then Call ShowScroll(2)
End Sub

Private Sub scrVsc_Scroll()
    Call ShowSize(-scrVsc.Value * 15#, -scrHsc.Value * 15#)
    If scrVsc.Value = 0 Then Call ShowScroll(2)
End Sub

Private Sub ShowSize(Optional lngTop As Single = 0, Optional lngLeft As Single = 0)
'����:��ʾ����ֽ�Ŵ�С
    Dim lngW As Long, lngH As Long
    Dim objFmt As RPTFmt
    
    Set objFmt = mobjReport.Fmts("_" & mbytCurrFmt)
    
    '��ӡ��ֽ��ֻ�Ǽ򵥵ؽ�ֽ�ſ�Ⱥ͸߶ȶԵ�
    If objFmt.ֽ�� = 1 Then
        lngW = objFmt.W: lngH = objFmt.H
    Else
        lngH = objFmt.W: lngW = objFmt.H
    End If
    
    picPaper.Width = Format(lngW * msngScale, "0.00")
    picPaper.Height = Format(lngH * msngScale, "0.00")
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    
    If picBack.Width > picPaper.Width + M_Shadow_W * 2 Then
        picPaper.Left = (picBack.Width - picPaper.Width - M_Shadow_W * 2) / 2
    Else
        picPaper.Left = lngLeft
    End If
    If picBack.Height > picPaper.Height + M_Shadow_W * 2 Then
        picPaper.Top = (picBack.Height - picPaper.Height - M_Shadow_W * 2) / 2
    Else
        picPaper.Top = lngTop
    End If
    
    picShadow.Top = picPaper.Top + M_Shadow_W
    picShadow.Left = picPaper.Left + M_Shadow_W
    
    sta.SimpleText = "��ӡ��:" & mobjReport.��ӡ�� & "   ֽ��:" & GetPaperName(objFmt.ֽ��, objFmt.W, objFmt.H) & " " & _
        IIF(objFmt.ֽ�� = 256, CInt(objFmt.W / Twip_mm) & "mm �� " & CInt(objFmt.H / Twip_mm) & "mm", "") & _
        IIF(objFmt.ֽ�� = 1, "   ����", "   ����")
    
    Call Refresh
End Sub

Private Sub ShowScroll(Optional bytType As Byte = 3)
'����:���ù�����
'����:bytType=3-���߶�����(ȱʡֵ),1-������Hsc,2-������Vsc
    
    If bytType = 3 Or bytType = 2 Then
        If picBack.ScaleHeight >= picPaper.Height + M_Shadow_W * 2 Then
            scrVsc.Enabled = False
        Else
            scrVsc.Max = (picPaper.Height + M_Shadow_W * 2 - picBack.ScaleHeight) / Screen.TwipsPerPixelX 'ת��Ϊ����Ϊ��λ
            Call ShowSize(0, picPaper.Left)
            scrVsc.Value = 0
            scrVsc.Enabled = True
        End If
    End If
    If bytType = 3 Or bytType = 1 Then
        If picBack.ScaleWidth >= picPaper.Width + M_Shadow_W * 2 Then
            scrHsc.Enabled = False
        Else
            scrHsc.Max = (picPaper.Width + M_Shadow_W * 2 - picBack.ScaleWidth) / Screen.TwipsPerPixelX
            Call ShowSize(picPaper.Top, 0)
            scrHsc.Value = 0
            scrHsc.Enabled = True
        End If
    End If
End Sub

Private Sub SetGridSame(mshS As Control, mshO As Control)
'����:������������ؼ�������ͬ
'˵��������ʱ����������������
    Dim i As Integer, j As Integer
    
    mshO.Redraw = False
    mshS.Redraw = False
    
    mshO.Width = mshS.Width
    mshO.Height = mshS.Height
    mshO.Rows = mshS.Rows
    mshO.Cols = mshS.Cols
    mshO.FixedCols = mshS.FixedCols
    mshO.FixedRows = mshS.FixedRows
    
    mshO.ForeColor = mshS.ForeColor
    mshO.BackColor = mshS.BackColor
    mshO.BackColorFixed = mshS.BackColorFixed
    mshO.ForeColorFixed = mshS.ForeColorFixed
    mshO.BackColorSel = mshS.BackColorSel
    mshO.ForeColorSel = mshS.ForeColorSel
    mshO.GridColor = mshS.GridColor
    mshO.GridColorFixed = mshS.GridColorFixed
    
    mshO.Font.Size = mshS.Font.Size
    mshO.Font.name = mshS.Font.name
    mshO.Font.Bold = mshS.Font.Bold
    mshO.Font.Underline = mshS.Font.Underline
    mshO.Font.Italic = mshS.Font.Italic
    
    For i = 0 To mshS.Rows - 1
        mshS.Row = i: mshO.Row = i
        mshO.RowHeight(i) = mshS.RowHeight(i)
        mshO.MergeRow(i) = mshS.MergeRow(i)
        For j = 0 To mshS.Cols - 1
            mshS.Col = j: mshO.Col = j
            mshO.CellAlignment = mshS.CellAlignment
            mshO.CellFontBold = mshS.CellFontBold
            mshO.CellFontName = mshS.CellFontName
            mshO.CellFontSize = mshS.CellFontSize
            mshO.CellFontItalic = mshS.CellFontItalic
            mshO.CellFontUnderline = mshS.CellFontUnderline
            mshO.TextMatrix(i, j) = mshS.TextMatrix(i, j)
            If i <= mshS.FixedRows - 1 Or j <= mshS.FixedCols - 1 Then
                mshO.CellBackColor = mshS.BackColorFixed
                mshO.CellForeColor = mshS.ForeColorFixed
            Else
                mshO.CellBackColor = mshS.BackColor
                mshO.CellForeColor = mshS.ForeColor
            End If
        Next
    Next
    For i = 0 To mshS.Cols - 1
        mshO.ColWidth(i) = mshS.ColWidth(i)
        mshO.ColAlignment(i) = mshS.ColAlignment(i)
        mshO.MergeCol(i) = mshS.MergeCol(i)
    Next
    
    mshO.Redraw = True
    mshS.Redraw = True
End Sub

Private Sub SetGridLine(idx As Integer)
'���ܣ�����ָ���������������,�����и����,�������������
'˵��������ʱ�ؼ���Ӧ�����ݶ���(Item)�����Ѿ����ڣ��Ҷ�Ӧ�ؼ��Ѿ�����������ͷ���
    Dim blnPre As Boolean, SinH As Single
    Dim X As Integer, Y As Integer, Z As Integer
    Dim tmpID As RelatID, i As Integer, j As Integer

    blnPre = msh(idx).Redraw
    msh(idx).Redraw = False
    
    If mobjReport.Items("_" & idx).���� = 4 Then '���ܱ��
        '�������������������
        If mobjReport.Ʊ�� Then
            SinH = 0: X = msh(idx).FixedRows
            For i = 0 To msh(idx).FixedRows - 1
                SinH = SinH + msh(idx).RowHeight(i)
            Next
            msh(idx).Rows = Abs(Int((-(msh(idx).Height - SinH)) / (mobjReport.Items("_" & idx).�и� * msngScale))) + X
            If msh(idx).Rows = X Then msh(idx).Rows = msh(idx).Rows + 2
            msh(idx).FixedRows = X
            For i = msh(idx).FixedRows To msh(idx).Rows - 1
                msh(idx).RowHeight(i) = mobjReport.Items("_" & idx).�и� * msngScale
            Next
        End If
    ElseIf mobjReport.Items("_" & idx).���� = 5 Then '���ܱ��
        X = msh(idx).FixedCols '���������Ŀ��
        Y = msh(idx).FixedRows - 1 '���������Ŀ��
        For Each tmpID In mobjReport.Items("_" & idx).SubIDs
            If mobjReport.Items("_" & tmpID.ID).���� = 9 Then Z = Z + 1 'ͳ����Ŀ��
        Next

        '���ܱ���������������
        msh(idx).Rows = Abs(Int(-msh(idx).Height / msh(idx).RowHeight(0))) + 1
        If msh(idx).Rows < msh(idx).FixedRows + 3 Then msh(idx).Rows = msh(idx).FixedRows + 3
        
        For i = msh(idx).FixedRows + 1 To msh(idx).Rows - 1
            msh(idx).RowHeight(i) = msh(idx).RowHeight(0)
            For j = 0 To msh(idx).FixedCols - 1
                msh(idx).TextMatrix(i, j) = msh(idx).TextMatrix(msh(idx).FixedRows + 2, j)
            Next
        Next
        
        '����к�����ࣺ���ܱ���������ͳ����
        If msh(idx).FixedRows > 1 Then
            X = 0
            For i = 0 To msh(idx).FixedCols - 1
                X = X + msh(idx).ColWidth(i) '��������ܿ��
            Next
            Y = 0
            For i = msh(idx).FixedCols To msh(idx).FixedCols + Z - 1
                Y = Y + msh(idx).ColWidth(i) 'һ��ͳ�����ܿ��
            Next
            '���� = ͳ������ * ÿ������ + �����������
            msh(idx).Cols = Abs(Int(-(msh(idx).Width - X) / Y)) * Z + msh(idx).FixedCols
            'ÿ���ȼ�������ͬ
            For i = msh(idx).FixedCols + Z To msh(idx).Cols - 1
                For j = 0 To msh(idx).FixedRows - 2
                    msh(idx).TextMatrix(j, i) = msh(idx).TextMatrix(j, msh(idx).FixedCols + 1)
                Next
            Next
            For i = msh(idx).FixedCols + Z To msh(idx).Cols - 1 Step Z
                For j = 1 To Z
                    msh(idx).TextMatrix(msh(idx).FixedRows - 1, i + j - 1) = _
                    msh(idx).TextMatrix(msh(idx).FixedRows - 1, msh(idx).FixedCols + j - 1)
                    msh(idx).ColWidth(i + j - 1) = msh(idx).ColWidth(msh(idx).FixedCols + j - 1)
                    msh(idx).ColAlignment(i + j - 1) = msh(idx).ColAlignment(msh(idx).FixedCols + j - 1)
                Next
            Next
        End If
    End If
    
    msh(idx).Redraw = blnPre
End Sub

Private Sub ShowItems()
    '���ܣ�����mobjReport������ʾ����Ԫ��
    Dim tmpItem As RPTItem, bytFormat As Byte
    
    For Each tmpItem In mobjReport.Items
        If tmpItem.��ʽ�� = mbytCurrFmt Then Call ShowItem(tmpItem.ID)
    Next
End Sub

Private Sub LoadReportFormat()
    Dim tmpFmt As RPTFmt
    
    With cboFormat
        .ComboItems.Clear
        For Each tmpFmt In mobjReport.Fmts
            .ComboItems.Add , "_" & tmpFmt.���, tmpFmt.˵��, 1
        Next
        .ComboItems(1).Selected = True
        Set .SelectedItem = .ComboItems(1)
        mbytCurrFmt = Mid(cboFormat.SelectedItem.Key, 2)
    End With
End Sub

Private Sub ShowItem(idx As Integer)
'���ܣ���ʾָ���ı���Ԫ��(ShowItems���Ӻ���,Ҳ�ɵ�������)
'������idx=mobjReport�е�Ԫ������
    Dim i As Integer, j As Integer, tmpID As RelatID, ObjSel As Control
    
    With mobjReport.Items("_" & idx)
        Select Case .����
            Case 1 '����
                Load lblLine(.ID)
                Set ObjSel = lblLine(.ID)
                ObjSel.Top = Format(.Y * msngScale, "0.00")
                ObjSel.Left = Format(.X * msngScale, "0.00")
                ObjSel.Height = Format(.H * msngScale, "0.00")
                ObjSel.Width = Format(.W * msngScale, "0.00")
                ObjSel.BackColor = .ǰ��
                If .���� Then ObjSel.BorderWidth = 2
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 2, 3 '��ǩ
                Load lbl(.ID)
                Set ObjSel = lbl(.ID)
                ObjSel.Top = Format(.Y * msngScale, "0.00")
                ObjSel.Left = Format(.X * msngScale, "0.00")
                ObjSel.Height = Format(.H * msngScale, "0.00")
                ObjSel.Width = Format(.W * msngScale, "0.00")
                ObjSel.ForeColor = .ǰ��
                ObjSel.BackColor = IIF(.���� = &HFFFFFF, lbl(0).BackColor, .����)
                ObjSel.Font.name = .����
                ObjSel.Font.Size = Format(.�ֺ� * msngScale, "0.0")
                ObjSel.Font.Bold = .����
                ObjSel.Font.Italic = .б��
                ObjSel.Font.Underline = .����
                ObjSel.BorderStyle = IIF(.�߿�, 1, 0)
                ObjSel.Alignment = IIF(.���� <> 0, IIF(.���� = 1, 2, 1), 0)
                ObjSel.Caption = .����
                ObjSel.AutoSize = .�Ե�
                If InStr(1, "|11,", "|" & .���� & ",") <> 0 Then
                    ObjSel.BorderStyle = 1
                    ObjSel.BackStyle = 0
                    If .���� = 10 Then ObjSel.Caption = ""
                End If
                ObjSel.ZOrder 0
                ObjSel.Visible = True
            Case 10 '����
                Load Shp(.ID)
                Set ObjSel = Shp(.ID)
                Load lblshp(.ID)
                lblshp(.ID).BackColor = picPaper.BackColor
                ObjSel.Top = Format(.Y * msngScale, "0.00")
                ObjSel.Left = Format(.X * msngScale, "0.00")
                ObjSel.Height = Format(.H * msngScale, "0.00")
                ObjSel.Width = Format(.W * msngScale, "0.00")
                lblshp(.ID).Top = ObjSel.Top
                lblshp(.ID).Left = ObjSel.Left
                lblshp(.ID).Width = ObjSel.Width
                lblshp(.ID).Height = ObjSel.Height
                ObjSel.BorderColor = .ǰ��
                ObjSel.BackColor = IIF(.���� = &HFFFFFF, Shp(0).BackColor, .����)
                ObjSel.BorderStyle = 1
                ObjSel.BackStyle = 0
                If .���� Then ObjSel.BorderWidth = 2
                
                ObjSel.ZOrder 1
                ObjSel.Visible = True
                lblshp(.ID).ZOrder 1
                lblshp(.ID).Visible = True
            Case 4, 5 '������,���ܱ��
                Load msh(.ID)
                Set ObjSel = msh(.ID)
                '��ʽ����
                ObjSel.Top = Format(.Y * msngScale, "0.00")
                ObjSel.Left = Format(.X * msngScale, "0.00")
                ObjSel.Height = Format(.H * msngScale, "0.00")
                ObjSel.Width = Format(.W * msngScale, "0.00")
                ObjSel.Font.Size = Format(.�ֺ� * msngScale, "0.0")
                
                '��������(����CopyIDs�Ѿ�����)
                i = 0
                For Each tmpID In .CopyIDs
                    i = i + 1
                    Load msh(tmpID.ID)
                    msh(tmpID.ID).Width = ObjSel.Width
                    msh(tmpID.ID).Height = ObjSel.Height
                    msh(tmpID.ID).Top = ObjSel.Top
                    msh(tmpID.ID).Left = ObjSel.Left + (ObjSel.Width - 15) * i
                    msh(tmpID.ID).Font.Size = ObjSel.Font.Size
                    msh(tmpID.ID).Tag = "C_" & .ID
                    msh(tmpID.ID).ZOrder
                    msh(tmpID.ID).Visible = True
                Next
                
                Call ReShowGrid(.ID)
                
                ObjSel.ZOrder
                ObjSel.Visible = True
            Case 11
                Load img(.ID)
                Set ObjSel = img(.ID)
                ObjSel.Picture = .ͼƬ
                ObjSel.Top = Format(.Y * msngScale, "0.00")
                ObjSel.Left = Format(.X * msngScale, "0.00")
                ObjSel.Height = Format(.H * msngScale, "0.00")
                ObjSel.Width = Format(.W * msngScale, "0.00")
                ObjSel.BorderStyle = IIF(.�߿�, 1, 0)
                ObjSel.ZOrder
                ObjSel.Visible = True
        End Select
    End With
End Sub

Private Sub ReShowGrid(idx As Integer)
'���ܣ�����mobjReport���������»��Ʊ������,��ʱˢ�·����ؼ�
'˵����1.mobjReport���������Ѵ���,2.��Ӧ�ؼ��Ѵ���

    Dim i As Integer, j As Integer, X As Integer, Y As Integer, Z As Integer
    Dim tmpID As RelatID, tmpItem As RPTItem, strCaption As String, sgnH As Long
    
    msh(idx).Redraw = False
    msh(idx).Clear
    With mobjReport.Items("_" & idx)
        If .���� = 4 Then '������
            '��ʽ����(λ�ü��ߴ粻��)
            msh(idx).ForeColor = .ǰ��
            msh(idx).ForeColorFixed = .ǰ��
            msh(idx).GridColor = .����
            msh(idx).GridColorFixed = IIF(.��ʽ = "", .����, Val(.��ʽ))
            
            msh(idx).BackColor = .����
            msh(idx).BackColorFixed = IIF(.���� = &HFFFFFF, lbl(0).BackColor, .����)
            
            msh(idx).Font.name = .����
            msh(idx).Font.Size = Format(.�ֺ� * msngScale, "0.0")
            msh(idx).Font.Bold = .����
            msh(idx).Font.Italic = .б��
            msh(idx).Font.Underline = .����
            
            '��������
            '����
            msh(idx).Cols = .SubIDs.count
            msh(idx).FixedCols = 0
            i = 0
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                
                If i = 0 Then '��С����
                    If mobjReport.Ʊ�� = False Then
                        msh(idx).Rows = UBound(Split(tmpItem.��ͷ, "|")) + 3
                        msh(idx).FixedRows = UBound(Split(tmpItem.��ͷ, "|")) + 1
                    Else
                        msh(idx).Rows = UBound(Split(tmpItem.��ͷ, "|")) + 3
                        msh(idx).FixedRows = UBound(Split(tmpItem.��ͷ, "|")) + 1
                    End If
                End If

                '����������
                msh(idx).ColWidth(tmpItem.���) = tmpItem.W * msngScale
                msh(idx).ColAlignment(tmpItem.���) = Switch(tmpItem.���� = 0, 1, tmpItem.���� = 1, 4, tmpItem.���� = 2, 7)
                msh(idx).TextMatrix(msh(idx).FixedRows, tmpItem.���) = tmpItem.����
                msh(idx).TextMatrix(msh(idx).FixedRows + 1, tmpItem.���) = tmpItem.����
                                    
                '�Զ����ͷ����
                For i = 0 To msh(idx).FixedRows - 1
                    On Error Resume Next
                    
                    Err = 0
                    strCaption = Split(Split(tmpItem.��ͷ, "|")(i), "^")(2)
                    If Err <> 0 Then strCaption = ""
                    If strCaption = "#" Then
                        msh(idx).TextMatrix(i, tmpItem.���) = ""
                    ElseIf strCaption = "��" Then
                        msh(idx).TextMatrix(i, tmpItem.���) = msh(idx).TextMatrix(i, tmpItem.��� - 1)
                    ElseIf strCaption = "��" Then
                        msh(idx).TextMatrix(i, tmpItem.���) = msh(idx).TextMatrix(i - 1, tmpItem.���)
                    Else
                        msh(idx).TextMatrix(i, tmpItem.���) = strCaption
                    End If
                    
                    Err = 0
                    sgnH = Split(Split(tmpItem.��ͷ, "|")(i), "^")(1)
                    If Err <> 0 Then sgnH = 250
                    msh(idx).RowHeight(i) = sgnH * msngScale
                    msh(idx).Row = i
                    msh(idx).Col = tmpItem.���
                    Err = 0
                    sgnH = Split(Split(tmpItem.��ͷ, "|")(i), "^")(0)
                    If Err <> 0 Then sgnH = 4
                    msh(idx).CellAlignment = sgnH
                Next
            Next
            
            For i = msh(idx).FixedRows To msh(idx).Rows - 1
                msh(idx).RowHeight(i) = .�и� * msngScale
            Next
            '�ϲ�����
            For i = 0 To msh(idx).FixedRows - 1
                msh(idx).MergeRow(i) = True
            Next
            For i = 0 To msh(idx).Cols - 1
                msh(idx).MergeCol(i) = True
            Next
            
            Call SetGridLine(.ID) '�������
            
            '��������(����CopyIDs�Ѿ�����)
            For Each tmpID In .CopyIDs
                Call SetGridSame(msh(idx), msh(tmpID.ID))
            Next
        ElseIf .���� = 5 Then '���ܱ��
            msh(idx).ForeColor = .ǰ��
            msh(idx).ForeColorFixed = .ǰ��
            msh(idx).GridColor = .����
            msh(idx).GridColorFixed = .����
            
            msh(idx).BackColor = .����
            msh(idx).BackColorFixed = IIF(.���� = &HFFFFFF, lbl(0).BackColor, .����)
            
            msh(idx).Font.name = .����
            msh(idx).Font.Size = Format(.�ֺ� * msngScale, "0.0")
            msh(idx).Font.Bold = .����
            msh(idx).Font.Italic = .б��
            msh(idx).Font.Underline = .����
            
            X = 0: Y = 0: Z = 0
            For Each tmpID In .SubIDs
                Select Case mobjReport.Items("_" & tmpID.ID).����
                    Case 7
                        X = X + 1 '���������
                    Case 8
                        Y = Y + 1 '���������
                    Case 9
                        Z = Z + 1 'ͳ������
                End Select
            Next
            '��С������
            msh(idx).Rows = Y + 4
            msh(idx).FixedRows = Y + 1
            If Y = 0 Then
                msh(idx).Cols = X + Z
            Else
                msh(idx).Cols = X + IIF(Z = 1, Z + 1, Z)
            End If
            msh(idx).FixedCols = X
            msh(idx).RowHeight(0) = .�и� * msngScale '�и�0�Ǳ�׼
            msh(idx).RowHeightMin = msh(idx).RowHeight(0)
            
            '������������
            For Each tmpID In .SubIDs
                Set tmpItem = mobjReport.Items("_" & tmpID.ID)
                Select Case tmpItem.����
                    Case 7 '�������
                        msh(idx).TextMatrix(msh(idx).FixedRows - 1, tmpItem.���) = "[" & tmpItem.���� & "]"
                        
                        For i = msh(idx).FixedRows To msh(idx).Rows - 1
                            msh(idx).TextMatrix(i, tmpItem.���) = tmpItem.����
                        Next
                        If tmpItem.���� <> "" Then
                            msh(idx).TextMatrix(msh(idx).FixedRows, tmpItem.���) = tmpItem.����
                        End If
                        
                        msh(idx).ColWidth(tmpItem.���) = tmpItem.W * msngScale
                        msh(idx).ColAlignment(tmpItem.���) = Switch(tmpItem.���� = 0, 1, tmpItem.���� = 1, 4, tmpItem.���� = 2, 7)
                    Case 8 '�������
                        For i = 0 To msh(idx).FixedCols - 1
                            msh(idx).TextMatrix(tmpItem.���, i) = "[" & tmpItem.���� & "]"
                        Next
                        
                        For i = msh(idx).FixedCols To msh(idx).Cols - 1
                            msh(idx).TextMatrix(tmpItem.���, i) = tmpItem.����
                        Next
                        If tmpItem.���� <> "" Then
                            msh(idx).TextMatrix(tmpItem.���, msh(idx).FixedCols) = tmpItem.����
                        End If
                    Case 9 'ͳ����
                        msh(idx).TextMatrix(msh(idx).FixedRows - 1, msh(idx).FixedCols + tmpItem.���) = "[" & tmpItem.���� & "]"
                        msh(idx).ColWidth(msh(idx).FixedCols + tmpItem.���) = tmpItem.W * msngScale
                        msh(idx).ColAlignment(msh(idx).FixedCols + tmpItem.���) = Switch(tmpItem.���� = 0, 1, tmpItem.���� = 1, 4, tmpItem.���� = 2, 7)
                End Select
            Next
            
            '�ϲ�����
            For i = 0 To msh(idx).FixedRows - 2
                msh(idx).MergeRow(i) = True
            Next
            For i = 0 To msh(idx).FixedCols - 1
                msh(idx).MergeCol(i) = True
            Next
            
            Call SetGridLine(.ID)
        End If
    End With
    msh(idx).Redraw = True
End Sub

Private Sub tbrScale_ButtonClick(ByVal Button As MSComctlLib.Button)
    tbrScale_ButtonMenuClick tbrScale.Buttons("Scale").ButtonMenus("ȫ����ʾ")
End Sub

Private Sub tbrScale_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim objFmt As RPTFmt
    
    Set objFmt = mobjReport.Fmts("_" & mbytCurrFmt)
    
    Select Case ButtonMenu.Text
        Case "ԭʼ��С"
            msngScale = 1
        Case "�ʺϿ��"
            msngScale = picBack.ScaleWidth / (objFmt.W + M_Shadow_W * 2)
        Case "�ʺϸ߶�"
            msngScale = picBack.ScaleHeight / (objFmt.H + M_Shadow_W * 2)
        Case "ȫ����ʾ"
            If picBack.ScaleWidth / (objFmt.W + M_Shadow_W * 2) < _
                picBack.ScaleHeight / (objFmt.H + M_Shadow_W * 2) Then
                msngScale = picBack.ScaleWidth / (objFmt.W + M_Shadow_W * 2)
            Else
                msngScale = picBack.ScaleHeight / (objFmt.H + M_Shadow_W * 2)
            End If
        Case Else
            msngScale = Val(ButtonMenu.Text) / 100
    End Select
    If msngScale = 0 Then msngScale = 1
    
    Call ReFlashReport
End Sub
