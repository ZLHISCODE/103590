VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSerialLayoutSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "版面设计"
   ClientHeight    =   5940
   ClientLeft      =   30
   ClientTop       =   570
   ClientWidth     =   7590
   Icon            =   "frmSerset.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame famImage 
      Caption         =   "图像"
      Height          =   4884
      Left            =   3996
      TabIndex        =   11
      Top             =   300
      Width           =   3288
      Begin VB.Frame Frame4 
         Caption         =   "自定义"
         Height          =   1575
         Left            =   330
         TabIndex        =   22
         Top             =   3120
         Width           =   2655
         Begin VB.CommandButton CmdImgApply 
            Caption         =   "应用"
            Height          =   350
            Left            =   840
            TabIndex        =   35
            Top             =   1020
            Width           =   1100
         End
         Begin MSComCtl2.UpDown UDImgCol 
            Height          =   255
            Left            =   2010
            TabIndex        =   33
            Top             =   540
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UDImgRow 
            Height          =   255
            Left            =   960
            TabIndex        =   36
            Top             =   540
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox TxtImgRow 
            Height          =   315
            Left            =   480
            TabIndex        =   30
            Text            =   "1"
            Top             =   510
            Width           =   735
         End
         Begin VB.TextBox TxtImgCol 
            Height          =   315
            Left            =   1530
            TabIndex        =   32
            Text            =   "1"
            Top             =   510
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "列:"
            Height          =   180
            Left            =   1530
            TabIndex        =   34
            Top             =   300
            Width           =   270
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "行:"
            Height          =   180
            Left            =   480
            TabIndex        =   31
            Top             =   300
            Width           =   270
         End
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   11
         Left            =   300
         Picture         =   "frmSerset.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   276
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   12
         Left            =   1200
         Picture         =   "frmSerset.frx":277C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   276
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   13
         Left            =   2100
         Picture         =   "frmSerset.frx":422E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   276
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   14
         Left            =   300
         Picture         =   "frmSerset.frx":5CE0
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   15
         Left            =   1200
         Picture         =   "frmSerset.frx":7792
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   16
         Left            =   2100
         Picture         =   "frmSerset.frx":9244
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   17
         Left            =   300
         Picture         =   "frmSerset.frx":ACF6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2076
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   18
         Left            =   1200
         Picture         =   "frmSerset.frx":C7A8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2076
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   19
         Left            =   2100
         Picture         =   "frmSerset.frx":E25A
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2076
         Width           =   900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3240
      TabIndex        =   10
      Top             =   5376
      Width           =   1100
   End
   Begin VB.Frame famSerial 
      Caption         =   "序列"
      Height          =   4896
      Left            =   252
      TabIndex        =   0
      Top             =   276
      Width           =   3288
      Begin VB.Frame Frame3 
         Caption         =   "自定义"
         Height          =   1575
         Left            =   300
         TabIndex        =   21
         Top             =   3150
         Width           =   2685
         Begin VB.CommandButton CmdSerialApply 
            Caption         =   "应用"
            Height          =   350
            Index           =   0
            Left            =   810
            TabIndex        =   28
            Top             =   1020
            Width           =   1100
         End
         Begin MSComCtl2.UpDown UDSerialCOL 
            Height          =   255
            Left            =   1980
            TabIndex        =   26
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UDSerialRow 
            Height          =   255
            Left            =   930
            TabIndex        =   29
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   450
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.TextBox TxtSerialRow 
            Height          =   315
            Left            =   450
            TabIndex        =   23
            Text            =   "1"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TxtSerialCol 
            Height          =   315
            Left            =   1500
            TabIndex        =   25
            Text            =   "1"
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "列:"
            Height          =   180
            Left            =   1500
            TabIndex        =   27
            Top             =   270
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "行:"
            Height          =   180
            Left            =   450
            TabIndex        =   24
            Top             =   270
            Width           =   270
         End
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   9
         Left            =   2100
         Picture         =   "frmSerset.frx":FD0C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2076
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   8
         Left            =   1200
         Picture         =   "frmSerset.frx":117BE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2076
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   7
         Left            =   300
         Picture         =   "frmSerset.frx":13270
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2076
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   6
         Left            =   2100
         Picture         =   "frmSerset.frx":14D22
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   5
         Left            =   1200
         Picture         =   "frmSerset.frx":167D4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   4
         Left            =   300
         Picture         =   "frmSerset.frx":18286
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1176
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   3
         Left            =   2100
         Picture         =   "frmSerset.frx":19D38
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   276
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   2
         Left            =   1200
         Picture         =   "frmSerset.frx":1B7EA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   276
         Width           =   900
      End
      Begin VB.CommandButton CmdSerial 
         Height          =   900
         Index           =   1
         Left            =   300
         Picture         =   "frmSerset.frx":1D29C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   276
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmSerialLayoutSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mthisForm As frmViewer

Public Sub zlShowMe(thisForm As frmViewer)
    Set mthisForm = thisForm
    Me.Show , mthisForm
End Sub

'图像应用
Private Sub CmdImgApply_Click()
    If Me.TxtImgCol <= G_INT_MAX_IMG_COL And Me.TxtImgCol >= 1 And Me.TxtImgRow <= G_INT_MAX_IMG_ROW And Me.TxtImgRow >= 1 Then
        Call subChangeImageLayout(Me.TxtImgRow, Me.TxtImgCol)
    Else
        Call subChangeImageLayout(1, 1)
    End If
End Sub

Private Sub CmdSerial_Click(Index As Integer)
    '设置序列布局和图像布局
    'Index从1-9 是序列布局；从11-19是图像布局
    
    Dim i As Integer, x As Integer, y As Integer, xx As Integer, Yy As Integer
    
    If Index = 1 Then
        mthisForm.intCountX = 1
        mthisForm.intCountY = 1
    ElseIf Index = 2 Then
        mthisForm.intCountX = 2
        mthisForm.intCountY = 1
    ElseIf Index = 3 Then
        mthisForm.intCountX = 1
        mthisForm.intCountY = 2
    ElseIf Index = 4 Then
        mthisForm.intCountX = 2
        mthisForm.intCountY = 2
    ElseIf Index = 5 Then
        mthisForm.intCountX = 4
        mthisForm.intCountY = 2
    ElseIf Index = 6 Then
        mthisForm.intCountX = 2
        mthisForm.intCountY = 4
    ElseIf Index = 7 Then
        mthisForm.intCountX = 4
        mthisForm.intCountY = 4
    ElseIf Index = 8 Then
        mthisForm.intCountX = 6
        mthisForm.intCountY = 4
    ElseIf Index = 9 Then
        mthisForm.intCountX = 8
        mthisForm.intCountY = 4
    ElseIf Index = 11 Then
        x = 1
        y = 1
    ElseIf Index = 12 Then
        x = 2
        y = 1
    ElseIf Index = 13 Then
        x = 1
        y = 2
    ElseIf Index = 14 Then
        x = 2
        y = 2
    ElseIf Index = 15 Then
        x = 4
        y = 2
    ElseIf Index = 16 Then
        x = 2
        y = 4
    ElseIf Index = 17 Then
        x = 4
        y = 4
    ElseIf Index = 18 Then
        x = 6
        y = 4
    ElseIf Index = 19 Then
        x = 8
        y = 4
    End If
    If Index > 10 And mthisForm.intSelectedSerial > 0 Then  '调整图像布局
        Call subChangeImageLayout(y, x)
    Else    '调整序列布局
        Call subChangeSeriesLayout(mthisForm)
    End If
End Sub

Private Sub subChangeImageLayout(intRows As Integer, intCols As Integer)
'------------------------------------------------
'功能：调整窗体中被选中viewer的图像行数和列数。
'参数：intCols--图像列数；intRows--图像行数。
'返回：无，直接调整viewer的行数和列数。
'------------------------------------------------
    Dim iBegin As Integer
    Dim iEnd As Integer
    Dim lngOldWidth As Long
    Dim lngOldHeight As Long
    Dim intImageIndex As Integer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    '判断是否选中全部序列，如果选中全部序列，需要对全部序列进行调整
    If mthisForm.isSelectAllSerial Then
        iBegin = 1
        iEnd = mthisForm.viewer.Count - 1
    Else
        iBegin = mthisForm.intSelectedSerial
        iEnd = iBegin
    End If
    
    '更改图像布局
    For i = iBegin To iEnd
        If mthisForm.viewer(i).Images.Count > 0 Then    '有图像才处理
            lngOldWidth = mthisForm.viewer(i).width / mthisForm.viewer(i).MultiColumns
            lngOldHeight = mthisForm.viewer(i).height / mthisForm.viewer(i).MultiRows
            intImageIndex = mthisForm.viewer(i).CurrentImage.Tag
            
            mthisForm.viewer(i).MultiColumns = intCols
            mthisForm.viewer(i).MultiRows = intRows
            
            '将当前的图像行数、列数设置保存到观片站运行参数里面
            mthisForm.MSFViewer.TextMatrix(i, 5) = intCols
            mthisForm.MSFViewer.TextMatrix(i, 6) = intRows
            
            '更改了图像布局后，需要重新显示图像
            Call subShowALLImage(mthisForm, mthisForm.viewer(i), intImageIndex, False)
            '如果图像曾经被缩放和移动，则保持图像的位置
            If mthisForm.viewer(i).Images.Count > 0 Then
                Call subScaleViewer(mthisForm.viewer(i), mthisForm.viewer(i).Images(1), lngOldWidth, lngOldHeight)
            End If
            '并重新显示滚动条
            Call subDisplayScrollBar(i, mthisForm, False)
        End If
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'序列应用
Private Sub CmdSerialApply_Click(Index As Integer)
    
    If Me.TxtSerialCol <= intMaxAreaX And Me.TxtSerialCol >= 1 And Me.TxtSerialRow <= intMaxAreaY And Me.TxtSerialRow >= 1 Then
        mthisForm.intCountX = TxtSerialCol
        mthisForm.intCountY = TxtSerialRow
    Else
        mthisForm.intCountX = 1
        mthisForm.intCountY = 1
    End If
    Call subChangeSeriesLayout(mthisForm)
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'2009用

    If mthisForm.intSelectedSerial = 0 Then '如果被选中序列为0 ，则禁止修改图像布局
        Me.famImage.Enabled = False
    Else        '在自定义中显示当前被选中序列的图像布局
        Me.TxtImgCol = mthisForm.viewer(mthisForm.intSelectedSerial).MultiColumns
        Me.TxtImgRow = mthisForm.viewer(mthisForm.intSelectedSerial).MultiRows
    End If
    '先根据系统设定，获取两个UpDown控件的最大值。
    UDSerialRow.Max = intMaxAreaY
    UDSerialCOL.Max = intMaxAreaX
    
    Me.TxtSerialCol = mthisForm.intCountX
    UDSerialCOL.Value = mthisForm.intCountX
    Me.TxtSerialRow = mthisForm.intCountY
    UDSerialRow.Value = mthisForm.intCountY
    
    If intMaxAreaX < 8 Or intMaxAreaY < 4 Then Me.CmdSerial(9).Enabled = False
    If intMaxAreaX < 6 Or intMaxAreaY < 4 Then Me.CmdSerial(8).Enabled = False
    If intMaxAreaX < 4 Or intMaxAreaY < 4 Then Me.CmdSerial(7).Enabled = False
    If intMaxAreaX < 2 Or intMaxAreaY < 4 Then Me.CmdSerial(6).Enabled = False
    If intMaxAreaX < 4 Or intMaxAreaY < 2 Then Me.CmdSerial(5).Enabled = False
    If intMaxAreaX < 2 Or intMaxAreaY < 2 Then Me.CmdSerial(4).Enabled = False
    If intMaxAreaX < 1 Or intMaxAreaY < 2 Then Me.CmdSerial(3).Enabled = False
    If intMaxAreaX < 2 Or intMaxAreaY < 1 Then Me.CmdSerial(6).Enabled = False
    
End Sub

Private Sub TxtImgCol_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtImgRow_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSerialCol_Change()
    If Val(TxtSerialCol) < 0 Or Val(TxtSerialCol) > intMaxAreaX Then
        MsgBox "序列纵向分布数必须在1和" & intMaxAreaX & "之间,请重新输入!", vbInformation, gstrSysName
        TxtSerialCol = 1
        TxtSerialCol.SetFocus
    End If
End Sub

Private Sub TxtSerialCol_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSerialRow_Change()
    If Val(TxtSerialRow) < 0 Or Val(TxtSerialRow) > intMaxAreaY Then
        MsgBox "序列纵向分布数必须在1和" & intMaxAreaY & "之间,请重新输入!", vbInformation, gstrSysName
        TxtSerialRow = 1
        TxtSerialRow.SetFocus
    End If
End Sub

Private Sub TxtSerialRow_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub UDImgCol_DownClick()
    If Me.TxtImgCol >= 2 Then
        Me.TxtImgCol = Me.TxtImgCol - 1
    End If
End Sub

Private Sub UDImgCol_UpClick()
    If Me.TxtImgCol < G_INT_MAX_IMG_COL Then
        Me.TxtImgCol = Me.TxtImgCol + 1
    End If
End Sub
Private Sub UDImgRow_DownClick()
    If Me.TxtImgRow >= 2 Then
        Me.TxtImgRow = Me.TxtImgRow - 1
    End If
End Sub

Private Sub UDImgRow_UpClick()
    If Me.TxtImgRow < G_INT_MAX_IMG_ROW Then
        Me.TxtImgRow = Me.TxtImgRow + 1
    End If
End Sub

Private Sub UDSerialCOL_DownClick()
    If Me.TxtSerialCol >= 2 Then
        Me.TxtSerialCol = Me.TxtSerialCol - 1
    End If
End Sub


Private Sub UDSerialCOL_UpClick()
    If Me.TxtSerialCol < intMaxAreaX Then
        Me.TxtSerialCol = Me.TxtSerialCol + 1
    End If
End Sub

Private Sub UDSerialRow_DownClick()
    If Me.TxtSerialRow >= 2 Then
        Me.TxtSerialRow = Me.TxtSerialRow - 1
    End If
End Sub

Private Sub UDSerialRow_UpClick()
    If Me.TxtSerialRow < intMaxAreaY Then
        Me.TxtSerialRow = Me.TxtSerialRow + 1
    End If
End Sub

