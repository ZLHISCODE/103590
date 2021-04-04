VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFilmConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "排版格式"
   ClientHeight    =   6090
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "frmFilmConfg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab sstabFilmConfig 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "排版格式"
      TabPicture(0)   =   "frmFilmConfg.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCustom"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "参数设置"
      TabPicture(1)   =   "frmFilmConfg.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame8 
         Caption         =   "设置"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   5295
         Begin VB.ComboBox cboPageRange 
            Height          =   300
            ItemData        =   "frmFilmConfg.frx":0044
            Left            =   2880
            List            =   "frmFilmConfg.frx":004E
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   4200
            Width           =   2000
         End
         Begin VB.ComboBox cboFormat 
            Height          =   315
            Left            =   360
            TabIndex        =   74
            Top             =   600
            Width           =   2000
         End
         Begin VB.ComboBox cboMedium 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":0062
            Left            =   360
            List            =   "frmFilmConfg.frx":006C
            TabIndex        =   73
            Top             =   2760
            Width           =   2000
         End
         Begin VB.ComboBox cboOrientation 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":0087
            Left            =   360
            List            =   "frmFilmConfg.frx":0091
            TabIndex        =   72
            Top             =   1320
            Width           =   2000
         End
         Begin VB.ComboBox cboFilmBox 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":00AA
            Left            =   360
            List            =   "frmFilmConfg.frx":00D2
            TabIndex        =   71
            Top             =   3480
            Width           =   2000
         End
         Begin VB.ComboBox cboMagnification 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":0132
            Left            =   360
            List            =   "frmFilmConfg.frx":0142
            TabIndex        =   70
            Top             =   2040
            Width           =   2000
         End
         Begin VB.ComboBox cboTrim 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":0168
            Left            =   360
            List            =   "frmFilmConfg.frx":0172
            TabIndex        =   69
            Top             =   4200
            Width           =   2000
         End
         Begin VB.ComboBox cboPriority 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":017F
            Left            =   2880
            List            =   "frmFilmConfg.frx":018C
            TabIndex        =   57
            Top             =   600
            Width           =   2000
         End
         Begin VB.ListBox lstCopies 
            Height          =   240
            ItemData        =   "frmFilmConfg.frx":01A0
            Left            =   2880
            List            =   "frmFilmConfg.frx":01C2
            TabIndex        =   56
            Top             =   2760
            Width           =   2000
         End
         Begin VB.ComboBox cboFilmSize 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":01E5
            Left            =   2880
            List            =   "frmFilmConfg.frx":01E7
            TabIndex        =   55
            Top             =   1320
            Width           =   2000
         End
         Begin VB.ComboBox cboResolution 
            Height          =   300
            ItemData        =   "frmFilmConfg.frx":01E9
            Left            =   2880
            List            =   "frmFilmConfg.frx":01F3
            TabIndex        =   54
            Top             =   3480
            Width           =   2000
         End
         Begin VB.ComboBox cboSmooth 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":0207
            Left            =   2880
            List            =   "frmFilmConfg.frx":0214
            TabIndex        =   53
            Top             =   2040
            Width           =   2000
         End
         Begin VB.Label Label1 
            Caption         =   "页码范围"
            Height          =   255
            Left            =   2880
            TabIndex        =   76
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label20 
            Caption         =   "格式"
            Height          =   255
            Left            =   360
            TabIndex        =   68
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label21 
            Caption         =   "优先级"
            Height          =   255
            Left            =   2880
            TabIndex        =   67
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "介质"
            Height          =   255
            Left            =   360
            TabIndex        =   66
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label23 
            Caption         =   "打印份数"
            Height          =   255
            Left            =   2880
            TabIndex        =   65
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "胶片方向"
            Height          =   255
            Left            =   360
            TabIndex        =   64
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label25 
            Caption         =   "胶片规格"
            Height          =   255
            Left            =   2880
            TabIndex        =   63
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label26 
            Caption         =   "片盒"
            Height          =   255
            Left            =   360
            TabIndex        =   62
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "分辨率"
            Height          =   255
            Left            =   2880
            TabIndex        =   61
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label28 
            Caption         =   "放大模式"
            Height          =   255
            Left            =   360
            TabIndex        =   60
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label29 
            Caption         =   "平滑模式"
            Height          =   255
            Left            =   2880
            TabIndex        =   59
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label30 
            Caption         =   "修整"
            Height          =   255
            Left            =   360
            TabIndex        =   58
            Top             =   3960
            Width           =   855
         End
      End
      Begin VB.Frame fraCustom 
         Height          =   4305
         Left            =   2880
         TabIndex        =   19
         Top             =   600
         Width           =   2436
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   1
            Left            =   876
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   300
            Width           =   1008
         End
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   2
            Left            =   876
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   816
            Width           =   1008
         End
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   3
            Left            =   876
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   1296
            Width           =   1008
         End
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   4
            Left            =   876
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   1776
            Width           =   1008
         End
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   5
            Left            =   876
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   2280
            Width           =   1008
         End
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   6
            Left            =   876
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   2760
            Width           =   1008
         End
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   7
            Left            =   876
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   3240
            Width           =   1008
         End
         Begin VB.TextBox txtC 
            Height          =   300
            Index           =   8
            Left            =   876
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   3696
            Width           =   1008
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   1
            Left            =   1620
            TabIndex        =   20
            Top             =   300
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   2
            Left            =   1620
            TabIndex        =   21
            Top             =   816
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   3
            Left            =   1620
            TabIndex        =   22
            Top             =   1296
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   4
            Left            =   1620
            TabIndex        =   23
            Top             =   1776
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   5
            Left            =   1620
            TabIndex        =   24
            Top             =   2280
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   6
            Left            =   1620
            TabIndex        =   25
            Top             =   2760
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   7
            Left            =   1620
            TabIndex        =   26
            Top             =   3240
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDC 
            Height          =   300
            Index           =   8
            Left            =   1620
            TabIndex        =   27
            Top             =   3696
            Width           =   252
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            Min             =   1
            Enabled         =   -1  'True
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   51
            Top             =   360
            Width           =   456
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   8
            Left            =   360
            TabIndex        =   50
            Top             =   3756
            Width           =   456
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   2
            Left            =   360
            TabIndex        =   49
            Top             =   876
            Width           =   456
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   3
            Left            =   360
            TabIndex        =   48
            Top             =   1356
            Width           =   456
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   4
            Left            =   360
            TabIndex        =   47
            Top             =   1836
            Width           =   456
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   5
            Left            =   360
            TabIndex        =   46
            Top             =   2340
            Width           =   456
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   6
            Left            =   360
            TabIndex        =   45
            Top             =   2820
            Width           =   456
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "第8行"
            Height          =   180
            Index           =   7
            Left            =   360
            TabIndex        =   44
            Top             =   3300
            Width           =   456
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   1
            Left            =   1920
            TabIndex        =   43
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   8
            Left            =   1944
            TabIndex        =   42
            Top             =   3756
            Width           =   180
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   2
            Left            =   1944
            TabIndex        =   41
            Top             =   876
            Width           =   180
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   3
            Left            =   1944
            TabIndex        =   40
            Top             =   1356
            Width           =   180
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   4
            Left            =   1944
            TabIndex        =   39
            Top             =   1836
            Width           =   180
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   5
            Left            =   1944
            TabIndex        =   38
            Top             =   2340
            Width           =   180
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   6
            Left            =   1944
            TabIndex        =   37
            Top             =   2820
            Width           =   180
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            Caption         =   "列"
            Height          =   180
            Index           =   7
            Left            =   1944
            TabIndex        =   36
            Top             =   3300
            Width           =   180
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "胶片参数"
         Height          =   1440
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   2448
         Begin VB.ComboBox cobSize 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":022F
            Left            =   828
            List            =   "frmFilmConfg.frx":023C
            TabIndex        =   16
            Text            =   "8INX10IN"
            Top             =   372
            Width           =   1332
         End
         Begin VB.ComboBox cobAspect 
            Height          =   315
            ItemData        =   "frmFilmConfg.frx":0260
            Left            =   804
            List            =   "frmFilmConfg.frx":026A
            TabIndex        =   15
            Text            =   "纵向"
            Top             =   864
            Width           =   1380
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "尺寸:"
            Height          =   180
            Left            =   180
            TabIndex        =   18
            Top             =   432
            Width           =   456
         End
         Begin VB.Label Label4 
            Caption         =   "方向:"
            Height          =   276
            Left            =   132
            TabIndex        =   17
            Top             =   900
            Width           =   528
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "行列定义"
         Height          =   2724
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2424
         Begin VB.OptionButton Option 
            Caption         =   "行自定义"
            Height          =   324
            Index           =   1
            Left            =   480
            TabIndex        =   13
            Top             =   348
            Width           =   1176
         End
         Begin VB.OptionButton Option 
            Caption         =   "列自定义"
            Height          =   180
            Index           =   2
            Left            =   480
            TabIndex        =   12
            Top             =   780
            Width           =   1116
         End
         Begin VB.OptionButton Option 
            Caption         =   "标准行列"
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   4
            Top             =   1176
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.Frame fraFormat 
            Height          =   1464
            Left            =   360
            TabIndex        =   5
            Top             =   1150
            Width           =   1890
            Begin VB.TextBox txtRow 
               Height          =   300
               Left            =   240
               TabIndex        =   9
               Text            =   "2"
               Top             =   960
               Width           =   980
            End
            Begin VB.TextBox txtCol 
               Height          =   300
               Left            =   240
               TabIndex        =   8
               Text            =   "2"
               Top             =   360
               Width           =   980
            End
            Begin MSComCtl2.UpDown UpDcol 
               Height          =   300
               Left            =   1200
               TabIndex        =   6
               Top             =   360
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   2
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDrow 
               Height          =   300
               Left            =   1200
               TabIndex        =   7
               Top             =   960
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   9
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "列"
               Height          =   180
               Left            =   1560
               TabIndex        =   11
               Top             =   420
               Width           =   180
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "行"
               Height          =   180
               Left            =   1560
               TabIndex        =   10
               Top             =   1020
               Width           =   180
            End
         End
      End
   End
   Begin VB.CommandButton cmdCnanel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   1
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton cndOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1080
      TabIndex        =   0
      Top             =   5520
      Width           =   1100
   End
End
Attribute VB_Name = "frmFilmConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public f As frmFilm
Private mblnOk As Boolean

Private Sub cmdCnanel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cndOK_Click()
    If Me.sstabFilmConfig.TabVisible(0) = True Then
        '设置排版格式
        Call f.subConfig
    ElseIf Me.sstabFilmConfig.TabVisible(1) = True Then
        '设置胶片参数
        Set f.clsTruePrinter = f.funFillPrinterParams(True)
    End If
    
    mblnOk = True
    
    Unload Me
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    On Error GoTo err
    
    mblnOk = False
    
    For i = 1 To 8
        lblC(i) = "第" & i & "行"
        lblH(i) = "列"
        txtC(i) = "2"
        Me.UpDC(i).Value = Val(txtC(i))
    Next
    Me.UpDrow.Value = Val(txtRow)
    Me.UpDcol.Value = Val(txtCol)
    If Me.sstabFilmConfig.TabVisible(1) Then        '在打印设置可见的情况下才填充combobox
        Dim strSQL As String
        Me.cboFilmSize.Clear
        Me.cboFormat.Clear
        Me.cobSize.Clear
        
        If blLocalRun = True Then
            strSQL = "SELECT 规格标识 as 名称 FROM 影像胶片规格"
            Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
        Else
            strSQL = "SELECT 名称 FROM 影像胶片规格"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        End If
        While Not rsTemp.EOF
            Me.cboFilmSize.AddItem rsTemp!名称
            Me.cobSize.AddItem rsTemp!名称
            rsTemp.MoveNext
        Wend
        
        If blLocalRun = True Then
            strSQL = "SELECT 格式标识 as 名称 FROM 影像打印格式"
            Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
        Else
            strSQL = "SELECT 名称 FROM 影像打印格式"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        End If
        While Not rsTemp.EOF
            Me.cboFormat.AddItem rsTemp!名称
            rsTemp.MoveNext
        Wend
    End If
    
    cboPageRange.ListIndex = 0
    
    Call subSZ
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Option_Click(Index As Integer)
    Call subSZ
End Sub

Private Sub subSZ()
Dim i As Integer

    If Me.Option(0) Then
        fraCustom.Enabled = False
        fraFormat.Enabled = True
    ElseIf Me.Option(1) Then            '行自定义
        fraCustom.Enabled = True
        fraFormat.Enabled = False
        For i = 1 To 8
            lblC(i) = "第" & i & "行"
            lblH(i) = "列"
            If Val(Me.txtRow) > 8 Then Me.txtRow = 8
            
            If i <= Val(Me.txtRow) Then
                Me.txtC(i).Enabled = True
                Me.UpDC(i).Enabled = True
            Else
                Me.txtC(i).Enabled = False
                Me.UpDC(i).Enabled = False
            End If
        Next
    ElseIf Me.Option(2) Then            '列自定义
        fraCustom.Enabled = True
        fraFormat.Enabled = False
        For i = 1 To 8
            lblC(i) = "第" & i & "列"
            lblH(i) = "行"
            If Val(Me.txtCol) > 8 Then Me.txtCol = 8
            
            If i <= Val(Me.txtCol) Then
                Me.txtC(i).Enabled = True
                Me.UpDC(i).Enabled = True
            Else
                Me.txtC(i).Enabled = False
                Me.UpDC(i).Enabled = False
            End If
        Next
    End If
End Sub

Private Sub txtC_Change(Index As Integer)
    If Val(txtC(Index)) < 1 Or Val(txtC(Index)) > 10 Then
        MsgBox "行列值必须在1-10之间", vbInformation, gstrSysName
        txtC(Index).SetFocus
    Else
        Me.UpDC(Index).Value = Val(txtC(Index))
    End If
End Sub

Private Sub txtCol_Change()
    If Val(txtCol) < 1 Or Val(txtCol) > 10 Then
        MsgBox "行列值必须在1-10之间", vbInformation, gstrSysName
        txtCol.SetFocus
    Else
        Me.UpDcol.Value = Val(txtCol)
    End If
End Sub

Private Sub txtRow_Change()
    If Val(txtRow) < 1 Or Val(txtRow) > 10 Then
        MsgBox "行列值必须在1-10之间", vbInformation, gstrSysName
        txtRow.SetFocus
    Else
        Me.UpDrow.Value = Val(txtRow)
    End If
    
End Sub

Private Sub UpDC_Change(Index As Integer)
    txtC(Index) = UpDC(Index).Value
End Sub

Private Sub UpDcol_Change()
    txtCol = UpDcol.Value

End Sub

Private Sub UpDrow_Change()
    txtRow = UpDrow.Value
End Sub

Public Function zlShowMe() As Boolean
    
    On Error GoTo err
    
    Call Me.Show(1, f)
    
    zlShowMe = mblnOk
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function
