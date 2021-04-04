VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFeeGroupCollectEdit 
   Caption         =   "财务组收款单"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11970
   Icon            =   "frmFeeGroupCollectEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   11970
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBasicInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   600
      ScaleHeight     =   735
      ScaleWidth      =   10335
      TabIndex        =   24
      Top             =   360
      Width           =   10335
      Begin VB.ComboBox cboNO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8160
         TabIndex        =   27
         Top             =   30
         Width           =   2040
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   26
         Top             =   375
         Width           =   1785
      End
      Begin VB.ComboBox cboDept 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7800
         TabIndex        =   30
         Top             =   90
         Width           =   210
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "缴款人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   29
         Top             =   420
         Width           =   630
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "缴款部门"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2880
         TabIndex        =   28
         Top             =   420
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   30
      ScaleHeight     =   4935
      ScaleWidth      =   11655
      TabIndex        =   7
      Top             =   3930
      Width           =   11655
      Begin VB.PictureBox picSubDetail 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         ScaleHeight     =   1455
         ScaleWidth      =   11535
         TabIndex        =   9
         Top             =   1920
         Width           =   11535
         Begin VB.TextBox txtTime 
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   23
            Top             =   1080
            Width           =   2625
         End
         Begin VB.TextBox txtNote 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            MaxLength       =   500
            TabIndex        =   2
            Top             =   0
            Width           =   10305
         End
         Begin VB.TextBox txtChargePrepay 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   14
            Top             =   360
            Width           =   2625
         End
         Begin VB.TextBox txtBorrowTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4920
            TabIndex        =   13
            Top             =   360
            Width           =   2625
         End
         Begin VB.TextBox txtLendTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8760
            TabIndex        =   12
            Top             =   360
            Width           =   2625
         End
         Begin VB.TextBox txtSuppose 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1080
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   720
            Width           =   2625
         End
         Begin VB.TextBox txtActual 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4920
            MaxLength       =   16
            TabIndex        =   3
            Top             =   720
            Width           =   2625
         End
         Begin VB.TextBox txtRemain 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   8760
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   720
            Width           =   2625
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "摘要"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   22
            Top             =   52
            Width           =   420
         End
         Begin VB.Label lblChargePrepay 
            AutoSize        =   -1  'True
            Caption         =   "冲预交"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   360
            TabIndex        =   21
            Top             =   412
            Width           =   630
         End
         Begin VB.Label lblBorrowTotal 
            AutoSize        =   -1  'True
            Caption         =   "借款合计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4080
            TabIndex        =   20
            Top             =   405
            Width           =   840
         End
         Begin VB.Label lblLendTotal 
            AutoSize        =   -1  'True
            Caption         =   "借出合计"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7920
            TabIndex        =   19
            Top             =   405
            Width           =   840
         End
         Begin VB.Label lblSuppose 
            AutoSize        =   -1  'True
            Caption         =   "现金应收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   18
            Top             =   772
            Width           =   840
         End
         Begin VB.Label lblActual 
            AutoSize        =   -1  'True
            Caption         =   "现金实收"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4080
            TabIndex        =   17
            Top             =   765
            Width           =   840
         End
         Begin VB.Label lblRemain 
            AutoSize        =   -1  'True
            Caption         =   "本次暂存"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7920
            TabIndex        =   16
            Top             =   765
            Width           =   840
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "收款时间"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   1132
            Width           =   840
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsSettleList 
         Height          =   1575
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   10455
         _cx             =   18441
         _cy             =   2778
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483636
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFeeGroupCollectEdit.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
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
      Begin VB.Line linMain 
         BorderColor     =   &H8000000C&
         X1              =   -120
         X2              =   10320
         Y1              =   4320
         Y2              =   4320
      End
   End
   Begin VB.PictureBox picGeneralInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   720
      ScaleHeight     =   3375
      ScaleWidth      =   7815
      TabIndex        =   6
      Top             =   1305
      Width           =   7815
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   31
         Top             =   30
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupCollectEdit.frx":0626
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   2055
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   10440
         _cx             =   18415
         _cy             =   3625
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483636
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupCollectEdit.frx":0B74
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
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
   Begin VB.PictureBox picCommand 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   8
      Top             =   7080
      Width           =   2775
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   360
         TabIndex        =   4
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1575
         TabIndex        =   5
         Top             =   0
         Width           =   1100
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeeGroupCollectEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngWorkerID As Long, mstrWorkerName As String, mstrIDs As String   '传入的收款人信息
Private mlngGroupID As Long     '缴款组ID
Private mlngModule As Long, mblnWarning As Boolean
Private mblnOK As Boolean   '确定收款标识
Private mstrCashSettle As String    '结算方式字符串

Private Enum EM_Pan '界面板块序号
    EM_Pan_基本信息 = 1
    EM_Pan_收费单轧帐信息 = 2
    EM_Pan_结算信息 = 3
    EM_Pan_操作界面 = 4
End Enum
Private mstrTitle As String '用于窗体个性化保存的窗体名

'Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        KeyCode = 0
'        zlCommFun.PressKey vbKeyTab
'    End If
'End Sub

Private Sub cboNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    If mblnWarning = True Then
        If MsgBox("将要取消收费操作，已经输入的项目都会丢失！", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, blnChecked As Boolean, blnTrans As Boolean, strSQL As String
    Dim strIDs As String, strRemainNo As String, lngID As Long, strDetails As String, strNO As String
    Dim strTemp As String, colSql As New Collection, strFixedSql As String, blnBatch As Boolean
    Dim rsTmp As New ADODB.Recordset, strSelIDs As String
    
    On Error GoTo errHandle
    
    For i = 1 To vsRollingCurtain.Rows - 1
        If Val(vsRollingCurtain.TextMatrix(i, vsRollingCurtain.ColIndex("选择"))) = -1 Then blnChecked = True
    Next i
    
    If blnChecked = False Then
        MsgBox "进行收费操作必须勾选至少一项轧帐记录！", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If InStr(Trim(txtNote.Text), "'") > 0 Then
        MsgBox "注意:" & vbCrLf & "   收款说明不允许有单引号!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtNote.Enabled And txtNote.Visible Then txtNote.SetFocus
        Exit Sub
    End If
    
    '问题号:110281,焦博,2017/08/15,把轧账说明的上限从50个字符调整为500个字符
    If zlCommFun.ActualLen(txtNote.Text) > 500 Then
        MsgBox "注意:" & vbCrLf & "   收款说明最多只能输入500个字符或250个汉字,请重新输入!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtNote.Enabled And txtNote.Visible Then txtNote.SetFocus
        Exit Sub
    End If
    
    With vsRollingCurtain
        If .Rows = 1 Then Exit Sub
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) = -1 Then
                strSelIDs = strSelIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            End If
        Next i
        strSelIDs = Mid(strSelIDs, 2)
    End With
    
    '并发检查
    If CheckCollectEdit(strSelIDs) = False Then
        MsgBox "注意:" & vbCrLf & "   选择记录中有记录因为并发原因已被收款或者作废" & vbCrLf & "请重新选择记录!", _
                vbCritical + vbDefaultButton1 + vbOKOnly, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    strNO = zlDatabase.GetNextNo(138)
    
    If Val(txtRemain.Text) <> 0 Then
        strRemainNo = zlDatabase.GetNextNo(141)
    Else
        strRemainNo = ""
    End If

    lngID = zlDatabase.GetNextId("人员收缴记录")
    blnBatch = False
    
    strFixedSql = "" & _
                  "Zl_小组收款记录_Insert(" & lngID & ",'" & strNO & "'," & mlngGroupID & "," _
                  & Val(txtRemain.Text) & ",'" & strRemainNo & "','" & mstrCashSettle & "','" & UserInfo.姓名 & "','" & mstrWorkerName & "'," _
                  & "Null," & Val(txtChargePrepay.Text) & "," & Val(txtBorrowTotal.Text) & "," & Val(txtLendTotal.Text) _
                  & ",'" & txtNote.Text & "',to_date('" & txtTime.Text & "','yyyy-MM-dd HH24:mi:ss')" & ",'"
    '批量处理
    With vsRollingCurtain
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) = -1 Then
                strTemp = strIDs
                strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                If zlCommFun.ActualLen(strIDs) >= 4000 Then
                    strTemp = Mid(strTemp, 2)
                    If blnBatch = False Then
                        strSQL = strFixedSql & strTemp & "',0)"
                        blnBatch = True
                    Else
                        strSQL = strFixedSql & strTemp & "',1)"
                    End If
                    colSql.Add strSQL
                    strIDs = "," & Val(.TextMatrix(i, .ColIndex("ID")))
                End If
            End If
        Next i
        strIDs = Mid(strIDs, 2)
        If strIDs <> "" Then
            If blnBatch = False Then
                strSQL = strFixedSql & strIDs & "',0)"
            Else
                strSQL = strFixedSql & strIDs & "',1)"
            End If
            colSql.Add strSQL
        End If
    End With
    
    With vsSettleList
        For i = 1 To .Rows - 1
            strTemp = strDetails
            If .TextMatrix(i, .ColIndex("结算方式")) <> "" Then
                If Val(.TextMatrix(i, .ColIndex("性质"))) = 1 Then
                    '现金只计算实收
                    strDetails = strDetails & .TextMatrix(i, .ColIndex("结算方式")) & "," & .TextMatrix(i, .ColIndex("结算号码")) & "," & _
                                              Val(txtActual.Text) & ",|"
                Else
                    strDetails = strDetails & .TextMatrix(i, .ColIndex("结算方式")) & "," & .TextMatrix(i, .ColIndex("结算号码")) & "," & _
                                              Val(.TextMatrix(i, .ColIndex("金额"))) & ",|"
                End If
                If zlCommFun.ActualLen(strDetails) >= 4000 Then
                    strSQL = "Zl_小组收款结算_Insert(" & lngID & ",'" & UserInfo.姓名 & "','" & mstrWorkerName & "','" & strTemp & "',0)"
                    colSql.Add strSQL
                    strDetails = .TextMatrix(i, .ColIndex("结算方式")) & "," & .TextMatrix(i, .ColIndex("结算号码")) & "," & _
                                 Val(.TextMatrix(i, .ColIndex("金额"))) & ",|"
                End If
            End If
        Next i
    End With
    If strDetails <> "" Then
        strSQL = "Zl_小组收款结算_Insert(" & lngID & ",'" & UserInfo.姓名 & "','" & mstrWorkerName & "','" & strDetails & "',1)"
        colSql.Add strSQL
    End If
    
    cboNO.AddItem strNO
    
    On Error GoTo errSql
    
    Call zlExecuteProcedureArrAy(colSql, Me.Caption)
    '保存成功，提示主程序同步刷新数据
    mblnOK = True
    Call frmFeeGroupManage.AutoPrint(lngID, strNO, 1)
    Unload Me
    Exit Sub
errSql:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建DOCKINGPANEL控件
    '编制:刘尔旋
    '日期:2013-09-04
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    With dkpMain
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(EM_Pan_基本信息, 2000, 1000, DockTopOf)
        objPanel.Handle = picBasicInfo.hWnd
        objPanel.Title = "基本信息"
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 50
        objPanel.MaxTrackSize.Height = 50
        Set objPanel = .CreatePane(EM_Pan_收费单轧帐信息, 2000, 400, DockBottomOf, objPanel)
        objPanel.Handle = picGeneralInfo.hWnd
        objPanel.Title = "收费员轧账信息"
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        objPanel.MinTrackSize.Height = 50
        Set objPanel = .CreatePane(EM_Pan_结算信息, 2000, 1000, DockBottomOf, objPanel)
        objPanel.Handle = picDetail.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        objPanel.Title = "本次结算明细"
        objPanel.MinTrackSize.Height = 100
        Set objPanel = .CreatePane(EM_Pan_操作界面, 2000, 300, DockBottomOf, objPanel)
        objPanel.Handle = picCommand.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.MinTrackSize.Height = 25
        objPanel.MaxTrackSize.Height = 25
        Set .PaintManager.CaptionFont = lblActual.Font
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub


Private Sub Form_Activate()
'    If cboDept.ListCount > 1 Then
'        If cboDept.Enabled And cboDept.Visible Then cboDept.SetFocus
'    Else
    With vsRollingCurtain
        If .Enabled And .Visible Then .SetFocus
        If .Rows >= 2 Then .Select 1, 0
    End With
'    End If
    With dkpMain.FindPane(EM_Pan_结算信息)
        .MinTrackSize.Height = picDetail.Height / 15
    End With
    With dkpMain.FindPane(EM_Pan_操作界面)
        .MinTrackSize.Height = picCommand.Height / 15
        .MaxTrackSize.Height = picCommand.Height / 15
    End With
End Sub

Private Sub Form_Load()
    Dim i As Integer, rsTmp As New ADODB.Recordset, strSQL As String
    
    txtTime.Text = zlDatabase.Currentdate
'    strSQL = "Select b.Id, b.编码, b.名称, a.缺省" & vbNewLine & _
'             "From 部门人员 A, 部门表 B" & vbNewLine & _
'             "Where a.部门id = b.Id And a.人员id = [1] And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
'             " 　   And (b.站点 = '" & gstrNodeNo & "' Or b.站点 Is Null)"
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngWorkerID)
'    With rsTmp
'        Do While Not .EOF
'            cboDept.AddItem "[" & !编码 & "]" & !名称
'            cboDept.ItemData(cboDept.NewIndex) = Val(Nvl(!ID))
'            If Val(Nvl(!缺省)) = 1 Then cboDept.ListIndex = cboDept.NewIndex
'            .MoveNext
'        Loop
'    End With
'    If cboDept.ListIndex < 0 And cboDept.ListCount <> 0 Then cboDept.ListIndex = 0
    
    Call SetDockingPanel
    Call TextBoxPropertySet
    Call SetVSGrid
    Call LoadGeneralInfo(mstrIDs)
    Call CaculateSettleInfo
    mstrTitle = "财务组收款单"
    RestoreWinState Me, App.ProductName, mstrTitle
    mblnWarning = False
End Sub

Private Sub LoadGeneralInfo(ByVal strIDs As String)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:读取指定的收费员的收费信息
    '入参:lngWorkerID--收费员ID
    '编制:刘尔旋
    '日期:2013-09-09
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    strSQL = "Select /*+ Rule*/ a.Id, a.No, a.登记时间, a.开始时间, a.终止时间, a.冲预交款, a.借入合计, a.借出合计, a.摘要, a.收款员" & vbNewLine & _
             "From 人员收缴记录 A, Table(f_Num2list([1])) B" & vbNewLine & _
             "Where a.记录性质 = 1 And a.作废时间 Is Null And a.小组收款时间 Is Null And a.财务收款时间 Is Null And" & vbNewLine & _
             "      a.Id = b.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs)
    
    With vsRollingCurtain
        .Rows = 1
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("选择")) = -1
            .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = NVL(rsTmp!No)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = NVL(rsTmp!登记时间)
            .TextMatrix(.Rows - 1, .ColIndex("收款员")) = NVL(rsTmp!收款员)
            '.TextMatrix(.Rows - 1, .ColIndex("收款部门")) = Nvl(rsTmp!部门名称)
            .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = NVL(rsTmp!开始时间)
            .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = NVL(rsTmp!终止时间)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Format(NVL(rsTmp!冲预交款), "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Format(NVL(rsTmp!借入合计), "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Format(NVL(rsTmp!借出合计), "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("备注")) = NVL(rsTmp!摘要)
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTmp!ID)
            rsTmp.MoveNext
        Loop
        .AutoSize 1, .Cols - 1
    End With
    '恢复个性化设置
    zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Caption, "收费单轧帐信息", False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetVSGrid()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建VS控件属性
    '编制:刘尔旋
    '日期:2013-09-09
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    With vsRollingCurtain
        .Rows = 1
        '.ColDataType(.ColIndex("ID")) = flexDTBoolean
        .Editable = flexEDKbdMouse
    End With
    vsSettleList.Editable = flexEDKbdMouse
    Call SetGrid
End Sub

Private Sub TextBoxPropertySet()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置文本框格式
    '编制:刘尔旋
    '日期:2013-09-05
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    txtName.Text = mstrWorkerName
    txtName.Locked = True
    txtName.Enabled = False
    txtChargePrepay.Locked = True
    txtChargePrepay.Enabled = False
    txtBorrowTotal.Locked = True
    txtBorrowTotal.Enabled = False
    txtLendTotal.Locked = True
    txtLendTotal.Enabled = False
    txtSuppose.Locked = True
    txtRemain.Locked = True
    txtTime.Enabled = False
    txtTime.Locked = True
End Sub

Public Function ShowMe(frmMain As Object, ByVal lngModule As Long, ByVal lngWorkerID As Long, ByVal strWorkerName As String, _
                       ByVal lngGroupID As Long, ByVal strIDs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:外部调用初始化接口
    '入参:frmMain--外部调用窗体
    '     lngModule--模块号
    '     lngWorkerID--收费员ID
    '     strWorkerName--收费员姓名
    '     lngGroupID--缴款组ID
    '     strIDs --收费条目ID集合: ID1,ID2,...IDn
    '返回:成功返回True,失败返回False
    '编制:刘尔旋
    '日期:2013-09-09
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mlngWorkerID = lngWorkerID
    mlngGroupID = lngGroupID
    mstrWorkerName = strWorkerName
    mstrIDs = strIDs
    mlngModule = lngModule
    Me.Show vbModal, frmMain
    ShowMe = mblnOK
    Exit Function
errHandle:
    ShowMe = False
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckCollectEdit(ByVal strSelIDs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:保存前并发检查
    '入参:选择的轧帐记录的ID-ID1,ID2,...,IDn
    '返回:成功返回True,失败返回False
    '编制:刘尔旋
    '日期:2013-10-14
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    strSQL = "Select /*+ Rule*/ a.Id " & vbNewLine & _
             "From 人员收缴记录 A, Table(f_Num2list([1])) B" & vbNewLine & _
             "Where a.记录性质 = 1 And (a.作废时间 Is Not Null Or a.小组收款时间 Is Not Null Or a.财务收款时间 Is Not Null) And" & vbNewLine & _
             "      a.Id = b.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSelIDs)
    '查到有并发操作记录
    If rsTmp.RecordCount >= 1 Then
        CheckCollectEdit = False
        Exit Function
    End If
    CheckCollectEdit = True
    Exit Function
errHandle:
    CheckCollectEdit = False
    If ErrCenter = 1 Then Resume
End Function

Private Sub CaculateSettleInfo()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:计算结算信息
    '编制:刘尔旋
    '日期:2013-9-5
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim i As Integer, strSQL As String, dblTotal As Double, strSelIDs As String
    Dim rsTmp As New ADODB.Recordset, blnAdd As Boolean, intRow As Integer
    Dim dblCharge As Double, dblBorrow As Double, dblLend As Double, dblTemp As Double
    With vsRollingCurtain
        If .Rows = 1 Then Exit Sub
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) = -1 Then
                strSelIDs = strSelIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
            End If
        Next i
        strSelIDs = Mid(strSelIDs, 2)
    End With
    With vsSettleList
        .Rows = 1
        If strSelIDs = "" Then GoTo EndCalc
        strSQL = "" & _
        "Select /*+ Rule*/" & vbNewLine & _
        "       a.结算方式, Trim(To_Char(Sum(a.金额), '9999999999" & gstrDec & "')) As 金额, Null As 结算号," & vbNewLine & _
        "       Decode(d.性质, 1, 1, 2, 2, 7, 3, 8, 4, 3, 5, 4, 6, 7) As 性质" & vbNewLine & _
        "From 人员收缴明细 A, Table(f_Num2list([1])) B, 结算方式 D" & vbNewLine & _
        "Where a.收缴id = b.Column_Value And a.结算方式 = d.名称" & vbNewLine & _
        "Group By 结算方式, 结算号, 性质" & vbNewLine & _
        "Order By 性质 Asc"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSelIDs)
        Set .DataSource = rsTmp
        If rsTmp.RecordCount = 0 Then
            .Clear 1
            .Rows = 2
        End If
        .AutoSize 0, .Cols - 1
        .ColWidth(.ColIndex("结算号码")) = 4500
    End With
    '恢复个性化设置
    'zl_vsGrid_Para_Restore mlngModule, vsSettleList, Me.Caption, "结算方式", False
EndCalc:
    dblTemp = 0
    
    With vsRollingCurtain
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) = -1 Then
                dblCharge = dblCharge + Val(.TextMatrix(i, .ColIndex("冲预交款")))
                dblBorrow = dblBorrow + Val(.TextMatrix(i, .ColIndex("借入合计")))
                dblLend = dblLend + Val(.TextMatrix(i, .ColIndex("借出合计")))
            End If
        Next i
        txtChargePrepay.Text = Format(dblCharge, "0.00")
        txtBorrowTotal.Text = Format(dblBorrow, "0.00")
        txtLendTotal.Text = Format(dblLend, "0.00")
    End With
    
    With vsSettleList
        '计算现金金额
        dblTemp = 0
        dblTotal = 0
        For i = 1 To .Rows - 1
            dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("金额")))
            If Val(.TextMatrix(i, .ColIndex("性质"))) = 1 Then
                mstrCashSettle = .TextMatrix(i, .ColIndex("结算方式"))
                dblTemp = dblTemp + Val(.TextMatrix(i, .ColIndex("金额")))
            End If
        Next i
        txtSuppose.Text = Format(dblTemp, "0.00")
        If Val(txtSuppose.Text) = 0 Then
            txtActual.Enabled = False
            txtActual.BackColor = &H80000000
        Else
            txtActual.Enabled = True
            txtActual.BackColor = &H80000005
        End If
        txtActual.Text = Format(dblTemp, "0.00")
        txtRemain.Text = Format(Val(txtSuppose.Text) - Val(txtActual.Text), "0.00")
    End With
    dkpMain.FindPane(EM_Pan_结算信息).Title = "本次结算明细:  " & Format(dblTotal, "0.00") & " 元"
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '窗体最小尺寸
    If Me.Width < 11745 Then Me.Width = 11745
    If Me.Height < 7065 Then Me.Height = 7680
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Caption, "收费单轧帐信息", False
    'zl_vsGrid_Para_Save mlngModule, vsSettleList, Me.Caption, "结算方式", False
    SaveWinState Me, App.ProductName, mstrTitle
End Sub

Private Sub picBasicInfo_Resize()
    cboNO.Left = picBasicInfo.Width - 2500
    lblNO.Left = cboNO.Left - 400
End Sub

Private Sub picCommand_Resize()
    On Error Resume Next
    cmdOK.Left = picCommand.Width - 3000
    cmdCancel.Left = picCommand.Width - 1800
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    vsSettleList.Width = picDetail.Width
    linMain.X1 = 0
    linMain.Y1 = picDetail.Height - 15
    linMain.X2 = picDetail.Width
    linMain.Y2 = picDetail.Height - 15
    picSubDetail.Left = 0
    picSubDetail.Width = picDetail.Width
    picSubDetail.Top = picDetail.Height - 50 - picSubDetail.Height
    vsSettleList.Height = picSubDetail.Top - 100
End Sub

Private Sub picGeneralInfo_Resize()
    On Error Resume Next
    vsRollingCurtain.Width = picGeneralInfo.Width
    vsRollingCurtain.Height = picGeneralInfo.Height
End Sub

Private Sub picSubDetail_Resize()
    On Error Resume Next
    '界面布局调整
    txtNote.Width = picSubDetail.Width - txtNote.Left - 300
    txtActual.Width = txtNote.Width / 4
    txtBorrowTotal.Width = txtNote.Width / 4
    txtChargePrepay.Width = txtNote.Width / 4
    txtLendTotal.Width = txtNote.Width / 4
    txtLendTotal.Left = picSubDetail.Width - txtLendTotal.Width - 300
    lblLendTotal.Left = txtLendTotal.Left - 960
    txtRemain.Width = txtNote.Width / 4
    txtRemain.Left = picSubDetail.Width - txtRemain.Width - 300
    lblRemain.Left = txtRemain.Left - 960
    txtSuppose.Width = txtNote.Width / 4
    txtTime.Width = txtNote.Width / 4
    lblBorrowTotal.Left = txtChargePrepay.Left + txtChargePrepay.Width + _
                          ((lblLendTotal.Left - txtChargePrepay.Left - txtChargePrepay.Width) - (txtBorrowTotal.Width + 960)) / 2
    txtBorrowTotal.Left = lblBorrowTotal.Left + 960
    lblActual.Left = lblBorrowTotal.Left
    txtActual.Left = txtBorrowTotal.Left
End Sub

Private Sub txtActual_Change()
    If Val(txtActual.Text) > Val(txtSuppose.Text) Then
        txtActual.Text = txtSuppose.Text
        Call zlControl.TxtSelAll(txtActual)
    End If
    txtRemain.Text = Format(Val(txtSuppose.Text) - Val(txtActual.Text), "0.00")
    mblnWarning = True
End Sub

Private Sub txtActual_GotFocus()
    Call zlControl.TxtSelAll(txtActual)
End Sub

Private Sub txtActual_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtActual_KeyPress(KeyAscii As Integer)
    '限定输入数字
    If (KeyAscii < Asc(".") Or KeyAscii = Asc("/") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
        KeyAscii = 0
        Exit Sub
    End If
    '小数点的判断
    If KeyAscii = Asc(".") And InStr(1, txtActual.Text, ".") > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtNote_GotFocus()
    zlCommFun.OpenIme True
    Call zlControl.TxtSelAll(txtNote)
End Sub

Private Sub txtNote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("‘") Then KeyAscii = 0
End Sub

Private Sub txtNote_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtRemain_Change()
    If Val(txtRemain.Text) = 0 Then
        txtRemain.ForeColor = &H80000008
    Else
        txtRemain.ForeColor = vbRed
    End If
End Sub

Private Sub txtSuppose_Change()
    If Val(txtSuppose.Text) = 0 Then
        txtSuppose.ForeColor = &H80000008
    Else
        txtSuppose.ForeColor = vbBlue
    End If
End Sub

Private Sub vsRollingCurtain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call zl_VsGridRowChange(vsRollingCurtain, OldRow, NewRow, OldCol, NewCol)
    With vsRollingCurtain
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
    With vsRollingCurtain
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsRollingCurtain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then zlCommFun.PressKey vbKeyTab: KeyCode = 0
End Sub

Private Sub vsRollingCurtain_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsRollingCurtain)
End Sub

Private Sub vsSettleList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    Call zl_VsGridRowChange(vsSettleList, OldRow, NewRow, OldCol, NewCol)
    With vsSettleList
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsSettleList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsSettleList.ColIndex("结算号码") Then Cancel = True
End Sub

Private Sub vsSettleList_GotFocus()
    Call zl_VsGridGotFocus(vsSettleList)
    With vsSettleList
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsSettleList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc(",") Or KeyAscii = Asc("|") Then KeyAscii = 0
    mblnWarning = True
End Sub

Private Sub vsSettleList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row <= 1 Then Exit Sub
    If KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0
    mblnWarning = True
End Sub

Private Sub vsSettleList_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsSettleList)
End Sub

Private Sub vsSettleList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '数据验证
    With vsSettleList
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("结算号码")
            If zlCommFun.ActualLen(strKey) > 10 Then
                MsgBox "结算号码超长,最多只能输入10个字符或5个汉字", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If InStr(1, strKey, "'") > 0 Or InStr(1, strKey, "|") > 0 Or InStr(1, strKey, ",") > 0 Then
                MsgBox "结算号码中不能包含特殊字符:',| ", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        End Select
    End With
End Sub

Private Sub SetGrid()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSF控件
    '编制:刘尔旋
    '日期:2013-10-13
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    With vsRollingCurtain
        For i = 0 To .Cols - 1
            If .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Or .ColKey(i) = "收款员" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "过滤" Or .ColKey(i) = "收款部门" Or .ColKey(i) = "选择" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Then .ColData(i) = "1|0"
        Next
    End With
    
    zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Caption, "收费单轧帐信息", False
End Sub

Private Sub vsRollingCurtain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call CaculateSettleInfo
    mblnWarning = True
End Sub

Private Sub vsRollingCurtain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then
        vsRollingCurtain.Select Row, 1
        Cancel = True
    End If
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Caption, "收费单轧帐信息", False
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub

