VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISBorrowEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   0
      Left            =   195
      ScaleHeight     =   7455
      ScaleWidth      =   9945
      TabIndex        =   2
      Top             =   75
      Width           =   9945
      Begin VB.Frame fra 
         Height          =   7005
         Left            =   60
         TabIndex        =   26
         Top             =   45
         Width           =   9225
         Begin VB.TextBox txtBorrowUser 
            Height          =   300
            Left            =   1065
            ScrollBars      =   2  'Vertical
            TabIndex        =   46
            Top             =   825
            Width           =   4500
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   3
            Left            =   5580
            Picture         =   "frmCISBorrowEdit.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1515
            Width           =   315
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   2
            Left            =   5580
            Picture         =   "frmCISBorrowEdit.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1185
            Width           =   315
         End
         Begin VB.CommandButton cmd 
            Height          =   300
            Index           =   0
            Left            =   5580
            Picture         =   "frmCISBorrowEdit.frx":D0A4
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   810
            Width           =   315
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   4740
            TabIndex        =   36
            Top             =   480
            Width           =   1170
         End
         Begin VB.TextBox txtConver 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   1095
            TabIndex        =   32
            Top             =   510
            Width           =   1260
         End
         Begin VB.TextBox txtConver 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   2625
            TabIndex        =   31
            Top             =   510
            Width           =   1260
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
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
            Height          =   225
            Index           =   5
            Left            =   510
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   150
            Width           =   1155
         End
         Begin VB.PictureBox picPane 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1410
            Index           =   1
            Left            =   315
            ScaleHeight     =   1410
            ScaleWidth      =   8700
            TabIndex        =   27
            Top             =   5100
            Width           =   8700
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   9
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   43
               Top             =   1065
               Width           =   1530
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   1
               Left            =   615
               ScrollBars      =   2  'Vertical
               TabIndex        =   42
               Top             =   1065
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   8
               Left            =   615
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               Top             =   720
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   7
               Left            =   615
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   375
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   6
               Left            =   630
               ScrollBars      =   2  'Vertical
               TabIndex        =   9
               Top             =   30
               Width           =   915
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   0
               Left            =   4665
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   720
               Width           =   3075
            End
            Begin VB.TextBox txtConver 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   240
               Index           =   2
               Left            =   4695
               TabIndex        =   29
               Top             =   420
               Width           =   1245
            End
            Begin VB.TextBox txtConver 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   240
               Index           =   3
               Left            =   6255
               TabIndex        =   28
               Top             =   420
               Width           =   1455
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   2
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   45
               Width           =   1530
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   3
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Top             =   390
               Width           =   1530
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   4
               Left            =   2355
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   720
               Width           =   1530
            End
            Begin MSComCtl2.DTPicker dtp 
               Height          =   300
               Index           =   2
               Left            =   4665
               TabIndex        =   17
               Top             =   390
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   91684867
               CurrentDate     =   39500
            End
            Begin MSComCtl2.DTPicker dtp 
               Height          =   300
               Index           =   3
               Left            =   6210
               TabIndex        =   19
               Top             =   390
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   529
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   91684867
               CurrentDate     =   39500
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "归还时间"
               Height          =   180
               Index           =   19
               Left            =   1605
               TabIndex        =   45
               Top             =   1110
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "归还人"
               Height          =   180
               Index           =   18
               Left            =   45
               TabIndex        =   44
               Top             =   1125
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "申请人"
               Height          =   180
               Index           =   0
               Left            =   30
               TabIndex        =   8
               Top             =   75
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "登记时间"
               Height          =   180
               Index           =   1
               Left            =   1605
               TabIndex        =   10
               Top             =   90
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "批准人"
               Height          =   180
               Index           =   6
               Left            =   45
               TabIndex        =   12
               Top             =   450
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "批准时间"
               Height          =   180
               Index           =   7
               Left            =   1605
               TabIndex        =   14
               Top             =   435
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "拒借人"
               Height          =   180
               Index           =   8
               Left            =   45
               TabIndex        =   20
               Top             =   780
               Width           =   540
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "拒借时间"
               Height          =   180
               Index           =   9
               Left            =   1605
               TabIndex        =   22
               Top             =   765
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "拒借理由"
               Height          =   180
               Index           =   10
               Left            =   3930
               TabIndex        =   24
               Top             =   780
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "借阅时间"
               Height          =   180
               Index           =   11
               Left            =   3930
               TabIndex        =   16
               Top             =   450
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "至"
               Height          =   180
               Index           =   12
               Left            =   6015
               TabIndex        =   18
               Top             =   435
               Width           =   330
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1725
            Index           =   1
            Left            =   1065
            TabIndex        =   7
            Top             =   1185
            Width           =   4500
            _cx             =   7937
            _cy             =   3043
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
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
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
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
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   2595
            TabIndex        =   34
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   122945539
            CurrentDate     =   39500
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1065
            TabIndex        =   3
            Top             =   480
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   122945539
            CurrentDate     =   39500
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Index           =   17
            Left            =   3930
            TabIndex        =   41
            Top             =   885
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Index           =   16
            Left            =   3915
            TabIndex        =   40
            Top             =   2955
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   210
            Index           =   15
            Left            =   7995
            TabIndex        =   35
            Top             =   150
            Width           =   120
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000011&
            X1              =   510
            X2              =   1815
            Y1              =   375
            Y2              =   375
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Index           =   2
            Left            =   2400
            TabIndex        =   1
            Top             =   525
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请理由"
            Height          =   180
            Index           =   3
            Left            =   3990
            TabIndex        =   4
            Top             =   525
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "申请借阅:"
            Height          =   180
            Index           =   13
            Left            =   60
            TabIndex        =   0
            Top             =   540
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No:"
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
            Height          =   210
            Index           =   14
            Left            =   90
            TabIndex        =   33
            Top             =   150
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "借阅人员:"
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   5
            Top             =   900
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "借阅病案:"
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   6
            Top             =   1245
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmCISBorrowEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mfrmMain As Object
Private mlngKey As Long
Private mlngReferKey As Long
Private mblnReading As Boolean
Private mstrSQL As String
Private mblnDataChanged As Boolean
Private mblnAllowModify As Boolean
Private mbytMode As Byte
Private mlngMoudal As Long
Private mstrPrivs As String

Private mblnBorrowAccount As Boolean '允许自由录入借阅原因
Private WithEvents mclsPatient As clsVsf
Attribute mclsPatient.VB_VarHelpID = -1

Public Event AfterDataChanged()
Public Event ViewDocument(ByVal strNo As String, ByVal lng病人ID As Long, ByVal lng主页ID As Long)

'######################################################################################################################
Public Property Let AllowModify(blnData As Boolean)
    mblnAllowModify = blnData
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = mblnAllowModify
End Property

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, ByVal lngMoudal As Long, ByVal blnAllowModify As Boolean, ByVal strPrivs As String, ByVal blnBorrowAccount As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Set mfrmMain = frmMain
    mblnAllowModify = blnAllowModify
    mlngMoudal = lngMoudal
    mstrPrivs = strPrivs
    mblnBorrowAccount = blnBorrowAccount
    If ExecuteCommand("初始控件") = False Or ExecuteCommand("初始数据") = False Then Exit Function
    Call ExecuteCommand("控件状态")
        
    DataChanged = False
End Function

Public Function AddPerson() As Boolean
    
    If cmd(0).Enabled And cmd(0).Visible Then
        Call cmd_Click(0)
    End If
    
    AddPerson = True
End Function

Public Function RemovePerson() As Boolean
    
    If cmd(1).Enabled And cmd(1).Visible Then
        Call cmd_Click(1)
    End If
    
    RemovePerson = True
End Function

Public Function AddPatient() As Boolean
    
    If cmd(2).Enabled And cmd(2).Visible Then
        Call cmd_Click(2)
    End If
    
    AddPatient = True
End Function

Public Function RemovePatient() As Boolean
    
    If cmd(3).Enabled And cmd(3).Visible Then
        Call cmd_Click(3)
    End If
    
    RemovePatient = True
    
End Function

Public Function ClearData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    ClearData = ExecuteCommand("清空数据")
End Function

Public Function RefreshData(ByVal lngKey As Long, ByVal blnAllowModify As Boolean, ByVal blnBorrowAccount As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    mbytMode = 2
    mblnBorrowAccount = blnBorrowAccount
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("初始数据")
            
    If ExecuteCommand("读取数据", mlngKey) = False Then Exit Function
    
    Call ExecuteCommand("控件状态")
    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function NewData(Optional ByVal lngReferKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = True
    mlngKey = 0
    mlngReferKey = lngReferKey
    
    mbytMode = 1
   
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("初始数据")
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("缺省数据")

    DataChanged = True
    
    dtp(0).SetFocus
        
    NewData = True
End Function

Public Function Aduit() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mbytMode = 3
    
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("缺省数据")

    DataChanged = True
    
    dtp(2).SetFocus
        
    Aduit = True
End Function

Public Function Revert() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mbytMode = 5
    
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("缺省数据")

    DataChanged = True
    
    txt(1).SetFocus
        
    Revert = True
End Function


Public Function Refuse() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mbytMode = 4
    
    Call ExecuteCommand("控件状态")
    Call ExecuteCommand("缺省数据")

    DataChanged = True
    
    Call LocationObj(txt(0))
        
    Refuse = True
End Function

Public Function ValidData(ByVal blnBorrowReason As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim i As Long
    Select Case mbytMode
    Case 1, 2
        
        If StrIsValid(cbo(0).Text, 255) = False Then
          cbo(0).SetFocus
          Exit Function
        End If
        
        If txtBorrowUser.Text = "" Or txtBorrowUser.Tag = "" Then
            ShowSimpleMsg "病案的借阅人员不能为空值，必须输入！"
            txtBorrowUser.SetFocus
            Exit Function
        End If
        
        With vsf(1)
            If .Rows = 2 And Val(.TextMatrix(1, .ColIndex("病人id"))) = 0 And Val(.TextMatrix(1, .ColIndex("主页id"))) = 0 Then
                ShowSimpleMsg "借阅的病人病案不能为空值，必须输入！"
                mclsPatient.SetFocus
                Exit Function
            End If
        End With
        
        If Format(dtp(1).Value, dtp(1).CustomFormat) < Format(dtp(0).Value, dtp(0).CustomFormat) Then
            ShowSimpleMsg "病案的借阅申请的借阅结束时间不能小于开始时间！"
            dtp(1).SetFocus
            Exit Function
        End If
        
        If DateDiff("d", dtp(0).Value, dtp(1).Value) > Val(GetPara("借阅最长期限", mfrmMain.模块号, "30")) Then
            
            ShowSimpleMsg "病案借阅的最长借阅时间不能超过" & Val(GetPara("借阅最长期限", mfrmMain.模块号, "30")) & "天！"
            dtp(1).SetFocus
            Exit Function

        End If
        
        If blnBorrowReason Then
            If cbo(0).Text = "" Then
                ShowSimpleMsg "请输入病案借阅申请理由!"
                cbo(0).SetFocus
                Exit Function
            End If
        End If
        
    Case 3
        If Format(dtp(3).Value, dtp(3).CustomFormat) < Format(dtp(2).Value, dtp(2).CustomFormat) Then
            ShowSimpleMsg "病案的批准借阅的借阅结束时间不能小于开始时间！"
            dtp(3).SetFocus
            Exit Function
        End If
        
        If DateDiff("d", dtp(2).Value, dtp(3).Value) > Val(GetPara("借阅最长期限", mfrmMain.模块号, "30")) Then
            
            ShowSimpleMsg "病案借阅的最长借阅时间不能超过" & Val(GetPara("借阅最长期限", mfrmMain.模块号, "30")) & "天！"
            dtp(3).SetFocus
            Exit Function

        End If
        
        '检查是否有已经借阅的病案
        With vsf(1)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("病案存储状态")) = "在院" Then
                    ValidData = True
                Else
                    MsgBox "选择的病案:[" & .TextMatrix(i, .ColIndex("姓名")) & "]已经被[" & .TextMatrix(i, .ColIndex("借出申请人")) & "]申请借出,请重新选择!", vbInformation, gstrSysName
                    ValidData = False
                    Exit Function
                End If
            Next
        End With
        
        
        
    Case 4
        If Trim(txt(0).Text) = "" Then
            ShowSimpleMsg "拒绝借阅时间，必须输入拒绝理由！"
            LocationObj txt(0)
            Exit Function
        End If
    Case 5

        
        
    End Select
    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset, ByRef lngKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    Select Case mbytMode
    Case 1, 2
        If mlngKey = 0 Then
            '新增
            lngKey = zlDatabase.GetNextId("病案借阅记录")
            txt(5).Text = zlDatabase.GetNextNo(91)
        Else
            '修改
            lngKey = mlngKey
            
            strSQL = "zl_病案借阅人员_Update(" & lngKey & ",Null)"
            Call SQLRecordAdd(rsSQL, strSQL)
        
            strSQL = "zl_病案借阅内容_Update(" & lngKey & ",Null)"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    
        strSQL = "zl_病案借阅记录_Update(" & lngKey & ",'" & txt(5).Text & "','" & txt(6).Text & "','" & cbo(0).Text & "',To_Date('" & Format(dtp(0).Value, dtp(0).CustomFormat) & " 00:00:00','yyyy-mm-dd hh24:mi:ss'),To_Date('" & Format(dtp(1).Value, dtp(1).CustomFormat) & " 23:59:59','yyyy-mm-dd hh24:mi:ss'),To_Date('" & txt(2).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
                    
        strTmp = ""
        strTmp = txtBorrowUser.Tag
        
        strSQL = "zl_病案借阅人员_Update(" & lngKey & ",'" & strTmp & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        strTmp = ""
        With vsf(1)
            For lngLoop = 1 To .Rows - 1
                If Val(.TextMatrix(lngLoop, .ColIndex("病人id"))) > 0 And Val(.TextMatrix(lngLoop, .ColIndex("主页id"))) > 0 Then
                    If strTmp = "" Then
                        strTmp = Val(.TextMatrix(lngLoop, .ColIndex("病人id"))) & ":" & Val(.TextMatrix(lngLoop, .ColIndex("主页id")))
                    Else
                        strTmp = strTmp & ";" & Val(.TextMatrix(lngLoop, .ColIndex("病人id"))) & ":" & Val(.TextMatrix(lngLoop, .ColIndex("主页id")))
                    End If
                End If
            Next
        End With
        strSQL = "zl_病案借阅内容_Update(" & lngKey & ",'" & strTmp & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case 3
        strSQL = "zl_病案借阅记录_Authorize(" & lngKey & ",To_Date('" & Format(dtp(2).Value, dtp(2).CustomFormat) & " 00:00:00','yyyy-mm-dd hh24:mi:ss'),To_Date('" & Format(dtp(3).Value, dtp(3).CustomFormat) & " 23:59:59','yyyy-mm-dd hh24:mi:ss'),'" & txt(7).Text & "',To_Date('" & txt(3).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case 4
        strSQL = "zl_病案借阅记录_Refuse(" & lngKey & ",'" & txt(8).Text & "','" & txt(0).Text & "',To_Date('" & txt(4).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case 5
        strSQL = "zl_病案借阅记录_Revert(" & lngKey & ",'" & txt(1).Text & "',To_Date('" & txt(9).Text & ":00','yyyy-mm-dd hh24:mi:ss'))"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    End Select
    
    SaveData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'######################################################################################################################
Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim lngNum As Long
        
    On Error GoTo errHand
    
    mblnReading = True
    Call SQLRecord(rsSQL)
    
    Select Case strCmd
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"

        Set mclsPatient = New clsVsf
        With mclsPatient
            Call .Initialize(Me.Controls, vsf(1), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            If AllowModify Then
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
            Else
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            End If
            Call .AppendColumn("姓名", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("性别", 600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("年龄", 600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("婚姻状况", 900, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("住院号", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病案号", 810, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("住院次数", 810, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("入院时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("出院时间", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm", , True)
            Call .AppendColumn("出院科室", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("借出状态", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("借出申请人", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("病案存储状态", 0, flexAlignLeftCenter, flexDTString, "", , True)
            
            
            If AllowModify Then
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("姓名"), True, vbVsfEditCommand)
                .IndicatorCol = 0
                Set .IndicatorIcon = GetImageList(16).ListImages("当前").Picture
            End If
            .AppendRows = True
        End With
                
        '借阅理由
        '----------------------------------------------------------------------------------------------------------
        cbo(0).Clear
        cbo(0).AddItem ""
        Set rs = gclsPackage.GetDictTableData("借阅理由")
        If rs.BOF = False Then
            Do While Not rs.EOF
                cbo(0).AddItem rs("名称").Value
                If rs("缺省标志").Value = 1 Then cbo(0).ListIndex = cbo(0).NewIndex
                rs.MoveNext
            Loop
        End If
        If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        blnAllowModify = mblnAllowModify
        If (mlngKey = 0 And mbytMode = 2) Or lbl(15).Caption <> "" Then blnAllowModify = False
        
        cmd(0).Enabled = blnAllowModify
        cmd(2).Enabled = blnAllowModify
        cmd(3).Enabled = blnAllowModify
        
        Select Case mbytMode
        Case 1, 2
            txt(0).Locked = Not blnAllowModify
            cbo(0).Locked = Not blnAllowModify
            txt(2).Locked = True
            txt(3).Locked = True
            txt(4).Locked = True
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = True
            txt(8).Locked = True
            txt(1).Locked = True
            txt(9).Locked = True
            
            dtp(0).Enabled = blnAllowModify
            dtp(1).Enabled = blnAllowModify
            dtp(2).Enabled = blnAllowModify
            dtp(3).Enabled = blnAllowModify
            
            If blnAllowModify Then
                txtBorrowUser.Enabled = True
                Call mclsPatient.InitializeEdit(True, True, True)
            Else
                txtBorrowUser.Enabled = False
                Call mclsPatient.InitializeEdit(False, False, False)
            End If
        Case 3          '批准
            txt(0).Locked = True
            cbo(0).Locked = True
            txt(2).Locked = True
            txt(3).Locked = False
            txt(4).Locked = True
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = False
            txt(8).Locked = True
            txt(1).Locked = True
            txt(9).Locked = True
            
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(2).Enabled = True
            dtp(3).Enabled = True
            txtBorrowUser.Enabled = False
            Call mclsPatient.InitializeEdit(False, False, False)
        Case 4          '拒借
            txt(0).Locked = False
            cbo(0).Locked = True
            txt(2).Locked = True
            txt(3).Locked = True
            txt(4).Locked = False
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = True
            txt(8).Locked = False
            txt(1).Locked = True
            txt(9).Locked = True
            
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(2).Enabled = False
            dtp(3).Enabled = False
            txtBorrowUser.Enabled = False
            Call mclsPatient.InitializeEdit(False, False, False)
        Case 5          '归还
            txt(0).Locked = True
            cbo(0).Locked = True
            txt(2).Locked = True
            txt(3).Locked = True
            txt(4).Locked = True
            txt(5).Locked = True
            txt(6).Locked = True
            txt(7).Locked = True
            txt(8).Locked = True
            txt(1).Locked = False
            txt(9).Locked = False
            
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(2).Enabled = False
            dtp(3).Enabled = False
            txtBorrowUser.Enabled = False
            Call mclsPatient.InitializeEdit(False, False, False)
        End Select
            
        For lngNum = 0 To 9
            If txt(lngNum).Locked Then
                txt(lngNum).Enabled = False
            Else
                txt(lngNum).Enabled = True
            End If
        Next
    '------------------------------------------------------------------------------------------------------------------
    Case "汇总信息"
        
        With vsf(1)
            If Val(.RowData(.Rows - 1)) > 0 Then
                lbl(16).Caption = "借阅的病案共有 " & .Rows - 1 & " 份"
            Else
                lbl(16).Caption = "借阅的病案共有 " & .Rows - 2 & " 份"
            End If
        End With
            
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"

        txt(0).MaxLength = GetMaxLength("病案借阅记录", "申请理由")
                    
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        ExecuteCommand = ExecuteCommand("读取数据", Val(varParam(0)))
        GoTo endHand
        
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
                
        txt(0).Text = ""
        cbo(0).Text = ""
        txt(2).Text = ""
        txt(3).Text = ""
        txt(4).Text = ""
        txt(5).Text = ""
        txt(6).Text = ""
        txt(7).Text = ""
        txt(8).Text = ""
        txt(1).Text = ""
        txt(9).Text = ""
        lbl(15).Caption = ""
        txtConver(0).Visible = True
        txtConver(1).Visible = True
        txtConver(2).Visible = True
        txtConver(3).Visible = True
        dtp(0).Enabled = False
        dtp(1).Enabled = False
        dtp(2).Enabled = False
        dtp(3).Enabled = False
        txtBorrowUser.Text = ""
        txtBorrowUser.Tag = ""
        mclsPatient.ClearGrid
        
        Call ExecuteCommand("汇总信息")
    '------------------------------------------------------------------------------------------------------------------
    Case "缺省数据"
        
        Select Case mbytMode
        Case 1, 2
            
            dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
            
            If Val(GetPara("病案借阅期限", mfrmMain.模块号, "7")) = 0 Then
                dtp(1).Value = Format(zlDatabase.Currentdate + 8, dtp(1).CustomFormat)
            Else
                dtp(1).Value = Format(zlDatabase.Currentdate + 1 + Val(GetPara("病案借阅期限", mfrmMain.模块号, "7")), dtp(1).CustomFormat)
            End If

            txtConver(0).Visible = False
            txtConver(1).Visible = False
            dtp(0).Enabled = True
            dtp(1).Enabled = True
            txt(6).Text = UserInfo.姓名
            txt(2).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        Case 3
            txt(7).Text = UserInfo.姓名
            txt(3).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            dtp(2).Value = dtp(0).Value
            dtp(3).Value = dtp(1).Value
            txtConver(2).Visible = False
            txtConver(3).Visible = False
        Case 4
            txt(8).Text = UserInfo.姓名
            txt(4).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        Case 5
            txt(1).Text = UserInfo.姓名
            txt(9).Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        
        Call ExecuteCommand("清空数据")
        mblnReading = True
        
        If Val(varParam(0)) > 0 Then
            Set rs = gclsPackage.GetBorrow(1, Val(varParam(0)))
            If rs.BOF = False Then
                txt(5).Text = zlCommFun.NVL(rs("No").Value)
                cbo(0).Text = zlCommFun.NVL(rs("申请理由").Value)
                txt(0).Text = zlCommFun.NVL(rs("拒借理由").Value)
                            
                txtConver(0).Visible = IsNull(rs("申请时间").Value)
                txtConver(1).Visible = IsNull(rs("申请期限").Value)
                txtConver(2).Visible = IsNull(rs("借阅时间").Value)
                txtConver(3).Visible = IsNull(rs("借阅期限").Value)
                
                If IsNull(rs("申请时间").Value) = False Then dtp(0).Value = Format(rs("申请时间").Value, dtp(0).CustomFormat)
                If IsNull(rs("申请期限").Value) = False Then dtp(1).Value = Format(rs("申请期限").Value, dtp(1).CustomFormat)
                If IsNull(rs("借阅时间").Value) = False Then dtp(2).Value = Format(rs("借阅时间").Value, dtp(2).CustomFormat)
                If IsNull(rs("借阅期限").Value) = False Then dtp(3).Value = Format(rs("借阅期限").Value, dtp(3).CustomFormat)
                
                txt(6).Text = zlCommFun.NVL(rs("申请人").Value)
                txt(7).Text = zlCommFun.NVL(rs("批准人").Value)
                txt(8).Text = zlCommFun.NVL(rs("拒借人").Value)
                txt(1).Text = zlCommFun.NVL(rs("收回人").Value)
                 
                If IsNull(rs("登记时间").Value) = False Then txt(2).Text = Format(rs("登记时间").Value, "yyyy-MM-dd HH:mm")
                If IsNull(rs("批准时间").Value) = False Then txt(3).Text = Format(rs("批准时间").Value, "yyyy-MM-dd HH:mm")
                If IsNull(rs("拒借时间").Value) = False Then txt(4).Text = Format(rs("拒借时间").Value, "yyyy-MM-dd HH:mm")
                If IsNull(rs("归还时间").Value) = False Then txt(9).Text = Format(rs("归还时间").Value, "yyyy-MM-dd HH:mm")
                
                dtp(0).Enabled = Not txtConver(0).Visible
                dtp(1).Enabled = Not txtConver(1).Visible
                dtp(2).Enabled = Not txtConver(2).Visible
                dtp(3).Enabled = Not txtConver(3).Visible
                
                Select Case rs("记录状态").Value
                Case 1
                    lbl(15).Caption = ""
                Case 2
                    lbl(15).Caption = "已批准"
                Case 3
                    lbl(15).Caption = "已拒绝"
                Case 4
                    lbl(15).Caption = "已归还"
                End Select
            End If
        End If
        
        If lbl(15).Caption = "" And mlngKey > 0 Then
            Call mclsPatient.ModifyColumn(0, "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
            
        Else
            Call mclsPatient.ModifyColumn(0, "", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        End If
    
        If Val(varParam(0)) > 0 Then
               
            Set rs = gclsPackage.GetBorrowPerson(Val(varParam(0)))
            If rs.BOF = False Then
                txtBorrowUser.Text = zlCommFun.NVL(rs!姓名)
                txtBorrowUser.Tag = zlCommFun.NVL(rs!ID, 0)
            End If
            
            Set rs = gclsPackage.GetBorrowPatient(Val(varParam(0)))
            If rs.BOF = False Then
                Call mclsPatient.LoadGrid(rs)
            End If
        End If
        
        Call ExecuteCommand("汇总信息")
        
    End Select

    ExecuteCommand = True
    
    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
    mblnReading = False
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If mblnBorrowAccount = False Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    Select Case Index
    Case 0

        Set rsData = gclsPackage.GetOperationPerson
        bytRet = ShowPubSelect(Me, txtBorrowUser, 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,", Me.Name & "\借阅人员选择", "请从下表中选择一个或多个借阅人员", rsData, rs, 8790, 4500, False, txtBorrowUser.Tag)
                    
        If bytRet = 1 Then
            If rs.RecordCount = 1 Then
                txtBorrowUser.Text = zlCommFun.NVL(rs("姓名").Value)
                txtBorrowUser.Tag = zlCommFun.NVL(rs("ID").Value, 0)
            End If
            DataChanged = True
        End If
   
    Case 1
        
'        lngLoop = vsf(0).Row
'        Call mclsPerson.DeleteRow(vsf(0).Row)
'
'        If lngLoop <= vsf(0).Rows - 1 Then
'            vsf(0).Row = lngLoop
'        Else
'            vsf(0).Row = vsf(0).Rows - 1
'        End If
        
    Case 2
    
        If frmSearchPatient.ShowEdit(Me, rs, mlngMoudal, mstrPrivs) Then
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                With vsf(1)
                    Do While Not rs.EOF
                                                                                
                        If mclsPatient.CheckHave(rs("ID").Value, False) = False Then
                            If Trim(.RowData(.Rows - 1)) <> "" And Trim(.RowData(.Rows - 1)) <> "0" Then .Rows = .Rows + 1
                            .RowData(.Rows - 1) = Trim(rs("ID").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("病人id")) = Val(rs("病人id").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("主页id")) = Val(rs("主页id").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("姓名")) = Trim(rs("姓名").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("性别")) = Trim(rs("性别").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("年龄")) = Trim(rs("年龄").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("婚姻状况")) = Trim(rs("婚姻状况").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("入院时间")) = Trim(rs("入院时间").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("出院时间")) = Trim(rs("出院时间").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("出院科室")) = Trim(rs("出院科室").Value)
                            
                            .TextMatrix(.Rows - 1, .ColIndex("住院号")) = Trim(rs("住院号").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("病案号")) = Trim(rs("病案号").Value)
                            .TextMatrix(.Rows - 1, .ColIndex("住院次数")) = Trim(rs("住院次数").Value)
                            
                            DataChanged = True
                        End If
                        
                        rs.MoveNext
                    Loop
                End With
            End If
        End If
        
    Case 3
                
        lngLoop = vsf(1).Row
        Call mclsPatient.DeleteRow(vsf(1).Row)
        
        If lngLoop <= vsf(1).Rows - 1 Then
            vsf(1).Row = lngLoop
        Else
            vsf(1).Row = vsf(1).Rows - 1
        End If
        
    End Select
    
    Call ExecuteCommand("汇总信息")
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    picPane(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    fra.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    txt(5).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsPatient = Nothing
End Sub

Private Sub mclsPatient_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call ExecuteCommand("汇总信息")
    DataChanged = True
End Sub

Private Sub mclsPatient_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf(1).RowData(Row)) = 0)
End Sub

Private Sub mclsPerson_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    Call ExecuteCommand("汇总信息")
    DataChanged = True
End Sub

Private Sub mclsPerson_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf(0).RowData(Row)) = 0)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        fra.Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        
        cbo(0).Move cbo(0).Left, cbo(0).Top, fra.Width - cbo(0).Left - 45
'        vsf(0).Move vsf(0).Left, vsf(0).Top, fra.Width - vsf(0).Left - 45
        
        txtBorrowUser.Move txtBorrowUser.Left, txtBorrowUser.Top, fra.Width - txtBorrowUser.Left - 45 - cmd(0).Width
        cmd(0).Move txtBorrowUser.Left + txtBorrowUser.Width + 15, txtBorrowUser.Top
        
        vsf(1).Move txtBorrowUser.Left, vsf(1).Top, txtBorrowUser.Width, fra.Height - vsf(1).Top - (picPane(1).Height + 45) - 75
        cmd(2).Move vsf(1).Left + vsf(1).Width + 15, vsf(1).Top
        cmd(3).Move cmd(2).Left, cmd(2).Top + cmd(2).Height + 15
        
        
        picPane(1).Move 30, vsf(1).Top + vsf(1).Height + 45, fra.Width - 60
        
        lbl(15).Move fra.Width - 900
        
        mclsPatient.AppendRows = True
    Case 1
        txt(0).Move txt(0).Left, txt(0).Top, picPane(Index).Width - txt(0).Left - 45
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
End Sub

Private Sub txtBorrowUser_KeyPress(KeyAscii As Integer)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim bytRet As Byte
    
    If KeyAscii = vbKeyReturn Then
        Set rsData = gclsPackage.GetOperationPerson(UCase(txtBorrowUser.Text))

         If ShowPubSelect(Me, txtBorrowUser, 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,", Me.Name & "\借阅人员过滤", "请从下表中选择一个借阅人员", rsData, rs, 8790, 4500, , txtBorrowUser.Tag, , True) = 1 Then

             txtBorrowUser.Text = zlCommFun.NVL(rs("姓名").Value)
             txtBorrowUser.Tag = zlCommFun.NVL(rs("ID").Value, 0)
             DataChanged = True
         End If
    End If
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    '编辑处理
    Select Case Index
    Case 0
        
    Case 1
        Call mclsPatient.AfterEdit(Row, Col)
    End Select
    
    DataChanged = True
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '编辑处理
    Select Case Index
    Case 0
 
    Case 1
        Call mclsPatient.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Select Case Index
    Case 0

    Case 1
        mclsPatient.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
    Case 0
'        mclsPerson.AppendRows = True
    Case 1
        mclsPatient.AppendRows = True
    End Select
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Index
    Case 0
'        Call mclsPerson.BeforeResizeColumn(Col, Cancel)
    Case 1
        Call mclsPatient.BeforeResizeColumn(Col, Cancel)
    End Select
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsf(Index)
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
'            If Col = .ColIndex("姓名") Then
'
'                Set rsData = gclsPackage.GetOperationPerson
'                bytRet = ShowPubSelect(Me, vsf(Index), 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,", Me.Name & "\借阅人员选择", "请从下表中选择一个或多个借阅人员", rsData, rs, 8790, 4500, True, Val(.RowData(Row)))
'
'                If bytRet = 1 Then
'
'                    For lngLoop = 1 To rs.RecordCount
'
'                        If mclsPerson.CheckHave(zlCommFun.NVL(rs("ID").Value), False) = False Then
'
'                            If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
'
'                            .EditText = zlCommFun.NVL(rs("姓名").Value)
'                            .TextMatrix(.Rows - 1, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
'                            .TextMatrix(.Rows - 1, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
'                            .TextMatrix(.Rows - 1, .ColIndex("科室")) = zlCommFun.NVL(rs("科室").Value)
'                            .RowData(.Rows - 1) = zlCommFun.NVL(rs("ID").Value, 0)
'
'                            DataChanged = True
'                        End If
'
'                        rs.MoveNext
'                    Next
'
'                    mclsPerson.AppendRows = True
'
'                    DataChanged = True
'
'                End If
'
'            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 1
            If frmSearchPatient.ShowEdit(Me, rs, mlngMoudal, mstrPrivs) Then
                If rs.RecordCount > 0 Then
                    rs.MoveFirst
                    With vsf(1)
                        Do While Not rs.EOF
                                                                                    
                            If mclsPatient.CheckHave(rs("ID").Value, False) = False Then
                                If Trim(.RowData(.Rows - 1)) <> "" And Trim(.RowData(.Rows - 1)) <> "0" Then .Rows = .Rows + 1
                                .RowData(.Rows - 1) = Trim(rs("ID").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("病人id")) = Val(rs("病人id").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("主页id")) = Val(rs("主页id").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("姓名")) = Trim(rs("姓名").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("性别")) = Trim(rs("性别").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("年龄")) = Trim(rs("年龄").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("婚姻状况")) = Trim(rs("婚姻状况").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("入院时间")) = Trim(rs("入院时间").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("出院时间")) = Trim(rs("出院时间").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("出院科室")) = Trim(rs("出院科室").Value)
                                
                                .TextMatrix(.Rows - 1, .ColIndex("住院号")) = Trim(rs("住院号").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("病案号")) = Trim(rs("病案号").Value)
                                .TextMatrix(.Rows - 1, .ColIndex("住院次数")) = Trim(rs("住院次数").Value)
                                
                                DataChanged = True
                            End If
                            
                            rs.MoveNext
                        Loop
                    End With
                End If
            End If
        End Select
    End With
    Call ExecuteCommand("汇总信息")
End Sub

Private Sub vsf_DblClick(Index As Integer)
    '编辑处理
    Select Case Index
    Case 0
'        Call mclsPerson.DbClick
    Case 1
        Call mclsPatient.DbClick
        
        If lbl(15).Caption = "已批准" And DataChanged = False Then
            With vsf(1)
                RaiseEvent ViewDocument(txt(5).Text, Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))))
            End With
        End If
        
    End Select
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '编辑处理
    Select Case Index
    Case 0
'        Call mclsPerson.KeyDown(KeyCode, Shift)
    Case 1
        Call mclsPatient.KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    
    'ToDo...
    If KeyAscii = vbKeyReturn Then Call vsf_DblClick(Index)
    
    '编辑处理,最后调用
    Select Case Index
    Case 0
'        Call mclsPerson.KeyPress(KeyAscii)
    Case 1
        Call mclsPatient.KeyPress(KeyAscii)
    End Select
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Select Case Index
    Case 0
'        Call mclsPerson.KeyPressEdit(KeyAscii)
    Case 1
        Call mclsPatient.KeyPressEdit(KeyAscii)
    End Select
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim StrText As String
    Dim bytRet As Byte
    Dim blnCard As Boolean
    Dim bytFilterMode As Byte
    
    With vsf(Index)
        
        If InStr(.EditText, "'") > 0 Then
            KeyCode = 0
            .EditText = ""
            Exit Sub
        End If
                            
        StrText = .EditText
        
        Select Case Index
        '----------------------------------------------------------------------------------------------------------
        Case 0
'            If KeyCode = vbKeyReturn Then
'                If Col = .ColIndex("姓名") Then
'
'                    Set rsData = gclsPackage.GetOperationPerson(UCase(StrText))
'
'                    If ShowPubSelect(Me, vsf(Index), 2, "编号,1200,0,;姓名,1200,0,;简码,900,0,;科室,1200,0,", Me.Name & "\借阅人员过滤", "请从下表中选择一个借阅人员", rsData, rs, 8790, 4500, , Val(.RowData(Row)), , True) = 1 Then
'
'                        If mclsPerson.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                            ShowSimpleMsg "选择的人员“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
'                            Exit Sub
'                        End If
'
'                        .EditText = zlCommFun.NVL(rs("姓名").Value)
'                        .Cell(flexcpData, Row, Col) = zlCommFun.NVL(rs("姓名").Value)
'                        .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
'                        .TextMatrix(Row, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
'                        .TextMatrix(Row, .ColIndex("科室")) = zlCommFun.NVL(rs("科室").Value)
'                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
'
'                        DataChanged = True
'                    Else
'                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
'                        .EditText = .Cell(flexcpData, Row, Col)
'                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
'                    End If
'
'                End If
'            Else
'                DataChanged = True
'            End If
        '----------------------------------------------------------------------------------------------------------
        Case 1
        
            If Col = .ColIndex("姓名") Then

                If KeyCode <> 8 And KeyCode <> 13 Then StrText = StrText & Chr(KeyCode)

                '检查非法字符
                If InStr(StrText, "'") > 0 Then
                    KeyCode = 0
                    ShowSimpleMsg "在个人姓名中有非法字符 ' ！"
                    .EditText = ""
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    Exit Sub
                End If

                '检查是否为就诊卡号码
                blnCard = VsfInputIsCard(vsf(Index), KeyCode, ParamInfo.系统号)
                If blnCard And Len(.EditText) = ParamInfo.就诊卡号码长度 - 1 And KeyCode <> 8 And KeyCode <> vbKeyReturn Then
                    .EditSelStart = Len(.EditText)
                    bytFilterMode = 1
                End If
            End If

            If KeyCode = vbKeyReturn Then
                If Col = .ColIndex("姓名") Then
                    If blnCard Then
                        '是就诊卡
                        bytFilterMode = 1
                    Else
                        '非就诊卡
                        blnCard = False
                        StrText = .EditText
                        
                        Select Case UCase(Left(StrText, 1))
                        Case "-", "A"                   '病人id
                            bytFilterMode = 2
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "+", "B"                   '住院号
                            bytFilterMode = 3
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "*", "D"                   '门诊号
                            bytFilterMode = 4
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "/", "C"                   '当前床号
                            bytFilterMode = 5
                            StrText = IIf(IsNumeric(Mid(StrText, 2)), Val(Mid(StrText, 2)), "0")
                        Case "F"                        '病案号
                            bytFilterMode = 7
                            StrText = Mid(StrText, 2)
                        Case Else                       '姓名
                            bytFilterMode = 6
                        End Select
                        
                    End If
                End If
            End If
            
            If Col = .ColIndex("姓名") Then
                
                If bytFilterMode > 0 Then
                    
                    Set rsData = gclsPackage.GetPatient(bytFilterMode, StrText)
                    
                    If rsData.RecordCount > 0 Then
                        If rsData.RecordCount = 1 Then
                            bytRet = 1
                            Set rs = rsData
                        Else
                            bytRet = ShowPubSelect(Me, vsf(Index), 2, "姓名,1200,0,0;性别,810,0,0;入院时间,1667,0,0;出院时间,1667,0,0;身份证号,1500,0,0", mfrmMain.Name & "\病人过滤选择", "请从下面选择一个病人", rsData, rs, 8790, 4500)
                        End If
                        
                        If bytRet = 1 Then
                        
                            If mclsPatient.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "选择的病人“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
                                
                                '还原原来的内容
                                .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                                .EditText = .Cell(flexcpData, Row, Col)
                                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                                Exit Sub
                            End If
                                      
                            .EditText = zlCommFun.NVL(rs("姓名"))
                            .Cell(flexcpData, Row, Col) = zlCommFun.NVL(rs("姓名"))
                            .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
                            .TextMatrix(Row, .ColIndex("性别")) = zlCommFun.NVL(rs("性别").Value)
                            .TextMatrix(Row, .ColIndex("年龄")) = zlCommFun.NVL(rs("年龄").Value)
                            .TextMatrix(Row, .ColIndex("婚姻状况")) = zlCommFun.NVL(rs("婚姻状况").Value)
                            .TextMatrix(Row, .ColIndex("入院时间")) = zlCommFun.NVL(rs("入院时间").Value)
                            .TextMatrix(Row, .ColIndex("出院时间")) = zlCommFun.NVL(rs("出院时间").Value)
                            .TextMatrix(Row, .ColIndex("出院科室")) = zlCommFun.NVL(rs("出院科室").Value)
                            .TextMatrix(Row, .ColIndex("病人id")) = zlCommFun.NVL(rs("病人id").Value)
                            .TextMatrix(Row, .ColIndex("主页id")) = zlCommFun.NVL(rs("主页id").Value)
                            DataChanged = True
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                            If blnCard Then
                                .Cell(flexcpData, Row, Col) = StrText
                                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                                KeyCode = 13
                            End If
                        Else
                            '还原原来的内容
                            .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                            .EditText = .Cell(flexcpData, Row, Col)
                            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                        End If
                    Else
                        '还原原来的内容
                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    End If
                End If
            
            End If
            
        End Select
    End With
    
    Call ExecuteCommand("汇总信息")
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case Index
        Case 0
'            Call mclsPerson.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 1
            Call mclsPatient.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        End Select
    End Select
End Sub

Private Sub vsf_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsf(Index)
        
        If .MouseCol = .ColIndex("姓名") And Index = 1 And (mbytMode = 1 Or mbytMode = 2) Then
            If .ToolTipText = "" Then .ToolTipText = "在姓名处查找病人的方法：1.'-'或'A'+病人id;2.'+'或'B'+住院号;3.'/'或'C'+床号;4.'*'或'D'+门诊号;5.其他按姓名查找"
        Else
            If .ToolTipText <> "" Then .ToolTipText = ""
        End If
    End With
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Select Case Index
    Case 0
'        Call mclsPerson.EditSelAll
    Case 1
        Call mclsPatient.EditSelAll
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Select Case Index
    Case 0
'        Call mclsPerson.BeforeEdit(Row, Col, Cancel)
    Case 1
        Call mclsPatient.BeforeEdit(Row, Col, Cancel)
    End Select
End Sub

