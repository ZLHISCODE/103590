VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReInvoice 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "票据收回选择"
   ClientHeight    =   4035
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
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
      Height          =   420
      Left            =   3715
      TabIndex        =   8
      Top             =   3495
      Width           =   1400
   End
   Begin VB.Frame fraTop 
      Height          =   60
      Left            =   0
      TabIndex        =   7
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   1200
      Width           =   1755
   End
   Begin VB.TextBox txtThis 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0.00"
      ToolTipText     =   "仅当改变缺省结算方式的金额时才产生"
      Top             =   2160
      Width           =   1755
   End
   Begin VB.CommandButton cmdOk 
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
      Height          =   420
      Left            =   3715
      TabIndex        =   1
      Top             =   3000
      Width           =   1400
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshInvoice 
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   5450
      _Version        =   393216
      Rows            =   5
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^ 选择|^    票据号      "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblPrompt 
      Caption         =   "请根据本次退费合计和实际收到的退票金额合计,选择对应的收回票据号,全选表示全部票据收回重打。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "可退总金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label lblMargin 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "本次退费金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   1440
   End
End
Attribute VB_Name = "frmReInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mstrInvoices As String
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mblnSelAll As Boolean

Public Function ShowMe(frmParent As Object, ByVal strNO As String, _
    ByVal cur可退金额 As Currency, _
    ByVal cur本次退款 As Currency, _
    ByRef strInvoices As String, _
    ByRef blnSelAll As Boolean, _
    Optional ByVal int票种 As Integer = 4) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进入退费发票号选择
    '入参:frmParent-调用的父窗口
    '     strNO-退费的单据号
    '     cur可退金额-可退的金额
    '     cur本次退款-本次退款额
    '     bln补结算-是否补结算
    '出参:strInVoices-退费选择的发票号
    '返回:点击确定,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-13 10:02:34
    '问题:27352
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Long

    mstrInvoices = ""
    mblnChange = False
    If Mid(strNO, 1, 1) = "," Then strNO = Mid(strNO, 2)
    Set rsTmp = GetInvoice(strNO, int票种)
    
    If rsTmp.RecordCount = 0 Then
        strInvoices = ""
        '对未打印发票的,直接返回true.
        ShowMe = True: Unload Me
        Exit Function
    End If
    
    If rsTmp.RecordCount = 1 Then
        '只有一张时,收回重打
        strInvoices = rsTmp!号码
        ShowMe = True
        blnSelAll = True
        Unload Me: Exit Function
    End If

    With mshInvoice
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 0) = "√"
            .TextMatrix(i, 1) = rsTmp!号码
            rsTmp.MoveNext
        Next
    End With
    txtTotal.Text = Format(cur可退金额, "0.00")
    txtThis.Text = Format(cur本次退款, "0.00")
    
    Me.Show 1, frmParent
    strInvoices = mstrInvoices
    blnSelAll = mblnSelAll
    ShowMe = mblnOk
End Function


Private Function GetInvoice(ByVal strNos As String, ByVal int票种 As Integer) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定单据所对应的发票使用集
    '返回:返回满足条件的单据发票
    '编制:刘兴洪
    '日期:2014-10-10 17:58:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errH

    strSQL = _
    "   Select A.号码" & vbNewLine & _
    "   From 票据使用明细 A" & vbNewLine & _
    "   Where A.性质 = 1 And a.原因 <> 6 " & vbNewLine & _
    "           And A.打印id = (Select Max(ID) From 票据打印内容 Where 数据性质 = [2] And NO = [1])" & vbNewLine & _
    "Minus" & vbNewLine & _
    "Select A.号码" & vbNewLine & _
    "From 票据使用明细 A" & vbNewLine & _
    "Where A.性质 = 2 And a.原因 <> 6 " & vbNewLine & _
    "   And A.打印id = (Select Max(ID) From 票据打印内容 Where 数据性质 = [2] And NO = [1])" & vbNewLine & _
    "Order By 号码"

    Set GetInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos, int票种)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim i As Long
    
    With mshInvoice
        
        If .Rows > 1 Then
            mstrInvoices = ""
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = "√" Then
                    mstrInvoices = mstrInvoices & "," & Trim(.TextMatrix(i, 1))
                End If
            Next
            If mstrInvoices = "" Then
                MsgBox "请至少选择一张票据!", vbInformation, gstrSysName
                Exit Sub
            End If
            mstrInvoices = Mid(mstrInvoices, 2)
            
            If .Rows - 1 = UBound(Split(mstrInvoices, ",")) + 1 Then
                If MsgBox("你确定要收回所有票据进行重打操作吗?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                mblnSelAll = True
            Else
                If MsgBox("共" & .Rows - 1 & "张票据,你选择了收回" & UBound(Split(mstrInvoices, ",")) + 1 & "张." & vbCrLf & _
                    "你确定要收回这些票据吗?" & vbCrLf & mstrInvoices, vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End With
    mblnChange = False
    mblnOk = True
    Unload Me
End Sub

Private Sub cmd取消_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub Form_Load()
    mblnChange = False
    mblnSelAll = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("你进行了相关票据选择的，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
End Sub

Private Sub mshInvoice_DblClick()
    Dim i As Long
    
    With mshInvoice
        If .Col = 0 Then
            If .Row = 0 Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 0) = IIf(.TextMatrix(i, 0) = "", "√", "")
                Next
            Else
                 .TextMatrix(.Row, 0) = IIf(.TextMatrix(.Row, 0) = "", "√", "")
            End If
            mblnChange = True
        End If
    End With
End Sub
