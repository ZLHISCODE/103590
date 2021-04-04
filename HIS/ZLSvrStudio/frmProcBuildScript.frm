VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcBuildScript 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "生成补充脚本"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   Icon            =   "frmProcBuildScript.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5340
      Index           =   0
      Left            =   120
      ScaleHeight     =   5340
      ScaleWidth      =   10380
      TabIndex        =   3
      Top             =   960
      Width           =   10380
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1140
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   105
         Width           =   1935
         _cx             =   3413
         _cy             =   2011
         Appearance      =   1
         BorderStyle     =   0
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "生成(&O)"
      Height          =   350
      Left            =   8145
      TabIndex        =   2
      Top             =   6420
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   9360
      TabIndex        =   1
      Top             =   6420
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   165
      Picture         =   "frmProcBuildScript.frx":6852
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   0
      Top             =   60
      Width           =   720
   End
   Begin MSComctlLib.ProgressBar pbr 
      Height          =   105
      Left            =   45
      TabIndex        =   5
      Top             =   6840
      Visible         =   0   'False
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "lbl"
      Height          =   180
      Left            =   60
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "根据不同的所有者，生成对应所有者的自定义过程到指定路径"
      Height          =   180
      Left            =   1320
      TabIndex        =   7
      Top             =   570
      Width           =   4860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "生成补充脚本"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmProcBuildScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjMain As Object
Private mclsVsf As clsVsf

Public Function ShowMe(ByVal objMain As Object) As Boolean
    On Error GoTo errHand
        
    Set mobjMain = objMain
    If ExecuteCommand("初始数据") = False Then Exit Function
    If ExecuteCommand("刷新数据") = False Then Exit Function
    Me.Show 1, mobjMain
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Sub cmdOK_Click()
    On Error GoTo errHand
    Dim rs As ADODB.Recordset
    Dim lngRow As Long
    Dim strFlag As String
    Dim lngKey As Long
    Dim objFileTemp As TextStream
    
    cmdOK.Enabled = False
    
    With vsf(0)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("所有者")) <> "" Then
                Set rs = gclsBase.GetProcByOwner(.TextMatrix(lngRow, .ColIndex("所有者")))
                If rs.BOF = False Then
                    lblTitle.Visible = True
                    pbr.Visible = True
                    pbr.Max = rs.RecordCount
                    pbr.value = 0
                    lblTitle.Caption = "正在生成过程脚本"
                    Set objFileTemp = gobjFile.CreateTextFile(.TextMatrix(lngRow, .ColIndex("脚本路径")), True)
                    Do While Not rs.EOF
                        If lngKey <> 0 And lngKey <> Nvl(rs("ID").value) Then
                            objFileTemp.Write vbCrLf & "/" & vbCrLf
                        End If
                        strFlag = Nvl(rs("内容").value)
                        objFileTemp.Write strFlag
                        lngKey = Nvl(rs("ID").value)
                        pbr.value = pbr.value + 1
                        rs.MoveNext
                    Loop
                    If strFlag <> "" Then
                        objFileTemp.Write vbCrLf & "/"
                    End If
'                    lblTitle.Caption = "生成成功！"
                    pbr.value = 0
                    pbr.Visible = False
                    lblTitle.Visible = False
                End If
            End If
        Next
    End With
    
    MsgBox "脚本已经生成成功！", vbInformation, Me.Caption
    cmdOK.Enabled = True
    
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
    cmdOK.Enabled = True
End Sub

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    Dim objItem As Object
    Dim intRow As Integer
    
    On Error GoTo errHand
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True)
            Call .ClearColumn
            Call .AppendColumn("所有者", 800, flexAlignLeftCenter, flexDTString, , "", False)
            Call .AppendColumn("脚本路径", 2000, flexAlignLeftCenter, flexDTString, , "", True)
            
            Call .InitializeEdit(True, False, False)
            Call .InitializeEditColumn(.ColIndex("脚本路径"), True, vbVsfEditCommand)

            .AppendRows = True
        End With
    Case "刷新数据"
        strSQL = "Select Distinct A.所有者,'" & App.Path & "\' || A.所有者 || 'Procedure.sql' As 脚本路径 From zlProcedure A,zlSystems B Where A.所有者 = B.所有者"
        Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "")
        If rs.BOF = False Then
            vsf(0).Rows = rs.RecordCount + 1
            For intRow = 0 To rs.RecordCount - 1
                vsf(0).TextMatrix(intRow + 1, vsf(0).ColIndex("所有者")) = Nvl(rs("所有者").value)
                vsf(0).TextMatrix(intRow + 1, vsf(0).ColIndex("脚本路径")) = Nvl(rs("脚本路径").value)
                rs.MoveNext
            Next
        End If
    End Select
    ExecuteCommand = True
    Exit Function
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    vsf(0).Move 15, 15, picPane(0).ScaleWidth - 30, picPane(0).ScaleHeight - 30
    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub



