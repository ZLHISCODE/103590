VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPathTrackView 
   AutoRedraw      =   -1  'True
   Caption         =   "病人临床路径"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10860
   Icon            =   "frmPathTrackView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   10860
   StartUpPosition =   3  '窗口缺省
   Begin XtremeSuiteControls.TabControl tbcPath 
      Height          =   3090
      Left            =   240
      TabIndex        =   0
      Top             =   255
      Width           =   5475
      _Version        =   589884
      _ExtentX        =   9657
      _ExtentY        =   5450
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPathTrackView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmPath As Object
Private mbyt场合 As Byte                    '0-住院临床路径跟踪;1-门诊临床路径跟踪

Public Sub ShowMe(frmParent As Object, vPati As TYPE_Pati, ByVal blnMoved As Boolean, Optional ByVal byt场合 As Byte = 0)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    mbyt场合 = byt场合
    If mbyt场合 = 0 Then
        strSql = "Select NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.住院号,B.出院病床 as 床号," & _
            " C.名称 as 科室,B.入院日期,B.出院日期" & _
            " From 病人信息 A,病案主页 B,部门表 C" & _
            " Where A.病人ID=B.病人ID And B.出院科室ID=C.ID And A.病人ID=[1] And B.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vPati.病人ID, vPati.主页ID)
        Me.tbcPath.Item(0).Caption = "姓名：" & rsTmp!姓名 & "　性别：" & Nvl(rsTmp!性别) & "　年龄：" & Nvl(rsTmp!年龄) & _
            "　科室：" & rsTmp!科室 & "　住院号：" & Nvl(rsTmp!住院号) & "　床号：" & Nvl(rsTmp!床号) & _
            "　第" & vPati.主页ID & "次住院：" & Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm") & _
            IIf(Not IsNull(rsTmp!出院日期), "-" & Format(rsTmp!出院日期, "yyyy-MM-dd HH:mm"), "")
        
        With vPati
            Call mfrmPath.zlRefresh(.病人ID, .主页ID, .病区ID, .科室ID, .病人状态, blnMoved)
        End With
    Else
        strSql = "Select NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,B.门诊号,C.名称 as 科室 " & _
            " From 病人信息 A,病人挂号记录 B,部门表 C" & _
            " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=[1] And B.ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, vPati.病人ID, vPati.挂号ID)
        Me.tbcPath.Item(0).Caption = "姓名：" & rsTmp!姓名 & "　性别：" & Nvl(rsTmp!性别) & "　年龄：" & Nvl(rsTmp!年龄) & _
            "　科室：" & rsTmp!科室 & "　门诊号：" & Nvl(rsTmp!门诊号)
        
        With vPati
            Call mfrmPath.zlRefresh(.病人ID, .挂号ID, .挂号NO, .科室ID, .病人状态, blnMoved)
        End With
    End If
    Me.Show , frmParent
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If mbyt场合 = 0 Then
        Set mfrmPath = New frmPathTable
    Else
        Set mfrmPath = New frmPathTableOut
    End If
    'TabControl
    '-----------------------------------------------------
    With Me.tbcPath
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        .InsertItem 0, "病人临床路径", mfrmPath.Hwnd, 0
    End With
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    Me.tbcPath.Left = 0
    Me.tbcPath.Top = 0
    Me.tbcPath.Width = Me.ScaleWidth
    Me.tbcPath.Height = Me.ScaleHeight
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    Unload mfrmPath
    Set mfrmPath = Nothing
End Sub
