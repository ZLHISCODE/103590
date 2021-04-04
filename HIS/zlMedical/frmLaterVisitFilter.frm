VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLaterVisitFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   2505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5895
   Icon            =   "frmLaterVisitFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   4515
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   11
         Top             =   1575
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75628547
         CurrentDate     =   38777
      End
      Begin VB.CheckBox chk 
         Caption         =   "只显当前要随访的人员(&3)"
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   1260
         Width           =   2730
      End
      Begin VB.CheckBox chk 
         Caption         =   "只显示随访期内的人员(&2)"
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   915
         Width           =   2670
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   300
         Index           =   0
         Left            =   4095
         TabIndex        =   4
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1560
         TabIndex        =   3
         Top             =   210
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1560
         TabIndex        =   12
         Top             =   1980
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75628547
         CurrentDate     =   38777
      End
      Begin VB.Label lblHint 
         AutoSize        =   -1  'True
         Caption         =   "(提示：按Del清除体检诊断结论)"
         Height          =   180
         Left            =   1500
         TabIndex        =   5
         Top             =   615
         Width           =   2610
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "体检结束时间(&5)"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   1
         Top             =   2055
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "体检开始时间(&4)"
         Height          =   180
         Index           =   3
         Left            =   150
         TabIndex        =   0
         Top             =   1635
         Width           =   1350
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "体检诊断结论(&1)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4695
      TabIndex        =   6
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4695
      TabIndex        =   7
      Top             =   510
      Width           =   1100
   End
End
Attribute VB_Name = "frmLaterVisitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mfrmMain As Object
Private mlngLoop As Long
Private mblnOK As Boolean

Private Type Items
    名称 As String
End Type

Private Type CONDITION
    结论id As Long
    随访期内 As Boolean             '只显示随访期内的人员
    开始时间 As String              '历史随访人员的体检开始时间
    结束时间 As String              '历史随访人员的体检结束时间
    随访人员 As Boolean             '只显示当前要随访的人员,前提是随访期内为真的
End Type

Private mConditon As CONDITION

Private usrSave As Items

Private mlng结论id As Long

Public Function ShowPara(ByVal frmMain As Object, ByRef 结论id As Long, ByRef 随访期内 As Boolean, ByRef 随访人员 As Boolean, ByRef 开始时间 As String, ByRef 结束时间 As String) As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    mblnOK = False
    
    mConditon.结论id = 结论id
    mConditon.随访期内 = 随访期内
    mConditon.随访人员 = 随访人员
    mConditon.开始时间 = 开始时间
    mConditon.结束时间 = 结束时间
    
    Set mfrmMain = frmMain
    '初始化
    
    chk(0).Value = IIf(mConditon.随访期内, 1, 0)
    chk(1).Value = IIf(mConditon.随访人员, 1, 0)
    
    If mConditon.随访期内 = False Then
        dtp(0).Enabled = True
    Else
        dtp(0).Enabled = False
    End If
    
    If mConditon.随访期内 = False Then
        dtp(1).Enabled = True
    Else
        dtp(1).Enabled = False
    End If
    
    dtp(0).Value = Format(mConditon.开始时间, dtp(0).CustomFormat)
    dtp(1).Value = Format(mConditon.结束时间, dtp(1).CustomFormat)
    
    strSQL = "Select * from 体检诊断建议 Where 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mConditon.结论id)
    If rs.BOF = False Then
        
        txt.Text = zlCommFun.NVL(rs("名称").Value)
        cmd(0).Tag = mConditon.结论id
        usrSave.名称 = txt.Text
        
    End If
    
    Me.Show 1, frmMain
    
    结论id = mConditon.结论id
    随访期内 = mConditon.随访期内
    随访人员 = mConditon.随访人员
    开始时间 = mConditon.开始时间
    结束时间 = mConditon.结束时间
    
    ShowPara = mblnOK
    
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    Select Case Index
    Case 0
    
        chk(1).Enabled = (chk(Index).Value = 1)
        
        dtp(0).Enabled = (chk(Index).Value = 0)
        dtp(1).Enabled = (chk(Index).Value = 0)
        
        lbl(3).Enabled = dtp(0).Enabled
        lbl(2).Enabled = dtp(1).Enabled
        
        If chk(Index).Value = 0 Then
            chk(1).Value = 0
        End If
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim rsData As New ADODB.Recordset
    
    strSQL = "SELECT -1 AS ID," & _
                        "0 AS 上级ID," & _
                        "0 AS 末级," & _
                        "'' AS 编码," & _
                        "'所有分类' AS 名称, " & _
                        "Null+0 AS 是否疾病,'' As 诊断建议 " & _
                "FROM dual "
                
    strSQL = strSQL & _
            " UNION ALL " & _
            "SELECT 序号 AS ID," & _
                        "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID," & _
                        "0 AS 末级," & _
                        "编码," & _
                        "名称, " & _
                        "Null+0 AS 是否疾病,'' As 诊断建议 " & _
                "FROM 体检诊断建议 " & _
                "WHERE NVL(末级,0)=0 " & _
                "START WITH 上级序号 is NULL CONNECT BY PRIOR 序号 = 上级序号 "
    
    strSQL = strSQL & _
                "UNION ALL " & _
                "SELECT A.序号 AS ID, " & _
                        "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "A.是否疾病,A.诊断建议 " & _
                "FROM 体检诊断建议 A " & _
                "WHERE NVL(A.末级,0)=1"
                    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt, "编码,900,0,1;名称,1800,0,0;诊断建议,2700,0,0", Me.Name & "\体检诊断选择", "请在下表中选择一个诊断结论。", rsData, rs, 8790, 5100) Then
    
        txt.Text = zlCommFun.NVL(rs("名称").Value)
        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
        usrSave.名称 = txt.Text
                
    End If

    txt.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
        
    mConditon.结论id = Val(cmd(0).Tag)
    mConditon.随访期内 = (chk(0).Value = 1)
    mConditon.随访人员 = (chk(1).Value = 1)
    
    mConditon.开始时间 = Format(dtp(0).Value, dtp(0).CustomFormat)
    mConditon.结束时间 = Format(dtp(1).Value, dtp(1).CustomFormat)

    mblnOK = True

    Unload Me
End Sub

Private Sub txt_Change()
    
    txt.Tag = "Changed"
    cmd(0).Tag = ""
    
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmd(0).Tag = ""
        txt.Text = ""
        txt.Tag = ""
        usrSave.名称 = ""
    End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        
        If txt.Tag = "Changed" Then
                            
            strText = UCase(txt.Text) & "%"
            strSQL = _
                    "SELECT A.序号 AS ID, " & _
                            "A.编码, " & _
                            "A.名称, " & _
                            "A.是否疾病,A.诊断建议 " & _
                    "FROM 体检诊断建议 A " & _
                    "WHERE NVL(末级,0)=1 "
                    
            strSQL = strSQL & " AND (A.编码 Like [1] OR Upper(A.名称) Like [2] OR Upper(A.简码) Like [2])"
            
            If ParamInfo.项目输入匹配方式 = 0 Then strTmp = "%" & strText
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
            
            If ShowTxtFilter(Me, txt, "编码,900,0,1;名称,1800,0,0;诊断建议,2700,0,0", Me.Name & "\体检结论过滤", "请从下面选择一个诊断结论", rsData, rs) Then
                
                txt.Text = zlCommFun.NVL(rs("名称"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                
                usrSave.名称 = txt.Text
            Else
                txt.Text = usrSave.名称
                Exit Sub
            End If
        End If
        
        zlCommFun.PressKey vbKeyTab
        zlCommFun.PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    
    If txt.Tag = "Changed" Then
        txt.Text = usrSave.名称
        txt.Tag = ""
    End If
    
End Sub
