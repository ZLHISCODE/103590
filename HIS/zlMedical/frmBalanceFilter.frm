VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBalanceFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   3075
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6330
   Icon            =   "frmBalanceFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "过滤条件"
      Height          =   2835
      Left            =   105
      TabIndex        =   21
      Top             =   135
      Width           =   4815
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   1410
         TabIndex        =   13
         Top             =   1530
         Width           =   1320
      End
      Begin VB.PictureBox picCmd 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3915
         ScaleHeight     =   255
         ScaleWidth      =   675
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1965
         Width           =   675
         Begin VB.CommandButton cmd 
            Caption         =   "&P"
            Height          =   240
            Index           =   0
            Left            =   15
            TabIndex        =   16
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "X"
            Height          =   240
            Index           =   0
            Left            =   285
            TabIndex        =   17
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "包含作废单据(&7)"
         Height          =   195
         Index           =   0
         Left            =   2940
         TabIndex        =   18
         Top             =   2400
         Width           =   1680
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1410
         TabIndex        =   15
         Top             =   1935
         Width           =   3075
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   3150
         TabIndex        =   11
         Top             =   1140
         Width           =   1320
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   3150
         TabIndex        =   7
         Top             =   750
         Width           =   1320
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1410
         TabIndex        =   5
         Top             =   750
         Width           =   1320
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1410
         TabIndex        =   9
         Top             =   1140
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75563011
         CurrentDate     =   38357
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   3150
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   75563011
         CurrentDate     =   38357
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "结算团体(&5)"
         Height          =   180
         Index           =   7
         Left            =   270
         TabIndex        =   14
         Top             =   2010
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   6
         Left            =   2835
         TabIndex        =   10
         Top             =   1215
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   5
         Left            =   2835
         TabIndex        =   6
         Top             =   825
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "结算人(&4)"
         Height          =   180
         Index           =   2
         Left            =   450
         TabIndex        =   12
         Top             =   1605
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "结算时间(&1)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   0
         Top             =   435
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "单据号(&2)"
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   4
         Top             =   840
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "票据号(&3)"
         Height          =   180
         Index           =   3
         Left            =   450
         TabIndex        =   8
         Top             =   1215
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   4
         Left            =   2835
         TabIndex        =   2
         Top             =   420
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   20
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5100
      TabIndex        =   19
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frmBalanceFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mfrmMain As Object
Private mstrCondition As String

Private Type Items
    名称 As String
End Type

Private usrSaveGroup As Items

Public Function ShowFilter(ByVal frmMain As Object, ByRef strCondition As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    mblnOK = False
            
    Set mfrmMain = frmMain
    mstrCondition = strCondition

    If ReadCondition = False Then Exit Function
            
    usrSaveGroup.名称 = txt(1).Text
    
    txt(1).Tag = ""
    
    Me.Show 1, mfrmMain
    
    If mblnOK Then strCondition = mstrCondition
    
    ShowFilter = mblnOK
    
End Function


Private Function ReadCondition() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    Dim varCondition As Variant
    
    On Error GoTo errHand
    
    '以'为分隔符
    '存储格式:开始时间'结束时间'开始单据号'结束单据号'开始票据号'结束票据号'结算人'体检团体'体检团体id'体检号'包括确认
    varCondition = Split(mstrCondition, "'")
    
    dtp(0).Value = Format(varCondition(0), dtp(0).CustomFormat)
    dtp(1).Value = Format(varCondition(1), dtp(1).CustomFormat)
    txt(0).Text = varCondition(2)
    txt(3).Text = varCondition(3)
    
    txt(2).Text = varCondition(4)
    txt(4).Text = varCondition(5)
    txt(5).Text = varCondition(6)
    
    txt(1).Text = varCondition(7)
    cmd(0).Tag = Val(varCondition(8))
    chk(0).Value = Val(varCondition(9))
    
    ReadCondition = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub chk_Click(Index As Integer)
    If Index = 1 Then
        txt(1).Enabled = (chk(Index).Value = 1)
        cmd(0).Enabled = txt(1).Enabled
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    gstrSQL = GetPublicSQL(SQL.体检团体选择)
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt(1), "编码,900,0,1;名称,1500,0,1;简码,900,0,1;地址,3000,0,1", Me.Name & "\体检团体选择", "请在下表中选择一个团体/单位。", rsData, rs, 8790, 5100) Then
    
        txt(1).Text = zlCommFun.NVL(rs("名称").Value)
        cmd(Index).Tag = zlCommFun.NVL(rs("ID").Value, 0)
        usrSaveGroup.名称 = txt(1).Text
                
    End If

    txt(1).SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click(Index As Integer)
    cmd(0).Tag = ""
    txt(1).Text = ""
    txt(1).Tag = ""
End Sub

Private Sub cmdOK_Click()
    
    '存储格式:开始时间'结束时间'开始单据号'结束单据号'开始票据号'结束票据号'结算人'体检团体'体检团体id'体检号'包括确认
    
    mstrCondition = ""
    mstrCondition = Format(dtp(0).Value, dtp(0).CustomFormat) & "'" & Format(dtp(1).Value, dtp(1).CustomFormat)
    mstrCondition = mstrCondition & "'" & Trim(txt(0).Text)
    mstrCondition = mstrCondition & "'" & Trim(txt(3).Text)
    mstrCondition = mstrCondition & "'" & Trim(txt(2).Text)
    mstrCondition = mstrCondition & "'" & Trim(txt(4).Text)
    mstrCondition = mstrCondition & "'" & Trim(txt(5).Text)
        
    If Val(cmd(0).Tag) > 0 Then
        mstrCondition = mstrCondition & "'" & Trim(txt(1).Text)
        mstrCondition = mstrCondition & "'" & Val(cmd(0).Tag)
    Else
        mstrCondition = mstrCondition & "''"
    End If
    mstrCondition = mstrCondition & "'" & chk(0).Value
    
    mblnOK = True
    
    Unload Me
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 1 Then
        txt(Index).Tag = "Changed"
        cmd(0).Tag = ""
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 5, 1
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        If txt(Index).Tag = "Changed" And Index = 1 Then
            
            gstrSQL = GetPublicSQL(SQL.团体过滤选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(txt(Index).Text) & "%")
            
            If ShowTxtFilter(Me, txt(Index), "名称,1800,0,0;编码,900,0,0;简码,900,0,0;联系人,900,0,0;电话,1200,0,0", Me.Name & "\团体过滤选择", "请从下面选择一个团体单位", rsData, rs) Then
                
                txt(1).Text = zlCommFun.NVL(rs("名称"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                
                usrSaveGroup.名称 = txt(1).Text
            Else
                txt(Index).Text = usrSaveGroup.名称
                Exit Sub
            End If
        End If
        
        zlCommFun.PressKey vbKeyTab
        
        Select Case Index
        Case 1
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
        End Select
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
        Case 0, 1, 2, 3, 4
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Dim intYear As Integer
    Dim strYear As String
    
    Select Case Index
    Case 5, 1
        zlCommFun.OpenIme False
    Case 0, 3
        
        '自动补齐单据号
        If (UCase(Left(txt(Index).Text, 1)) < "A" Or UCase(Left(txt(Index).Text, 1)) > "Z") And Trim(txt(Index).Text) <> "" Then
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            txt(Index).Text = strYear & Right("0000000" & txt(Index).Text, 7)
        End If
    End Select
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    
    If Index = 1 Then
        If txt(Index).Tag = "Changed" Then
            txt(Index).Text = usrSaveGroup.名称
            txt(Index).Tag = ""
        End If
    End If
End Sub


