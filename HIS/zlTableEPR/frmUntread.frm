VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmUntread 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "版本回退"
   ClientHeight    =   3555
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5700
   Icon            =   "frmUntread.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdUntread 
      Caption         =   "回退(&U)"
      Height          =   375
      Left            =   2790
      TabIndex        =   3
      Top             =   2955
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   4095
      TabIndex        =   2
      Top             =   2955
      Width           =   1230
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2085
      Left            =   285
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      _cx             =   8916
      _cy             =   3678
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   255
      Picture         =   "frmUntread.frx":058A
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "该病历审阅修订情况如下，可以逐步回退以撤消对病历的修订和签名。"
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   195
      Width           =   4500
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUntread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mfParent As Object, mstrPrivs As String

Public Function ShowMe(ByVal fParent As Object, ByVal strPrivs As String) As Boolean
'功能：显示病历的版本修订变化情况，让用户决定执行回退
'返回：成功与否
Dim rsTemp As New ADODB.Recordset
    mblnOK = False
    On Error GoTo errHand
1    Set mfParent = fParent: mstrPrivs = strPrivs
2    gstrSQL = "Select 要素表示, 内容文本, 对象属性, 终止版,对象类型" & vbNewLine & _
            "From (Select 要素表示, 内容文本, 对象属性, 终止版,对象类型" & vbNewLine & _
            "       From 电子病历内容" & vbNewLine & _
            "       Where 文件id = [1] And 对象类型 In (6, 7, 8) And nvl(终止版,0)>0 " & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select Distinct 0 要素表示, '修订' 内容文本, '|0;;;;' 对象属性, 终止版-0.1 终止版,6 对象类型" & vbNewLine & _
            "       From 电子病历内容" & vbNewLine & _
            "       Where 文件id = [1] And 对象类型 Not In (6, 7, 8))" & vbNewLine & _
            "Order By 终止版 Desc"
3    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mfParent.Document.EPRPatiRecInfo.ID)
4    If rsTemp.EOF Then
5        MsgBox "当前没有可以回退的签名版本！", vbInformation, gstrSysName
6        Exit Function
7    Else
8        If rsTemp.RecordCount = 1 Then
9            MsgBox "当前没有可以回退的签名版本！", vbInformation, gstrSysName
10            Exit Function
11        End If
12    End If
    
13    With Me.vfgThis
14        .Clear
15        .Tag = mfParent.Document.EPRPatiRecInfo.ID
16        .Cols = 6: .Rows = rsTemp.RecordCount + 1
17        .ColWidth(0) = 1200: .ColWidth(1) = 1200: .ColWidth(2) = 1800: .ColWidth(3) = 0:    .ColWidth(4) = 0:   .ColWidth(5) = 0
18        .TextMatrix(0, 0) = "签名级别": .TextMatrix(0, 1) = "签名人": .TextMatrix(0, 2) = "签名时间"
19        .TextMatrix(0, 3) = "签名版本":   .TextMatrix(0, 4) = "签名方式":   .TextMatrix(0, 5) = "签名类型"
20        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
21        Do Until rsTemp.EOF
22            .TextMatrix(rsTemp.AbsolutePosition, 0) = Decode(mfParent.Document.EPRPatiRecInfo.病历种类, 4, Decode(rsTemp!要素表示, 3, "护士长", 1, "护士", "修订"), Decode(rsTemp!要素表示, 3, "主任医师", 2, "主治医师", 1, "经治医师", "修订"))
23            .TextMatrix(rsTemp.AbsolutePosition, 1) = Nvl(rsTemp!内容文本)
24            .TextMatrix(rsTemp.AbsolutePosition, 2) = Split(Split(rsTemp!对象属性, "|")(1), ";")(4)
25            .TextMatrix(rsTemp.AbsolutePosition, 3) = CInt(rsTemp!终止版)
26            .TextMatrix(rsTemp.AbsolutePosition, 4) = Val(Split(Split(rsTemp!对象属性, "|")(1), ";")(0))
27            .TextMatrix(rsTemp.AbsolutePosition, 5) = CInt(rsTemp!对象类型)
28            rsTemp.MoveNext
29        Loop
30    End With
    
31    Me.Show vbModal
32    ShowMe = mblnOK
    Exit Function
errHand:
    Call MsgBox("frmUntread:ShowMe错误行：" & Erl(), vbInformation, gstrSysName)
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    ShowMe = False
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdUntread_Click()
Dim objESign As Object  '电子签名接口部件
On Error GoTo errHand
    If vfgThis.TextMatrix(vfgThis.Row, 5) = 6 Then
        If mfParent.Document.EPRPatiRecInfo.保存人 <> UserInfo.姓名 And InStr(mstrPrivs, "回退他人签名") = 0 Then
            MsgBox "最后保存人与当前操作者不是同一人，不能回退！", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf vfgThis.TextMatrix(vfgThis.Row, 5) = 7 Or vfgThis.TextMatrix(vfgThis.Row, 5) = 8 Then
        If vfgThis.TextMatrix(vfgThis.Row, 1) <> UserInfo.姓名 And InStr(mstrPrivs, "回退他人签名") = 0 Then
            MsgBox "需要回退的签名与当前操作者不是同一人，不能回退！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("注意：回退操作将不可恢复！是否继续？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    
    If vfgThis.TextMatrix(vfgThis.Row, 4) = 2 Then
        '数字签名验证
        Err.Clear: On Error Resume Next
        If objESign Is Nothing Then
            Set objESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err = 0
        End If
        If Not objESign Is Nothing Then
            If objESign.Initialize(gcnOracle, glngSys) Then
                If Not objESign.CheckCertificate(UserInfo.用户名) Then Exit Sub
            Else
                MsgBox "取消已签名文件时需要再次验证签名，但系统没有设置签名认证中心，不能取消。", vbOKOnly + vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            MsgBox "签名部件初始化失败！", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error GoTo errHand
    gstrSQL = "Zl_电子病历内容_Untread(" & vfgThis.Tag & "," & vfgThis.TextMatrix(vfgThis.Row, 3) & "," & IIf(vfgThis.TextMatrix(vfgThis.Row, 0) <> "修订", 1, 0) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "回退"
    '只有三行，0行为固定行，1行为签名，2行为最初原始记录
    If vfgThis.Rows = 3 Then mfParent.Document.mReadOnly = 0: mfParent.Document.ET = TabET_单病历编辑              '回退到未签名状态，可以再次签名
    mblnOK = True: Unload Me
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mfParent = Nothing
End Sub

Private Sub vfgThis_DblClick()
    If cmdUntread.Enabled Then
        Call cmdUntread_Click
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim blnEnable As Boolean


    cmdUntread.Enabled = IIf(vfgThis.TextMatrix(vfgThis.Row, 5) = 6, vfgThis.Row = 1, True)
End Sub


