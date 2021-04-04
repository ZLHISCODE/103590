VERSION 5.00
Begin VB.Form frmMediPriceCpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "售价计算器"
   ClientHeight    =   2505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "frmMediPriceCpt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3990
      Picture         =   "frmMediPriceCpt.frx":000C
      TabIndex        =   9
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      Picture         =   "frmMediPriceCpt.frx":0156
      TabIndex        =   10
      Top             =   900
      Width           =   1100
   End
   Begin VB.TextBox txt零售价 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   915
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1950
      Width           =   1935
   End
   Begin VB.TextBox txt成本价 
      Height          =   300
      Left            =   900
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -90
      TabIndex        =   1
      Top             =   375
      Width           =   5925
   End
   Begin VB.Label lbl零售价 
      AutoSize        =   -1  'True
      Caption         =   "零售价"
      Height          =   180
      Left            =   300
      TabIndex        =   7
      Top             =   2010
      Width           =   540
   End
   Begin VB.Label lbl差价让利比 
      AutoSize        =   -1  'True
      Caption         =   "差价让利比：60%"
      Height          =   180
      Left            =   300
      TabIndex        =   6
      Top             =   1170
      Width           =   1350
   End
   Begin VB.Label lbl加成率 
      AutoSize        =   -1  'True
      Caption         =   "加成率：18%"
      Height          =   180
      Left            =   300
      TabIndex        =   5
      Top             =   855
      Width           =   990
   End
   Begin VB.Label lbl指导售价 
      AutoSize        =   -1  'True
      Caption         =   "指导售价：10.65元/合"
      Height          =   180
      Left            =   300
      TabIndex        =   4
      Top             =   555
      Width           =   1800
   End
   Begin VB.Label lbl成本价 
      AutoSize        =   -1  'True
      Caption         =   "成本价"
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   1635
      Width           =   540
   End
   Begin VB.Label lblMediName 
      AutoSize        =   -1  'True
      Caption         =   "阿莫西林 0.125g*12粒*2板*1盒 海南三叶"
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   3330
   End
End
Attribute VB_Name = "frmMediPriceCpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim intFormat As String
Dim strKind As String

Public Function ShowMe(intUnit As Integer, lngMediId As Long) As Double
    
    On Error GoTo errHandle
    gstrSql = "select I.ID,I.类别,I.编码,I.名称,I.规格,I.产地,I.计算单位,P.药库单位,P.药库包装," & _
             "       P.指导零售价,P.指导差价率,P.差价让利比,P.成本价" & _
             " from 收费项目目录 I,药品规格 P" & _
             " where I.ID=P.药品ID and I.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    With rsTemp
        If .RecordCount <= 0 Then Unload Me: Exit Function
        
        strKind = Nvl(!类别)
        Me.lblMediName.Caption = !编码 & "-" & !名称 & " " & Nvl(!规格) & " " & Nvl(!产地)
        Me.lbl加成率.Tag = Format(100 / (1 - Nvl(!指导差价率, 0) / 100) - 100, "0.00000")
        Me.lbl加成率.Caption = "加成率：" & Me.lbl加成率.Tag & "%"
        Me.lbl差价让利比.Tag = Nvl(!差价让利比, 0)
        Me.lbl差价让利比.Caption = "差价让利比：" & Me.lbl差价让利比.Tag & "%"
        If intUnit = 0 Then
            intFormat = 7
            Me.lbl指导售价.Tag = FormatEx(Nvl(!指导零售价, 0), intFormat)
            Me.lbl指导售价.Caption = "指导售价：" & Me.lbl指导售价.Tag & "元/" & Nvl(!计算单位)
            Me.txt成本价.Text = FormatEx(Nvl(!成本价, 0), intFormat)
        Else
            intFormat = 2
            Me.lbl指导售价.Tag = FormatEx(Nvl(!指导零售价, 0) * Nvl(!药库包装, 1), intFormat)
            Me.lbl指导售价.Caption = "指导售价：" & Me.lbl指导售价.Tag & "元/" & Nvl(!药库单位)
            Me.txt成本价.Text = FormatEx(Nvl(!成本价, 0) * Nvl(!药库包装, 1), intFormat)
        End If
    End With
    Me.Show 1
    
    ShowMe = Val(Me.txt零售价.Text)
    Unload Me
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCanc_Click()
    Me.txt零售价.Text = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub txt成本价_GotFocus()
    Me.txt成本价.SelStart = 0: Me.txt成本价.SelLength = 100
End Sub

Private Sub txt成本价_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt成本价_LostFocus()
    Dim dblSalePrice As Double
    Me.txt成本价.Text = FormatEx(Val(Me.txt成本价.Text), intFormat)
    dblSalePrice = Val(Me.txt成本价.Text) * (1 + Val(Me.lbl加成率.Tag) / 100)
    If strKind <> "7" Then
        dblSalePrice = dblSalePrice + (Val(Me.lbl指导售价.Tag) - dblSalePrice) * (1 - Val(Me.lbl差价让利比.Tag) / 100)
        If dblSalePrice > Val(Me.lbl指导售价.Tag) Then dblSalePrice = Val(Me.lbl指导售价.Tag)
    End If
    Me.txt零售价.Text = FormatEx(dblSalePrice, intFormat)
End Sub

Private Sub txt零售价_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
