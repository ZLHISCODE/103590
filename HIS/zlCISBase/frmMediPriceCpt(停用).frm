VERSION 5.00
Begin VB.Form frmMediPriceCpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ۼۼ�����"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3990
      Picture         =   "frmMediPriceCpt.frx":000C
      TabIndex        =   9
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3990
      Picture         =   "frmMediPriceCpt.frx":0156
      TabIndex        =   10
      Top             =   900
      Width           =   1100
   End
   Begin VB.TextBox txt���ۼ� 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   915
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1950
      Width           =   1935
   End
   Begin VB.TextBox txt�ɱ��� 
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
   Begin VB.Label lbl���ۼ� 
      AutoSize        =   -1  'True
      Caption         =   "���ۼ�"
      Height          =   180
      Left            =   300
      TabIndex        =   7
      Top             =   2010
      Width           =   540
   End
   Begin VB.Label lbl��������� 
      AutoSize        =   -1  'True
      Caption         =   "��������ȣ�60%"
      Height          =   180
      Left            =   300
      TabIndex        =   6
      Top             =   1170
      Width           =   1350
   End
   Begin VB.Label lbl�ӳ��� 
      AutoSize        =   -1  'True
      Caption         =   "�ӳ��ʣ�18%"
      Height          =   180
      Left            =   300
      TabIndex        =   5
      Top             =   855
      Width           =   990
   End
   Begin VB.Label lblָ���ۼ� 
      AutoSize        =   -1  'True
      Caption         =   "ָ���ۼۣ�10.65Ԫ/��"
      Height          =   180
      Left            =   300
      TabIndex        =   4
      Top             =   555
      Width           =   1800
   End
   Begin VB.Label lbl�ɱ��� 
      AutoSize        =   -1  'True
      Caption         =   "�ɱ���"
      Height          =   180
      Left            =   300
      TabIndex        =   2
      Top             =   1635
      Width           =   540
   End
   Begin VB.Label lblMediName 
      AutoSize        =   -1  'True
      Caption         =   "��Ī���� 0.125g*12��*2��*1�� ������Ҷ"
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
    gstrSql = "select I.ID,I.���,I.����,I.����,I.���,I.����,I.���㵥λ,P.ҩ�ⵥλ,P.ҩ���װ," & _
             "       P.ָ�����ۼ�,P.ָ�������,P.���������,P.�ɱ���" & _
             " from �շ���ĿĿ¼ I,ҩƷ��� P" & _
             " where I.ID=P.ҩƷID and I.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    With rsTemp
        If .RecordCount <= 0 Then Unload Me: Exit Function
        
        strKind = Nvl(!���)
        Me.lblMediName.Caption = !���� & "-" & !���� & " " & Nvl(!���) & " " & Nvl(!����)
        Me.lbl�ӳ���.Tag = Format(100 / (1 - Nvl(!ָ�������, 0) / 100) - 100, "0.00000")
        Me.lbl�ӳ���.Caption = "�ӳ��ʣ�" & Me.lbl�ӳ���.Tag & "%"
        Me.lbl���������.Tag = Nvl(!���������, 0)
        Me.lbl���������.Caption = "��������ȣ�" & Me.lbl���������.Tag & "%"
        If intUnit = 0 Then
            intFormat = 7
            Me.lblָ���ۼ�.Tag = FormatEx(Nvl(!ָ�����ۼ�, 0), intFormat)
            Me.lblָ���ۼ�.Caption = "ָ���ۼۣ�" & Me.lblָ���ۼ�.Tag & "Ԫ/" & Nvl(!���㵥λ)
            Me.txt�ɱ���.Text = FormatEx(Nvl(!�ɱ���, 0), intFormat)
        Else
            intFormat = 2
            Me.lblָ���ۼ�.Tag = FormatEx(Nvl(!ָ�����ۼ�, 0) * Nvl(!ҩ���װ, 1), intFormat)
            Me.lblָ���ۼ�.Caption = "ָ���ۼۣ�" & Me.lblָ���ۼ�.Tag & "Ԫ/" & Nvl(!ҩ�ⵥλ)
            Me.txt�ɱ���.Text = FormatEx(Nvl(!�ɱ���, 0) * Nvl(!ҩ���װ, 1), intFormat)
        End If
    End With
    Me.Show 1
    
    ShowMe = Val(Me.txt���ۼ�.Text)
    Unload Me
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCanc_Click()
    Me.txt���ۼ�.Text = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub txt�ɱ���_GotFocus()
    Me.txt�ɱ���.SelStart = 0: Me.txt�ɱ���.SelLength = 100
End Sub

Private Sub txt�ɱ���_KeyPress(KeyAscii As Integer)
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

Private Sub txt�ɱ���_LostFocus()
    Dim dblSalePrice As Double
    Me.txt�ɱ���.Text = FormatEx(Val(Me.txt�ɱ���.Text), intFormat)
    dblSalePrice = Val(Me.txt�ɱ���.Text) * (1 + Val(Me.lbl�ӳ���.Tag) / 100)
    If strKind <> "7" Then
        dblSalePrice = dblSalePrice + (Val(Me.lblָ���ۼ�.Tag) - dblSalePrice) * (1 - Val(Me.lbl���������.Tag) / 100)
        If dblSalePrice > Val(Me.lblָ���ۼ�.Tag) Then dblSalePrice = Val(Me.lblָ���ۼ�.Tag)
    End If
    Me.txt���ۼ�.Text = FormatEx(dblSalePrice, intFormat)
End Sub

Private Sub txt���ۼ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub
