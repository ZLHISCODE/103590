VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmInsElement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ҫ��"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   Icon            =   "frmInsElement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3690
      Index           =   0
      Left            =   585
      TabIndex        =   28
      Tag             =   "1000"
      Top             =   1065
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   6509
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2505
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsElement.frx":058A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsElement.frx":0B24
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsElement.frx":10BE
            Key             =   "item"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVBar 
      BackColor       =   &H8000000C&
      Height          =   5850
      Left            =   3420
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5850
      ScaleWidth      =   30
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   5760
      Left            =   3555
      ScaleHeight     =   5760
      ScaleWidth      =   4620
      TabIndex        =   21
      Top             =   45
      Width           =   4620
      Begin VB.CheckBox chkDyn 
         Caption         =   "�Զ���(&K)"
         Height          =   225
         Left            =   3150
         TabIndex        =   16
         Top             =   2498
         Width           =   1110
      End
      Begin VB.CheckBox chkItemMust 
         Caption         =   "����Ҫ��"
         Height          =   210
         Left            =   2580
         TabIndex        =   32
         ToolTipText     =   "�Ƿ����Ҫ�أ�������������Ŀ�ж���"
         Top             =   5025
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3285
         TabIndex        =   20
         ToolTipText     =   "[ESC]�������˳�"
         Top             =   5415
         Width           =   1100
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "����(&I)"
         Height          =   350
         Left            =   1725
         TabIndex        =   31
         ToolTipText     =   "[F8]ȷ��"
         Top             =   5415
         Width           =   1100
      End
      Begin VB.CheckBox chkProtect 
         Caption         =   "��������(&P)"
         Height          =   225
         Left            =   915
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1485
      End
      Begin VB.CheckBox chkToString 
         Caption         =   "�Զ�תΪ�ı�(&X)"
         Height          =   225
         Left            =   2595
         TabIndex        =   29
         Top             =   4710
         Width           =   1710
      End
      Begin VB.ComboBox cbo�滻�� 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "frmInsElement.frx":1658
         Left            =   915
         List            =   "frmInsElement.frx":165A
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   4665
         Width           =   1560
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   2
         Left            =   105
         TabIndex        =   25
         Top             =   5340
         Width           =   4305
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   1
         Left            =   105
         TabIndex        =   24
         Top             =   2355
         Width           =   4305
      End
      Begin VB.CheckBox chk��̬ 
         Caption         =   "չ��(&E)"
         Height          =   225
         Left            =   2190
         TabIndex        =   15
         Top             =   2498
         Width           =   930
      End
      Begin VB.OptionButton opt�̶� 
         Caption         =   "������ʱ����Ҫ��(&A)"
         Height          =   180
         Index           =   0
         Left            =   1125
         TabIndex        =   1
         Top             =   585
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton opt�̶� 
         Caption         =   "����̶�����Ҫ��(&B)"
         Height          =   180
         Index           =   1
         Left            =   1125
         TabIndex        =   2
         Top             =   885
         Width           =   2775
      End
      Begin VB.TextBox txt��λ 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1635
         Width           =   1080
      End
      Begin VB.TextBox txtֵ�� 
         Height          =   1230
         IMEMode         =   1  'ON
         Left            =   915
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2820
         Width           =   3285
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         ItemData        =   "frmInsElement.frx":165C
         Left            =   915
         List            =   "frmInsElement.frx":165E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1635
         Width           =   1080
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   915
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1275
         Width           =   3285
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   915
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1995
         Width           =   1080
      End
      Begin VB.TextBox txtС�� 
         Height          =   300
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1995
         Width           =   1080
      End
      Begin VB.ComboBox cbo��ʾ 
         Height          =   300
         ItemData        =   "frmInsElement.frx":1660
         Left            =   915
         List            =   "frmInsElement.frx":1662
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2460
         Width           =   1125
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   0
         Left            =   105
         TabIndex        =   22
         Top             =   1155
         Width           =   4305
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   150
         Picture         =   "frmInsElement.frx":1664
         Top             =   90
         Width           =   480
      End
      Begin VB.Label lblҪ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����������µ���ʱҪ�أ�����б���ѡ��Ҫ����Ϊ��ʱ��̶�Ҫ�ز��룺"
         Height          =   360
         Left            =   705
         TabIndex        =   0
         Top             =   120
         Width           =   3420
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblֵ�� 
         AutoSize        =   -1  'True
         Caption         =   "ֵ��(&V)"
         Height          =   180
         Left            =   195
         TabIndex        =   17
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   195
         TabIndex        =   3
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&T)"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   1695
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&L)"
         Height          =   180
         Left            =   195
         TabIndex        =   9
         Top             =   2055
         Width           =   630
      End
      Begin VB.Label lblС�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "С��(&D)"
         Height          =   180
         Left            =   2415
         TabIndex        =   11
         Top             =   2055
         Width           =   630
      End
      Begin VB.Label lbl��λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ(&U)"
         Height          =   180
         Left            =   2415
         TabIndex        =   7
         Top             =   1695
         Width           =   630
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʾ(&F)"
         Height          =   180
         Left            =   195
         TabIndex        =   13
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label lbl��д˵�� 
         AutoSize        =   -1  'True
         Caption         =   "�Էֺŷָ���д��ѡ����ֵ�����磺A;B;C;D"
         Height          =   390
         Left            =   105
         TabIndex        =   19
         Top             =   4095
         Width           =   4305
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl tbcKind 
      Height          =   5445
      Left            =   180
      TabIndex        =   27
      Top             =   210
      Width           =   2850
      _Version        =   589884
      _ExtentX        =   5027
      _ExtentY        =   9604
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmInsElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'################################################################################################################
'�������
Private mblnOK As Boolean
Private mblnOnlyAutoElement As Boolean
Private mblnInRich As Boolean
'################################################################################################################
'��ʱ����
Dim Element As cTabElement

'################################################################################################################
'## ���ܣ�  �ϼ�������ñ�����Ľӿں��������ݲ���������ʾ����
'##
'## ������  frmParent       :������
'##         oElement        :���������Ҫ�ض���
'##         blnExtend       :�Ƿ������չ����������̬������
'##         blnCanProtect   :�Ƿ���������Ҫ��Ϊ������������
'##         blnOnlyAutoElement:ֻ��ʹ���Զ��滻Ҫ��
'################################################################################################################
Public Function ShowMe(ByRef frmParent As Object, Optional oElement As cTabElement, _
    Optional blnExtend As Boolean = True, Optional blnCanProtect As Boolean = False, _
    Optional blnOnlyAutoElement As Boolean = False, Optional blnInRich As Boolean = False) As Boolean
Dim aryTemp() As String
Dim lngCount As Long

    mblnOK = False:     mblnInRich = blnInRich
    
    If blnCanProtect Then
        '�������ñ���
        chkProtect.Enabled = True
    Else
        chkProtect.Enabled = False
    End If
    If blnExtend = False Then Me.chk��̬.Visible = False
    mblnOnlyAutoElement = blnOnlyAutoElement
    
    '��д��Ҫѡ�������
    aryTemp = Split("0-��ֵ;1-����;2-����;3-�߼�", ";")
    Me.cbo����.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo����.AddItem aryTemp(lngCount)
    Next
    Me.cbo����.ListIndex = 1
    
    aryTemp = Split("0-������;1-�Զ��滻;2-�ֵ���Ŀ", ";")
    Me.cbo�滻��.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�滻��.AddItem aryTemp(lngCount)
    Next
    Me.cbo�滻��.ListIndex = 0
    chkToString.Visible = False
    
    Set Element = New cTabElement
    If oElement.Ҫ������ = "" Then
        cmdInsert.Caption = "����(&I)"
        Call zlRefElementByObject(Element, True)
    Else
        cmdInsert.Caption = "�޸�(&I)"
        Call oElement.Clone(Element)
        Call zlRefElementByObject(Element, True)
    End If
    
    '��ʾ����
    Me.Show 1
    If mblnOK = False Then ShowMe = False: Exit Function
    
    '���ؽ������
    Element.Clone oElement
    ShowMe = True
    Unload Me
End Function

Private Sub cbo��ʾ_Click()
    Me.txtֵ��.Enabled = True
    Select Case Left(Me.cbo����.Text, 1)
    Case 0
        Select Case Left(Me.cbo��ʾ.Text, 1)
        Case 0: Me.lbl��д˵��.Caption = "���԰�����Сֵ;���ֵ����ʽָ����ֵ���ƣ����磺0;100"
        Case 1: Me.lbl��д˵��.Caption = "���԰�����Сֵ;���ֵ����ʽָ����ֵ���ƣ����磺0;100"
        Case 2: Me.lbl��д˵��.Caption = "��Ҫ���ֺ�(;)�ָ�ָ���ų��ѡ�Ĳ�ͬ��ֵ�����磺1;3;5"
        End Select
    Case 1
        Select Case Left(Me.cbo��ʾ.Text, 1)
        Case 0: Me.lbl��д˵��.Caption = "�����ı����룬����Ҫ����ֵ������": Me.txtֵ��.Enabled = False: Me.txtֵ��.Text = ""
        Case 2: Me.lbl��д˵��.Caption = "��Ҫ���ֺ�(;)�ָ�ָ�������ų��ѡ�����֣����磺����;�쳣"
        Case 3: Me.lbl��д˵��.Caption = "��Ҫ���ֺ�(;)�ָ�ָ����ѡ����ֵ�����磺η��;��ʹ;����"
        End Select
    Case 2 '����
        Me.lbl��д˵��.Caption = "���԰�����Сֵ;���ֵ����ʽָ�����ڷ�Χ�����磺" & Format(Now, "yyyy-MM-dd" & " 00:00:00") & ";" & Format(Now + 1, "yyyy-MM-dd" & " 00:00:00") & "��������10,8,19���־�����ʽ�ֱ��ʾ����,ʱ��,����ʱ��"
    Case 3 '�߼�
        Me.lbl��д˵��.Caption = "������;����ʽָ��ֵ�����ƣ����磺Y;N"
    End Select
    Select Case Left(Me.cbo��ʾ.Text, 1)
    Case 0, 1
        If Left(Me.cbo����.Text, 1) <> 2 Then '�������Ͳ���չ��
            Me.chk��̬.Enabled = False: Me.chk��̬.Value = 0
        Else                                 '�ڻ�ϱ༭������չ��
            Me.chk��̬.Enabled = Not mblnInRich: Me.chk��̬.Value = 0
        End If
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    Case 2, 3                                 '��ѡ��ѡ����չ��
        Me.chk��̬.Enabled = True
        chkDyn.Enabled = True
    Case 0 And Left(Me.cbo��ʾ.Text, 1) = 3
        Me.chk��̬.Enabled = True: Me.chk��̬.Value = 0
        chkDyn.Enabled = False: chkDyn.Value = vbUnchecked
    End Select
End Sub

Private Sub cbo��ʾ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����_Change()
    Me.picBack.Tag = ""
End Sub

Private Sub cbo����_Click()
Dim aryTemp() As String
Dim lngCount As Long
    '0-��ֵ��1-���֣�2-���ڣ�3-�߼�
    Me.txtС��.Enabled = Me.opt�̶�(0).Value
    Select Case Left(Me.cbo����.Text, 1)
    Case 0              '��ֵ
        aryTemp = Split("0-�ı�;1-����", ";"):: txt��λ.Enabled = True
    Case 1              '�ı�
        Me.txtС��.Text = 0: Me.txtС��.Enabled = False: txt��λ.Enabled = True
        aryTemp = Split("0-�ı�;2-��ѡ;3-��ѡ", ";")
    Case 2              '����
        txtС��.Enabled = False: txt��λ.Enabled = False: txt��λ.Text = ""
        txt����.Text = 10: txt����.Enabled = True
        aryTemp = Split("0-�ı�", ";")
    Case 3              '�߼�
        txtС��.Enabled = False: txt��λ.Enabled = False: txt��λ.Text = ""
        txt����.Text = 2: txt����.Enabled = False
        aryTemp = Split("0-�ı�", ";")
    End Select
    Me.cbo��ʾ.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo��ʾ.AddItem aryTemp(lngCount)
    Next
    Me.cbo��ʾ.ListIndex = 0
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�滻��_Click()
    If Me.cbo�滻��.ListIndex = 2 Then  '�ֵ���Ŀ
        Me.cbo��ʾ.ListIndex = 0: Me.cbo��ʾ.Enabled = False
        chkToString.Visible = False
    Else
        chkToString.Visible = (cbo�滻��.ListIndex = 1) '�滻��Ŀ
        Me.cbo��ʾ.Enabled = True
    End If
End Sub
Private Sub chkDyn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk��̬_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdInsert_Click()
Dim aryTemp
    If Me.opt�̶�(0).Value Then
        If Trim(Me.txt����.Text) = "" Then MsgBox "������Ҫ�����ƣ�", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > 40 Then MsgBox "���Ƴ��������40���ַ���20�����֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.txt��λ.Text), vbFromUnicode)) > 10 Then MsgBox "��λ���������10���ַ���5�����֣���", vbInformation, gstrSysName: Me.txt��λ.SetFocus: Exit Sub
        If Val(Me.txt����.Text) = 0 Then MsgBox "δ��ȷ���ó��ȣ�", vbExclamation, gstrSysName: Me.txt����.SetFocus: Exit Sub
        If Val(Me.txtС��.Text) <> 0 And Val(Me.txt����.Text) - Val(Me.txtС��.Text) < 2 Then MsgBox "δ��ȷ���ó��ȣ�", vbExclamation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    Else
        If Val(Me.picBack.Tag) = 0 Then MsgBox "���ǰ��涨ѡ��Ĺ̶�����Ҫ�أ�", vbExclamation, gstrSysName: Exit Sub
    End If
    Select Case Left(Me.cbo��ʾ.Text, 1)
    Case 2, 3
        If Trim(Me.txtֵ��.Text) = "" Then MsgBox "��ѡ��ѡ���ͣ��������ÿ�ѡ��Ŀ��", vbExclamation, gstrSysName: Me.txtֵ��.SetFocus: Exit Sub
    End Select
    
    Select Case Left(cbo����.Text, 1)
        Case 2 '����
            If Val(txt����.Text) < 10 Then txt����.Text = 8
            If Val(txt����.Text) >= 10 And Val(txt����.Text) < 19 Then txt����.Text = 10
            If Val(txt����.Text) >= 19 Then txt����.Text = 19
            aryTemp = Split(Trim(Me.txtֵ��.Text), ";")
            If UBound(aryTemp) > 0 Then
                If Not IsDate(aryTemp(0)) Or Len(aryTemp(0)) <> txt����.Text Then
                    MsgBox "������ֵ����Сֵ����Ϊ�����ͻ�ʱ���ͣ��ҳ��������ʽһ����", vbExclamation, gstrSysName
                    txtֵ��.SelStart = 0: txtֵ��.SelLength = Len(txtֵ��): Me.txtֵ��.SetFocus: Exit Sub
                End If
                If Not IsDate(aryTemp(1)) Or Len(aryTemp(0)) <> txt����.Text Then
                    MsgBox "������ֵ�����ֵ����Ϊ�����ͻ�ʱ���ͣ��ҳ��������ʽһ����", vbExclamation, gstrSysName
                    txtֵ��.SelStart = 0: txtֵ��.SelLength = Len(txtֵ��): Me.txtֵ��.SetFocus: Exit Sub
                End If
            ElseIf UBound(aryTemp) = 0 Then
                MsgBox "������ֵ�����/��Сֵ����Ϊ�����ͻ�ʱ���ͣ��ҳ��������ʽһ����", vbExclamation, gstrSysName
                txtֵ��.SelStart = 0: txtֵ��.SelLength = Len(txtֵ��): Me.txtֵ��.SetFocus: Exit Sub
            End If
        Case 3 '�߼�
            If Trim(txtֵ��.Text) = "" Then txtֵ��.Text = "��;��"
            If InStr(txtֵ��.Text, ";") = 0 Then MsgBox "�߼����ͣ���������������ѡ��Ŀ��", vbExclamation, gstrSysName: Me.txtֵ��.SetFocus: Exit Sub
    End Select
    
    With Element
        .Ҫ������ = Trim(Me.txt����.Text)
        .����Ҫ��ID = IIf(Me.opt�̶�(0).Value, 0, Val(Me.picBack.Tag))
        .Ҫ������ = Left(Me.cbo����.Text, 1)
        .Ҫ�س��� = Val(Me.txt����.Text)
        .Ҫ��С�� = IIf(.Ҫ������ = 0, Val(Me.txtС��.Text), 0)
        .Ҫ�ص�λ = Trim(Me.txt��λ.Text)
        .Ҫ�ر�ʾ = Left(Me.cbo��ʾ.Text, 1)
        .�滻�� = IIf(Me.opt�̶�(0).Value, 0, Me.cbo�滻��.ListIndex)
        .�Զ�ת�ı� = IIf(Me.chk��̬.Visible, IIf(Me.chkToString.Value = vbChecked, True, False), False)
        .���� = Me.chkItemMust.Value
        .������̬ = IIf(Me.chk��̬.Visible, Me.chk��̬.Value, 0)
        .��̬�� = chkDyn.Value
        If chkProtect.Enabled Then
            .�������� = IIf(chkProtect.Value = vbChecked, True, False)
        End If
        
        If mblnInRich Then
        Select Case .Ҫ������
            Case "����ҽʦǩ��", "����ҽʦǩ��", "����ҽʦǩ��"
                MsgBox "����Ҫ��ϱ༭�������ǩ��Ҫ��", vbInformation, gstrSysName
                Exit Sub
        End Select
        End If
        
        Select Case .Ҫ������
            Case 0  '��ֵ
                If Trim(Me.txtֵ��.Text) = "" Then
                    .Ҫ��ֵ�� = ""
                Else
                    aryTemp = Split(Trim(Me.txtֵ��.Text), ";")
                    .Ҫ��ֵ�� = Val(aryTemp(0)) & ";" & Val(aryTemp(1))
                End If
            Case 2  '����
                If Trim(Me.txtֵ��.Text) = "" Then
                    Select Case .Ҫ�س���
                        Case 8
                            .Ҫ��ֵ�� = "00:00:00;23:59:59"
                        Case 10
                            .Ҫ��ֵ�� = "1901-01-01;3000-01-01"
                        Case 19
                            .Ҫ��ֵ�� = "1901-01-01 00:00:00;3000-01-01 23:59:59"
                    End Select
                Else
                    aryTemp = Split(Trim(Me.txtֵ��.Text), ";")
                    Select Case .Ҫ�س���
                        Case 8      'ʱ����
                            .Ҫ��ֵ�� = Format(aryTemp(0), "hh:mm:ss") & ";" & Format(aryTemp(1), "hh:mm:ss")
                        Case 10     'ʱ����
                            .Ҫ��ֵ�� = Format(aryTemp(0), "yyyy-MM-dd") & ";" & Format(aryTemp(1), "yyyy-MM-dd")
                        Case 19     '��ʱ����
                            .Ҫ��ֵ�� = Format(aryTemp(0), "yyyy-MM-dd hh:mm:ss") & ";" & Format(aryTemp(1), "yyyy-MM-dd hh:mm:ss")
                    End Select
                End If
            Case 3  '�߼�
                aryTemp = Split(Trim(Me.txtֵ��.Text), ";")
                .Ҫ��ֵ�� = IIf(Trim(aryTemp(0)) = "", "��", Trim(aryTemp(0))) & ";" & IIf(Trim(aryTemp(1)) = "", "��", Trim(aryTemp(1)))
            Case 1      '�ı�
                Select Case .Ҫ�ر�ʾ
                Case 2, 3
                    .Ҫ��ֵ�� = Trim(Me.txtֵ��.Text)
                    If chkDyn.Value = 1 And InStr(.Ҫ��ֵ��, "�Զ���") = 0 Then .Ҫ��ֵ�� = .Ҫ��ֵ�� & ";�Զ���"
                Case Else
                    .Ҫ��ֵ�� = ""
                End Select
        End Select
        
        If .������̬ = 1 Then 'չ����ʽĬ���ı�����
            Dim T As Variant, i As Long, strContent As String
            T = Split(.Ҫ��ֵ��, ";")
            For i = 0 To UBound(T)
                strContent = strContent & IIf(.Ҫ�ر�ʾ = 3, "��", "��") & T(i) & IIf(i = UBound(T), "", "")   '������
            Next
            If .Ҫ������ = 2 Then 'ʱ���չ����ʽ��ʾҪ������
                .�����ı� = ""
            Else
                .�����ı� = strContent
            End If
        Else
            If .Ҫ������ <> 3 Then
                .�����ı� = ""
            Else
                .�����ı� = Split(.Ҫ��ֵ��, ";")(1)
            End If
        End If
    End With
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call cmdInsert_Click
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF8 Then
        Call cmdInsert_Click
    End If
End Sub

Private Sub Form_Load()
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
    With Me.tbcKind
        .SetImageList Me.imgList
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .Color = xtpTabColorOffice2003
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
    End With
    
    '�����Ѿ����õ�������������
    Err = 0: On Error GoTo errHand
    gstrSQL = "select ����,���� from ������������ order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > Me.tvwClass.Count Then Load Me.tvwClass(.AbsolutePosition - 1)
            Me.tbcKind.InsertItem(.AbsolutePosition - 1, !���� & "." & !����, Me.tvwClass(.AbsolutePosition - 1).hWnd, 0).Tag = "" & !����
            .MoveNext
        Loop
    End With
    
    Dim intKind As Long
    gstrSQL = "select ID,�ϼ�ID,����,����,����" & _
            " From ������������" & _
            " Where ���� = [1]" & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
    For intKind = 0 To Me.tvwClass.Count - 1
        Me.tvwClass(intKind).Nodes.Clear
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.tbcKind.Item(intKind).Tag))
        With rsTemp
            Do While Not .EOF
                If IsNull(!�ϼ�id) Then
                    Set objNode = Me.tvwClass(intKind).Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
                Else
                    Set objNode = Me.tvwClass(intKind).Nodes.Add("_" & !�ϼ�id, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
                End If
                objNode.Sorted = True
                objNode.Tag = IIf(IsNull(!����), "", !����)
                objNode.ExpandedImage = "expend"
                .MoveNext
            Loop
        End With
        If Me.tvwClass(intKind).Nodes.Count > 0 Then Me.tvwClass(intKind).Nodes(1).Selected = True
    Next
    If Me.tbcKind.ItemCount > 0 Then Me.tbcKind.Item(0).Selected = True
    Call RestoreWinState(Me, App.ProductName)
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    With Me.tbcKind
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
        .Left = Me.ScaleLeft: .Width = Me.picBack.Left - .Left - 30
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub opt�̶�_Click(Index As Integer)
    Me.txt����.Enabled = Me.opt�̶�(0).Value
    Me.cbo����.Enabled = Me.opt�̶�(0).Value
    Me.txt����.Enabled = Me.opt�̶�(0).Value
    Me.txtС��.Enabled = Me.opt�̶�(0).Value
    If Me.opt�̶�(0).Value = True Then
        Me.cbo�滻��.Tag = Me.cbo�滻��.ListIndex: Me.cbo�滻��.ListIndex = 0
    Else
        Me.cbo�滻��.ListIndex = Val(Me.cbo�滻��.Tag)
    End If
End Sub

Private Sub opt�̶�_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub tvwClass_DblClick(Index As Integer)
    If Me.tvwClass(Index).SelectedItem Is Nothing Then Exit Sub
    If Left(Me.tvwClass(Index).SelectedItem.Key, 1) <> "I" Then Exit Sub
    Call zlRefElementByString(Me.tvwClass(Index).SelectedItem.Tag)
    Me.opt�̶�(1).Value = True
End Sub

Private Sub tvwClass_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
    If Node.Children > 0 Then Exit Sub
    If Left(Node.Key, 1) <> "_" Then Exit Sub
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "select  ID,����,������,����,����,С��,С��,��λ,��ʾ��,��ֵ��,�滻��,����,��̬��" & _
            " from ����������Ŀ I" & _
            " where ����ID=[1]"
    If mblnOnlyAutoElement Then
        gstrSQL = gstrSQL & " and �滻��=1"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Mid(Node.Key, 2)))
    With rsTemp
        Do While Not .EOF
            If InStr("����ҽʦǩ��,����ҽʦǩ��,����ҽʦǩ��", !������) = 0 Then
                Set objNode = Me.tvwClass(Index).Nodes.Add(Node.Key, tvwChild, "I" & !ID, "[" & !���� & "]" & !������, "item")
                objNode.Tag = !������ & "|" & !ID & "|" & !���� & "|" & !���� & "|" & !С�� & "|" & !��λ
                Select Case Val("" & !��ʾ��)
                Case 5: objNode.Tag = objNode.Tag & "|1||0" & "|" & !�滻�� & "|0|0|" & !���� & "|" & Nvl(!��̬��, 0)
                Case 4: objNode.Tag = objNode.Tag & "|2|" & !��ֵ�� & "|0" & "|" & !�滻�� & "|0|0|" & !���� & "|" & Nvl(!��̬��, 0)
                Case Else: objNode.Tag = objNode.Tag & "|" & !��ʾ�� & "|" & !��ֵ�� & "|0" & "|" & !�滻�� & "|0|0|" & !���� & "|" & Nvl(!��̬��, 0)
                End Select
            End If
            .MoveNext
        Loop
    End With
    Node.Expanded = True
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt����_Change()
    Me.picBack.Tag = ""
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��λ_Change()
    ValidControlText txt��λ
End Sub

Private Sub txt��λ_GotFocus()
    Me.txt��λ.SelStart = 0: Me.txt��λ.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��λ_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
    Me.picBack.Tag = ""
End Sub

Private Sub txtС��_Change()
    Me.picBack.Tag = ""
End Sub

Private Sub txtС��_GotFocus()
    Me.txtС��.SelStart = 0: Me.txtС��.SelLength = 100
End Sub

Private Sub txtС��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtֵ��_Change()
    ValidControlText txtֵ��
    If Left(cbo����, 1) <> 3 Then '���߼���
        'ȥ�������ַ���������
        txtֵ�� = Replace(txtֵ��, "��", "")
        txtֵ�� = Replace(txtֵ��, "��", "")
        txtֵ�� = Replace(txtֵ��, "��", "")
        txtֵ�� = Replace(txtֵ��, "��", "")
    End If
    If cbo����.ListIndex = 1 And Left(Me.cbo��ʾ.Text, 1) <> 0 Then
        '�ı�����ѡ/��ѡ
        On Error Resume Next
        Dim lngNum As Long, T As Variant
        T = Split(txtֵ��.Text, ";")
        txt����.Text = Len(txtֵ��.Text) + (UBound(T) + 1) * 2 + 4
        Err.Clear
    End If
End Sub

Private Sub txtֵ��_GotFocus()
    If Me.cbo��ʾ.ListIndex = 0 Then
        Call zlCommFun.OpenIme(False)
    End If
End Sub

Private Sub txtֵ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
    Case vbKeyReturn: KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    Case Else
        If Me.cbo����.ListIndex = 0 Then
            If InStr("0123456789.;-", Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ��Ԫ�ذ��ն���������д���༭�ؼ�
'##
'## ������
'##         Element     :���������Ҫ�ض���
'##         blnFromOut  :�Ƿ��ⲿ�ṩ�޸ĵ�Ԫ��
'################################################################################################################
Public Sub zlRefElementByObject(ByRef Ele As cTabElement, Optional blnFromOut As Boolean)
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
    Dim intKind As Integer, lngItemId As Long
    lngItemId = Val(Ele.����Ҫ��ID)
    If lngItemId <> 0 Then
        gstrSQL = "Select c.����, i.����id, i.Id From ����������Ŀ i, ������������ c Where i.����id = c.Id And i.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngItemId)
        For intKind = 0 To Me.tbcKind.ItemCount - 1
            If Val(Me.tbcKind.Item(intKind).Tag) = rsTemp!���� Then
                Me.tbcKind.Item(intKind).Selected = True
                On Error Resume Next
                Set objNode = Nothing
                Set objNode = Me.tvwClass(intKind).Nodes("_" & rsTemp!����id)
                If Not (objNode Is Nothing) Then Call tvwClass_NodeClick(intKind, objNode)
                
                Set objNode = Nothing
                Set objNode = Me.tvwClass(intKind).Nodes("I" & lngItemId)
                If Not (objNode Is Nothing) Then
                    objNode.Selected = True
                    objNode.EnsureVisible
                End If
                Err.Clear
                Exit For
            End If
        Next
    End If
    
    Dim strElement As String
    strElement = Ele.Ҫ������ & "|" & Ele.����Ҫ��ID & "|" & Ele.Ҫ������ & "|" & Ele.Ҫ�س��� & "|" & Ele.Ҫ��С�� & "|" & Ele.Ҫ�ص�λ _
                & "|" & Ele.Ҫ�ر�ʾ & "|" & Ele.Ҫ��ֵ�� & "|" & Ele.������̬ & "|" & Ele.�滻�� & "|" & IIf(Ele.�Զ�ת�ı�, 1, 0) _
                & "|" & IIf(Ele.��������, 1, 0) & "|" & Ele.���� & "|" & Ele.��̬��
    Call zlRefElementByString(strElement, blnFromOut)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'################################################################################################################
'## ���ܣ�  ��Ԫ�ذ������Էֽ���д���༭�ؼ�
'##
'## ������
'##         strElement  :��|�ָ���Ԫ�����Դ�
'##         blnFromOut  :�Ƿ��ⲿ�ṩ�޸ĵ�Ԫ��
'################################################################################################################
Public Sub zlRefElementByString(ByVal strElement As String, Optional blnFromOut As Boolean)
Dim aryTemp() As String
Dim lngCount As Long
    aryTemp = Split(strElement, "|")
    Me.txt����.Text = aryTemp(0)
    Me.cbo����.ListIndex = IIf(aryTemp(0) = "", 1, Val(aryTemp(2)))
    Me.txt����.Text = Val(aryTemp(3))
    Me.txtС��.Text = Val(aryTemp(4))
    Me.txt��λ.Text = aryTemp(5)
    For lngCount = 0 To Me.cbo��ʾ.ListCount - 1
        If Val(Left(Me.cbo��ʾ.List(lngCount), 1)) = Val(aryTemp(6)) Then
            Me.cbo��ʾ.ListIndex = lngCount: Exit For
        End If
    Next
    Me.txtֵ��.Text = aryTemp(7)
    If UBound(aryTemp) >= 8 And Me.chk��̬.Enabled Then Me.chk��̬.Value = aryTemp(8)
    Me.cbo�滻��.ListIndex = Val(aryTemp(9)): Me.cbo�滻��.Tag = Val(aryTemp(9))
    Me.chkToString.Value = IIf(Val(aryTemp(10)) = 0, vbUnchecked, vbChecked)
    Me.chkToString.Visible = (Me.cbo�滻��.ListIndex = 1)
    Me.chkProtect.Value = IIf(Val(aryTemp(11)) = 1, vbChecked, vbUnchecked)
    Me.chkItemMust.Value = aryTemp(12)
    Me.chkDyn.Value = Val(aryTemp(13))

    'ID��������ã����������������б������¼����
    If blnFromOut Then
        If Val(aryTemp(1)) = 0 Then
            Me.opt�̶�(0).Value = True
        Else
            Me.opt�̶�(1).Value = True
        End If
    End If
    Me.picBack.Tag = Val(aryTemp(1))
End Sub
