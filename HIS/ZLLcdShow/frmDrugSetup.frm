VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDrugSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6015
   Icon            =   "frmDrugSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   6015
   StartUpPosition =   1  '����������
   Begin VB.Frame Fra�������� 
      Caption         =   "�������������"
      Height          =   5895
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   5535
      Begin VB.Frame �������� 
         Caption         =   "��������"
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   5295
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   0
            Left            =   240
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdCall 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   52
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl���� 
            Caption         =   "����;�Ӵ�;24"
            Height          =   225
            Index           =   0
            Left            =   1320
            TabIndex        =   54
            Top             =   303
            Width           =   2120
         End
         Begin VB.Shape shpCall 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.Frame fraѡ������ʾ 
         Caption         =   "ѡ������ʾ"
         Height          =   4335
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   5295
         Begin VB.CommandButton cmdPreparingColor 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   37
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   36
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chk��ʾ����ҩ 
            Caption         =   "��ʾ����ҩ"
            Height          =   200
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CheckBox chk��ʾ����ҩ 
            Caption         =   "��ʾ����ҩ"
            Height          =   200
            Left            =   120
            TabIndex        =   34
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox chk��ʾ�������� 
            Caption         =   "��ʾ��������"
            Height          =   200
            Left            =   120
            TabIndex        =   33
            Top             =   3120
            Width           =   1575
         End
         Begin VB.CheckBox chk��ʾ���� 
            Caption         =   "��ʾ����"
            Height          =   200
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdWin 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   31
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   1
            Left            =   240
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   2
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   3
            Left            =   240
            TabIndex        =   28
            Top             =   2280
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   4
            Left            =   240
            TabIndex        =   27
            Top             =   3360
            Width           =   975
         End
         Begin VB.CommandButton cmdOther 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   26
            Top             =   3360
            Width           =   975
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2160
            TabIndex        =   25
            Text            =   "3"
            Top             =   2707
            Width           =   615
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   720
            TabIndex        =   24
            Text            =   "1"
            Top             =   2707
            Width           =   615
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   720
            TabIndex        =   23
            Text            =   "1"
            Top             =   1627
            Width           =   615
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2160
            TabIndex        =   22
            Text            =   "3"
            Top             =   1627
            Width           =   615
         End
         Begin VB.TextBox txtContent 
            Height          =   300
            Left            =   1200
            TabIndex        =   21
            Top             =   3840
            Width           =   3975
         End
         Begin VB.Label lbl���� 
            Caption         =   "����ҩ"
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   50
            Top             =   1278
            Width           =   2120
         End
         Begin VB.Label lbl���� 
            Caption         =   "����ҩ"
            Height          =   195
            Index           =   3
            Left            =   1320
            TabIndex        =   49
            Top             =   2358
            Width           =   2120
         End
         Begin VB.Shape shpWin 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   510
            Width           =   375
         End
         Begin VB.Label lbl���� 
            Caption         =   "����"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   48
            Top             =   558
            Width           =   2120
         End
         Begin VB.Shape shpOther 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00408000&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   3385
            Width           =   375
         End
         Begin VB.Label lbl���� 
            Caption         =   "��������"
            Height          =   195
            Index           =   4
            Left            =   1320
            TabIndex        =   47
            Top             =   3438
            Width           =   2120
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   1680
            TabIndex        =   45
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lbl����ҩSum 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   3600
            TabIndex        =   43
            Top             =   2707
            Width           =   615
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   3120
            TabIndex        =   41
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lbl����ҩSum 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   3600
            TabIndex        =   40
            Top             =   1627
            Width           =   615
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   1680
            TabIndex        =   39
            Top             =   1680
            Width           =   375
         End
         Begin VB.Shape shpPreparing 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00408000&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   2305
            Width           =   375
         End
         Begin VB.Shape shpCalling 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   1225
            Width           =   375
         End
         Begin VB.Label lblContent 
            Caption         =   "��ʾ����"
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   3893
            Width           =   735
         End
      End
      Begin VB.TextBox txt��ѯʱ�� 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   4080
         TabIndex        =   19
         Text            =   "10"
         Top             =   5490
         Width           =   615
      End
      Begin VB.TextBox txt��ҳʱ�� 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1320
         TabIndex        =   18
         Text            =   "10"
         Top             =   5490
         Width           =   615
      End
      Begin VB.Label lblˢ��ʱ�� 
         Caption         =   "LCDˢ��ʱ��"
         Height          =   195
         Left            =   2880
         TabIndex        =   58
         Top             =   5550
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "��"
         Height          =   195
         Left            =   2040
         TabIndex        =   57
         Top             =   5550
         Width           =   255
      End
      Begin VB.Label lbl��ҳʱ�� 
         Caption         =   "��ҳʱ��"
         Height          =   195
         Left            =   480
         TabIndex        =   56
         Top             =   5550
         Width           =   735
      End
      Begin VB.Label lbl�� 
         Caption         =   "��"
         Height          =   195
         Left            =   4800
         TabIndex        =   55
         Top             =   5550
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   9000
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   9000
      Width           =   1100
   End
   Begin VB.Frame Fra��ʾ���� 
      Caption         =   "��ʾ����"
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   5535
      Begin VB.Frame fra�кŴ��� 
         Caption         =   "�кŴ���"
         Height          =   1215
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   3255
         Begin VB.ListBox lst�кŴ��� 
            Columns         =   1
            ForeColor       =   &H80000012&
            Height          =   900
            IMEMode         =   3  'DISABLE
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.Frame Fra��ʾģʽ 
         Caption         =   "��ʾģʽ"
         Height          =   1215
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton Opt���� 
            Caption         =   "�ര��"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Opt���� 
            Caption         =   "������"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame frmRect 
      Caption         =   "Һ����λ�ã��ֱ���Ϊ��λ��"
      Height          =   1150
      Left            =   240
      TabIndex        =   0
      Top             =   1770
      Width           =   5535
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Top             =   310
         Width           =   1695
      End
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   2
         Left            =   3600
         TabIndex        =   3
         Top             =   310
         Width           =   1695
      End
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   3
         Left            =   840
         TabIndex        =   2
         Top             =   710
         Width           =   1695
      End
      Begin VB.TextBox txtRect 
         Height          =   300
         Index           =   4
         Left            =   3600
         TabIndex        =   1
         Top             =   710
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "��"
         Height          =   255
         Index           =   0
         Left            =   405
         TabIndex        =   8
         Top             =   340
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   3165
         TabIndex        =   7
         Top             =   345
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "��ȣ�"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "�߶ȣ�"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   5
         Top             =   750
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   240
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   0
      Left            =   480
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   1
      Left            =   2040
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   2
      Left            =   2520
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   3
      Left            =   3120
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   4
      Left            =   1440
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDrugSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrWins As String

Private Sub chk��ʾ����_Click()
    If chk��ʾ����.Value = 1 Then
        cmdFont(1).Enabled = True
        Me.cmdWin.Enabled = True
    Else
        cmdFont(1).Enabled = False
        Me.cmdWin.Enabled = False
    End If
End Sub

Private Sub chk��ʾ����ҩ_Click()
    If chk��ʾ����ҩ.Value = 1 Then
        cmdFont(2).Enabled = True
        Me.cmdColor.Enabled = True
        Me.txt����ҩ����.Enabled = True
        Me.txt����ҩ����.Enabled = True
    Else
        cmdFont(2).Enabled = False
        Me.cmdColor.Enabled = False
        Me.txt����ҩ����.Enabled = False
        Me.txt����ҩ����.Enabled = False
    End If
End Sub

Private Sub chk��ʾ����ҩ_Click()
    If chk��ʾ����ҩ.Value = 1 Then
        cmdFont(3).Enabled = True
        Me.cmdPreparingColor.Enabled = True
        Me.txt����ҩ����.Enabled = True
        Me.txt����ҩ����.Enabled = True
    Else
        cmdFont(3).Enabled = False
        Me.cmdPreparingColor.Enabled = False
        Me.txt����ҩ����.Enabled = False
        Me.txt����ҩ����.Enabled = False
    End If
End Sub

Private Sub chk��ʾ��������_Click()
    If chk��ʾ��������.Value = 1 Then
        cmdFont(4).Enabled = True
        Me.cmdOther.Enabled = True
    Else
        cmdFont(4).Enabled = True
        Me.cmdOther.Enabled = True
    End If
End Sub

Private Sub cmdCall_Click()
    dlgColor.Color = shpCall.FillColor
    dlgColor.ShowColor
    shpCall.FillColor = dlgColor.Color
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click(Index As Integer)
    Dim strReg As String
    Dim str���� As String
    
    On Error GoTo err
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    dlgFont(Index).Flags = cdlCFBoth
    dlgFont(Index).CancelError = False  '�ѵ�ȡ������������
'    dlgFont(Index).FontName = GetSetting("ZLSOFT", strReg, "����(" & Index & ")", "����")
'    dlgFont(Index).FontBold = GetSetting("ZLSOFT", strReg, "����(" & Index & ")", "False")
'    dlgFont(Index).FontItalic = GetSetting("ZLSOFT", strReg, "б��(" & Index & ")", "False")
'    dlgFont(Index).FontSize = GetSetting("ZLSOFT", strReg, "�ֺ�(" & Index & ")", "14")
    dlgFont(Index).ShowFont
    On Error GoTo 0
    '��������
    SaveSetting "ZLSOFT", strReg, "����(" & Index & ")", dlgFont(Index).FontName
    SaveSetting "ZLSOFT", strReg, "����(" & Index & ")", dlgFont(Index).FontBold
    SaveSetting "ZLSOFT", strReg, "б��(" & Index & ")", dlgFont(Index).FontItalic
    SaveSetting "ZLSOFT", strReg, "�ֺ�(" & Index & ")", dlgFont(Index).FontSize
    Me.lbl����(Index) = dlgFont(Index).FontName & "," & IIf(dlgFont(Index).FontBold, "����,", "") & IIf(dlgFont(Index).FontItalic, "б��,", "") & dlgFont(Index).FontSize
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strReg As String
    Dim strWin As String
    Dim i As Integer
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    SaveSetting "ZLSOFT", strReg, "��", Val(txtRect(1).Text)
    SaveSetting "ZLSOFT", strReg, "��", Val(txtRect(2).Text)
    SaveSetting "ZLSOFT", strReg, "���", Val(txtRect(3).Text)
    SaveSetting "ZLSOFT", strReg, "�߶�", Val(txtRect(4).Text)
    
    SaveSetting "ZLSOFT", strReg, "����ģʽ", IIf(Me.Opt����(0).Value = True, 0, 1)
    
    
    For i = 0 To Me.lst�кŴ���.ListCount - 1
        If lst�кŴ���.Selected(i) Then
            strWin = strWin & "," & lst�кŴ���.List(i)
        End If
    Next
    strWin = Mid(strWin, 2)
    SaveSetting "ZLSOFT", strReg, "����", strWin
    
    SaveSetting "ZLSOFT", strReg, "��������ɫ", shpCall.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����", Me.chk��ʾ����.Value
    SaveSetting "ZLSOFT", strReg, "������ɫ", shpWin.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ��������", Me.chk��ʾ��������.Value
    SaveSetting "ZLSOFT", strReg, "����������ɫ", shpOther.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����ҩ", Me.chk��ʾ����ҩ.Value
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.lbl����ҩSum.Caption
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.txt����ҩ����.Text
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.txt����ҩ����.Text
    SaveSetting "ZLSOFT", strReg, "����ҩ��ɫ", shpCalling.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����ҩ", Me.chk��ʾ����ҩ.Value
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.lbl����ҩSum.Caption
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.txt����ҩ����.Text
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.txt����ҩ����.Text
    SaveSetting "ZLSOFT", strReg, "����ҩ��ɫ", shpPreparing.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ҳʱ��", Me.txt��ҳʱ��.Text
    SaveSetting "ZLSOFT", strReg, "ˢ��ʱ��", Me.txt��ѯʱ��.Text
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����", Me.txtContent.Text
    
    Unload Me
End Sub


Private Sub cmdOther_Click()
    dlgColor.Color = shpOther.FillColor
    dlgColor.ShowColor
    shpOther.FillColor = dlgColor.Color
End Sub

Private Sub cmdWin_Click()
    dlgColor.Color = shpWin.FillColor
    dlgColor.ShowColor
    shpWin.FillColor = dlgColor.Color
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Integer
    Dim strWin As String
    Dim Index  As Integer
    
    strReg = "����ģ��\ҩ���Ŷӽк�\Һ������"
    
    Me.Opt����(Val(GetSetting("ZLSOFT", strReg, "����ģʽ", "0"))).Value = True
    
    strWin = GetSetting("ZLSOFT", strReg, "����", "")
    '���ش���
    LoadWin
    '�ָ�ѡ�д���
    For i = 0 To Me.lst�кŴ���.ListCount - 1
        If InStr(1, strWin, lst�кŴ���.List(i)) > 0 Then
            lst�кŴ���.Selected(i) = True
        End If
    Next
    
    '������Ļ��Ϣ
    txtRect(1).Text = GetSetting("ZLSOFT", strReg, "��", "1024")
    txtRect(2).Text = GetSetting("ZLSOFT", strReg, "��", "0")
    txtRect(3).Text = GetSetting("ZLSOFT", strReg, "���", "1024")
    txtRect(4).Text = GetSetting("ZLSOFT", strReg, "�߶�", "768")
    
    
    shpCall.FillColor = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
    
    Me.chk��ʾ����.Value = GetSetting("ZLSOFT", strReg, "��ʾ����", 1)
    shpWin.FillColor = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
    
    Me.chk��ʾ��������.Value = GetSetting("ZLSOFT", strReg, "��ʾ��������", 1)
    shpOther.FillColor = GetSetting("ZLSOFT", strReg, "����������ɫ", vbGreen)
    
    Me.chk��ʾ����ҩ.Value = Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1"))
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "3"))
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "3"))
    Me.lbl����ҩSum.Caption = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "9"))
    shpCalling.FillColor = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
    
    Me.chk��ʾ����ҩ.Value = Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1"))
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "3"))
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "3"))
    Me.lbl����ҩSum.Caption = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "9"))
    shpPreparing.FillColor = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
    
    Me.txt��ҳʱ��.Text = GetSetting("ZLSOFT", strReg, "��ҳʱ��", "5")
    Me.txt��ѯʱ��.Text = GetSetting("ZLSOFT", strReg, "ˢ��ʱ��", "10")
    Me.txtContent.Text = GetSetting("ZLSOFT", strReg, "��ʾ����", "")
    
    For Index = 0 To Me.dlgFont.UBound
        dlgFont(Index).Flags = cdlCFBoth
        dlgFont(Index).CancelError = False  '�ѵ�ȡ������������
        dlgFont(Index).FontName = GetSetting("ZLSOFT", strReg, "����(" & Index & ")", "����")
        dlgFont(Index).FontBold = GetSetting("ZLSOFT", strReg, "����(" & Index & ")", "False")
        dlgFont(Index).FontItalic = GetSetting("ZLSOFT", strReg, "б��(" & Index & ")", "False")
        dlgFont(Index).FontSize = GetSetting("ZLSOFT", strReg, "�ֺ�(" & Index & ")", "14")
        Me.lbl����(Index) = dlgFont(Index).FontName & "," & IIf(dlgFont(Index).FontBold, "����,", "") & IIf(dlgFont(Index).FontItalic, "б��,", "") & dlgFont(Index).FontSize
    Next

End Sub

Public Function ShowMe(ByVal strWins As String, ByVal frmParent As Form) As Boolean
'����˵����strWins���ڴ�����ʽΪ������1,����2��
    mstrWins = strWins
    
    Me.Show 1, frmParent
    
    ShowMe = True
End Function

Private Sub lbl����ҩ������_Click()

End Sub

Private Sub Opt����_Click(Index As Integer)
    Me.fra�кŴ���.Enabled = IIf(Index = 0, False, True)
End Sub

Private Sub txtRect_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.txt����ҩ����.Text <> "" Then
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
    End If
End Sub

Private Sub txt����ҩ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txt����ҩ����.Text = Me.txt����ҩ����.Text
    
    If Me.txt����ҩ����.Text <> "" Then
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
    End If
    
    If Me.txt����ҩ����.Text <> "" Then
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
    End If
End Sub

Private Sub txt����ҩ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.txt����ҩ����.Text <> "" Then
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * (Me.txt����ҩ����.Text)
    End If
End Sub

Private Sub txt����ҩ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.txt����ҩ����.Text = Me.txt����ҩ����.Text
    
    If Me.txt����ҩ����.Text <> "" Then
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
    End If
    
    If Me.txt����ҩ����.Text <> "" Then
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
    End If
End Sub

Private Sub txt��ѯʱ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub cmdPreparingColor_Click()
    dlgColor.Color = shpPreparing.FillColor
    dlgColor.ShowColor
    shpPreparing.FillColor = dlgColor.Color
End Sub


Private Sub cmdColor_Click()
    dlgColor.Color = shpCalling.FillColor
    dlgColor.ShowColor
    shpCalling.FillColor = dlgColor.Color
End Sub

Private Sub LoadWin()
    Dim i As Integer
    
    For i = 0 To UBound(Split(mstrWins, ","))
        Me.lst�кŴ���.AddItem Split(mstrWins, ",")(i)
    Next
    
End Sub
