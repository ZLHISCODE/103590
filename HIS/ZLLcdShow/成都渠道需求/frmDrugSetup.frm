VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDrugSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   6408
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   11424
   Icon            =   "frmDrugSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6408
   ScaleWidth      =   11424
   StartUpPosition =   1  '����������
   Begin VB.Frame Frm��������1 
      Caption         =   "�������������1"
      Height          =   2745
      Left            =   240
      TabIndex        =   47
      Top             =   3000
      Width           =   5535
      Begin VB.Frame ʱ������ 
         Caption         =   "ʱ������"
         Height          =   735
         Left            =   150
         TabIndex        =   66
         Top             =   1830
         Width           =   5265
         Begin VB.TextBox txt��ѯʱ�� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   1290
            TabIndex        =   68
            Text            =   "10"
            Top             =   270
            Width           =   450
         End
         Begin VB.TextBox txt��ҳʱ�� 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   2940
            TabIndex        =   67
            Text            =   "10"
            Top             =   270
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblˢ��ʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "LCDˢ��ʱ��"
            Height          =   180
            Left            =   210
            TabIndex        =   72
            Top             =   330
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   180
            Left            =   3420
            TabIndex        =   71
            Top             =   330
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lbl��ҳʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ҳʱ��"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2220
            TabIndex        =   70
            Top             =   330
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl�� 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   1800
            TabIndex        =   69
            Top             =   330
            Width           =   180
         End
      End
      Begin VB.Frame �������� 
         Caption         =   "��������"
         Height          =   795
         Left            =   120
         TabIndex        =   64
         Top             =   1020
         Width           =   5295
         Begin VB.CommandButton cmd���� 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   240
            TabIndex        =   65
            Top             =   270
            Width           =   975
         End
         Begin VB.Shape shp���� 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   1320
            Top             =   300
            Width           =   375
         End
      End
      Begin VB.Frame �������� 
         Caption         =   "��������"
         Height          =   795
         Left            =   120
         TabIndex        =   48
         Top             =   210
         Width           =   5295
         Begin VB.CommandButton cmdCall 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   50
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   0
            Left            =   240
            TabIndex        =   49
            Top             =   240
            Width           =   975
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
         Begin VB.Label lbl���� 
            Caption         =   "����;�Ӵ�;24"
            Height          =   225
            Index           =   0
            Left            =   1350
            TabIndex        =   51
            Top             =   330
            Width           =   2115
         End
      End
   End
   Begin VB.Frame Fra�������� 
      Caption         =   "�������������2"
      Height          =   5625
      Left            =   5790
      TabIndex        =   17
      Top             =   120
      Width           =   5535
      Begin VB.Frame fraѡ������ʾ 
         Caption         =   "ѡ������ʾ"
         Height          =   5325
         Left            =   120
         TabIndex        =   18
         Top             =   210
         Width           =   5295
         Begin VB.CheckBox chk����ҩ��� 
            Caption         =   "��ʾ����ҩ���"
            Height          =   200
            Left            =   1890
            TabIndex        =   73
            Top             =   960
            Width           =   1665
         End
         Begin VB.TextBox txt�ѹ������� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2160
            TabIndex        =   58
            Text            =   "3"
            Top             =   3630
            Width           =   615
         End
         Begin VB.TextBox txt�ѹ������� 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   57
            Text            =   "1"
            Top             =   3630
            Width           =   615
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   5
            Left            =   240
            TabIndex        =   56
            Top             =   3215
            Width           =   975
         End
         Begin VB.CheckBox chk��ʾ�ѹ��� 
            Caption         =   "��ʾ�ѹ���"
            Height          =   200
            Left            =   120
            TabIndex        =   55
            Top             =   2970
            Width           =   1335
         End
         Begin VB.CommandButton cmdTimeoutColor 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   54
            Top             =   3215
            Width           =   975
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   300
            Left            =   660
            TabIndex        =   52
            Text            =   "1"
            Top             =   2610
            Width           =   615
         End
         Begin VB.CommandButton cmdPreparingColor 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   34
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   33
            Top             =   1175
            Width           =   975
         End
         Begin VB.CheckBox chk��ʾ����ҩ 
            Caption         =   "��ʾ����ҩ"
            Height          =   200
            Left            =   120
            TabIndex        =   32
            Top             =   1950
            Width           =   1335
         End
         Begin VB.CheckBox chk��ʾ����ҩ 
            Caption         =   "��ʾ����ҩ"
            Height          =   200
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox chk��ʾ�������� 
            Caption         =   "��ʾ��������"
            Height          =   200
            Left            =   120
            TabIndex        =   30
            Top             =   4110
            Width           =   1575
         End
         Begin VB.CheckBox chk��ʾ���� 
            Caption         =   "��ʾ����"
            Height          =   200
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdWin 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   28
            Top             =   485
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   485
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   2
            Left            =   240
            TabIndex        =   26
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   3
            Left            =   240
            TabIndex        =   25
            Top             =   2190
            Width           =   975
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "��������"
            Height          =   350
            Index           =   4
            Left            =   240
            TabIndex        =   24
            Top             =   4350
            Width           =   975
         End
         Begin VB.CommandButton cmdOther 
            Caption         =   "������ɫ"
            Height          =   350
            Left            =   3720
            TabIndex        =   23
            Top             =   4350
            Width           =   975
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2160
            TabIndex        =   22
            Text            =   "3"
            Top             =   2610
            Width           =   615
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   720
            TabIndex        =   21
            Text            =   "1"
            Top             =   1590
            Width           =   615
         End
         Begin VB.TextBox txt����ҩ���� 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2160
            TabIndex        =   20
            Text            =   "3"
            Top             =   1590
            Width           =   615
         End
         Begin VB.TextBox txtContent 
            Height          =   300
            Left            =   1200
            TabIndex        =   19
            Top             =   4830
            Width           =   3975
         End
         Begin VB.Shape shpTimeout 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label lbl�ѹ������� 
            Caption         =   "����"
            Height          =   195
            Left            =   1680
            TabIndex        =   63
            Top             =   3690
            Width           =   375
         End
         Begin VB.Label lbl�ѹ���Sum 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   3600
            TabIndex        =   62
            Top             =   3630
            Width           =   615
         End
         Begin VB.Label lbl�ѹ������� 
            Caption         =   "����"
            Height          =   195
            Left            =   3120
            TabIndex        =   61
            Top             =   3690
            Width           =   375
         End
         Begin VB.Label lbl�ѹ������� 
            Caption         =   "����"
            Height          =   195
            Left            =   240
            TabIndex        =   60
            Top             =   3690
            Width           =   375
         End
         Begin VB.Label lbl���� 
            Caption         =   "�ѹ���"
            Height          =   195
            Index           =   5
            Left            =   1320
            TabIndex        =   59
            Top             =   3293
            Width           =   2115
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   2670
            Width           =   375
         End
         Begin VB.Label lbl���� 
            Caption         =   "����ҩ"
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   46
            Top             =   1253
            Width           =   2115
         End
         Begin VB.Label lbl���� 
            Caption         =   "����ҩ"
            Height          =   195
            Index           =   3
            Left            =   1320
            TabIndex        =   45
            Top             =   2238
            Width           =   2115
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
            TabIndex        =   44
            Top             =   563
            Width           =   2120
         End
         Begin VB.Shape shpOther 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00408000&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   4380
            Width           =   375
         End
         Begin VB.Label lbl���� 
            Caption         =   "��������"
            Height          =   195
            Index           =   4
            Left            =   1320
            TabIndex        =   43
            Top             =   4425
            Width           =   2115
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   1680
            TabIndex        =   42
            Top             =   2670
            Width           =   375
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   3120
            TabIndex        =   41
            Top             =   2670
            Width           =   375
         End
         Begin VB.Label lbl����ҩSum 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   3600
            TabIndex        =   40
            Top             =   2610
            Width           =   615
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   240
            TabIndex        =   39
            Top             =   1650
            Width           =   375
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   3120
            TabIndex        =   38
            Top             =   1650
            Width           =   375
         End
         Begin VB.Label lbl����ҩSum 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Height          =   300
            Left            =   3600
            TabIndex        =   37
            Top             =   1590
            Width           =   615
         End
         Begin VB.Label lbl����ҩ���� 
            Caption         =   "����"
            Height          =   195
            Left            =   1680
            TabIndex        =   36
            Top             =   1650
            Width           =   375
         End
         Begin VB.Shape shpPreparing 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00408000&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   2185
            Width           =   375
         End
         Begin VB.Shape shpCalling 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   300
            Left            =   4800
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lblContent 
            Caption         =   "��ʾ����"
            Height          =   195
            Left            =   360
            TabIndex        =   35
            Top             =   4890
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   9090
      TabIndex        =   11
      Top             =   5910
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   7230
      TabIndex        =   10
      Top             =   5910
      Width           =   1100
   End
   Begin VB.Frame Fra��ʾ���� 
      Caption         =   "��ʾ����"
      Height          =   1575
      Left            =   210
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
            Height          =   864
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
      Left            =   90
      Top             =   5910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   0
      Left            =   90
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   1
      Left            =   600
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   2
      Left            =   1140
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   3
      Left            =   1680
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   4
      Left            =   2220
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Index           =   5
      Left            =   2730
      Top             =   5400
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
        Me.chk����ҩ���.Enabled = True
    Else
        cmdFont(2).Enabled = False
        Me.cmdColor.Enabled = False
        Me.txt����ҩ����.Enabled = False
        Me.txt����ҩ����.Enabled = False
        Me.chk����ҩ���.Enabled = False
    End If
End Sub

Private Sub chk��ʾ����ҩ_Click()
    If chk��ʾ����ҩ.Value = 1 Then
        cmdFont(3).Enabled = True
        Me.cmdPreparingColor.Enabled = True
        'Me.txt����ҩ����.Enabled = True
        Me.txt����ҩ����.Enabled = True
    Else
        cmdFont(3).Enabled = False
        Me.cmdPreparingColor.Enabled = False
        'Me.txt����ҩ����.Enabled = False
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

Private Sub chk��ʾ�ѹ���_Click()
    If chk��ʾ�ѹ���.Value = 1 Then
        cmdFont(3).Enabled = True
        Me.cmdTimeoutColor.Enabled = True
        Me.txt�ѹ�������.Enabled = True
    Else
        cmdFont(3).Enabled = False
        Me.cmdTimeoutColor.Enabled = False
        Me.txt�ѹ�������.Enabled = False
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
    If gobjComLib.ErrCenter = 1 Then
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
    SaveSetting "ZLSOFT", strReg, "������ɫ", shp����.FillColor
    SaveSetting "ZLSOFT", strReg, "��������ɫ", shpCall.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����", Me.chk��ʾ����.Value
    SaveSetting "ZLSOFT", strReg, "������ɫ", shpWin.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ��������", Me.chk��ʾ��������.Value
    SaveSetting "ZLSOFT", strReg, "����������ɫ", shpOther.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����ҩ", Me.chk��ʾ����ҩ.Value
    SaveSetting "ZLSOFT", strReg, "����ҩ���", Me.chk����ҩ���.Value
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.lbl����ҩSum.Caption
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Val(Me.txt����ҩ����.Text)
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Val(Me.txt����ҩ����.Text)
    SaveSetting "ZLSOFT", strReg, "����ҩ��ɫ", shpCalling.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����ҩ", Me.chk��ʾ����ҩ.Value
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Me.lbl����ҩSum.Caption
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Val(Me.txt����ҩ����.Text)
    SaveSetting "ZLSOFT", strReg, "����ҩ����", Val(Me.txt����ҩ����.Text)
    SaveSetting "ZLSOFT", strReg, "����ҩ��ɫ", shpPreparing.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ʾ�ѹ���", Me.chk��ʾ�ѹ���.Value
    SaveSetting "ZLSOFT", strReg, "�ѹ�������", Me.lbl�ѹ���Sum.Caption
    SaveSetting "ZLSOFT", strReg, "�ѹ�������", Val(Me.txt�ѹ�������.Text)
    SaveSetting "ZLSOFT", strReg, "�ѹ�������", Val(Me.txt�ѹ�������.Text)
    SaveSetting "ZLSOFT", strReg, "�ѹ�����ɫ", shpTimeout.FillColor
    
    SaveSetting "ZLSOFT", strReg, "��ҳʱ��", Val(Me.txt��ҳʱ��.Text)
    SaveSetting "ZLSOFT", strReg, "ˢ��ʱ��", Val(Me.txt��ѯʱ��.Text)
'    SaveSetting "ZLSOFT", strReg, "����ʱ��", Val(Me.txt����ʱ��.Text)
    
    SaveSetting "ZLSOFT", strReg, "��ʾ����", Me.txtContent.Text
    
    Unload Me
End Sub


Private Sub cmdOther_Click()
    dlgColor.Color = shpOther.FillColor
    dlgColor.ShowColor
    shpOther.FillColor = dlgColor.Color
End Sub

Private Sub cmdTimeoutColor_Click()
    dlgColor.Color = shpTimeout.FillColor
    dlgColor.ShowColor
    shpTimeout.FillColor = dlgColor.Color
End Sub

Private Sub cmdWin_Click()
    dlgColor.Color = shpWin.FillColor
    dlgColor.ShowColor
    shpWin.FillColor = dlgColor.Color
End Sub

Private Sub cmd����_Click()
    dlgColor.Color = shp����.FillColor
    dlgColor.ShowColor
    shp����.FillColor = dlgColor.Color
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
    
    shp����.FillColor = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
    shpCall.FillColor = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
    
    Me.chk��ʾ����.Value = GetSetting("ZLSOFT", strReg, "��ʾ����", 1)
    If Me.chk��ʾ����.Value = 1 Then
        cmdFont(1).Enabled = True
        Me.cmdWin.Enabled = True
    Else
        cmdFont(1).Enabled = False
        Me.cmdWin.Enabled = False
    End If
    shpWin.FillColor = GetSetting("ZLSOFT", strReg, "������ɫ", vbGreen)
    
    Me.chk��ʾ��������.Value = GetSetting("ZLSOFT", strReg, "��ʾ��������", 1)
    If chk��ʾ��������.Value = 1 Then
        cmdFont(4).Enabled = True
        Me.cmdOther.Enabled = True
    Else
        cmdFont(4).Enabled = True
        Me.cmdOther.Enabled = True
    End If
    shpOther.FillColor = GetSetting("ZLSOFT", strReg, "����������ɫ", vbGreen)
    
    Me.chk��ʾ����ҩ.Value = Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1"))
    If Me.chk��ʾ����ҩ.Value = 1 Then
        cmdFont(2).Enabled = True
        Me.cmdColor.Enabled = True
        Me.txt����ҩ����.Enabled = True
        Me.txt����ҩ����.Enabled = True
        Me.chk����ҩ���.Enabled = True
    Else
        cmdFont(2).Enabled = False
        Me.cmdColor.Enabled = False
        Me.txt����ҩ����.Enabled = False
        Me.txt����ҩ����.Enabled = False
        Me.chk����ҩ���.Enabled = False
    End If
    Me.chk����ҩ���.Value = Val(GetSetting("ZLSOFT", strReg, "����ҩ���", "0"))
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "5"))
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "2"))
    Me.lbl����ҩSum.Caption = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "10"))
    shpCalling.FillColor = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
    
    Me.chk��ʾ����ҩ.Value = Val(GetSetting("ZLSOFT", strReg, "��ʾ����ҩ", "1"))
    If Me.chk��ʾ����ҩ.Value = 1 Then
        cmdFont(3).Enabled = True
        Me.cmdPreparingColor.Enabled = True
        Me.txt����ҩ����.Enabled = True
    Else
        cmdFont(3).Enabled = False
        Me.cmdPreparingColor.Enabled = False
        Me.txt����ҩ����.Enabled = False
    End If
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "5"))
    Me.txt����ҩ����.Text = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "2"))
    Me.lbl����ҩSum.Caption = Val(GetSetting("ZLSOFT", strReg, "����ҩ����", "10"))
    shpPreparing.FillColor = GetSetting("ZLSOFT", strReg, "����ҩ��ɫ", vbGreen)
    
    Me.chk��ʾ�ѹ���.Value = Val(GetSetting("ZLSOFT", strReg, "��ʾ�ѹ���", "1"))
    If Me.chk��ʾ�ѹ���.Value = 1 Then
        cmdFont(3).Enabled = True
        Me.cmdTimeoutColor.Enabled = True
        Me.txt�ѹ�������.Enabled = True
    Else
        cmdFont(3).Enabled = False
        Me.cmdTimeoutColor.Enabled = False
        Me.txt�ѹ�������.Enabled = False
    End If
    Me.txt�ѹ�������.Text = Val(GetSetting("ZLSOFT", strReg, "�ѹ�������", "5"))
    Me.txt�ѹ�������.Text = Val(GetSetting("ZLSOFT", strReg, "�ѹ�������", "1"))
    Me.lbl�ѹ���Sum.Caption = Val(GetSetting("ZLSOFT", strReg, "�ѹ�������", "5"))
    shpTimeout.FillColor = GetSetting("ZLSOFT", strReg, "�ѹ�����ɫ", vbGreen)
    
    Me.txt��ҳʱ��.Text = GetSetting("ZLSOFT", strReg, "��ҳʱ��", "5")
    Me.txt��ѯʱ��.Text = GetSetting("ZLSOFT", strReg, "ˢ��ʱ��", "10")
'    Me.txt����ʱ��.Text = GetSetting("ZLSOFT", strReg, "����ʱ��", "10")
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
        txt����ҩ����.Text = txt����ҩ����.Text
        txt�ѹ�������.Text = txt����ҩ����.Text
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
        Me.lbl�ѹ���Sum.Caption = Val(Me.txt�ѹ�������.Text) * Val(Me.txt�ѹ�������.Text)
    End If
End Sub

Private Sub txt����ҩ����_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����ҩ����_KeyUp(KeyCode As Integer, Shift As Integer)
    'Me.txt����ҩ����.Text = Me.txt����ҩ����.Text
    
    If Me.txt����ҩ����.Text <> "" Then
        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
    End If
    
'    If Me.txt����ҩ����.Text <> "" Then
'        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
'    End If
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�ѹ�������_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.txt�ѹ�������.Text <> "" Then
        Me.lbl�ѹ���Sum.Caption = Val(Me.txt�ѹ�������.Text) * Val(Me.txt�ѹ�������.Text)
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
    'Me.txt����ҩ����.Text = Me.txt����ҩ����.Text
    
'    If Me.txt����ҩ����.Text <> "" Then
'        Me.lbl����ҩSum.Caption = Val(Me.txt����ҩ����.Text) * Val(Me.txt����ҩ����.Text)
'    End If
    
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

Private Sub txt�ѹ�������_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
