VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStudyListCfg 
   BorderStyle     =   0  'None
   Caption         =   "����б�����"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraListColorCfg 
      Caption         =   "�б���ɫ����"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7785
      Begin VB.Frame Frame2 
         Caption         =   "��ɫ��ʾ����"
         Height          =   615
         Left            =   3960
         TabIndex        =   37
         ToolTipText     =   "����б���������ɫ���ͣ�Ϊǰ��ɫʱ�����б��ǰ��ɫ����֮������ɫ��"
         Top             =   4680
         Width           =   2055
         Begin VB.OptionButton optListColorMark 
            Caption         =   "����ɫ"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   39
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optListColorMark 
            Caption         =   "ǰ��ɫ"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   840
         TabIndex        =   34
         Top             =   4680
         Width           =   2895
         Begin VB.CheckBox chkNameColColorCfg 
            Caption         =   "��������ɫ��ʾ"
            Height          =   180
            Left            =   120
            TabIndex        =   36
            ToolTipText     =   "������ɫ���ݲ���������ʾ��"
            Top             =   0
            Width           =   1800
         End
         Begin VB.CheckBox chkOrdinaryNameColColorCfg 
            Caption         =   "����ȱʡ����������ɫ"
            Height          =   255
            Left            =   600
            TabIndex        =   35
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   10
         Left            =   2655
         TabIndex        =   32
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtAudit 
         Height          =   270
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtStudy 
         Height          =   270
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "0"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox txtReport 
         Height          =   270
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   27
         Text            =   "0"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtCheckIn 
         Height          =   270
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   25
         Text            =   "0"
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtEnreg 
         Height          =   270
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   23
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "�ָ�Ĭ��(&D)"
         Height          =   375
         Left            =   6080
         TabIndex        =   21
         Top             =   5400
         Width           =   1335
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   9
         Left            =   5170
         TabIndex        =   19
         Top             =   3120
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   8
         Left            =   2650
         TabIndex        =   18
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   7
         Left            =   2650
         TabIndex        =   16
         Top             =   3120
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   6
         Left            =   2650
         TabIndex        =   14
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   5
         Left            =   5190
         TabIndex        =   12
         Top             =   4080
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   4
         Left            =   2650
         TabIndex        =   10
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   3
         Left            =   2655
         TabIndex        =   8
         Top             =   4080
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   2
         Left            =   5175
         TabIndex        =   6
         Top             =   3600
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   0
         Left            =   2650
         TabIndex        =   4
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "��"
         Height          =   255
         Index           =   1
         Left            =   2650
         TabIndex        =   2
         Top             =   960
         Width           =   255
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   10
         Left            =   1560
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�Ѳ��أ�"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   33
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "״̬��������        ������"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "״̬��������        ������"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "״̬��������        ������"
         Height          =   255
         Left            =   3120
         TabIndex        =   26
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "״̬��������        ������"
         Height          =   255
         Left            =   3120
         TabIndex        =   24
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "״̬��������        ������"
         Height          =   255
         Left            =   3120
         TabIndex        =   22
         Top             =   480
         Width           =   2415
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   9
         Left            =   4080
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�Ѿܾ���"
         Height          =   255
         Index           =   9
         Left            =   3360
         TabIndex        =   20
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape shpColor 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   8
         Left            =   1560
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�ѵǼǣ�"
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   7
         Left            =   1560
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "����ɣ�"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   15
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   6
         Left            =   1560
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "����ˣ�"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   13
         Top             =   2400
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   5
         Left            =   4095
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "����У�"
         Height          =   255
         Index           =   5
         Left            =   3375
         TabIndex        =   11
         Top             =   4080
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   4
         Left            =   1560
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�ѱ��棺"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   9
         Top             =   1920
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   3
         Left            =   1560
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�����У�"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   7
         Top             =   4080
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   2
         Left            =   4080
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�����У�"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   5
         Top             =   3600
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   1560
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�Ѽ�飺"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape shpColor 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   1560
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "�ѱ�����"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   2640
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmStudyListCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptId As Long     '��¼��ǰ����ID
Private mblnRefreshed  As Boolean '�жϸý����Ƿ��Ѿ�ˢ��

Private Sub chkNameColColorCfg_Click()
    If chkNameColColorCfg.value = 1 Then
        chkOrdinaryNameColColorCfg.Enabled = True
    Else
        chkOrdinaryNameColColorCfg.value = 0
        chkOrdinaryNameColColorCfg.Enabled = False
    End If
End Sub

Private Sub cmdColor_Click(Index As Integer)
    dlgColor.Color = shpColor(Index).FillColor
    dlgColor.ShowColor
    shpColor(Index).FillColor = dlgColor.Color
    
    '���������Ѿ�ˢ��
    mblnRefreshed = True
End Sub


Private Sub LoadDefaultCfg()
    shpColor(10).FillColor = ColorConstants.vbYellow
    shpColor(9).FillColor = ColorConstants.vbRed
    shpColor(7).FillColor = ColorConstants.vbGreen
    
    shpColor(0).FillColor = ColorConstants.vbWhite
    shpColor(1).FillColor = ColorConstants.vbWhite
    shpColor(2).FillColor = ColorConstants.vbWhite
    shpColor(3).FillColor = ColorConstants.vbWhite
    shpColor(4).FillColor = ColorConstants.vbWhite
    shpColor(5).FillColor = ColorConstants.vbWhite
    shpColor(6).FillColor = ColorConstants.vbWhite
    shpColor(8).FillColor = ColorConstants.vbWhite
    
    txtEnreg.Text = "0"
    txtCheckIn.Text = "0"
    txtStudy.Text = "0"
    txtReport.Text = "0"
    txtAudit.Text = "0"
End Sub

Private Sub cmdDefault_Click()
    Call LoadDefaultCfg
    
    mblnRefreshed = True
End Sub

Private Sub Form_Load()
    mblnRefreshed = False
    mlngDeptId = -1
End Sub

Private Sub Form_Resize()
   fraListColorCfg.Left = (Me.ScaleWidth - fraListColorCfg.Width) / 2
End Sub


Public Sub zlRefresh(lngDeptID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTemp As Long
             
    On Error GoTo err
    
    mlngDeptId = lngDeptID
    
    Call LoadDefaultCfg
    
    strSQL = "select ID ,����ID,������,����ֵ from Ӱ�����̲��� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId)
    
    While Not rsTemp.EOF
        Select Case rsTemp!������
            Case "�ѵǼ�"
                shpColor(8).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�ѱ���"
                shpColor(1).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "������"
                shpColor(2).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�Ѽ��"
                shpColor(0).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "������"
                shpColor(3).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�ѱ���"
                shpColor(4).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�����"
                shpColor(6).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�����"
                shpColor(7).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�����"
                shpColor(5).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�Ѿܾ�"
                shpColor(9).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�Ѳ���"
                shpColor(10).FillColor = Val(Nvl(rsTemp!����ֵ))
            Case "�ǼǺ�����"
                txtEnreg.Text = Val(Nvl(rsTemp!����ֵ))
            Case "����������"
                txtCheckIn.Text = Val(Nvl(rsTemp!����ֵ))
            Case "��������"
                txtStudy.Text = Val(Nvl(rsTemp!����ֵ))
            Case "���������"
                txtReport.Text = Val(Nvl(rsTemp!����ֵ))
            Case "��˺�����"
                txtAudit.Text = Val(Nvl(rsTemp!����ֵ))
            Case "��ɫ��ʾ����"
                If Val(Nvl(rsTemp!����ֵ)) = 0 Then
                    optListColorMark(0).value = True
                Else
                    optListColorMark(1).value = True
                End If
        End Select
        rsTemp.MoveNext
    Wend
    
    chkNameColColorCfg.value = Val(GetDeptPara(mlngDeptId, "������ɫ����", 0))
    If chkNameColColorCfg.value = 0 Then
        chkOrdinaryNameColColorCfg.value = 0
        chkOrdinaryNameColColorCfg.Enabled = False
    Else
        chkOrdinaryNameColColorCfg.Enabled = True
        chkOrdinaryNameColColorCfg.value = Val(GetDeptPara(mlngDeptId, "ȱʡ���Ͳ���������ɫ����", 0))
    End If
    
    mblnRefreshed = True
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub


Public Sub zlSave()
    Dim i As Integer, strInput As String
    Dim strSQL As String
    
    If Not mblnRefreshed Then Exit Sub      'û��ˢ���򲻱���
    If mlngDeptId < 0 Then Exit Sub
    
      
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�ѵǼ�','" & shpColor(8).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�ѱ���','" & shpColor(1).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '������','" & shpColor(2).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�Ѽ��','" & shpColor(0).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '������','" & shpColor(3).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�ѱ���','" & shpColor(4).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�����','" & shpColor(6).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�����','" & shpColor(7).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�����','" & shpColor(5).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�Ѿܾ�','" & shpColor(9).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�Ѳ���','" & shpColor(10).FillColor & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '�ǼǺ�����','" & Val(txtEnreg.Text) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '����������','" & Val(txtCheckIn.Text) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��������','" & Val(txtStudy.Text) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '���������','" & Val(txtReport.Text) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��˺�����','" & Val(txtAudit.Text) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '������ɫ����','" & chkNameColColorCfg.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", 'ȱʡ���Ͳ���������ɫ����','" & chkOrdinaryNameColColorCfg.value & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strSQL = "ZL_Ӱ�����̲���_UPDATE( " & mlngDeptId & ", '��ɫ��ʾ����','" & IIf(optListColorMark(0).value = True, 0, 1) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
End Sub


Private Sub txtAudit_Change()
    mblnRefreshed = True
End Sub

Private Sub txtCheckIn_Change()
    mblnRefreshed = True
End Sub

Private Sub txtEnreg_Change()
    mblnRefreshed = True
End Sub

Private Sub txtReport_Change()
    mblnRefreshed = True
End Sub

Private Sub txtStudy_Change()
    mblnRefreshed = True
End Sub
