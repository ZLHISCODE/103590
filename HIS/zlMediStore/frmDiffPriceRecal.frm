VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDiffPriceRecal 
   Caption         =   "ҩƷ��ʼ���"
   ClientHeight    =   7305
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   11760
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   5
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   10410
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   6
         Top             =   950
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   23
         Top             =   4500
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   22
         Top             =   4500
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   21
         Top             =   4500
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   20
         Top             =   4500
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ⷿ"
         Height          =   180
         Left            =   270
         TabIndex        =   19
         Top             =   660
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ��ʼ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   18
         Top             =   45
         Width           =   11535
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   17
         Top             =   4155
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   16
         Top             =   4440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   15
         Top             =   4440
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   14
         Top             =   4440
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   13
         Top             =   4440
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "����ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label txtCheckDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9600
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ʱ��"
         Height          =   180
         Left            =   8640
         TabIndex        =   9
         Top             =   660
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "�̵���ϼƣ�"
         Height          =   180
         Left            =   1920
         TabIndex        =   8
         Top             =   3840
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8730
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7410
      TabIndex        =   3
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   2
      Top             =   5040
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   5100
      Width           =   1815
   End
   Begin VB.CommandButton cmd�̶��� 
      Caption         =   "�̶���(&L)"
      Height          =   350
      Left            =   6090
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0000
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":021A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0434
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":064E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0868
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0A82
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0C9C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":0EB6
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":10D0
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":12EA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1504
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":171E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1938
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1B52
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1D6C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceRecal.frx":1F86
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   6945
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiffPriceRecal.frx":21A0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCode 
      Caption         =   "����ҩƷ"
      Height          =   180
      Left            =   1530
      TabIndex        =   25
      Top             =   5145
      Width           =   720
   End
End
Attribute VB_Name = "frmDiffPriceRecal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSelectStock As String           '�Ƿ��ѡ�ⷿ
Private mint�༭״̬ As Integer             '1����ʼ��棻2���ֹ�¼���ۣ�
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean                '��һ����ʾ
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭

Private Const mlngColorRed As Long = vbRed
Private Const mlngColorBlue As Long = vbBlue
Private Const mlngColorBlack As Long = vbBlack
Private mlngCurrColor As Long
Private mlngNextColor As Long
Private blnColorRefresh As Boolean

Private mstrMsg As String

Private mlng�ⷿ As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��
Private mintMaxMoneyBit As Integer          'ҩƷ�����н��С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntColҩ�� As Integer = 2
Private Const mconIntCol��Ʒ�� As Integer = 3
Private Const mconIntCol��� As Integer = 4
Private Const mconIntCol��� As Integer = 5
Private Const mconIntCol�ۼ۵�λ As Integer = 6
Private Const mconIntColС��λ���� As Integer = 7
Private Const mconIntColҩ�ⵥλ As Integer = 8
Private Const mconIntCol��λ���� As Integer = 9
Private Const mconIntCol����� As Integer = 10
Private Const mconIntCol����� As Integer = 11
Private Const mconIntColʵ�ʲ�� As Integer = 12
Private Const mconIntColS  As Integer = 13              '������

Private Function GetAllDrug() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strҩ�� As String
    
    On Error GoTo errHandle
    
    gstrSQL = "Select Distinct A.ҩƷid, '[' || E.���� || ']' As ҩƷ����, E.���� As ͨ����, C.���� As ��Ʒ��, E.���, E.���㵥λ As �ۼ۵�λ, S.������� As С��װ����," & _
        " A.ҩ�ⵥλ, S.������� / A.ҩ���װ As ���װ����, S.�����, S.����� " & _
        " From ҩƷ��� A, �շ���ĿĿ¼ E, �շ���Ŀ���� C, " & _
        " (Select ҩƷid, Sum(ʵ������) �������, Sum(ʵ�ʽ��) �����, Sum(ʵ�ʲ��) ����� " & _
        " From ҩƷ��� Where Nvl(�Ƿ��ʼ,0) = 1 " & _
        " Group By ҩƷid) S " & _
        " Where A.ҩƷid = E.ID And A.ҩƷid = C.�շ�ϸĿid(+) And C.����(+) = 3 And A.ҩƷid = S.ҩƷid " & _
        " Order By ҩƷ����"
        
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[��ȡ���н��ҩƷ]")
    
    If rsTmp.EOF Then
        GetAllDrug = False
        Exit Function
    End If
    
    Call initGrid
    
    Call FS.StopFlash
    
    With mshBill
        .Redraw = False
        Do While Not rsTmp.EOF
            .TextMatrix(.rows - 1, 0) = rsTmp!ҩƷid
            
            If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                strҩ�� = rsTmp!ͨ����
            Else
                strҩ�� = IIf(IsNull(rsTmp!��Ʒ��), rsTmp!ͨ����, rsTmp!��Ʒ��)
            End If
           
            .TextMatrix(.rows - 1, mconIntColҩ��) = rsTmp!ҩƷ���� & strҩ��
          
            .TextMatrix(.rows - 1, mconIntCol��Ʒ��) = IIf(IsNull(rsTmp!��Ʒ��), "", rsTmp!��Ʒ��)
                    
            .TextMatrix(.rows - 1, mconIntCol���) = rsTmp!���
            .TextMatrix(.rows - 1, mconIntCol�ۼ۵�λ) = rsTmp!�ۼ۵�λ
            .TextMatrix(.rows - 1, mconIntColС��λ����) = rsTmp!С��װ����
            .TextMatrix(.rows - 1, mconIntColҩ�ⵥλ) = rsTmp!ҩ�ⵥλ
            .TextMatrix(.rows - 1, mconIntCol��λ����) = rsTmp!���װ����
            .TextMatrix(.rows - 1, mconIntCol�����) = zlStr.FormatEx(rsTmp!�����, gtype_UserSysParms.P9_���ý���λ��)
            .TextMatrix(.rows - 1, mconIntCol�����) = zlStr.FormatEx(rsTmp!�����, gtype_UserSysParms.P9_���ý���λ��)

            Call zlControl.StaShowPercent(rsTmp.AbsolutePosition / rsTmp.RecordCount, staThis.Panels(2), frmDiffPriceRecal)
            rsTmp.MoveNext
            If Not rsTmp.EOF Then .rows = .rows + 1
        Loop
        Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
        .Redraw = True
    End With
    
    DoEvents
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    
    GetAllDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal int�༭״̬ As Integer)
    mblnSave = False
    mblnSuccess = False
    mint�༭״̬ = int�༭״̬
    mblnChange = False
    mblnFirst = True
        
    Set mfrmMain = FrmMain
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption & IIf(mint�༭״̬ = 2, "(���¼��)", "")
    
    Me.Show vbModal, FrmMain
    
End Sub
'��ʼ���༭�ؼ�
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        .ClearBill
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol�ۼ۵�λ) = "�ۼ۵�λ"
        .TextMatrix(0, mconIntColС��λ����) = "С��װ����"
        .TextMatrix(0, mconIntColҩ�ⵥλ) = "ҩ�ⵥλ"
        .TextMatrix(0, mconIntCol��λ����) = "���װ����"
        .TextMatrix(0, mconIntCol�����) = "�����"
        .TextMatrix(0, mconIntCol�����) = "�����"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 500
        .ColWidth(mconIntColҩ��) = 3000
        
        '��Ʒ���д���
        If gintҩƷ������ʾ = 2 Then
            '��ʾ��Ʒ����
            .ColWidth(mconIntCol��Ʒ��) = 2000
        Else
            '��������ʾ��Ʒ����
            .ColWidth(mconIntCol��Ʒ��) = 0
        End If
        
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol�ۼ۵�λ) = 800
        .ColWidth(mconIntColС��λ����) = 1000
        .ColWidth(mconIntColҩ�ⵥλ) = 800
        .ColWidth(mconIntCol��λ����) = 1000
        .ColWidth(mconIntCol�����) = 1000
        .ColWidth(mconIntCol�����) = 1000
        .ColWidth(mconIntColʵ�ʲ��) = 1000
        
        
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntColҩ��) = 5
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol�ۼ۵�λ) = 5
        .ColData(mconIntColС��λ����) = 5
        .ColData(mconIntColҩ�ⵥλ) = 5
        .ColData(mconIntCol��λ����) = 5
        .ColData(mconIntCol�����) = 5
        .ColData(mconIntCol�����) = 5
        .ColData(mconIntColʵ�ʲ��) = 4
                
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol�ۼ۵�λ) = flexAlignCenterCenter
        .ColAlignment(mconIntColС��λ����) = flexAlignRightCenter
        .ColAlignment(mconIntColҩ�ⵥλ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol��λ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�����) = flexAlignRightCenter
        .ColAlignment(mconIntColʵ�ʲ��) = flexAlignRightCenter
                
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
        
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    MsgBox "�Բ�����ʱû�а�����"
End Sub

Private Sub CmdSave_Click()
    Dim lngRow As Long
    Dim dbl��� As Double
    Dim lngҩƷID As Long
    Dim strTmp As String
    Dim intDrugCount As Integer
    
    On Error GoTo errHandle
    
    gcnOracle.BeginTrans
        
    With mshBill
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, 0)) > 0 And Val(.TextMatrix(lngRow, mconIntColʵ�ʲ��)) > 0 Then
                intDrugCount = intDrugCount + 1
                lngҩƷID = Val(.TextMatrix(lngRow, 0))
                
                '����С��װ������
                dbl��� = Round(Val(.TextMatrix(lngRow, mconIntColʵ�ʲ��)) / Val(.TextMatrix(lngRow, mconIntColС��λ����)), 7)
                
                strTmp = IIf(strTmp = "", "", strTmp & "|")
                strTmp = strTmp & lngҩƷID & "," & dbl���
                
                If intDrugCount > 99 Then
                    gstrSQL = "Zl_ҩƷ���_Update('" & strTmp & "' )"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
                    strTmp = ""
                    intDrugCount = 0
                End If
            End If
        Next
        If strTmp <> "" Then
            gstrSQL = "Zl_ҩƷ���_Update('" & strTmp & "' )"
            Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End With
    
    gcnOracle.CommitTrans
    MsgBox "��۱�����ϣ�", vbInformation + vbOKOnly, gstrSysName
    Unload Me
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_Activate()
    mshBill.ClearBill
    If GetAllDrug = False Then
        Exit Sub
        Unload Me
    End If
End Sub

'=========================================================================================
Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - cmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With
    
    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    
    txtCheckDate.Left = mshBill.Left + mshBill.Width - txtCheckDate.Width
    lblCheckDate.Left = txtCheckDate.Left - lblCheckDate.Width - 100
    
    LblStock.Left = mshBill.Left
    txtStock.Left = LblStock.Left + LblStock.Width + 100
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With
    
    With Lbl��������
        .Top = Lbl������.Top
        .Left = Txt������.Left + Txt������.Width + 250
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Txt�������
        .Top = Lbl������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = Lbl�������.Left - 200 - .Width
    End With
    
    With Lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = Pic����.TextWidth(.Caption) + 200
        
        lblCheckSum.Left = .Left + .Width + 100
        lblCheckSum.Top = .Top
        lblCheckSum.Width = Pic����.TextWidth(lblCheckSum.Caption) + 200
    End With
    
    With mshBill
        .Height = Pic����.Height - .Top - 100
    End With
    
    With cmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = cmdCancel.Left - .Width - 100
        .Top = cmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = cmdCancel.Top
    End With
        
    With lblCode
        .Top = cmdCancel.Top + 50
    End With
    With txtCode
        .Top = cmdCancel.Top + 30
    End With
    
    With cmd�̶���
        .Left = CmdSave.Left - .Width - 150
        .Top = CmdSave.Top
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        Select Case .Col
            Case mconIntColʵ�ʲ��
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select

    End With
End Sub


Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strkey = UCase(Trim(.Text))
        Select Case .Col
            Case mconIntColʵ�ʲ��
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ��𣬲�۽�����Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Abs(Val(strkey)) < 0.00001 Then
                        MsgBox "�Բ��𣬲�۽��ľ���ֵ���벻С��0.00001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "��۽�����С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strkey = zlStr.FormatEx(strkey, gtype_UserSysParms.P9_���ý���λ��, , True)
                    .Text = strkey
                    
                End If
        End Select
    End With
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        FindRow mshBill, mconIntColҩ��, txtCode.Text, True
    End If
End Sub


