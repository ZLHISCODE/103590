VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmCashPayAll 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ɿ��¼"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashPayAll.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtGroups 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      Left            =   6810
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   735
      Width           =   2490
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9360
      TabIndex        =   17
      Top             =   8805
      Width           =   9360
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   9330
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   420
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1530
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ����(&S)"
         Height          =   420
         Left            =   1650
         TabIndex        =   18
         Top             =   120
         Width           =   1530
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   420
         Left            =   7755
         TabIndex        =   14
         Top             =   120
         Width           =   1530
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   420
         Left            =   6225
         TabIndex        =   13
         Top             =   120
         Width           =   1530
      End
   End
   Begin VB.ComboBox cboTimes 
      Height          =   360
      Left            =   5190
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1257
      Width           =   1500
   End
   Begin VB.ComboBox cboType 
      Height          =   360
      Left            =   3880
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1257
      Width           =   1335
   End
   Begin VB.ComboBox cbo�ɿ�� 
      Height          =   360
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   750
      Width           =   2250
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   420
      Left            =   6810
      TabIndex        =   8
      ToolTipText     =   "�ȼ���F5"
      Top             =   1227
      Width           =   1530
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   750
      Width           =   1530
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1250
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   168951811
      CurrentDate     =   36904
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5490
      Left            =   80
      TabIndex        =   21
      Top             =   3240
      Width           =   10005
      Begin VB.TextBox txtRquareEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4050
         Width           =   1620
      End
      Begin VB.TextBox txtRquareEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4050
         Width           =   1635
      End
      Begin VB.TextBox txtLoanEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3615
         Width           =   1635
      End
      Begin VB.TextBox txtLoanEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3615
         Width           =   1620
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Index           =   2
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4545
         Width           =   8115
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   -120
         TabIndex        =   25
         Top             =   120
         Width           =   9390
      End
      Begin VB.TextBox txtItem 
         Height          =   360
         Index           =   3
         Left            =   1095
         MaxLength       =   50
         TabIndex        =   12
         Top             =   4950
         Width           =   5010
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Index           =   4
         Left            =   7215
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4950
         Width           =   1995
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Index           =   1
         Left            =   1095
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3165
         Width           =   4980
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshIncome 
         Height          =   3855
         Left            =   6180
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin ZL9BillEdit.BillEdit mshCash 
         Height          =   2490
         Left            =   30
         TabIndex        =   10
         Top             =   600
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   4392
         Enabled         =   -1  'True
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   3
         RowHeight0      =   360
         RowHeightMin    =   300
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lblLoan 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����ѿ�"
         Height          =   240
         Index           =   3
         Left            =   45
         TabIndex        =   36
         Top             =   4110
         Width           =   960
      End
      Begin VB.Label lblLoan 
         AutoSize        =   -1  'True
         Caption         =   "��ֵ"
         Height          =   240
         Index           =   2
         Left            =   3870
         TabIndex        =   38
         Top             =   4110
         Width           =   480
      End
      Begin VB.Label lblLoan 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   240
         Index           =   1
         Left            =   3870
         TabIndex        =   34
         Top             =   3675
         Width           =   480
      End
      Begin VB.Label lblLoan 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   240
         Index           =   0
         Left            =   540
         TabIndex        =   32
         Top             =   3675
         Width           =   480
      End
      Begin VB.Label lblItem 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ɿ�ϼ�"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   31
         Top             =   4605
         Width           =   960
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ϸ��"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   30
         Tag             =   "������ϸ��"
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "ժҪ(&D)"
         Height          =   240
         Index           =   6
         Left            =   210
         TabIndex        =   11
         Top             =   5010
         Width           =   840
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   240
         Index           =   7
         Left            =   6450
         TabIndex        =   29
         Top             =   5010
         Width           =   720
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ����"
         Height          =   240
         Index           =   4
         Left            =   330
         TabIndex        =   28
         Top             =   3225
         Width           =   720
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ϸ��"
         Height          =   240
         Index           =   3
         Left            =   6225
         TabIndex        =   27
         Tag             =   "������ϸ��"
         Top             =   285
         Width           =   1200
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRec 
      Height          =   1560
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   2752
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      MergeCells      =   3
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblGroups 
      AutoSize        =   -1  'True
      Caption         =   "��Ա����"
      Height          =   240
      Left            =   5820
      TabIndex        =   40
      Top             =   780
      Width           =   960
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   240
      Index           =   8
      Left            =   2880
      TabIndex        =   16
      Top             =   810
      Width           =   480
   End
   Begin VB.Label lblTimePeriod 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   7455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ɿ�Ǽǿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2707
      TabIndex        =   0
      Top             =   195
      Width           =   2250
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ɿ���"
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   810
      Width           =   720
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ֹʱ��"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   1317
      Width           =   960
   End
End
Attribute VB_Name = "frmCashPayAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gPayMoneyEdit
    PM_ȫ��ɿ� = 0
    PM_���սɿ� = 1
End Enum
Private mEditType As gPayMoneyEdit   '0-ȫ��ɿ�;1-���սɿ�

Private mblnOK  As Boolean
Private mstr�ɿ��� As String, mlng�ɿ���ID As Long
Private mstrLast As String '��¼������ȡ�ɹ��Ľ�ֹʱ��
Private mrsDetail As ADODB.Recordset
Private mrsTimes As ADODB.Recordset '��ǰ�ɿ����ڵ�ѡ����շ����͵Ľɿ����
Private mlng��ID As Long    '-1ʱ,������:�ݲ���
Private Const CONDFormat = "yyyy-MM-dd HH:mm:ss"
Private Const CONRecHead = "���|850|1,����|600|1,��ʼʱ��|2600|4,��ֹʱ��|2600|4"

Public Function ShowMe(ByVal strUser As String, ByVal lngUserID As Long, frmParent As Object, _
    Optional ByVal EditType As gPayMoneyEdit = PM_ȫ��ɿ�) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������(��ʾ��ɿ�)
    '���:strUser-�ɿ���Ա
    '       lngUserID-�ɿ���ԱID
    '       frmParent-���õ�������
    '       EditType-���ù���(0-ȫ��ɿ�;1-���սɿ�)
    '����:
    '����:
    '����:���˺�
    '����:2010-11-29 14:16:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType
    mstr�ɿ��� = strUser: mlng�ɿ���ID = lngUserID
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function
Private Sub zlSetDefaultDate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡʱ��
    '����:
    '����:���˺�
    '����:2009-10-14 14:26:23
    '�����:25752
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mEditType = PM_ȫ��ɿ� Then
        ''ȫ��ɿ�
        strSQL = "Select Max(��ֹʱ��) as ��ֹʱ�� From  �շ�����¼ Where �տ�Ա=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���)
        If IsNull(rsTemp!��ֹʱ��) Then
            dtpDate.Value = zlDatabase.Currentdate
        Else
            dtpDate.Value = CDate(Format(rsTemp!��ֹʱ��, CONDFormat))
        End If
    Else
        dtpDate.Value = CDate(Format(zlDatabase.Currentdate, dtpDate.CustomFormat))
        Call dtpDate_Change
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
'���ܣ���ʼ��������ʾ
    Set mrsDetail = New ADODB.Recordset
    mstrLast = "" '���Ϊ��
    
    If Not Visible Then
        If mEditType = PM_ȫ��ɿ� Then 'ȫ��ɿ�
            lblTimePeriod.Visible = False
            cboType.Visible = False
            cboTimes.Visible = False
            
            lblItem(1).Caption = "��ֹʱ��"
            dtpDate.CustomFormat = CONDFormat
            cmdRefresh.Left = dtpDate.Left + dtpDate.Width + 100
            
            lblTimePeriod.Visible = True
            mshRec.Visible = False
            fraMain.Top = lblTimePeriod.Top + lblTimePeriod.Height
            Me.Height = Me.Height - (mshRec.Height - lblTimePeriod.Height)
            
        Else '���սɿ�
            lblTimePeriod.Visible = True
            cboType.Visible = True
            cboTimes.Visible = True
            
            lblItem(1).Caption = "�ɿ�����"
            dtpDate.CustomFormat = "yyyy-MM-dd 00:00:00"
            dtpDate.Width = txtItem(0).Width
            cboType.Left = dtpDate.Left + dtpDate.Width + 100
            cboTimes.Left = cboType.Left + cboType.Width + 100
            cmdRefresh.Left = cboTimes.Left + cboTimes.Width + 100
            
            lblTimePeriod.Visible = False
            mshRec.Visible = True
        End If
    End If
    
    lblItem(2).Caption = lblItem(2).Tag
    lblItem(3).Caption = lblItem(3).Tag
    txtItem(1).Text = ""
    txtItem(2).Text = ""
    
    With mshCash
        .AllowAddRow = False
        .Font.Size = 12
        .TxtEditFont.Size = 12
        .Active = False
        
        .ClearBill
        .RowHeight(0) = .RowHeightMin
        .TextMatrix(0, 0) = "���㷽ʽ"
        .TextMatrix(0, 1) = "���"
        .TextMatrix(0, 2) = "�����"
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        
        .MsfObj.ColAlignmentFixed(0) = 4
        .MsfObj.ColAlignmentFixed(1) = 4
        .MsfObj.ColAlignmentFixed(2) = 4
        
        .ColWidth(0) = 1300
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        
        .ColData(2) = 4
    End With
    
    With mshIncome
        .Clear
        .Rows = 2
        .TextMatrix(0, 0) = "������Ŀ"
        .TextMatrix(0, 1) = "���"
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColWidth(0) = 1350
        .ColWidth(1) = 1350
    End With
End Sub

Private Sub cboType_Click()
    Call LoadCashTimes(mstr�ɿ���, dtpDate.Value, Val(cboType.ItemData(cboType.ListIndex)))
End Sub

Private Sub cmdRefresh_Click()
'���ܣ�������ȡ��ǰ�ɿ���Ա�Ľɿ���ϸ
    Dim rsTmp As New ADODB.Recordset, strIF As String
    Dim strSQL As String, strSub As String, i As Long, bytFlag As Byte
    Dim datBegin As Date, datEnd As Date, strTable As String, strWhere As String
    Dim cur�ɿ�ϼ� As Currency, cur����ϼ� As Currency
    Dim cur����ϼ� As Currency, curԤ���ϼ� As Currency
    Dim dbl��� As Double, dbl��� As Double
    Dim dbl�������� As Double, dbl�˿��� As Double, dbl��ֵ�� As Double
    
    If dtpDate.Value > zlDatabase.Currentdate Then
        MsgBox "�ɿ��ֹʱ�䲻ӦԽ����ǰϵͳʱ�䡣", vbInformation, gstrSysName
        dtpDate.SetFocus: Exit Sub
    End If
    
    Call InitFace
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
    If mEditType = PM_���սɿ� Then
        bytFlag = Val(cboType.ItemData(cboType.ListIndex))  '0-ȫ��,Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨',6,'���ѿ�')
        If bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5 Then strIF = " And ��¼���� = " & IIf(bytFlag = 3, 1, bytFlag)
    End If
    
    '����:42376:ִ��״̬<>9
    '��ȡ�ɿ�Ա�ϴνɿ��ֹʱ��
    If mEditType = PM_���սɿ� Then
        '���սɿ�
        If cboTimes.ListIndex = 0 Then
            strSQL = "Select Min(��ʼʱ��) ��ʼʱ��,Max(��ֹʱ��) ��ֹʱ�� From �շ�����¼ Where �տ�Ա=[1] And ����=[2]"
            If cboType.ListIndex > 0 Then strSQL = strSQL & " And ���� = [3]"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, dtpDate.Value, bytFlag)
            If rsTmp.RecordCount > 0 Then
                If Not IsNull(rsTmp!��ʼʱ��) Then datBegin = DateAdd("s", -1, rsTmp!��ʼʱ��)
                If Not IsNull(rsTmp!��ֹʱ��) Then datEnd = rsTmp!��ֹʱ��
            End If
            If datBegin = CDate(0) Then datBegin = CDate(Format(DateAdd("d", -1, dtpDate.Value), "yyyy-MM-dd ") & "23:59:59")
            If datEnd = CDate(0) Then datEnd = CDate(Format(dtpDate.Value, "yyyy-MM-dd ") & "23:59:59")
        Else    '�ض���cboType.ListIndex > 0
            mrsTimes.Filter = "����=" & cboTimes.ListIndex
            datBegin = mrsTimes!��ʼʱ��
            datEnd = mrsTimes!��ֹʱ��
        End If
        If zlDatabase.DateMoved(Format(datBegin, CONDFormat), , , Me.Caption) Then
            MsgBox "�ϴνɿ�ʱ��:" & Format(datBegin, CONDFormat) & vbCrLf & "�������һ����ʷ����ת��֮ǰ,Ҫ�ɿ�Ĳ���������ת���! " & _
                vbCrLf & "����ϵͳ����Ա��ϵת������,�����״νɿ�����ֹ��ɿʽ!", vbInformation, gstrSysName
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        
    Else
        'ȫ��ɿ�
        strSQL = "Select Max(��ֹʱ��) as ��ֹʱ�� From ��Ա�ɿ��¼ Where �տ�Ա=[1] And ��ֹʱ�� is Not NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���)
        If rsTmp.RecordCount > 0 Then
            If Not IsNull(rsTmp!��ֹʱ��) Then datBegin = rsTmp!��ֹʱ��
        End If
        If datBegin = CDate(0) Then
            datBegin = CDate(Format("1990-01-01 00:00:00", CONDFormat))
            
            If zlDatabase.DateMoved(Format(datBegin, CONDFormat), , , Me.Caption) Then
                If MsgBox("��ǰ�տ�Աδ�ɹ���,������ϴ���ʷ����ת��֮ǰ�����տ�����,��ǰ���ݿ��ܲ���������ȷ��Ҫ������", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Else
            If zlDatabase.DateMoved(Format(datBegin, CONDFormat), , , Me.Caption) Then
                MsgBox "�ϴνɿ�ʱ��:" & Format(datBegin, CONDFormat) & vbCrLf & "�������һ����ʷ����ת��֮ǰ,Ҫ�ɿ�Ĳ���������ת��󱸱�! " & _
                    vbCrLf & "����ϵͳ����Ա��ϵת������,�����״νɿ�����ֹ��ɿʽ!", vbInformation, gstrSysName
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        datEnd = dtpDate.Value
        lblTimePeriod.Caption = "ʱ�䷶Χ:" & Format(datBegin, CONDFormat) & "~" & Format(datEnd, CONDFormat)
    End If
    
        
    
    '��ȡ�ýɿ�Ա���θ��ֽ��㷽ʽ��Ӧ�ɽ��
    '-----------------------------------------------------------------------------------------------
    '�շѲ��ݣ��շѡ��Һš��������շѳ�Ԥ��(������ʾ)
    strSQL = ""
    If mEditType = PM_ȫ��ɿ� Or _
        mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5) Then
           strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.����ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=1"
        '0-ȫ��ɿ�;1-���սɿ�
        'bytFlag: 0-ȫ��,Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨')
       If bytFlag = 5 And mEditType = PM_���սɿ� Then
            '���սɿ���Ϊ��ֻͳ�ƾ��￨
            strTable = "" & _
            " Select Distinct ����ID" & _
            " From סԺ���ü�¼ A" & _
            " Where Nvl(���ʷ���,0)=0 And ��¼״̬<>0 And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3]" & strIF & _
            "        And Not Exists(" & strSub & ") "
        ElseIf InStr(1, "05", bytFlag) = 0 And mEditType = PM_���սɿ� Then
            '���սɿ�,���Ǿ��￨��ȫ���ɿ�
            strTable = "" & _
            " Select Distinct ����ID" & _
            " From ������ü�¼ A" & _
            " Where Nvl(���ʷ���,0)=0  and nvl(����״̬,0)<>1 And ��¼״̬<>0 And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3]" & strIF & _
            "           And Not Exists(" & strSub & ") "
        Else
            strTable = "" & _
            " Select Distinct ����ID" & _
            " From סԺ���ü�¼ A" & _
            " Where Nvl(���ʷ���,0)=0 And ��¼״̬<>0 And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3]" & strIF & _
            "        And Not Exists(" & strSub & ") " & _
            " Union  " & _
            " Select Distinct ����ID" & _
            " From ������ü�¼ A" & _
            " Where Nvl(���ʷ���,0)=0 And ��¼״̬<>0 and nvl(����״̬,0)<>1 And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3]" & strIF & _
            "        And Not Exists(" & strSub & ") "
        End If
        
        strSQL = _
        " Select Decode(Mod(B.��¼����,10),1,'[��Ԥ����]',B.���㷽ʽ) as ���㷽ʽ,Sum(B.��Ԥ��) as ���" & _
        " From ( " & strTable & ") A,����Ԥ����¼ B" & _
        " Where A.����ID=B.����ID And nvl(B.У�Ա�־,0) =0" & _
        " Group by Decode(Mod(B.��¼����,10),1,'[��Ԥ����]',B.���㷽ʽ)"
    End If
        
    '���ʲ��ݣ����ʲ�����ʳ�Ԥ��(������ʾ)
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 2) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=2"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Decode(Mod(B.��¼����,10),1,'[��Ԥ����]',B.���㷽ʽ) as ���㷽ʽ,Sum(B.��Ԥ��) as ���" & _
            " From ���˽��ʼ�¼ A,����Ԥ����¼ B" & _
            " Where A.ID=B.����ID And A.����״̬ Is Null And A.����Ա����||''=[1] And A.�շ�ʱ��>[2] And A.�շ�ʱ��<=[3] And Not Exists(" & strSub & ")" & _
            " Group by Decode(Mod(B.��¼����,10),1,'[��Ԥ����]',B.���㷽ʽ)"
    End If
    
    '��Ԥ�����ݣ�ֱ����Ԥ��
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 1) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=3"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            "   Select ���㷽ʽ,Sum(���) as ���" & _
            "   From ����Ԥ����¼ A" & _
            "   Where ��¼����=1 And ����Ա����||''=[1] And �տ�ʱ��>[2] And �տ�ʱ��<=[3] And Not Exists(" & strSub & ") " & _
            "   Group by ���㷽ʽ"
    End If
    
    '���ѿ�:
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 6) Then
        '0-ȫ��,Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨',6,'���ѿ�')
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=6"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.���㷽ʽ,Sum(A.ʵ�ս��) as ���" & _
            " From ���˿������¼ A, ���˿������¼ B" & _
            " Where a.������� = b.�������(+) And a.��¼���� In (1, 3) And b.��¼����(+) = 3 And a.����Ա����||''=[1] And a.�Ǽ�ʱ��>[2] And a.�Ǽ�ʱ��<=[3] And Not Exists(" & strSub & ") " & _
            " Group by ���㷽ʽ"
        
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=5"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.���㷽ʽ, Sum(A.ʵ�ս��) as ���" & _
            " From ���˿������¼ A, ���˿������¼ B" & _
            " Where a.������� = b.�������(+) And a.��¼���� In (2, 3) And b.��¼����(+) = 3 And a.����Ա����||''=[1] And a.�Ǽ�ʱ��>[2] And a.�Ǽ�ʱ��<=[3] And Not Exists(" & strSub & ") " & _
            " Group by ���㷽ʽ"
    End If
    
 
     
    '����ͳ��
    '-----------------------------------------------------------------------------------------------
    'ֱ�ӿۼ��ֽ�
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=4"
        
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
             "Select A.���㷽ʽ, Sum(nvl(a.�����,0)) as ���" & _
            " From ��Ա����¼ A" & _
            " Where  A.�����||''=[1] And A.���ʱ��>[2] And A.���ʱ��<=[3] And A.ȡ��ʱ�� is NULL And Not Exists(" & strSub & ")" & _
            " Group by A.���㷽ʽ"
        
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
             "Select A.���㷽ʽ,-1* Sum(nvl(a.�����,0)) as ���" & _
            " From ��Ա����¼ A" & _
            " Where  A.�����||''=[1] And A.���ʱ��>[2] And A.���ʱ��<=[3] And A.ȡ��ʱ�� is NULL And Not Exists(" & strSub & ")" & _
            " Group by A.���㷽ʽ"
    End If
    
    strSQL = "Select ���㷽ʽ,Sum(���) as ��� From (" & strSQL & ") Group by ���㷽ʽ Having Sum(���)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
    If Not rsTmp.EOF Then
        With mshCash
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Nvl(rsTmp!���㷽ʽ)
                .TextMatrix(i, 1) = Format(Nvl(rsTmp!���, 0), "0.00")
                cur����ϼ� = cur����ϼ� + Nvl(rsTmp!���, 0)
                If Nvl(rsTmp!���㷽ʽ) <> "[��Ԥ����]" Then
                    cur�ɿ�ϼ� = cur�ɿ�ϼ� + Nvl(rsTmp!���, 0)
                End If
                rsTmp.MoveNext
            Next
        End With
    End If
  
    
    '��ȡ�ýɿ�Ա���νɿ��Ӧ��������
    '-----------------------------------------------------------------------------------------------
    '�շѲ��ݣ��շ�,�Һ�,����,�շѳ�Ԥ��
    strSQL = ""
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.����ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=1"
        strWhere = " Where Nvl(���ʷ���,0)=0  and nvl(ִ��״̬,0)<>9 And ��¼״̬<>0 And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3]" & strIF & _
        "        And Not Exists(" & strSub & ")  "
        
        '0-ȫ��ɿ�;1-���սɿ�
        'bytFlag: 0-ȫ��,Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨')
        If bytFlag = 5 And mEditType = PM_���սɿ� Then
             '���սɿ���Ϊ��ֻͳ�ƾ��￨
             strTable = "סԺ���ü�¼"
         ElseIf InStr(1, "05", bytFlag) = 0 And mEditType = PM_���սɿ� Then
             '���սɿ�,���Ǿ��￨��ȫ���ɿ�
             strTable = "������ü�¼"
         Else
            strTable = " ( " & _
             " Select ������ĿID,Sum(���ʽ��) as ���ʽ��" & _
             " From סԺ���ü�¼ A" & _
             " Where Nvl(���ʷ���,0)=0 And ��¼״̬<>0 And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3]" & strIF & _
             "        And Not Exists(" & strSub & ")  " & _
             " Group by ������ĿID  " & _
             " Union ALL " & _
             " Select ������ĿID,Sum(���ʽ��) as ���ʽ��" & _
             " From ������ü�¼ A" & _
             " Where Nvl(���ʷ���,0)=0 And ��¼״̬<>0 and nvl(A.����״̬,0)<>1 And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3]" & strIF & _
             "        And Not Exists(" & strSub & ")  " & _
             " Group by ������ĿID ) "
             strWhere = ""
         End If
        strSQL = _
        " Select ������ĿID,Sum(���ʽ��) as ���" & _
        " From " & strTable & " A" & _
                 strWhere & _
        " Group by ������ĿID"
    End If
        
    '���ʲ��ݣ����ʲ���,���ʳ�Ԥ��
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 2) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=2"
        
        'bytFlag: 0-ȫ��,Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨')
        '���ܴ���������ʵ����,��ˣ�����ȫ����
        strTable = "" & _
        " Select B.������ĿID,Sum(B.���ʽ��) as ���ʽ��" & _
        " From ���˽��ʼ�¼ A,������ü�¼ B" & _
        " Where B.���ʷ���=1 And A.����״̬ Is Null And A.ID=B.����ID And A.����Ա����||''=[1] And A.�շ�ʱ��>[2] And A.�շ�ʱ��<=[3]" & _
        "       And Not Exists(" & strSub & ") " & _
        " Group by B.������ĿID " & _
        " Union ALL  " & _
        " Select B.������ĿID,Sum(B.���ʽ��) as ���ʽ��" & _
        " From ���˽��ʼ�¼ A,סԺ���ü�¼ B" & _
        " Where B.���ʷ���=1 And A.����״̬ Is Null And A.ID=B.����ID And A.����Ա����||''=[1] And A.�շ�ʱ��>[2] And A.�շ�ʱ��<=[3]" & _
        "       And Not Exists(" & strSub & ") " & _
        " Group by B.������ĿID"
        
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.������ĿID,Sum(A.���ʽ��) as ���" & _
            " From  (" & strTable & ") A" & _
            " Group by A.������ĿID"
    End If
    
    If strSQL <> "" Then
        strSQL = "Select B.����,B.����,Sum(A.���) as ��� From (" & strSQL & ") A,������Ŀ B" & _
            " Where A.������ĿID=B.ID Group by B.����,B.���� Having Sum(A.���)<>0 Order by B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
        If Not rsTmp.EOF Then
            With mshIncome
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(i, 0) = Nvl(rsTmp!����)
                    .TextMatrix(i, 1) = Format(Nvl(rsTmp!���, 0), "0.00")
                    cur����ϼ� = cur����ϼ� + Nvl(rsTmp!���, 0)
                    rsTmp.MoveNext
                Next
            End With
        End If
    End If
    
    '��ȡ�ýɿ�Ա���νɿ��Ӧ��Ԥ���տ�
    '-----------------------------------------------------------------------------------------------
    '��Ԥ�����ݣ�ֱ����Ԥ��
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 1) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=3"
        strSQL = "Select Sum(���) as ���" & _
            " From ����Ԥ����¼ A" & _
            " Where ��¼����=1 And ����Ա����||''=[1] And �տ�ʱ��>[2] And �տ�ʱ��<=[3] And Not Exists(" & strSub & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
        If Not rsTmp.EOF Then curԤ���ϼ� = Nvl(rsTmp!���, 0)
    End If
    
   '���ѿ�:
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 6) Then
        '0-ȫ��,Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨',6,'���ѿ�')
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=6"
        strSQL = "" & _
            " Select  Sum(A.ʵ�ս��) as ���" & _
            " From ���˿������¼ A, ���˿������¼ B" & _
            " Where a.������� = b.�������(+) And a.��¼���� In (1, 3) And b.��¼����(+) = 3 And a.����Ա����||''=[1] And a.�Ǽ�ʱ��>[2] And a.�Ǽ�ʱ��<=[3] And Not Exists(" & strSub & ") "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl�������� = Nvl(rsTmp!���, 0)
        
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=5"
        strSQL = "" & _
            " Select  Sum(A.ʵ�ս��) as ���" & _
            " From ���˿������¼ A, ���˿������¼ B" & _
            " Where a.������� = b.�������(+) And a.��¼���� In (2, 3) And b.��¼����(+) = 3 And a.����Ա����||''=[1] And a.�Ǽ�ʱ��>[2] And a.�Ǽ�ʱ��<=[3] And Not Exists(" & strSub & ") "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl��ֵ�� = Nvl(rsTmp!���, 0)
    End If
        
        
    '����ͳ��
    '-----------------------------------------------------------------------------------------------
    'ֱ�ӿۼ��ֽ�
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=4"
        strSQL = "Select Sum(�����) as ���" & _
            " From ��Ա����¼ A" & _
            " Where  �����||''=[1] And ���ʱ��>[2] And ���ʱ��<=[3] And ȡ��ʱ�� is NULL And Not Exists(" & strSub & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl��� = Nvl(rsTmp!���, 0)
        
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=4"
        strSQL = "Select Sum(�����) as ���" & _
            " From ��Ա����¼ A" & _
            " Where  �����||''=[1] And ���ʱ��>[2] And ���ʱ��<=[3] And ȡ��ʱ�� is NULL And Not Exists(" & strSub & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl��� = Nvl(rsTmp!���, 0)
    End If
    
    
    
    '��ȡ�ýɿ�Ա���νɿ����
    '-----------------------------------------------------------------------------------------------
    '�շѲ��ݣ��շ�,�Һ�,����,�շѳ�Ԥ��
    strSQL = ""
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.����ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=1"
        '0-ȫ��ɿ�;1-���սɿ�
        '����:44344
        'bytFlag: 0-ȫ��,Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨')
        strWhere = " where Nvl(���ʷ���,0)=0 And ��¼״̬<>0 and nvl(ִ��״̬,0)<>9   And ����Ա����||''=[1] And �Ǽ�ʱ��>[2] And �Ǽ�ʱ��<=[3] And Not Exists(" & strSub & ")" & strIF
        If bytFlag = 5 And mEditType = PM_���սɿ� Then
             '���սɿ���Ϊ��ֻͳ�ƾ��￨
             strTable = "סԺ���ü�¼"
         ElseIf InStr(1, "05", bytFlag) = 0 And mEditType = PM_���սɿ� Then
             '���սɿ�,���Ǿ��￨��ȫ���ɿ�
             strTable = "������ü�¼"
         Else
            strTable = " ( " & _
             " Select Distinct ����ID" & _
             " From סԺ���ü�¼ A" & _
               strWhere & _
             "  " & _
             " Union ALL " & _
             " Select Distinct ����ID " & _
             " From ������ü�¼ A" & _
               strWhere & _
             " ) "
             strWhere = ""
         End If
        
        strSQL = _
            " Select Distinct 1 as ����,����ID as ��¼ID" & _
            " From " & strTable & " A" & _
            "  " & strWhere
    End If
        
    '���ʲ��ݣ����ʲ���,���ʳ�Ԥ��
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 2) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=2"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 2 as ����,ID as ��¼ID" & _
            " From ���˽��ʼ�¼ A" & _
            " Where ����Ա����||''=[1] And �շ�ʱ��>[2] And �շ�ʱ��<=[3] And ����״̬ Is Null And Not Exists(" & strSub & ")"
    End If
    
    '��Ԥ�����ݣ�ֱ����Ԥ��
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 1) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=3"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 3 as ����,ID as ��¼ID" & _
            " From ����Ԥ����¼ A" & _
            " Where ��¼����=1 And ����Ա����||''=[1] And �տ�ʱ��>[2] And �տ�ʱ��<=[3] And Not Exists(" & strSub & ")"
    End If
    
    '-���ѿ�
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� And (bytFlag = 0 Or bytFlag = 6) Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=6"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 6 as ����,ID as ��¼ID" & _
            " From ���˿������¼ A, ���˿������¼ B" & _
            " Where a.������� = b.�������(+) And a.��¼���� In (1, 3) And b.��¼����(+) = 3 And a.����Ա����||''=[1] And a.�Ǽ�ʱ��>[2] And a.�Ǽ�ʱ��<=[3] And Not Exists(" & strSub & ")"
        
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=5"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 5 as ����,ID as ��¼ID" & _
            " From ���˿������¼ A, ���˿������¼ B" & _
            " Where a.������� = b.�������(+) And a.��¼���� In (2, 3) And b.��¼����(+) = 3 And a.����Ա����||''=[1] And a.�Ǽ�ʱ��>[2] And a.�Ǽ�ʱ��<=[3] And   Not Exists(" & strSub & ")  "
        
    End If
 
    '��ȡ����¼����:0-ȫ��ɿ�;1-���սɿ�
    '   ���ݰ��սɿ���ϰ�ʱ�䣬����ȫ��ɿ�Ľ�ֹʱ�䣬ͳ����Ӧʱ�䷶Χ�ڵ�"��Ա����¼"�� _
    '��û��ȡ���Ľ��ʱ��Ϊ׼�������ֽ���㷽ʽ������Ϊ�ɿ����
    If mEditType = PM_ȫ��ɿ� Or mEditType = PM_���սɿ� Then
        strSub = "Select Y.��¼ID From ��Ա�ɿ��¼ X,��Ա�ɿ���� Y Where Y.��¼ID=A.ID And X.�տ�Ա=[1] And X.ID=Y.����ID And Y.����=4"
        '���
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 4 as ����,ID as ��¼ID" & _
            " From  ��Ա����¼ A" & _
            " Where �����||''=[1] And ���ʱ��>[2] And ���ʱ��<=[3] and ȡ��ʱ�� is NULL And Not Exists(" & strSub & ")"
        
        '���
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 4 as ����,ID as ��¼ID" & _
            " From  ��Ա����¼ A" & _
            " Where �����||''=[1] And ���ʱ��>[2] And ���ʱ��<=[3] And ȡ��ʱ�� is NULL And Not Exists(" & strSub & ")"
    End If
    
    strSQL = "Select /*+ Rule*/ ����,��¼ID From (" & strSQL & ")"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�ɿ���, datBegin, datEnd)
    
    '��ʾ�ϼ���Ϣ
    '-----------------------------------------------------------------------------------------------
    lblItem(2).Caption = lblItem(2).Tag & Format(cur����ϼ�, "0.00")
    lblItem(3).Caption = lblItem(3).Tag & Format(cur����ϼ�, "0.00")
    txtItem(1).Text = Format(curԤ���ϼ�, "0.00")
    
    txtLoanEdit(0).Text = Format(dbl���, "0.00")
    txtLoanEdit(1).Text = Format(dbl���, "0.00")
    
    txtRquareEdit(0).Text = Format(dbl�������� - dbl�˿���, "0.00")
    txtRquareEdit(1).Text = Format(dbl��ֵ��, "0.00")
    
    If cur�ɿ�ϼ� <> 0 Then
        txtItem(2).Text = Format(cur�ɿ�ϼ�, "0.00Ԫ") & " ��" & zlCommFun.UppeMoney(cur�ɿ�ϼ�) & "��"
    Else
        txtItem(2).Text = Format(cur�ɿ�ϼ�, "0.00Ԫ")
    End If
    
    '��ǳɹ�
    mshCash.Active = True
    mshCash.Row = 1: mshCash.Col = 2
    mstrLast = "To_Date('" & Format(datEnd, CONDFormat) & "','YYYY-MM-DD HH24:MI:SS')"
    Screen.MousePointer = 0
    mshCash.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    Dim arrSQL() As Variant, i As Long, k As Long
    Dim strDate As String, lng����ID As Long, lngID As Long, blnTrans As Boolean
    Dim dblSumMoney As Double
    
    If InStr(txtItem(3).Text, "'") > 0 Then
        MsgBox "ժҪ��Ϣ�а����Ƿ����ַ���", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtItem(3).Text) > txtItem(3).MaxLength Then
        MsgBox "ժҪ��Ϣ�а���������̫�࣬������� " & txtItem(3).MaxLength \ 2 & " �����ֻ� " & txtItem(3).MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    
    With mshCash
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 And .TextMatrix(i, 0) <> "[��Ԥ����]" Then
                If InStr(.TextMatrix(i, 2), "'") > 0 Then
                    MsgBox "���㷽ʽ [" & .TextMatrix(i, 0) & "] �Ľ�����а����Ƿ��ַ���", vbInformation, gstrSysName
                    .Row = i: .Col = 2: .SetFocus: Exit Sub
                ElseIf zlCommFun.ActualLen(.TextMatrix(i, 2)) > 10 Then
                    MsgBox "���㷽ʽ [" & .TextMatrix(i, 0) & "] �Ľ���Ź������������10���ַ���", vbInformation, gstrSysName
                    .Row = i: .Col = 2: .SetFocus: Exit Sub
                End If
                dblSumMoney = dblSumMoney + Val(.TextMatrix(i, 1))
'                If Val(.TextMatrix(i, 1)) < 0 Then
'                    MsgBox "���㷽ʽ [" & .TextMatrix(i, 0) & "] �Ľ���Ϊ������", vbInformation, gstrSysName
'                    .Row = i: .Col = 1: .SetFocus: Exit Sub
'                End If
                k = k + 1
            End If
        Next
    End With
    
    '���˺� ����:????����ʡ����ҽԺ   ����:2010-12-06 11:09:29
    '       ʵ����������ܶ�Ҳ����ָ���������ģ���������һ��֧Ʊ���յ�Ǯû���˵Ķ��ʱ��
    '    '���˺�:25694,
    '    If dblSumMoney < 0 Then
    '        MsgBox "�����ܶ�[" & Format(dblSumMoney, "####0.00;-####0.00;0;0") & "]����Ϊ����,���顣", vbInformation, gstrSysName
    '        mshCash.SetFocus: Exit Sub
    '    End If

    
    If k = 0 Or mstrLast = "" Or mrsDetail.State = 0 Then
        MsgBox "û����ȡ��Ч�Ľɿ��", vbInformation, gstrSysName
        dtpDate.SetFocus: Exit Sub
    End If
    If mrsDetail.RecordCount = 0 Then
        MsgBox "û����ȡ��Ч�Ľɿ��", vbInformation, gstrSysName
        dtpDate.SetFocus: Exit Sub
    End If
    
    '����SQL���
    arrSQL = Array()
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, CONDFormat) & "','YYYY-MM-DD HH24:MI:SS')"
    lng����ID = zlDatabase.GetNextId("��Ա�ɿ��¼")
    With mshCash
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 And .TextMatrix(i, 0) <> "[��Ԥ����]" Then
                If lngID = 0 Then
                    lngID = lng����ID
                Else
                    lngID = zlDatabase.GetNextId("��Ա�ɿ��¼")
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_��Ա�ɿ��¼_Insert(" & lngID & "," & lng����ID & "," & strDate & "," & _
                    "'" & mstr�ɿ��� & "','" & UserInfo.���� & "','" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 1)) & "," & _
                    "'" & .TextMatrix(i, 2) & "','" & txtItem(3).Text & "'," & mstrLast & "," & cbo�ɿ��.ItemData(cbo�ɿ��.ListIndex) & ")"
            End If
        Next
    End With
    If mrsDetail.RecordCount <> 0 Then mrsDetail.MoveFirst
    For i = 1 To mrsDetail.RecordCount
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_��Ա�ɿ����_Insert(" & lng����ID & "," & mrsDetail!���� & "," & mrsDetail!��¼ID & ")"
        mrsDetail.MoveNext
    Next
    
    '����ɿ��¼
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '��ӡƱ��
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me, "����ID=" & lng����ID, 2)
    
    Screen.MousePointer = 0
    mstrLast = "" '���Ϊ���Թر�
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdPrint_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me
End Sub

Private Sub dtpDate_Change()
    If mEditType = PM_���սɿ� Then
        Call LoadCashType(mstr�ɿ���, dtpDate.Value)
        Call LoadRec(mstr�ɿ���, dtpDate.Value)
    End If
End Sub

Private Sub LoadCashType(ByVal strOperator As String, ByVal datThis As Date)
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
 
    strSQL = "Select Distinct ����, Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨',6, '���ѿ�') ����˵��" & vbNewLine & _
            "From �շ�����¼" & vbNewLine & _
            "Where �տ�Ա = [1] And ���� = [2]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOperator, datThis)
    With cboType
        .Clear
        .AddItem "ȫ�����"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0  '����click�¼�
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!����˵��
            .ItemData(.NewIndex) = rsTmp!����
            rsTmp.MoveNext
        Next
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRec(ByVal strOperator As String, ByVal datThis As Date)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    With mshRec
        .Redraw = False
        Call zlControl.MshSetFormat(mshRec, CONRecHead, Me.Caption, , , True)
        
        strSQL = "Select Decode(����, 1, 'Ԥ����', 2, '����', 3, '�շ�', 4, '�Һ�', 5, '���￨',6,'���ѿ�') ����," & vbNewLine & _
                "       Row_Number() Over(Partition By ���� Order By ��ʼʱ��) ����," & vbNewLine & _
                "       ��ʼʱ��, ��ֹʱ��" & vbNewLine & _
                "From �շ�����¼" & vbNewLine & _
                "Where �տ�Ա = [1] And ���� = [2]" & vbNewLine & _
                "Order By ����, ����"

        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOperator, datThis)
        If rsTmp.RecordCount > 0 Then
            .Rows = .FixedRows + rsTmp.RecordCount
            .MergeCol(0) = True
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = rsTmp!����
                .TextMatrix(i, 1) = rsTmp!����
                .TextMatrix(i, 2) = Format(rsTmp!��ʼʱ��, CONDFormat)
                .TextMatrix(i, 3) = Format(rsTmp!��ֹʱ��, CONDFormat)
                
                rsTmp.MoveNext
            Next
        Else
            .Rows = .FixedRows + 1
        End If
        
        .Redraw = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCashTimes(ByVal strOperator As String, ByVal datThis As Date, ByVal bytFlag As Byte)
    Dim strSQL As String, i As Long
    
    With cboTimes
        .Clear
        .AddItem "ȫ������"
        .ListIndex = .NewIndex  '����click�¼�
        
        If bytFlag <> 0 Then
            strSQL = "Select Rownum ����, ��ʼʱ��,��ֹʱ�� From " & _
            "(Select ��ʼʱ��,��ֹʱ�� From �շ�����¼ Where �տ�Ա=[1] And ���� = [2] And ���� = [3] Order By ��ʼʱ��)"
        
            On Error GoTo errH
            Set mrsTimes = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOperator, datThis, bytFlag)
            For i = 1 To mrsTimes.RecordCount
                .AddItem "��" & i & "�νɿ�"
                mrsTimes.MoveNext
            Next
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    ElseIf KeyCode = vbKeyF5 Then
        Call cmdRefresh_Click
    ElseIf KeyCode = 13 Then
        If Not ActiveControl Is mshCash Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    
    mblnOK = False
    
    txtItem(0).Text = mstr�ɿ���
    Set rsTmp = GetPersonnelDept(mlng�ɿ���ID)
    Call zlControl.CboAddData(cbo�ɿ��, rsTmp, True)
    If cbo�ɿ��.ListCount > 0 Then cbo�ɿ��.ListIndex = 0
    txtItem(4).Text = UserInfo.����
    
    Call InitFace
    Call zlSetDefaultDate
    Call LoadGroups
   
End Sub
Private Sub LoadGroups()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����Ա������Ϣ
    '����:���˺�
    '����:2010-11-29 14:30:43
    '����:33633
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    '����:33633
    gstrSQL = "" & _
    "   Select A.������,A.ID From ����ɿ���� A ,�ɿ��Ա��� B " & _
    "   Where A.ID=B.��ID And B.��ԱID=[1] and A.ɾ������>=sysdate"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng�ɿ���ID)
    txtGroups.Text = ""
    mlng��ID = -1
    txtGroups.Visible = True: lblGroups.Visible = True
    If Not rsTemp.EOF Then
        txtGroups.Text = Nvl(rsTemp!������)
        mlng��ID = Val(Nvl(rsTemp!ID))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mstrLast <> "" Then
        If MsgBox("ȷʵҪ�����ɿ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    mstr�ɿ��� = ""
    mstrLast = ""
    Set mrsDetail = Nothing
End Sub

Private Sub mshCash_EnterCell(Row As Long, Col As Long)
    If mshCash.TextMatrix(Row, 0) = "[��Ԥ����]" Then
        mshCash.ColData(2) = 0
    Else
        mshCash.ColData(2) = 4
    End If
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtItem(Index))
End Sub

Private Sub txtItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txtItem(Index).Locked Then
        glngTXTProc = GetWindowLong(txtItem(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtItem(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txtItem(Index).Locked Then
        Call SetWindowLong(txtItem(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
