VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDiffPriceAdjustCard 
   Caption         =   "����۵�����"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmDiffPriceAdjustCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   9
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   255
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   10
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9975
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   1425
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   2
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
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   4
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   960
         TabIndex        =   26
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "������ϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   24
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   1920
         TabIndex        =   23
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "����ۺϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   20
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   18
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   17
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   16
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "����۵�����"
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
         TabIndex        =   15
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   14
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   13
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   12
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   11
         Top             =   4500
         Width           =   720
      End
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
            Picture         =   "frmDiffPriceAdjustCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1000
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
            Picture         =   "frmDiffPriceAdjustCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiffPriceAdjustCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDiffPriceAdjustCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDiffPriceAdjustCard.frx":3080
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Caption         =   "����"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmDiffPriceAdjustCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbln�޸������� As Boolean           '�����޸�������
Private mbln��������    As Boolean          '����ʱ���ݺ��ۼ�1
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mblnFirst As Boolean                '��һ����ʾ
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭

Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Dim mstrPrivs As String                     'Ȩ��

Private mint����� As Integer             '��ʾ�������ϳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
'���˺�:2007/06/10:����10813
Private mstrTime_Start As String            '���뵥�ݱ༭�ĵ���ʱ�� ,��Ҫ�ж��Ƿ񵥾ݱ����˸��Ĺ�,����༭��,���ܽ������
Private mstrTime_End As String
Private Const mlngModule = 1715
Private Const mstrCaption As String = "����۵�����"
Private mstr�ظ����� As String '��¼�ظ�������

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntCol���� As Integer = 2
Private Const mconIntCol��� As Integer = 3
Private Const mconIntCol���� As Integer = 4
Private Const mconIntColʵ������ As Integer = 5
Private Const mconIntCol����ϵ�� As Integer = 6
Private Const mconIntCol���� As Integer = 7
Private Const mconIntCol��λ As Integer = 8
Private Const mconIntCol���� As Integer = 9
Private Const mconIntColЧ�� As Integer = 10
Private Const mconIntColһ���Բ��� As Integer = 11
Private Const mconIntCol���Ч�� As Integer = 12
Private Const mconIntCol�������    As Integer = 13
Private Const mconIntCol���ʧЧ�� As Integer = 14
Private Const mconIntCol����� As Integer = 15
Private Const mconintCol����� As Integer = 16
Private Const mconintcol�ɱ��� As Integer = 17
Private Const mconintcol�³ɱ��� As Integer = 18
Private Const mconintCol������ As Integer = 19
Private Const mconIntColS  As Integer = 20              '������
'=========================================================================================


'�������������
Private Function GetDepend() As Boolean
    Dim rsTemp As New Recordset

    GetDepend = False
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   SELECT B.Id " & _
        "   FROM ҩƷ�������� A, ҩƷ������ B " & _
        "   Where A.���id = B.ID AND A.���� = 33 "
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "����۵���"
    
    If rsTemp.EOF Then
        MsgBox "û�������������Ͽ���۵���������������������������ã�", vbInformation + vbOKOnly, gstrSysName
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
        Optional int��¼״̬ As Integer = 1, Optional ByVal strPrivs As String, Optional blnSuccess As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʾ��༭����,��Ψһ���
    '--�����:
    '--������:
    '--��  ��:blnSuccess
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String

    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mblnFirst = True
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub

    mbln�޸������� = IIf(Val(zlDatabase.GetPara("�޸Ĳɹ��޼�", glngSys, mlngModule, "0")) = 1, 1, 0) = 1
   
    
    Call GetRegInFor(g˽��ģ��, "����۵�������", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
    
    
    If mint�༭״̬ = 1 Then
'        If mbln�������� Then
'            mstr���ݺ� = NextNo(71)
'        End If
        mblnEdit = True
        txtNO.Locked = True
        txtNO.TabStop = True

        txtNO = mstr���ݺ�
        txtNO.Tag = mstr���ݺ�
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRownew mshBill, mconIntCol����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram stbThis, gSystem_Para.int���뷽ʽ
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    If mint�༭״̬ = 1 Then
        mshBill.ClearBill
        
        Dim str����ID As String, lng�ⷿID As Long, int��۲����� As Integer
        If frmDiffPriceAdjustCondition.GetCondition(mfrmMain, str����ID, lng�ⷿID, int��۲�����) = True Then
        
            Screen.MousePointer = 11
            SearchData str����ID, lng�ⷿID, int��۲�����
            Screen.MousePointer = 0
        Else
            Unload Me
            Exit Sub
        End If
        
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    Else
        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '����
            Case 2
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
            Case 3
                '�޸ĵĵ����ѱ����
                MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
        End Select
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mconIntCol����, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim strReg As String
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
        
        '���˺�:2007/06/10:����10813
        mstrTime_End = GetBillInfo(18, txtNO.Tag)
        If mstrTime_End = "" Then
            MsgBox "ע��:" & vbCrLf & "  �õ����Ѿ�����������Աɾ��,���ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If mstrTime_End <> mstrTime_Start Then
            If MsgBox("ע��:" & vbCrLf & "  �õ����Ѿ�����������Ա�༭�����ܼ���!" & vbCrLf & "  �Ƿ�����ˢ�µ���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                Call initCard
            End If
            Exit Sub
        End If
                
        If SaveCheck = True Then
            strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
            
        strReg = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
        If Val(strReg) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                printbill
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
        
'    If mbln�������� Then
'        mstr���ݺ� = NextNo(71)
'        txtNO = mstr���ݺ�
'    End If
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    txtժҪ.Text = ""
    mblnChange = False
    If txtNO.Tag <> "" Then Me.stbThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNO.Tag
    
End Sub

Private Sub Form_Load()
   Dim strReg As String

    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
         
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO.Text
    initCard
    RestoreWinState Me, App.ProductName, mstrCaption
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    
    On Error GoTo ErrHandle
    '�ⷿ
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    
    strCompare = Mid(strOrder, 1, 1)
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            txtStock = .List(.ListIndex)
            txtStock.Tag = .ItemData(.ListIndex)
            
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "" & _
                "   Select distinct b.id,b.���� " & _
                "   From ҩƷ�շ���¼ a,���ű� b  " & _
                "   Where a.�ⷿid=b.id and A.���� =18 and  a.no=[1]"
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                
                
                If rsTemp.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsTemp!����
                txtStock.Tag = rsTemp!Id
                
                rsTemp.Close
            End If
            
            Select Case mintUnit
                Case 0
                    strUnitQuantity = "c.���㵥λ AS ��λ, A.��д���� as ��������,'1' as ����ϵ��,"
                Case Else
                    strUnitQuantity = "b.��װ��λ AS ��λ,(A.��д���� / B.����ϵ��) AS ��������,B.����ϵ�� as ����ϵ��,"
            End Select
            
            gstrSQL = "" & _
                "   Select * " & _
                "   From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ, c.���," & _
                "                   A.����, A.����,a.Ч��,a.�������,a.���Ч�� as ���ʧЧ��,b.һ���Բ���,b.���Ч��,a.����," & _
                "                   zlSpellCode(c.����) ����," & strUnitQuantity & _
                "                   A.�ɱ��� as �����,nvl(a.���ۼ�,0) as �����,A.��� as ������,(nvl(a.���ۼ�,0)-nvl(a.�ɱ���,0))/a.��д���� as �ɱ���,a.���� as �³ɱ���, " & _
                "                   a.ժҪ,������,��������,�����,�������,a.�ⷿid " & _
                "           FROM ҩƷ�շ���¼ A, ��������  b,�շ���ĿĿ¼ c" & _
                "           Where A.ҩƷid = B.����id and a.ҩƷid=c.id " & _
                "                   AND A.��¼״̬ =[2]" & _
                "                   AND A.���� =18 AND A.No = [1]" & _
                "           ) " & _
                " ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�, mint��¼״̬)
            
            If rsTemp.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            '���˺�:2007/06/10:����10813
            mstrTime_Start = GetBillInfo(18, mstr���ݺ�)
            
            Txt������ = rsTemp!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û���
            End If
            Txt�������� = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
            
            Txt����� = IIf(IsNull(rsTemp!�����), "", rsTemp!�����)
            Txt������� = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsTemp!ժҪ), "", rsTemp!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            With mshBill
                Do While Not rsTemp.EOF
                    
                    intRow = rsTemp.AbsolutePosition
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsTemp.Fields(0)
                    .TextMatrix(intRow, mconIntCol����) = rsTemp!������Ϣ
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mconIntCol��λ) = rsTemp!��λ
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
                   
                    .TextMatrix(intRow, mconIntColһ���Բ���) = zlStr.Nvl(rsTemp!һ���Բ���)
                    .TextMatrix(intRow, mconIntCol���Ч��) = zlStr.Nvl(rsTemp!���Ч��)
                    .TextMatrix(intRow, mconIntCol�������) = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mconIntCol���ʧЧ��) = IIf(IsNull(rsTemp!���ʧЧ��), "", Format(rsTemp!���ʧЧ��, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mconIntCol�����) = Format(rsTemp!�����, mFMT.FM_���)
                    .TextMatrix(intRow, mconintCol�����) = Format(IIf(IsNull(rsTemp!�����), 0, rsTemp!�����), mFMT.FM_���)
                    .TextMatrix(intRow, mconintCol������) = Format(rsTemp!������, mFMT.FM_���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
                    .TextMatrix(intRow, mconIntColʵ������) = Format(IIf(IsNull(rsTemp!��������), "0", rsTemp!��������), mFMT.FM_����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsTemp!����ϵ��
                    .TextMatrix(intRow, mconintcol�ɱ���) = Format(rsTemp!�ɱ��� * rsTemp!����ϵ��, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mconintcol�³ɱ���) = Format(rsTemp!�³ɱ��� * rsTemp!����ϵ��, mFMT.FM_�ɱ���)
                    
                    rsTemp.MoveNext
                Loop
            End With
            rsTemp.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    mint����� = Get������(Val(txtStock.Tag))
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ʼ���༭�ؼ�
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntCol����) = "���������"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = "ʧЧ��"
        
        .TextMatrix(0, mconIntColһ���Բ���) = "һ���Բ���"
        .TextMatrix(0, mconIntCol���Ч��) = "���Ч��"
        .TextMatrix(0, mconIntCol���ʧЧ��) = "���ʧЧ��"
        .TextMatrix(0, mconIntCol�������) = "�������"
         
        .TextMatrix(0, mconintCol�����) = "�����"
        .TextMatrix(0, mconIntCol�����) = "�����"
        .TextMatrix(0, mconintCol������) = "������"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColʵ������) = "�������"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconintcol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconintcol�³ɱ���) = "�³ɱ���"

        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntColʵ������) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        
        .ColWidth(mconIntCol����) = 2500
        .ColWidth(mconIntCol���) = 1000
        .ColWidth(mconIntCol����) = 1000
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol����) = 1000
        .ColWidth(mconIntColЧ��) = 1000
        
        .ColWidth(mconIntColһ���Բ���) = 0
        .ColWidth(mconIntCol���Ч��) = 0
        .ColWidth(mconIntCol���ʧЧ��) = 1000
        .ColWidth(mconIntCol�������) = 1000
        
        .ColWidth(mconIntCol�����) = 1200
        .ColWidth(mconintCol�����) = 1200
        .ColWidth(mconintcol�ɱ���) = 1200
        .ColWidth(mconintcol�³ɱ���) = 1200
        .ColWidth(mconintCol������) = 1200
        
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColЧ��) = 5
        
        .ColData(mconIntColһ���Բ���) = 5
        .ColData(mconIntCol���Ч��) = 5
        .ColData(mconIntCol���ʧЧ��) = 5
        .ColData(mconIntCol�������) = 2
          
        .ColData(mconintCol�����) = 5
        .ColData(mconIntCol�����) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColʵ������) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconintcol�ɱ���) = 5
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
            
            
            .ColData(mconIntCol����) = 1
            .ColData(mconintCol������) = 4
            .ColData(mconintcol�³ɱ���) = 4
            
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
            
            txtժҪ.Enabled = False
            
            .ColData(mconintCol������) = 5
            .ColData(mconintcol�³ɱ���) = 5
        End If
        
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        
        .ColAlignment(mconIntColһ���Բ���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���Ч��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���ʧЧ��) = flexAlignCenterCenter
        .ColAlignment(mconIntCol�������) = flexAlignCenterCenter
        
        .ColAlignment(mconIntCol�����) = flexAlignRightCenter
        .ColAlignment(mconIntColʵ������) = flexAlignRightCenter
        .ColAlignment(mconintCol�����) = flexAlignRightCenter
        
        .ColAlignment(mconintCol������) = flexAlignRightCenter
        .ColAlignment(mconintcol�³ɱ���) = flexAlignRightCenter
        .ColAlignment(mconintcol�ɱ���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntCol����
        .LocateCol = mconIntCol����
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mconIntCol����) = 0
    End With
    txtժҪ.MaxLength = sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
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
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
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
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    
End Sub

Private Function SaveCheck() As Boolean
    Dim strNo As String
    Dim str����� As String
    
    mblnSave = False
    SaveCheck = False
    str����� = UserInfo.�û���
    strNo = txtNO.Tag
    On Error GoTo ErrHandle
    
    gstrSQL = "zl_���Ͽ���۵���_Verify('" & strNo & "','" & str����� & "')"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntCol����) = 0 Then
        Exit Sub
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ�������������ϣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, txtStock.Tag, , txtStock.Tag, False, , , , , , , , , , , , mstrPrivs & ";�鿴�ɱ���;", , False)
    If RecReturn.RecordCount > 0 Then
        With mshBill
            Dim strUnit As String
            Dim intUnit As Integer
            
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                    IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                    IIf(IsNull(RecReturn!����), "0", RecReturn!����), _
                     IIf(IsNull(RecReturn!ʵ������), "0", RecReturn!ʵ������), _
                    IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��)) Then
                    
                    If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                    .Row = .Row + 1
                End If
                
                .Col = mconintcol�³ɱ���
                RecReturn.MoveNext
            Next
            
            mshBill.Row = int�����
            
            If mstr�ظ����� <> "" Then
                MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                mstr�ظ����� = ""
            End If
            
'            If RecReturn.RecordCount = 1 Then
'
'                SetColValue .Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                    IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                    IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
'                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                    IIf(IsNull(RecReturn!����), "0", RecReturn!����), _
'                     IIf(IsNull(RecReturn!ʵ������), "0", RecReturn!ʵ������), _
'                    IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��)
'                 .Col = mconintcol�³ɱ���
'
'            End If
        End With
        RecReturn.Close
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconintcol�³ɱ��� Or .Col = mconintCol������ Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            
            Select Case .Col
                Case mconintcol�³ɱ���
                   intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.�ɱ���С��, g_С��λ��.obj_ɢװС��.�ɱ���С��)
                Case mconintCol������
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.���С��, g_С��λ��.obj_ɢװС��.���С��)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        
        Select Case .Col
            Case mconIntCol����
                .TxtCheck = False
                .MaxLength = 40
                'ֻ�������в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                
            Case mconintCol������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890-"
          Case mconIntCol�������
                .TxtCheck = True
                .Value = Format(sys.Currentdate, "yyyy-mm-dd")
                .TextMask = "1234567890-"
                .MaxLength = 10
            Case mconintcol�³ɱ���
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim dblʵ������ As Double
    Dim dblMoney As Double
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int����� As Integer
    
    int����� = mshBill.Row
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            
            Case mconIntCol����
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, txtStock.Tag, , txtStock.Tag, strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, False, , , , , , , , , , , mstrPrivs, , False)
                    
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                                IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                IIf(IsNull(RecReturn!����), "0", RecReturn!����), _
                                IIf(IsNull(RecReturn!ʵ������), "0", RecReturn!ʵ������), _
                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��)) Then
                                
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        RecReturn.MoveNext
                    Next
                    
                    mshBill.Row = int�����
                    
                    If mstr�ظ����� <> "" Then
                        MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                        mstr�ظ����� = ""
                    End If
                    
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                                IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                                IIf(IsNull(RecReturn!����), "0", RecReturn!����), _
'                                IIf(IsNull(RecReturn!ʵ������), "0", RecReturn!ʵ������), _
'                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call ��ʾ�����
                End If
           
          Case mconIntCol�������
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "������ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol���Ч��)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("�����������Ѿ��������ʧЧ��(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol���Ч��)), CDate(strKey)), "yyyy-mm-dd") & "),�Ƿ�Ҫ�������!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    '����ʧЧ��
                    .TextMatrix(.Row, mconIntCol���ʧЧ��) = Format(DateAdd("m", Val(.TextMatrix(.Row, mconIntCol���Ч��)), CDate(strKey)), "yyyy-mm-dd")
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol�������) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
            Case mconintcol�³ɱ���
               If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�ɱ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0.001 Then
                        MsgBox "�ɱ��۱������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "�ɱ��۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(strKey, 4)
                    .TextMatrix(.Row, .Col) = .Text
                End If
      
                If strKey <> "" Then
                    strKey = Format(strKey, mFMT.FM_�ɱ���)
                    .Text = strKey
                    .TextMatrix(.Row, mconintcol�³ɱ���) = .Text
                    
                    '�����۵�����(�����������������*�ɱ���-�����)
                    dblʵ������ = Val(.TextMatrix(.Row, mconIntColʵ������))
                    dblMoney = Val(.TextMatrix(.Row, mconIntCol�����)) - dblʵ������ * Val(strKey) - Val(.TextMatrix(.Row, mconintCol�����))
                       .TextMatrix(.Row, mconintCol������) = Format(dblMoney, mFMT.FM_���)
                End If
            Case mconintCol������
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "������������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "�������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Abs(Val(strKey)) < 0.01 Then
                        MsgBox "������ı������0.01,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "���������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_���)
                    .Text = strKey
                    
                    '����ɱ���(�ɱ���=(�����-�����-������)/ʵ������)
                    dblʵ������ = Val(.TextMatrix(.Row, mconIntColʵ������))
                    If dblʵ������ <> 0 Then
                        dblMoney = Val(.TextMatrix(.Row, mconIntCol�����)) - Val(.TextMatrix(.Row, mconintCol�����)) - Val(strKey)
                        dblMoney = dblMoney / dblʵ������
                        .TextMatrix(.Row, mconintcol�³ɱ���) = Format(dblMoney, mFMT.FM_�ɱ���)
                    End If
                
                End If
                Call ��ʾ�ϼƽ��
        End Select
    End With
End Sub

'�Ӳ���������ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
    ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal str���� As String, ByVal strЧ�� As String, _
    ByVal num����� As Double, ByVal lng���� As Long, ByVal num�������� As Double, _
    ByVal num����ϵ�� As Double, ByVal num����� As Double) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    SetColValue = False
    gstrSQL = "Select һ���Բ���,���Ч�� from �������� where ����id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
    
    
    With mshBill
        
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID And Val(.TextMatrix(lngRow, mconIntCol����)) = lng���� Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & str���� & "��"  '����¼�����ظ�������
                    'Call MsgBox("�������ϡ�" & str���� & "(" & lng���� & ")���Ѿ����ڣ�������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mconIntCol�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        
        .TextMatrix(intRow, mconIntColһ���Բ���) = zlStr.Nvl(rsTemp!һ���Բ���)
        .TextMatrix(intRow, mconIntCol���Ч��) = zlStr.Nvl(rsTemp!���Ч��)
        
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntColʵ������) = Format(num�������� / IIf(num����ϵ�� = 0, 1, num����ϵ��), mFMT.FM_����)
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol����) = lng����
        .TextMatrix(intRow, mconIntCol�����) = Format(num�����, mFMT.FM_���)
        .TextMatrix(intRow, mconintCol�����) = Format(num�����, mFMT.FM_���)
        .TextMatrix(intRow, mconintcol�ɱ���) = Format(Get�ɱ���(lng����ID, txtStock.Tag, lng����) * num����ϵ��, mFMT.FM_�ɱ���)
        
    End With
    Call ��ʾ�����
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
        
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "���ݺų���,���������" & CInt(txtNO.MaxLength / 2) & "�����֣���ò�Ҫ���֣���" & txtNO.MaxLength & "���ַ�!"
            txtNO.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol����)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconintCol������))) = "" Then
                        MsgBox "��" & intLop & "���������ϵĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol������)) > 9999999999999# Then
                        MsgBox "��" & intLop & "���������ϵĵ�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol������
                        Exit Function
                    End If
                    
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function


Private Function SaveCard() As Boolean
    Dim lng������ID As Long
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿID As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng����ID As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim dbl�������� As Double
    Dim dbl����� As Double
    Dim dbl����� As Double
    Dim dbl������ As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim dat�������� As String
    Dim rs������ As New Recordset
    Dim str������� As String
    Dim str���Ч�� As String
    Dim dbl�³ɱ��� As Double
      
    Dim intRow As Integer
    
    On Error GoTo ErrHandle
    SaveCard = False
    
    '����������������ID����Ҫ�����в��϶�Ҫ����
    
    gstrSQL = "SELECT B.Id " _
        & " FROM ҩƷ�������� A, ҩƷ������ B " _
        & "Where A.���id = B.ID " _
      & "AND A.���� = 33 "
    
    zlDatabase.OpenRecordset rs������, gstrSQL, mstrCaption
    
    If rs������.EOF Then
        MsgBox "û�������������Ͽ���۵���������������������������ã�", vbInformation + vbOKOnly, gstrSysName
        rs������.Close
        Exit Function
    End If
    lng������ID = rs������.Fields(0)
    rs������.Close
    
    With mshBill
        chrNo = Trim(txtNO)
        lng�ⷿID = txtStock.Tag
        
        If mint�༭״̬ = 1 Then   'mbln�������� Or
            If chrNo <> "" Then
                If CheckNOExists(71, chrNo) Then Exit Function
            End If
            If chrNo = "" Then chrNo = sys.GetNextNo(71, lng�ⷿID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        dat�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        gcnOracle.BeginTrans
        If mint�༭״̬ = 2 Then        '�޸�
            
            gstrSQL = "zl_���Ͽ���۵���_Delete('" & mstr���ݺ� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
        End If
            
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = Val(.TextMatrix(intRow, 0))
                str���� = .TextMatrix(intRow, mconIntCol����)
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng����ID = Val(.TextMatrix(intRow, mconIntCol����))
                strЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                str������� = IIf(.TextMatrix(intRow, mconIntCol�������) = "", "", .TextMatrix(intRow, mconIntCol�������))
                str���Ч�� = IIf(.TextMatrix(intRow, mconIntCol���ʧЧ��) = "", "", .TextMatrix(intRow, mconIntCol���ʧЧ��))
                
                dbl�������� = Round(Val(.TextMatrix(intRow, mconIntColʵ������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dbl����� = Round(Val(.TextMatrix(intRow, mconIntCol�����)), g_С��λ��.obj_���С��.���С��)
                dbl����� = Round(Val(.TextMatrix(intRow, mconintCol�����)), g_С��λ��.obj_���С��.���С��)
                dbl������ = Round(Val(.TextMatrix(intRow, mconintCol������)), g_С��λ��.obj_���С��.���С��)
                dbl�³ɱ��� = Round(Val(.TextMatrix(intRow, mconintcol�³ɱ���)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), g_С��λ��.obj_���С��.�ɱ���С��)
                lng��� = intRow
                
                'zl_���Ͽ���۵���_INSERT( /*������ID_IN*/, /*NO_IN*/, /*���_IN*/,
                    '/*�ⷿID_IN*/, /*����ID_IN*/, /*����_IN*/, /*��������_IN*/,
                    '/*�����_IN*/, /*������_IN*/, /*������_IN*/, /*��������_IN*/,
                    '/*����_IN*/, /*����_IN*/, /*Ч��_IN*/*�������_IN*/,/*���Ч��_IN*//, /*ժҪ_IN*/ );
                    
                gstrSQL = "zl_���Ͽ���۵���_INSERT(" & lng������ID & ",'" & chrNo & "'," & lng��� & "," & _
                     lng�ⷿID & "," & lng����ID & "," & lng����ID & "," & dbl�������� & "," & _
                     dbl����� & "," & dbl����� & "," & dbl������ & ",'" & str������ & "',to_date('" & dat�������� & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                     str���� & "','" & str���� & "'," & _
                     IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & "," & _
                    IIf(str������� = "", "Null", "to_date('" & Format(str�������, "yyyy-MM-dd") & "','yyyy-mm-dd')") & "," & _
                    IIf(str���Ч�� = "", "Null", "to_date('" & Format(str���Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ",'" & _
                     strժҪ & "'," & dbl�³ɱ��� & ")"
                
                Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
            End If
        Next
        gcnOracle.CommitTrans
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim dbl����� As Double
    Dim dbl������ As Double
    Dim dbl����� As Double
    
    Dim intLop As Integer
    
    dbl����� = 0
    dbl������ = 0
    
    With mshBill
        For intLop = 1 To .Rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl����� = dbl����� + Val(.TextMatrix(intLop, mconintCol�����))
                dbl����� = dbl����� + Val(.TextMatrix(intLop, mconIntCol�����))
                dbl������ = dbl������ + Val(.TextMatrix(intLop, mconintCol������))
            End If
        Next
    End With
    
    
    lblPurchasePrice.Caption = "�����ϼƣ�" & Format(dbl�����, mFMT.FM_���)
    lblSalePrice.Caption = "����ۺϼƣ�" & Format(dbl�����, mFMT.FM_���)
    lblDifference.Caption = "������ϼƣ�" & Format(dbl������, mFMT.FM_���)
    
End Sub

Private Sub ��ʾ�����()
    
    If mint�༭״̬ = 4 Then Exit Sub
    With mshBill
        If .TextMatrix(.Row, mconIntCol����) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        stbThis.Panels(2).Text = "���������ϵ�ǰ�����Ϊ[" & .TextMatrix(.Row, mconIntColʵ������) & "]" & .TextMatrix(.Row, mconIntCol��λ)
    End With
End Sub

Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    ImeLanguage True
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    ImeLanguage False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'��ӡ����
Private Sub printbill()
    Dim strUnit As String
    Dim int��λϵ�� As Integer
    Dim strNo As String
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1715", mint��¼״̬, mintUnit, 1715, "�������ϲ�۵�����", strNo
    
End Sub


Private Sub SearchData(ByVal str����ID, ByVal lng�ⷿID As Long, _
    ByVal intRate As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���������������ȡ�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------

    
    Dim rsTemp As New Recordset  '�������Ͽ���¼��
    
    Dim strPhysic As String, i As Long
    Dim sngLevel As Single
    Dim intRecordCount As Integer
    
    Dim strUnit As String
    Dim strUnitQuantity As String
    
    On Error GoTo ErrHandle:
    
    
    '���ý�����ʾ����
    
    stbThis.Panels(2).Text = "���ڶ�" & txtStock & "���������Ͻ����Զ���ۼ���"
        
    
    '�����������ϲ�ѯ����(��������)
    strPhysic = " And (c.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or c.����ʱ�� is NULL)"
    
    If str����ID <> "" Then
            strPhysic = strPhysic & " And d.����ID IN(" & str����ID & ")"
    End If
    
    DoEvents
    Select Case mintUnit
        Case 0
            strUnitQuantity = "c.���㵥λ AS ��λ, nvl(b.��������,0) AS ��������, '1' as ����ϵ��,decode(nvl(b.ƽ���ɱ���,0),0,a.�ɱ���,b.ƽ���ɱ���) �ɱ���,"
                
        Case Else
            strUnitQuantity = "a.��װ��λ AS ��λ,(nvl(b.��������,0)/a.����ϵ��) AS ��������,a.����ϵ�� as ����ϵ��,decode(nvl(b.ƽ���ɱ���,0),0,a.�ɱ���,b.ƽ���ɱ���) �ɱ���,"
    End Select
    
    gstrSQL = "" & _
        "   SELECT distinct  b.ҩƷid ����id,c.����,c.���� AS ��Ʒ����, " & _
        "           c.���, decode(b.�ϴβ���,NULL,c.����,b.�ϴβ���) AS ����,b.����,b.�ϴ����� as ����, b.Ч��,b.���Ч�� as ���ʧЧ��,a.���Ч��,a.һ���Բ���," & _
        "           add_months(b.���Ч��,-a.���Ч��) as �������," & _
        "           B.ʵ�ʽ��, B.ʵ�ʲ��, " & strUnitQuantity & _
        "           DECODE (SIGN (B.ʵ�ʲ��/B.ʵ�ʽ��*100-(A.ָ�������+" & intRate & ")),1,-(ʵ�ʲ��-B.ʵ�ʽ��*A.ָ�������/100)," & _
        "           DECODE (SIGN(B.ʵ�ʲ��/B.ʵ�ʽ��*100-(A.ָ�������-" & intRate & ")),-1,B.ʵ�ʽ��*A.ָ�������/100-ʵ�ʲ��)) AS ��۵����� " & _
        "   FROM �������� A,�շ���ĿĿ¼ c,������ĿĿ¼ d, (Select �ⷿid, ҩƷid, ����, Ч��, ����,ʵ������ ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, ��׼�ĺ�,ƽ���ɱ��� From ҩƷ��� Where ����=1 and  Nvl(ʵ�ʽ��,0)<>0) B " & _
        "   Where A.����id=c.id and a.����id=d.id  and A.����ID = B.ҩƷID and B.����(+)=1 AND B.�ⷿid =[1]" & _
        "           AND ((B.����>0 AND B.ʵ������>0)  OR NVL(B.����,0)=0) " & _
        "           AND ( NVL(B.ʵ�ʽ��,0)<>0 " & _
        "           AND (B.ʵ�ʲ��/Nvl(B.ʵ�ʽ��,1)*100>(A.ָ�������+" & intRate & ") OR B.ʵ�ʲ��/Nvl(B.ʵ�ʽ��,1)*100<A.ָ�������-" & intRate & _
        "               )) " & _
                strPhysic & _
        "   Order by c.����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ڼ����������Ͽ������", lng�ⷿID)
    
    
    intRecordCount = rsTemp.RecordCount
    
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    If intRecordCount = 0 Then
        MsgBox "δ����ȷ��ȡ�������Ͽ������,�����Ի��ֹ������������ϣ�", vbInformation, gstrSysName: Exit Sub
    End If
    
    DoEvents:
    mshBill.Redraw = False
    
    rsTemp.MoveFirst
    i = 1
    With mshBill
        Do While Not rsTemp.EOF
           If i > 1 Then .Rows = .Rows + 1
           .TextMatrix(i, 0) = rsTemp!����ID
           .TextMatrix(i, mconIntCol����) = "[" & rsTemp!���� & "]" & rsTemp!��Ʒ����
           .TextMatrix(i, mconIntCol���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
           .TextMatrix(i, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
           .TextMatrix(i, mconIntCol��λ) = IIf(IsNull(rsTemp!��λ), "", rsTemp!��λ)
           .TextMatrix(i, mconIntCol����) = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
           .TextMatrix(i, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
           .TextMatrix(i, mconIntColЧ��) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd"))
            
            .TextMatrix(i, mconIntColһ���Բ���) = zlStr.Nvl(rsTemp!һ���Բ���)
            .TextMatrix(i, mconIntCol���Ч��) = zlStr.Nvl(rsTemp!���Ч��)
            .TextMatrix(i, mconIntCol�������) = IIf(IsNull(rsTemp!�������), "", Format(rsTemp!�������, "yyyy-mm-dd"))
            .TextMatrix(i, mconIntCol���ʧЧ��) = IIf(IsNull(rsTemp!���ʧЧ��), "", Format(rsTemp!���ʧЧ��, "yyyy-mm-dd"))
            
           .TextMatrix(i, mconIntColʵ������) = rsTemp!��������
           .TextMatrix(i, mconIntCol�����) = Format(rsTemp!ʵ�ʽ��, mFMT.FM_���)
           .TextMatrix(i, mconintCol�����) = Format(rsTemp!ʵ�ʲ��, mFMT.FM_���)
           .TextMatrix(i, mconintCol������) = Format(rsTemp!��۵�����, mFMT.FM_���)
           .TextMatrix(i, mconintcol�³ɱ���) = ""
           .TextMatrix(i, mconIntCol����ϵ��) = rsTemp!����ϵ��
           .TextMatrix(i, mconintcol�ɱ���) = Format(rsTemp!�ɱ��� * rsTemp!����ϵ��, mFMT.FM_�ɱ���)
    
            Call ShowPercent(i / intRecordCount)
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Redraw = True
    End With
    rsTemp.Close
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    
    stbThis.Panels(2).Text = ""
    mshBill.Row = 1
    mshBill.Col = mconintCol������
    If Me.Visible = True Then
        mshBill.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mshBill.Redraw = True
    Call SaveErrLog
    
End Sub

Private Sub ShowPercent(sngPercent As Single)
'����:��״̬���ϸ��ݰٷֱ���ʾ��ǰ�������(��)
    Dim intAll As Integer
    intAll = stbThis.Panels(2).Width / TextWidth("��") - 4
    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "��")
End Sub


