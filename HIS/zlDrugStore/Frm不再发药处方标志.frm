VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm���ٷ�ҩ������־ 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "Frm���ٷ�ҩ������־.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5581.047
   ScaleMode       =   0  'User
   ScaleWidth      =   9000
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   350
      Left            =   7830
      TabIndex        =   15
      Top             =   720
      Width           =   825
   End
   Begin VB.Frame fraSelect 
      Height          =   1110
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   8655
      Begin VB.TextBox txtBillNo 
         Height          =   300
         Left            =   3870
         TabIndex        =   2
         Top             =   630
         Width           =   1590
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Left            =   1290
         TabIndex        =   1
         Top             =   630
         Width           =   1590
      End
      Begin VB.Frame fraSelectFlag 
         Caption         =   "�Ƿ��ѱ�ǲ��ٷ�ҩ"
         Height          =   825
         Left            =   5670
         TabIndex        =   7
         Top             =   180
         Width           =   1905
         Begin VB.OptionButton optUnFlag 
            Caption         =   "δ���"
            Height          =   195
            Left            =   450
            TabIndex        =   9
            Top             =   270
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton optFlag 
            Caption         =   "�ѱ��"
            Height          =   195
            Left            =   450
            TabIndex        =   8
            Top             =   540
            Width           =   1005
         End
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1290
         TabIndex        =   10
         Top             =   255
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   125370368
         CurrentDate     =   38455
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   3870
         TabIndex        =   17
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   529
         _Version        =   393216
         Format          =   125370368
         CurrentDate     =   38455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   3060
         TabIndex        =   16
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   315
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��  ��"
         Height          =   180
         Left            =   540
         TabIndex        =   12
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�� �� ��"
         Height          =   180
         Left            =   3060
         TabIndex        =   11
         Top             =   690
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   5
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   4
      Top             =   5205
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "���(&O)"
      Height          =   350
      Left            =   6120
      TabIndex        =   3
      Top             =   5205
      Width           =   1275
   End
   Begin VB.Frame fraGrid 
      Height          =   3840
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   8655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshBill 
         Height          =   3435
         Left            =   50
         TabIndex        =   14
         Top             =   180
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   6059
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         GridColor       =   -2147483631
         GridColorFixed  =   8421504
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1440
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "Frm���ٷ�ҩ������־"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''����

'�ؼ��򴰿ڱ���
Private Const M_STR_FRASELECT_CAPTION As String = "��ѯ����"                             '��ѯ�����

'ҩƷ����
Private Const M_STR_ҩƷ_FRM_CAPTION As String = "ֹͣ��ҩ��־"                          '���ڱ���
Private Const M_STR_ҩƷ_FRAGRID_CHECK_CAPTION As String = "�ѱ��ֹͣ��ҩ������Ϣ"      '������Ϣ�����-�������ʱ
Private Const M_STR_ҩƷ_FRAGRID_UNCHECK_CAPTION As String = "δ���ֹͣ��ҩ������Ϣ"    '������Ϣ�����-δ�����ʱ
Private Const M_STR_ҩƷ_FRASELECTFLAG_CAPTION As String = "�Ƿ��ѱ��ֹͣ��ҩ"          '��Ǻ�δ���ѡ������

'���Ĵ���
Private Const M_STR_����_FRM_CAPTION As String = "ֹͣ���ϱ�־"                          '���ڱ���
Private Const M_STR_����_FRAGRID_CHECK_CAPTION As String = "�ѱ��ֹͣ���ϴ�����Ϣ"      '������Ϣ�����-�������ʱ
Private Const M_STR_����_FRAGRID_UNCHECK_CAPTION As String = "δ���ֹͣ���ϴ�����Ϣ"    '������Ϣ�����-δ�����ʱ
Private Const M_STR_����_FRASELECTFLAG_CAPTION As String = "�Ƿ��ѱ��ֹͣ����"          '��Ǻ�δ���ѡ������


'ȷ�ϰ�ť����
Private Const M_STR_CMDOK_CHECK As String = "���(&F)"
Private Const M_STR_CMDOK_UNCHECK As String = "�ָ����(&U)"

'�������
Private Const M_STR_CHECK_NAME As String = "�ѱ��"
Private Const M_STR_UNCHECK_NAME As String = "δ���"

'�ѡ�δ��Ǽ�¼��ɫ
Const M_LNG_CHECKED_COLOR = &HC0C0C0
Const M_LNG_UNCHECKED_COLOR = &H8000000E

'�̶���Ԫ����ɫ
Const M_LNG_FIXEDCOLS_COLOR = &H8000000F

'ѡ���б���ɫ
Const M_LNG_SELECTEDCOLS_COLOR = &HFFC0C0

'Ĭ���б���ɫ
Const M_LNG_DEFAULTCOLS_COLOR = &H8000000E

'�ؼ��򴰿���ʾ
Private Const M_STR_PATI_INPUT_DESC As String = "�������룺*����ţ�+סԺ�ţ�-����ID�����Ҳ���"                       '�����������ʾ
Private Const M_STR_FLAG_DESC As String = "ѡ��δ���-��ѯδ����ǵ�δ��������ѡ���ѱ��-��ѯ������ǵ�δ������"        '��Ǻ�δ���ѡ�����ʾ
Private Const M_STR_GRID_DESC As String = "�ڵ�һ�д������δ��������Ҳ����ȡ���Ѵ򹴵���"                              '������¼�б���ʾ
Private Const M_STR_BILLNO_DESC As String = "���봦���ţ�֧������������������ȡ�ô�����Ϣ"                                '�������������ʾ

'����Ĭ�Ͽ�ȡ��߶�
Private Const M_STR_FRM_WIDTH As Long = 9000
Private Const M_STR_FRM_HEIGHT As Long = 6000

'�����ַ������ָ�ʽ
Private Const M_STR_DEFAULT_ORA_NUMERIC_FORMAT As String = "9999990.00000"               'Ĭ��Oracle����ʽ
Private Const M_STR_DEFAULT_ORA_DATE_FORMAT As String = "yyyy-mm-dd hh24:mi:ss"          'Ĭ��Oracleʱ���ʽ
Private Const M_STR_VB_DATE_FORMAT As String = "yyyy-mm-dd hh:mm:ss"                     'Ĭ��VBʱ���ʽ


''''����

'���ڻ�ؼ��������
Private mstrFrmCaption As String
Private mstrFraGridCheckCaption As String
Private mstrFraGridUnCheckCaption As String
Private mstrFraSelectFlagCaption As String

Private mstr�����շ��뷢ҩ���� As String
Private mstrסԺ�����뷢ҩ���� As String
Private mint�������� As Integer                  '1-������ҩ���� 2-���ŷ�ҩ����
Private mlng��ҩҩ��ID As Long
Private mstrҩ�� As String
Private mstrSystemNumericFormat As String            '����ʽ
Private mstrSystemAmountFormat As String             '������ʽ
Private mblnIsChecked As Boolean                      '�жϼ�¼ԭʼ״̬�Ƿ�Ϊ�������
Private mintBillType As Integer                      '������� 1-ҩƷ���� 2-���Ĵ���
Private mint�������� As Integer                       '��������סԺ��0-����סԺ��1-���2-סԺ
Private mbln������ As Boolean
Private mlngҩ��ID As Long
Private mintNumberDigit As Integer          '����С��λ��


Private mstrPrivs As String                              'Ȩ�޴�

Private mIntCheckStock As Integer               '����飺0-�����;1-���,��������;2-���,�����ֹ

Private mstrֹͣ As String
Private mstr�ָ� As String

Private mBillCol As BILLCOL

Private mstrDeptNode As String

''''ȫ�ֱ���
Public gstrParentName As String                     '�����������

''''����

'����������
Private Type BILLCOL
    BillCols As Integer

    Flag  As String
    FlagCol  As Integer
    FlagColWidth As Long
    FlagColAlig As Integer
       
    Tag  As String
    TagCol  As Integer
    TagColWidth As Long
    TagColAlig As Integer
        
    Id As String
    IdCol As Integer
    IdColWidth As Long
    IdColAlig As Integer
        
    Pati As String
    PatiCol As Integer
    PatiColWidth As Long
    PatiColAlig As Integer
    
    Bill As String
    BILLCOL As Integer
    BillColWidth As Long
    BillColAlig As Integer
    
    NO As String
    NoCol As Integer
    NoColWidth As Long
    NoColAlig As Integer
    
    Drug As String
    DrugCol As Integer
    DrugColWidth As Long
    DrugColAlig As Integer
    
    Spec As String
    SpecCol As Integer
    SpecColWidth As Long
    SpecColAlig As Integer
    
    Unit As String
    UnitCol As Integer
    UnitColWidth As Long
    UnitColAlig As Integer
    
    UnitPrice As String
    UnitPriceCol As Integer
    UnitPriceColWidth As Long
    UnitPriceColAlig As Integer
    UnitPriceFormat As String
    
    count As String
    CountCol As Integer
    CountColWidth As Long
    CountColAlig As Integer
    
    Amount As String
    AmountCol As Integer
    AmountColWidth As Long
    AmountColAlig As Integer
        
    Price As String
    PriceCol As Integer
    PriceColWidth As Long
    PriceColAlig As Integer
    PriceColFormat As String
    
    Group As String
    GroupCol As Integer
    GroupColWidth As Long
    GroupColAlig As Integer
    
    BillDate As String
    BillDateCol As Integer
    BillDateColWidth As Long
    BillDateColAlig As Integer
    
    Category As String
    CategoryCol As Integer
    CategoryColWidth As Integer
    CategoryColAlig As Integer
    
    ��¼���� As String
    ��¼����Col As Integer
    ��¼����ColWidth As Integer
    ��¼����ColAlig As Integer
    
    �����־ As String
    �����־Col As Integer
    �����־ColWidth As Integer
    �����־ColAlig As Integer
    
    ȱҩ As String
    ȱҩCol As Integer
    ȱҩColWidth As Integer
    ȱҩColAlig As Integer
End Type

Public Property Get In_�����() As Integer
    In_����� = mIntCheckStock
End Property

Public Property Let In_�����(ByVal vNewValue As Integer)
    mIntCheckStock = vNewValue
End Property

'--50313��zdt��:������Խ��ܷ��ϲ���id��ֵ
Public Property Get In_�ⷿid() As Long
    In_�ⷿid = mlng��ҩҩ��ID
End Property

Public Property Let In_�ⷿid(ByVal vNewValue As Long)
    mlng��ҩҩ��ID = vNewValue
End Property

Public Property Let In_��������(ByVal vNewValue As Integer)
    mint�������� = vNewValue
End Property

Public Property Get In_��������() As Integer
    In_�������� = mint��������
End Property

'����Ƿ�Ϊ����
Private Function CheckIsDate(dtInput As Date) As Boolean
    CheckIsDate = IsDate(dtInput)
End Function

'����Ƿ�Ϊ����
Private Function CheckIsNumber(bytInput As Byte) As Boolean
    Dim strTmp As String
    strTmp = "0123456789"
    CheckIsNumber = (InStr(strTmp, bytInput) > 0)
    
End Function


Private Function GetPatiName(ByVal strInput As String) As String
    Dim intLen As Integer
    Dim strTmp As String
    Dim strsql As String
    Dim blnTmp As Boolean
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    
    blnTmp = True
    strTmp = Trim(strInput)
    intLen = Len(strTmp)
    
    strsql = "Select distinct ���� From ������Ϣ "
    
    If InStr("*-+", Mid(strTmp, 1, 1)) > 0 Then
        For n = 2 To intLen
            If InStr("0123456789", Mid(strTmp, n, 1)) = 0 Then
                blnTmp = False
                Exit For
            End If
        Next
        If blnTmp = True Then
            Select Case Mid(strTmp, 1, 1)
                Case "*"
                    strsql = strsql & " where �����=[1]"
                Case "+"
                    strsql = strsql & " where סԺ��=[1]"
                Case "-"
                    strsql = strsql & " where ����ID=[1]"
                Case Else
            End Select
        Else
            GetPatiName = strInput
            Exit Function
        End If
    Else
        GetPatiName = strInput
        Exit Function
    End If
    
    On Error GoTo err

    Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, Val(Mid(txtPati.Text, 2)))
    
    If rs.RecordCount > 0 Then
        GetPatiName = rs!����
    Else
        GetPatiName = strInput
    End If
    
    rs.Close
    Exit Function
   
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Get��������()
    Dim rsData As ADODB.Recordset
    Dim int������� As Integer
    Dim int���ʱ�� As Integer
    
    On Error GoTo errHandle
    If gstrParentName = "frmҩƷ������ҩNew" Then
        mlng��ҩҩ��ID = Val(zlDatabase.GetPara("��ҩҩ��", glngSys, 1341, 0))
    ElseIf gstrParentName = "Frm���ŷ�ҩ����New" Then
        mlng��ҩҩ��ID = Val(zlDatabase.GetPara("��ҩҩ��", glngSys, 1342, 0))
        If mlng��ҩҩ��ID = 0 Then
            mlng��ҩҩ��ID = mlngҩ��ID
        End If
    End If
    
    int������� = zlDatabase.GetPara(241, glngSys, , 0)
    int���ʱ�� = zlDatabase.GetPara(242, glngSys, , 0)
    mstrDeptNode = GetDeptStationNode(mlng��ҩҩ��ID)
    mbln������ = ((int������� = 1 Or int������� = 3) And int���ʱ�� = 2)
        
    
    If mlng��ҩҩ��ID > 0 Then
        gstrSQL = "Select ���� From ���ű� Where ID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Get��������", mlng��ҩҩ��ID)
        
        If rsData.RecordCount > 0 Then
            mstrҩ�� = rsData!����
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'ȡϵͳ����
Private Sub Getϵͳ����()
    Dim rs As New ADODB.Recordset
    Dim intTmp As Integer
    Dim n As Integer
    
   '�����շ��뷢ҩ����
    On Error GoTo errHandle
    mstr�����շ��뷢ҩ���� = zlDatabase.GetPara(15, glngSys, , "0")
    'סԺ�����뷢ҩ����
    mstrסԺ�����뷢ҩ���� = zlDatabase.GetPara(16, glngSys, , "0")
    
    '���ý���λ��
    intTmp = CInt(zlDatabase.GetPara(9, glngSys))
    mstrSystemNumericFormat = "0."
    
    For n = 1 To intTmp
        mstrSystemNumericFormat = mstrSystemNumericFormat & "0"
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniControls()
    '����Ĭ�Ͽ�ȡ��߶�
    Me.Width = M_STR_FRM_WIDTH
    Me.Height = M_STR_FRM_HEIGHT
    
    '���ڻ�ؼ�����
    Me.Caption = mstrFrmCaption
    fraSelect.Caption = M_STR_FRASELECT_CAPTION
    fraGrid.Caption = mstrFraGridUnCheckCaption
    fraSelectFlag.Caption = mstrFraSelectFlagCaption
    CmdOK.Caption = M_STR_CMDOK_CHECK
    
    '�ؼ���ʾ��Ϣ
    txtPati.ToolTipText = M_STR_PATI_INPUT_DESC
    txtBillNo.ToolTipText = M_STR_BILLNO_DESC
    mshBill.ToolTipText = M_STR_GRID_DESC
    fraSelectFlag.ToolTipText = M_STR_FLAG_DESC
    
    If zlStr.IsHavePrivs(mstrPrivs, mstrֹͣ) = False Then
        optUnFlag.Enabled = False
        optFlag.Value = True
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, mstr�ָ�) = False Then
        optFlag.Enabled = False
        optUnFlag.Value = True
    End If
    
End Sub

Private Sub IniGrid()
    With mBillCol
        .BillCols = 19
    
        .Flag = ""
        .FlagCol = 0
        .FlagColWidth = 400
        .FlagColAlig = 1
        
        .Tag = "���"
        .TagCol = 1
        .TagColWidth = 0
        .TagColAlig = 1
           
        .Pati = "��������"
        .PatiCol = 2
        .PatiColWidth = 1000
        .PatiColAlig = 1
        
        .Bill = "������"
        .BILLCOL = 3
        .BillColWidth = 800
        .BillColAlig = 1
        
        .NO = "���"
        .NoCol = 4
        .NoColWidth = 600
        .NoColAlig = 1
        
        Select Case mintBillType
            Case 1
                .Drug = "ҩƷ����"
            Case 2
                .Drug = "��������"
            Case Else
        End Select
        .DrugCol = 5
        .DrugColWidth = 1500
        .DrugColAlig = 1
        
        .Spec = "���"
        .SpecCol = 6
        .SpecColWidth = 1000
        .SpecColAlig = 1
        
        .Unit = "��λ"
        .UnitCol = 7
        .UnitColWidth = 600
        .UnitColAlig = 1
        
        .UnitPrice = "����"
        .UnitPriceCol = 8
        .UnitPriceColWidth = 800
        .UnitPriceColAlig = 7
        .UnitPriceFormat = mstrSystemNumericFormat
        
        .count = "����"
        .CountCol = 9
        Select Case mintBillType
            Case 1
                .CountColWidth = 600
            Case 2
                .CountColWidth = 0
            Case Else
        End Select
        .CountColAlig = 7
        
        .Amount = "����"
        .AmountCol = 10
        .AmountColWidth = 800
        .AmountColAlig = 7
            
        .Price = "���"
        .PriceCol = 11
        .PriceColWidth = 1000
        .PriceColAlig = 7
        .PriceColFormat = mstrSystemNumericFormat
        
        .Group = "����"
        .GroupCol = 12
        .GroupColWidth = 600
        .GroupColAlig = 7
        
        .BillDate = "��������"
        .BillDateCol = 13
        .BillDateColWidth = 2000
        .BillDateColAlig = 1
        
        .Category = "��ҩ��ʽ"
        .CategoryCol = 14
        .CategoryColWidth = 0
        .CategoryColAlig = 1
        
        .Id = "Id"
        .IdCol = 15
        .IdColWidth = 0
        .IdColAlig = 7
    
        .��¼���� = "��¼����"
        .��¼����Col = 16
        .��¼����ColWidth = 0
        .��¼����ColAlig = 7
        
        .�����־ = "�����־"
        .�����־Col = 17
        .�����־ColWidth = 0
        .�����־ColAlig = 7
        
        .ȱҩ = "ȱҩ"
        .ȱҩCol = 18
        .ȱҩColWidth = 0
        .ȱҩColAlig = 7
    End With
    
    With mshBill
        .Clear
        .Cols = mBillCol.BillCols
        .rows = 1
        .rows = 2
        
        .TextMatrix(0, mBillCol.AmountCol) = mBillCol.Amount
        .TextMatrix(0, mBillCol.BILLCOL) = mBillCol.Bill
        .TextMatrix(0, mBillCol.BillDateCol) = mBillCol.BillDate
        .TextMatrix(0, mBillCol.CategoryCol) = mBillCol.Category
        .TextMatrix(0, mBillCol.CountCol) = mBillCol.count
        .TextMatrix(0, mBillCol.DrugCol) = mBillCol.Drug
        .TextMatrix(0, mBillCol.GroupCol) = mBillCol.Group
        .TextMatrix(0, mBillCol.NoCol) = mBillCol.NO
        .TextMatrix(0, mBillCol.PatiCol) = mBillCol.Pati
        .TextMatrix(0, mBillCol.PriceCol) = mBillCol.Price
        .TextMatrix(0, mBillCol.SpecCol) = mBillCol.Spec
        .TextMatrix(0, mBillCol.UnitCol) = mBillCol.Unit
        .TextMatrix(0, mBillCol.UnitPriceCol) = mBillCol.UnitPrice
        .TextMatrix(0, mBillCol.FlagCol) = mBillCol.Flag
        .TextMatrix(0, mBillCol.TagCol) = mBillCol.Tag
        .TextMatrix(0, mBillCol.IdCol) = mBillCol.Id
        .TextMatrix(0, mBillCol.��¼����Col) = mBillCol.��¼����
        .TextMatrix(0, mBillCol.�����־Col) = mBillCol.�����־
        .TextMatrix(0, mBillCol.ȱҩCol) = mBillCol.ȱҩ
        
        .ColWidth(mBillCol.AmountCol) = mBillCol.AmountColWidth
        .ColWidth(mBillCol.BILLCOL) = mBillCol.BillColWidth
        .ColWidth(mBillCol.BillDateCol) = mBillCol.BillDateColWidth
        .ColWidth(mBillCol.CategoryCol) = mBillCol.CategoryColWidth
        .ColWidth(mBillCol.CountCol) = mBillCol.CountColWidth
        .ColWidth(mBillCol.DrugCol) = mBillCol.DrugColWidth
        .ColWidth(mBillCol.GroupCol) = mBillCol.GroupColWidth
        .ColWidth(mBillCol.NoCol) = mBillCol.NoColWidth
        .ColWidth(mBillCol.PatiCol) = mBillCol.PatiColWidth
        .ColWidth(mBillCol.PriceCol) = mBillCol.PriceColWidth
        .ColWidth(mBillCol.SpecCol) = mBillCol.SpecColWidth
        .ColWidth(mBillCol.UnitCol) = mBillCol.UnitColWidth
        .ColWidth(mBillCol.UnitPriceCol) = mBillCol.UnitPriceColWidth
        .ColWidth(mBillCol.FlagCol) = mBillCol.FlagColWidth
        .ColWidth(mBillCol.TagCol) = mBillCol.TagColWidth
        .ColWidth(mBillCol.IdCol) = mBillCol.IdColWidth
        .ColWidth(mBillCol.��¼����Col) = mBillCol.��¼����ColWidth
        .ColWidth(mBillCol.�����־Col) = mBillCol.�����־ColWidth
        .ColWidth(mBillCol.ȱҩCol) = mBillCol.ȱҩColWidth
        
        .ColAlignment(mBillCol.AmountCol) = mBillCol.AmountColAlig
        .ColAlignment(mBillCol.BILLCOL) = mBillCol.BillColAlig
        .ColAlignment(mBillCol.BillDateCol) = mBillCol.BillDateColAlig
        .ColAlignment(mBillCol.CategoryCol) = mBillCol.CategoryColAlig
        .ColAlignment(mBillCol.CountCol) = mBillCol.CountColAlig
        .ColAlignment(mBillCol.DrugCol) = mBillCol.DrugColAlig
        .ColAlignment(mBillCol.GroupCol) = mBillCol.GroupColAlig
        .ColAlignment(mBillCol.NoCol) = mBillCol.NoColAlig
        .ColAlignment(mBillCol.PatiCol) = mBillCol.PatiColAlig
        .ColAlignment(mBillCol.PriceCol) = mBillCol.PriceColAlig
        .ColAlignment(mBillCol.SpecCol) = mBillCol.SpecColAlig
        .ColAlignment(mBillCol.UnitCol) = mBillCol.UnitColAlig
        .ColAlignment(mBillCol.UnitPriceCol) = mBillCol.UnitPriceColAlig
        .ColAlignment(mBillCol.FlagCol) = mBillCol.FlagColAlig
        .ColAlignment(mBillCol.TagCol) = mBillCol.TagColAlig
        .ColAlignment(mBillCol.IdCol) = mBillCol.IdColAlig
        .ColAlignment(mBillCol.��¼����Col) = mBillCol.��¼����ColAlig
        .ColAlignment(mBillCol.�����־Col) = mBillCol.�����־ColAlig
        .ColAlignment(mBillCol.ȱҩCol) = mBillCol.ȱҩColAlig
    End With
    
    Dim n As Long
    With mshBill
        .Row = 0
        For n = 0 To .Cols - 1
            .Col = n
            .CellBackColor = M_LNG_FIXEDCOLS_COLOR
        Next
    End With
    
End Sub
Public Function GetUnit(ByVal lngҩ��ID As Long) As String
    '����ָ���ⷿ�����ݡ�NO���õ�ҩƷ��λ
    Dim intUnit As Integer
    Dim rstemp As New ADODB.Recordset
    
    '����ϵͳ�����趨�ĵ�λ��ʾ����
    Select Case mintBillType
        Case 1      'ȡҩƷ��λ
            intUnit = Val(zlDatabase.GetPara("ҩ������", glngSys, 1341))
                If intUnit = 0 Then
                    'ȡ��ǰ�����Ĳ�����Դ
                    intUnit = mint��������
                End If
                If intUnit = 1 Then
                    GetUnit = GetSpecUnit1(lngҩ��ID, 2)
                Else
                    GetUnit = GetSpecUnit1(lngҩ��ID, 3)
                End If
        Case 2      'ȡ���ĵ�λ
            intUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, 1723, "0"))
            GetUnit = IIf(intUnit = 1, "��װ��λ", "�ۼ۵�λ")
        Case Else
    End Select
    
End Function

Public Function GetSpecUnit1(ByVal lng�ⷿid As Long, ByVal int��Χ As Integer) As String
    Dim strobjTemp As String                    '�����������ַ���
    Dim strWorkTemp As String                   '���湤�������ַ���
    Dim strUnit As String
    Dim rsProperty As New ADODB.Recordset
    Dim strsql As String
    
    
    '����ָ���ָⷿ�����÷�Χ�ĵ�λ
    On Error GoTo ErrHand
    
    gstrSQL = "Select Nvl(����,1) AS ��λ From ҩƷ�ⷿ��λ Where �ⷿID=[1] And ���÷�Χ=[2]"
    
    Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��λ", lng�ⷿid, int��Χ)
    
   
    If rsProperty.RecordCount = 1 Then
        strUnit = rsProperty!��λ
    Else
'        MsgBox "�ÿⷿδ���ÿⷿ��λ�����ݲ��������Լ��������ȡȱʡ��λ��" & _
'            vbCrLf & "ȱʡ��λ�Ĺ���" & _
'            vbCrLf & "  ���������סԺ�������סԺ�ģ�ȡסԺ��λ" & _
'            vbCrLf & "  ������������ģ�ȡ���ﵥλ" & _
'            vbCrLf & "  ����ҩ�����Եģ�ȡҩ�ⵥλ" & _
'            vbCrLf & "  ����ȡ�ۼ۵�λ", vbInformation, gstrSysName
        
        gstrSQL = "SELECT distinct �������,�������� From ��������˵�� Where ����ID =[1]"
        
        Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ��λ", lng�ⷿid)
            
        'ȡ������󼰲�������
        With rsProperty
            Do While Not .EOF
                strobjTemp = strobjTemp & .Fields(0)
                strWorkTemp = strWorkTemp & .Fields(1)
                .MoveNext
            Loop
            .Close
        End With
        If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
            'סԺ��λ
            strUnit = 3
        ElseIf InStr(strobjTemp, "1") <> 0 Then
            '���ﵥλ
            strUnit = 2
        ElseIf InStr(strWorkTemp, "ҩ��") <> 0 Then
            'ҩ�ⵥλ
            strUnit = 4
        Else
            '�ۼ۵�λ����Ҫ���Ƽ���
            strUnit = 1
        End If
    End If
    
    'ת��Ϊ��ʵ�ĵ�λ���ظ�������
    GetSpecUnit1 = Switch(strUnit = 1, "�ۼ۵�λ", strUnit = 2, "���ﵥλ", strUnit = 3, "סԺ��λ", strUnit = 4, "ҩ�ⵥλ")
    If glngSys / 100 = 8 Then
        'ҩ��ֻ���ۼ۵�λ��ҩ�ⵥλ
        GetSpecUnit1 = IIf(strUnit = 1, "�ۼ۵�λ", "ҩ�ⵥλ")
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub GetDetailBill()
    Dim strUnit As String           'ҩƷ��λ
    Dim rs As New ADODB.Recordset
    Dim strInputFlag As String
    Dim strSubSql As String
    Dim strsql As String
    Dim n As Long
    Dim strStartDate As String
    Dim strEndDate As String
    Dim strTmp As String
    
    On Error GoTo err
    
    strStartDate = Format(dtpStartDate, "yyyy-mm-dd 00:00:01")
    strEndDate = Format(dtpEndDate, "yyyy-mm-dd 23:59:59")
    
    Call IniGrid
    
    '''''''''''''����SQL���
    
    ''''select �Ӿ�
    strsql = "select distinct "
    strsql = strsql & " b.���� as " & mBillCol.Pati & ","
    strsql = strsql & " a.No as " & mBillCol.Bill & ","
    strsql = strsql & " a.��� as " & mBillCol.NO & ","
    strsql = strsql & " Decode(e.����,NULL,d.����,e.����) as " & mBillCol.Drug & ","
    strsql = strsql & " d.��� as " & mBillCol.Spec & ","
    strsql = strsql & " a.���ۼ� as " & mBillCol.UnitPrice & ","
    strsql = strsql & " a.���۽�� as " & mBillCol.Price & ","
    strsql = strsql & " NVL(a.����,0) as " & mBillCol.Group & ","
    strsql = strsql & " a.�������� as " & mBillCol.BillDate & ","
    strsql = strsql & " NVL(a.��ҩ��ʽ,-999) as " & mBillCol.Category & ","
    strsql = strsql & " NVL(a.����,1) as " & mBillCol.count & ","
    strsql = strsql & " a.id as " & mBillCol.Id & ","
    strsql = strsql & " b.��¼���� as " & mBillCol.��¼���� & ","
    strsql = strsql & " b.�����־ as " & mBillCol.�����־ & ","
    
        
    'ҩƷ��λ�����ۡ�����
    strUnit = GetUnit(mlng��ҩҩ��ID)
'    strUnit = "סԺ��λ"
    Select Case strUnit
    Case "�ۼ۵�λ"
        strSubSql = "1"
        strsql = strsql & " D.���㵥λ as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.���ۼ�,'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.ʵ������,'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "���ﵥλ"
        strSubSql = "Decode(�����װ,Null,1,0,1,�����װ)"
        strsql = strsql & " F.���ﵥλ as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.���ۼ�*Decode(F.�����װ,Null,1,0,1,F.�����װ),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.ʵ������/Decode(F.�����װ,Null,1,0,1,F.�����װ),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "סԺ��λ"
        strSubSql = "Decode(סԺ��װ,Null,1,0,1,סԺ��װ)"
        strsql = strsql & " F.סԺ��λ as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.���ۼ�*Decode(F.סԺ��װ,Null,1,0,1,F.סԺ��װ),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.ʵ������/Decode(F.סԺ��װ,Null,1,0,1,F.סԺ��װ),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "ҩ�ⵥλ"
        strSubSql = "Decode(ҩ���װ,Null,1,0,1,ҩ���װ)"
        strsql = strsql & " F.ҩ�ⵥλ as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.���ۼ�*Decode(F.ҩ���װ,Null,1,0,1,F.ҩ���װ),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.ʵ������/Decode(F.ҩ���װ,Null,1,0,1,F.ҩ���װ),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    Case "��װ��λ"         '����
        strSubSql = "Decode(��װ��λ,Null,1,0,1,��װ��λ)"
        strsql = strsql & " F.��װ��λ as " & mBillCol.Unit & ","
        strsql = strsql & " ltrim(to_char(A.���ۼ�*Decode(F.����ϵ��,Null,1,0,1,F.����ϵ��),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.UnitPrice & ","
        strsql = strsql & " ltrim(to_char(A.ʵ������/Decode(F.����ϵ��,Null,1,0,1,F.����ϵ��),'" & M_STR_DEFAULT_ORA_NUMERIC_FORMAT & "')) as " & mBillCol.Amount
    End Select
    
    If mIntCheckStock > 0 And mblnIsChecked = True Then
        strsql = strsql & ",Decode(Sign(Nvl(K.��������, 0) - A.ʵ������ * Nvl(A.����, 1)), -1, 1, 0) " & mBillCol.ȱҩ
    Else
        strsql = strsql & ",0 " & mBillCol.ȱҩ
    End If
    
    ''''from�Ӿ�
    strsql = strsql & " From ҩƷ�շ���¼ A,������ü�¼ B,���ű� C,�շ���ĿĿ¼ D,�շ���Ŀ���� E,���ű� P " & IIf(mbln������, ",��������¼ Q,���������ϸ K ", "")
    Select Case mintBillType
        Case 1
            strsql = strsql & ",ҩƷ��� F"
        Case 2
            strsql = strsql & ",�������� F"
        Case Else
    End Select
    
    'Ҫ�����ʱ������ҩƷ����
    If mIntCheckStock > 0 Then
        strsql = strsql & ",(Select ҩƷid, Nvl(����, 0) ����, Nvl(��������, 0) �������� From ҩƷ��� Where ���� = 1 And �ⷿid = [1]) K "
    End If
    
    ''''where �Ӿ�
    strsql = strsql & " where A.����id=B.Id And A.ҩƷid=D.Id And D.ID=E.�շ�ϸĿID(+) And B.�շ�ϸĿid=D.Id " & IIf(mbln������, " and b.ҽ�����=k.ҽ��id(+) and Q.id(+)=K.��id and K.����ύ(+)=1 And (b.ҽ����� is null or nvl(q.�����,0) = 1)", "") & _
         " And A.�ⷿid+0=C.Id  AND E.����(+)=3 And A.�Է�����ID = P.ID And Nvl(B.����״̬,0)<>1 "
    
    If mstrDeptNode <> "" Then
        strsql = strsql & " And (P.վ�� = [6] Or P.վ�� Is Null)"
    End If
    
    Select Case mintBillType
        Case 1
            strsql = strsql & " AND A.ҩƷid=F.ҩƷid "
        Case 2
            strsql = strsql & " AND A.ҩƷid=F.����id "
        Case Else
    End Select
    
    '����Ĳ�ѯ����
    strsql = strsql & " and Mod(A.��¼״̬,3)=1 and A.����� is null "                 'δ��ҩ����
        
    '���ݲ����͸����ڴ��ݵĲ�ѯ����
    If mint�������� = 1 Then
        strsql = strsql & " and A.���� in(8,9) "
        If mstr�����շ��뷢ҩ���� = "0" Then
            strsql = strsql & " and A.�ⷿid+0=[1]"
        End If
    ElseIf mint�������� = 2 Then
        '����ǲ��ŷ�ҩ����ֻ��ѯסԺ���ü�¼
        strsql = strsql & " and A.���� in(9,10) "
        If mstrסԺ�����뷢ҩ���� = "0" Then
            strsql = strsql & " and A.�ⷿid+0=[1]"
        End If
    Else
        strsql = strsql & " and A.�ⷿid+0=[1]"
    End If
        
    '�û�ѡ��Ĳ�ѯ����
    strsql = strsql & " and A.��������>=[2]  and A.��������<=[3] "
    
    If txtPati.Text <> "" Then
        strsql = strsql & " and B.����=[4] "
    End If
    
    If txtBillNo.Text <> "" Then
        strsql = strsql & " and A.no=[5] "
    End If
    
    If Not mblnIsChecked Then
        strsql = strsql & " and NVL(A.��ҩ��ʽ,-999)<>-1"
    Else
        strsql = strsql & " and A.��ҩ��ʽ=-1"
        
        If mIntCheckStock > 0 Then
            strsql = strsql & " And A.ҩƷid = K.ҩƷid(+) And Nvl(A.����, 0) = K.����(+) "
        End If
    End If
    
    If mint�������� = 1 Then
        '����Ǵ�����ҩ����ϲ�������ü�¼��סԺ���ü�¼
        strTmp = Replace(strsql, "b.����", "nvl(R.����,b.����)")
        strTmp = Replace(strTmp, "B.����", "nvl(R.����,b.����)")
        strTmp = Replace(strTmp, "������ü�¼ B", "סԺ���ü�¼ B,������ҳ R")
        strTmp = Replace(strTmp, "And Nvl(B.����״̬,0)<>1", " And B.����id=R.����id And B.��ҳID=R.��ҳID ")
        strTmp = Replace(strTmp, "in(8,9)", "in(9,10)")
        
        strsql = strsql & " Union All " & strTmp
        
        ''''order�Ӿ�
        strsql = strsql & " order by ������,��� "
    ElseIf mint�������� = 0 Or mint�������� = 2 Then
        If mint�������� = 0 Then
            strTmp = Replace(strsql, "b.����", "nvl(R.����,b.����)")
            strTmp = Replace(strTmp, "B.����", "nvl(R.����,b.����)")
            strTmp = Replace(strTmp, "������ü�¼ B", "סԺ���ü�¼ B,������ҳ R")
            strTmp = Replace(strTmp, "And Nvl(B.����״̬,0)<>1", " And B.����id=R.����id And B.��ҳID=R.��ҳID ")
            
            strsql = strsql & " Union All " & strTmp
        ElseIf mint�������� = 2 Then
            strsql = Replace(strsql, "b.����", "nvl(R.����,b.����)")
            strsql = Replace(strsql, "B.����", "nvl(R.����,b.����)")
            strsql = Replace(strsql, "������ü�¼ B", "סԺ���ü�¼ B,������ҳ R")
            strsql = Replace(strsql, "And Nvl(B.����״̬,0)<>1", " And B.����id=R.����id And B.��ҳID=R.��ҳID ")
        End If
        
        ''''order�Ӿ�
        strsql = strsql & " order by ������,��� "
    Else
        ''''order�Ӿ�
        strsql = strsql & " order by A.No,A.��� "
    End If
    
    
    
    '''''''''''''���Ϲ���SQL���
    
    
    ''''��ѯ���
    Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlng��ҩҩ��ID, CDate(strStartDate), CDate(strEndDate), txtPati.Text, txtBillNo.Text, mstrDeptNode)
    
        
    DoEvents
    Me.MousePointer = 11
        
    With mshBill
        .Redraw = False

        If (Not rs.EOF) And (Not rs.BOF) Then
            CmdOK.Enabled = True
        End If
        Do While Not rs.EOF
        
            If .rows >= 2 And .TextMatrix(1, 1) <> "" Then
                .rows = .rows + 1
            End If
            .TextMatrix(.rows - 1, mBillCol.AmountCol) = FormatEx(rs.Fields(mBillCol.Amount).Value, mintNumberDigit)
            .TextMatrix(.rows - 1, mBillCol.BILLCOL) = rs.Fields(mBillCol.Bill).Value
            .TextMatrix(.rows - 1, mBillCol.BillDateCol) = rs.Fields(mBillCol.BillDate).Value
            .TextMatrix(.rows - 1, mBillCol.CategoryCol) = rs.Fields(mBillCol.Category).Value
            .TextMatrix(.rows - 1, mBillCol.CountCol) = rs.Fields(mBillCol.count).Value
            .TextMatrix(.rows - 1, mBillCol.DrugCol) = rs.Fields(mBillCol.Drug).Value
            .TextMatrix(.rows - 1, mBillCol.GroupCol) = rs.Fields(mBillCol.Group).Value
            .TextMatrix(.rows - 1, mBillCol.NoCol) = rs.Fields(mBillCol.NO).Value
            .TextMatrix(.rows - 1, mBillCol.PatiCol) = IIf(IsNull(rs.Fields(mBillCol.Pati).Value), "", rs.Fields(mBillCol.Pati).Value)
            .TextMatrix(.rows - 1, mBillCol.PriceCol) = Format(rs.Fields(mBillCol.Price).Value, mBillCol.PriceColFormat)
            .TextMatrix(.rows - 1, mBillCol.SpecCol) = IIf(IsNull(rs.Fields(mBillCol.Spec).Value), "", rs.Fields(mBillCol.Spec).Value)
            .TextMatrix(.rows - 1, mBillCol.UnitCol) = IIf(IsNull(rs.Fields(mBillCol.Unit).Value), "", rs.Fields(mBillCol.Unit).Value)
            .TextMatrix(.rows - 1, mBillCol.UnitPriceCol) = Format(rs.Fields(mBillCol.UnitPrice).Value, mBillCol.UnitPriceFormat)
            .TextMatrix(.rows - 1, mBillCol.IdCol) = rs.Fields(mBillCol.Id).Value
            .TextMatrix(.rows - 1, mBillCol.��¼����Col) = rs.Fields(mBillCol.��¼����).Value
            .TextMatrix(.rows - 1, mBillCol.�����־Col) = rs.Fields(mBillCol.�����־).Value
            .TextMatrix(.rows - 1, mBillCol.ȱҩCol) = rs.Fields(mBillCol.ȱҩ).Value
            
            .TextMatrix(.rows - 1, mBillCol.FlagCol) = ""
            
            .Col = 0
            .Row = .rows - 1
            
            .TextMatrix(.rows - 1, mBillCol.TagCol) = M_STR_UNCHECK_NAME
            Set .CellPicture = LoadResPicture("unchecked", vbResBitmap)
            
            'ȱҩҩƷ�ú�ɫ������
            If mIntCheckStock > 0 And Val(.TextMatrix(.rows - 1, mBillCol.ȱҩCol)) = 1 Then
                For n = 0 To .Cols - 1
                    .Col = n
                    .CellForeColor = vbRed
                Next
            End If
            
            rs.MoveNext
        Loop
        .Col = 0
        .Row = 0
        
        .TextMatrix(.rows - 1, mBillCol.TagCol) = M_STR_UNCHECK_NAME
        Set mshBill.CellPicture = LoadResPicture("unchecked", vbResBitmap)
        
        .Redraw = True
    End With
    
    DoEvents
    Me.MousePointer = 0
    
    Exit Sub

err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdCancel_Click()
'    Call SaveFlexState(mshBill, Me.Name)
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim n As Long
    Dim int�շ��뷢ҩ���� As Integer
    Dim int���� As Integer
    Dim blnBeginTrans As Boolean
    Dim arrSql As Variant
    
    If mshBill.rows < 2 Then
        Exit Sub
    End If
    
    Select Case mintBillType
        Case 1
            If mint�������� = 1 Then
                int�շ��뷢ҩ���� = mstr�����շ��뷢ҩ����
            Else
                int�շ��뷢ҩ���� = mstrסԺ�����뷢ҩ����
            End If
        Case 2
            int�շ��뷢ҩ���� = 0
        Case Else
    End Select
    
    mblnIsChecked = optFlag.Value
    arrSql = Array()
    
    On Error GoTo err
    For n = 1 To mshBill.rows - 1
        If mblnIsChecked = True Then
            If mshBill.TextMatrix(n, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                If Val(mshBill.TextMatrix(n, mBillCol.��¼����Col)) = 1 Or (Val(mshBill.TextMatrix(n, mBillCol.��¼����Col)) = 2 And (Val(mshBill.TextMatrix(n, mBillCol.�����־Col))) = 1 Or (Val(mshBill.TextMatrix(n, mBillCol.�����־Col))) = 4) Then
                    int���� = 1
                Else
                    int���� = 2
                End If
                
                gstrSQL = "Zl_����ҩ�������_Unchecked(" & mshBill.TextMatrix(n, mBillCol.IdCol) & "," & int�շ��뷢ҩ���� & "," & int���� & ")"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        End If
        If mblnIsChecked = False Then
            If mshBill.TextMatrix(n, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                gstrSQL = "Zl_����ҩ�������_Checked(" & mshBill.TextMatrix(n, mBillCol.IdCol) & "," & int�շ��뷢ҩ���� & ")"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        End If
    Next
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For n = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(n)), Me.Caption & "-�����")
    Next
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    Call GetDetailBill
    
    Exit Sub
            
err:
    '����ѿ������񣬲���δ�ύ�������ʱ�ع�����
    If blnBeginTrans Then
        gcnOracle.RollbackTrans
    End If
    
    MsgBox "��ʾ������ʧ�ܡ�"
    Call SaveErrLog
End Sub



Private Sub cmdRefresh_Click()
    If mblnIsChecked Then
        CmdOK.Caption = M_STR_CMDOK_UNCHECK
        fraGrid.Caption = mstrFraGridCheckCaption
    Else
        CmdOK.Caption = M_STR_CMDOK_CHECK
        fraGrid.Caption = mstrFraGridUnCheckCaption
    End If

    Call GetDetailBill
End Sub

'
Private Sub dtpStartDate_Change()
'         Call GetDetailBill
End Sub


Private Sub Form_Load()
    Dim intUnit As Integer
    Dim rstemp As Recordset
    
    mstrPrivs = gstrprivs
    
    dtpStartDate.Value = Format(Date, "yyyy-mm-01")
    dtpEndDate.Value = Format(Date, "yyyy-mm-dd")
    
    mblnIsChecked = False
    
    mstrSystemAmountFormat = "0"
    
   
      
    Select Case mintBillType
        Case 1
            mstrFrmCaption = M_STR_ҩƷ_FRM_CAPTION
            mstrFraGridCheckCaption = M_STR_ҩƷ_FRAGRID_CHECK_CAPTION
            mstrFraGridUnCheckCaption = M_STR_ҩƷ_FRAGRID_UNCHECK_CAPTION
            mstrFraSelectFlagCaption = M_STR_ҩƷ_FRASELECTFLAG_CAPTION
        Case 2
            mstrFrmCaption = M_STR_����_FRM_CAPTION
            mstrFraGridCheckCaption = M_STR_����_FRAGRID_CHECK_CAPTION
            mstrFraGridUnCheckCaption = M_STR_����_FRAGRID_UNCHECK_CAPTION
            mstrFraSelectFlagCaption = M_STR_����_FRASELECTFLAG_CAPTION
        Case Else
    End Select
    
    Call Getϵͳ����
    Call Get��������
    
    Call IniGrid
    
    Me.Caption = Me.Caption & "-" & mstrҩ��
    If mIntCheckStock = 1 Then
        lblComment.Caption = "������������ҩƷ�ú�ɫ�����ʶ��"
    Else
        lblComment.Caption = "������������ҩƷ�ú�ɫ�����ʶ�����ָܻ���־��"
    End If
    
     '���ݸ�������������жϴ������ͣ�������������Ƹı䣬����ҲҪ����Ӧ�ı�
    Select Case gstrParentName
        Case "frmҩƷ������ҩNew"
            mint�������� = 1
            mintBillType = 1
            mstrֹͣ = "ֹͣ��ҩ"
            mstr�ָ� = "�ָ���ҩ"
            
            Select Case GetDrugUnit(mlng��ҩҩ��ID, Me.Caption)
                Case "�ۼ۵�λ"             '�ۼ۵�λ����Ҫ���Ƽ���
                    intUnit = 1
                Case "���ﵥλ"
                    intUnit = 2
                Case "סԺ��λ"
                    intUnit = 3
                Case "ҩ�ⵥλ"
                    intUnit = 4
            End Select
    
            gstrSQL = "select ���� from ҩƷ���ľ��� where ����=0 and ��� = 1 And ���� = 3 And ��λ = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", intUnit)
            
        Case "Frm���ŷ�ҩ����New"
            mint�������� = 2
            mintBillType = 1
            mstrֹͣ = "ֹͣ��ҩ"
            mstr�ָ� = "�ָ���ҩ"
    
            Select Case GetDrugUnit(mlng��ҩҩ��ID, Me.Caption)
                Case "�ۼ۵�λ"             '�ۼ۵�λ����Ҫ���Ƽ���
                    intUnit = 1
                Case "���ﵥλ"
                    intUnit = 2
                Case "סԺ��λ"
                    intUnit = 3
                Case "ҩ�ⵥλ"
                    intUnit = 4
            End Select
    
            gstrSQL = "select ���� from ҩƷ���ľ��� where ����=0 and ��� = 1 And ���� = 3 And ��λ = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", intUnit)
        Case "frm���ķ��Ź���"
            mbln������ = False
            mintBillType = 2
            mstrֹͣ = "ֹͣ����"
            mstr�ָ� = "�ָ�����"
            
            '��ȡ��������
            intUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, 1723))
            
            gstrSQL = "select ���� from ҩƷ���ľ��� where ����=0 and ��� = 2 And ���� = 3 And ��λ = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ��������", intUnit)
            
            
        Case "frmStuffRxSend"
            mbln������ = False
            mintBillType = 2
            mstrֹͣ = "ֹͣ����"
            mstr�ָ� = "�ָ�����"
        Case "frmStuffDeptSend"
            mbln������ = False
            mintBillType = 2
            mstrֹͣ = "ֹͣ����"
            mstr�ָ� = "�ָ�����"
        Case Else
    End Select
    
    If Not rstemp.EOF Then
        mintNumberDigit = rstemp!����
    Else
        mintNumberDigit = 5
    End If
    
    Call IniControls
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    If Me.Width < M_STR_FRM_WIDTH Then Me.Width = M_STR_FRM_WIDTH
    If Me.Height < M_STR_FRM_HEIGHT Then Me.Height = M_STR_FRM_HEIGHT

'    With CmdHelp
'        .Top = Me.ScaleHeight - .Height - 100
'    End With
'
'    With CmdCancel
'        .Top = CmdHelp.Top
'        .Left = Me.ScaleWidth - .Width - 100
'    End With
'    With CmdOK
'        .Top = CmdHelp.Top
'        .Left = CmdCancel.Left - .Width - 100
'    End With
'
'    With fraSelect
'        .Width = Me.ScaleWidth - .Left - 50
'    End With
'
'    With fraGrid
'        .Height = CmdOK.Top - 1400
'        .Width = Me.ScaleWidth - .Left - 50
'    End With
'
'    With mshBill
'        .Left = 50
'        .Height = CmdOK.Top - 1600
'        .Width = fraGrid.Width - .Left - 150
'    End With
    
    With CmdHelp
        .Move .Left, Me.ScaleHeight - .Height - 100
    End With

    With CmdCancel
        .Move Me.ScaleWidth - .Width - 100, CmdHelp.Top
    End With
    
    With CmdOK
        .Move CmdCancel.Left - .Width - 100, CmdHelp.Top
    End With

    With fraSelect
        .Move .Left, .Top, Me.ScaleWidth - .Left - 50
    End With

    With fraGrid
        .Move .Left, .Top, Me.ScaleWidth - .Left - 50, CmdOK.Top - 1400
    End With

    With mshBill
        .Move 50, .Top, fraGrid.Width - .Left - 150, CmdOK.Top - 1600
    End With
    
    With lblComment
        .Top = CmdHelp.Top + 100
    End With
End Sub

Private Sub mshBill_Click()
    Dim n As Long
    Dim i As Long
    Dim lngColor As Long
    Dim lngCurRow As Long
    Dim lngCurCol As Long
    Dim blnWarn As Boolean
    Dim blnWarnDo As Boolean
    
'    Debug.Print "row:" & mshBill.Row & " col:" & mshBill.Col
    
    With mshBill
        .Redraw = False
        lngCurRow = .Row
        lngCurCol = .Col
        If .rows > 1 And .TextMatrix(.rows - 1, mBillCol.BILLCOL) <> "" Then
            '���ѡ����ǵ�һ�У�������ǻ�ȡ����Ǵ���
            If .Col = 0 Then
                If .Row = 0 Then
                    If .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                        For n = 0 To .rows - 1
                            .TextMatrix(n, mBillCol.TagCol) = M_STR_UNCHECK_NAME
                            .Row = n
                            .Col = mBillCol.FlagCol
                            Set .CellPicture = LoadResPicture("unchecked", vbResBitmap)
                            If .Row > 0 Then
                                For i = 0 To .Cols - 1
                                    .Row = n
                                    .Col = i
                                    .CellBackColor = M_LNG_UNCHECKED_COLOR
                                Next
                            End If
                        Next
                    Else
                        For n = 0 To .rows - 1
                            blnWarnDo = True
                            
                            'ȱҩ�������
                            If n > 0 And mIntCheckStock > 0 And Val(.TextMatrix(n, mBillCol.ȱҩCol)) = 1 Then
                                If mIntCheckStock = 1 Then
                                    '��治������ʱ����ȱҩ��¼���ѣ���������������д���
                                    If blnWarn = False Then
                                        blnWarn = True
                                        If MsgBox("���ڻָ���־����ÿ�治���ҩƷ���Ƿ�ָ���־��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                            blnWarnDo = False
                                        End If
                                     End If
                                ElseIf mIntCheckStock = 2 Then
                                    '�ϸ���ƿ��ʱ������ȱҩ��¼
                                    blnWarnDo = False
                                End If
                            End If
                            
                            If blnWarnDo = True Then
                                .TextMatrix(n, mBillCol.TagCol) = M_STR_CHECK_NAME
                                .Row = n
                                .Col = mBillCol.FlagCol
                                Set .CellPicture = LoadResPicture("checked", vbResBitmap)
                                If .Row > 0 Then
                                    For i = 0 To .Cols - 1
                                        .Row = n
                                        .Col = i
                                        .CellBackColor = M_LNG_CHECKED_COLOR
                                    Next
                                End If
                            End If
                        Next
                    End If
                ElseIf .Row > 0 Then
                    If .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                        .TextMatrix(.Row, mBillCol.TagCol) = M_STR_UNCHECK_NAME
                        .Col = mBillCol.FlagCol
                        Set .CellPicture = LoadResPicture("unchecked", vbResBitmap)
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = M_LNG_UNCHECKED_COLOR
                        Next
                    Else
                        'ȱҩ�������
                        If mIntCheckStock > 0 And Val(.TextMatrix(.Row, mBillCol.ȱҩCol)) = 1 Then
                            If mIntCheckStock = 1 Then
                                '��治������ʱ����ȱҩ��¼���ѣ���������������д���
                                If MsgBox("���ڻָ���־����ÿ�治���ҩƷ���Ƿ�ָ���־��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    Exit Sub
                                End If
                            ElseIf mIntCheckStock = 2 Then
                                '�ϸ���ƿ��ʱ������ȱҩ��¼
                                Exit Sub
                            End If
                        End If
                            
                        .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME
                        .Col = mBillCol.FlagCol
                        Set .CellPicture = LoadResPicture("checked", vbResBitmap)
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = M_LNG_CHECKED_COLOR
                        Next
                    End If
                
                End If
            '���ѡ��Ĳ��ǵ�һ�У�����ѡ���д���
            ElseIf .Row > 0 Then
                For n = 1 To .rows - 1
                    .Row = n
                    .Col = 0
                    If .CellBackColor = M_LNG_SELECTEDCOLS_COLOR Then
                        If .TextMatrix(.Row, mBillCol.TagCol) = M_STR_CHECK_NAME Then
                            lngColor = M_LNG_CHECKED_COLOR
                        Else
                            lngColor = M_LNG_DEFAULTCOLS_COLOR
                        End If
                        For i = 0 To .Cols - 1
                            .Col = i
                            .CellBackColor = lngColor
                        Next
                    End If
                Next
                .Row = lngCurRow
                .Col = lngCurCol
                lngColor = M_LNG_SELECTEDCOLS_COLOR
                For i = 0 To .Cols - 1
                    .Col = i
                    .CellBackColor = lngColor
                Next
            End If
        End If
        
        .Redraw = True
    End With
                        
End Sub


Private Sub optFlag_Click()
    mblnIsChecked = True
    CmdOK.Enabled = False
    
    If mIntCheckStock = 0 Then
        lblComment.Visible = False
    Else
        lblComment.Visible = True
    End If
End Sub

Private Sub optUnFlag_Click()
    mblnIsChecked = False
    CmdOK.Enabled = False
    
    lblComment.Visible = False
End Sub

Private Sub txtBillNo_GotFocus()
    Call zlControl.TxtSelAll(txtBillNo)

End Sub

Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        If Not (Len(Trim(txtBillNo.Text)) = 0 Or txtBillNo.SelLength = Len(txtBillNo.Text)) And _
            InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0:  Exit Sub
        End If
    End If
    
    If KeyAscii = 13 And txtBillNo.Text <> "" Then
        
        txtBillNo.Text = zlCommFun.GetFullNO(txtBillNo.Text, 13)
'        Call GetDetailBill
        
    End If

End Sub


Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(Trim(txtPati.Text)) <> 0 Then
'        If InStr(1, "*-+", Mid(Trim(txtPati.Text), 1, 1)) = 0 Then
'            txtPati.Text = "*" & Trim(txtPati.Text)
'        End If
        txtPati.Text = GetPatiName(txtPati.Text)
        Call OS.PressKey(vbKeyTab)
    End If
End Sub


Private Sub txtPati_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(Trim((txtPati.Text))) > 0 Then
        If InStr("*-+", Mid(txtPati.Text, 1, 1)) > 0 Then
            If InStr("0123456789", Chr(KeyCode)) = 0 Then
                Exit Sub
            End If
        End If
    End If
End Sub

Public Sub showMe(ByVal frmParent As Form, ByVal lngҩ��ID As Long)
    mlngҩ��ID = lngҩ��ID
    Me.Show 1, frmParent
End Sub




