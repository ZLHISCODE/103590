VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmAdjust 
   Caption         =   "ҩƷ���۵�"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   Icon            =   "frmAdjust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   10110
   StartUpPosition =   1  '����������
   Begin VB.CheckBox Chk���� 
      Caption         =   "ʱ��ҩƷ��Ϊ��������(&D)"
      Enabled         =   0   'False
      Height          =   210
      Left            =   2505
      TabIndex        =   6
      Top             =   3450
      Width           =   2370
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2865
      Left            =   5790
      TabIndex        =   15
      Top             =   30
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   2880
      Left            =   5775
      TabIndex        =   17
      Top             =   15
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   5080
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ZL9BillEdit.BillEdit bfgPrice 
      Height          =   2955
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5212
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
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
   Begin VB.CommandButton cmdCpt 
      Caption         =   "����ǰ������(&T)��"
      Height          =   350
      Left            =   6075
      Picture         =   "frmAdjust.frx":0442
      TabIndex        =   12
      Top             =   3915
      Width           =   1965
   End
   Begin VB.CommandButton cmdPstor 
      Caption         =   "��ӡ���䶯��(&S)��"
      Height          =   350
      Left            =   8085
      Picture         =   "frmAdjust.frx":058C
      TabIndex        =   13
      Top             =   3915
      Width           =   1965
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)��"
      Height          =   350
      Left            =   8880
      Picture         =   "frmAdjust.frx":06D6
      TabIndex        =   11
      Top             =   900
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8880
      Picture         =   "frmAdjust.frx":0820
      TabIndex        =   10
      Top             =   480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8880
      Picture         =   "frmAdjust.frx":096A
      TabIndex        =   9
      Top             =   45
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -435
      TabIndex        =   16
      Top             =   3810
      Width           =   16815
   End
   Begin VB.TextBox txtRegistrar 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   6285
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3015
      Width           =   2445
   End
   Begin MSComCtl2.DTPicker dtpRunDate 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   6285
      TabIndex        =   8
      Top             =   3390
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
      Format          =   127270915
      CurrentDate     =   36846.5833333333
   End
   Begin VB.TextBox txtSummary 
      Height          =   300
      Left            =   825
      TabIndex        =   2
      Top             =   3015
      Width           =   4485
   End
   Begin VB.CheckBox chkImmediately 
      Caption         =   "���м۸�������Ч(&I)"
      Height          =   210
      Left            =   75
      TabIndex        =   5
      Top             =   3450
      Width           =   2040
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdStore 
      Height          =   1815
      Left            =   30
      TabIndex        =   14
      Top             =   4305
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   14737632
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���䶯��"
      Height          =   180
      Left            =   60
      TabIndex        =   18
      Top             =   4050
      Width           =   1080
   End
   Begin VB.Label lblRegistrar 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   5655
      TabIndex        =   3
      Top             =   3075
      Width           =   540
   End
   Begin VB.Label lblRunDate 
      AutoSize        =   -1  'True
      Caption         =   "ִ������"
      Height          =   180
      Left            =   5475
      TabIndex        =   7
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label lblSummary 
      AutoSize        =   -1  'True
      Caption         =   "����˵��"
      Height          =   180
      Left            =   30
      TabIndex        =   1
      Top             =   3075
      Width           =   720
   End
End
Attribute VB_Name = "frmAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnImmediately As Boolean
'�Ƿ���Ҫ������Ч��ʱ��ҩƷ������Ҫ��
Public lngBillId As Long
'��������:0-���۴���;����-��ʾlngBillIdȷ������ʷ���۵�
Public lngMediId As Long
'��������:0-δָ������ҩƷ;����-����ʱֱ����ʾlngMediId��ԭ�۸����
Public intUnit As Integer   '0-�ۼ۵�λ;1-���ﵥλ;2-ҩ�ⵥλ;3-סԺ��λ
'��������:0-δָ������ҩƷ;����-����ʱֱ����ʾlngMediId��ԭ�۸����
Private BlnModify As Boolean

'--------���۵��г���--------------
Const conColҩƷid As Integer = 0
Const conColƷ�� As Integer = 1
Const conCol��� As Integer = 2
Const conCol���� As Integer = 3
Const conCol��λ As Integer = 4
Const conColԭ�� As Integer = 5
Const conCol�ּ� As Integer = 6
Const conCol����ID As Integer = 7
Const conColOld����ID As Integer = 8
Const conCol�������� As Integer = 9

'---------------------------------
Dim rsTemp As New ADODB.Recordset
Private StrFindStyle As String
Dim intCount As Integer
Dim objItem As ListItem
Dim objNode As Node
Dim dtToday As Date

Private Const mconintPriceBit As Integer = 7            '����С��λ��
Private Sub bfgPrice_CommandClick()
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long
    Dim RecReturn As New ADODB.Recordset
    
    On Error GoTo errHandle
    If Me.bfgPrice.Col = conColƷ�� Then
'        With Me.bfgPrice
'            Set RecReturn = FrmҩƷѡ����.ShowMe(Me, 1)
'            If RecReturn.EOF Then Exit Sub
'            LngmediIDThis = RecReturn!ҩƷID
'            If LngmediIDThis = 0 Then Exit Sub
'
'            '�Ǳ��ҩƷ���˳�
'            gstrSQL = " Select Nvl(�Ƿ���,0) ��� From ҩƷ��� Where ID=[1]"
'            Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
'            With RecCheck
'                '�����ʱ��ҩƷ�����ۼ�¼����������Ч
'                If !��� = 1 And blnImmediately = False Then
'                    blnImmediately = True
'                    chkImmediately.Value = 1
'                    chkImmediately.Enabled = False
'                End If
'                If !��� = 1 Then Chk����.Enabled = True
'            End With
'
'            If chkImmediately.Value = 1 Then
'
'                    '�ж��Ƿ���δִ�е���ʷ�۸�
'                gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And �շ�ϸĿID=[1]"
'                Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
'
'                With RecCheck
'                    If Not .EOF Then
'                        If Not IsNull(!Records) Then
'                            If !Records <> 0 And chkImmediately.Value = 1 Then
'                                MsgBox "��ҩƷ����δִ�м۸񣬲�������Ϊ����ִ�У�", vbInformation, gstrSysName
'                                If chkImmediately.Enabled Then
'                                    chkImmediately.Value = 0
'                                Else
'                                    Exit Sub
'                                End If
'                            End If
'                        End If
'                    End If
'                End With
'            End If
'
'            .TextMatrix(.Row, conColҩƷid) = RecReturn!ҩƷID
'            .TextMatrix(.Row, conColƷ��) = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!��Ʒ��
'            .TextMatrix(.Row, conCol���) = IIf(IsNull(RecReturn!���), "", RecReturn!���)
'            .TextMatrix(.Row, conCol����) = IIf(IsNull(RecReturn!����), "", RecReturn!����)
'            .TextMatrix(.Row, conCol��λ) = IIf(IsNull(RecReturn!�ۼ۵�λ), "", RecReturn!�ۼ۵�λ)
'            Call getMediPrice(RecReturn!ҩƷID)
'            .CmdVisible = False
'            .Col = conCol�ּ�
'            BlnModify = True
'        End With
    Else
        With Me.tvwItem
            .Left = bfgPrice.Left + bfgPrice.MsfObj.CellLeft
            .Top = Me.bfgPrice.Top + Me.bfgPrice.CellTop + Me.bfgPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            For intCount = 1 To .Nodes.count
                If InStr(1, .Nodes(intCount).Text & "-", Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, Me.bfgPrice.Col) & "-") > 0 Then
                    .Nodes(intCount).Selected = True
                    Exit For
                End If
            Next
            .SetFocus
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub bfgPrice_EnterCell(Row As Long, Col As Long)
    Select Case Col
    Case conCol��������
        Me.bfgPrice.TextMatrix(Row, conCol�ּ�) = zlStr.FormatEx(Me.bfgPrice.TextMatrix(Row, conCol�ּ�), 7)
    End Select
    
End Sub

Private Sub bfgPrice_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    If KeyCode <> 13 Then Exit Sub
    On Error GoTo errHandle
    Select Case Me.bfgPrice.Col
    Case conColƷ��
        If Trim(Me.bfgPrice.Text) = "" Then Exit Sub
        strInput = UCase(Me.bfgPrice.Text)
        
        Me.lvwItem.Tag = conColƷ��
        With Me.lvwItem.ColumnHeaders
            .Clear
            .Add , "����", "����", 900
            .Add , "ͨ������", "ͨ������", 2000
            .Add , "���", "���", 1200
            .Add , "����", "����", 1100
            .Add , "��λ", "��λ", 450
        End With
        
        gstrSQL = " Select Distinct D.ҩƷID,C.����,NVL(A.����,C.����) ͨ������,C.���,C.����,C.���㵥λ ��λ,Nvl(C.�Ƿ���,0) ���" & _
                      " From �շ���Ŀ���� A,ҩƷ��� D," & _
                      "     (Select B.* From �շ���Ŀ���� A,�շ���ĿĿ¼ B" & _
                      "     Where A.�շ�ϸĿID=B.ID And B.��� In ('5','6','7') " & _
                      "           ANd (A.���� Like [1] Or A.���� Like [1] Or B.���� Like [1])) C " & _
                      " WHERE D.ҩƷID=C.ID And (C.����ʱ�� Is Null Or C.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))" & _
                      " And D.ҩƷID=A.�շ�ϸĿID(+) and A.����(+)=3 and A.����(+)=1"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, StrFindStyle & strInput & "%")
                       
        With rsTemp
            If .EOF Then
                MsgBox "δ�ҵ����ҩƷ�����������룡", vbInformation, gstrSysName
                Cancel = True
                bfgPrice.TxtSetFocus
                Exit Sub
            End If
            
            '�����ʱ��ҩƷ�����ۼ�¼����������Ч
            If !��� = 1 And blnImmediately = False Then
                blnImmediately = True
                chkImmediately.Value = 1
                chkImmediately.Enabled = False
            End If
            If !��� = 1 Then Chk����.Enabled = True
            
            If chkImmediately.Value = 1 Then
                Dim RecCheck As New ADODB.Recordset
                
                gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And �շ�ϸĿID=[1] " & _
                        GetPriceClassString("")
                
                Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(rsTemp!ҩƷid))
                
                With RecCheck
                    If Not .EOF Then
                        If Not IsNull(!Records) Then
                            If !Records <> 0 And chkImmediately.Value = 1 Then
                                MsgBox "��ҩƷ����δִ�м۸񣬲�������Ϊ����ִ�У�", vbInformation, gstrSysName
                                If chkImmediately.Enabled Then
                                    chkImmediately.Value = 0
                                Else
                                    Cancel = True
                                    bfgPrice.TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End With
            End If
            
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ҩƷid, !����)
                objItem.SubItems(1) = !ͨ������
                objItem.SubItems(2) = IIf(IsNull(!���), "", !���)
                objItem.SubItems(3) = IIf(IsNull(!����), "", !����)
                objItem.SubItems(4) = IIf(IsNull(!��λ), "", !��λ)
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            If Me.lvwItem.ListItems.count = 0 Then
                MsgBox "��ҩƷ������", vbExclamation, gstrSysName
                Cancel = True
                Exit Sub
            ElseIf Me.lvwItem.ListItems.count = 1 Then
                lvwItem_DblClick
                Cancel = True
                Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = Me.bfgPrice.Left
            .Top = Me.bfgPrice.Top + Me.bfgPrice.CellTop + Me.bfgPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            .SetFocus
        End With
        Cancel = True
    Case conCol��������
        If Trim(Me.bfgPrice.Text) = "" Then Exit Sub
        strInput = UCase(Me.bfgPrice.Text)
        
        Me.lvwItem.Tag = conCol��������
        With Me.lvwItem.ColumnHeaders
            .Clear
            .Add , "����", "����", 600
            .Add , "����", "����", 1000
        End With
        
        gstrSQL = "select id,����,���� from ������Ŀ U" & _
                    " where rownum<100 and nvl(����ʱ��,to_date('3000-01-01','YYYY-MM-DD'))>trunc(sysdate)+1" & _
                    "      and NOT exists(select 1 from ������Ŀ D where D.�ϼ�id=U.id)" & _
                    "      and (���� like [1] or ���� like [1] or ���� like [1])"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, strInput & "%")
        
        With rsTemp
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !Id, !����)
                objItem.SubItems(1) = !����
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            If Me.lvwItem.ListItems.count = 0 Then
                MsgBox "����Ŀ������", vbExclamation, gstrSysName
                Cancel = True
                Exit Sub
            ElseIf Me.lvwItem.ListItems.count = 1 Then
                lvwItem_DblClick
                Cancel = True
                Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = bfgPrice.Left + bfgPrice.MsfObj.CellLeft
            .Top = Me.bfgPrice.Top + Me.bfgPrice.CellTop + Me.bfgPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 2000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 2000
            End If
            .Visible = True
            .SetFocus
        End With
        Cancel = True
    Case conCol�ּ�
        Dim lngҩƷID As Long
        With bfgPrice
            lngҩƷID = Val(bfgPrice.TextMatrix(bfgPrice.Row, conColҩƷid))
            If lngҩƷID = 0 Then Exit Sub
            
            '�ּ۲��ܴ���ָ�����ۼ�
            gstrSQL = " Select Nvl(ָ�����ۼ�,0) ָ�����ۼ� From ҩƷ��� Where ҩƷID=[1]"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngҩƷID)
            
            If Val(.Text) > rsTemp!ָ�����ۼ� Then
                MsgBox "�ּ۲��ܴ���ָ�����ۼۣ���" & Format(rsTemp!ָ�����ۼ�, "#####0.0000000;-#####0.0000000; ;") & "��"
                Cancel = True
                .TxtSetFocus
            End If
            BlnModify = True
        End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkImmediately_Click()
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer
    
    On Error GoTo errHandle
    If chkImmediately.Value = 1 Then
        'ѭ���ж�����ҩƷ
        For IntCheck = 1 To bfgPrice.rows - 1
            LngmediIDThis = Val(bfgPrice.TextMatrix(IntCheck, conColҩƷid))
            If LngmediIDThis <> 0 Then
                '�ж��Ƿ���δִ�е���ʷ�۸�
                
                 gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And �շ�ϸĿID=[1]" & _
                 GetPriceClassString("")
                 
                 Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
                 
                 With RecCheck
                    If Not .EOF Then
                        If Not IsNull(!Records) Then
                            If !Records <> 0 Then
                                MsgBox "ҩƷ" & bfgPrice.TextMatrix(IntCheck, conColƷ��) & "����δִ�м۸񣬲�������Ϊ����ִ�У�", vbInformation, gstrSysName
                                chkImmediately.Value = 0
                                Exit Sub
                            End If
                        End If
                    End If
                End With
            End If
        Next
    End If
    
    If Me.chkImmediately.Value Then
        Me.dtpRunDate.Enabled = False
    Else
        Me.dtpRunDate.Enabled = True
    End If
    
    On Error Resume Next
    Me.bfgPrice.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCanc_Click()
    lngBillId = 0
    lngMediId = 0
    Unload Me
End Sub



Private Sub cmdOK_Click()
'    Dim strID As String, LngCurID As Long
'    Dim ArrayID
'    Dim lngAdjId As Long
'    Dim strOldId As String
'    Dim strNewId As String
'
'    '����������Ϸ���
'    If CheckPrice = False Then Exit Sub
'    '�����ʱִ�У�����ù���zl_ҩƷ�շ���¼_Adjust
'
'    dtToday = Sys.Currentdate()
'    Err = 0
'    On Error GoTo ErrHand
'    With rsTemp
'        gstrSQL = "select �շѼ�Ŀ_ID.nextval from dual"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSQL)
'        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "cmdOK_Click")
'        Call SQLTest
'
'        lngAdjId = .Fields(0).Value
'    End With
'
'    gcnOracle.BeginTrans
'    With Me.bfgPrice
'        strOldId = ""
'        strNewId = ""
'        strID = ""
'        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
'            LngCurID = Sys.NextId("�շѼ�Ŀ")
'            strID = strID & IIf(strID = "", "", ",") & LngCurID
'            If CLng(.RowData(intCount)) <> 0 Then
'                If .RowData(intCount) <> -1 And InStr(1, strOldId & ",", "," & .RowData(intCount) & ",") > 0 Then
'                    gcnOracle.RollbackTrans: .SetFocus: Exit Sub
'                    MsgBox "��һ�ε����в��ܶ���ͬƷ��(" & .TextMatrix(intCount, conColƷ��) & ")�ظ�����", vbExclamation, gstrSysName
'                End If
'                If .RowData(intCount) = -1 And InStr(1, strNewId & ",", "," & .TextMatrix(intCount, conColҩƷid) & ",") > 0 Then
'                    gcnOracle.RollbackTrans: .SetFocus: Exit Sub
'                    MsgBox "���ܶ���ͬƷ��(" & .TextMatrix(intCount, conColƷ��) & ")�ظ����ü۸�", vbExclamation, gstrSysName
'                End If
''                If .TextMatrix(intCount, conColԭ��) = .TextMatrix(intCount, conCol�ּ�) Then
''                    MsgBox .TextMatrix(intCount, conColƷ��) & " �ּ�δ����������", vbExclamation, gstrSysName
''                    gcnOracle.RollbackTrans:.SetFocus:Exit Sub
''                End If
'                If .RowData(intCount) <> -1 Then
'                    strOldId = strOldId & "," & .RowData(intCount)
'                Else
'                    strNewId = strNewId & "," & .TextMatrix(intCount, conColҩƷid)
'                End If
'
'                If Val(.TextMatrix(intCount, conCol�ּ�)) <> 0 Then
'                    '������һ�εļ۸��¼��ִֹ��
'                    gstrSQL = "zl_�շѼ�Ŀ_stop(" & .TextMatrix(intCount, conColҩƷid) & ","
'                    If Me.chkImmediately.Value Then
'                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    Else
'                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    End If
'                    gstrSQL = gstrSQL & ")"
'                    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'                    '�����۸��¼
'                    gstrSQL = "zl_�շѼ�Ŀ_Insert(" & LngCurID & "," & IIf(.RowData(intCount) = -1, "NUll", .RowData(intCount)) & _
'                              "," & .TextMatrix(intCount, conColҩƷid) & "," & Val(.TextMatrix(intCount, conColOld����ID)) & "," & Val(.TextMatrix(intCount, conColԭ��)) & "," & Val(.TextMatrix(intCount, conCol�ּ�)) & _
'                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtRegistrar.Text) & "',"
'                    If Me.chkImmediately.Value Then
'                        gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    Else
'                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    End If
'                    gstrSQL = gstrSQL & ",0)"
'                    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
'                End If
'
'            End If
'        Next
'    End With
'
'    'ѭ��ִ�й���
'    ArrayID = Split(strID, ",")
'    For intCount = 0 To UBound(ArrayID)
'        If Me.chkImmediately.Value Or bfgPrice.RowData(intCount + 1) = -1 Then
'            gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & ArrayID(intCount) & "," & Chk����.Value & ")"
'            Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption & "-�����۸������¼")
'        End If
'    Next
'
'    gcnOracle.CommitTrans
'    lngBillId = 0
'    lngMediId = 0
'
'    BlnModify = False
'    Unload Me
    Exit Sub
    
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Me.bfgPrice.SetFocus
End Sub

Private Sub cmdCpt_Click()
    Dim lngMediId As Long
    Dim dblOldPrice As Double
    Dim dblNewPrice As Double
    
    Dim intRows As Integer
    
    Me.hgdStore.Redraw = False
    Me.hgdStore.rows = 1
    
    On Error GoTo errHandle
    With Me.bfgPrice
        For intCount = 1 To .rows - 1
            lngMediId = Val(.TextMatrix(intCount, conColҩƷid))
            dblOldPrice = Val(.TextMatrix(intCount, conColԭ��))
            dblNewPrice = Val(.TextMatrix(intCount, conCol�ּ�))
            If lngMediId <> 0 And dblOldPrice <> dblNewPrice Then
                gstrSQL = "SELECT DISTINCT D.���� AS �ⷿ,'['||C.����||']'||NVL(A.����,C.����) AS ҩƷ," & _
                    "      C.���,C.����,C.���㵥λ AS ��λ,S.����,S.����,S.����" & _
                    " FROM " & _
                    "      (SELECT S.�ⷿID,S.ҩƷID,S.�ϴ����� ����,S.ʵ������ AS ����,S.����" & _
                    "      FROM ҩƷ��� S" & _
                    "      WHERE S.ʵ������<>0 AND S.ҩƷID=[1] And S.����=1) S, " & _
                    "      ���ű� D,ҩƷ��� M,�շ���Ŀ���� A,�շ���ĿĿ¼ C" & _
                    " WHERE D.ID=S.�ⷿID AND M.ҩƷID=C.ID AND S.ҩƷID=M.ҩƷID     " & _
                    " AND M.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 AND A.����(+)=1" & _
                    " ORDER BY '['||C.����||']'||NVL(A.����,C.����),S.����"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)
                
                With rsTemp
                    intRows = Me.hgdStore.rows
                    Me.hgdStore.rows = Me.hgdStore.rows + .RecordCount
                    Do While Not .EOF
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 0) = !�ⷿ
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 1) = !ҩƷ
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 2) = IIf(IsNull(!���), "", !���) & IIf(IsNull(!����), "", "|" & !����)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 3) = IIf(IsNull(!��λ), "", !��λ)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 4) = IIf(IsNull(!����), "", !����)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 5) = Format(!����, "0.00000")
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 6) = zlStr.FormatEx(dblOldPrice, mconintPriceBit)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 7) = zlStr.FormatEx(dblNewPrice, mconintPriceBit)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 8) = Format(!���� * (dblNewPrice - dblOldPrice), "0.00")
                        .MoveNext
                    Loop
                End With
            End If
        Next
    End With
    
    If Me.hgdStore.rows < 2 Then
        Me.hgdStore.rows = 2
    End If
    Me.hgdStore.FixedRows = 1
    Me.hgdStore.Redraw = True
    If Me.bfgPrice.Active Then Me.bfgPrice.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdPrint_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Trim(Me.bfgPrice.TextMatrix(1, 0)) = "" Then Exit Sub
    objPrint.Title.Text = "ҩƷ����֪ͨ��"
    
    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(Me.chkImmediately.Value, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txtRegistrar.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.bfgPrice.MsfObj
    objPrint.PageFooter = 2
     
    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing

End Sub

Private Sub cmdPstor_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Me.cmdCpt.Enabled Then
        Call cmdCpt_Click
    End If
    If Trim(Me.hgdStore.TextMatrix(1, 0)) = "" Then Exit Sub
    
    objPrint.Title.Text = "���ۿ��䶯��"
    
    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(Me.chkImmediately.Value, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txtRegistrar.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(Sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.hgdStore
    objPrint.PageFooter = 2
     
    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing

End Sub

Private Sub dtpRunDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.cmdOk.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyEscape Then
        If Me.ActiveControl.Name = "tvwItem" Then
            tvwItem.Visible = False
            bfgPrice.SetFocus
        ElseIf Me.ActiveControl.Name = "lvwItem" Then
            lvwItem.Visible = False
            bfgPrice.SetFocus
        Else
            cmdCanc_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    BlnModify = False
    StrFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    
    On Error GoTo errHandle
    With Me.hgdStore
        .rows = 2
        .Cols = 9
        .Redraw = False
        .TextMatrix(0, 0) = "�ⷿ"
        .TextMatrix(0, 1) = "ҩƷ"
        .TextMatrix(0, 2) = "���|����"
        .TextMatrix(0, 3) = "��λ"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "����"
        .TextMatrix(0, 6) = "ԭ��"
        .TextMatrix(0, 7) = "�ּ�"
        .TextMatrix(0, 8) = "�������"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 2800
        .ColWidth(2) = 1350
        .ColWidth(3) = 400
        .ColWidth(4) = 800
        .ColWidth(5) = 1000
        .ColWidth(6) = 900
        .ColWidth(7) = 900
        .ColWidth(8) = 1050
    
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
    
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
        .ColAlignmentFixed(6) = 4
        .ColAlignmentFixed(7) = 4
        .ColAlignmentFixed(8) = 4
        .Redraw = True
    End With
    
    With Me.bfgPrice
        .Cols = 10
        .MsfObj.FixedCols = 0
        .TextMatrix(0, conColҩƷid) = "ҩƷid"
        .TextMatrix(0, conColƷ��) = "Ʒ��"
        .TextMatrix(0, conCol���) = "���"
        .TextMatrix(0, conCol����) = "����"
        .TextMatrix(0, conCol��λ) = "��λ"
        .TextMatrix(0, conColԭ��) = "ԭ��"
        .TextMatrix(0, conCol�ּ�) = "�ּ�"
        .TextMatrix(0, conCol����ID) = "����id"
        .TextMatrix(0, conColOld����ID) = "ԭ����id"
        .TextMatrix(0, conCol��������) = "������Ŀ"
        
        .ColWidth(conColҩƷid) = 0
        .ColWidth(conColƷ��) = 2800
        .ColWidth(conCol���) = 1200
        .ColWidth(conCol����) = 1000
        .ColWidth(conCol��λ) = 400
        .ColWidth(conColԭ��) = 975
        .ColWidth(conCol�ּ�) = 1000
        .ColWidth(conCol����ID) = 0
        .ColWidth(conColOld����ID) = 0
        .ColWidth(conCol��������) = 1000
        
        .ColData(conColҩƷid) = 5
        .ColData(conColƷ��) = 1
        .ColData(conCol���) = 5
        .ColData(conCol����) = 5
        .ColData(conCol��λ) = 5
        .ColData(conColԭ��) = 5
        .ColData(conCol�ּ�) = 4
        .ColData(conCol����ID) = 5
        .ColData(conColOld����ID) = 5
        .ColData(conCol��������) = 1

        .ColAlignment(conColҩƷid) = 1
        .ColAlignment(conColƷ��) = 1
        .ColAlignment(conCol���) = 1
        .ColAlignment(conCol����) = 1
        .ColAlignment(conCol��λ) = 4
        .ColAlignment(conColԭ��) = 7
        .ColAlignment(conCol�ּ�) = 7
        .ColAlignment(conCol����ID) = 1
        .ColAlignment(conColOld����ID) = 1
        .ColAlignment(conCol��������) = 1

        .PrimaryCol = conColƷ��
        .LocateCol = conColƷ��
    End With
    
    Dim StrToday As String
    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    If lngBillId = 0 Then
        '������۱༭״̬
        Me.bfgPrice.Active = True
        If blnImmediately Then
            Me.chkImmediately.Value = 1
            Me.chkImmediately.Enabled = False
        End If
        
        Me.lblTitle.Caption = "���䶯��(���ڵ���δ���棬��ӳ�Ŀ����ܲ�׼ȷ)"
        Me.dtpRunDate.MinDate = DateAdd("s", 1, Format(Sys.Currentdate, "yyyy-MM-dd"))
        Me.dtpRunDate.Value = DateAdd("d", 1, Format(Sys.Currentdate, "YYYY-MM-DD"))
        Me.txtRegistrar.Text = gstrUserName
        With rsTemp
            
            gstrSQL = "select id,�ϼ�id,����,���� from ������Ŀ start with �ϼ�id is null connect by  prior id=�ϼ�id order by level"
            If .State = adStateOpen Then .Close
            
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Form_Load")
            Call SQLTest
            
            Me.tvwItem.Nodes.Clear
            Do While Not .EOF
                If IsNull(!�ϼ�ID) Then
                    Me.tvwItem.Nodes.Add , , "_" & !Id, !���� & "_" & !����
                Else
                    Me.tvwItem.Nodes.Add "_" & !�ϼ�ID, 4, "_" & !Id, !���� & "_" & !����
                End If
                .MoveNext
            Loop
        End With
        
        If lngMediId = 0 Then Exit Sub
        '���ָ�����ȵ��۵�ҩƷ����ֱ�ӽ���ҩƷ����
            
         gstrSQL = " Select Distinct D.ҩƷID,'['||C.����||']'||NVL(A.����,C.����) Ʒ��,C.���,C.����,C.���㵥λ ��λ" & _
                  " From �շ���Ŀ���� A,ҩƷ��� D,�շ���ĿĿ¼ C " & _
                  " Where D.ҩƷID =[1] And D.ҩƷID=C.ID" & _
                  " And D.ҩƷID=A.�շ�ϸĿID(+) And A.����(+)=3 And A.����(+)=1"
         Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)
         
         With rsTemp
            If .RecordCount = 0 Then Exit Sub
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conColҩƷid) = !ҩƷid
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conColƷ��) = !Ʒ��
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol���) = IIf(IsNull(!���), "", !���)
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol����) = IIf(IsNull(!����), "", !����)
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol��λ) = IIf(IsNull(!��λ), "", !��λ)
            Call getMediPrice(lngMediId)
            Me.bfgPrice.Col = conCol�ּ�
        End With
    
    Else
        '���������ʾ״̬
        Me.bfgPrice.Active = False
        Me.cmdOk.Visible = False
        Me.cmdCanc.Caption = "����(&C)"
        Me.cmdCanc.Top = Me.cmdOk.Top
        Me.txtSummary.Enabled = False
        Me.chkImmediately.Value = 0
        Me.chkImmediately.Enabled = False
        Me.dtpRunDate.Enabled = False
        
        Dim strBills As String
        Dim strUnit As String
        
        strBills = ""
        
        Select Case intUnit
            Case 1
                strUnit = ",C.���㵥λ AS ��λ,P.ԭ��,P.�ּ�"
            Case 2
                strUnit = ",M.���ﵥλ AS ��λ,P.ԭ��*M.�����װ As ԭ��,P.�ּ�*M.�����װ As �ּ�"
            Case 3
                strUnit = ",M.ҩ�ⵥλ AS ��λ,P.ԭ��*M.ҩ���װ As ԭ��,P.�ּ�*M.ҩ���װ As �ּ�"
            Case 4
                strUnit = ",M.סԺ��λ AS ��λ,P.ԭ��*M.סԺ��װ As ԭ��,P.�ּ�*M.סԺ��װ As �ּ�"
        End Select
        
        gstrSQL = "SELECT DISTINCT P.ID,M.ҩƷID,'['||C.����||'] '||NVL(A.����,C.����) AS Ʒ��,C.���,C.���� " & strUnit & _
            "      ,P.������ĿID,I.���� AS ��������,TO_CHAR(P.ִ������,'YYYY-MM-DD HH24:MI:SS') ִ������,P.�䶯ԭ��,P.����˵��,P.������" & _
            " FROM �շѼ�Ŀ P,ҩƷ��� M,������Ŀ I,�շ���Ŀ���� A,�շ���ĿĿ¼ C" & _
            " WHERE P.�շ�ϸĿID=M.ҩƷID AND P.������ĿID=I.ID AND M.ҩƷID=C.ID " & _
            " AND M.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 AND A.����(+)=1" & _
            " AND P.ID=[1] " & _
            GetPriceClassString("P") & _
            " ORDER BY P.ID"                            '�����IDȡ���Ǽ۸��¼ID����һ��ID
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngBillId)
        
        Me.bfgPrice.rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            strBills = strBills & "," & rsTemp!Id
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conColҩƷid) = rsTemp!ҩƷid
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conColƷ��) = rsTemp!Ʒ��
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol��λ) = IIf(IsNull(rsTemp!��λ), "", rsTemp!��λ)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conColԭ��) = zlStr.FormatEx(rsTemp!ԭ��, mconintPriceBit)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol�ּ�) = zlStr.FormatEx(rsTemp!�ּ�, mconintPriceBit)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol����ID) = rsTemp!������ĿID
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol��������) = rsTemp!��������
            Me.txtSummary = IIf(IsNull(rsTemp!����˵��), "", rsTemp!����˵��)
            Me.txtRegistrar.Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
            Me.dtpRunDate.Value = rsTemp!ִ������
            
            If rsTemp!ִ������ <= StrToday And rsTemp!�䶯ԭ�� = 0 Then        'δ���е��ۼ���,��ִ�м���
                gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & Val(rsTemp!ҩƷid) & ")"
                Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption & "-�����۸������¼")
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        
        If rsTemp!ִ������ > StrToday Then
            '���ִ��ʱ��δ������ֻ��ģ����ʾ���䶯
            Me.lblTitle.Caption = "���䶯��(����ִ��ʱ��δ������ӳ�Ŀ����ܲ�׼ȷ)"
            Call cmdCpt_Click
        Else
            'ִ��ʱ���ѵ����϶�Ҳ�����˵��ۼ��㣬ֱ�Ӵ��շ���¼��ȡ���۱䶯���
            Me.cmdCpt.Enabled = False
            Me.lblTitle.Caption = "���䶯��"

            Select Case intUnit
                Case 1
                    strUnit = ",C.���㵥λ AS ��λ,S.ԭ��,S.�ּ�,S.����"
                Case 2
                    strUnit = ",M.���ﵥλ AS ��λ,S.ԭ��*M.�����װ As ԭ��,S.�ּ�*M.�����װ As �ּ�,S.����/M.�����װ As ����"
                Case 3
                    strUnit = ",M.ҩ�ⵥλ AS ��λ,S.ԭ��*M.ҩ���װ As ԭ��,S.�ּ�*M.ҩ���װ As �ּ�,S.����/M.ҩ���װ As ����"
                Case 4
                    strUnit = ",M.סԺ��λ AS ��λ,S.ԭ��*M.סԺ��װ As ԭ��,S.�ּ�*M.סԺ��װ As �ּ�,S.����/M.סԺ��װ As ����"
            End Select

            gstrSQL = "SELECT DISTINCT S.ID,D.���� AS �ⷿ,'['||C.����||']'||NVL(A.����,C.����) AS ҩƷ,C.���,C.����,S.����,S.�������" & strUnit & _
                " FROM (SELECT ID,�ⷿID,ҩƷID,����,��д���� AS ����,�ɱ��� AS ԭ��,���ۼ� AS �ּ�,���۽�� AS �������" & _
                "       FROM " & _
                "       (SELECT P.ID,N.�ⷿID,N.ҩƷID,N.����,N.��д����,N.�ɱ���,N.���ۼ�,N.���۽��" & _
                "       FROM ҩƷ�շ���¼ N," & _
                "           (SELECT ID,�շ�ϸĿID,ִ������,��ֹ���� FROM �շѼ�Ŀ" & _
                "           WHERE ID=[1]" & GetPriceClassString("") & ") P" & _
                "       WHERE N.ҩƷID=P.�շ�ϸĿID AND ����=13" & _
                "       AND N.������� BETWEEN P.ִ������ AND NVL(��ֹ����,SYSDATE))) S," & _
                "       ���ű� D,ҩƷ��� M,�շ���Ŀ���� A,�շ���ĿĿ¼ C" & _
                " WHERE S.�ⷿID=D.ID AND S.ҩƷID=M.ҩƷID AND M.ҩƷID=C.ID " & _
                " AND M.ҩƷID=A.�շ�ϸĿID(+) AND A.����(+)=3 AND A.����(+)=1" & _
                " ORDER BY '['||C.����||']'||NVL(A.����,C.����),S.����"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(strBills, 2))
            
            If rsTemp.RecordCount > 0 Then Me.hgdStore.rows = rsTemp.RecordCount + 1
            Do While Not rsTemp.EOF
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 0) = rsTemp!�ⷿ
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 1) = rsTemp!ҩƷ
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 2) = IIf(IsNull(rsTemp!���), "", rsTemp!���) & IIf(IsNull(rsTemp!����), "", "|" & rsTemp!����)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 3) = IIf(IsNull(rsTemp!��λ), "", rsTemp!��λ)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 4) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 5) = Format(rsTemp!����, "0.00000")
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 6) = zlStr.FormatEx(rsTemp!ԭ��, mconintPriceBit)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 7) = zlStr.FormatEx(rsTemp!�ּ�, mconintPriceBit)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 8) = Format(rsTemp!�������, "0.00")
                rsTemp.MoveNext
            Loop
           
        End If
            
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height < 5000 Then
        Me.Height = 5000
    End If
    If Me.Width < 9720 Then
        Me.Width = 9720
    End If
    
    Me.cmdOk.Left = Me.ScaleWidth - Me.cmdOk.Width - 150
    Me.cmdCanc.Left = Me.cmdOk.Left
    Me.cmdPrint.Left = Me.cmdOk.Left
    
    
    Me.bfgPrice.Width = Me.cmdOk.Left - 150
    Me.txtRegistrar.Left = Me.bfgPrice.Left + Me.bfgPrice.Width - Me.txtRegistrar.Width
    Me.lblRegistrar.Left = txtRegistrar.Left - lblRegistrar.Width - 50
    Me.txtSummary.Width = lblRegistrar.Left - txtSummary.Left - 300
    
    Me.dtpRunDate.Left = Me.bfgPrice.Left + Me.bfgPrice.Width - Me.dtpRunDate.Width
    Me.lblRunDate.Left = dtpRunDate.Left - lblRunDate.Width - 50
    
    Me.cmdPstor.Left = Me.cmdOk.Left + Me.cmdOk.Width - Me.cmdPstor.Width
    Me.cmdCpt.Left = Me.cmdPstor.Left - 45 - Me.cmdCpt.Width
    
    Me.hgdStore.Width = Me.ScaleWidth - 30
    Me.hgdStore.Height = Me.ScaleHeight - 30 - Me.hgdStore.Top

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If BlnModify Then If MsgBox("��ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1: Exit Sub
    SaveWinState Me
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwItem
        .Sorted = False
        .SortKey = ColumnHeader.Index - 2
        .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwItem_DblClick()
    Dim LngmediIDThis As Long
    Dim RecCheck As New ADODB.Recordset
    
    On Error GoTo errHandle
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    Set objItem = Me.lvwItem.SelectedItem
    LngmediIDThis = Mid(objItem.Key, 2)
    If LngmediIDThis = 0 Then Exit Sub
    If Me.lvwItem.Tag = conColƷ�� Then
        
        '�Ǳ��ҩƷ���˳�

       gstrSQL = " Select Nvl(�Ƿ���,0) ��� From �շ���ĿĿ¼ Where ҩƷID=[1]"
       Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
       
       '�����ʱ��ҩƷ�����ۼ�¼����������Ч
       If RecCheck!��� = 1 And blnImmediately = False Then
           blnImmediately = True
           chkImmediately.Value = 1
           chkImmediately.Enabled = False
       End If
        
        If chkImmediately.Value = 1 Then
            '�ж��Ƿ���δִ�е���ʷ�۸�
            
            gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And �շ�ϸĿID=[1] " & _
                    GetPriceClassString("")
            
            Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
            
            If Not RecCheck.EOF Then
                If Not IsNull(RecCheck!Records) Then
                    If RecCheck!Records <> 0 And chkImmediately.Value = 1 Then
                        MsgBox "��ҩƷ����δִ�м۸񣬲�������Ϊ����ִ�У�", vbInformation, gstrSysName
                        If chkImmediately.Enabled Then
                            chkImmediately.Value = 0
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
        With Me.bfgPrice
            .TextMatrix(.Row, conColҩƷid) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, conColƷ��) = "[" & objItem.Text & "] " & objItem.SubItems(1)
            .TextMatrix(.Row, conCol���) = objItem.SubItems(2)
            .TextMatrix(.Row, conCol����) = objItem.SubItems(3)
            .TextMatrix(.Row, conCol��λ) = objItem.SubItems(4)
            Call getMediPrice(.TextMatrix(.Row, conColҩƷid))
            .CmdVisible = False
            .Col = conCol�ּ�
        End With
    Else
        With Me.bfgPrice
            .TextMatrix(.Row, conCol����ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, conCol��������) = objItem.SubItems(1)
            .CmdVisible = False
            .Col = conCol��������
        End With
    End If
    Me.lvwItem.Visible = False
    bfgPrice.SetFocus
    BlnModify = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    lvwItem_DblClick
End Sub

Private Sub lvwItem_LostFocus()
    Me.lvwItem.Visible = False
End Sub

Private Sub tvwItem_DblClick()
    If Me.tvwItem.SelectedItem Is Nothing Then Exit Sub
    If Me.tvwItem.SelectedItem.Selected = False Then Exit Sub
    If Me.tvwItem.SelectedItem.Children = 0 Then
        tvwItem_NodeClick Me.tvwItem.SelectedItem
        Me.tvwItem.Visible = False
    End If
    bfgPrice.SetFocus
End Sub

Private Sub tvwItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    tvwItem_DblClick
End Sub

Private Sub tvwItem_LostFocus()
    Me.tvwItem.Visible = False
End Sub

Private Sub tvwItem_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Children = 0 Then
        BlnModify = True
        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol����ID) = Mid(Node.Key, 2)
        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol��������) = Mid(Node.Text, InStr(1, Node.Text, "_") + 1)
    End If
End Sub

Private Function getMediPrice(lngMediId As Long)
    Dim blnʱ�� As Boolean
    Dim rsʱ�� As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(�Ƿ���,0) ��� From �շ���ĿĿ¼ Where ID=[1]"
    Set rsʱ�� = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)

    blnʱ�� = (rsʱ��!��� = 1)
    If blnʱ�� Then Chk����.Enabled = True

'    If blnʱ�� Then
'        '��ʾʱ��ҩƷ���ۣ�ȡ�����/���������Ϊ��۸�
'        gstrSQL = "" & _
'            " SELECT P.ID,DECODE(K.�������,0,P.�ּ�,K.�����/NVL(K.�������,1)) �ּ�,P.ִ������,P.������ĿID,I.���� AS ��������" & _
'            " FROM �շѼ�Ŀ P,������Ŀ I," & _
'            "   (SELECT ҩƷID,SUM(ʵ�ʽ��) �����,SUM(ʵ������) �������" & _
'            "    FROM ҩƷ��� WHERE ����=1 And ҩƷID=" & lngMediId & _
'            "    GROUP BY ҩƷID ) K" & _
'            " WHERE P.������ĿID=I.ID AND P.�շ�ϸĿID=K.ҩƷID(+) AND P.�շ�ϸĿID=" & lngMediId & _
'            "       AND NVL(P.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')"
'    Else
'        '��ʱ��ҩƷ���ۣ�ȡ��۸��¼�еļ۸�
'        gstrSQL = "" & _
'            " SELECT P.ID,P.�ּ�,P.ִ������,P.������ĿID,I.���� AS ��������" & _
'            " FROM �շѼ�Ŀ P,������Ŀ I" & _
'            " WHERE P.������ĿID=I.ID AND P.�շ�ϸĿID=" & lngMediId & _
'            "       AND NVL(P.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')"
'    End If
'    If .State = adStateOpen Then .Close
'
    
    If blnʱ�� Then
        '��ʾʱ��ҩƷ���ۣ�ȡ�����/���������Ϊ��۸�
        gstrSQL = "" & _
            " SELECT P.ID,DECODE(K.�������,0,P.�ּ�,K.�����/NVL(K.�������,1)) �ּ�,P.ִ������,P.������ĿID,I.���� AS ��������" & _
            " FROM �շѼ�Ŀ P,������Ŀ I," & _
            "   (SELECT ҩƷID,SUM(ʵ�ʽ��) �����,SUM(ʵ������) �������" & _
            "    FROM ҩƷ��� WHERE ����=1 And ҩƷID=[1] " & _
            "    GROUP BY ҩƷID ) K" & _
            " WHERE P.������ĿID=I.ID AND P.�շ�ϸĿID=K.ҩƷID(+) AND P.�շ�ϸĿID=[1] " & _
            GetPriceClassString("P") & _
            "       AND NVL(P.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')"
    Else
        '��ʱ��ҩƷ���ۣ�ȡ��۸��¼�еļ۸�
        gstrSQL = "" & _
            " SELECT P.ID,P.�ּ�,P.ִ������,P.������ĿID,I.���� AS ��������" & _
            " FROM �շѼ�Ŀ P,������Ŀ I" & _
            " WHERE P.������ĿID=I.ID AND P.�շ�ϸĿID=[1] " & _
            "       AND NVL(P.��ֹ����,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')" & _
            GetPriceClassString("P")
    End If
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.bfgPrice.RowData(Me.bfgPrice.Row) = !Id
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColԭ��) = zlStr.FormatEx(!�ּ�, mconintPriceBit)
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol�ּ�) = zlStr.FormatEx(!�ּ�, mconintPriceBit)
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol����ID) = !������ĿID
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColOld����ID) = !������ĿID
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol��������) = !��������
            If Me.dtpRunDate.MinDate <= DateAdd("d", 1, !ִ������) Then
                Me.dtpRunDate.MinDate = DateAdd("d", 1, CDate(Format(!ִ������, "yyyy-MM-dd 00:00:00")))
            End If
            If Me.dtpRunDate.Value <= Me.dtpRunDate.MinDate Then
                Me.dtpRunDate.Value = Me.dtpRunDate.MinDate
            End If
        Else
            Me.bfgPrice.RowData(Me.bfgPrice.Row) = -1
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColԭ��) = Format(0, "0.0000000")
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol�ּ�) = Format(0, "0.0000000")
            If Me.bfgPrice.Row > 1 Then
                Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol����ID) = Me.bfgPrice.TextMatrix(Me.bfgPrice.Row - 1, conCol����ID)
                Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColOld����ID) = Me.bfgPrice.TextMatrix(Me.bfgPrice.Row - 1, conCol����ID)
                Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol��������) = Me.bfgPrice.TextMatrix(Me.bfgPrice.Row - 1, conCol��������)
            Else
                For Each objNode In Me.tvwItem.Nodes
                    If objNode.Children = 0 Then
                        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol����ID) = Mid(objNode.Key, 2)
                        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColOld����ID) = Mid(objNode.Key, 2)
                        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol��������) = objNode.Text
                    End If
                Next
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.dtpRunDate.Enabled Then Me.dtpRunDate.SetFocus
End Sub

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim RecCheck As New ADODB.Recordset
    '����ִ�м۸��Ƿ���ȷ
    '�Լ�������Ŀ��ͬ��������ּ��Ƿ���ԭ����ͬ
    CheckPrice = False
    With bfgPrice
        For IntCheck = 1 To .rows - 1
            If Val(.TextMatrix(IntCheck, conColҩƷid)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, conCol�ּ�))) Then
                    MsgBox "��" & IntCheck & "�е�ҩƷ�ּ��к��зǷ��ַ���", vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(.TextMatrix(IntCheck, conCol�ּ�)) = 0 Then
                    MsgBox "��" & IntCheck & "�е�ҩƷ�ּ۲���Ϊ�գ�", vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(.TextMatrix(IntCheck, conColOld����ID)) = Val(.TextMatrix(IntCheck, conCol����ID)) Then
                    If Val(.TextMatrix(IntCheck, conCol�ּ�)) = Val(.TextMatrix(IntCheck, conColԭ��)) Then
                        MsgBox "��" & IntCheck & "�е�ҩƷ�ּ���ԭ����ͬ������ִ�е��ۣ�"
                        Exit Function
                    End If
                End If
            End If
            
        Next
    End With
    CheckPrice = True
End Function
