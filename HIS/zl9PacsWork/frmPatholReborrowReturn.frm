VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholReborrowReturn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ĺ黹"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   Icon            =   "frmPatholReborrowReturn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtReturnedMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtMemo 
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox txtEnregMan 
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4440
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox rtfAdvice 
      Height          =   855
      Left            =   4320
      TabIndex        =   16
      Top             =   3960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      TextRTF         =   $"frmPatholReborrowReturn.frx":000C
   End
   Begin VB.TextBox txtDoctor 
      Height          =   300
      Left            =   7920
      TabIndex        =   13
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtHospital 
      Height          =   300
      Left            =   4320
      TabIndex        =   11
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox txtReturnMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtBorrowMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtHoldMoney 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   5
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpReturnDate 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   64684035
      CurrentDate     =   40903
   End
   Begin VB.TextBox txtReturnPepole 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin zl9PACSWork.ucFlexGrid ufgBackDetail 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   10215
      _ExtentX        =   16325
      _ExtentY        =   5106
      DefaultCols     =   ""
      GridRows        =   201
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      DataFontCharset =   134
      DataFontWeight  =   400
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10200
      TabIndex        =   28
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label14 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10200
      TabIndex        =   27
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6600
      TabIndex        =   26
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "����Ѻ��"
      Height          =   255
      Left            =   6600
      TabIndex        =   24
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "��    ע��"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "�� �� �ˣ�"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "���������"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "����ҽʦ��"
      Height          =   255
      Left            =   6960
      TabIndex        =   14
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "����ҽԺ��"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "����Ѻ��"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "����Ѻ��"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "�۽�Ѻ��"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "�黹���ڣ�"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "�� �� �ˣ�"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmPatholReborrowReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentBorrowGrid As ucFlexGrid
Private mlngBorrowId As Long
Public blnIsOk As Boolean


Public Sub ShowBorrowReturnWindow(ufgBorrow As ucFlexGrid, owner As Object)
    If Not ufgBorrow.IsSelectionRow Then
        Call err.Raise(0, "ShowBorrowReturnWindow", "û��ѡ����Ҫ�黹�Ľ��ļ�¼��")
        Exit Sub
    End If
    
    Set mufgParentBorrowGrid = ufgBorrow
    
    mlngBorrowId = Val(ufgBorrow.KeyValue(ufgBorrow.SelectionRow))
    
    blnIsOk = False
    
    Call ReadBorrowInfToFace
    Call LoadReturnMaterialData(mlngBorrowId)
    
    Call Me.Show(1, owner)
End Sub

Private Function GetReturnMoney(ByVal lngBorrowId As Long)
'ȡ������Ѻ��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetReturnMoney = 0
    
    strSQL = "select sum(�˻�Ѻ��) as ����ֵ from ����黹��Ϣ where ����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBorrowId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetReturnMoney = Val(Nvl(rsData!����ֵ))
End Function


Private Function CheckDataIsValid() As String
'�жϹ黹�������Ƿ���Ч
    Dim i As Long

    CheckDataIsValid = ""
    
    For i = 1 To ufgBackDetail.GridRows - 1
        If ufgBackDetail.GetRowCheck(i) Then
            If Val(ufgBackDetail.Text(i, gstrPatholCol_��������)) < Val(ufgBackDetail.Text(i, gstrPatholCol_ʵ������)) Then
                CheckDataIsValid = "�ڲ��Ϲ滮�б��У�ʵ�����������ܴ��ڴ�������������"
                
                Call ufgBackDetail.SetFocus
                Call ufgBackDetail.LocateRow(i)
                
                Exit Function
            End If
        End If
    Next i
    
End Function


Private Sub ReadBorrowInfToFace()
'��ȡ��ؽ�����Ϣ������
    txtReturnPepole.Text = mufgParentBorrowGrid.Text(mufgParentBorrowGrid.SelectionRow, gstrPatholCol_������)
    
    txtBorrowMoney.Text = mufgParentBorrowGrid.Text(mufgParentBorrowGrid.SelectionRow, gstrPatholCol_Ѻ��)   '����Ѻ��
    txtReturnedMoney.Text = GetReturnMoney(mlngBorrowId)    '����Ѻ��
    txtHoldMoney.Text = 0   '����Ѻ��
    txtReturnMoney.Text = Val(txtBorrowMoney.Text) - Val(txtReturnedMoney.Text) '����Ѻ��
    
    txtEnregMan.Text = UserInfo.����
    
    dtpReturnDate.value = zlDatabase.Currentdate

End Sub

Private Sub cmdCancel_Click()
'ȡ���黹
On Error Resume Next
    blnIsOk = False
    
    Call Me.Hide
err.Clear
End Sub

Private Sub cmdSure_Click()
On Error GoTo ErrHandle
    Dim strInf As String
    
    If Trim(txtReturnPepole.Text) = "" Then
        Call MsgBoxD(Me, "�黹�˲���Ϊ�ա�", vbOKOnly, Me.Caption)
        Call txtReturnPepole.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtHospital.Text) = "" Then
        Call MsgBoxD(Me, "����ҽԺ����Ϊ�ա�", vbOKOnly, Me.Caption)
        Call txtHospital.SetFocus
        
        Exit Sub
    End If
    
    If Trim(txtDoctor.Text) = "" Then
        Call MsgBoxD(Me, "����ҽʦ����Ϊ�ա�", vbOKOnly, Me.Caption)
        Call txtDoctor.SetFocus
        
        Exit Sub
    End If
    
    If Trim(rtfAdvice.Text) = "" Then
        Call MsgBoxD(Me, "�����������Ϊ�ա�", vbOKOnly, Me.Caption)
        Call rtfAdvice.SetFocus
        
        Exit Sub
    End If
    
    If Not ufgBackDetail.IsCheckedRow Then
        If MsgBoxD(Me, "ȷ�ϲ�ѡ���κβ��Ͻ��й黹�����𣿶�δ��ѡ�Ĳ��ϣ�ϵͳ���Զ�����ʧ����", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    End If
    
    strInf = CheckDataIsValid()
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call MaterialReturnProcess
    
    '���½��Ĺ黹״̬
    Call UpdateBorrowReturnState
    
    blnIsOk = True
    
    Call Me.Hide
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub UpdateBorrowReturnState()
'���½��Ĺ黹״̬
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnFind As Boolean
    Dim chkState As CheckState
    Dim strValue As String
    
    strSQL = "select �黹״̬ from ���������Ϣ where id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngBorrowId)
    
    If rsData.RecordCount > 0 Then
'        Call mufgParentBorrowGrid.GetFieldDisplayText(gstrPatholCol_�黹״̬, Val(Nvl(rsData!�黹״̬)), blnFind, chkState, strValue)
        Call mufgParentBorrowGrid.SyncData(mufgParentBorrowGrid.SelectionRow, gstrPatholCol_�黹״̬, Val(Nvl(rsData!�黹״̬)), True)
    End If
End Sub

Private Sub MaterialReturnProcess()
'���Ϲ黹����
    Dim i As Long
    Dim lngReturnCount As Long
    
    gcnOracle.BeginTrans
    
    On Error GoTo errTrans
    Call zlDatabase.ExecuteProcedure("ZL_����黹_������¼(" & _
                                        mlngBorrowId & ",'" & _
                                        txtReturnPepole.Text & "'," & _
                                        zlStr.To_Date(dtpReturnDate.value) & "," & _
                                        Val(txtReturnMoney.Text) & ",'" & _
                                        txtHospital.Text & "','" & _
                                        txtDoctor.Text & "','" & _
                                        rtfAdvice.Text & "','" & _
                                        txtEnregMan.Text & "','" & _
                                        txtMemo.Text & "')", Me.Caption)
                                        
    For i = 1 To ufgBackDetail.GridRows - 1
        lngReturnCount = IIf(ufgBackDetail.GetRowCheck(i), Val(ufgBackDetail.Text(i, gstrPatholCol_ʵ������)), 0)
        
        Call zlDatabase.ExecuteProcedure("ZL_����黹_���Ϲ黹(" & _
                                            mlngBorrowId & "," & _
                                            ufgBackDetail.KeyValue(i) & "," & _
                                            lngReturnCount & "," & _
                                            zlStr.To_Date(dtpReturnDate.value) & ",'" & _
                                            txtEnregMan.Text & "')", Me.Caption)
        
    Next i
    
    gcnOracle.CommitTrans
Exit Sub
errTrans:
    gcnOracle.RollbackTrans
End Sub


Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitBorrowReturnList
    
    blnIsOk = False
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitBorrowReturnList()
    '��������
    ufgBackDetail.GridRows = glngStandardRowCount
    '�����и�
    ufgBackDetail.RowHeightMin = glngStandardRowHeight
    
    '��ʼ���黹�б�
    ufgBackDetail.IsKeepRows = False
    ufgBackDetail.DefaultColNames = gstrMaterialBorrowReturnCols
    ufgBackDetail.ColNames = gstrMaterialBorrowReturnCols
    ufgBackDetail.ColConvertFormat = gstrMaterialBorrowReturnConvertFormat
End Sub


Private Sub LoadReturnMaterialData(ByVal lngBorrowId As Long)
'��ȡ��Ҫ�黹�Ĳ�����Ϣ
    Dim strSQL As String
    
    strSQL = "select a.�鵵id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, '����' as �������, " & _
             " case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, " & _
             " (nvl(a.��������, 0) - nvl(a.�黹����, 0)) as ��������, (nvl(a.��������, 0) - nvl(a.�黹����, 0)) as ʵ������, a.�黹״̬, e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
             " from ��������Ϣ d, ����ȡ����Ϣ c, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a " & _
             " Where c.����ҽ��id = d.����ҽ��id And b.�Ŀ�id = c.�Ŀ�id and e.id=b.����ID And a.�鵵id = b.ID And b.������Դ = 1 And a.�黹״̬<>1 and a.����id = [1] " & _
         " Union All " & _
             " select a.�鵵id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, '��Ƭ' as �������, " & _
             " decode(o.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, " & _
             " (nvl(a.��������, 0) - nvl(a.�黹����, 0)) as ��������,(nvl(a.��������, 0) - nvl(a.�黹����, 0)) as ʵ������,a.�黹״̬, e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
             " from ��������Ϣ d, ����ȡ����Ϣ c, ������Ƭ��Ϣ o, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a " & _
             " Where c.����ҽ��id = d.����ҽ��id And o.����ҽ��id = c.����ҽ��id " & _
             " and b.��Ƭid = o.id and c.�Ŀ�id= o.�Ŀ�id and e.id=b.����ID and a.�鵵id=b.id and b.������Դ=2 and a.�黹״̬<>1 and a.����id=[1] " & _
         " Union All " & _
             " select a.�鵵id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, " & _
             " decode(o.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
             " decode(o.�ؼ�ϸĿ,0,decode(o.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || q.�������� || decode(o.��������,-1,'-��',0,'','-��' || o.��������) || ')' as ������ϸ, " & _
             " (nvl(a.��������, 0) - nvl(a.�黹����, 0)) as ��������,(nvl(a.��������, 0) - nvl(a.�黹����, 0)) as ʵ������,a.�黹״̬, e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
             " from ��������Ϣ d, ����ȡ����Ϣ c, ��������Ϣ q, �����ؼ���Ϣ o, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a " & _
             " Where c.����ҽ��id = d.����ҽ��id And q.����ID = o.����ID And o.����ҽ��id = c.����ҽ��id " & _
             " and b.�ؼ�id = o.id and e.id=b.����ID and a.�鵵id=b.id and b.������Դ=3 and a.�黹״̬<>1 and a.����id=[1] "
             
    Set ufgBackDetail.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngBorrowId)
    Call ufgBackDetail.RefreshData
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    
err.Clear
End Sub

Private Sub txtHoldMoney_Change()
On Error Resume Next
    txtReturnMoney.Text = Val(txtBorrowMoney.Text) - Val(txtReturnedMoney.Text) - Val(txtHoldMoney.Text) '����Ѻ��
err.Clear
End Sub


Private Sub ufgBackDetail_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    If Col <> ufgBackDetail.GetColIndex(gstrPatholCol_ʵ������) Then Exit Sub
    
    If Val(ufgBackDetail.Text(Row, gstrPatholCol_ʵ������)) < Val(ufgBackDetail.Text(Row, gstrPatholCol_��������)) And _
        Val(ufgBackDetail.Text(Row, gstrPatholCol_ʵ������)) > 0 Then
        Call ufgBackDetail.SetRowCheck(Row, True)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

