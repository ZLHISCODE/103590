VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#2.1#0"; "zl9PacsControl.ocx"
Begin VB.Form frmPatholReborrowNew 
   Caption         =   "���ĵǼ�"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14055
   Icon            =   "frmPatholReborrowNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14055
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture0 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   120
      ScaleHeight     =   7935
      ScaleWidth      =   9735
      TabIndex        =   25
      Top             =   120
      Width           =   9735
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   100
         Left            =   0
         TabIndex        =   38
         Top             =   3975
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   185
         MousePointer    =   7
         SplitWidth      =   100
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   3000
         Con2MinSize     =   2000
         Control1Name    =   "Picture1"
         Control2Name    =   "Picture2"
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3860
         Left            =   0
         ScaleHeight     =   3855
         ScaleWidth      =   9735
         TabIndex        =   35
         Top             =   4075
         Width           =   9735
         Begin VB.CommandButton cmdCancelLend 
            Caption         =   "�������(&R)"
            Height          =   400
            Left            =   8160
            TabIndex        =   36
            Top             =   2880
            Width           =   1215
         End
         Begin zl9PACSWork.ucFlexGrid ufgMaterialEnreged 
            Height          =   2775
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4895
            DefaultCols     =   ""
            KeyName         =   "��"
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   0
         ScaleHeight     =   3975
         ScaleWidth      =   9735
         TabIndex        =   26
         Top             =   0
         Width           =   9735
         Begin VB.Frame framNameQuery 
            Height          =   615
            Left            =   2280
            TabIndex        =   44
            Top             =   0
            Width           =   6135
            Begin VB.TextBox txtPatholName 
               Height          =   300
               Left            =   720
               TabIndex        =   45
               Top             =   200
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker dtpStartDate 
               Height          =   300
               Left            =   3120
               TabIndex        =   46
               Top             =   200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   57475075
               CurrentDate     =   40899
            End
            Begin MSComCtl2.DTPicker dtpEndDate 
               Height          =   300
               Left            =   4680
               TabIndex        =   47
               Top             =   200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   57475075
               CurrentDate     =   40899
            End
            Begin VB.Label Label13 
               Caption         =   "�� ����"
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label14 
               Caption         =   "�������ڣ�"
               Height          =   255
               Left            =   2280
               TabIndex        =   49
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "��"
               Height          =   255
               Left            =   4470
               TabIndex        =   48
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.CommandButton cmdLend 
            Caption         =   "�������(&L)"
            Height          =   400
            Left            =   8160
            TabIndex        =   32
            Top             =   3480
            Width           =   1215
         End
         Begin VB.CheckBox chkMaterial 
            Caption         =   "�ؼ����"
            Height          =   180
            Index           =   2
            Left            =   2160
            TabIndex        =   31
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CheckBox chkMaterial 
            Caption         =   "��Ƭ����"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   30
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CheckBox chkMaterial 
            Caption         =   "�������"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton cmdQuery 
            Caption         =   "���ϲ�ѯ(&Q)"
            Height          =   400
            Left            =   8520
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtPatholNum 
            Height          =   330
            Left            =   720
            TabIndex        =   27
            Top             =   200
            Width           =   1455
         End
         Begin zl9PACSWork.ucFlexGrid ufgMaterialEnreg 
            Height          =   2655
            Left            =   0
            TabIndex        =   33
            Top             =   720
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4683
            DefaultCols     =   ""
            GridRows        =   201
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Label Label12 
            Caption         =   "����ţ�"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8175
      Left            =   10080
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin RichTextLib.RichTextBox rtfReason 
         Height          =   1695
         Left            =   1080
         TabIndex        =   43
         Top             =   4080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2990
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmPatholReborrowNew.frx":000C
      End
      Begin VB.CheckBox chkBorrowType 
         Caption         =   "�ڲ�����"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ ��(&C)"
         Height          =   400
         Left            =   2520
         TabIndex        =   23
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "ȷ ��(&S)"
         Height          =   400
         Left            =   1320
         TabIndex        =   22
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox txtMemo 
         Height          =   300
         Left            =   1080
         TabIndex        =   21
         Top             =   6000
         Width           =   2415
      End
      Begin VB.TextBox txtEnregPeople 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   6480
         Width           =   2415
      End
      Begin VB.TextBox txtAddress 
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox txtMobilePhone 
         Height          =   300
         Left            =   1080
         TabIndex        =   14
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtBorrowDays 
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Text            =   "30"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtMoney 
         Height          =   300
         Left            =   1080
         TabIndex        =   10
         Text            =   "0"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtCardNum 
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cbxCardType 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpBorrowDate 
         Height          =   300
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   57475075
         CurrentDate     =   40898
      End
      Begin VB.TextBox txtBorrowPeople 
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   42
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label16 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   41
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label17 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   40
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label16 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   39
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "��ע˵����"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   6060
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "�� �� �ˣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   6540
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "����ԭ��"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4140
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "��ϵ��ַ��"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "��ϵ�绰��"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "����������"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "����Ѻ��"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "֤�����룺"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "֤�����ͣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "�������ڣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "�� �� �ˣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholReborrowNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugState = False


Private mufgParentBorrowGrid As ucFlexGrid

Private mlngBorrowId As Long
Private mblnIsUpdate As Boolean
Private mblnIsEnter As Boolean

Public blnIsOk As Boolean



Public Sub ShowNewBorrowWindow(ufgParentGrid As ucFlexGrid, owner As Object)
'��ʾ�������Ĵ���
    mblnIsUpdate = False
    blnIsOk = False
    mlngBorrowId = -1
    
    Set mufgParentBorrowGrid = ufgParentGrid
    
    txtEnregPeople.Text = UserInfo.����
    
    Call Me.Show(1, owner)
End Sub


Public Sub ShowUpdateBorrowWindow(ufgParentGrid As ucFlexGrid, owner As Object)
'��ʾ�������Ĵ���

    
    mblnIsUpdate = True
    blnIsOk = False
    mlngBorrowId = ufgParentGrid.KeyValue(ufgParentGrid.SelectionRow)
    
    Set mufgParentBorrowGrid = ufgParentGrid
    
    Call ConfigUpdateData
    Call LoadBorrowMaterialDetail
    
    Call Me.Show(1, owner)
End Sub


Private Sub ConfigUpdateData()
'���ø�������
    Dim blnFind As Boolean
    
    If mlngBorrowId <= 0 Then Exit Sub
    
    With mufgParentBorrowGrid
        txtBorrowPeople.Text = .Text(.SelectionRow, gstrPatholCol_������)
        dtpBorrowDate.value = CDate(.Text(.SelectionRow, gstrPatholCol_��������))
        
        cbxCardType.ListIndex = .GetFieldDataValue(gstrPatholCol_֤������, .Text(.SelectionRow, gstrPatholCol_֤������), blnFind)
        
        txtCardNum.Text = .Text(.SelectionRow, gstrPatholCol_֤������)
        txtMoney.Text = .Text(.SelectionRow, gstrPatholCol_Ѻ��)
        txtBorrowDays.Text = .Text(.SelectionRow, gstrPatholCol_��������)
        txtMobilePhone.Text = .Text(.SelectionRow, gstrPatholCol_��ϵ�绰)
        txtAddress.Text = .Text(.SelectionRow, gstrPatholCol_��ϵ��ַ)
        rtfReason.Text = .Text(.SelectionRow, gstrPatholCol_����ԭ��)
        txtMemo.Text = .Text(.SelectionRow, gstrPatholCol_��ע)
        txtEnregPeople.Text = .Text(.SelectionRow, gstrPatholCol_�Ǽ���)
        chkBorrowType.value = IIf(.Text(.SelectionRow, gstrPatholCol_��������) = "�ڲ�����", True, False)
    End With
End Sub


Private Sub LoadBorrowMaterialDetail()
'�����ѽ��Ĳ�����ϸ
    Dim strSQL As String
    
    If mlngBorrowId <= 0 Then Exit Sub
    

    strSQL = " select b.id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, '����' as �������," & _
            " case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, " & _
            " nvl(a.��������, 0) as ��������,e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
            " from ��������Ϣ d, ����ȡ����Ϣ c, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a " & _
            " Where c.����ҽ��id = d.����ҽ��id And c.�Ŀ�id = b.�Ŀ�id and e.id=b.����ID And a.�鵵id = b.ID And b.������Դ = 1 And a.����id = [1] " & _
        " Union All " & _
            " select b.id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, '��Ƭ' as �������, " & _
            " decode(o.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, " & _
            " nvl(a.��������, 0) as ��������,e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
            " from ��������Ϣ d, ����ȡ����Ϣ c, ������Ƭ��Ϣ o, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a " & _
            " Where c.����ҽ��id = d.����ҽ��id And o.����ҽ��id = c.����ҽ��id " & _
            " and o.id = b.��Ƭid and e.id=b.����ID and a.�鵵id=b.id and b.������Դ=2 and a.����id=[1] " & _
        " Union All " & _
            " select b.id, d.�������,d.�����,c.���,c.�걾����,c.ȡ��λ��, " & _
            " decode(o.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
            " decode(o.�ؼ�ϸĿ,0,decode(o.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || q.�������� || decode(o.��������,-1,'-��',0,'','-��' || o.��������) || ')' as ������ϸ, " & _
            " nvl(a.��������, 0) as ��������, e.��������, e.��ϸ��ַ, " & _
            " '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ�� " & _
            " from ��������Ϣ d, ����ȡ����Ϣ c, ��������Ϣ q, �����ؼ���Ϣ o, ��������Ϣ e, ����鵵��Ϣ b, ������Ĺ��� a " & _
            " Where c.����ҽ��id = d.����ҽ��id And q.����ID = o.����ID And o.����ҽ��id = c.����ҽ��id " & _
            " and o.id = b.�ؼ�id and e.id=b.����ID and a.�鵵id=b.id and b.������Դ=3 and a.����id=[1] "

    Set ufgMaterialEnreged.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngBorrowId)
    Call ufgMaterialEnreged.RefreshData
End Sub


Private Sub FilterQueryMaterialData()
'���˲�ѯ���Ĳ�������
    Dim strFilter As String
    
    strFilter = ""
    
    If ufgMaterialEnreg.DataGrid.Rows < 2 Then Exit Sub
    
    If chkMaterial(0).value <> 0 Then
        strFilter = "�������='����'"
    End If
    
    If chkMaterial(1).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & "�������='��Ƭ'"
    End If
    
    If chkMaterial(2).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & "�������='����' or �������='����' or �������='��Ⱦ'"
    End If
    
    ufgMaterialEnreg.AdoData.Filter = strFilter
    
    Call ufgMaterialEnreg.RefreshData
    
End Sub


Private Sub chkMaterial_Click(Index As Integer)
On Error GoTo ErrHandle
    Call FilterQueryMaterialData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function LendMaterial() As String
'�������

    Dim lngNewRow As Long
    Dim i As Long
    Dim strLog As String
    
    
    strLog = ""
    
    For i = 1 To ufgMaterialEnreg.GridRows - 1
    
        If ufgMaterialEnreg.GetRowCheck(i) Then
            If ufgMaterialEnreged.FindRowIndex(ufgMaterialEnreg.Text(i, gstrPatholCol_ID), gstrPatholCol_ID, True) < 1 Then
                If Val(ufgMaterialEnreg.Text(i, gstrPatholCol_�������)) > Val(ufgMaterialEnreg.Text(i, gstrPatholCol_�ɽ�����)) Then
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    
                    strLog = strLog & "��ѡ���� [�����:" & ufgMaterialEnreg.Text(i, gstrPatholCol_�����) & " �Ŀ��:" & ufgMaterialEnreg.Text(i, gstrPatholCol_�Ŀ��) & _
                                    " ������ϸ:" & ufgMaterialEnreg.Text(i, gstrPatholCol_������ϸ) & "] ����������ܴ��ڿɽ�������δ�ܽ��н��������"
                Else
                
                    lngNewRow = ufgMaterialEnreged.NewRow
                
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_ID) = ufgMaterialEnreg.Text(i, gstrPatholCol_ID)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_�������) = ufgMaterialEnreg.Text(i, gstrPatholCol_�������)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_�����) = ufgMaterialEnreg.Text(i, gstrPatholCol_�����)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_�Ŀ��) = ufgMaterialEnreg.Text(i, gstrPatholCol_�Ŀ��)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_�걾����) = ufgMaterialEnreg.Text(i, gstrPatholCol_�걾����)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_ȡ��λ��) = ufgMaterialEnreg.Text(i, gstrPatholCol_ȡ��λ��)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_������ϸ) = ufgMaterialEnreg.Text(i, gstrPatholCol_������ϸ)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_�������) = ufgMaterialEnreg.Text(i, gstrPatholCol_�������)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_��������) = ufgMaterialEnreg.Text(i, gstrPatholCol_�������)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_��������) = ufgMaterialEnreg.Text(i, gstrPatholCol_��������)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_���λ��) = ufgMaterialEnreg.Text(i, gstrPatholCol_���λ��)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_��ϸ��ַ) = ufgMaterialEnreg.Text(i, gstrPatholCol_��ϸ��ַ)
                
                    Call ufgMaterialEnreged.SetRowCheck(lngNewRow, False)
                End If

            Else
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    strLog = strLog & "��ѡ���� [�����:" & ufgMaterialEnreg.Text(i, gstrPatholCol_�����) & " �Ŀ��:" & ufgMaterialEnreg.Text(i, gstrPatholCol_�Ŀ��) & _
                                    " ������ϸ:" & ufgMaterialEnreg.Text(i, gstrPatholCol_������ϸ) & "] ���ڽ�������б��У������ٴν��н��������"
            End If
        End If
    Next i
    
    Call ufgMaterialEnreged.LocateRow(lngNewRow)
    
    LendMaterial = strLog
End Function


Private Sub CancelLend()
'�������Ͻ��
    Dim i As Long
    
    For i = ufgMaterialEnreged.GridRows - 1 To 1 Step -1
        If ufgMaterialEnreged.GetRowCheck(i) Then
            Call ufgMaterialEnreged.RemoveRow(i)
        End If
    Next i
End Sub


Private Sub cmdCancel_Click()
On Error GoTo ErrHandle
    blnIsOk = False
    
    Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancelLend_Click()
'�������Ͻ��
On Error GoTo ErrHandle
    If Not ufgMaterialEnreged.IsCheckedRow Then
        Call MsgBoxD(Me, "�빴ѡ��Ҫ��������Ĳ��ϡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call CancelLend
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdLend_Click()
'�������
On Error GoTo ErrHandle
    Dim strInf As String
    
    If Not ufgMaterialEnreg.IsCheckedRow Then
        Call MsgBoxD(Me, "�빴ѡ��Ҫ����Ĳ��ϼ�¼��", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strInf = LendMaterial
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
         '�������ʾ����������⣬�򵯳���ʾ�󽫽��㶨λ����������ϣ������û��޸�
        If InStr(strInf, "����") > 0 Then
            '���������Ĭ�ϳ�1
            ufgMaterialEnreg.Text(ufgMaterialEnreg.SelectionRow, gstrPatholCol_�������) = 1
            Call ufgDataGridSetFocus(ufgMaterialEnreg, ufgMaterialEnreg.SelectionRow, ufgMaterialEnreg.GetColIndexWithRowCheck)
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdQuery_Click()
On Error GoTo ErrHandle
    Call QueryPatholMaterialData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtPatholNum_KeyPress(KeyAscii As Integer)
'�س���ݲ�ѯ
On Error GoTo ErrHandle

    If KeyAscii = 13 Then
         '���ò�ѯ����
         Call QueryPatholMaterialData
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtPatholName_KeyPress(KeyAscii As Integer)
'�س�ִ�в�ѯ
On Error GoTo ErrHandle

    If KeyAscii = 13 Then
         '���ò�ѯ����
         Call QueryPatholMaterialData
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub QueryPatholMaterialData()
'��ѯ�������
    Dim strSQL As String
    Dim strFilter As String
    Dim strLinkTable As String
    
    
    strFilter = " and d.����ʱ�� between [1] and [2] "
    
    strLinkTable = ""
    
    If txtPatholNum.Text <> "" Then
        strFilter = " and d.�����=[3] "
    Else
        If txtPatholName.Text <> "" Then
            'strLinkTable = "(select id from ����ҽ����¼ a, ������Ϣ b where a.����ID=b.����ID and a.���ID is null and b.����" & IIf(InStr(txtPatholName.Text, "%") > 1, " like [4]", "=[4]") & ") h "
            'strFilter = strFilter & " and d.ҽ��ID=h.ID "
'            strFilter = strFilter & " " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.����  like [4]", " and h.���� =[4]")
        End If
    End If
        
    
    'ͳ����ʧ�Ĳ�������(��ѯ��������ʱ��ֻ��ͳ��δ�黹�Ľ������������ֹ黹�ͽ�����ʧ����ļ�¼����������ʧ��������ʧ���������ֳ�)
    strLinkTable = IIf(strLinkTable <> "", strLinkTable & ",", "") & _
                    " (select nvl(sum(��ʧ����),0) as ��ʧ����, �鵵ID " & _
                    " from ������ʧ��Ϣ a, ����鵵��Ϣ b, ��������Ϣ d Where a.�鵵ID = b.ID And b.����ҽ��id = d.����ҽ��id " & _
                    Replace(strFilter, "and d.ҽ��ID=h.ID", "") & " group by �鵵ID ) x, " & _
                    " (select (nvl(sum(��������), 0) - nvl(sum(�黹����), 0)) as �ѽ�����, �鵵ID " & _
                    " from ������Ĺ��� a, ����鵵��Ϣ b, ��������Ϣ d where a.�鵵ID = b.ID And b.����ҽ��id = d.����ҽ��id and a.�黹״̬=0" & _
                    Replace(strFilter, "and d.ҽ��ID=h.ID", "") & " group by �鵵ID" & ") y"
    
    
    
    strSQL = "select /*+ Rule*/ * from (select d.�������, d.�����, h.����, a.id, c.���, c.�걾����, c.ȡ��λ��, '����' as �������, " & _
            " case when (c.������ - nvl(x.��ʧ����,0) - nvl(y.�ѽ�����, 0)) <= 0 then 0 else 1 end as �������," & _
            " case when c.����ID is null then '����ȡ��' else '��ȡ��' end as ������ϸ, " & _
            " (c.������ - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0) ) as �ɽ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬, a.����״̬," & _
            " f.��������, '����:' || f.�������� || ' ���:' || f.������� || ' ����:' || f.�������� as ���λ��, f.��ϸ��ַ " & _
            " from ����鵵��Ϣ a, ����ȡ����Ϣ c, ��������Ϣ d, ��������Ϣ f, ����ҽ����¼ h, " & strLinkTable & _
            " where a.�Ŀ�id=c.�Ŀ�id and c.����ҽ��id=d.����ҽ��id and d.ҽ��ID=h.id and h.���ID is null and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) and a.����id=f.id and f.����״̬=1 " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.����  like [4]", IIf(txtPatholName.Text = "", "", " and h.���� =[4]")) & strFilter & _
        " Union All " & _
            " select d.�������, d.�����,h.����, a.id, c.���, c.�걾����, c.ȡ��λ��, '��Ƭ' as �������, " & _
            " case when (b.��Ƭ�� - nvl(x.��ʧ����,0) - nvl(y.�ѽ�����, 0)) <= 0 then 0 else 1 end as �������, " & _
            " decode(b.��Ƭ��ʽ,0,'����',1,'����',2,'����',3,'����',4,'��Ƭ',5,'��Ⱦ',6,'��Ƭ','����') as ������ϸ, " & _
            " (b.��Ƭ�� - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0)) as �ɽ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬, a.����״̬, " & _
            " e.��������, '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ��, e.��ϸ��ַ " & _
            " from ����鵵��Ϣ a, ������Ƭ��Ϣ b, ����ȡ����Ϣ c, ��������Ϣ d, ��������Ϣ e, ����ҽ����¼ h," & strLinkTable & _
            " where a.��Ƭid=b.id and b.�Ŀ�id=c.�Ŀ�id and c.����ҽ��id=d.����ҽ��id and d.ҽ��ID=h.id and h.���ID is null and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) and a.����id=e.id  and e.����״̬=1 " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.����  like [4]", IIf(txtPatholName.Text = "", "", " and h.���� =[4]")) & strFilter & _
        " Union All " & _
            " select d.�������, d.�����,h.����, a.id, c.���, c.�걾����, c.ȡ��λ��, decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����') as �������, " & _
            " case when (1 - nvl(x.��ʧ����,0) - nvl(y.�ѽ�����, 0)) <= 0 then 0 else 1 end as �������, " & _
            " decode(b.�ؼ�ϸĿ,0,decode(b.�ؼ�����,0, '����',1,'��Ⱦ',2,'����'),1,'����',2,'����ҩ',3,'ӫ��',4,'��ͨ') || '(' || f.�������� || decode(b.��������,-1,'-��',0,'','-��' || b.��������) || ')' as ������ϸ, " & _
            " (1 - nvl(x.��ʧ����,0)  - nvl(y.�ѽ�����, 0)) as �ɽ�����, nvl(x.��ʧ����,0) as ��ʧ����, nvl(y.�ѽ�����,0) as �ѽ�����, a.���״̬, a.����״̬, " & _
            " e.��������, '����:' || e.�������� || ' ���:' || e.������� || ' ����:' || e.�������� as ���λ��, e.��ϸ��ַ " & _
            " from ����鵵��Ϣ a, �����ؼ���Ϣ b, ����ȡ����Ϣ c, ��������Ϣ d, ��������Ϣ e, ��������Ϣ f, ����ҽ����¼ h, " & strLinkTable & _
            " where a.�ؼ�id=b.id and b.�Ŀ�id=c.�Ŀ�id and c.����ҽ��id=d.����ҽ��id and d.ҽ��ID=h.id and h.���ID is null and a.id = x.�鵵ID(+) and a.id=y.�鵵id(+) " & _
            " and a.����id=e.id  and e.����״̬=1 and b.����ID=f.����ID " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.����  like [4]", IIf(txtPatholName.Text = "", "", " and h.���� =[4]")) & strFilter & _
        ") order by �ɽ����� desc,�������, ���,������ϸ,���״̬"


    Set ufgMaterialEnreg.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                    CDate(Format(dtpStartDate.value, "yyyy-mm-dd 00:00:00")), _
                                                    CDate(Format(dtpEndDate.value, "yyyy-mm-dd 23:59:59")), _
                                                    txtPatholNum.Text, _
                                                    txtPatholName.Text)
                                                    
    Call ufgMaterialEnreg.RefreshData
                                                          

    If ufgMaterialEnreg.AdoData.RecordCount <= 0 Then
        Call MsgBoxD(Me, "δ��ѯ��������ݡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
End Sub

Private Function CheckDataIsValid() As String
'��������Ƿ���Ч����Ч�򷵻ؿ��ַ���
    If ufgMaterialEnreged.ShowingDataRowCount <= 0 Then
        CheckDataIsValid = "û��ѡȡ�ɽ��ĵĲ��ϡ�"
        Call ufgMaterialEnreged.SetFocus
        
        Exit Function
    End If
    
    If Trim(txtBorrowPeople.Text) = "" Then
        CheckDataIsValid = "�����˲���Ϊ�ա�"
        Call txtBorrowPeople.SetFocus
        
        Exit Function
    End If
    
    If Trim(txtCardNum.Text) = "" Then
        CheckDataIsValid = "֤�����벻��Ϊ�ա�"
        Call txtCardNum.SetFocus
        
        Exit Function
    End If
    
    
    If Val(txtBorrowDays.Text) <= 0 Then
        CheckDataIsValid = "������������С�ڻ����0��"
        Call txtBorrowDays.SetFocus
        
        Exit Function
    End If
    
    If Trim(rtfReason.Text) = "" Then
        CheckDataIsValid = "����ԭ����Ϊ�ա�"
        Call rtfReason.SetFocus
        
        Exit Function
    End If
End Function


Private Sub NewBorrow()
'��������
    Dim i As Integer
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngNewBorrowId As Long
    Dim lngNewRecordIndex As Long
    
    strSQL = "select Zl_�������_��������([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12]) as ����ֵ from dual"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                            txtBorrowPeople.Text, _
                                            CDate(dtpBorrowDate.value), _
                                            cbxCardType.ListIndex, _
                                            txtCardNum.Text, _
                                            Val(txtMoney.Text), _
                                            Val(txtBorrowDays.Text), _
                                            txtMobilePhone.Text, _
                                            txtAddress.Text, _
                                            rtfReason.Text, _
                                            UserInfo.����, _
                                            txtMemo.Text, _
                                            IIf(chkBorrowType.value <> 0, 0, 1) _
                                            )
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "NewBorrow", "δ�ɹ���ȡ������Ľ���ID,���β���ʧ�ܡ�")
        Exit Sub
    End If
    
    lngNewBorrowId = Val(Nvl(rsData!����ֵ))
    
    Call gcnOracle.BeginTrans
    
On Error GoTo errTrans
    For i = 1 To ufgMaterialEnreged.GridRows - 1
        Call zlDatabase.ExecuteProcedure("Zl_�������_��������(" & lngNewBorrowId & "," & _
                                                                ufgMaterialEnreged.Text(i, gstrPatholCol_ID) & "," & _
                                                                Val(ufgMaterialEnreged.Text(i, gstrPatholCol_��������)) & ")", _
                                                                Me.Caption)
    Next i
    
    Call gcnOracle.CommitTrans
    
    
    
    With mufgParentBorrowGrid
        lngNewRecordIndex = .NewRow
        
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_ID, lngNewBorrowId, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_���ĺ�, lngNewBorrowId, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_������, txtBorrowPeople.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_��������, Format(dtpBorrowDate.value, "yyyy-mm-dd"), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_�黹����, Format(dtpBorrowDate.value + Val(txtBorrowDays.Text), "yyyy-mm-dd"), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_֤������, cbxCardType.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_֤������, txtCardNum.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_Ѻ��, Val(txtMoney.Text), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_��������, Val(txtBorrowDays.Text), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_��ϵ�绰, txtMobilePhone.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_��ϵ��ַ, txtAddress.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_����ԭ��, rtfReason.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_��ע, txtMemo.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_��������, IIf(chkBorrowType.value <> 0, "�ڲ�����", "�ⲿ����"), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_�黹״̬, "δ�黹", True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_ȷ��״̬, "δȷ��", True)
        
        Call .LocateRow(lngNewRecordIndex)
        
    End With
    
    Exit Sub
errTrans:
    Call gcnOracle.RollbackTrans
End Sub


Private Sub UpdateBorrow()
'���½���
    Dim i As Integer
        
    Call gcnOracle.BeginTrans
    
On Error GoTo errTrans

    '���½��ļ�¼
    Call zlDatabase.ExecuteProcedure("Zl_�������_���½���(" & _
                                            mlngBorrowId & ",'" & _
                                            txtBorrowPeople.Text & "'," & _
                                            zlStr.To_Date(dtpBorrowDate.value) & "," & _
                                            cbxCardType.ListIndex & ",'" & _
                                            txtCardNum.Text & "'," & _
                                            Val(txtMoney.Text) & "," & _
                                            Val(txtBorrowDays.Text) & ",'" & _
                                            txtMobilePhone.Text & "','" & _
                                            txtAddress.Text & "','" & _
                                            rtfReason.Text & "','" & _
                                            txtEnregPeople.Text & "','" & _
                                            txtMemo.Text & "'," & _
                                            IIf(chkBorrowType.value <> 0, 0, 1) & ")", Me.Caption)

    'ɾ�����н��Ĳ���
    Call zlDatabase.ExecuteProcedure("Zl_�������_�������(" & mlngBorrowId & ")", Me.Caption)

    For i = 1 To ufgMaterialEnreged.GridRows - 1
        Call zlDatabase.ExecuteProcedure("Zl_�������_��������(" & mlngBorrowId & "," & _
                                                                ufgMaterialEnreged.Text(i, gstrPatholCol_ID) & "," & _
                                                                Val(ufgMaterialEnreged.Text(i, gstrPatholCol_��������)) & ")", _
                                                                Me.Caption)
    Next i
    
    Call gcnOracle.CommitTrans
    
    
    With mufgParentBorrowGrid
        Call .SyncText(.SelectionRow, gstrPatholCol_������, txtBorrowPeople.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_��������, dtpBorrowDate.value)
        Call .SyncText(.SelectionRow, gstrPatholCol_֤������, cbxCardType.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_֤������, txtCardNum.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_Ѻ��, Val(txtMoney.Text))
        Call .SyncText(.SelectionRow, gstrPatholCol_��������, Val(txtBorrowDays.Text))
        Call .SyncText(.SelectionRow, gstrPatholCol_��ϵ�绰, txtMobilePhone.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_��ϵ��ַ, txtAddress.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_����ԭ��, rtfReason.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_��ע, txtMemo.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_��������, IIf(chkBorrowType.value <> 0, "�ڲ�����", "�ⲿ����"))
        Call .SyncText(.SelectionRow, gstrPatholCol_�黹״̬, "δ�黹")
        Call .SyncText(.SelectionRow, gstrPatholCol_ȷ��״̬, "δȷ��")
        
'        Call .LocateRow(.SelectRowIndex)
        
    End With
    
    Exit Sub
errTrans:
    Call gcnOracle.RollbackTrans
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub

Private Sub cmdSure_Click()
'ȷ�ϲ��Ͻ���
On Error GoTo ErrHandle
    Dim strInf As String
    
    strInf = CheckDataIsValid()
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not mblnIsUpdate Then
        Call NewBorrow
    Else
        Call UpdateBorrow
    End If
    
    blnIsOk = True
    
    Call Me.Hide
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
'    #If DebugState = True Then
'        Call InitDebugObject(1294, Me, "zlhis", "HIS")
'    #End If
    
    dtpBorrowDate.value = zlDatabase.Currentdate
    
    dtpStartDate.value = Format(DateAdd("m", -6, dtpBorrowDate.value), "yyyy-mm-dd")
    dtpEndDate.value = Format(dtpBorrowDate.value, "yyyy-mm-dd")

    Call LoadCardType
    
    Call InitMaterialList
    Call InitMaterialEnregedList

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadCardType()
'0-���֤,1-ѧ��֤,2-����֤,3-��ʻ֤,4-����,5-�籣��,6-�м�֤,7-����
    Call cbxCardType.AddItem("0-���֤")
    Call cbxCardType.AddItem("1-ѧ��֤")
    Call cbxCardType.AddItem("2-����֤")
    Call cbxCardType.AddItem("3-��ʻ֤")
    Call cbxCardType.AddItem("4-����")
    Call cbxCardType.AddItem("5-�籣��")
    Call cbxCardType.AddItem("6-�м�֤")
    Call cbxCardType.AddItem("7-����")
    
    cbxCardType.ListIndex = 0
End Sub


Private Sub InitMaterialList()
    '��������
    ufgMaterialEnreg.GridRows = glngStandardRowCount
    '�����и�
    ufgMaterialEnreg.RowHeightMin = glngStandardRowHeight
    
    '��ʼ�����ϲ�ѯ�б�
    ufgMaterialEnreg.IsKeepRows = False
    ufgMaterialEnreg.DefaultColNames = gstrMaterialBorrowEnregCols
    ufgMaterialEnreg.ColNames = gstrMaterialBorrowEnregCols
    ufgMaterialEnreg.ColConvertFormat = gstrMaterialBorrowEnregConvertFormat
End Sub



Private Sub InitMaterialEnregedList()
    '��������
    ufgMaterialEnreged.GridRows = glngStandardRowCount
    '�����и�
    ufgMaterialEnreged.RowHeightMin = glngStandardRowHeight

    '��ʼ�����ϲ�ѯ�б�
    ufgMaterialEnreged.IsKeepRows = False
    ufgMaterialEnreged.DefaultColNames = gstrMaterialBorrowEnregedCols
    ufgMaterialEnreged.ColNames = gstrMaterialBorrowEnregedCols
    ufgMaterialEnreged.ColConvertFormat = gstrMaterialBorrowEnregConvertFormat
End Sub

Private Sub Form_Resize()
On Error Resume Next
        
    Picture0.Left = 120
    Picture0.Top = 120
    Picture0.Width = Me.ScaleWidth - Frame2.Width - 360
    Picture0.Height = Me.ScaleHeight - 240
    
    Frame2.Top = 0
    Frame2.Left = Me.ScaleWidth - Frame2.Width - 120
    Frame2.Height = Me.ScaleHeight - 120
    
    Call ucSplitter1.RePaint
err.Clear
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
    ufgMaterialEnreg.Left = 0
    ufgMaterialEnreg.Top = framNameQuery.Top + framNameQuery.Height + 120
    ufgMaterialEnreg.Width = Picture1.ScaleWidth
    ufgMaterialEnreg.Height = Picture1.ScaleHeight - cmdLend.Height - framNameQuery.Height - 360
    
    cmdLend.Top = ufgMaterialEnreg.Top + ufgMaterialEnreg.Height + 120
    cmdLend.Left = Picture1.ScaleWidth - cmdLend.Width
    
    chkMaterial(0).Top = cmdLend.Top + 60
    chkMaterial(1).Top = cmdLend.Top + 60
    chkMaterial(2).Top = cmdLend.Top + 60
    
    
err.Clear
End Sub


Private Sub Picture2_Resize()
On Error Resume Next
    ufgMaterialEnreged.Top = 0
    ufgMaterialEnreged.Left = 0
    ufgMaterialEnreged.Width = Picture2.ScaleWidth
    ufgMaterialEnreged.Height = Picture2.ScaleHeight - cmdCancelLend.Height - 120
    
    cmdCancelLend.Top = ufgMaterialEnreged.Height + 120
    cmdCancelLend.Left = Picture2.ScaleWidth - cmdCancelLend.Width
err.Clear
End Sub

Private Sub txtPatholName_Change()
On Error Resume Next
    dtpStartDate.Enabled = IIf(txtPatholName.Text = "", False, True)
    dtpEndDate.Enabled = IIf(txtPatholName.Text = "", False, True)
    
    err.Clear
End Sub

Private Sub ufgMaterialEnreg_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInf As String

    If mblnIsEnter Then
        If Not ufgMaterialEnreg.IsCheckedRow Then
            Call MsgBoxD(Me, "�빴ѡ��Ҫ����Ĳ��ϼ�¼��", vbOKOnly, Me.Caption)
            Exit Sub
        End If

        strInf = LendMaterial

        If strInf <> "" Then
           Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
           '�������ʾ����������⣬�򵯳���ʾ�󽫽��㶨λ����������ϣ������û��޸�
           If InStr(strInf, "����") > 0 Then
                '���������Ĭ�ϳ�1
                 ufgMaterialEnreg.Text(ufgMaterialEnreg.SelectionRow, gstrPatholCol_�������) = 1
                 Call ufgDataGridSetFocus(ufgMaterialEnreg, Row, Col - 1)
           End If
        End If
        
        mblnIsEnter = False
    End If
End Sub

Private Sub ufgMaterialEnreg_OnCheckChanged(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
        '���õ�Ԫ��õ����㷽��
        Call ufgDataGridSetFocus(ufgMaterialEnreg, Row, Col)
    err.Clear
End Sub

Private Sub ufgDataGridSetFocus(ufgData As ucFlexGrid, ByVal Row As Long, ByVal Col As Long)
'ʹĳ�ĵ�Ԫ��õ����㣬������ڱ༭״̬
    ufgData.DataGrid.SetFocus
    If Col = ufgData.GetColIndexWithRowCheck Then
        If ufgData.GetRowCheck(Row) Then
            Call ufgData.DataGrid.Select(Row, Col + 1)
            Call ufgData.DataGrid.ShowCell(Row, Col + 1)
            Call ufgData.DataGrid.EditCell
        End If
    End If
End Sub

Private Sub ufgMaterialEnreg_OnColsNameReSet()
On Error GoTo ErrHandle

    If ufgMaterialEnreg.DataGrid.Rows > 1 Then Call QueryPatholMaterialData

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMaterialEnreg_OnKeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     '�ж��Ƿ��µ��ǻس�������������浽ģ�������
     mblnIsEnter = IIf(KeyAscii = 13, True, False)
End Sub



Private Sub ufgMaterialEnreg_OnKeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    If KeyCode = 13 Then
'        MsgBox "KeyUpEdit�¼�"
'    End If
End Sub

Private Sub ufgMaterialEnreg_OnNewRow(ByVal Row As Long)
On Error Resume Next
    If Val(ufgMaterialEnreg.Text(Row, gstrPatholCol_�ɽ�����)) <= 0 Then
        Call ufgMaterialEnreg.DisableCheck(Row, ufgMaterialEnreg.GetColIndexWithRowCheck)
    End If
    
    err.Clear
End Sub

Private Sub ufgMaterialEnreg_OnSelChange()
On Error Resume Next
    Dim lngFindRow As Long
    
    If Not ufgMaterialEnreg.IsSelectionRow Then Exit Sub
    
    lngFindRow = ufgMaterialEnreged.FindRowIndex(ufgMaterialEnreg.Text(ufgMaterialEnreg.SelectionRow, gstrPatholCol_ID), gstrPatholCol_ID, True)
    
    If lngFindRow >= 1 Then
        Call ufgMaterialEnreged.LocateRow(lngFindRow)
    End If
err.Clear
End Sub

Private Sub ufgMaterialEnreged_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrHandle

    Call LoadBorrowMaterialDetail
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

