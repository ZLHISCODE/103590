VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISAduitFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   7695
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7830
   Icon            =   "frmCISAduitFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame frmFind 
      Caption         =   "��������"
      Height          =   6855
      Left            =   210
      TabIndex        =   42
      Top             =   105
      Width           =   7410
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   4
         Left            =   1365
         TabIndex        =   7
         Top             =   870
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   5
         Left            =   5235
         TabIndex        =   8
         Top             =   855
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin VB.CommandButton cmdҩƷ��Ϣ 
         Height          =   300
         Left            =   6945
         Picture         =   "frmCISAduitFilter.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   6375
         Width           =   300
      End
      Begin VB.CommandButton cmd�������� 
         Height          =   300
         Left            =   6930
         Picture         =   "frmCISAduitFilter.frx":685E
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5940
         Width           =   300
      End
      Begin VB.CommandButton cmd������� 
         Height          =   300
         Left            =   3045
         Picture         =   "frmCISAduitFilter.frx":D0B0
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6375
         Width           =   300
      End
      Begin VB.CommandButton cmdסԺҽʦ 
         Height          =   300
         Left            =   3045
         Picture         =   "frmCISAduitFilter.frx":13902
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5940
         Width           =   300
      End
      Begin VB.Frame Frame3 
         Caption         =   "��������"
         Height          =   765
         Left            =   285
         TabIndex        =   22
         Top             =   3375
         Width           =   6960
         Begin VB.OptionButton opt 
            Caption         =   "ҽ������(&F)"
            Height          =   240
            Index           =   2
            Left            =   4665
            TabIndex        =   25
            Top             =   360
            Width           =   1470
         End
         Begin VB.OptionButton opt 
            Caption         =   "��ҽ������(&E)"
            Height          =   195
            Index           =   1
            Left            =   2475
            TabIndex        =   24
            Top             =   360
            Width           =   1500
         End
         Begin VB.OptionButton opt 
            Caption         =   "���в���(&D)"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   23
            Top             =   345
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "����״̬"
         Height          =   945
         Left            =   285
         TabIndex        =   16
         Top             =   2310
         Width           =   6960
         Begin VB.CheckBox chk 
            Caption         =   "�ύ����(&6)"
            Height          =   195
            Index           =   5
            Left            =   180
            TabIndex        =   47
            Top             =   255
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "���մ���(&7)"
            Height          =   195
            Index           =   0
            Left            =   2490
            TabIndex        =   17
            Top             =   255
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "���鵵(&B)"
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   19
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������(&8)"
            Height          =   195
            Index           =   2
            Left            =   4680
            TabIndex        =   18
            Top             =   270
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "��鷴��(&9)"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   20
            Top             =   585
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������(&A)"
            Height          =   195
            Index           =   4
            Left            =   2490
            TabIndex        =   21
            Top             =   600
            Value           =   1  'Checked
            Width           =   1335
         End
      End
      Begin VB.ListBox lst 
         Height          =   1320
         Left            =   1365
         Style           =   1  'Checkbox
         TabIndex        =   27
         Top             =   4290
         Width           =   5880
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   5235
         TabIndex        =   15
         Top             =   1815
         Width           =   2010
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1365
         TabIndex        =   13
         Top             =   1815
         Width           =   2010
      End
      Begin VB.TextBox txtסԺҽʦ 
         Height          =   300
         Left            =   1365
         TabIndex        =   29
         Top             =   5940
         Width           =   1650
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   5235
         TabIndex        =   32
         Top             =   5940
         Width           =   1650
      End
      Begin VB.TextBox txt������� 
         Height          =   300
         Left            =   1365
         TabIndex        =   35
         Top             =   6375
         Width           =   1650
      End
      Begin VB.TextBox txtҩƷ��Ϣ 
         Height          =   300
         Left            =   5235
         TabIndex        =   38
         Top             =   6375
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1365
         TabIndex        =   1
         Top             =   360
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   5235
         TabIndex        =   2
         Top             =   375
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   6
         Left            =   1365
         TabIndex        =   10
         Top             =   1320
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   7
         Left            =   5235
         TabIndex        =   11
         Top             =   1320
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   3
         Left            =   5235
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   1365
         TabIndex        =   4
         Top             =   750
         Visible         =   0   'False
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   19136515
         CurrentDate     =   38083
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   4110
         TabIndex        =   45
         Top             =   900
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��Ժ����(&2)"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   6
         Top             =   915
         Width           =   990
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   10995
         Y1              =   5775
         Y2              =   5775
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   10995
         Y1              =   5790
         Y2              =   5805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   11
         Left            =   4110
         TabIndex        =   46
         Top             =   1395
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��ʱ��(&3)"
         Height          =   180
         Index           =   10
         Left            =   285
         TabIndex        =   9
         Top             =   1365
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   9
         Left            =   4110
         TabIndex        =   44
         Top             =   405
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ύʱ��(&1)"
         Height          =   180
         Index           =   8
         Left            =   285
         TabIndex        =   0
         Top             =   420
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������(&G)"
         Height          =   180
         Index           =   6
         Left            =   285
         TabIndex        =   26
         Top             =   4305
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����(&4)"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   12
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���(&5)"
         Height          =   180
         Index           =   4
         Left            =   4110
         TabIndex        =   14
         Top             =   1860
         Width           =   990
      End
      Begin VB.Label labסԺҽʦ 
         AutoSize        =   -1  'True
         Caption         =   "סԺҽʦ(&H)"
         Height          =   180
         Left            =   285
         TabIndex        =   28
         Top             =   6000
         Width           =   990
      End
      Begin VB.Label lab�������� 
         AutoSize        =   -1  'True
         Caption         =   "��������(&I)"
         Height          =   180
         Left            =   4110
         TabIndex        =   31
         Top             =   6000
         Width           =   990
      End
      Begin VB.Label lab������� 
         AutoSize        =   -1  'True
         Caption         =   "�������(&J)"
         Height          =   180
         Left            =   285
         TabIndex        =   34
         Top             =   6435
         Width           =   990
      End
      Begin VB.Label labҩƷ��Ϣ 
         AutoSize        =   -1  'True
         Caption         =   "ҩƷ��Ϣ(&K)"
         Height          =   180
         Left            =   4110
         TabIndex        =   37
         Top             =   6435
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�鵵����(&2)"
         Height          =   180
         Index           =   5
         Left            =   285
         TabIndex        =   3
         Top             =   795
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   7
         Left            =   4110
         TabIndex        =   43
         Top             =   825
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5190
      TabIndex        =   40
      Top             =   7140
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6480
      TabIndex        =   41
      Top             =   7140
      Width           =   1100
   End
End
Attribute VB_Name = "frmCISAduitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private mrsParam As New ADODB.Recordset
Private mblnDataChanged As Boolean
Private mblnOK As Boolean

'######################################################################################################################

Public Function ShowPara(ByVal frmMain As Object, ByRef rsParam As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnOK = False
    
    Set mrsParam = CopyRecordStruct(rsParam)
    Call CopyRecordData(rsParam, mrsParam)
                
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    If ExecuteCommand("��ȡ����") = False Then Exit Function
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        Call DeleteRecordData(rsParam)
        Call CopyRecordData(mrsParam, rsParam)
        ShowPara = mblnOK
    End If
    
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim strTmp As String
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        chk(0).Value = 1
        chk(1).Value = 1
        chk(2).Value = 1
        chk(3).Value = 1
        chk(4).Value = 1
        chk(5).Value = 1
        opt(0).Value = True
        
        dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
        dtp(1).Value = Format(zlDatabase.Currentdate, dtp(1).CustomFormat)
        dtp(2).Value = Format(zlDatabase.Currentdate, dtp(2).CustomFormat)
        dtp(3).Value = Format(zlDatabase.Currentdate, dtp(3).CustomFormat)

        dtp(4).Value = Format(zlDatabase.Currentdate, dtp(4).CustomFormat)
        dtp(5).Value = Format(zlDatabase.Currentdate, dtp(5).CustomFormat)
        
        dtp(6).Value = Format(zlDatabase.Currentdate, dtp(6).CustomFormat)
        dtp(7).Value = Format(zlDatabase.Currentdate, dtp(7).CustomFormat)
        
        cbo(0).Clear
        cbo(0).AddItem ""
        Set rs = gclsPackage.GetBaseCode("����")
        If rs.BOF = False Then
            Call AddComboData(cbo(0), rs, "����", "����", , False)
        End If
                
        cbo(1).Clear
        cbo(1).AddItem ""
        Set rs = gclsPackage.GetBaseCode("���ƽ��")
        If rs.BOF = False Then
            Call AddComboData(cbo(1), rs, "����", "����", , False)
        End If
        
        lst.Clear
        Set rs = gclsPackage.GetInsureKind()
        If rs.BOF = False Then
            Do While Not rs.EOF
                lst.AddItem rs("����").Value
                lst.ItemData(lst.NewIndex) = rs("���").Value
                rs.MoveNext
            Loop
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        
        chk(0).Value = Val(ParamRead(mrsParam, "���մ���"))
        chk(1).Value = Val(ParamRead(mrsParam, "�ܾ�����"))
        chk(2).Value = Val(ParamRead(mrsParam, "�������"))
        chk(3).Value = Val(ParamRead(mrsParam, "��鷴��"))
        chk(4).Value = Val(ParamRead(mrsParam, "�������"))
        chk(5).Value = Val(ParamRead(mrsParam, "�ύ����"))
        cbo(0).Text = Trim(ParamRead(mrsParam, "��ǰ����"))
        cbo(1).Text = Trim(ParamRead(mrsParam, "��Ժ���"))
        
        txtסԺҽʦ.Text = Trim(ParamRead(mrsParam, "סԺҽʦ"))
        txt��������.Tag = Trim(ParamRead(mrsParam, "��������"))
        txt�������.Text = Trim(ParamRead(mrsParam, "�������"))
        txtҩƷ��Ϣ.Tag = Trim(ParamRead(mrsParam, "ҩƷ��Ϣ"))
        
        '��ȡ��������
        If txt��������.Tag <> "" Then
            gstrSQL = "Select ����,���� From ��������Ŀ¼ where ID = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt��������.Tag)
            If rs Is Nothing Then
                txt��������.Tag = ""
            ElseIf rs.EOF Or rs.BOF Then
                txt��������.Tag = ""
            Else
                txt��������.Text = rs!���� & "[" & rs.Fields!���� & "]"
            End If
        End If
        '��ȡҩƷ����
        If txtҩƷ��Ϣ.Tag <> "" Then
        
            gstrSQL = "select ���� from ҩƷĿ¼ where ҩƷID = [1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(txtҩƷ��Ϣ.Tag))
            If rs Is Nothing Then
                txtҩƷ��Ϣ.Tag = ""
            ElseIf rs.EOF Or rs.BOF Then
                txtҩƷ��Ϣ.Tag = ""
            Else
                txtҩƷ��Ϣ.Text = rs.Fields!����
            End If
        End If
        
        intLoop = Val(ParamRead(mrsParam, "��������"))
        If intLoop >= 0 And intLoop <= 2 Then opt(intLoop).Value = True
        
        strTmp = Trim(ParamRead(mrsParam, "ҽ������"))
        If strTmp <> "" Then
            strTmp = "," & strTmp & ","
            For intLoop = 0 To lst.ListCount - 1
                If InStr(strTmp, "," & lst.ItemData(intLoop) & ",") > 0 Then
                    lst.Selected(intLoop) = True
                End If
            Next
        End If
        
        If ParamRead(mrsParam, "��鿪ʼʱ��") <> "" Then dtp(0).Value = Format(ParamRead(mrsParam, "��鿪ʼʱ��"), dtp(0).CustomFormat)
        If ParamRead(mrsParam, "������ʱ��") <> "" Then dtp(1).Value = Format(ParamRead(mrsParam, "������ʱ��"), dtp(1).CustomFormat)
        If ParamRead(mrsParam, "�鵵��ʼʱ��") <> "" Then dtp(2).Value = Format(ParamRead(mrsParam, "�鵵��ʼʱ��"), dtp(2).CustomFormat)
        If ParamRead(mrsParam, "�鵵����ʱ��") <> "" Then dtp(3).Value = Format(ParamRead(mrsParam, "�鵵����ʱ��"), dtp(3).CustomFormat)
        If ParamRead(mrsParam, "��Ժ��ʼʱ��") <> "" Then dtp(4).Value = Format(ParamRead(mrsParam, "��Ժ��ʼʱ��"), dtp(4).CustomFormat)
        If ParamRead(mrsParam, "��Ժ����ʱ��") <> "" Then dtp(5).Value = Format(ParamRead(mrsParam, "��Ժ����ʱ��"), dtp(5).CustomFormat)
        If ParamRead(mrsParam, "��Ժ��ʼʱ��") <> "" Then dtp(6).Value = Format(ParamRead(mrsParam, "ҽ����ʼʱ��"), dtp(6).CustomFormat)
        If ParamRead(mrsParam, "��Ժ����ʱ��") <> "" Then dtp(7).Value = Format(ParamRead(mrsParam, "ҽ������ʱ��"), dtp(7).CustomFormat)
        
        lst.Enabled = opt(2).Value
        DataChanged = False
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"
        
        If chk(0).Value = 0 And chk(1).Value = 0 And chk(2).Value = 0 And chk(3).Value = 0 And chk(4).Value = 0 And chk(5).Value = 0 Then
            ShowSimpleMsg "���մ��󡢾ܾ����ա�����������鷴������ѡ��һ�"
            chk(0).SetFocus
            Exit Function
        End If
        
        If Abs(DateDiff("m", dtp(1).Value, dtp(0).Value)) > 3 Or Abs(DateDiff("m", dtp(3).Value, dtp(2).Value)) > 3 Or Abs(DateDiff("m", dtp(5).Value, dtp(4).Value)) > 3 Then
            If MsgBox("�����õ�ʱ�䷶Χ������3���£����ܻ�������Ƿ������", vbYesNo + vbDefaultButton2, ParamInfo.��Ʒ����) = vbNo Then
                Exit Function
            End If
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"

        
        Call ParamWrite(mrsParam, "���մ���", chk(0).Value)
        Call ParamWrite(mrsParam, "�ܾ�����", chk(1).Value)
        Call ParamWrite(mrsParam, "�������", chk(2).Value)
        Call ParamWrite(mrsParam, "��鷴��", chk(3).Value)
        Call ParamWrite(mrsParam, "�������", chk(4).Value)
        Call ParamWrite(mrsParam, "�ύ����", chk(5).Value)
        Call ParamWrite(mrsParam, "��ǰ����", cbo(0).Text)
        Call ParamWrite(mrsParam, "��Ժ���", cbo(1).Text)
        
        If opt(0).Value Then
            Call ParamWrite(mrsParam, "��������", 0)
        ElseIf opt(1).Value Then
            Call ParamWrite(mrsParam, "��������", 1)
        Else
            Call ParamWrite(mrsParam, "��������", 2)
        End If
        
        strTmp = ""
        For intLoop = 0 To lst.ListCount - 1
            If lst.Selected(intLoop) Then
                strTmp = strTmp & "," & lst.ItemData(intLoop)
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        Call ParamWrite(mrsParam, "ҽ������", strTmp)
        
        Call ParamWrite(mrsParam, "��鿪ʼʱ��", Format(dtp(0).Value, dtp(0).CustomFormat))
        Call ParamWrite(mrsParam, "������ʱ��", Format(dtp(1).Value, dtp(1).CustomFormat))
        Call ParamWrite(mrsParam, "�鵵��ʼʱ��", Format(dtp(2).Value, dtp(2).CustomFormat))
        Call ParamWrite(mrsParam, "�鵵����ʱ��", Format(dtp(3).Value, dtp(3).CustomFormat))
        Call ParamWrite(mrsParam, "��Ժ��ʼʱ��", Format(dtp(4).Value, dtp(4).CustomFormat))
        Call ParamWrite(mrsParam, "��Ժ����ʱ��", Format(dtp(5).Value, dtp(5).CustomFormat))
        Call ParamWrite(mrsParam, "ҽ����ʼʱ��", Format(dtp(6).Value, dtp(6).CustomFormat))
        Call ParamWrite(mrsParam, "ҽ������ʱ��", Format(dtp(7).Value, dtp(7).CustomFormat))
        
        Call ParamWrite(mrsParam, "סԺҽʦ", txtסԺҽʦ.Text)
        Call ParamWrite(mrsParam, "��������", IIf(txt��������.Text = "", "", txt��������.Tag))
        Call ParamWrite(mrsParam, "�������", txt�������.Text)
        Call ParamWrite(mrsParam, "ҩƷ��Ϣ", IIf(txtҩƷ��Ϣ.Text = "", "", txtҩƷ��Ϣ.Tag))
        
    End Select
    
    ExecuteCommand = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

'######################################################################################################################

Private Sub chk_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If DataChanged Then
        If ExecuteCommand("У������") = False Then Exit Sub
        If ExecuteCommand("�������") Then
            mrsParam.Filter = ""
            mblnOK = True
            DataChanged = False
        End If
    End If
    Unload Me
End Sub

Private Sub cmd��������_Click()
On Error GoTo ErrH
    SelectSick
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd�������_Click()
On Error GoTo ErrH
    SelectCheck
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdҩƷ��Ϣ_Click()
On Error GoTo ErrH
    SelectDrug
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdסԺҽʦ_Click()
On Error GoTo ErrH
    SelectDoctor
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub dtp_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub lst_ItemCheck(Item As Integer)
    DataChanged = True
End Sub

Private Sub lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    DataChanged = True
    
    lst.Enabled = (Index = 2)
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt��������_Change()
    DataChanged = True
End Sub

Private Sub txt�������_Change()
    DataChanged = True
End Sub

Private Sub txtҩƷ��Ϣ_Change()
    DataChanged = True
End Sub

Private Sub txtסԺҽʦ_Change()
    DataChanged = True
End Sub

Private Sub txtסԺҽʦ_KeyPress(KeyAscii As Integer)
    If Trim(txtסԺҽʦ.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectDoctor txtסԺҽʦ.Text
    End If
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If Trim(txt��������.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectSick txt��������.Text
    End If
End Sub

Private Sub txtҩƷ��Ϣ_KeyPress(KeyAscii As Integer)
    If Trim(txtҩƷ��Ϣ.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectDrug txtҩƷ��Ϣ.Text
    End If
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)

    If Trim(txt�������.Text) = "" Then Exit Sub
    If KeyAscii = 13 Then
        SelectCheck txt�������.Text
    End If

End Sub

'ѡ��ҽ��
Private Sub SelectDoctor(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "Select c.ID,c.���,c.���� As ����"
        gstrSQL = gstrSQL & vbCrLf & "From ��Ա�� C, ��Ա����˵�� D"
        gstrSQL = gstrSQL & vbCrLf & "Where  c.id = d.��Աid And D.��Ա���� = 'ҽ��'"
        gstrSQL = gstrSQL & vbCrLf & "And (c.���� like '%'||[1]||'%' or ���� like '%'||[1]||'%')"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
        bytRet = ShowPubSelect(Me, txtסԺҽʦ, 2, "���,1200,0,;����,1200,0,", Me.Name & "\ҽ��ѡ��", "����±���ѡ��һ��ҽ��", rsTmp, rsResult, 8790, 4500, False)
    Else
        gstrSQL = gstrSQL & vbCrLf & "Select Id,�ϼ�id,0 As ĩ��,���� as ���,���� From ���ű�"
        gstrSQL = gstrSQL & vbCrLf & "Start With �ϼ�id Is Null"
        gstrSQL = gstrSQL & vbCrLf & "Connect By Prior ID = �ϼ�id"
        gstrSQL = gstrSQL & vbCrLf & "Union All"
        gstrSQL = gstrSQL & vbCrLf & "Select c.id,b.����id As �ϼ�Id,1 As ĩ��,c.���,c.���� As ����"
        gstrSQL = gstrSQL & vbCrLf & "From ������Ա b,��Ա�� C, ��Ա����˵�� D"
        gstrSQL = gstrSQL & vbCrLf & "Where c.Id=b.��Աid and c.id = d.��Աid And D.��Ա���� = 'ҽ��' And b.ȱʡ=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        bytRet = ShowPubSelect(Me, txtסԺҽʦ, 3, "���,1200,0,;����,1200,0,", Me.Name & "\ҽ��ѡ��", "����±���ѡ��һ��ҽ��", rsTmp, rsResult, 8790, 4500, False)
    End If
    
    If rsResult Is Nothing Then
        txtסԺҽʦ.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txtסԺҽʦ.Text = ""
    Else
        txtסԺҽʦ.Text = rsResult!����
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

'ѡ�񼲲�
Private Sub SelectSick(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    gstrSQL = gstrSQL & vbCrLf & "select ID,����,���� from ��������Ŀ¼"
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "Where (���� like '%'||[1]||'%' or ���� like '%'||[1]||'%' or ���� like '%'||[1]||'%')"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
    bytRet = ShowPubSelect(Me, txtסԺҽʦ, 2, "����,1200,0,;����,1200,0,", Me.Name & "\����ѡ��", "����±���ѡ��һ�ּ���", rsTmp, rsResult, 8790, 4500, False)
    
    If rsResult Is Nothing Then
        txt��������.Text = ""
        txt��������.Tag = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt��������.Text = ""
        txt��������.Tag = ""
    Else
        txt��������.Tag = rsResult!ID
        txt��������.Text = rsResult!���� & "[" & rsResult!���� & "]"
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

'ѡ��ҩƷ
Private Sub SelectDrug(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    If strShortName = "" Then
        gstrSQL = gstrSQL & vbCrLf & "select 5 as ID,null as �ϼ�ID,0 as ĩ��,'5' as ����,'����ҩ' as ����,'����ҩ' as ͨ������,'' as ��� from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 6 as ID,null as �ϼ�ID,0 as ĩ��,'6' as ����,'�г�ҩ' as ����,'�г�ҩ' as ͨ������,'' as ��� from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 7 as ID,null as �ϼ�ID,0 as ĩ��,'7' as ����,'�в�ҩ' as ����,'�в�ҩ' as ͨ������,'' as ��� from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "Select * from("
        gstrSQL = gstrSQL & vbCrLf & "select a.ҩƷID as ID,decode(b.���ʷ���,'����ҩ',5,'�г�ҩ',6,7) as �ϼ�ID,1 as ĩ��,a.���� as ����,a.���� as ����,b.ͨ������ as ͨ������ ,a.���"
        gstrSQL = gstrSQL & vbCrLf & "from ҩƷĿ¼ a,ҩƷ��Ϣ b"
        gstrSQL = gstrSQL & vbCrLf & "Where a.ҩ��ID = b.ҩ��ID"
        gstrSQL = gstrSQL & vbCrLf & "Order by b.���ʷ���"
        gstrSQL = gstrSQL & vbCrLf & ")"
    Else
        gstrSQL = gstrSQL & vbCrLf & "select 5 as ID,null as �ϼ�ID,0 as ĩ��,'5' as ����,'����ҩ' as ����,'����ҩ' as ͨ������,'' as ��� from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 6 as ID,null as �ϼ�ID,0 as ĩ��,'6' as ����,'�г�ҩ' as ����,'�г�ҩ' as ͨ������,'' as ��� from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "select 7 as ID,null as �ϼ�ID,0 as ĩ��,'7' as ����,'�в�ҩ' as ����,'�в�ҩ' as ͨ������,'' as ��� from dual union all"
        gstrSQL = gstrSQL & vbCrLf & "Select * from("
        gstrSQL = gstrSQL & vbCrLf & "select a.ҩƷID as ID,decode(b.���ʷ���,'����ҩ',5,'�г�ҩ',6,7) as �ϼ�ID,1 as ĩ��,a.���� as ����,a.���� as ����,b.ͨ������ as ͨ������ ,a.���"
        gstrSQL = gstrSQL & vbCrLf & "from ҩƷĿ¼ a,ҩƷ��Ϣ b"
        gstrSQL = gstrSQL & vbCrLf & "Where a.ҩ��ID = b.ҩ��ID"
        gstrSQL = gstrSQL & vbCrLf & "And (a.���� like '%' || [1] || '%' or a.����  like '%' || [1] || '%' or zlSpellCode(a.����)  like '%' || [1] || '%')"
        gstrSQL = gstrSQL & vbCrLf & "Order by b.���ʷ���"
        gstrSQL = gstrSQL & vbCrLf & ")"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
    bytRet = ShowPubSelect(Me, txtסԺҽʦ, 3, "����,1200,0,;����,1200,0,;ͨ������,1200,0,;���,800,0,", Me.Name & "\����ѡ��", "����±���ѡ��һ�ּ���", rsTmp, rsResult, 8790, 4500, False)
    
    If rsResult Is Nothing Then
        txtҩƷ��Ϣ.Text = ""
        txtҩƷ��Ϣ.Tag = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txtҩƷ��Ϣ.Text = ""
        txtҩƷ��Ϣ.Tag = ""
    Else
        txtҩƷ��Ϣ.Tag = rsResult!ID
        txtҩƷ��Ϣ.Text = rsResult!����
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

'ѡ��������
Private Sub SelectCheck(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo ErrH
    gstrSQL = ""
    gstrSQL = gstrSQL & vbCrLf & "select ���� as ID,����,���� from ���Ƽ������"
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "Where (���� like '%'||[1]||'%' or ���� like '%'||[1]||'%' or ����  like '%' || [1] || '%')"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strShortName))
    bytRet = ShowPubSelect(Me, txtסԺҽʦ, 2, "����,1200,0,;����,1200,0,", Me.Name & "\�������ѡ��", "����±���ѡ��һ�ּ������", rsTmp, rsResult, 8790, 4500, False)
    
    If rsResult Is Nothing Then
        txt�������.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt�������.Text = ""
    Else
        txt�������.Text = rsResult!����
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

