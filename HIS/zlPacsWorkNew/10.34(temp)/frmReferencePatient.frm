VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReferencePatient 
   Caption         =   "��������"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "frmReferencePatient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkUseOldCheckNo 
      Caption         =   "Ӧ�ù������˵ļ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   8880
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9240
      TabIndex        =   6
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisRelating 
      Caption         =   "ȡ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6840
      TabIndex        =   5
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdRelating 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   4
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "��ѯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   8760
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ѯ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   10335
      Begin VB.OptionButton optFilter 
         Caption         =   "���� ="
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   1800
         TabIndex        =   25
         Top             =   772
         Width           =   3075
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1800
         TabIndex        =   18
         Top             =   1725
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "IC����   ="
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   7080
         TabIndex        =   16
         Top             =   1252
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "���֤�� ="
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5520
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   1800
         TabIndex        =   14
         Top             =   1252
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "���￨�� ="
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   7080
         TabIndex        =   12
         Top             =   772
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "סԺ��   ="
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5520
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   7080
         TabIndex        =   10
         Top             =   285
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "�����   ="
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5520
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtFilter 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Top             =   292
         Width           =   3075
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "����     ="
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frmRelated 
      Caption         =   "�ѹ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10335
      Begin MSComctlLib.ListView lvwRelated 
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2143
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�ȴ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   10335
      Begin MSComctlLib.ListView lvwToBeRelate 
         Height          =   1335
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2355
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwStudies 
         Height          =   1335
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2355
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Label lblPatientInfo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   10095
   End
   Begin VB.Label lblPatientInfo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   480
      Width           =   10095
   End
   Begin VB.Label lblPatientInfo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmReferencePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr���� As String      '����������������ʾʱ��Ĭ�ϲ�ѯ����
Private mlngOrderID As Long     '���˱��μ���ҽ��ID
Private mlngPatietID As Long    '����ID
Private mlngRelatingID As Long  '��ǰ�Ĺ���ID
Private mblnInCheck As Boolean  '�Ƿ��ڳ�����Ƹ���Check��״̬�У����ٴ���Check״̬
Private mfrmParent As Form      '������
Private mlngDetpID As Long      '��ǰ����ID
Private mlngStudyNoBuildType As Long        '�������ɷ�ʽ,0-�������� 1-�����ҵ���
Private mstrModality As String              'Ӱ�����


Public Sub zlShowMe(lngOrderID As Long, str���� As String, frmParent As Form, blnShow As Boolean, lngDetpID As Long)
'��ʾ�������˵Ĵ���
'������ lngOrderID --- ҽ��ID
'       str���� --- ��������
'       frmParent --- ������
'       blnShow --- û�пɹ����Ĳ������Ƿ���ʾ���壬True-��ʾ��False-����ʾ
'       mlngDetpID --- ִ�п���ID��������ȡ�������̲���

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim rsToBeRelate As ADODB.Recordset
    
    mstr���� = str����
    mlngOrderID = lngOrderID
    Set mfrmParent = frmParent
    mlngRelatingID = 0
    mlngDetpID = lngDetpID
    mlngStudyNoBuildType = Val(GetDeptPara(mlngDetpID, "�������ɷ�ʽ", 0))
    
    
    On Error GoTo err
    '��ѯ����¼��ǰ�Ĳ���ID
    strSql = "Select a.����id,b.����, b.�Ա�, b.����,to_char(b.��������,'yyyy-mm-dd') ��������, " & _
             " b.�����,b.סԺ��,b.���￨��, " & _
             " b.���֤��,b.ְҵ,b.����,b.����״��,nvl(b.��ͥ��ַ,b.������λ) ��ַ,nvl(b.��ͥ�绰,b.��ϵ�˵绰) �绰 " & _
             " From ����ҽ����¼ a ,������Ϣ b Where a.����id=b.����id and id= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ����ID", mlngOrderID)
    If rsTemp.EOF = True Then Exit Sub
    
    mlngPatietID = rsTemp!����ID
    lblPatientInfo(0).Caption = " ������" & Nvl(rsTemp!����) & " �Ա�" & Nvl(rsTemp!�Ա�) & _
            " ���䣺" & Nvl(rsTemp!����) & " �������ڣ�" & Nvl(rsTemp!��������)
    lblPatientInfo(1).Caption = " ����ţ�" & Nvl(rsTemp!�����) & " סԺ�ţ�" & Nvl(rsTemp!סԺ��) & _
            " ���￨�ţ�" & Nvl(rsTemp!���￨��) & " ���֤�ţ�" & Nvl(rsTemp!���֤��)
    lblPatientInfo(2).Caption = " ���壺" & Nvl(rsTemp!����) & " �绰��" & Nvl(rsTemp!�绰) & " ��ַ��" & Nvl(rsTemp!��ַ)
    
    
    strSql = "Select Ӱ����� From Ӱ�����¼ Where ҽ��ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ѯӰ�����", mlngOrderID)
    If rsTemp.EOF = False Then
        mstrModality = Nvl(rsTemp!Ӱ�����)
    End If
        
    '��ѯ�Ƿ���ͬ������û�й����Ĳ���
    strSql = "Select Distinct a.����id, b.����, b.����id, a.����, a.�Ա�, a.����, a.��������, a.�����, a.סԺ��, a.���￨��, a.�ѱ�," & _
             " a.ҽ�Ƹ��ʽ , a.���֤��, a.ְҵ, a.����, a.����״��, a.��ַ, a.�绰,b.Ӱ�����,c.ִ�п���ID " & _
             " From (Select ����id, ����, �Ա�, ����, To_Char(��������, 'yyyy-mm-dd') As ��������, �����, סԺ��, ���￨��, �ѱ�, " & _
             "        ҽ�Ƹ��ʽ, ���֤��, ְҵ, ����, ����״��, Nvl(��ͥ��ַ, ������λ) As ��ַ, Nvl(��ͥ�绰, ��ϵ�˵绰) �绰 " & _
             "       From ������Ϣ Where ���� = [1] And ����id <> [2]) a, Ӱ�����¼ b, ����ҽ����¼ c " & _
             " Where c.����id = a.����id And c.ID = b.ҽ��ID And b.����id Is Null Order By a.����id "
    
    Set rsToBeRelate = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", mstr����, mlngPatietID)
    
    If mlngStudyNoBuildType = 1 Then
        '�������ɷ�ʽ=1����ֻ��ѯ�����ҵļ��
        rsToBeRelate.Filter = "ִ�п���ID = " & mlngDetpID
    Else
        '�������ɷ�ʽ=0����ֻ��ѯ��Ӱ�����ļ��
        rsToBeRelate.Filter = "Ӱ����� = '" & mstrModality & "'"
    End If
    
    '���û�й����Ĳ��ˣ��Ҳ���ʾ���壬���˳�
    If rsToBeRelate.EOF = True And blnShow = False Then
        Exit Sub
    Else
        '��ʼ������
        Call InitLists
        
        Call FillToBeRelateList(rsToBeRelate)
        
        '������Ѿ��������б�
        strSql = "Select Distinct b.����ID,a.����,a.����ID,b.����,b.�Ա�, b.����,to_char(b.��������,'yyyy-mm-dd') ��������," & _
             " b.�����,b.סԺ��,b.���￨��,b.�ѱ�,b.ҽ�Ƹ��ʽ,b.����ID, " & _
             " b.���֤��,b.ְҵ,b.����,b.����״��,nvl(b.��ͥ��ַ,b.������λ) ��ַ,nvl(b.��ͥ�绰,b.��ϵ�˵绰) �绰 " & _
             " From (Select ҽ��id,����,����ID From Ӱ�����¼ Where ����id =(Select ����ID From Ӱ�����¼ Where ҽ��id=[1])) a, " & _
             " ������Ϣ b, ����ҽ����¼ c " & _
             " Where c.����id = b.����id And a.ҽ��id = c.Id and b.����ID <> [2] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", mlngOrderID, mlngPatietID)
        Call FillRelatedList(rsTemp)
        
        '��ʾ����
        Me.Show 1, mfrmParent
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitLists()
'��ʼ�����ѹ������͡��ȴ��������б�
    
    On Error GoTo err
    '��ʼ�����ѹ������б�
    With lvwRelated.ColumnHeaders
        .Clear
        .Add , , "����", 1000
        .Add , , "�Ա�", 800
        .Add , , "����", 800
        .Add , , "��������", 1200
        .Add , , "�����", 1000
        .Add , , "סԺ��", 1000
        .Add , , "����", 1000
        .Add , , "���￨��", 1000
        .Add , , "���֤��", 1400
        .Add , , "����", 800
        .Add , , "�绰", 1000
        .Add , , "��ַ", 2000
        .Add , , "����ID", 0
    End With
    lvwRelated.ListItems.Add , , "Temp"
    
    '��ʼ�����ȴ��������б�
    With lvwToBeRelate.ColumnHeaders
        .Clear
        .Add , , "����", 1000
        .Add , , "�Ա�", 800
        .Add , , "����", 800
        .Add , , "��������", 1200
        .Add , , "�����", 1000
        .Add , , "סԺ��", 1000
        .Add , , "����", 1000
        .Add , , "���￨��", 1000
        .Add , , "���֤��", 1400
        .Add , , "����", 800
        .Add , , "�绰", 1000
        .Add , , "��ַ", 2000
        .Add , , "����ID", 0
    End With
    lvwToBeRelate.ListItems.Add , , "Temp"
    
    '��ʼ������顱�б�
    With lvwStudies.ColumnHeaders
        .Clear
        .Add , , "���", 800
        .Add , , "Ӱ�����", 1000
        .Add , , "����", 2000
        .Add , , "��ͼʱ��", 2000
        .Add , , "ҽ������", 4000
    End With
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillToBeRelateList(rsToBeRelate As ADODB.Recordset)
'��䡰�ȴ��������б�
'������ rsToBeRelate --- �ȴ������б�

    Dim tmpItem As MSComctlLib.ListItem
    Dim strPatientID As String
    Dim strRelatedPatientID As String   '��¼�ѹ����Ĳ���ID
    Dim i As Integer
    
    On Error GoTo err
    
    '��¼�ѹ����Ĳ���ID
    strRelatedPatientID = ""
    If lvwRelated.ListItems.Count >= 1 Then
        For i = 1 To lvwRelated.ListItems.Count
            strRelatedPatientID = strRelatedPatientID & "," & Mid(lvwRelated.ListItems(i).Key, 2)
        Next i
    End If
    
    '���ȴ��������б�
    strPatientID = ""
    lvwToBeRelate.ListItems.Clear
    
    While rsToBeRelate.EOF = False
        If InStr(strPatientID, rsToBeRelate("����ID")) = 0 And InStr(strRelatedPatientID, rsToBeRelate("����ID")) = 0 Then
            strPatientID = strPatientID & "," & rsToBeRelate("����ID")
            Set tmpItem = lvwToBeRelate.ListItems.Add(, "_" & rsToBeRelate("����ID"), rsToBeRelate("����"))
            tmpItem.SubItems(1) = Nvl(rsToBeRelate("�Ա�"))
            tmpItem.SubItems(2) = Nvl(rsToBeRelate("����"))
            tmpItem.SubItems(3) = Nvl(rsToBeRelate("��������"))
            tmpItem.SubItems(4) = Nvl(rsToBeRelate("�����"))
            tmpItem.SubItems(5) = Nvl(rsToBeRelate("סԺ��"))
            tmpItem.SubItems(6) = Nvl(rsToBeRelate("����"))
            tmpItem.SubItems(7) = Nvl(rsToBeRelate("���￨��"))
            tmpItem.SubItems(8) = Nvl(rsToBeRelate("���֤��"))
            tmpItem.SubItems(9) = Nvl(rsToBeRelate("����"))
            tmpItem.SubItems(10) = Nvl(rsToBeRelate("�绰"))
            tmpItem.SubItems(11) = Nvl(rsToBeRelate("��ַ"))
            tmpItem.SubItems(12) = Nvl(rsToBeRelate("����ID"))
        End If
        rsToBeRelate.MoveNext
    Wend
    
    '��д����б�
    If lvwToBeRelate.ListItems.Count >= 1 Then
        Call lvwToBeRelate_ItemClick(lvwToBeRelate.ListItems(1))
    Else
        lvwStudies.ListItems.Clear
    End If
        
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillRelatedList(rsRelated As ADODB.Recordset)
'��䡰�ѹ������б�
'������ rsRelated --- �ѹ��������ݼ�

    Dim tmpItem As MSComctlLib.ListItem
    Dim strPatientID As String
    
    On Error GoTo err
    
    strPatientID = ""
    lvwRelated.ListItems.Clear
    
    While rsRelated.EOF = False
        If InStr(strPatientID, rsRelated("����ID")) = 0 Then
            strPatientID = strPatientID & "," & rsRelated("����ID")
            Set tmpItem = lvwRelated.ListItems.Add(, "_" & rsRelated("����ID"), rsRelated("����"))
            tmpItem.SubItems(1) = Nvl(rsRelated("�Ա�"))
            tmpItem.SubItems(2) = Nvl(rsRelated("����"))
            tmpItem.SubItems(3) = Nvl(rsRelated("��������"))
            tmpItem.SubItems(4) = Nvl(rsRelated("�����"))
            tmpItem.SubItems(5) = Nvl(rsRelated("סԺ��"))
            tmpItem.SubItems(6) = Nvl(rsRelated("����"))
            tmpItem.SubItems(7) = Nvl(rsRelated("���￨��"))
            tmpItem.SubItems(8) = Nvl(rsRelated("���֤��"))
            tmpItem.SubItems(9) = Nvl(rsRelated("����"))
            tmpItem.SubItems(10) = Nvl(rsRelated("�绰"))
            tmpItem.SubItems(11) = Nvl(rsRelated("��ַ"))
            tmpItem.SubItems(12) = Nvl(rsRelated("����ID"))
            mlngRelatingID = Nvl(rsRelated("����ID"))
        End If
        rsRelated.MoveNext
    Wend
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdDisRelating_Click()
'�ѡ��ѹ������б��б�ѡ�е���Ŀ���óɡ��ȴ�������
    Dim ItemSelected As MSComctlLib.ListItem
    Dim ItemDeleteRelate As MSComctlLib.ListItem
    Dim lngPatientID As Long
    Dim i As Integer
    
    On Error GoTo err
    
    '�жϵ�ǰ���ѹ������б����Ƿ��б�ѡ�е���Ŀ��ѭ����ѡ�е���Ŀ��ȡ������
    For Each ItemSelected In lvwRelated.ListItems
        lngPatientID = Mid(ItemSelected.Key, 2)
        If ItemSelected.Checked = True Then
            'ȡ�����ݿ�Ĺ���
            Call DeleteRelating(lngPatientID)
            '�ѵ�ǰ���ѹ������б���ѡ�е���Ŀ�ƶ������ȴ��������б���
            Set ItemDeleteRelate = lvwToBeRelate.ListItems.Add(, ItemSelected.Key, ItemSelected.Text)
            ItemDeleteRelate.SubItems(1) = ItemSelected.SubItems(1)
            ItemDeleteRelate.SubItems(2) = ItemSelected.SubItems(2)
            ItemDeleteRelate.SubItems(3) = ItemSelected.SubItems(3)
            ItemDeleteRelate.SubItems(4) = ItemSelected.SubItems(4)
            ItemDeleteRelate.SubItems(5) = ItemSelected.SubItems(5)
            ItemDeleteRelate.SubItems(6) = ItemSelected.SubItems(6)
            ItemDeleteRelate.SubItems(7) = ItemSelected.SubItems(7)
            ItemDeleteRelate.SubItems(8) = ItemSelected.SubItems(8)
            ItemDeleteRelate.SubItems(9) = ItemSelected.SubItems(9)
            ItemDeleteRelate.SubItems(10) = ItemSelected.SubItems(10)
            ItemDeleteRelate.SubItems(11) = ItemSelected.SubItems(11)
            ItemDeleteRelate.SubItems(12) = ItemSelected.SubItems(12)
        End If
    Next
    
    'ɾ���б��б�ѡ�е���Ŀ
    For i = lvwRelated.ListItems.Count To 1 Step -1
        If lvwRelated.ListItems(i).Checked = True Then
            lvwRelated.ListItems.Remove i
        End If
    Next i
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
'��ѯ�ȴ������Ĳ�����Ϣ
    Dim i As Integer
    Dim blnQuery As Boolean
    Dim intFilterIndex As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFilter As String
    
    intFilterIndex = -1
    On Error GoTo err
        For i = 0 To 6
            If optFilter(i).value = True Then
                blnQuery = True
                intFilterIndex = i
                Exit For
            End If
        Next i
        
        If blnQuery = False Then
            MsgBoxD mfrmParent, "��ѡ���������ٲ�ѯ", vbOKOnly, "��ʾ��Ϣ"
        Else
            strSql = "Select Distinct b.����ID,a.����,a.����ID,nvl(b.����,a.����) ����,nvl(b.�Ա�,a.�Ա�) �Ա�, nvl(b.����,a.����) ����,to_char(b.��������,'yyyy-mm-dd') ��������," & _
             " b.�����,b.סԺ��,b.���￨��,b.�ѱ�,b.ҽ�Ƹ��ʽ,b.����ID,a.Ӱ�����,c.ִ�п���ID, " & _
             " b.���֤��,b.ְҵ,b.����,b.����״��,nvl(b.��ͥ��ַ,b.������λ) ��ַ,nvl(b.��ͥ�绰,b.��ϵ�˵绰) �绰 " & _
             " From Ӱ�����¼ a, ������Ϣ b, ����ҽ����¼ c " & _
             " Where c.����id = b.����id And a.ҽ��id = c.Id and b.����ID <> [1] "
            If mlngRelatingID <> 0 Then
                strFilter = strFilter & " and (a.����ID <> [2] Or a.����id Is Null) "
            End If
            Select Case intFilterIndex
            Case 0  '����
                strFilter = strFilter & " and b.���� = [3] "
            Case 1  '���￨
                strFilter = strFilter & " and b.���￨�� = [4] "
            Case 2  'IC��
                strFilter = strFilter & " and b.IC���� = [5] "
            Case 3  '�����
                strFilter = strFilter & " and b.����� = [6] "
            Case 4  'סԺ��
                strFilter = strFilter & " and b.סԺ�� = [7] "
            Case 5  '���֤��
                strFilter = strFilter & " and b.���֤�� = [8] "
            Case 6  '����
                strFilter = strFilter & " and a.���� = [9] "
            End Select
            
            strSql = strSql & strFilter
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", mlngPatietID, mlngRelatingID, CStr(txtFilter(0).Text), _
                        CStr(txtFilter(1).Text), CStr(txtFilter(2).Text), CLng(Val(txtFilter(3).Text)), CLng(Val(txtFilter(4).Text)), _
                        CStr(txtFilter(5).Text), CStr(txtFilter(6).Text))
            
            If mlngStudyNoBuildType = 1 Then
                '�������ɷ�ʽ=1����ֻ��ѯ�����ҵļ��
                rsTemp.Filter = "ִ�п���ID = " & mlngDetpID
            Else
                '�������ɷ�ʽ=0����ֻ��ѯ��Ӱ�����ļ��
                rsTemp.Filter = "Ӱ����� = '" & mstrModality & "'"
            End If
    
            Call FillToBeRelateList(rsTemp)
        End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRelating_Click()
'�ѡ��ȴ������б���ѡ�е���Ŀ���óɸ���ǰ���˹���
    Dim ItemSelected As MSComctlLib.ListItem
    Dim ItemRelated As MSComctlLib.ListItem
    Dim lngPatientID As Long
    Dim i As Integer
    Dim str����  As String
    Dim arr����() As String
    Dim strReturn As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    str���� = "||"
    
    '�жϵ�ǰ���ȴ������б��Ƿ��б�ѡ�е���Ŀ,ѭ����ѡ�е���Ŀ�����ù���
    For Each ItemSelected In lvwToBeRelate.ListItems
        lngPatientID = Mid(ItemSelected.Key, 2)
        If ItemSelected.Checked = True Then
            '�����ݿ������ù���
            Call SetRelating(lngPatientID)
            '�ѵ�ǰ���ȴ��������б���ѡ�е���Ŀ�ƶ������ѹ������б���
            Set ItemRelated = lvwRelated.ListItems.Add(, ItemSelected.Key, ItemSelected.Text)
            ItemRelated.SubItems(1) = ItemSelected.SubItems(1)
            ItemRelated.SubItems(2) = ItemSelected.SubItems(2)
            ItemRelated.SubItems(3) = ItemSelected.SubItems(3)
            ItemRelated.SubItems(4) = ItemSelected.SubItems(4)
            ItemRelated.SubItems(5) = ItemSelected.SubItems(5)
            ItemRelated.SubItems(6) = ItemSelected.SubItems(6)
            ItemRelated.SubItems(7) = ItemSelected.SubItems(7)
            ItemRelated.SubItems(8) = ItemSelected.SubItems(8)
            ItemRelated.SubItems(9) = ItemSelected.SubItems(9)
            ItemRelated.SubItems(10) = ItemSelected.SubItems(10)
            ItemRelated.SubItems(11) = ItemSelected.SubItems(11)
            ItemRelated.SubItems(12) = ItemSelected.SubItems(12)
            
            '��ѯ��ǰҪ�����Ĳ����Ѿ��еļ���
            strSql = "Select Ӱ�����,����,a.ִ�п���ID From Ӱ�����¼ a,����ҽ����¼ b Where a.ҽ��ID=b.Id And  b.����ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˼���", lngPatientID)
            
            If mlngStudyNoBuildType = 1 Then
                '�������ɷ�ʽ=1����ֻ��ѯ�����ҵļ��
                rsTemp.Filter = "ִ�п���ID = " & mlngDetpID
            Else
                '�������ɷ�ʽ=0����ֻ��ѯ��Ӱ�����ļ��
                rsTemp.Filter = "Ӱ����� = '" & mstrModality & "'"
            End If
            
            If rsTemp.RecordCount > 1 Then
                For i = 1 To rsTemp.RecordCount
                    If InStr(str����, "||" & rsTemp("����") & "||") = 0 Then
                        str���� = str���� & rsTemp("����") & "||" 'ȷ��ÿһ�����Ŷ�����һ��||��Χ
                    End If
                    rsTemp.MoveNext
                Next i
            Else
                If InStr(str����, "||" & ItemSelected.SubItems(6) & "||") = 0 Then
                    str���� = str���� & ItemSelected.SubItems(6) & "||" 'ȷ��ÿһ�����Ŷ�����һ��||��Χ
                End If
            End If
             
            
        End If
    Next
    
    'ɾ���б��б�ѡ�е���Ŀ
    For i = lvwToBeRelate.ListItems.Count To 1 Step -1
        If lvwToBeRelate.ListItems(i).Checked = True Then
            lvwToBeRelate.ListItems.Remove i
        End If
    Next i
    
    '�����Ƿ��Զ�Ӧ�ü���
    If chkUseOldCheckNo.value = 1 And str���� <> "||" Then
        '�Ƿ��ж�����ţ�����ж����ͬ�ļ��ţ�����ʾ�û�ѡ��
        arr���� = Split(str����, "||")
        If UBound(arr����) > 2 Then
            '�ж�����ţ���ʾ�û��Լ�ѡ��
            For i = 1 To UBound(arr����) - 1
                strReturn = strReturn & i & "----" & arr����(i) & vbCrLf
            Next i
            strReturn = InputBox("���ι�����ʹ���˶������" & vbCrLf & "��������ѡ������һ������" & vbCrLf & vbCrLf _
                        & strReturn & vbCrLf & "���������Ч��ţ���ʾ��Ӧ���κ�һ�����š�", "ѡ�����", "1")
            If Val(strReturn) >= 1 And Val(strReturn) <= UBound(arr����) - 1 Then
                strReturn = arr����(Val(strReturn))
                Call subSetCheckNo(mlngOrderID, strReturn)
            End If
        Else
            'ֻ��һ�����ţ�ֱ���޸�
            strReturn = arr����(1)
            Call subSetCheckNo(mlngOrderID, strReturn)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetRelating(lngPatientID As Long)
'���ù���
'������ lngPatientID -- ��Ҫ�������Ĳ���ID
    Dim strSql As String
    
    On Error GoTo err
    
    strSql = "ZL_Ӱ���������(" & mlngOrderID & "," & mlngPatietID & "," & lngPatientID & ")"
    zlDatabase.ExecuteProcedure strSql, "��������"
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DeleteRelating(lngPatientID As Long)
'ȡ������
'������lngPatientID -- ��Ҫ��ȡ�������Ĳ���ID
    Dim strSql As String

    On Error GoTo err
        
    strSql = "ZL_Ӱ��ȡ����������(" & mlngOrderID & "," & lngPatientID & ")"
    zlDatabase.ExecuteProcedure strSql, "��������"
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim strRegPath As String
    
    strRegPath = "����ģ��\" & App.ProductName & "\frmReferencePatient"
    
    chkUseOldCheckNo.value = Val(GetSetting("ZLSOFT", strRegPath, "Ӧ�ù������˵ļ���", 0))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    strRegPath = "����ģ��\" & App.ProductName & "\frmReferencePatient"
    Call SaveSetting("ZLSOFT", strRegPath, "Ӧ�ù������˵ļ���", chkUseOldCheckNo.value)
End Sub

Private Sub lvwToBeRelate_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '���ù���,����Ƿ��и�����������Ŀ��ͬʱҲѡ����
    Dim i As Integer
    
    If mblnInCheck = True Then Exit Sub
    
    mblnInCheck = True
    For i = 1 To lvwToBeRelate.ListItems.Count
        If lvwToBeRelate.ListItems(i).SubItems(12) <> "" And lvwToBeRelate.ListItems(i).SubItems(12) = Item.SubItems(12) Then
            lvwToBeRelate.ListItems(i).Checked = Item.Checked
        End If
    Next i
    mblnInCheck = False
End Sub

Private Sub lvwToBeRelate_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '��ʾ��ǰ���˶�Ӧ�����м��
    Dim strSql As String
    Dim rsStudies As ADODB.Recordset
    Dim lngPatientID As Long
    
    On Error GoTo err
    
    lngPatientID = Mid(Item.Key, 2)
    strSql = "Select Ӱ�����,ҽ��ID,��������,ҽ������,����,b.ִ�п���ID From Ӱ�����¼ a,����ҽ����¼ b Where a.ҽ��ID=b.Id And b.����id=[1]  And b.���id Is Null order by ��������"
    Set rsStudies = zlDatabase.OpenSQLRecord(strSql, "��ȡ���˼����Ϣ", lngPatientID)
    
    If mlngStudyNoBuildType = 1 Then
        '�������ɷ�ʽ=1����ֻ��ѯ�����ҵļ��
        rsStudies.Filter = "ִ�п���ID = " & mlngDetpID
    Else
        '�������ɷ�ʽ=0����ֻ��ѯ��Ӱ�����ļ��
        rsStudies.Filter = "Ӱ����� = '" & mstrModality & "'"
    End If
    
    Call FillStudies(rsStudies)
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
    optFilter(Index).value = True
End Sub

Private Sub txtFilter_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdQuery_Click
    End If
End Sub

Private Sub FillStudies(rsStudies As ADODB.Recordset)
    '������б�
    Dim tmpItem As MSComctlLib.ListItem
    Dim i As Integer
    
    On Error GoTo err
    lvwStudies.ListItems.Clear
    i = 1
    
    While rsStudies.EOF = False
        Set tmpItem = lvwStudies.ListItems.Add(, "_" & rsStudies("ҽ��ID"), i)
        tmpItem.SubItems(1) = Nvl(rsStudies("Ӱ�����"))
        tmpItem.SubItems(2) = Nvl(rsStudies("����"))
        tmpItem.SubItems(3) = Nvl(rsStudies("��������"))
        tmpItem.SubItems(4) = Nvl(rsStudies("ҽ������"))
        rsStudies.MoveNext
        i = i + 1
    Wend
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subSetCheckNo(lngҽ��ID As Long, str���� As String)
'------------------------------------------------
'���ܣ������µļ���
'������ lngҽ��ID--ҽ��ID
'       str���� -- �µļ���
'���أ���
'----------------------------------------------
    Dim strSql As String
    
    On Error GoTo err
    
    strSql = "Zl_Ӱ�����_����( " & lngҽ��ID & "," & str���� & ")"
    zlDatabase.ExecuteProcedure strSql, "�����µļ���"
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
