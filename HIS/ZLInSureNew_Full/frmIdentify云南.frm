VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�����˾�������ѡ��"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fra�������� 
      Caption         =   "��������"
      Height          =   5145
      Left            =   210
      TabIndex        =   3
      Top             =   1440
      Width           =   4305
      Begin MSComctlLib.ListView lvw���� 
         Height          =   3975
         Left            =   270
         TabIndex        =   6
         Top             =   1020
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   7011
         View            =   1
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils32"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "����(&B)"
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   540
         Width           =   1215
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "��ͨ(&A)"
         Height          =   285
         Index           =   0
         Left            =   1380
         TabIndex        =   4
         Top             =   540
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   1
         Left            =   270
         Picture         =   "frmIdentify����.frx":000C
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.Frame fra������ 
      Caption         =   "ҽ��������"
      Height          =   1125
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   4305
      Begin VB.OptionButton optҽ���� 
         Caption         =   "�ſ�(&2)"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   2910
         TabIndex        =   2
         Top             =   570
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optҽ���� 
         Caption         =   "IC��(&1)"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   1410
         TabIndex        =   1
         Top             =   570
         Width           =   1215
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   0
         Left            =   300
         Picture         =   "frmIdentify����.frx":0E4E
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   405
      Left            =   4770
      TabIndex        =   8
      Top             =   870
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   4770
      TabIndex        =   7
      Top             =   270
      Width           =   1305
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   4980
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIdentify����.frx":1C90
            Key             =   "Disease"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum card����
    cardIC = 0
    card�ſ� = 1
End Enum

Dim mint���� As Integer
Dim mstr������ As String
Dim mstr�������� As String
Dim mlng����ID As Long
Dim mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If opt����(1).Value = True Then
        If lvw����.ListItems.Count = 0 Then
            MsgBox "����ҽ�����ֹ�����������������ϿɵĲ��֡�", vbInformation, gstrSysName
            Exit Sub
        End If
    
        If lvw����.SelectedItem Is Nothing Then
            MsgBox "��ѡ�񼲲����͡�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    mstr������ = IIf(optҽ����(cardIC).Value = True, cardIC, card�ſ�)
    If opt����(0).Value = True Then
        mlng����ID = 0
        mstr�������� = ""
    Else
        mlng����ID = Mid(lvw����.SelectedItem.Key, 2)
        mstr�������� = lvw����.SelectedItem.SubItems(1)
    End If
    
    '����ʹ�õĿ�����
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", mstr������
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKey1 Or KeyCode = vbKeyF1 Or KeyCode = vbKeyNumpad1 Then
        optҽ����(cardIC) = 0
    ElseIf KeyCode = vbKey2 Or KeyCode = vbKeyF2 Or KeyCode = vbKeyNumpad2 Then
        optҽ����(card�ſ�) = 1
    End If
End Sub

Public Function GetIdentifyMode(ByVal intInsure As Integer, ByVal bytType As Byte, str������ As String, lng����ID As Long, str�������� As String) As Boolean
'���ܣ���������֤��ģʽ
'������bytType     0-���1-סԺ��2-�������֤
'      str������   0-IC��,1-�ſ�
'      lng����ID   0-��ͨ����,����Ϊ��������
'���أ��ɹ�ΪTrue
    Dim bln������� As Boolean
    Dim rsTemp As New ADODB.Recordset, lst As ListItem
    
    mblnOK = False
    mint���� = intInsure
    mstr������ = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "������", "1") 'ȱʡΪ�ſ�
    mlng����ID = 0
    mstr�������� = ""
    
    '����ע�����Ϣ����ǰһ��ʹ�õĿ�����
    optҽ����(IIf(mstr������ = "0", 0, 1)).Value = True
    
    '���������֤���ͣ���ʾ������Ϣ
'    If bytType = 0 Or bytType = 1 Then
        '������סԺ������Ҫѡ�񼲲�
'        gstrSQL = "select ����ֵ from ���ղ��� where ����=" & mint���� & " and ����=0 and ������='֧�����Բ������ֲ�'"
'        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'        If rsTemp.EOF = False Then
'            If rsTemp("����ֵ") = "1" Then
                'Ҫ�������ⲡѡ��
                bln������� = True
'            End If
'        End If
'    End If
    
    If bln������� = False Then
        '������ʹ��������Ŀǰֻ����ͨ�ʻ���֤ʱ���������
        opt����(1).Enabled = False
    Else
        '��Ժ֧��ʹ�����ⲡ��
        If bytType = 0 Then
            '���ʹ�����Բ������ֲ�
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.��� in (0,1,2) and A.����=" & mint����
        Else
            'סԺ��ʹ����ͨ��
            'Modified by ZYB 2004-10-12 ����
            '-------------------------------
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.��� in (0,1,2) and A.����=" & mint���� & IIf(mint���� = TYPE_���Ͻ�ˮ, "", " And A.���� IN ('0094','0093')")
        End If
        
        Call OpenRecordset(rsTemp, "ҽ�������֤")
        Do Until rsTemp.EOF
            Set lst = lvw����.ListItems.Add(, "K" & rsTemp("ID"), rsTemp("����"), "Disease", "Disease")
            lst.SubItems(1) = rsTemp("����")
            
            rsTemp.MoveNext
        Loop
    End If
    
    frmIdentify����.Show vbModal
    GetIdentifyMode = mblnOK
    If mblnOK = True Then
        str������ = mstr������
        If mint���� = TYPE_���Ͻ�ˮ Then str������ = 3
        lng����ID = mlng����ID
        str�������� = mstr��������
    End If
End Function

Private Sub Form_Load()
    If mint���� = TYPE_���Ͻ�ˮ Then
        optҽ����(1).Visible = False
        optҽ����(0).Caption = "��ˮ��"
    End If
    optҽ����(0).Value = True
End Sub

Private Sub fra������_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub lvw����_DblClick()
'���Ĳ���Ա���������
'    Call cmdOK_Click
End Sub

Private Sub lvw����_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvw����.Drag 0
End Sub

Private Sub lvw����_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If Not lvw����.HitTest(x, y) Is Nothing Then
            lvw����.Drag 1
        End If
    End If
End Sub

Private Sub opt����_Click(Index As Integer)
    lvw����.Enabled = (opt����(1).Value = True)
    
    If lvw����.Enabled = False Then
        lvw����.BackColor = &H8000000F '��ť����
    Else
        lvw����.BackColor = &H80000005 '���ڱ���
    End If
End Sub

Private Sub optҽ����_Click(Index As Integer)

End Sub
