VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain_�������󲡰��ӿ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������ϴ�"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmMain_�������󲡰��ӿ�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7125
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmd����ϴ���־ 
      Caption         =   "���(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5880
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "���ָ�����˵Ĳ��������ϴ���־"
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����(&M)"
      Height          =   350
      Left            =   5880
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CheckBox chkȫѡ 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ȫѡ"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5010
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1590
      Width           =   675
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3270
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain_�������󲡰��ӿ�.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd�ϴ� 
      Caption         =   "�ϴ�(&U)"
      Height          =   350
      Left            =   5880
      TabIndex        =   13
      Top             =   780
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw�����嵥 
      Height          =   2265
      Left            =   150
      TabIndex        =   12
      Top             =   1800
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���˱��"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "סԺ��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "��Ժ����"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�ϴ�"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.CommandButton CDM���� 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   5880
      TabIndex        =   9
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5880
      TabIndex        =   14
      Top             =   3630
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������������(&S)"
      Height          =   1425
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5565
      Begin VB.CheckBox chkδ�ϴ� 
         Caption         =   "����ʾδ�ϴ�����"
         Height          =   255
         Left            =   1020
         TabIndex        =   17
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2205
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         Left            =   4080
         TabIndex        =   8
         Top             =   690
         Width           =   1275
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1470
         TabIndex        =   6
         Top             =   690
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   300
         Left            =   1470
         TabIndex        =   2
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   113311747
         CurrentDate     =   39071
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   300
         Left            =   4080
         TabIndex        =   4
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   113311747
         CurrentDate     =   39071
      End
      Begin VB.Label lblסԺ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3180
         TabIndex        =   7
         Top             =   750
         Width           =   810
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   750
         TabIndex        =   5
         Top             =   750
         Width           =   630
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   990
      End
      Begin VB.Label lbl��ʼ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Label lbl�����嵥 
      BackColor       =   &H00C0C0C0&
      Caption         =   "�����嵥"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   1590
      Width           =   5565
   End
End
Attribute VB_Name = "frmMain_�������󲡰��ӿ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const gverControl = 99  ' 0-��֧�ֶ�̬ҽ��(9.19��ǰ),1-֧�ֶ�̬ҽ���޸��Ӳ���(9.22��ǰ) , _

Private strSQL As String

Private RSPATIENT As New ADODB.Recordset        '��ȡ���������Ĳ���
Private RSREC As New ADODB.Recordset

Private Type TRECORD_INFO
    C1ͳ������ As String
    C2ҽ�ƻ������ As String
    C3סԺ�� As String
    C4�շѲ���Ա As String
    C5���ʽ As String
    C6����סԺ���� As Integer
    C7������� As String
    C8���˱�� As String
    C9���� As String
    C10�Ա� As String
    C11�������� As String
    C12���� As String
    C13ְҵ As String
    C14������ As String
    C15���� As String
    C16���� As String
    C17���֤�� As String
    C18������λ As String
    C19��λ��ַ As String
    C20��λ�绰 As String
    C21��λ�������� As String
    C22���ڵ�ַ As String
    C23�������� As String
    C24��ϵ�� As String
    C25�벡�˹�ϵ As String
    C26��ϵ��ַ As String
    C27��ϵ�绰 As String
    C28��Ժ���� As String
    C29��Ժ���� As String
    C30��Ժ���� As String
    C31ת�ƿƱ� As String
    C32��Ժ���� As String
    C33��Ժ���� As String
    C34��Ժ���� As String
    C35��Ժ���� As String
    C36��Ժ��ȷ������ As String
    C37����ҩ�� As String
    C38HBSAG As String
    C39HCV_AB As String
    C40HIV_AB As String
    C41�������Ժ As Integer
    C42��Ժ���Ժ As Integer
    C43��ǰ������ As Integer
    C44�ٴ��벡�� As Integer
    C45�����벡�� As Integer
    C46���ȴ��� As Integer
    C47���ȳɹ����� As Integer
    C48������ As String
    C49����ҽʦ As String
    C50����ҽʦ As String
    C51סԺҽʦ As String
    C52����ҽʦ As String
    C53�о���ʵϰҽʦ As String
    C54ʵϰҽʦ As String
    C55����Ա As String
    C56�������� As String
    C57�ʿ�ҽʦ As String
    C58�ʿػ�ʦ As String
    C59�������� As String
    C60ʬ���־ As String
    C61�������Ƽ�����Ϊ��Ժ��һ�� As Integer
    C62�����־ As Integer
    C63�������� As Integer
    C64ʾ�̲��� As Integer
    C65Ѫ�� As Integer
    C66RH As Integer
    C67����Ѫ��Ӧ��־ As Integer
    C68�����ϸ�� As Currency
    C69����ѪС�� As Currency
    C70����Ѫ�� As Currency
    C71ȫѪ As Currency
    C72���� As Currency
    C73������ As String
    C74����ʱ�� As String
End Type

Private Sub READPATIENTS()
    Dim lvwItem As ListItem
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHand
    lvw�����嵥.ListItems.Clear
    Me.chkȫѡ.Value = 0
    
    '�������
    strSQL = " AND B.��Ժ���� BETWEEN TO_DATE('" & Format(dtp��ʼ����.Value, "YYYY-MM-DD") & " 00:00:00','YYYY-MM-DD HH24:MI:SS')" & _
    " AND TO_DATE('" & Format(dtp��������.Value, "YYYY-MM-DD") & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    If Trim(Me.txt����.Text) <> "" Then
        strSQL = strSQL & " AND A.���� LIKE '" & Trim(Me.txt����.Text) & "%'"
    End If
    If Trim(Me.txtסԺ��.Text) <> "" Then
        strSQL = strSQL & " AND A.סԺ�� LIKE '" & Trim(Me.txtסԺ��.Text) & "%'"
    End If
    
    '��ȡ���������ĳ�Ժ����
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    If Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34 Then
        strSQL = " SELECT A.����ID,B.��ҳID,C.ҽ����,A.����,A.סԺ��,B.��Ժ����,B.�Ƿ��ϴ� " & _
             " FROM ������Ϣ A,������ҳ B,�����ʻ� C" & _
             " WHERE A.����ID=B.����ID AND A.��ҳID=B.��ҳID " & _
             " AND A.����ID=C.����ID AND C.����=" & TYPE_�������� & IIf(chkδ�ϴ�.Value = 1, " AND NVL(B.�Ƿ��ϴ�,0)=0 ", "") & strSQL
    Else
        strSQL = " SELECT A.����ID,B.��ҳID,C.ҽ����,A.����,A.סԺ��,B.��Ժ����,B.�Ƿ��ϴ� " & _
             " FROM ������Ϣ A,������ҳ B,�����ʻ� C" & _
             " WHERE A.����ID=B.����ID AND A.סԺ����=B.��ҳID " & _
             " AND A.����ID=C.����ID AND C.����=" & TYPE_�������� & IIf(chkδ�ϴ�.Value = 1, " AND NVL(B.�Ƿ��ϴ�,0)=0 ", "") & strSQL
        
    End If
    Call OpenRecordset(RSPATIENT, "��ȡ���������Ĳ���", strSQL)
    With RSPATIENT
        Do While Not .EOF
            Set lvwItem = lvw�����嵥.ListItems.Add(, "K" & .AbsolutePosition, Nvl(!ҽ����), 1)
            lvwItem.SubItems(1) = Nvl(!סԺ��) & "_" & !��ҳID
            lvwItem.SubItems(2) = Nvl(!����)
            lvwItem.SubItems(3) = Format(!��Ժ����, "YYYY-MM-DD")
            lvwItem.SubItems(4) = IIf(Nvl(!�Ƿ��ϴ�, 0) = 0, "��", "��")
            lvwItem.Tag = !����ID & "_" & !��ҳID
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    MsgBox Err.Description, vbInformation, gstrSysName
    Resume
End Sub

Private Function UPLOADREC(ByVal lng����ID As Long, ByVal lng��ҳID As Long, STRERR As String) As Boolean
    '----------------------------------------------------------------
    '�ָ�����,�ֶζ�ȡ��Ӧ�����ݲ�����
    '----------------------------------------------------------------
    '���ݴ���Ĳ��˱�ʶ�ϴ�������Ϣ
    Dim arr������Ŀ
    Dim STR������Ŀ As String
    Dim STR��Ժ��� As String
    Dim RECORD_INFO As TRECORD_INFO
    Dim bln34 As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHand
    
    STR������Ŀ = GetSetting("ZLSOFT", "˽��ģ��\FRMSET", "������Ŀ", "")
    arr������Ŀ = Split(STR������Ŀ, "|")
    
    gcn����.BeginTrans
    
    '----------------------------------------------------------------
    '1��RECORD_INFO
    RECORD_INFO.C1ͳ������ = gstrҽ����������
    RECORD_INFO.C2ҽ�ƻ������ = Trim(gstrҽԺ����)
    'ȡҽ������
    strSQL = " SELECT ҽ����" & _
             " FROM �����ʻ�" & _
             " WHERE ����=" & TYPE_�������� & " AND ����ID=" & lng����ID
    Call OpenRecordset(RSREC, "ȡҽ������", strSQL)
    RECORD_INFO.C8���˱�� = RSREC!ҽ����
    
    'ȡ���˻�����Ϣ
    gstrSQL = "Select �汾�� From zlSystems Where ��� = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS�汾��")
    bln34 = Split(rsTmp!�汾��, ".")(0) = 10 And Split(rsTmp!�汾��, ".")(1) >= 34
    
    #If gverControl < 6 Then
        If bln34 Then
            strSQL = " SELECT A.סԺ��,B.ҽ�Ƹ��ʽ,B.��ҳID,B.������,A.����,A.�Ա�,A.��������,B.����״��,B.ְҵ,A.�����ص�," & _
                "        H.���� AS ����,B.����,A.���֤��,A.������λ,B.��λ��ַ,B.��λ�绰,B.��λ�ʱ�,B.��ͥ��ַ,B.�����ʱ�," & _
                "        B.��ϵ������,B.��ϵ�˹�ϵ,B.��ϵ�˵�ַ,B.��ϵ�˵绰,B.��Ժ����,D.���� AS ��Ժ����,E.���� AS ��Ժ����," & _
                "        B.��Ժ����,F.���� AS ��Ժ����,B.��Ժ����,B.ȷ������,B.���ȴ���,B.�ɹ�����,B.��Ժ��ʽ," & _
                "        B.��ĿԱ����,NVL(B.��Ŀ����,SYSDATE) AS ��Ŀ����,B.ʬ���־,B.�����־,B.��������,B.Ѫ��,B.סԺҽʦ" & _
                " FROM ������Ϣ A,������ҳ B,��Լ��λ C,���ű� D,���ű� E,���ű� F,���� H" & _
                " WHERE A.����ID=B.����ID AND A.��ҳID=B.��ҳID AND A.��ͬ��λID=C.ID(+)" & _
                " AND B.��Ժ����ID=D.ID(+) AND B.��Ժ����ID=E.ID(+) AND B.��Ժ����ID=F.ID(+) " & _
                " AND A.����=H.���� AND B.����ID=[1] AND B.��ҳID=[2]"
        Else
            strSQL = " SELECT A.סԺ��,B.ҽ�Ƹ��ʽ,B.��ҳID,B.������,A.����,A.�Ա�,A.��������,B.����״��,B.ְҵ,A.�����ص�," & _
                "        H.���� AS ����,B.����,A.���֤��,A.������λ,B.��λ��ַ,B.��λ�绰,B.��λ�ʱ�,B.��ͥ��ַ,B.�����ʱ�," & _
                "        B.��ϵ������,B.��ϵ�˹�ϵ,B.��ϵ�˵�ַ,B.��ϵ�˵绰,B.��Ժ����,D.���� AS ��Ժ����,E.���� AS ��Ժ����," & _
                "        B.��Ժ����,F.���� AS ��Ժ����,B.��Ժ����,B.ȷ������,B.���ȴ���,B.�ɹ�����,B.��Ժ��ʽ," & _
                "        B.��ĿԱ����,NVL(B.��Ŀ����,SYSDATE) AS ��Ŀ����,B.ʬ���־,B.�����־,B.��������,B.Ѫ��,B.סԺҽʦ" & _
                " FROM ������Ϣ A,������ҳ B,��Լ��λ C,���ű� D,���ű� E,���ű� F,���� H" & _
                " WHERE A.����ID=B.����ID AND A.סԺ����=B.��ҳID AND A.��ͬ��λID=C.ID(+)" & _
                " AND B.��Ժ����ID=D.ID(+) AND B.��Ժ����ID=E.ID(+) AND B.��Ժ����ID=F.ID(+) " & _
                " AND A.����=H.���� AND B.����ID=[1] AND B.��ҳID=[2]"
        End If
    #Else
        If bln34 Then
            strSQL = " SELECT A.סԺ��,B.ҽ�Ƹ��ʽ,B.��ҳID,B.������,A.����,A.�Ա�,A.��������,B.����״��,B.ְҵ,A.�����ص�," & _
                "        H.���� AS ����,B.����,A.���֤��,A.������λ,B.��λ��ַ,B.��λ�绰,B.��λ�ʱ�,B.��ͥ��ַ,B.��ͥ��ַ�ʱ� As �����ʱ�," & _
                "        B.��ϵ������,B.��ϵ�˹�ϵ,B.��ϵ�˵�ַ,B.��ϵ�˵绰,B.��Ժ����,D.���� AS ��Ժ����,E.���� AS ��Ժ����," & _
                "        B.��Ժ����,F.���� AS ��Ժ����,B.��Ժ����,B.ȷ������,B.���ȴ���,B.�ɹ�����,B.��Ժ��ʽ," & _
                "        B.��ĿԱ����,NVL(B.��Ŀ����,SYSDATE) AS ��Ŀ����,B.ʬ���־,B.�����־,B.��������,B.Ѫ��,B.סԺҽʦ" & _
                " FROM ������Ϣ A,������ҳ B,��Լ��λ C,���ű� D,���ű� E,���ű� F,���� H" & _
                " WHERE A.����ID=B.����ID AND A.��ҳID=B.��ҳID AND A.��ͬ��λID=C.ID(+)" & _
                " AND B.��Ժ����ID=D.ID(+) AND B.��Ժ����ID=E.ID(+) AND B.��Ժ����ID=F.ID(+) " & _
                " AND A.����=H.���� AND B.����ID=[1] AND B.��ҳID=[2]"
        Else
            strSQL = " SELECT A.סԺ��,B.ҽ�Ƹ��ʽ,B.��ҳID,B.������,A.����,A.�Ա�,A.��������,B.����״��,B.ְҵ,A.�����ص�," & _
                "        H.���� AS ����,B.����,A.���֤��,A.������λ,B.��λ��ַ,B.��λ�绰,B.��λ�ʱ�,B.��ͥ��ַ,B.��ͥ��ַ�ʱ� As �����ʱ�," & _
                "        B.��ϵ������,B.��ϵ�˹�ϵ,B.��ϵ�˵�ַ,B.��ϵ�˵绰,B.��Ժ����,D.���� AS ��Ժ����,E.���� AS ��Ժ����," & _
                "        B.��Ժ����,F.���� AS ��Ժ����,B.��Ժ����,B.ȷ������,B.���ȴ���,B.�ɹ�����,B.��Ժ��ʽ," & _
                "        B.��ĿԱ����,NVL(B.��Ŀ����,SYSDATE) AS ��Ŀ����,B.ʬ���־,B.�����־,B.��������,B.Ѫ��,B.סԺҽʦ" & _
                " FROM ������Ϣ A,������ҳ B,��Լ��λ C,���ű� D,���ű� E,���ű� F,���� H" & _
                " WHERE A.����ID=B.����ID AND A.סԺ����=B.��ҳID AND A.��ͬ��λID=C.ID(+)" & _
                " AND B.��Ժ����ID=D.ID(+) AND B.��Ժ����ID=E.ID(+) AND B.��Ժ����ID=F.ID(+) " & _
                " AND A.����=H.���� AND B.����ID=[1] AND B.��ҳID=[2]"
        End If
    #End If
    Call OpenRecordset(RSREC, "ȡ���˻�����Ϣ", strSQL)
    STR��Ժ��� = TRANDATA("��Ժ���", Nvl(RSREC!��Ժ��ʽ))
    With RSREC
        RECORD_INFO.C31ת�ƿƱ� = "��"
        
        RECORD_INFO.C3סԺ�� = Nvl(!סԺ��) & "_" & !��ҳID
        RECORD_INFO.C5���ʽ = TRANDATA("ҽ�Ƹ��ʽ", !ҽ�Ƹ��ʽ)
        RECORD_INFO.C6����סԺ���� = !��ҳID
        RECORD_INFO.C7������� = Nvl(!סԺ��)
        RECORD_INFO.C9���� = !����
        RECORD_INFO.C10�Ա� = TRANDATA("�Ա�", !�Ա�)
        RECORD_INFO.C11�������� = Format(!��������, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C12���� = TRANDATA("����", Nvl(!����״��, "δ��"))
        RECORD_INFO.C13ְҵ = ToVarchar(Nvl(!ְҵ), 20)
        RECORD_INFO.C14������ = Nvl(!�����ص�)
        RECORD_INFO.C15���� = !����
        RECORD_INFO.C16���� = 86 '!����
        RECORD_INFO.C17���֤�� = Nvl(!���֤��)
        RECORD_INFO.C18������λ = Nvl(!������λ)
        RECORD_INFO.C19��λ��ַ = Nvl(!��λ��ַ)
        RECORD_INFO.C20��λ�绰 = Nvl(!��λ�绰)
        RECORD_INFO.C21��λ�������� = Nvl(!��λ�ʱ�)
        RECORD_INFO.C22���ڵ�ַ = Nvl(!��ͥ��ַ)
        RECORD_INFO.C23�������� = Nvl(!�����ʱ�)
        RECORD_INFO.C24��ϵ�� = Nvl(!��ϵ������)
        RECORD_INFO.C25�벡�˹�ϵ = TRANDATA("�벡�˹�ϵ", Nvl(!��ϵ�˹�ϵ))
        RECORD_INFO.C26��ϵ��ַ = Nvl(!��ϵ�˵�ַ)
        RECORD_INFO.C27��ϵ�绰 = Nvl(!��ϵ�˵绰)
        RECORD_INFO.C28��Ժ���� = Format(!��Ժ����, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C29��Ժ���� = !��Ժ����
        RECORD_INFO.C30��Ժ���� = Nvl(!��Ժ����)
        RECORD_INFO.C32��Ժ���� = Format(!��Ժ����, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C33��Ժ���� = !��Ժ����
        RECORD_INFO.C34��Ժ���� = !��Ժ����
        RECORD_INFO.C35��Ժ���� = TRANDATA("��Ժ����", Nvl(!��Ժ����))
        RECORD_INFO.C36��Ժ��ȷ������ = Format(!ȷ������, "YYYY-MM-DD HH:MM:SS")
        RECORD_INFO.C46���ȴ��� = Nvl(!���ȴ���, 0)
        RECORD_INFO.C47���ȳɹ����� = Nvl(!�ɹ�����, 0)
        RECORD_INFO.C51סԺҽʦ = Nvl(!סԺҽʦ)
        RECORD_INFO.C55����Ա = Nvl(!��ĿԱ����)
        RECORD_INFO.C60ʬ���־ = TRANDATA("ʬ���־", Nvl(!ʬ���־, "��"))
        RECORD_INFO.C62�����־ = TRANDATA("�����־", Nvl(!�����־))
        RECORD_INFO.C63�������� = Nvl(!��������, 0)
        RECORD_INFO.C65Ѫ�� = TRANDATA("Ѫ��", Nvl(!Ѫ��))
        RECORD_INFO.C73������ = ToVarchar(Nvl(!��ĿԱ����), 20)
        RECORD_INFO.C74����ʱ�� = Format(Now, "YYYY-MM-DD HH:MM:SS")
    End With
    
    'ȡ������ҳ�ӱ�
    Dim STR��Ϣֵ As String
    strSQL = "SELECT UPPER(��Ϣ��) AS ��Ϣ��,��Ϣֵ FROM ������ҳ�ӱ� WHERE ����ID=" & lng����ID & " AND ��ҳID=" & lng��ҳID
    Call OpenRecordset(RSREC, "ȡ������ҳ�ӱ�", strSQL)
    With RSREC
        Do While Not .EOF
            STR��Ϣֵ = Nvl(!��Ϣֵ)
            Select Case !��Ϣ��
            Case "HBSAG"
                RECORD_INFO.C38HBSAG = TRANDATA("HBSAG", STR��Ϣֵ)
            Case "HCV-AB"
                RECORD_INFO.C39HCV_AB = TRANDATA("HCV_AB", STR��Ϣֵ)
            Case "HIV-AB"
                RECORD_INFO.C40HIV_AB = TRANDATA("HIV_AB", STR��Ϣֵ)
            Case "������"
                RECORD_INFO.C48������ = STR��Ϣֵ
            Case "����ҽʦ"
                RECORD_INFO.C49����ҽʦ = STR��Ϣֵ
            Case "����ҽʦ"
                RECORD_INFO.C50����ҽʦ = STR��Ϣֵ
            Case "����ҽʦ"
                RECORD_INFO.C52����ҽʦ = STR��Ϣֵ
            Case "�о���ʵϰҽʦ"
                RECORD_INFO.C53�о���ʵϰҽʦ = STR��Ϣֵ
            Case "ʵϰҽʦ"
                RECORD_INFO.C54ʵϰҽʦ = STR��Ϣֵ
            Case "��������"
                RECORD_INFO.C56�������� = TRANDATA("��������", STR��Ϣֵ)
            Case "�ʿ�ҽʦ"
                RECORD_INFO.C57�ʿ�ҽʦ = STR��Ϣֵ
            Case "�ʿػ�ʿ"
                RECORD_INFO.C58�ʿػ�ʦ = STR��Ϣֵ
            Case "����"
                RECORD_INFO.C61�������Ƽ�����Ϊ��Ժ��һ�� = TRANDATA("����", STR��Ϣֵ)
            Case "ʾ�̲���"
                RECORD_INFO.C64ʾ�̲��� = TRANDATA("ʾ�̲���", STR��Ϣֵ)
            Case "RH"
                RECORD_INFO.C66RH = TRANDATA("RH", STR��Ϣֵ)
            Case "��Ѫ��Ӧ"
                RECORD_INFO.C67����Ѫ��Ӧ��־ = TRANDATA("��Ѫ��Ӧ", STR��Ϣֵ)
            Case "���ϸ��"
                RECORD_INFO.C68�����ϸ�� = Val(STR��Ϣֵ)
            Case "��ѪС��"
                RECORD_INFO.C69����ѪС�� = Val(STR��Ϣֵ)
            Case "��Ѫ��"
                RECORD_INFO.C70����Ѫ�� = Val(STR��Ϣֵ)
            Case "��ȫѪ"
                RECORD_INFO.C71ȫѪ = Val(STR��Ϣֵ)
            Case "������"
                RECORD_INFO.C72���� = Val(STR��Ϣֵ)
            End Select
            .MoveNext
        Loop
    End With
    
    '��ѡһ�ֹ���ҩ��
    strSQL = " SELECT ����ҩ�� FROM ���˹���ҩ�� WHERE ����ID=" & lng����ID
    Call OpenRecordset(RSREC, "��ѡһ�ֹ���ҩ��", strSQL)
    Do While Not RSREC.EOF
        RECORD_INFO.C37����ҩ�� = RECORD_INFO.C37����ҩ�� & "," & Nvl(RSREC!����ҩ��)
        RSREC.MoveNext
    Loop
    RECORD_INFO.C37����ҩ�� = Mid(RECORD_INFO.C37����ҩ��, 2)
    RECORD_INFO.C37����ҩ�� = ToVarchar(RECORD_INFO.C37����ҩ��, 50)
'    'ȡ�������ֽ�����Ӳ�����ҳ�ӱ��ж�ȡ��ֻҪ��д�˲����Ķ��������ݣ�
'    STRSQL = "SELECT �ȼ� FROM �������ֽ�� WHERE ����ID=" & LNG����ID & " AND ��ҳID=" & LNG��ҳID
'    CALL OPENRECORD(RSREC, STRSQL, "ȡ�������ֽ��")
'    IF RSREC.RECORDCOUNT <> 0 THEN
'        RECORD_INFO.C56�������� = RSREC!�ȼ�
'    END IF
    'ȡ��������
    strSQL = "SELECT ����Ա����,�շ�ʱ�� FROM ���˽��ʼ�¼ WHERE ID = (SELECT MAX(ID) FROM ���˽��ʼ�¼ WHERE ����ID=" & lng����ID & " AND ��¼״̬=1)"
    Call OpenRecordset(RSREC, "ȡ��������", strSQL)
    If RSREC.RecordCount <> 0 Then
'        RECORD_INFO.C4�շѲ���Ա = RSREC!����Ա����
        RECORD_INFO.C59�������� = Format(RSREC!�շ�ʱ��, "YYYY-MM-DD HH:MM:SS")
    End If
    'ȡ������
    strSQL = "SELECT ��������,NVL(�������,0) AS ������� FROM ��Ϸ������ WHERE ����ID=" & lng����ID & " AND ��ҳID=" & lng��ҳID
    Call OpenRecordset(RSREC, "ȡ������", strSQL)
    With RSREC
        Do While Not .EOF
            Select Case !��������
            Case 1  '�������Ժ
                RECORD_INFO.C41�������Ժ = !�������
            Case 2  '��Ժ���Ժ
                RECORD_INFO.C42��Ժ���Ժ = !�������
            Case 3  '�����벡��
                RECORD_INFO.C45�����벡�� = !�������
            Case 4  '�ٴ��벡��
                RECORD_INFO.C44�ٴ��벡�� = !�������
            Case 6  '��ǰ������
                RECORD_INFO.C43��ǰ������ = !�������
            End Select
            .MoveNext
        Loop
    End With
    
    '�������ݱ�:MEDICAL_RECORD_INFO
'    gcn����.Execute " DELETE MEDICAL_RECORD_INFO " & _
'                  " WHERE AREAID='" & RECORD_INFO.C1ͳ������ & "'" & _
'                  " AND PERSONAL_NUMBER='" & RECORD_INFO.C8���˱�� & "'" & _
'                  " AND RESIDENCE_NO='" & RECORD_INFO.C3סԺ�� & "'"
    strSQL = " INSERT INTO MEDICAL_RECORD_INFO" & _
         " (AREAID,HOSPITAL_NUMBER,RESIDENCE_NO,CHARGE_NUMBER,PAY_MODE,IN_COUNT,MEDICAL_RECORD_NO,PERSONAL_NUMBER, " & _
         " NAME,SEX,BIRTH_DATE,MARITAL_STATUS,STATUS,BIRTH_ADDRESS,NATIONALITY,CITIZENSHIP,IDENTITY_NUMBER, " & _
         " UNIT_NAME,UNIT_ADDRESS,UNIT_PHONE,UNIT_ZIPCODE,REGISTER_ADDRESS,REGISTER_ZIPCODE,CONTACT_PERSON, " & _
         " RELATIONSHIP,CONTACT_ADDRESS,CONTACT_PHONE,ADMISSION_DATE,ADMISSION_DEPT,IN_DEPT_ZONE,DEPT_TRANSFERED_TO, " & _
         " DISCHARGE_DATE,DISCHARGE_DEPT,OUT_DEPT_ZONE,PAT_ADM_CONDITION,DIAGNOSIS_DATE,ALERGY_DRUGS,HBsAg,HCV_Ab, " & _
         " HIV_Ab,CLINIC_INHOSPITAL,IN_OUT,BEFORE_AFTER_TREATMENT,CLINIC_PATHOLOGY,EMIT_PATHOLOGY,EMER_TREAT_TIMES,ESC_EMER_TIMES, " & _
         " DIRECTOR,DIRECTOR_DOCTOR,ATTENDING_DOCTOR,INHOSPITAL_DOCTOR,REFRESH_DOCTOR,GRADUATE_DOCTOR,INTERM,CODE_NAME, " & _
         " MEDICAL_RECORD_MASS,CONTROL_DOCTOR,CONTROL_NURSE,BAL_DATE,BODY_EXAMINE_FLAG,FIRST_FLAG,FOLLOW_FLAG,FOLLOW_TERM, " & _
         " TEACH_MR_FLAG,BLOOD_TYPE,Rh,BLOOD_TRAN_REACT_FLAG,ERYTHROCYTE,HEMOBLAST,PLASM,BLOOD,OTHER_BLOOD,HANDLE,HANDLE_DATE)" & _
         " VALUES ("
    strSQL = strSQL & _
         "'" & RECORD_INFO.C1ͳ������ & "','" & RECORD_INFO.C2ҽ�ƻ������ & "','" & RECORD_INFO.C3סԺ�� & "','" & RECORD_INFO.C4�շѲ���Ա & "'," & _
         "'" & RECORD_INFO.C5���ʽ & "'," & RECORD_INFO.C6����סԺ���� & ",'" & RECORD_INFO.C7������� & "','" & RECORD_INFO.C8���˱�� & "'," & _
         "'" & RECORD_INFO.C9���� & "','" & RECORD_INFO.C10�Ա� & "','" & RECORD_INFO.C11�������� & "','" & RECORD_INFO.C12���� & "'," & _
         "'" & RECORD_INFO.C13ְҵ & "','" & RECORD_INFO.C14������ & "','" & RECORD_INFO.C15���� & "','" & RECORD_INFO.C16���� & "'," & _
         "'" & RECORD_INFO.C17���֤�� & "','" & RECORD_INFO.C18������λ & "','" & RECORD_INFO.C19��λ��ַ & "','" & RECORD_INFO.C20��λ�绰 & "'," & _
         "'" & RECORD_INFO.C21��λ�������� & "','" & RECORD_INFO.C22���ڵ�ַ & "','" & RECORD_INFO.C23�������� & "','" & RECORD_INFO.C24��ϵ�� & "'," & _
         "'" & RECORD_INFO.C25�벡�˹�ϵ & "','" & RECORD_INFO.C26��ϵ��ַ & "','" & RECORD_INFO.C27��ϵ�绰 & "','" & RECORD_INFO.C28��Ժ���� & "'," & _
         "'" & RECORD_INFO.C29��Ժ���� & "','" & RECORD_INFO.C30��Ժ���� & "','" & RECORD_INFO.C31ת�ƿƱ� & "','" & RECORD_INFO.C32��Ժ���� & "'," & _
         "'" & RECORD_INFO.C33��Ժ���� & "','" & RECORD_INFO.C34��Ժ���� & "','" & RECORD_INFO.C35��Ժ���� & "','" & RECORD_INFO.C36��Ժ��ȷ������ & "'," & _
         "'" & RECORD_INFO.C37����ҩ�� & "','" & RECORD_INFO.C38HBSAG & "','" & RECORD_INFO.C39HCV_AB & "','" & RECORD_INFO.C40HIV_AB & "'," & _
         "'" & RECORD_INFO.C41�������Ժ & "','" & RECORD_INFO.C42��Ժ���Ժ & "','" & RECORD_INFO.C43��ǰ������ & "','" & RECORD_INFO.C44�ٴ��벡�� & "'," & _
         "'" & RECORD_INFO.C45�����벡�� & "'," & RECORD_INFO.C46���ȴ��� & "," & RECORD_INFO.C47���ȳɹ����� & ",'" & RECORD_INFO.C48������ & "'," & _
         "'" & RECORD_INFO.C49����ҽʦ & "','" & RECORD_INFO.C50����ҽʦ & "','" & RECORD_INFO.C51סԺҽʦ & "','" & RECORD_INFO.C52����ҽʦ & "',"
    strSQL = strSQL & _
         "'" & RECORD_INFO.C53�о���ʵϰҽʦ & "','" & RECORD_INFO.C54ʵϰҽʦ & "','" & RECORD_INFO.C55����Ա & "','" & RECORD_INFO.C56�������� & "'," & _
         "'" & RECORD_INFO.C57�ʿ�ҽʦ & "','" & RECORD_INFO.C58�ʿػ�ʦ & "','" & RECORD_INFO.C59�������� & "','" & RECORD_INFO.C60ʬ���־ & "'," & _
         "'" & RECORD_INFO.C61�������Ƽ�����Ϊ��Ժ��һ�� & "','" & RECORD_INFO.C62�����־ & "'," & RECORD_INFO.C63�������� & ",'" & RECORD_INFO.C64ʾ�̲��� & "'," & _
         "'" & RECORD_INFO.C65Ѫ�� & "','" & RECORD_INFO.C66RH & "','" & RECORD_INFO.C67����Ѫ��Ӧ��־ & "'," & RECORD_INFO.C68�����ϸ�� & "," & _
         "" & RECORD_INFO.C69����ѪС�� & "," & RECORD_INFO.C70����Ѫ�� & "," & RECORD_INFO.C71ȫѪ & "," & RECORD_INFO.C72���� & ",'" & RECORD_INFO.C73������ & "','" & RECORD_INFO.C74����ʱ�� & "')"
    gcn����.Execute strSQL
    
    '----------------------------------------------------------------
    '2��DIAGNOSIS
    'TODO:���ƽ�������д:ÿ����ϼ�¼��Ҫ��д���ƽ�����Ͱ���Ժ����Ĵ������д
    strSQL = " SELECT A.�������,A.��ϴ���,A.�������,B.���� AS ��������,A.�������,A.��¼��,NVL(A.��¼����,SYSDATE) AS ��¼����" & _
             " FROM ������ϼ�¼ A,��������Ŀ¼ B" & _
             " WHERE A.����ID=B.ID AND A.��¼��Դ=3 AND A.�������<8 AND A.����ID=" & lng����ID & " AND A.��ҳID=" & lng��ҳID
    Call OpenRecordset(RSREC, "����ϼ�¼", strSQL)
    With RSREC
        Do While Not .EOF
            strSQL = " INSERT INTO MR_DIAGNOSIS" & _
                     " (HOSPITAL_NUMBER,MEDICAL_RECORD_NO,IN_COUNT,DIAGNOSIS_TYPE,DIAGNOSIS_NO,ILLNESS_CODE,DIAGNOSIS_DESC,DIAGNOSIS_DATE,TREAT_RESULT,HANDLE,HANDLE_DATE)" & _
                     " VALUES (" & _
                     "'" & RECORD_INFO.C2ҽ�ƻ������ & "','" & RECORD_INFO.C7������� & "'," & RECORD_INFO.C6����סԺ���� & "," & _
                     "'" & TRANDATA("�������", !�������) & "'," & Nvl(!��ϴ���, 0) & ",'" & Nvl(!��������) & "','" & Nvl(!�������) & "'," & _
                     "'" & Format(!��¼����, "YYYY-MM-DD HH:MM:SS") & "','" & STR��Ժ��� & "','" & ToVarchar(Nvl(!��¼��), 20) & "','" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
            gcn����.Execute strSQL
            .MoveNext
        Loop
    End With
    
    '----------------------------------------------------------------
    '3��OPERATION
    strSQL = " SELECT B.����,B.����,A.�п�,A.����,A.��������,A.��������,A.����ҽʦ,A.��һ����,A.�ڶ�����,A.����ҽʦ,A.��¼��,NVL(A.��¼����,SYSDATE) AS ��¼���� " & _
             " FROM ���������¼ A ,��������Ŀ¼ B " & _
             " WHERE A.����ID=" & lng����ID & " AND A.��ҳID=" & lng��ҳID & " AND A.��������ID=B.ID"
    Call OpenRecordset(RSREC, "ȡ���������¼", strSQL)
    With RSREC
        Do While Not .EOF
            strSQL = " INSERT INTO MR_OPERATION" & _
                     " (HOSPITAL_NUMBER,MEDICAL_RECORD_NO,IN_COUNT,OPERATION_NO,OPERATION_CODE,OPERATION_DESC,WOUND_GRADE," & _
                     " HEAL,OPERATING_DATE,ANAESTHESIA_METHOD,OPERATOR,IASIST1,IASIST2,ANAESTHESIA_OPERATOR,HANDLE,HANDLE_DATE)" & _
                     " VALUES (" & _
                     "'" & RECORD_INFO.C2ҽ�ƻ������ & "','" & RECORD_INFO.C7������� & "'," & RECORD_INFO.C6����סԺ���� & "," & _
                     "" & .AbsolutePosition & ",'" & !���� & "','" & !���� & "','" & TRANDATA(Nvl(!�п�), Nvl(!����)) & "','" & TRANDATA(Nvl(!�п�), Nvl(!����)) & "'," & _
                     "'" & Format(!��������, "YYYY-MM-DD HH:MM:SS") & "','" & TRANDATA("��������", !��������) & "','" & !����ҽʦ & "'," & _
                     "'" & Nvl(!��һ����) & "','" & Nvl(!�ڶ�����) & "','" & ToVarchar(Nvl(!����ҽʦ), 20) & "','" & ToVarchar(Nvl(!��¼��), 20) & "','" & Format(Now, "YYYY-MM-DD") & "')"
            gcn����.Execute strSQL
            .MoveNext
        Loop
    End With
    
    '----------------------------------------------------------------
    '4��RECEIPT_DETAIL
    strSQL = " SELECT ������,��� FROM ���˷��� WHERE ����ID=" & lng����ID & " AND ��ҳID=" & lng��ҳID
    Call OpenRecordset(RSREC, "ȡ���˷���", strSQL)
    With RSREC
        Do While Not .EOF
            strSQL = " INSERT INTO MR_RECEIPT_DETAIL" & _
                     " (HOSPITAL_NUMBER,MEDICAL_RECORD_NO,IN_COUNT,RECEIPT_NAME,ITEM_COST,SEND_FLAG,HANDLE,HANDLE_DATE)" & _
                     " VALUES (" & _
                     "'" & RECORD_INFO.C2ҽ�ƻ������ & "','" & RECORD_INFO.C7������� & "'," & RECORD_INFO.C6����סԺ���� & "," & _
                     "'" & GET������Ŀ����(!������, arr������Ŀ) & "'," & !��� & ",0,'" & gstrUserName & "','" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "')"
            gcn����.Execute strSQL
            .MoveNext
        Loop
    End With
    
    gcn����.CommitTrans
    '�����ϴ����
    strSQL = "ZL_������ҳ_�ϴ�(" & lng����ID & "," & lng��ҳID & ",1)"
    gcnOracle.Execute strSQL, , adCmdStoredProc
    
    UPLOADREC = True
    Exit Function
errHand:
    STRERR = Err.Description
    Debug.Print "Error SQL:" & strSQL
    gcn����.RollbackTrans
End Function

Private Sub CDM����_CLICK()
    Call READPATIENTS
End Sub

Private Sub CHKȫѡ_CLICK()
    Dim BLNSEL As Boolean
    Dim LNGDO As Long, LNGMAX As Long
    
    BLNSEL = (chkȫѡ.Value = 1)
    LNGMAX = lvw�����嵥.ListItems.Count
    
    For LNGDO = 1 To LNGMAX
        lvw�����嵥.ListItems(LNGDO).Checked = BLNSEL
    Next
    chkȫѡ.Caption = IIf(BLNSEL, "ȫ��", "ȫѡ")
End Sub

Private Sub CMD����_CLICK()
'    frmSet.Show 1, Me
End Sub

Private Sub CMD����_Click()
    Unload Me
End Sub

Private Sub cmd����ϴ���־_Click()
    Dim lng����ID As Long, lng��ҳID As Long
    On Error GoTo errHand
    
    If lvw�����嵥.ListItems.Count = 0 Then Exit Sub
    If lvw�����嵥.SelectedItem Is Nothing Then Exit Sub
    lng����ID = Val(Split(lvw�����嵥.SelectedItem.Tag, "_")(0))
    lng��ҳID = Val(Split(lvw�����嵥.SelectedItem.Tag, "_")(1))
    
    If MsgBox("��ȷ��Ҫ����ò��˵Ĳ��������ϴ���־��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '�����ϴ����
    strSQL = "ZL_������ҳ_�ϴ�(" & lng����ID & "," & lng��ҳID & ",0)"
    gcnOracle.Execute strSQL, , adCmdStoredProc
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CMD�ϴ�_CLICK()
    Dim STR������Ŀ As String
    Dim LNGDO As Long, LNGMAX As Long
    Dim lng����ID As Long, lng��ҳID As Long
    Dim STR���� As String, STRERR As String
    
    STR������Ŀ = GetSetting("ZLSOFT", "˽��ģ��\FRMSET", "������Ŀ", "")
    If STR������Ŀ = "" Then
        MsgBox "�����ڲ��������н��в�����Ŀ���գ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    LNGMAX = lvw�����嵥.ListItems.Count
    For LNGDO = 1 To LNGMAX
        If lvw�����嵥.ListItems(LNGDO).Checked Then
            STR���� = lvw�����嵥.ListItems(LNGDO).SubItems(2)
            lng����ID = Split(lvw�����嵥.ListItems(LNGDO).Tag, "_")(0)
            lng��ҳID = Split(lvw�����嵥.ListItems(LNGDO).Tag, "_")(1)
            
            If lvw�����嵥.SelectedItem.SubItems(4) = "��" Then
                Me.Caption = "���������ϴ� �����ϴ�:" & STR���� & "������,���Ժ�..."
                If Not UPLOADREC(lng����ID, lng��ҳID, STRERR) Then
                    If MsgBox("�ϴ�����[" & STR���� & "]ʱ��������,�����ϴ�����������" & vbCrLf & _
                        STRERR, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Me.Caption = "���������ϴ�"
                            Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    Me.Caption = "���������ϴ�"
    Call READPATIENTS
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If ҽ����ʼ��_�������� = False Then
        Unload Me
        Exit Sub
    End If
    
    Me.dtp��������.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    Me.dtp��ʼ����.Value = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd")
End Sub

Private Function TRANDATA(ByVal STR��Ϣ�� As String, ByVal STR��Ϣֵ As String) As String
    '���ݽӿ��ĵ�ת��HIS�е�ֵ
    Select Case STR��Ϣ��
    Case "ҽ�Ƹ��ʽ"
        Select Case STR��Ϣֵ
        Case "������ҽ�Ʊ���"
            TRANDATA = 1
        Case "��ҵ����"
            TRANDATA = 2
        Case "�Է�ҽ��"
            TRANDATA = 3
        Case "����ҽ��"
            TRANDATA = 4
        Case "��ͳ��"
            TRANDATA = 5
        Case Else   '����
            TRANDATA = 6
        End Select
    Case "�Ա�"
        Select Case STR��Ϣֵ
        Case "��"
            TRANDATA = 1
        Case Else   'Ů
            TRANDATA = 2
        End Select
    Case "����"
        Select Case STR��Ϣֵ
        Case "δ��"
            TRANDATA = 1
        Case "�ѻ�"
            TRANDATA = 2
        Case "���"
            TRANDATA = 3
        Case Else   'ɥ
            TRANDATA = 4
        End Select
    Case "�벡�˹�ϵ"
        Select Case STR��Ϣֵ
        Case "��ż"
            TRANDATA = 1
        Case "��", "Ů"
            TRANDATA = 2
        Case "��ĸ"
            TRANDATA = 3
        Case Else   '����\��Ů\�游\��ĸ\���˵ȵ�,����������
            TRANDATA = 9
        End Select
    Case "ʬ���־", "����", "�����־", "ʾ�̲���", "RH", "��Ѫ��Ӧ"
        Select Case STR��Ϣֵ
        Case "��"
            TRANDATA = 1
        Case Else
            TRANDATA = 2
        End Select
    Case "Ѫ��"
        Select Case STR��Ϣֵ
        Case "A"
            TRANDATA = 1
        Case "B"
            TRANDATA = 2
        Case "AB"
            TRANDATA = 3
        Case "O"
            TRANDATA = 4
        Case Else
            TRANDATA = 5
        End Select
    Case "��������"
        Select Case STR��Ϣֵ
        Case "ȫ��"
            TRANDATA = 1
        Case "����"
            TRANDATA = 3
        Case Else
            TRANDATA = 2
        End Select
    Case "��������"
        Select Case STR��Ϣֵ
        Case "��"
            TRANDATA = 1
        Case "��"
            TRANDATA = 2
        Case Else
            TRANDATA = 3
        End Select
    Case "���ƽ��", "��Ժ���"
        Select Case STR��Ϣֵ
        Case "����", "����"
            TRANDATA = 1
        Case "��ת"
            TRANDATA = 2
        Case "δ��"
            TRANDATA = 3
        Case "����"
            TRANDATA = 4
        Case Else
            TRANDATA = 5
        End Select
    Case "HBSAG", "HCV_AB", "HIV_AB"
        Select Case STR��Ϣֵ
        Case "����"
            TRANDATA = 1
        Case "����"
            TRANDATA = 2
        Case Else
            TRANDATA = 0
        End Select
    Case "��Ժ����"
        Select Case STR��Ϣֵ
        Case "Σ"
            TRANDATA = 1
        Case "��"
            TRANDATA = 2
        Case Else
            TRANDATA = 3
        End Select
    Case "�п�", "����"     '�ӿ�����ͳһ�жϵ�
        Select Case STR��Ϣֵ
        Case "��/��"
            TRANDATA = "01"
        Case "��/��"
            TRANDATA = "02"
        Case "��/��"
            TRANDATA = "03"
        Case "��/��"
            TRANDATA = "04"
        Case "��/��"
            TRANDATA = "05"
        Case "��/��"
            TRANDATA = "06"
        Case "��/��"
            TRANDATA = "07"
        Case "��/��"
            TRANDATA = "08"
        Case Else
            TRANDATA = "09"
        End Select
    Case "�������"
        Select Case STR��Ϣֵ
        Case 5, 6, 7
            TRANDATA = Val(STR��Ϣֵ) - 1
        Case 1, 2, 3
            TRANDATA = Val(STR��Ϣֵ)
        End Select
    End Select
End Function

Private Function GET������Ŀ����(ByVal STRNAME As String, ByVal ARRNAME As Variant) As String
    Dim intDO As Integer, intCOUNT As Integer
    intCOUNT = UBound(ARRNAME)
    For intDO = 0 To intCOUNT
        If STRNAME = Split(ARRNAME(intDO), ",")(0) Then
            GET������Ŀ���� = Split(Split(ARRNAME(intDO), ",")(1), "-")(0)
            Exit Function
        End If
    Next
End Function


Private Sub lvw�����嵥_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvw�����嵥.ListItems.Count = 0 Then Exit Sub
    If lvw�����嵥.SelectedItem Is Nothing Then Exit Sub
    
    If lvw�����嵥.SelectedItem.SubItems(4) = "��" Then
        cmd����ϴ���־.Enabled = True
    Else
        cmd����ϴ���־.Enabled = False
    End If
End Sub
