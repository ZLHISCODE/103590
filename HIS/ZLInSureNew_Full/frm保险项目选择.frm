VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm������Ŀѡ�� 
   AutoRedraw      =   -1  'True
   Caption         =   "ҽ����Ŀѡ��"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm������Ŀѡ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7845
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   3690
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7845
      TabIndex        =   5
      Top             =   4350
      Width           =   7845
      Begin VB.CommandButton cmdRequery 
         Caption         =   "������ϸ"
         Height          =   350
         Left            =   3900
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ�б�"
         Height          =   350
         Left            =   2790
         TabIndex        =   10
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ϸ����(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   3
      Top             =   270
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   7144
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2434
      EndProperty
   End
   Begin MSComctlLib.ListView lvwClass 
      Height          =   3990
      Left            =   15
      TabIndex        =   1
      Top             =   285
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7038
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   15
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ��.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ��.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ����(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ��ϸ(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   2
      Top             =   30
      Width           =   4710
   End
End
Attribute VB_Name = "frm������Ŀѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String '�������,ҽ����ĿDetailCode
Private mrsDetail As ADODB.Recordset
Private mblnOK As Boolean
Private mint���� As Integer
Private mint���õ��� As Integer '����ר�ã�0��ʾ����������1��ʾ����������ɾ������˵���Ŀ��

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    '����ѡ����Ŀ����
    If mint���� = TYPE_�˳ɺ˹�ҵ Then
        mstrCode = Mid(lvwDetail.SelectedItem.Key, 2) & "|" & lvwDetail.SelectedItem.SubItems(1) & "|" & Mid(lvwClass.SelectedItem.Key, 2)
    ElseIf mint���� = TYPE_�������� Then
        mstrCode = Mid(lvwDetail.SelectedItem.Key, 3) & "|" & lvwDetail.SelectedItem.SubItems(1) & "|" & Mid(lvwClass.SelectedItem.Key, 2)
    Else
        mstrCode = Mid(lvwDetail.SelectedItem.Key, 2)
    End If
    mblnOK = True
    Unload Me
End Sub

Public Function GetCode(strCode As String, ByVal int���� As Integer) As Boolean
'���ܣ����һ���շ���Ŀ��ҽ������
'������strCode ����Ϊ��������������
'���أ��ɹ�����True
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    
    mblnOK = False
    mint���� = int����
    
    On Error GoTo ErrH
    
    Set rsTmp = New ADODB.Recordset
    Set mrsDetail = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    mrsDetail.CursorLocation = adUseClient
    
    Select Case int����
        Case TYPE_���Ͻ�ˮ
            With gcnSybase
                If .State = adStateOpen Then .Close
                .Provider = "MSDataShape"
                '�̶�ʹ�ø��û�������������ַ���
'                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "yyzf", "yhcsi2000"
                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "his", "his"
                If .State = adStateClosed Then Exit Function
            End With
            
            rsTmp.Open "Select Upper(SFDLBM) as CODE,SFDLMC as NAME From BG01SFXMDL order by CODE", gcnSybase, adOpenKeyset
            mrsDetail.Open "Select Upper(SFDLBM) as CLASSCODE,Upper(SFXMBM) as CODE,XMMC as NAME from v_bg02fwxm order by CLASSCODE,CODE", gcnSybase, adOpenKeyset
        Case TYPE_����ʡ, TYPE_������
            With gcnSybase
                If .State = adStateOpen Then .Close
                .Provider = "MSDataShape"
                '�̶�ʹ�ø��û�������������ַ���
'                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "yyzf", "yhcsi2000"
                .Open "Driver={Microsoft ODBC for Oracle};Server=" & "si2000", "his", "his"
                If .State = adStateClosed Then Exit Function
            End With
            
            rsTmp.Open "Select Upper(SFDLBM) as CODE,SFDLMC as NAME From BG01SFXMDL order by CODE", gcnSybase, adOpenKeyset
            mrsDetail.Open "select Upper(SFDLBM) as CLASSCODE,Upper(SFXMBM) as CODE,xmmc NAME,gg ���,dw ��λ,jx ����,cd ����," & _
                           " DECODE(tjdm,1,'�����ؼ�',2,'���๫��',3,'����ҹ�',5,'������ҩ',6,'���ٹ���',31,'����ҹ�','ȫ�Է�') AS ��� " & _
                           " from v_bg02fwxm Where YAB060 IN ('$$$$'," & IIf(int���� = TYPE_������, "'0101'", "'0000'") & ") order by CLASSCODE,CODE", gcnSybase, adOpenKeyset
        Case TYPE_�ɶ���
            If ҽ����ʼ��_�ɶ� = False Then Exit Function
            
            rsTmp.Open "Select Upper(sfdlbm) as CODE,sfdlmc as NAME From sfxmdl order by CODE", gcnSybase, adOpenKeyset
            mrsDetail.Open "Select Upper(sfdlbm) as CLASSCODE,Upper(sfxmbm) as CODE,xmmc as NAME from ypsfxmb order by CLASSCODE,CODE", gcnSybase, adOpenKeyset
        Case TYPE_������, type_�ɶ�����
            gstrSQL = "Select ������� as CLASSCODE,���� AS CODE,trim(����) AS NAME ,���� " & IIf(int���� = TYPE_������, ",��ע", "") & _
                           " from ������Ŀ where ����=[1] order by �������,����"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", int����)
        Case TYPE_�Թ���, TYPE_�ɶ���ũҽ, TYPE_�ɶ�����, Is > 900
            'ҽ������
            gstrSQL = "Select ���� AS CODE,���� AS NAME From ����֧������ where ����=[1] order by ����"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", int����)
            
            '����ҩ��
            gstrSQL = "Select ������� as CLASSCODE,���� AS CODE,���� AS NAME ,����,��ע " & _
                           " from ������Ŀ where ����=[1] order by �������,����"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", int����)
        Case TYPE_������
            gstrSQL = "Select ���� AS CODE,���� AS NAME From ����֧������ where ����=" & int���� & " order by ����"
            rsTmp.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
            
            gstrSQL = "SELECT A.����  AS CODE,A.���� AS NAME,A.����,A.��λ,A.������� as CLASSCODE,C.���� AS ���� " & _
                      "     ,A.�Ƿ���ҩ,A.�Ƿ�ҽ��,A.���۸�����,A.�����Ը�����,A.�۸�,A.��Ŀ�ں�,A.��������,A.˵�� " & _
                      "  FROM ������Ŀ A,���� C " & _
                      "  WHERE A.����=" & TYPE_������ & " AND A.���ͱ���=c.����(+) "
            mrsDetail.Open gstrSQL, gcn����, adOpenStatic, adLockReadOnly
        Case TYPE_ͭ��
            gstrSQL = "Select ���� AS CODE,���� AS NAME From ����֧������ where ����=" & int���� & " order by ����"
            rsTmp.Open gstrSQL, gcnͭ��, adOpenStatic, adLockReadOnly
            
            gstrSQL = "SELECT A.����  AS CODE,A.���� AS NAME,A.����,A.��λ,A.������� as CLASSCODE,C.���� AS ���� " & _
                      "     ,A.�Ƿ���ҩ,A.�Ƿ�ҽ��,A.���۸�����,A.�����Ը�����,A.�۸�,A.��Ŀ�ں�,A.��������,A.˵�� " & _
                      "  FROM ������Ŀ A,���� C " & _
                      "  WHERE A.����=" & TYPE_ͭ�� & " AND A.���ͱ���=c.����(+) "
            mrsDetail.Open gstrSQL, gcnͭ��, adOpenStatic, adLockReadOnly
        'Modified by ���� 20031218 ����������
        Case TYPE_��������, TYPE_����ʡ, TYPE_������, TYPE_��ƽ��, TYPE_�Ĵ�üɽ, TYPE_������, TYPE_��ɽ
            gstrSQL = "Select ���� AS CODE,���� AS NAME From ����֧������ where ����=" & int���� & " order by ����"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", int����)
            If rsTmp.RecordCount = 0 Then
                MsgBox "������ɱ��մ�������á�", vbInformation, gstrSysName
                Exit Function
            End If
            
            gstrSQL = "Select ������� as ClassCode ,���� AS CODE,���� AS NAME,����,��ע From ������Ŀ where ����=[1] order by ����"
            Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", int����)
            
        Case TYPE_�ɶ��ڽ�
            gstrSQL = "Select ���� AS CODE,���� AS NAME From ����֧������ where ����=" & TYPE_�ɶ��ڽ� & " order by ����"
            
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "��ȡ������Ŀ"
            gstrSQL = "SELECT A.����  AS CODE,A.���� AS NAME,A.����,A.������� as CLASSCODE,��ע " & _
                      "  FROM ������Ŀ A " & _
                      "  WHERE A.����=" & TYPE_�ɶ��ڽ� & _
                      "  order by ���� "
            zlDatabase.OpenRecordset mrsDetail, gstrSQL, "��ȡ������Ŀ"
        Case TYPE_����
            gstrSQL = "Select ���� AS CODE,���� AS NAME From ����֧������ where ����=" & TYPE_���� & " order by ����"
            
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "��ȡ������Ŀ"
            gstrSQL = "SELECT A.����  AS CODE,A.���� AS NAME,A.����,A.������� as CLASSCODE,��ע " & _
                      "  FROM ������Ŀ A " & _
                      "  WHERE A.����=" & TYPE_���� & _
                      "  order by ���� "
            zlDatabase.OpenRecordset mrsDetail, gstrSQL, "��ȡ������Ŀ"
        
        Case TYPE_�˳ɺ˹�ҵ
            
            gstrSQL = "SELECT 0  AS CODE,'ҩƷ' AS NAME from dual  union all "
            gstrSQL = gstrSQL & "SELECT 1  AS CODE,'����' AS NAME from dual  union all "
            gstrSQL = gstrSQL & "SELECT 2  AS CODE,'����' AS NAME from dual   "
            
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "��ȡ������Ŀ"
            
            gstrSQL = "select 0  ��ע, 'ҩƷ' as ���,xmdm CODE,xmmc  Name,pl Ʒ��,zfbl �Ը�����,0 CLASSCODE " & _
                     " from  YB_YD  " & _
                     " union all  " & _
                     " select 1  ��ע, '����' as ���,xmdm CODE,xmmc Name,pl Ʒ��,zfbl �Ը�����,1 CLASSCODE " & _
                     " from   YB_ZLML " & _
                     " union all  " & _
                     " select 2  ��ע, '����' as ���,xmdm CODE,xmmc Name,pl Ʒ��,zfbl �Ը�����,2 CLASSCODE" & _
                     " from  YB_FWSS " & _
                     " "
            mrsDetail.Open gstrSQL, gcnSQLSEVER_�˳�, adOpenStatic, adLockReadOnly
            
        Case TYPE_��������
            
            
            gstrSQL = "SELECT 1  AS CODE,'ҩƷ' AS NAME from dual  union all "
            gstrSQL = gstrSQL & "SELECT 2  AS CODE,'������Ŀ' AS NAME from dual  "
            zlDatabase.OpenRecordset rsTmp, gstrSQL, "��ȡ������Ŀ"

            gstrSQL = "" & _
                     " Select LB CLASSCODE, LB ��ע,decode(LB,1,'ҩƷ','����') as ���,LB||BM CODE,MC Name,PYBM ������, " & _
                     "        YPBM1 ����1,PYBM1 ����1������,YPBM2 ����2,PYBM2 ����2������,YPBM3 ����3,PYBM3 ����3������, " & _
                     "        YPJX  ����,JG �۸�,decode(YPLX,1,'�г�ҩ',2,'�в�ҩ',3,'��ҩ','') ҩƷ����, " & _
                     "        decode(BXLX,1,'����',2,'�Է�',3,'����','') ��������,GUIG ҩƷ��� " & _
                     " From YY_YPFZB " & _
                     " union all  " & _
                     " select ��� CLASSCODE,��� ��ע,decode(���,1,'ҩƷ','����') as ��� ,���||���� CODE,���� Name,���� ������, " & _
                     "        '' ����1,'' ����1������,'' ����2,'' ����2������,'' ����3,'' ����3������, " & _
                     "        ''  ����,0 �۸�,decode(ҩƷ����,1,'�г�ҩ',2,'�в�ҩ',3,'��ҩ','') ҩƷ����, " & _
                     "        decode(��������,1,'����',2,'�Է�',3,'����','') ��������,'' ҩƷ��� " & _
                     " From �շ���Ŀ������Ϣ"
                        
            
            mrsDetail.Open gstrSQL, gcnOracle_��ľ����, adOpenStatic, adLockReadOnly
            
        Case Else
            Exit Function
    End Select
    
    'Ϊ��ϸ���Ӷ�����ʾ����
    Dim fld As ADODB.Field
    For Each fld In mrsDetail.Fields
        If fld.Name <> "CLASSCODE" And fld.Name <> "NAME" And fld.Name <> "CODE" Then
            If fld.Name <> "��ע" Then
                lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
            Else
                '����ע���зֽ�
                'Modified by ���� 20031218 ����������
                If int���� = TYPE_�������� Or int���� = TYPE_����ʡ Or int���� = TYPE_������ Or int���� = TYPE_��ƽ�� Then
                    lvwDetail.ColumnHeaders.Add , , "���", 1000
                    lvwDetail.ColumnHeaders.Add , , "��λ", 1000
                    lvwDetail.ColumnHeaders.Add , , "��Ʊ����", 600
                    lvwDetail.ColumnHeaders.Add , , "��ѧ��", 1200
                    lvwDetail.ColumnHeaders.Add , , "��Ʒ��", 1200
                    lvwDetail.ColumnHeaders.Add , , "����", 1500
                    lvwDetail.ColumnHeaders.Add , , "����", 600
                    lvwDetail.ColumnHeaders.Add , , "�Ƿ�ҽ��", 1000, lvwColumnCenter
                ElseIf int���� = TYPE_����ʡ Or int���� = TYPE_������ Then
                    lvwDetail.ColumnHeaders.Add , , "���", 1000
                    lvwDetail.ColumnHeaders.Add , , "����", 600
                    lvwDetail.ColumnHeaders.Add , , "��λ", 1000
                    lvwDetail.ColumnHeaders.Add , , "����", 1500
                    lvwDetail.ColumnHeaders.Add , , "���", 1200
                ElseIf int���� = TYPE_�Ĵ�üɽ Then
                    lvwDetail.ColumnHeaders.Add , , "��λ", 1000
                    lvwDetail.ColumnHeaders.Add , , "���", 1000
                    lvwDetail.ColumnHeaders.Add , , "�Ƿ�ҽ��", 800
                    lvwDetail.ColumnHeaders.Add , , "�������", 1000, lvwColumnCenter
                ElseIf int���� = TYPE_�Թ��� Then
                    lvwDetail.ColumnHeaders.Add , , "��λ", 1000
                    lvwDetail.ColumnHeaders.Add , , "�Ƿ�ҽ��", 1000, lvwColumnCenter
                    lvwDetail.ColumnHeaders.Add , , "�Ƿ���ҩ", 1000, lvwColumnCenter
                    lvwDetail.ColumnHeaders.Add , , "����", 1000
                'Modified By ���� ��������ɳ
                ElseIf int���� = TYPE_������ Then
                    lvwDetail.ColumnHeaders.Add , , "���", 1000
                    lvwDetail.ColumnHeaders.Add , , "����", 1000
                    lvwDetail.ColumnHeaders.Add , , "����", 1000
                    If mint���õ��� = 1 Then lvwDetail.ColumnHeaders.Add , , "����", 1000
                ElseIf int���� = TYPE_��ɽ Then
                    lvwDetail.ColumnHeaders.Add , , "��������", 1000
                    lvwDetail.ColumnHeaders.Add , , "������Ŀ����", 1000
                    lvwDetail.ColumnHeaders.Add , , "������Ŀ", 1000
                    lvwDetail.ColumnHeaders.Add , , "ҽ������", 1000
                ElseIf int���� = TYPE_�˳ɺ˹�ҵ Then
                    lvwDetail.ColumnHeaders.Add , , "���", 1000
                    lvwDetail.ColumnHeaders.Add , , "Ʒ��", 1000
                    lvwDetail.ColumnHeaders.Add , , "�Ը�����", 1000
                ElseIf int���� = TYPE_�������� Then
                    
                    lvwDetail.ColumnHeaders.Add , , "���", 1000
                    lvwDetail.ColumnHeaders.Add , , "������", 1000
                    lvwDetail.ColumnHeaders.Add , , "����1", 1000
                    lvwDetail.ColumnHeaders.Add , , "����1������", 800
                    lvwDetail.ColumnHeaders.Add , , "����2", 1000
                    lvwDetail.ColumnHeaders.Add , , "����2������", 800
                    lvwDetail.ColumnHeaders.Add , , "����3", 1000
                    lvwDetail.ColumnHeaders.Add , , "����3������", 800
                    lvwDetail.ColumnHeaders.Add , , "����", 1000
                    lvwDetail.ColumnHeaders.Add , , "�۸�", 800
                    lvwDetail.ColumnHeaders.Add , , "ҩƷ����", 800
                    lvwDetail.ColumnHeaders.Add , , "��������", 800
                    lvwDetail.ColumnHeaders.Add , , "ҩƷ���", 1000
                ElseIf int���� = TYPE_������ Then
                    lvwDetail.ColumnHeaders.Add , , "����޼�", 1000
                    lvwDetail.ColumnHeaders.Add , , "�Ը�����", 1000
                    lvwDetail.ColumnHeaders.Add , , "������Ŀ", 1000
                    lvwDetail.ColumnHeaders.Add , , "������Ŀ", 1000
                    lvwDetail.ColumnHeaders.Add , , "���ⱨ��", 1000
                    lvwDetail.ColumnHeaders.Add , , "���ɽ���", 1000
                End If
            End If
        End If
    Next
    
    '��ʼ������
    If rsTmp.State = adStateOpen Then
        If Not rsTmp.EOF Then
            lvwClass.ListItems.Clear
            For i = 1 To rsTmp.RecordCount
                Set objItem = lvwClass.ListItems.Add(, "_" & rsTmp("CODE"), rsTmp("CODE"), , "Class")
                objItem.SubItems(1) = IIf(IsNull(rsTmp("NAME")), "", rsTmp("NAME"))
                rsTmp.MoveNext
            Next
        End If
    Else
        '�����������û�д����
        lblClass.Visible = False
        lvwClass.Visible = False
        picSplit.Visible = False
        Call lvwClass.ListItems.Add(, "_1", "1", , "Class")
    End If
    If int���� = TYPE_������ Or int���� = type_�ɶ����� Or int���� = TYPE_�������� Or _
    int���� = TYPE_�Ĵ�üɽ Or int���� = TYPE_��ɽ Or int���� = TYPE_������ Or _
    int���� = TYPE_����ʡ Or int���� = TYPE_������ Or int���� = TYPE_��ƽ�� Or int���� = TYPE_�ɶ��ڽ� Or int���� = TYPE_���� Then
        '��ϸ���Ը���
        cmdRequery.Visible = True
    End If
    
    
    If Not mrsDetail.EOF Then
       If mstrCode <> "" Then
            '���Ҵ�����벢��λ
            mrsDetail.Filter = "CODE Like '" & UCase(mstrCode) & "%'"
            If Not mrsDetail.EOF Then
                lvwClass.ListItems("_" & mrsDetail("CLASSCODE")).Selected = True
            ElseIf lvwClass.ListItems.Count > 0 Then
                lvwClass.ListItems(1).Selected = True
            End If
            Call lvwClass_ItemClick(lvwClass.SelectedItem)
            lvwClass.SelectedItem.EnsureVisible
        Else
            If lvwClass.ListItems.Count > 0 Then
                lvwClass.ListItems(1).Selected = True
            End If
            Call lvwClass_ItemClick(lvwClass.SelectedItem)
        End If
        
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    
    frm������Ŀѡ��.Show 1
    '����ֵ
    If mblnOK = True Then
        strCode = mstrCode
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdPrint_Click()
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "������Ŀ"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "ҽ�����ࣺ" & lvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
    End Select

End Sub

Private Sub cmdRequery_Click()
    Dim str�������� As String
    Dim str��ע As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnȫ�� As Boolean
    Dim blnReturn As Boolean
    
    If MsgBox("���������ܻỨ�Ƚϳ���ʱ�䣬�Ƿ������" & vbCrLf & vbCrLf & "����ע�⣬������ֻ����ҽ����Ŀ��ϸ������������Ӧ��ϵ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    With rsTemp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Fields.Append "CLASSCODE", adVarChar, 6 '�������
        'Modified By ���� 2003-12-09 ��������ɽ
        If mint���� = TYPE_��ɽ Then
            .Fields.Append "CODE", adVarChar, 40     '����
        Else
            .Fields.Append "CODE", adVarChar, 20     '����
        End If
        .Fields.Append "NAME", adVarChar, 300     '����
        .Fields.Append "PY", adVarChar, 150       'ƴ������
        .Fields.Append "MEMO", adVarChar, 500     '��ע
        .Open
    End With
    
    blnȫ�� = True
    Me.Caption = "ҽ����Ŀѡ�����ڶ�ȡ���ļ��������ȡ������Ŀ��ϸ�����Ժ�......��"
    If mint���� = TYPE_������ Then
        blnReturn = ҽ����Ŀ_����(rsTemp)
    ElseIf mint���� = type_�ɶ����� Then
        blnReturn = ҽ����Ŀ_�ɶ�����(rsTemp)
    ElseIf mint���� = TYPE_�������� Or mint���� = TYPE_����ʡ Or mint���� = TYPE_������ Or mint���� = TYPE_��ƽ�� Then
        blnReturn = ҽ����Ŀ_��������(rsTemp)
    ElseIf mint���� = TYPE_�Ĵ�üɽ Then
        blnReturn = ҽ����Ŀ_�Ĵ�üɽ(rsTemp)
    ElseIf mint���� = TYPE_��ɽ Then
        blnReturn = ҽ����Ŀ_��ɽ(rsTemp, blnȫ��)
    ElseIf mint���� = TYPE_������ Then
        blnReturn = ҽ����Ŀ_������(rsTemp)
    ElseIf mint���� = TYPE_�ɶ��ڽ� Then
        If MsgBox("�Ƿ����ԭ����ҽ����Ŀ��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) <> vbYes Then
            blnȫ�� = False
        End If
        blnReturn = ҽ����Ŀ_�ɶ��ڽ�(rsTemp)
    ElseIf mint���� = TYPE_���� Then
        If MsgBox("�Ƿ����ԭ����ҽ����Ŀ��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) <> vbYes Then
            blnȫ�� = False
        End If
        blnReturn = ҽ����Ŀ_����(rsTemp)
    End If
    
    If blnReturn = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.Caption = "ҽ����Ŀѡ�����ڸ���ҽ����Ŀ......��"
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    If blnȫ�� Then
        gstrSQL = "zl_������Ŀ_Clear(" & mint���� & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ҽ����Ŀѡ��")
    End If
    
    '���±�����Ŀ
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        str��ע = Nvl(rsTemp("MEMO"))
        If mint���� = TYPE_������ Then
            str�������� = Split(str��ע, "^^")(1)
            If Trim(str��������) <> "" Then
                'ֻҪ��Ϊ�գ�˵����ҩƷ��Ŀ�����·�������
                gstrSQL = "ZL_���·�������('" & rsTemp("CODE") & "','" & str�������� & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���·�������")
            End If
            str��ע = Split(str��ע, "^^")(0)
        End If
        
        '���뱣����Ŀ
        gstrSQL = "zl_������Ŀ_Insert(" & mint���� & ",'" & rsTemp("CODE") & "','" & ToVarchar(rsTemp("NAME"), 300) & _
            "','" & ToVarchar(rsTemp("PY"), 150) & "','" & ToVarchar(rsTemp("CLASSCODE"), 6) & "','" & ToVarchar(str��ע, 500) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҽ����Ŀ")
        Me.Caption = "ҽ����Ŀѡ�����ڸ���ҽ����Ŀ���Ѳ���" & rsTemp.AbsolutePosition & "����¼��"
        rsTemp.MoveNext
    Loop
    
    '���±��ղ���
    If mint���� = TYPE_������ Then
        Me.Caption = "ҽ����Ŀѡ�����ڶ�ȡ���ļ��������ȡ���ռ�����ϸ�����Ժ�......��"
        If Not ����Ŀ¼_���� Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
    End If
    gcnOracle.CommitTrans
    '����װ����ϸ
    mrsDetail.Requery
    Call lvwClass_ItemClick(lvwClass.SelectedItem)
    MousePointer = vbDefault
    Me.Caption = "ҽ����Ŀѡ��"
    MsgBox "������ɡ�", vbInformation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
    MousePointer = vbDefault
End Sub
Private Function ҽ����Ŀ_�ɶ��ڽ�(ByVal rsTemp As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Const COL_��Ŀ����   As Long = 1
    Const COL_���� As Long = 2
    Const COL_���� As Long = 3
    Const COL_����  As Long = 4
    Const COL_����  As Long = 5
    
    Err = 0
    On Error GoTo errHand:
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    
    'ѡ��ָ���ļ�
    On Error Resume Next
    Err = 0
    With Dlg
        .Filter = "EXCEL�ļ�(*.xls)|*.xls"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Function
        strFile = .FileName
    End With
    
    '����EXCEL����
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "EXCELδ��ȷ��װ������ȷ��װEXCEL���İ�������У�", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHand:
    Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ����......��"
    
    'ȡEXCEL�ļ�������
    With ObjExcel
        .Workbooks.Open strFile
        
        'ȡ���е�ֵ
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, COL_����) <> "" Then
                rsTemp.AddNew
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, COL_����)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_����)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_����)), 10)
                rsTemp("CLASSCODE") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_��Ŀ����)), 6), "'", "")
                rsTemp.Update
                Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ���ݣ��ѻ�ȡ" & rsTemp.RecordCount & "����¼��"
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '�ر�EXCEL����
    ObjExcel.quit
    Set ObjExcel = Nothing
    ҽ����Ŀ_�ɶ��ڽ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    If mint���� = TYPE_������ Then
        mint���õ��� = 0
        gstrSQL = "Select ����ֵ From ���ղ��� Where ������='���õ���' And ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���õ���", TYPE_������)
        If Not rsTemp.EOF Then
            mint���õ��� = Nvl(rsTemp!����ֵ, 0)
        End If
    End If
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = lvwClass.Width
    
    On Error Resume Next
    
    lvwClass.Left = 0: lvwClass.Top = lblClass.Top + lblClass.Height
    lvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = lvwClass.Top
    picSplit.Left = lvwClass.Left + lvwClass.Width
    picSplit.Height = lvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If lvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    
    lvwDetail.Top = lvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = lvwClass.Height
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwDetail_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If lvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        lvwClass.Width = lvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        lvwDetail.Left = lvwDetail.Left + x
        lvwDetail.Width = lvwDetail.Width - x
    End If
End Sub

Private Sub lvwdetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwDetail.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwDetail.SortOrder = lvwDescending
    Else
        lvwDetail.SortOrder = lvwAscending
    End If
    lvwDetail.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwDetail.SelectedItem Is Nothing Then lvwDetail.SelectedItem.EnsureVisible
End Sub

Private Sub lvwclass_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwClass.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwClass.SortOrder = lvwDescending
    Else
        lvwClass.SortOrder = lvwAscending
    End If
    lvwClass.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwClass.SelectedItem Is Nothing Then lvwClass.SelectedItem.EnsureVisible
End Sub

Private Sub lvwClass_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, objItem As ListItem
    Dim lngCount As Long, str�� As String, bln���⴦�� As Boolean
    Dim BLNSEL As Boolean
    Dim varPart As Variant
    
    
    Me.MousePointer = vbHourglass
    lvwDetail.ListItems.Clear
    If Item Is Nothing Then Exit Sub
    
    mrsDetail.Filter = "CLASSCODE='" & Mid(Item.Key, 2) & "'"
    If Not mrsDetail.EOF Then
        For i = 1 To mrsDetail.RecordCount
            Set objItem = lvwDetail.ListItems.Add(, "_" & mrsDetail("CODE"), mrsDetail("CODE"), , "Detail")
            objItem.SubItems(1) = IIf(IsNull(mrsDetail("NAME")), "", mrsDetail("NAME"))
            objItem.Tag = mrsDetail("CLASSCODE")
            
            '��ʾ�������
            With lvwDetail.ColumnHeaders
                For lngCount = 3 To lvwDetail.ColumnHeaders.Count
                    str�� = .Item(lngCount).Text
                    bln���⴦�� = False
                    
                    'Modified by ���� 20031218 ����������
                    If mint���� = TYPE_�������� Or mint���� = TYPE_����ʡ Or mint���� = TYPE_������ Or mint���� = TYPE_��ƽ�� Then
                        '��ע���ֶ����������ǣ���񡢷�Ʊ���ơ��Ƿ�ҽ��
                        If InStr(1, ",���,��λ,��Ʊ����,�Ƿ�ҽ��,��ѧ��,��Ʒ��,����,����,", str��) <> 0 Then
                            bln���⴦�� = True
                            varPart = Split(IIf(IsNull(mrsDetail("��ע")), "", mrsDetail("��ע")), "|")
                            Select Case str��
                                Case "���"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "��λ"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "��Ʊ����"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = varPart(2)
                                Case "�Ƿ�ҽ��"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lvwDetail.ColumnHeaders.Count - 1) = varPart(3)
                                Case "��ѧ��"
                                    If UBound(varPart) >= 4 Then objItem.SubItems(lngCount - 1) = varPart(4)
                                Case "��Ʒ��"
                                    If UBound(varPart) >= 5 Then objItem.SubItems(lngCount - 1) = varPart(5)
                                Case "����"
                                    If UBound(varPart) >= 6 Then objItem.SubItems(lngCount - 1) = varPart(6)
                                Case "����"
                                    If UBound(varPart) >= 7 Then objItem.SubItems(lngCount - 1) = varPart(7)
                            End Select
                        End If
                    'Modified By ���� ��������ɳ
                    ElseIf mint���� = TYPE_������ Then
                        If str�� = "����" Or str�� = "���" Or str�� = "����" Or str�� = "����" Then
                            bln���⴦�� = True
                            varPart = Split(IIf(IsNull(mrsDetail("��ע")), "", mrsDetail("��ע")), "||")
                            Select Case str��
                                Case "���"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "����"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = varPart(2)
                                Case "����"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = varPart(3)
                                Case "����"
                                    If UBound(varPart) >= 4 Then objItem.SubItems(lngCount - 1) = varPart(4)
                            End Select
                        End If
                    ElseIf mint���� = TYPE_�Ĵ�üɽ Then
                        '��ע���ֶ����������ǣ���񡢷�Ʊ���ơ��Ƿ�ҽ��
                        If str�� = "���" Or str�� = "�Ƿ�ҽ��" Or str�� = "��λ" Or str�� = "�������" Then
                            bln���⴦�� = True
                            varPart = Split(IIf(IsNull(mrsDetail("��ע")), "", mrsDetail("��ע")), "|")
                            Select Case str��
                                Case "��λ"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "���"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "�Ƿ�ҽ��"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = varPart(2)
                                Case "�������"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = varPart(3)
                            End Select
                        End If
                    ElseIf mint���� = TYPE_�Թ��� Then
                        '��ע���ֶ����������ǣ����ͱ��롢�Ƿ�ҽ�����Ƿ���ҩ����λ
                        If str�� = "��λ" Or str�� = "�Ƿ���ҩ" Or str�� = "�Ƿ�ҽ��" Or str�� = "����" Then
                            bln���⴦�� = True
                            varPart = Split(IIf(IsNull(mrsDetail("��ע")), "", mrsDetail("��ע")), "|")
                            If UBound(varPart) >= 4 Then
                                If str�� = "��λ" Then
                                    objItem.SubItems(lngCount - 1) = varPart(3)
                                ElseIf str�� = "�Ƿ���ҩ" Then
                                    objItem.SubItems(lngCount - 1) = IIf(varPart(2) = "1", "��", "��")
                                ElseIf str�� = "�Ƿ�ҽ��" Then
                                    objItem.SubItems(lngCount - 1) = IIf(varPart(1) = "1", "��", "��")
                                Else          '"����"
                                    objItem.SubItems(lngCount - 1) = varPart(4)
                                End If
                            End If
                        End If
                    ElseIf mint���� = TYPE_��ɽ Then
                        If str�� = "��������" Or str�� = "������Ŀ����" Or str�� = "������Ŀ" Or str�� = "ҽ������" Then
                            bln���⴦�� = True
                            varPart = Split(IIf(IsNull(mrsDetail("��ע")), "", mrsDetail("��ע")), "|")
                            Select Case str��
                                Case "��������"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "������Ŀ����"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "������Ŀ"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(2)) = 0, "��", "��")
                                Case "ҽ������"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = varPart(3)
                            End Select
                        End If
                    ElseIf mint���� = TYPE_������ Then
                        If str�� = "����޼�" Or str�� = "�Ը�����" Or str�� = "������Ŀ" Or str�� = "������Ŀ" _
                        Or str�� = "���ⱨ��" Or str�� = "���ɽ���" Then
                            bln���⴦�� = True
                            varPart = Split(IIf(IsNull(mrsDetail("��ע")), "", mrsDetail("��ע")), "|")
                            Select Case str��
                                Case "����޼�"
                                    If UBound(varPart) >= 0 Then objItem.SubItems(lngCount - 1) = varPart(0)
                                Case "�Ը�����"
                                    If UBound(varPart) >= 1 Then objItem.SubItems(lngCount - 1) = varPart(1)
                                Case "������Ŀ"
                                    If UBound(varPart) >= 2 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(2)) = 0, "��", "��")
                                Case "������Ŀ"
                                    If UBound(varPart) >= 3 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(3)) = 0, "��", "��")
                                Case "���ⱨ��"
                                    If UBound(varPart) >= 4 Then objItem.SubItems(lngCount - 1) = IIf(Val(varPart(4)) = 0, "��ͨ��Ŀ", IIf(Val(varPart(4)) = 1, "������Ա����֧���������Է���Ŀ", "����ֱ��֧����Ŀ"))
                                Case "���ɽ���" '01-��ͨ��Ŀ��02-���ɽ�����շ�Χ����Ŀ��03-ҽ���չ���Ա������Ŀ�� 04-����ֱ��֧����Ŀ��
                                                '05-���ɽ�������Է���Ŀ��06-��������ҽԺ����ȫ�Է���Ŀ
                                    If UBound(varPart) >= 5 Then
                                        objItem.SubItems(lngCount - 1) = IIf(Val(varPart(5)) = 1, "��ͨ��Ŀ", _
                                                                         IIf(Val(varPart(5)) = 2, "���ɽ�����շ�Χ����Ŀ", _
                                                                         IIf(Val(varPart(5)) = 3, "ҽ���չ���Ա������Ŀ", _
                                                                         IIf(Val(varPart(5)) = 4, "����ֱ��֧����Ŀ", _
                                                                         IIf(Val(varPart(5)) = 5, "���ɽ�������Է���Ŀ", "��������ҽԺ����ȫ�Է���Ŀ")))))
                                    End If
                            End Select
                        End If
                    End If
                    
                    If bln���⴦�� = False Then
                        'û�н������⴦��
                        objItem.SubItems(lngCount - 1) = IIf(IsNull(mrsDetail(.Item(lngCount).Text)), "", mrsDetail(.Item(lngCount).Text))
                    End If
                Next
            End With
                        
            If InStr(mrsDetail("CODE"), mstrCode) > 0 And Not BLNSEL Then
                objItem.Selected = True
                BLNSEL = True
            End If
            mrsDetail.MoveNext
        Next
        If Not BLNSEL And lvwDetail.ListItems.Count > 0 Then lvwDetail.ListItems(1).Selected = True
        lvwDetail.SelectedItem.EnsureVisible
    End If
    Call zlControl.LvwSetColWidth(lvwDetail)
    Me.MousePointer = vbDefault
End Sub

Private Sub txtFind_Change()
'���ܣ������û���������ݲ���ƥ�������
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvwDetail.ListItems.Count = 0 Then Exit Sub
    
    Set lst = lvwDetail.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '���ı�������������ƥ��
        lngSubItems = lvwDetail.ColumnHeaders.Count - 1
        For Each lst In lvwDetail.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Function ҽ����Ŀ_��������(rsTemp As ADODB.Recordset) As Boolean
'���ܣ����¸���������ҽ����Ŀ
    Const cOL���� As Long = 1
    Const COL�վݷ�Ŀ As Long = 2
    Const cOL���� As Long = 3
    Const COL��� As Long = 4
    Const COL��λ As Long = 5
    Const COL�Ƿ�ҽ�� As Long = 7
    Const COL���� As Long = 8
    Const COLƴ�� As Long = 9
    Const COL��ѧ�� As Long = 10
    Const COL��Ʒ�� As Long = 11
    Const COL���� As Long = 12
    Const COL���� As Long = 13
    
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    Dim rs���� As New ADODB.Recordset
    
    
    gstrSQL = "Select ����,���� From ����֧������ Where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    
    'ѡ��ָ���ļ�
    On Error Resume Next
    Err = 0
    With Dlg
        .Filter = "EXCEL�ļ�(*.xls)|*.xls"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Function
        strFile = .FileName
    End With
    
    '����EXCEL����
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "EXCELδ��ȷ��װ������ȷ��װEXCEL���İ�������У�", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ����......��"
    
    'ȡEXCEL�ļ�������
    With ObjExcel
        .Workbooks.Open strFile
        
        'ȡ���е�ֵ
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL����) <> "" Then
                rsTemp.AddNew
                
                rs����.Filter = "����='" & Trim(.ActiveSheet.Cells(lngRow, COL����)) & "'"
                If rs����.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs����("����")
                End If
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COLƴ��)), 10)
                rsTemp("MEMO") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL���)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL��λ)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL�վݷ�Ŀ)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL�Ƿ�ҽ��)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL��ѧ��)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL��Ʒ��)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL����)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL����)), 1000)
                rsTemp.Update
                Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ���ݣ��ѻ�ȡ" & rsTemp.RecordCount & "����¼��"
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '�ر�EXCEL����
    ObjExcel.quit
    Set ObjExcel = Nothing
    ҽ����Ŀ_�������� = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ҽ����Ŀ_�Ĵ�üɽ(rsTemp As ADODB.Recordset) As Boolean
'���ܣ����¸���������ҽ����Ŀ
    Const cOL���� As Long = 1
    Const cOL���� As Long = 2
    Const COL��λ As Long = 3
    Const COL��� As Long = 4
    Const COL�Ƿ�ҽ�� As Long = 5
    Const COL���� As Long = 6
    Const cOL���� As Long = 7
    Const COL������� As Long = 8
    
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    Dim int������� As Integer
    Dim rs���� As New ADODB.Recordset
    
    gstrSQL = "Select ����,���� From ����֧������ Where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    
    'ѡ��ָ���ļ�
    On Error Resume Next
    Err = 0
    With Dlg
        .Filter = "EXCEL�ļ�(*.xls)|*.xls"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Function
        strFile = .FileName
    End With
    
    '����EXCEL����
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "EXCELδ��ȷ��װ������ȷ��װEXCEL���İ�������У�", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ����......��"
    
    'ȡEXCEL�ļ�������
    With ObjExcel
        .Workbooks.Open strFile
        
        'ȡ���е�ֵ
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL����) <> "" Then
                rsTemp.AddNew
                
                rs����.Filter = "����='" & Trim(.ActiveSheet.Cells(lngRow, COL����)) & "'"
                If rs����.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs����("����")
                End If
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 10)
                If Trim(.ActiveSheet.Cells(lngRow, COL�������)) = "���סԺ" Then
                    int������� = 3
                Else
                    If Trim(.ActiveSheet.Cells(lngRow, COL�������)) = "����" Then
                        int������� = 1
                    Else
                        int������� = 2
                    End If
                End If
                rsTemp("MEMO") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL��λ)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL���)) & _
                                "|" & Trim(.ActiveSheet.Cells(lngRow, COL�Ƿ�ҽ��)) & _
                                "|" & int�������, 50)
                rsTemp.Update
                Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ���ݣ��ѻ�ȡ" & rsTemp.RecordCount & "����¼��"
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '�ر�EXCEL����
    ObjExcel.quit
    Set ObjExcel = Nothing
    ҽ����Ŀ_�Ĵ�üɽ = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ҽ����Ŀ_����(rsTemp As ADODB.Recordset) As Boolean
    '���ܣ����¸���������ҽ����Ŀ
    Const cOL���� As Long = 1
    Const cOL���� As Long = 2
    Const cOL���� As Long = 3
    Const COL���� As Long = 4
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    
    Dim int������� As Integer
    Dim rs���� As New ADODB.Recordset
    
    gstrSQL = "Select ����,���� From ����֧������ Where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    
    'ѡ��ָ���ļ�
    On Error Resume Next
    Err = 0
    With Dlg
        .Filter = "EXCEL�ļ�(*.xls)|*.xls"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Function
        strFile = .FileName
    End With
    
    
    '����EXCEL����
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "EXCELδ��ȷ��װ������ȷ��װEXCEL���İ�������У�", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ����......��"
    
    'ȡEXCEL�ļ�������
    With ObjExcel
        .Workbooks.Open strFile
        
        'ȡ���е�ֵ
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL����) <> "" Then
                rsTemp.AddNew
                
                rs����.Filter = "����='" & Trim(.ActiveSheet.Cells(lngRow, COL����)) & "'"
                If rs����.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs����("����")
                End If
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 10)
                rsTemp.Update
                Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ���ݣ��ѻ�ȡ" & rsTemp.RecordCount & "����¼��"
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '�ر�EXCEL����
    ObjExcel.quit
    Set ObjExcel = Nothing
    ҽ����Ŀ_���� = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ҽ����Ŀ_������(rsTemp As ADODB.Recordset) As Boolean
    Dim str���� As String, str���� As String, str���� As String, str���� As String
    Dim str��� As String, str���� As String, str���� As String, str�������� As String
    Dim str���� As String, int���� As Integer, strTmp As String
    Dim rs���� As New ADODB.Recordset
    Dim classInsure As New clsInsure
    '���»�ȡҽ����Ŀ
    On Error GoTo errHand
    
    If Not classInsure.InitInsure(gcnOracle, TYPE_������) Then Exit Function
    
    '���û�����ô������˳�
    gstrSQL = "Select ����,���� From ����֧������ Where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_������)
    If rs����.RecordCount < 4 Then
        MsgBox "������ȷ�������˱��մ������ʹ�ñ����ܣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ʼ�������ĵ�ҽ����Ŀ
    '----��ȡ������Ŀ----
    If Not ���ýӿ�_׼��_������(Function_������.��Ŀƥ��_ȡ��Ŀ��Ϣ) Then Exit Function
    '0-������Ŀ;1-ҩƷ
    gstrField_������ = "match_type"
    gstrValue_������ = "0"
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������() Then Exit Function
    If Not ���ýӿ�_ָ����¼��_������("diseaseinfo") Then Exit Function
'    (1)��match_type="0"(������Ŀ)ʱ�����ݼ������������ݣ�
'    ���    �ֶ�    �ֶ�˵��    ��󳤶�    ��ע
'    1   item_code  ��Ŀ����    20
'    2   item_name  ��Ŀ����    50
'    3   price      ����        12
'    4   code_wb    �����      20
'    5   code_py    ƴ����      20
'    (2)��match_type="1"(ҩƷ)ʱ�����ݼ������������ݣ�
'    ���    �ֶ�    �ֶ�˵��    ��󳤶�    ��ע
'    1   medi_code      ҩƷ����    20
'    2   medi_name      ҩƷ����    50
'    3   model_name     ��������    12
'    4   factory        ��������    50
'    5   standard       ���        20
'    6   medi_item_type ҩƷ����    1   "1"����ҩ   "2"���г�ҩ    "3"���в�ҩ
'    7   Staple_flag    ��������    1   "1"������   "2"������      "9"��ȫ�Է�
'    8   medi_item_name ҩƷ��������10
'    9   code_wb        �����      20
'   10   code_py        ƴ����      20
    int���� = 0
    str���� = "������Ŀ"
    rs����.Filter = "����='" & str���� & "'"
    If rs����.RecordCount > 0 Then
        str���� = rs����("����")
    End If
    If ���ýӿ�_��¼��_������ Then
        Do While True
            Call ���ýӿ�_��ȡ����_������("item_code", str����)
            Call ���ýӿ�_��ȡ����_������("item_name", str����)
            Call ���ýӿ�_��ȡ����_������("code_py", str����)
            If mint���õ��� = 1 Then
                Call ���ýӿ�_��ȡ����_������("price", str����)
            End If
            '��ע���ݸ�ʽ������||���||����||����^^ƥ�����к�
            Call AddRecord(rsTemp, str����, ToVarchar(str����, 300), ToVarchar(str����, 150), "0|| || || ||" & str���� & "^^", ToVarchar(str����, 6))
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    '----ȡҩƷ��Ϣ----
    If Not ���ýӿ�_׼��_������(Function_������.��Ŀƥ��_ȡ��Ŀ��Ϣ) Then Exit Function
    '0-������Ŀ;1-ҩƷ
    gstrField_������ = "match_type"
    gstrValue_������ = "1"
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������() Then Exit Function
    If Not ���ýӿ�_ָ����¼��_������("diseaseinfo") Then Exit Function
    If ���ýӿ�_��¼��_������ Then
        Do While True
            Call ���ýӿ�_��ȡ����_������("medi_code", str����)
            Call ���ýӿ�_��ȡ����_������("medi_name", str����)
            Call ���ýӿ�_��ȡ����_������("code_py", str����)
            Call ���ýӿ�_��ȡ����_������("standard", str���)
            Call ���ýӿ�_��ȡ����_������("model_name", str����)
            Call ���ýӿ�_��ȡ����_������("factory", str����)
            If mint���õ��� = 1 Then
                Call ���ýӿ�_��ȡ����_������("price", str����)
            End If
            
            'ȡҩƷ���ͼ�������Ϣ
            Call ���ýӿ�_��ȡ����_������("medi_item_type", strTmp)
            int���� = Val(strTmp)
            str���� = IIf(int���� = 1, "����ҩ", IIf(int���� = 2, "�г�ҩ", "�в�ҩ"))
            rs����.Filter = "����='" & str���� & "'"
            If rs����.RecordCount > 0 Then
                str���� = rs����("����")
            End If
            
            'ȡ��������
            Call ���ýӿ�_��ȡ����_������("staple_flag", strTmp)
            If Val(strTmp) = 1 Then
                strTmp = "����ҩƷ"
            ElseIf Val(strTmp) = 2 Then
                strTmp = "����ҩƷ"
            Else
                strTmp = "�ǻ���ҩƷ"
            End If
            
            Call AddRecord(rsTemp, str����, ToVarchar(str����, 300), ToVarchar(str����, 150), int���� & "||" & str��� & "||" & str���� & "||" & str���� & "||" & str���� & "^^" & strTmp, ToVarchar(str����, 6))
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    ҽ����Ŀ_������ = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ҽ����Ŀ_��ɽ(rsTemp As ADODB.Recordset, bln���·�ʽ As Boolean) As Boolean
'���ܣ����¸���������ҽ����Ŀ
    Dim str���� As String
    Const COL���� As Long = 1   '0:ҩƷ;1-����;2-����
    Const cOL���� As Long = 2
    Const COLҽ������ As Long = 3
    Const cOL���� As Long = 4
    Const cOL���� As Long = 5
    Const COL�������� As Long = 6
    Const COL������Ŀ���� As Long = 7
    Const COL������Ŀ As Long = 8
    
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    Dim rs���� As New ADODB.Recordset
    
    gstrSQL = "Select ����,���� From ����֧������ Where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    
    'ѡ��ָ���ļ�
    On Error Resume Next
    Err = 0
    With Dlg
        .Filter = "EXCEL�ļ�(*.xls)|*.xls"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Function
        strFile = .FileName
    End With
    
    '����EXCEL����
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "EXCELδ��ȷ��װ������ȷ��װEXCEL���İ�������У�", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHandle
    bln���·�ʽ = Not ����ģʽ(TYPE_��ɽ)       '���ʾ���������ʾȫ��������㲻��������ȡ��
    Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ����......��"
    
    'ȡEXCEL�ļ�������
    With ObjExcel
        .Workbooks.Open strFile
        
        'ȡ���е�ֵ
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, cOL����) <> "" Then
                str���� = Trim(.ActiveSheet.Cells(lngRow, COL����))
                Select Case str����
                Case "0"
                    str���� = "ҩƷ"
                Case "1"
                    str���� = "����"
                Case "2"
                    str���� = "����"
                End Select
                
                rsTemp.AddNew
                
                rs����.Filter = "����='" & str���� & "'"
                If rs����.RecordCount > 0 Then
                    rsTemp("ClassCode") = rs����("����")
                End If
                rsTemp("Code") = .ActiveSheet.Cells(lngRow, cOL����)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, cOL����)), 10)
                rsTemp("MEMO") = Trim(.ActiveSheet.Cells(lngRow, COL��������)) & "|" & Trim(.ActiveSheet.Cells(lngRow, COL������Ŀ����)) & "|" & Trim(.ActiveSheet.Cells(lngRow, COL������Ŀ)) & "|" & Trim(.ActiveSheet.Cells(lngRow, COLҽ������))
                rsTemp.Update
                Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ���ݣ��ѻ�ȡ" & rsTemp.RecordCount & "����¼��"
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '�ر�EXCEL����
    ObjExcel.quit
    Set ObjExcel = Nothing
    ҽ����Ŀ_��ɽ = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ����Ŀ¼_����() As Boolean
    Dim lngRecords As Long
    Dim lngNextID As Long
    Dim str���� As String, str���� As String, str���� As String
    Dim blnInsert As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim classInsure As New clsInsure
    '���»�ȡҽ����Ŀ
    On Error GoTo errHand
    
    lngRecords = 1
    If Not classInsure.InitInsure(gcnOracle, TYPE_������) Then Exit Function
    '----ȡ������Ϣ----
    If Not ���ýӿ�_׼��_������(Function_������.��Ŀƥ��_ȡ��Ŀ��Ϣ) Then Exit Function
    '0-������Ŀ;1-ҩƷ;2-����
    gstrField_������ = "match_type"
    gstrValue_������ = "2"
    If Not ���ýӿ�_д��ڲ���_������(1) Then Exit Function
    If Not ���ýӿ�_ִ��_������() Then Exit Function
    If Not ���ýӿ�_ָ����¼��_������("diseaseinfo") Then Exit Function
    If ���ýӿ�_��¼��_������ Then
        '����ɾ�������²��벡����Ϣ����Ϊ����ID������������ϵ��ֻ���µĲ��ֲ��ܲ��룬���в���ͨ���޸�ʵ��
'        gstrSQL = "zl_���ղ���_DELETEALL(" & TYPE_������ & ")"
'        Call zlDatabase.ExecuteProcedure(gstrSQL, "ɾ������������ҽ������")
        gstrSQL = "Select ID,���� From ���ղ��� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���в���Ŀ¼", TYPE_������)
        
        Do While True
            Call ���ýӿ�_��ȡ����_������("icd", str����)
            Call ���ýӿ�_��ȡ����_������("disease", str����)
            Call ���ýӿ�_��ȡ����_������("code_py", str����)
            str���� = Replace(str����, "'", "")
            
            With rsTemp
                .Filter = "����='" & str���� & "'"
                blnInsert = (.RecordCount = 0)
            End With
            
            '���±��ռ���
            If blnInsert Then
                lngNextID = zlDatabase.GetNextID("���ղ���")
                gstrSQL = "zl_���ղ���_INSERT(" & lngNextID & "," & TYPE_������ & ",'" & str���� & _
                            "','" & str���� & "','" & str���� & "',0,NULL,NULL)"
            Else
                lngNextID = rsTemp!ID
                gstrSQL = "zl_���ղ���_UPDATE(" & lngNextID & ",'" & str���� & _
                            "','" & str���� & "','" & str���� & "',0,NULL,NULL)"
            End If
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Me.Caption = "ҽ����Ŀѡ�����ڸ���ҽ������Ŀ¼���Ѳ���" & lngRecords & "����¼��"
            lngRecords = lngRecords + 1
            
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
        rsTemp.Filter = 0
    End If
    
    ����Ŀ¼_���� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AddRecord(rsObj As ADODB.Recordset, ByVal str���� As String, ByVal str���� As String, _
str���� As String, ByVal str��ע As String, ByVal str���� As String)
    With rsObj
        .AddNew
        !CODE = str����
        !Name = Replace(str����, "'", "")
        !py = Replace(str����, "'", "")
        !Memo = Replace(str��ע, "'", "")
        !ClassCode = str����
        .Update
    End With
End Sub

Private Function ����ģʽ(ByVal lng���� As Long) As Boolean
    Dim intReturn As Integer
    Dim rsTemp As New ADODB.Recordset
    '����Ƿ��Ѵ���ҽ����Ŀ���������ʾ�������أ��ٱ�ʾȫ������
    gstrSQL = "Select 1 From ������Ŀ Where ����=[1] And Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ��Ѵ���ҽ����Ŀ", lng����)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '˵���Ѵ��ڼ�¼����ʾ����Ա����ȡ�ķ�ʽ
    intReturn = MsgBox("�����Ѵ������ݣ������Ƿ��ȡ�������أ�" & vbCrLf & "�Ǳ�ʾ�������أ����ʾȫ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName)
    ����ģʽ = (intReturn = vbYes)
End Function
