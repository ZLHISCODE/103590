VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm������Ŀѡ������ 
   AutoRedraw      =   -1  'True
   Caption         =   "ҽ����Ŀѡ��"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "frm������Ŀѡ������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7815
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7815
      TabIndex        =   4
      Top             =   4320
      Width           =   7815
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
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   7
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ϸ����(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   5
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   2
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
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3915
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
            Picture         =   "frm������Ŀѡ������.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ������.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   10
      Top             =   270
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
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
      TabIndex        =   1
      Top             =   30
      Width           =   4710
   End
End
Attribute VB_Name = "frm������Ŀѡ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mstrCode As String
Private mstrName As String
Private mdblҽԺ�۸� As Double
Private mobjStream As TextStream
Private mobjFileSystem As New FileSystemObject
Private mblnOK As Boolean
Private mcnYB As New ADODB.Connection   'ҽ��ǰ�÷���������
Private Const strFile = "C:\CQYB_YH\ERR.LOG"
Private mErrFile As TextStream
'��������ҽ�������� 204-04-07 ��Ҫ�Ǽ��˺��������޸���ҩƷ�����Ƽ����֣�������ƺ������ŵ�����

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str��׼���� As String, str�޼� As String, str������Ŀ As String, str������ As String
    
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Լ۸�����ж�
    Call GetValueByCol("�޼�", str�޼�)
    If str�޼� <> "" And mdblҽԺ�۸� > 0 Then
        Call GetValueByCol("��׼����", str��׼����)
        Call GetValueByCol("������Ŀ��־", str������Ŀ)
        Call GetValueByCol("������", str������)
        
        If mint���� = TYPE_������ Then
            If �۸��ж�_����(mdblҽԺ�۸�, Val(str��׼����), str�޼�, str������Ŀ = "��", Val(str������)) = False Then
                Exit Sub
            End If
        Else
            If �۸��ж�_����������(mdblҽԺ�۸�, Val(str��׼����), str�޼�, str������Ŀ = "��", Val(str������)) = False Then
                Exit Sub
            End If
        End If
    End If
    
    '����ѡ����Ŀ����
    mstrCode = lvwDetail.SelectedItem.Text
    '��Ʒ������Ŀ����ֻ������һ����Ч
    Call GetValueByCol("��Ʒ��", mstrName)
    Call GetValueByCol("��Ŀ����", mstrName)
    
    mblnOK = True
    Unload Me
End Sub

Private Sub GetValueByCol(ByVal strColumnName As String, strValue As String)
    Dim lngCount As Long, lngIndex As Long

    For lngCount = 1 To lvwDetail.ColumnHeaders.Count
        If lvwDetail.ColumnHeaders(lngCount).Text = strColumnName Then
            lngIndex = lngCount
            Exit For
        End If
    Next
    
    If lngIndex > 0 Then
        strValue = lvwDetail.SelectedItem.SubItems(lngIndex - 1)
    End If
End Sub

Public Function GetCode(strCode As String, STRNAME As String, ByVal dblҽԺ�۸� As Double, ByVal int���� As Integer) As Boolean
'���ܣ����һ���շ���Ŀ��ҽ������
'������strCode ����Ϊ��������������
'���أ��ɹ�����True
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim nod As Node
    
    mblnOK = False
    mstrCode = strCode
    mdblҽԺ�۸� = dblҽԺ�۸�
    mint���� = int����
    
    On Error GoTo ErrH
    
    '���ȶ���������������
    gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, int����)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
        Select Case rsTemp("������")
            Case "ҽ��������"
                strServer = strTemp
            Case "ҽ���û���"
                strUser = strTemp
            Case "ҽ���û�����"
                strPass = strTemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    
    If int���� = TYPE_���������� Then
        '��������ҽ�������� 204-03-29
        On Error Resume Next
        If Not ҽ����ʼ��_���������� Then
            Unload Me
            Exit Function
        End If
    End If
    
    '��ʾҩƷ���
    gstrSQL = "select BH id,FXBH �ϼ�ID,LBDM ����,LBMC ����,'Y' ���,level ���� from YPML_LBDM start with FXBH=0 connect by prior BH=FXBH " & _
             " Union All " & _
             " select LBDM as id,'0' �ϼ�ID,LBDM ����,LBMC ����,'Z' ���,1 ���� from zlxm_lbdm2 " & _
             " order by ��� Desc,����,����"
    
    If int���� = TYPE_ɽ�� Then
        gstrSQL = "Select 11 id,0 as �ϼ�ID,'11' ����,'��ҩ' as ����,'Y' ���,1 as ���� from dual " & _
                   " union all " & _
                   " Select 12 id,0 as �ϼ�ID,'12' ����,'�г�ҩ' as ����,'Y' ���,1 as ���� from dual " & _
                    " union all " & _
                   " Select 13 id,0 as �ϼ�ID,'13' ����,'�в�ҩ' as ����,'Y' ���,1 as ���� from dual " & _
                    " union all " & _
                   " Select 90 id,0 as �ϼ�ID,'90' ����,'������Ŀ' as ����,'Z' ���,1 as ���� from dual "
        
    End If
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, mcnYB, adOpenStatic, adLockReadOnly

    If rsTemp.EOF = True Then
        MsgBox "ҽ��ǰ�÷�������û��ҩƷ������ݣ��޷�ѡ��", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    tvwClass.Nodes.Clear
    Do Until rsTemp.EOF
        If rsTemp("�ϼ�id") = 0 Then
            Set nod = tvwClass.Nodes.Add(, , rsTemp("���") & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "Class", "Class")
        Else
            Set nod = tvwClass.Nodes.Add(rsTemp("���") & rsTemp("�ϼ�id"), tvwChild, rsTemp("���") & rsTemp("id"), "��" & rsTemp("����") & "��" & rsTemp("����"), "Class", "Class")
        End If
        nod.Sorted = True
        rsTemp.MoveNext
    Loop
    
    tvwClass.Nodes(1).Selected = True
    Call FillList
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    frm������Ŀѡ������.Show vbModal, frm������Ŀ
    '����ֵ
    If mblnOK = True Then
        strCode = mstrCode
        STRNAME = mstrName
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FillList()
'���ܣ���ʾ��ǰ����µ�ҽ����ϸ
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, fld As ADODB.Field
    Dim str������ As String, blnColSet As Boolean
    Dim lngCol  As Long
    Dim varValue As Variant
    Dim strҽԺ�ȼ� As String
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandle
    With tvwClass.SelectedItem
        str������ = Mid(.Text, 2, InStr(.Text, "��") - 2)
    End With
    
    rsTemp.CursorLocation = adUseClient
    '��ʱ���б���ˢ��
    LockWindowUpdate lvwDetail.hwnd
    lvwDetail.ListItems.Clear
    
    If mint���� = TYPE_������ Then
        If Left(tvwClass.SelectedItem.Key, 1) = "Y" Then
            '��ǰѡ���ǵ�һ��ҩƷ���
            gstrSQL = "select YPLSH  ҽ������,YPBM ҩƷ����,REPLACE(TYM,chr(39),'') ͨ������,REPLACE(SPM,chr(39),'') ��Ʒ��,SPMZJM ��Ʒ������,YCMC ҩ������,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ� " & _
                      "      ,PFJ ������,BZDJ ��׼����,ZFBL �Ը�����,JX ����,BZSL ��װ����,BZDW ��װ��λ,HL ����,HLDW ������λ,RL ����,RLDW ������λ " & _
                      "      ,DECODE(CFYBZ,1,'��') ����ҩ��־,decode(GMP,1,'��') GMP��־,decode(YPXJFS,1,'�޼�',2,'��ҽԺ�ȼ��޼�',3,'���������޼�',20,'������') �޼�,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼���� " & _
                      "  FROM YPML WHERE LBDM='" & str������ & "'"
        Else
            '��ǰѡ���ǵ�һ���������
            gstrSQL = "Select XMLSH ҽ������,XMBM ���Ʊ���,REPLACE(XMMC,chr(39),'') ��Ŀ����,REPLACE(ZJM,chr(39),'') ����,decode(FYDJ,1,'����',2,'����','�Է�') ���õȼ�,DW ��λ " & _
                     "       ,TPJ ������,BZJ ��׼����,ZZBL ��ְ�Ը�����,TXBL �����Ը�����,decode(XJFS,1,'ͳһ�޼�',2,'��ҽԺ�ȼ�����',3,'������ҽԺ��׼��������') �޼� " & _
                     "       ,TQFYDJ ��Ⱥ��Ŀ�ȼ�,TQZFBL ��Ⱥ�Ը�����,TQBZDJ ��Ⱥ��׼����,decode(TPXMBZ,1,'��') ������Ŀ��־,BZ ��ע " & _
                     "   FROM ZLXM WHERE LBDM2='" & str������ & "'"
        End If
    Else
        If mint���� = TYPE_ɽ�� Then
            'ȡҽԺ�ȼ�,
            gstrSQL = "Select * from ���ղ��� where ����=[1] and ������='ҽԺ�ȼ�'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽԺ�ȼ�", TYPE_ɽ��)
            If rsTemp.EOF Then
                strҽԺ�ȼ� = ""
            Else
                strҽԺ�ȼ� = " where aka101=" & Val(Mid(rsTemp!����ֵ, 1, 2))
            End If
        
            If Left(tvwClass.SelectedItem.Key, 1) = "Y" Then
                '��ǰѡ���ǵ�һ��ҩƷ���
                gstrSQL = "Select aka060 ҽ������,aka065  ҩƷ�ȼ�,aka061  ͨ������,aka074  ���,aka067  ��λ,aka068  ��߼۸�,aka069  �Ը�����,1 as ��ע,aka070  ����,aka060  ҩƷ����,aka062  ��Ʒ��,zka003  ҩƷ����,aka064  ����ҩ��־,aka063  �շ����,aka066  ��Ʒ������,"
                gstrSQL = gstrSQL & "aka071  ÿ������,aka072  ʹ��Ƶ��,aka073  �÷�,ckc050  �޶�����,aae013  ��ע,aae035 �������"
                gstrSQL = gstrSQL & " From ka02 where zkA003 like '" & str������ & "%'"
            Else
                '��ǰѡ���ǵ�һ���������
                gstrSQL = "Select aka090  ҽ������,aka091 AS ��Ŀ����,aka065  ��Ŀ�ȼ�,aka068  ��߼۸�,aka069  �Ը�����,2 as ��ע,aka101  ҽԺ�ȼ�,aka063  �շ����,aka066  ����,aae035  �������,aae013 ��ע" & _
                         " From ka03 " & strҽԺ�ȼ� & _
                                    " Union All "
                gstrSQL = gstrSQL & " Select aka100  ҽ������,aka102  ������ʩ����,aka103  �����ȼ�,aka104  ����޼�,0   �Ը�����,3 as ��ע,aka101  ҽԺ�ȼ�,aka063  �շ����,aka066  ����,aae035  �������,'������ʩ' ��ע"
                gstrSQL = gstrSQL & " From ka04 " & strҽԺ�ȼ�

            End If
            
        Else
            If Left(tvwClass.SelectedItem.Key, 1) = "Y" Then
                '��ǰѡ���ǵ�һ��ҩƷ���
                gstrSQL = "select ��ˮ�� ҽ������,���� ҩƷ����,ͨ���� ͨ������,��Ʒ�� ��Ʒ��,��Ʒ�������� ��Ʒ������,ҩ������,decode(��Ŀ�ȼ�,1,'����',2,'����','�Է�') ���õȼ� " & _
                          "      ,������,��׼����,�Ը�����,����,��װ����,��װ��λ,����,������λ,����,������λ " & _
                          "      ,DECODE(����ҩ��־,1,'��') ����ҩ��־,decode(GMP��־,1,'��') GMP��־,decode(�޼۷�ʽ,1,'�޼�') �޼�,��Ⱥ��Ŀ�ȼ�,��Ⱥ�Ը�����,��Ⱥ��׼���� " & _
                          "  FROM �м��_ҩƷĿ¼ WHERE ���� like '" & str������ & "%'"
            Else
                '��ǰѡ���ǵ�һ���������
                gstrSQL = "Select ��ˮ�� ҽ������,��Ŀ���� ���Ʊ���,��Ŀ����,������ ����,decode(��Ŀ�ȼ�,1,'����',2,'����','�Է�') ���õȼ�,��λ " & _
                         "       ,������,��׼����,��ְ���� ��ְ�Ը�����,���ݱ��� �����Ը�����,decode(�޼۷�ʽ,1,'ͳһ�޼�',2,'��ҽԺ�ȼ�����',3,'������ҽԺ��׼��������') �޼� " & _
                         "       ,��Ⱥ��Ŀ�ȼ�,��Ⱥ�Ը�����,��Ⱥ��Ŀ����,decode(������Ŀ��־,1,'��') ������Ŀ��־,��ע " & _
                         "   FROM �м��_������Ŀ Where ��Ŀ���� like '" & str������ & "%'"
            End If
        End If
    End If
    Call OpenRecordset_OtherBase(rsTemp, "ҽ��������ϸ", gstrSQL, mcnYB)
    
    If lvwDetail.ColumnHeaders.Count <> rsTemp.Fields.Count Then
        '���´����ͷ
        blnColSet = True
        lvwDetail.ColumnHeaders.Clear
        For Each fld In rsTemp.Fields
            lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
        Next
    End If
        
    Do Until rsTemp.EOF
        Set lst = lvwDetail.ListItems.Add(, "K" & rsTemp("ҽ������"), rsTemp("ҽ������"), "Detail", "Detail")
        
        '����ListView�����������ݿ�ȡ��
        For lngCol = 2 To lvwDetail.ColumnHeaders.Count
            varValue = rsTemp(lvwDetail.ColumnHeaders(lngCol).Text).Value
            lst.SubItems(lngCol - 1) = IIf(IsNull(varValue), "", varValue)
        Next
        rsTemp.MoveNext
    Loop
    If blnColSet = True Then
        '���¶��н����˴���
        If lvwDetail.ListItems.Count > 0 Then Call zlControl.LvwSetColWidth(lvwDetail)
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockWindowUpdate 0
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "������Ŀ"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "ҽ�����ࣺ" & tvwClass.SelectedItem.Text
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
    Dim StrInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnȫ�� As Boolean
    
    If MsgBox("���������ܻỨ�Ƚϳ���ʱ�䣬�Ƿ������" & vbCrLf & vbCrLf & "����ע�⣬������ֻ����ҽ����Ŀ��ϸ������������Ӧ��ϵ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    Me.Caption = "ҽ����Ŀѡ�����ڶ�ȡ���ļ��������ȡ������Ŀ��ϸ�����Ժ�......��"
    
    '��������ҽ�������� 204-04-07
    '��鱾����ȫ�����»�����������(�޸�����)
    gstrSQL = "Select 1 From zlcq.�м��_ҩƷĿ¼ where rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鱾����ȫ�����»�����������")
    blnȫ�� = (rsTemp.RecordCount = 0)
    If Not blnȫ�� Then
        If MsgBox("��Ҫ��ʼ�������أ����ȷ��������������أ����ȡ������ȫ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            blnȫ�� = True
        End If
    End If
    
    mcnYB.BeginTrans
    gcnOracle.BeginTrans
    
    'ɾ����־�ļ�
    Set mobjFileSystem = New FileSystemObject
    If mobjFileSystem.FileExists(strFile) Then mobjFileSystem.DeleteFile (strFile)
    Set mErrFile = mobjFileSystem.CreateTextFile(strFile)
    
    If Not AnalyFile_YPML(blnȫ��) Then
        mErrFile.Close
        MousePointer = vbDefault
        mcnYB.RollbackTrans
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    If Not AnalyFile_ZLML(blnȫ��) Then
        mErrFile.Close
        MousePointer = vbDefault
        mcnYB.RollbackTrans
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    If Not AnalyFile_BZML(blnȫ��) Then
        mErrFile.Close
        MousePointer = vbDefault
        mcnYB.RollbackTrans
        gcnOracle.RollbackTrans
        Exit Sub
    End If
    
    mErrFile.Close
    mcnYB.CommitTrans
    gcnOracle.CommitTrans
    
    '����װ����ϸ
    Call FillList
    MousePointer = vbDefault
    Me.Caption = "ҽ����Ŀѡ��"
End Sub

Private Sub Form_Load()
    cmdRequery.Visible = (mint���� = TYPE_����������)
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    
    lvwDetail.Top = tvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = tvwClass.Height
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
        If tvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
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

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
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

Private Function AnalyFile_YPML(Optional ByVal blnȫ�� As Boolean = True) As Boolean
    '�����ӿڷ��ص�ҩƷĿ¼�ļ��������浽�м��
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strDeal As String, StrInput As String
    Dim str���ʱ�� As String, intMode As Integer       '����ʹ�ã������ʱ�估������ʽ������ɾ���ģ�
    Dim intCol_In As Integer, intCols_In As Integer
    Dim str��ˮ�� As String, STRERR As String
    Dim arrCol
'    �������͡���1������2���޸�3��ɾ��
    Const int���ʱ�� As Integer = 23
    Const strFile_ȫ�� As String = "C:\CQYB_YH\YPML.txt"
    Const strFile_���� As String = "C:\CQYB_YH\TEMP.txt"
    Dim objStream As TextStream
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If blnȫ�� Then
        StrInput = "|" & strFile_ȫ��
        Call ���ýӿ�_׼��_����������("02", StrInput)
    Else
        '��ȡ�����ı��ʱ�䣨������������أ��϶����ڼ�¼��
        gstrSQL = "Select Max(���ʱ��) ʱ�� From zlcq.�м��_ҩƷĿ¼"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ı��ʱ��")
        str���ʱ�� = Format(rsTemp!ʱ��, "yyyyMMdd HH:mm:ss")
        StrInput = str���ʱ�� & "|" & strFile_����
        Call ���ýӿ�_׼��_����������("18", StrInput)
    End If
    If Not ���ýӿ�_����������() Then Exit Function
    
    If Not mobjFileSystem.FileExists(IIf(blnȫ��, strFile_ȫ��, strFile_����)) Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile(IIf(blnȫ��, strFile_ȫ��, strFile_����), ForReading, False, TristateMixed)
    If blnȫ�� Then mcnYB.Execute "ZL_�м��_ҩƷĿ¼_DELETEALL()", , adCmdStoredProc
    
    StrInput = "ZL_�м��_ҩƷĿ¼_Insert("
    Do While Not mobjStream.AtEndOfStream
        strData = Replace(mobjStream.ReadLine, """", "")
        arrCol = Split(strData, vbTab)
        lngCols = UBound(arrCol)
        strDeal = ""
        For lngCol = 0 To lngCols
            '�������ȫ��,�������һ���ֶ�,���жϲ�������
            If Not blnȫ�� And lngCol = lngCols Then
                '���û�ȡĿ¼��ϸ��Ϣ�������ݲ�������
                intMode = IIf(arrCol(1) = "INSERT", 1, IIf(arrCol(1) = "UPDATE", 2, 3))
                If intMode = 1 Or intMode = 2 Then
                    str��ˮ�� = arrCol(2)
                    StrInput = arrCol(2) & "|" & strFile_ȫ��
                    Call ���ýӿ�_׼��_����������("02", StrInput)
                    If ���ýӿ�_���������� Then
                        If mobjFileSystem.FileExists(strFile_ȫ��) Then
                            Set objStream = mobjFileSystem.OpenTextFile(strFile_ȫ��)
                            strData = Replace(objStream.ReadLine, """", "")
                            objStream.Close
                            arrCol = Split(strData, vbTab)
                            intCols_In = UBound(arrCol)
                            strDeal = ""
                            
                            For intCol_In = 0 To intCols_In
                                Select Case intCol_In
                                Case int���ʱ��
                                    '�������ڸ�ʽ��ͬ����Ҫת��
                                    strDate = ReplaceStr(arrCol(intCol_In))
                                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                                    strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                                    strDeal = strDeal & strDate
                                Case Else
                                    If Trim(arrCol(intCol_In)) = "" Then
                                        strDeal = strDeal & ",NULL"
                                    Else
                                        strDeal = strDeal & ",'" & ReplaceStr(arrCol(intCol_In)) & "'"
                                    End If
                                End Select
                            Next
                            
                            Select Case intMode
                            Case 1
                                strDeal = "ZL_�м��_ҩƷĿ¼_Insert(" & Mid(strDeal, 2) & ")"
                            Case 2
                                strDeal = "ZL_�м��_ҩƷĿ¼_Update(" & Mid(strDeal, 2) & ")"
                            End Select
                        End If
                    End If
                Else
                    strDeal = "ZL_�м��_ҩƷĿ¼_Delete('" & str��ˮ�� & "')"
                End If
            Else
                If blnȫ�� Then
                Select Case lngCol
                Case int���ʱ��
                    '�������ڸ�ʽ��ͬ����Ҫת��
                    strDate = ReplaceStr(arrCol(lngCol))
                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                    strDate = ",to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                    strDeal = strDeal & strDate
                Case Else
                    If Trim(arrCol(lngCol)) = "" Then
                        strDeal = strDeal & ",NULL"
                    Else
                        strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                    End If
                End Select
                End If
            End If
        Next
        If blnȫ�� Then strDeal = StrInput & Mid(strDeal, 2) & ")"
        mcnYB.Execute strDeal, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_YPML = True
    Exit Function
errHand:
    STRERR = "��ǰ��:" & mobjStream.Line - 1 & "�����:" & Err.Number & "������Ϣ:" & Err.Description
    mErrFile.WriteLine STRERR
    Resume Next
End Function

Private Function AnalyFile_ZLML(Optional ByVal blnȫ�� As Boolean = True) As Boolean
    '�����ӿڷ��ص�ҩƷĿ¼�ļ��������浽�м��
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim strData As String, strDate As String
    Dim strDeal As String, StrInput As String
    Dim str���ʱ�� As String, intMode As Integer       '����ʹ�ã������ʱ�估������ʽ������ɾ���ģ�
    Dim intCol_In As Integer, intCols_In As Integer
    Dim str��ˮ�� As String, STRERR As String
    Dim arrRow, arrCol
    Dim strHosCode As String, bln������Ŀ As Boolean
    
    Const int���ʱ�� As Integer = 14
    Const int������Ŀ���� As Integer = 17
    Const intҽ�ƻ������� As Integer = 16
    Const strFile_ȫ�� As String = "C:\CQYB_YH\ZLML.txt"
    Const strFile_���� As String = "C:\CQYB_YH\TEMP.txt"
    Dim objStream As TextStream
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If blnȫ�� Then
        StrInput = "|" & strFile_ȫ��
        Call ���ýӿ�_׼��_����������("03", StrInput)
    Else
        '��ȡ�����ı��ʱ�䣨������������أ��϶����ڼ�¼��
        gstrSQL = "Select Max(���ʱ��) ʱ�� From zlcq.�м��_������Ŀ"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ı��ʱ��")
        str���ʱ�� = Format(rsTemp!ʱ��, "yyyyMMdd HH:mm:ss")
        StrInput = str���ʱ�� & "|" & strFile_����
        Call ���ýӿ�_׼��_����������("19", StrInput)
    End If
    If Not ���ýӿ�_����������() Then Exit Function
    
    If Not mobjFileSystem.FileExists(IIf(blnȫ��, strFile_ȫ��, strFile_����)) Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile(IIf(blnȫ��, strFile_ȫ��, strFile_����), ForReading, False, TristateMixed)
    If blnȫ�� Then mcnYB.Execute "ZL_�м��_������Ŀ_DELETEALL()", , adCmdStoredProc
    
    StrInput = "ZL_�м��_������Ŀ_Insert("
    Do While Not mobjStream.AtEndOfStream
        strData = Replace(mobjStream.ReadLine, """", "")
        arrCol = Split(strData, vbTab)
        lngCols = UBound(arrCol)
        strDeal = ""
        For lngCol = 0 To lngCols
            If Not blnȫ�� And lngCol = lngCols Then
                '���û�ȡĿ¼��ϸ��Ϣ�������ݲ�������
                intMode = IIf(arrCol(1) = "INSERT", 1, IIf(arrCol(1) = "UPDATE", 2, 3))
                If intMode = 1 Or intMode = 2 Then
                    str��ˮ�� = arrCol(2)
                    StrInput = arrCol(2) & "|" & strFile_ȫ��
                    Call ���ýӿ�_׼��_����������("02", StrInput)
                    If ���ýӿ�_���������� Then
                        If mobjFileSystem.FileExists(strFile_ȫ��) Then
                            Set objStream = mobjFileSystem.OpenTextFile(strFile_ȫ��)
                            strData = Replace(objStream.ReadLine, """", "")
                            objStream.Close
                            arrCol = Split(strData, vbTab)
                            intCols_In = UBound(arrCol)
                            strDeal = ""
                            
                            For intCol_In = 0 To intCols_In
                                Select Case lngCol
                                Case int���ʱ��
                                    '�������ڸ�ʽ��ͬ����Ҫת��
                                    strDate = ReplaceStr(arrCol(lngCol))
                                    strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                                    strDate = ",to_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                                    strDeal = strDeal & strDate
                                Case intҽ�ƻ�������
                                    strHosCode = ReplaceStr(arrCol(lngCol))
                                    strDeal = strDeal & ",'" & Trim(arrCol(lngCol)) & "'"
                                Case int������Ŀ����
                                    bln������Ŀ = False
                                    If strHosCode = gComInfo_����������.ҽԺ���� Then
                                        If Val(arrCol(lngCol)) = 3 Then
                                            '������Ŀ
                                            bln������Ŀ = True
                                        End If
                                    End If
                                    strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                                Case Else
                                    strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                                End Select
                            Next
                        
                            Select Case intMode
                            Case 1
                                strDeal = "ZL_�м��_������Ŀ_Insert(" & Mid(strDeal, 2) & IIf(bln������Ŀ, ",1", "") & ")"
                            Case 2
                                strDeal = "ZL_�м��_������Ŀ_Update(" & Mid(strDeal, 2) & IIf(bln������Ŀ, ",1", "") & ")"
                            End Select
                        End If
                    End If
                Else
                    strDeal = "ZL_�м��_������Ŀ_Delete('" & str��ˮ�� & "')"
                End If
            Else
                If blnȫ�� Then
                    Select Case lngCol
                    Case int���ʱ��
                        '�������ڸ�ʽ��ͬ����Ҫת��
                        strDate = ReplaceStr(arrCol(lngCol))
                        strDate = Mid(strDate, 1, 4) & "-" & Mid(strDate, 5, 2) & "-" & Mid(strDate, 7)
                        strDate = ",to_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')"
                        strDeal = strDeal & strDate
                    Case intҽ�ƻ�������
                        strHosCode = ReplaceStr(arrCol(lngCol))
                        strDeal = strDeal & ",'" & Trim(arrCol(lngCol)) & "'"
                    Case int������Ŀ����
                        bln������Ŀ = False
                        If strHosCode = gComInfo_����������.ҽԺ���� Then
                            If Val(arrCol(lngCol)) = 3 Then
                                '������Ŀ
                                bln������Ŀ = True
                            End If
                        End If
                        strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                    Case Else
                        strDeal = strDeal & ",'" & ReplaceStr(arrCol(lngCol)) & "'"
                    End Select
                End If
            End If
        Next
        If blnȫ�� Then strDeal = StrInput & Mid(strDeal, 2) & IIf(bln������Ŀ, ",1", "") & ")"
        mcnYB.Execute strDeal, , adCmdStoredProc
    Loop
    mobjStream.Close
    
    AnalyFile_ZLML = True
    Exit Function
errHand:
    STRERR = "��ǰ��:" & mobjStream.Line - 1 & "�����:" & Err.Number & "������Ϣ:" & Err.Description
    mErrFile.WriteLine STRERR
    Resume Next
End Function

Private Function AnalyFile_BZML(Optional ByVal blnȫ�� As Boolean = True) As Boolean
    '�����ӿڷ��ص�ҩƷĿ¼�ļ��������浽�м��
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim str���� As String, str���� As String, str���� As String, str��� As String
    Dim strDeal As String, StrInput As String, strData As String
    Dim arrRow, arrCol
    Dim lngNextID As Long
    Dim str���ʱ�� As String, intMode As Integer         '1-����;2-�޸�;3-ɾ��
    Dim STRERR As String
    
    Const strFile_ȫ�� As String = "C:\CQYB_YH\BZML.txt"
    Dim rs���� As New ADODB.Recordset
    
    On Error GoTo errHand
    
    StrInput = strFile_ȫ��
    Call ���ýӿ�_׼��_����������("04", StrInput)
    If Not ���ýӿ�_����������() Then Exit Function
    
    If Not mobjFileSystem.FileExists(strFile_ȫ��) Then Exit Function
    Set mobjStream = mobjFileSystem.OpenTextFile(strFile_ȫ��, ForReading, False, TristateMixed)
    
    '�����в���
    gstrSQL = "Select ID,���� From ���ղ��� Where ����=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���в���Ŀ¼", TYPE_����������)
    
    Do While Not mobjStream.AtEndOfStream
        strData = Replace(mobjStream.ReadLine, """", "")
        arrCol = Split(strData, vbTab)
        
        str���� = ReplaceStr(arrCol(0))
        str���� = ReplaceStr(arrCol(1))
        str���� = ReplaceStr(arrCol(4))
        str��� = Val(arrCol(2)) - 1
        If Val(str���) < 0 Then str��� = 0
        
        With rs����
            .Filter = "����='" & str���� & "'"
            intMode = IIf(.RecordCount = 0, 1, 2)
        End With
        
        '���±��ռ���
        Select Case intMode
        Case 1
            lngNextID = zlDatabase.GetNextID("���ղ���")
            gstrSQL = "zl_���ղ���_INSERT(" & lngNextID & "," & TYPE_���������� & ",'" & str���� & _
                        "','" & str���� & "','" & str���� & "'," & str��� & ",NULL,NULL)"
        Case 2
            lngNextID = rs����!ID
            gstrSQL = "zl_���ղ���_UPDATE(" & lngNextID & ",'" & str���� & _
                        "','" & str���� & "','" & str���� & "'," & str��� & ",NULL,NULL)"
        Case Else
            lngNextID = rs����!ID
            gstrSQL = "zl_���ղ���_DELETE(" & lngNextID & ")"
        End Select
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Loop
    mobjStream.Close
    
    AnalyFile_BZML = True
    Exit Function
errHand:
    STRERR = "��ǰ��:" & mobjStream.Line - 1 & "�����:" & Err.Number & "������Ϣ:" & Err.Description
    mErrFile.WriteLine STRERR
    Resume Next
End Function

Private Function ReplaceStr(ByVal StrInput As String) As String
    ReplaceStr = Trim(Replace(StrInput, "'", ""))
    ReplaceStr = Replace(ReplaceStr, """", "")
End Function


