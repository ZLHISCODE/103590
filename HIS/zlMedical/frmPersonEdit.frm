VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonEdit 
   Caption         =   "�ܼ���Ա"
   ClientHeight    =   5865
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   10095
   Icon            =   "frmPersonEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk 
      Caption         =   "�����ͬʱ���б�������(&3)"
      Height          =   195
      Left            =   3600
      TabIndex        =   0
      Top             =   105
      Width           =   3150
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   675
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   2745
   End
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   75
      ScaleHeight     =   615
      ScaleWidth      =   10650
      TabIndex        =   9
      Top             =   4740
      Width           =   10650
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   8115
         TabIndex        =   12
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9315
         TabIndex        =   11
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   90
         TabIndex        =   10
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.Frame fra2 
      Height          =   3615
      Left            =   645
      TabIndex        =   7
      Top             =   315
      Width           =   6210
      Begin zl9Medical.VsfGrid vsfPerson 
         Height          =   3045
         Left            =   45
         TabIndex        =   8
         Top             =   150
         Width           =   4755
         _extentx        =   8387
         _extenty        =   5371
      End
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   11
      Left            =   8700
      Picture         =   "frmPersonEdit.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "����������Ϣ"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   14
      Left            =   8265
      Picture         =   "frmPersonEdit.frx":15AC
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "����Ϣд��IC��"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   15
      Left            =   7875
      Picture         =   "frmPersonEdit.frx":7DFE
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "��IC������Ϣ"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   16
      Left            =   8250
      Picture         =   "frmPersonEdit.frx":E650
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "��λ��Աѡ��"
      Top             =   525
      Width           =   345
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   5505
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPersonEdit.frx":14EA2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12726
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4020
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":15736
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1A7A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1AA9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1B034
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1B5CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1B728
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "&1.���"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   75
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "&2.��Ŀ"
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   405
      Width           =   540
   End
End
Attribute VB_Name = "frmPersonEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngLoop As Long
Private mblnDataChange As Boolean
Private mrsPersons As New ADODB.Recordset                 '������ʱ�����Ա
Private mblnChanged As Boolean
Private mblnGroup As Boolean
Private mlngGroup As Long
Private mstrGroup As String
Private mbytMode As Byte
Private mlngKey As Long
Private mblnNo As Boolean
Private mblnRegister As Boolean

Private Enum mPersonCol
    ���� = 1
    �����
    ������
    �Ա�
    ����
    ����״��
    ��������
    ���֤
    ����
    ����
    ѧ��
    ְҵ
    ���
    ��ϵ������
    ��ϵ�˵绰
    �����ʼ�
    ��ϵ�˵�ַ
    ������λ
    ����id
    IC����
    ���￨��
    ǰ��ɫ
    �¼�
End Enum

'�������Զ�����̻���************************************************************************************************
'������Աȱʡֵ
Private Function SetDefault(ByVal intRow As Integer) As Boolean
    
    '�Ȱ����ж�ȡ
    With vsfPerson
        If intRow > 1 Then
            .TextMatrix(intRow, mPersonCol.�Ա�) = .TextMatrix(intRow - 1, mPersonCol.�Ա�)
            .TextMatrix(intRow, mPersonCol.����״��) = .TextMatrix(intRow - 1, mPersonCol.����״��)
        End If
        
    End With
    
End Function

Private Function CountGroup() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:�����ͳ����Ŀ�������������С�Ů��
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    
    If mblnGroup Then
        strTmp = """" & cbo.Text & """�����"
    End If
    
    If mblnGroup Then
        lngCount1 = 0
        lngCount2 = 0
        
        For lngLoop = 1 To vsfPerson.Rows - 1
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.����)) <> "" Then
                If InStr(vsfPerson.TextMatrix(lngLoop, mPersonCol.�Ա�), "��") > 0 Then
                    lngCount1 = lngCount1 + 1
                Else
                    lngCount2 = lngCount2 + 1
                End If
            End If
        Next
        
        strTmp = strTmp & "������Ա" & lngCount1 + lngCount2 & "��(����:" & lngCount1 & "��,Ů��:" & lngCount2 & "��)"
    End If
    
    stbThis.Panels(2).Text = strTmp
    
End Function

Private Function CheckHavePerson(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsfPerson.Rows - 1
        If Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.����id)) = lngKey And vsfPerson.Row <> lngLoop And Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.����id)) > 0 Then
            CheckHavePerson = True
            Exit Function
        End If
    Next
End Function

Private Property Let DataChange(ByVal vData As Boolean)
        mblnDataChange = vData
End Property

Private Property Get DataChange() As Boolean
        DataChange = mblnDataChange
End Property

Private Function GetPatientInfo(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    strSQL = "SELECT A.* FROM ������Ϣ A WHERE A.����id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        

        If mlngGroup <> Val(zlCommFun.NVL(rs("��ͬ��λid"))) And Val(zlCommFun.NVL(rs("��ͬ��λid"))) > 0 And mlngGroup > 0 Then

            If MsgBox("���ǵ�ǰ�������Ա���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function

        End If
        
        vsfPerson.EditText = zlCommFun.NVL(rs("����"))
        vsfPerson.Cell(flexcpData, vsfPerson.Row, vsfPerson.Col) = zlCommFun.NVL(rs("����").Value)
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
        
        Call SetDefault(vsfPerson.Row)
        
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����"))
        
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��"))
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������) = Format(zlCommFun.NVL(rs("��������")), "yyyy-MM-dd")
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) = zlCommFun.NVL(rs("����id"))
        
        DataChange = True
        
    End If
    
    GetPatientInfo = True
    
End Function


Public Function ShowEdit(ByVal frmMain As Object, _
                        ByVal lngKey As Long, _
                        ByRef rsPersons As ADODB.Recordset, _
                        Optional blnGroup As Boolean = False, _
                        Optional ByVal bytMode As Byte = 1, _
                        Optional lngGroup As Long, _
                        Optional ByRef blnRegister As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant

    mblnNo = True
    mblnStartUp = True
    mblnOK = False

    Set mfrmMain = frmMain
    
    mstrGroup = ""
    
    mlngGroup = lngGroup
    mlngKey = lngKey
    Call CopyRecord(rsPersons, mrsPersons)
    mblnGroup = blnGroup
    mbytMode = bytMode

    Call ClearData
    If InitData = False Then Exit Function
    If ReadData() = False Then Exit Function

    DataChange = False
    
    mblnNo = False
    
    Call cbo_Click
    
    Call vsfPerson_AfterRowColChange(0, 0, vsfPerson.Row, vsfPerson.Col)
    
    Me.Show 1, frmMain

    rsPersons.Filter = ""
    If mblnOK Then Call CopyRecord(mrsPersons, rsPersons)
    blnRegister = mblnRegister
    
    ShowEdit = mblnOK

End Function

Private Function ClearData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    cbo.Clear
    Call ResetVsf(vsfPerson)

'    vsfPerson.AppendRow = True

    DataChange = False


End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHand
    
    chk.Visible = (mbytMode = 2)
    chk.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & Me.Name, "��������", "1"))
    
    cbo.AddItem "ȱʡ"

    With vsfPerson
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "����", 1080, 1, "...", 1, GetMaxLength("������Ϣ", "����")
        .NewColumn "�����", 810, 1
        .NewColumn "������", 900, 1, , 1
        .NewColumn "�Ա�", 750, 1, GetCombList("SELECT ���� FROM �Ա�"), 1, GetMaxLength("������Ϣ", "�Ա�")
        .NewColumn "����", 600, 1, , 1, GetMaxLength("������Ϣ", "����")
        .NewColumn "����״��", 900, 1, GetCombList("SELECT ���� FROM ����״��"), 1, GetMaxLength("������Ϣ", "����״��")
        .NewColumn "��������", 990, 1, , 1
        .NewColumn "���֤", 1800, 1, , 1, GetMaxLength("������Ϣ", "���֤��")
        .NewColumn "����", 0, 1, , , GetMaxLength("������Ϣ", "����")
        .NewColumn "����", 0, 1, , , GetMaxLength("������Ϣ", "����")
        .NewColumn "ѧ��", 0, 1, , , GetMaxLength("������Ϣ", "ѧ��")
        .NewColumn "ְҵ", 0, 1, , , GetMaxLength("������Ϣ", "ְҵ")
        .NewColumn "���", 0, 1, , , GetMaxLength("������Ϣ", "���")
        .NewColumn "��ϵ������", 0, 1, , , GetMaxLength("������Ϣ", "��ϵ������")
        .NewColumn "��ϵ�˵绰", 0, 1, , , GetMaxLength("������Ϣ", "��ϵ�˵绰")
        .NewColumn "�����ʼ�", 0, 1, , , GetMaxLength("������Ϣ", "�����ʼ�")
        .NewColumn "��ϵ�˵�ַ", 0, 1, , , GetMaxLength("������Ϣ", "��ϵ�˵�ַ")
        .NewColumn "������λ", 0, 1, , , GetMaxLength("������Ϣ", "������λ")
        .NewColumn "����id", 0, 1
        .NewColumn "IC����", 0, 1
        .NewColumn "���￨��", 0, 1
        
        .NewColumn "ǰ��ɫ", 0, 1
        .NewColumn "�¼�", 0, 1
'        .NewColumn "", 15, 1
'        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .Body.GridColorFixed = &HC1C1C1
'        .AppendRow = True
        
        .Body.ColEditMask(mPersonCol.��������) = "0000-00-00"
        
    End With

    If mblnGroup = False Then
        cbo.Visible = False
        lbl(4).Visible = False
    End If
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand


    '��ȡ�����������Ŀ
    
    mblnNo = True
    
    cbo.Clear

    gstrSQL = "SELECT A.������� AS ���, rownum AS ID FROM ������ A WHERE A.�Ǽ�id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo.AddItem rs("���").Value
            rs.MoveNext
        Loop
    Else
        cbo.AddItem "ȱʡ"
    End If

    '��ȡ�����Ŀ
    
    If cbo.ListCount > 0 Then cbo.ListIndex = 0
    
    mblnNo = False
    
    Call cbo_Click

    ReadData = True

    Exit Function

errHand:

    If ErrCenter = 1 Then Resume

End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    For lngLoop = 1 To vsfPerson.Rows - 1
        If Val(vsfPerson.RowData(lngLoop)) = lngKey And vsfPerson.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function


Private Function SaveItems(ByVal strGroup As String) As Boolean

    Dim lngLoop As Long

    On Error GoTo errHand

    '������ѡ��ļ�����Ŀ
    mrsPersons.Filter = ""
    mrsPersons.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"

    Call DeleteRecord(mrsPersons)

    For lngLoop = 1 To vsfPerson.Rows - 1

        If vsfPerson.TextMatrix(lngLoop, mPersonCol.����) <> "" Then
            mrsPersons.AddNew

            mrsPersons("���").Value = strGroup
            mrsPersons("����id").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.����id)
            mrsPersons("IC����").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.IC����)
            mrsPersons("������").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.������)
            mrsPersons("����").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.����)
            mrsPersons("�����").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.�����)
            mrsPersons("���֤").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.���֤)
            mrsPersons("�Ա�").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.�Ա�)
            mrsPersons("��������").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.��������)
            mrsPersons("����״��").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.����״��)
            mrsPersons("����").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.����)
            mrsPersons("����").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.����)
            mrsPersons("����").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.����)
            mrsPersons("ѧ��").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.ѧ��)
            mrsPersons("ְҵ").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.ְҵ)
            mrsPersons("���").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.���)
            mrsPersons("��ϵ������").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.��ϵ������)
            mrsPersons("��ϵ�˵绰").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.��ϵ�˵绰)
            mrsPersons("�����ʼ�").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.�����ʼ�)
            mrsPersons("��ϵ�˵�ַ").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.��ϵ�˵�ַ)
            mrsPersons("������λ").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.������λ)
'            mrsPersons("���￨��").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.���￨��)
            mrsPersons("�¼�").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.�¼�)
            mrsPersons("ǰ��ɫ").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.ǰ��ɫ)

        End If
    Next

    SaveItems = True

errHand:

End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    For lngLoop = 1 To vsfPerson.Rows - 1
        
        If vsfPerson.EditMode(mPersonCol.����) = 1 Then
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.����), GetMaxLength("������Ϣ", "����")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.����
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.���֤), GetMaxLength("������Ϣ", "���֤��")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.���֤
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
                        
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.����״��), GetMaxLength("������Ϣ", "����״��")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.����״��
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.�����ʼ�), GetMaxLength("�����Ա����", "�����ʼ�")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.�����ʼ�
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.�Ա�), GetMaxLength("������Ϣ", "�Ա�")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.�Ա�
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.��������)) <> "" Then
                
                If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.��������), CHECKFORMAT.����) = False Then
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.��������
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
                End If
            End If
            
                        
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.�����ʼ�), CHECKFORMAT.�����ʼ�) = False Then
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.�����ʼ�
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
                
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.���֤), CHECKFORMAT.���֤��) = False Then
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.���֤
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
        End If
    Next
    
    ValidEdit = True

End Function

Private Function ReadItems(ByVal strGroup As String) As Boolean

    mrsPersons.Filter = ""
    mrsPersons.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"
    If mrsPersons.RecordCount > 0 Then
        mrsPersons.MoveFirst
        Call FillGrid(vsfPerson, mrsPersons)
    End If

    ReadItems = True

End Function

Private Sub cbo_Click()
    If mblnNo Then Exit Sub
    
    If mstrGroup <> cbo.Text Then
        Call SaveItems(mstrGroup)
        
        mstrGroup = cbo.Text
        
        Call ResetVsf(vsfPerson)
        Call ReadItems(mstrGroup)
    End If
    
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

        zlCommFun.PressKey vbKeyTab

    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim clsCard As Object
    Dim strInfo() As String
    Dim lngLoop As Long
    Dim strParam As String
    Dim varParam As Variant
    Dim strItem As String
    Dim strValue As String
    Dim strCardNo1 As String
    Dim strCardNo2 As String
    
    On Error GoTo errHand
    Select Case Index
    Case 11
        
        strParam = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) & "'"
        strParam = strParam & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������), "yyyy-MM-dd") & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) & "'"
        
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������)
                
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(1)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) = varParam(2)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = varParam(3)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������) = varParam(4)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) = varParam(5)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) = Val(varParam(0))
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(6)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(7)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) = varParam(8)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) = varParam(9)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) = varParam(10)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) = varParam(11)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) = varParam(12)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) = varParam(14)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) = varParam(15)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(16)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������) = varParam(17)
            
            
        End If
        
        If vsfPerson.Visible Then vsfPerson.SetFocus
        
    Case 14    'д��
                    
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            ReDim strInfo(1 To 16)
            
            strCardNo1 = clsCard.GetCardNo
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����)
            If strCardNo2 <> "" Then
                '�����п������͵�ǰ�Ŀ�����ͬһ�ſ�
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "�˿����ǵ�ǰ���˵Ŀ���"
                    Exit Sub
                End If
            Else
                '����û�п�
                
                If strCardNo1 = "" Then
                
                    '�¿����Զ�����
                    strCardNo1 = "11111111"
                    strCardNo2 = strCardNo1
                    
                    'д����
                    If clsCard.SetCardNo(strCardNo1) = False Then Exit Sub
                    vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����) = strCardNo2
                    
                Else
                
                    '�����¿�
                    ShowSimpleMsg "�˿������¿������ܽ���д�������"
                    Exit Sub
                End If
            End If
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����) = strCardNo2
            
            strInfo(1) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
            strInfo(2) = "���֤��=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤)
            strInfo(3) = "�Ա�=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�)
            strInfo(4) = "��������=" & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������), "yyyy-MM-dd")
            strInfo(5) = "����״��=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��)
            strInfo(6) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
            strInfo(7) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
            strInfo(8) = "ѧ��=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��)
            strInfo(9) = "ְҵ=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ)
            strInfo(10) = "���=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���)
            strInfo(11) = "��ϵ������=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������)
            strInfo(12) = "��ϵ�˵绰=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰)
            strInfo(13) = "��ϵ�˵�ַ=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ)
            strInfo(14) = "�����ʼ�=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�)
            strInfo(15) = "������λ=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ)
            strInfo(16) = "����=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����)
                                    
            If clsCard.SetPatient(strInfo) Then
                ShowSimpleMsg "���µ�ǰ������Ϣ�ɹ���"
            End If
        End If
        If vsfPerson.Visible Then vsfPerson.SetFocus
    Case 15    '����
        
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            strCardNo1 = clsCard.GetCardNo
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����)
            
            If strCardNo2 <> "" Then
                '��¼�Ĳ����п������͵�ǰ�Ŀ�����ͬһ�ſ�
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "�˿����ǵ�ǰ���˵Ŀ���"
                    Exit Sub
                End If
            Else
            
                '����û�п����򽫵�ǰ�Ŀ��Ÿ�������
                strCardNo2 = strCardNo1
                                
            End If
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC����) = strCardNo2
            
            If GetPatientID(strCardNo2) > 0 Then
                
                '��ϵͳ���ҵ��˲���
                
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) = GetPatientID(strCardNo2)
                Call GetPatientInfo(Val(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id)))
                
            ElseIf clsCard.GetPatient(strInfo) Then
                For lngLoop = LBound(strInfo) To UBound(strInfo)
                    If InStr(strInfo(lngLoop), "=") > 0 Then
                        strItem = Mid(strInfo(lngLoop), 1, InStr(strInfo(lngLoop), "=") - 1)
                        strValue = Mid(strInfo(lngLoop), InStr(strInfo(lngLoop), "=") + 1)
                        
                        Select Case strItem
                        Case "����"
                        
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            
                        Case "���֤��"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) = strValue
                                                        
                        Case "�Ա�"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = strValue
                            
                        Case "��������"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������) = strValue
                            
                        Case "����״��"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) = strValue
                            
                        Case "����"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            
                        Case "����"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            
                        Case "ѧ��"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) = strValue
                            
                        Case "ְҵ"
                        
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) = strValue
                            
                        Case "���"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) = strValue
                            
                        Case "��ϵ������"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) = strValue
                            
                        Case "��ϵ�˵绰"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) = strValue
                            
                        Case "��ϵ�˵�ַ"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) = strValue
                            
                        Case "�����ʼ�"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) = strValue
                            
                        Case "������λ"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) = strValue
                            
                        Case "����"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = strValue
                            
                        End Select
                    End If
                Next
                
                
            End If
        End If
        If vsfPerson.Visible Then vsfPerson.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case 16
        'ѡ��λ��Ա
        Dim rsData As New ADODB.Recordset
        Dim rs As New ADODB.Recordset
        
        If frmSelectGroupPerson.ShowFilter(Me, mlngGroup, rs) Then
            rs.Filter = 0
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then

                If Val(vsfPerson.RowData(1)) > 0 Then
                    If MsgBox("�Ƿ�Ҫ�����ѡ����ܼ���Ա��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call ResetVsf(vsfPerson)
                    End If
                End If

                rs.MoveFirst

                Do While Not rs.EOF
                    
                    If CheckHavePerson(rs("ID").Value) = False Then
                        With vsfPerson
                        
                            .Row = .Rows - 1
                            If Val(.RowData(.Row)) > 0 Then
                                .Rows = .Rows + 1
                                .Row = .Rows - 1
                            End If
            
                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����").Value)
                            .TextMatrix(.Row, mPersonCol.������) = zlCommFun.NVL(rs("������").Value)
                            .TextMatrix(.Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
                            .TextMatrix(.Row, mPersonCol.��������) = zlCommFun.NVL(rs("��������").Value)
                            .TextMatrix(.Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��").Value)
                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(.Row, mPersonCol.ѧ��) = zlCommFun.NVL(rs("ѧ��").Value)
                            .TextMatrix(.Row, mPersonCol.ְҵ) = zlCommFun.NVL(rs("ְҵ").Value)
                            .TextMatrix(.Row, mPersonCol.���) = zlCommFun.NVL(rs("���").Value)
                            .TextMatrix(.Row, mPersonCol.��ϵ������) = zlCommFun.NVL(rs("��ϵ������").Value)
                            .TextMatrix(.Row, mPersonCol.��ϵ�˵绰) = zlCommFun.NVL(rs("��ϵ�˵绰").Value)
'                            .TextMatrix(.Row, mPersonCol.�����ʼ�) = zlCommFun.NVL(rs("�����ʼ�").Value)
                            .TextMatrix(.Row, mPersonCol.��ϵ�˵�ַ) = zlCommFun.NVL(rs("��ϵ�˵�ַ").Value)
                            .TextMatrix(.Row, mPersonCol.������λ) = zlCommFun.NVL(rs("������λ").Value)
                            .TextMatrix(.Row, mPersonCol.����id) = zlCommFun.NVL(rs("ID").Value, 0)
                            .TextMatrix(.Row, mPersonCol.IC����) = zlCommFun.NVL(rs("IC����").Value)
                            .TextMatrix(.Row, mPersonCol.���￨��) = zlCommFun.NVL(rs("���￨��").Value)
        
                            .RowData(.Row) = zlCommFun.NVL(rs("ID").Value)
                        End With
                    End If

                    rs.MoveNext

                Loop

                DataChange = True
            End If

        End If

        Call EnterFocus(vsfPerson)
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = -1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()

    Dim lngKey As Long

    If Trim(cbo.Text) <> "" Then Call SaveItems(Trim(cbo.Text))

    If ValidEdit = False Then Exit Sub

    mrsPersons.Filter = ""

    mblnOK = True
    DataChange = False
    mblnRegister = (chk.Value = 1)
    
    Unload Me

End Sub


Private Sub Form_Load()

    glngFormW = 10770
    glngFormH = 6780
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
       
    With fra2
        .Left = 0
        .Top = -90 + IIf(cbo.Visible, cbo.Height + cbo.Top + 30, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - picButton.Height + 90 - stbThis.Height
    End With
       
    With picButton
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height
        .Width = fra2.Width
    End With
    
    
    With vsfPerson
        .Left = 45
        .Top = 120
        .Width = fra2.Width - .Left - 45
        .Height = fra2.Height - .Top - 45
    End With
    
    With cmd(11)
        .Left = Me.ScaleWidth - .Width - 45
    End With
    
    With cmd(16)
        .Left = cmd(11).Left - .Width - 45
        .Top = cmd(11).Top
    End With
    
    With cmd(14)
        .Left = cmd(16).Left - .Width - 45
    End With
    
    With cmd(15)
        .Left = cmd(14).Left - .Width - 45
    End With
      
    cmdCancel.Left = picButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If DataChange Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If

    Call SaveWinState(Me, App.ProductName)
    SaveSetting "ZLSOFT", "˽��ģ��\" & Me.Name, "��������", chk.Value
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim strParam As String
    Dim varParam As Variant
        
    On Error GoTo errHand

    Select Case Button.Key
    Case "��ϸ����"
        
        strParam = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) & "'"
        strParam = strParam & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������), "yyyy-MM-dd") & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) & "'"
        
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) & "'"
        
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������)
                        
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(1)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���֤) = varParam(2)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�Ա�) = varParam(3)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��������) = varParam(4)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����״��) = varParam(5)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����id) = Val(varParam(0))
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(6)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(7)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ѧ��) = varParam(8)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.ְҵ) = varParam(9)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.���) = varParam(10)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ������) = varParam(11)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵绰) = varParam(12)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.�����ʼ�) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.��ϵ�˵�ַ) = varParam(14)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������λ) = varParam(15)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.����) = varParam(16)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.������) = varParam(17)
            
            
        End If

    End Select

    Exit Sub

errHand:
        If ErrCenter = 1 Then Resume
End Sub


Private Sub vsfPerson_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngPos As Long
    
    If Col = mPersonCol.�������� Then
        If Trim(vsfPerson.TextMatrix(Row, Col)) <> "" Then
        
            vsfPerson.TextMatrix(Row, Col) = zlCommFun.AddDate(vsfPerson.TextMatrix(Row, Col))
            
            If IsDate(vsfPerson.TextMatrix(Row, Col)) = False Then
                vsfPerson.TextMatrix(Row, Col) = ""
            End If
        End If
    End If
    
    
    If Col = mPersonCol.�����ʼ� Then
        If Trim(vsfPerson.TextMatrix(Row, Col)) <> "" Then
            
            lngPos = InStr(vsfPerson.TextMatrix(Row, Col), "@")
            
            If lngPos = 0 Then
                vsfPerson.TextMatrix(Row, Col) = ""
            Else
                If Trim(Mid(vsfPerson.TextMatrix(Row, Col), 1, lngPos - 1)) = "" Then
                    vsfPerson.TextMatrix(Row, Col) = ""
                ElseIf Trim(Mid(vsfPerson.TextMatrix(Row, Col), lngPos + 1)) = "" Then
                    vsfPerson.TextMatrix(Row, Col) = ""
                End If
            End If
        End If
    End If
    
    If Col = mPersonCol.�Ա� Then
        Call CountGroup
    End If
    
End Sub

Private Sub vsfPerson_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    '���ñ༭״̬
    If Val(vsfPerson.TextMatrix(NewRow, mPersonCol.�¼�)) = 1 Then
        
        If vsfPerson.EditMode(mPersonCol.����) <> 0 Then
            vsfPerson.EditMode(mPersonCol.����) = 0
            vsfPerson.EditMode(mPersonCol.���֤) = 0
            vsfPerson.EditMode(mPersonCol.��������) = 0
            vsfPerson.EditMode(mPersonCol.�Ա�) = 0
            vsfPerson.EditMode(mPersonCol.����) = 0
            vsfPerson.EditMode(mPersonCol.����״��) = 0
            
            vsfPerson.ComboList(mPersonCol.�Ա�) = ""
            vsfPerson.ComboList(mPersonCol.����״��) = ""
            vsfPerson.ComboList(mPersonCol.����) = ""
            
            cmd(11).Enabled = False
            cmd(14).Enabled = False
            cmd(15).Enabled = False
        End If
        
    Else
        If vsfPerson.EditMode(mPersonCol.����) <> 1 Then
            vsfPerson.EditMode(mPersonCol.����) = 1
            vsfPerson.EditMode(mPersonCol.���֤) = 1
            vsfPerson.EditMode(mPersonCol.����) = 1
            vsfPerson.EditMode(mPersonCol.��������) = 1
            vsfPerson.EditMode(mPersonCol.�Ա�) = 1
            vsfPerson.EditMode(mPersonCol.����״��) = 1
            
            vsfPerson.ComboList(mPersonCol.����) = "..."
            vsfPerson.ComboList(mPersonCol.�Ա�) = GetCombList("SELECT ���� FROM �Ա�")
            vsfPerson.ComboList(mPersonCol.����״��) = GetCombList("SELECT ���� FROM ����״��")
            cmd(11).Enabled = True
            cmd(14).Enabled = True
            cmd(15).Enabled = True
        End If
      
    End If
End Sub

Private Sub vsfPerson_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Val(vsfPerson.TextMatrix(Row, mPersonCol.�¼�)) = 1 And mbytMode = 2 Then
        
        Cancel = True
        Exit Sub
        
    End If
    
End Sub

Private Sub vsfPerson_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error Resume Next
    
    If Val(vsfPerson.TextMatrix(NewRow, mPersonCol.�����)) > 0 Then
        vsfPerson.EditMode(mPersonCol.�����) = 0
    Else
        vsfPerson.EditMode(mPersonCol.�����) = 1
    End If
End Sub

Private Sub vsfPerson_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset

    If frmPatientFind.ShowFind(Me, lngKey) Then
        If lngKey > 0 Then

            gstrSQL = "SELECT A.* FROM ������Ϣ A WHERE A.����id=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then

                If mlngGroup <> Val(zlCommFun.NVL(rs("��ͬ��λid"))) And Val(zlCommFun.NVL(rs("��ͬ��λid"))) > 0 And mlngGroup > 0 Then

                    If MsgBox("���ǵ�ǰ�������Ա���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

                End If
                
                vsfPerson.EditText = zlCommFun.NVL(rs("����"))
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("����").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                
                Call SetDefault(Row)
                
                vsfPerson.TextMatrix(Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����"))
                vsfPerson.TextMatrix(Row, mPersonCol.������) = zlCommFun.NVL(rs("������"))
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                vsfPerson.TextMatrix(Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��"))
                vsfPerson.TextMatrix(Row, mPersonCol.��������) = Format(zlCommFun.NVL(rs("��������")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����id) = zlCommFun.NVL(rs("����id"))
                
                DataChange = True

            End If

        End If
    End If

End Sub

Private Sub vsfPerson_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strText As String
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strSvrText As String
    Dim rsData As New ADODB.Recordset
    Dim blnCard As Boolean
    
    If Chr(KeyCode) = "'" Then KeyCode = 0
    
    If Col = mPersonCol.���� Then
        
        strText = vsfPerson.EditText
        If KeyCode <> 8 And KeyCode <> 13 Then
            strText = strText & Chr(KeyCode)
        End If
        
        If InStr(vsfPerson.EditText, "'") > 0 Then
            KeyCode = 0
            ShowSimpleMsg "�ڸ����������зǷ��ַ� ' ��"
            vsfPerson.EditText = ""
            vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
        
        blnCard = InputIsCard(vsfPerson.EditText, KeyCode)
        
        If blnCard And Len(vsfPerson.EditText) = ParamInfo.���￨���볤�� - 1 And KeyCode <> 8 And KeyCode <> vbKeyReturn Then
            vsfPerson.Body.EditSelStart = Len(vsfPerson.EditText)
            strInput = strInput & " AND C.���￨��=[1] "
        End If
        
        If KeyCode = vbKeyReturn Then
            If blnCard Then
                '�Ǿ��￨
                strInput = strInput & " AND C.���￨��=[1] "
            Else
                '�Ǿ��￨
                blnCard = False

                strText = vsfPerson.EditText
    
                Select Case UCase(Left(strText, 1))
                Case "-", "A"                 '����id,���￨��
                    strInput = strInput & " AND C.����id=[1]"
                Case "+", "B"                 'סԺ��
                    strInput = " AND C.סԺ��=[1]"
                Case "*", "D"                 '�����
                    strInput = strInput & " AND C.�����=[1]"
                Case "/", "C"                 '��ǰ����
                    strInput = strInput & " AND C.��ǰ����=[1]"
                Case Else
                    strSvrText = vsfPerson.Cell(flexcpData, Row, Col)
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                End Select
            End If
        End If
            
        If strInput <> "" Then
            gstrSQL = GetPublicSQL(SQL.��Ա����ѡ��, strInput)
            
            If blnCard Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strText))
            ElseIf UCase(Left(strText, 1)) = "/" Or UCase(Left(strText, 1)) = "C" Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Mid(strText, 2)))
            Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(strText, 2)))
            End If
            
            If ShowGrdFilter(Me, vsfPerson, "����,1200,0,0;�Ա�,810,0,0;��������,1200,0,0;����״��,900,0,0;���֤��,1500,0,0", Me.Name & "\��Ա����ѡ��Grid", "�������ѡ��һ����Ա", rsData, rs, , , , False) Then
                                                                        
                vsfPerson.EditText = zlCommFun.NVL(rs("����"))
                strText = vsfPerson.EditText
                If mlngGroup <> Val(zlCommFun.NVL(rs("��ͬ��λid"))) And Val(zlCommFun.NVL(rs("��ͬ��λid"))) > 0 And mlngGroup > 0 Then

                    If MsgBox("���ǵ�ǰ�������Ա���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        KeyCode = 0
                        vsfPerson.EditText = ""
                        vsfPerson.TextMatrix(Row, Col) = strSvrText
                        Cancel = True
                        Exit Sub
                    End If
                End If
    
                If CheckHavePerson(Val(zlCommFun.NVL(rs("ID")))) Then
                    ShowSimpleMsg "���ˡ�" & zlCommFun.NVL(rs("����").Value) & "���Ѿ����ڣ�"
                    KeyCode = 0
                    vsfPerson.EditText = ""
                    vsfPerson.TextMatrix(Row, Col) = strSvrText
                    Cancel = True
                    Exit Sub
                End If
    
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("����").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                Call SetDefault(Row)
                vsfPerson.TextMatrix(Row, mPersonCol.���֤) = zlCommFun.NVL(rs("���֤��"))
                vsfPerson.TextMatrix(Row, mPersonCol.��������) = Format(zlCommFun.NVL(rs("��������")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.�Ա�) = zlCommFun.NVL(rs("�Ա�").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����״��) = zlCommFun.NVL(rs("����״��").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.����id) = zlCommFun.NVL(rs("ID"))
                vsfPerson.TextMatrix(Row, mPersonCol.����) = zlCommFun.NVL(rs("����"))
                vsfPerson.TextMatrix(Row, mPersonCol.�����) = zlCommFun.NVL(rs("�����"))
                vsfPerson.TextMatrix(Row, mPersonCol.������) = zlCommFun.NVL(rs("������"))
                                
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.��ɫ
                
                vsfPerson.EditMode(mPersonCol.�����) = 0
    
                If blnCard Then
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                    vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                    KeyCode = 13
                End If
                
                DataChange = True
            Else
                'ȡ���˱���ѡ��
                    
                vsfPerson.EditMode(mPersonCol.�����) = 1
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.��ɫ
                
                vsfPerson.Cell(flexcpData, Row, Col) = vsfPerson.EditText
                vsfPerson.EditText = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, mPersonCol.�����) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.���֤) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.����id) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.��������) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.����) = ""
                Call SetDefault(Row)
                
            End If
        ElseIf KeyCode = vbKeyReturn Then
            '�²��ˣ��������������
            
            vsfPerson.EditMode(mPersonCol.�����) = 1
            vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.��ɫ
            vsfPerson.TextMatrix(Row, mPersonCol.����id) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.�����) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.���֤) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.��������) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.����) = ""
            
            Call SetDefault(Row)
        End If
    End If
End Sub

Private Sub vsfPerson_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)

    If KeyAscii = vbKeyReturn Then

        If Col = 1 Then
            If Trim(vsfPerson.TextMatrix(Row, Col)) = "" Then
                
                KeyAscii = 0
                
                cmdOK.SetFocus
                Cancel = True

            End If
        End If
    End If

End Sub

Private Sub vsfPerson_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    Select Case Col
    Case mPersonCol.�����
        '���������Ƿ����
        If Val(vsfPerson.EditText) > 0 Then
            gstrSQL = "Select 1 From ������Ϣ Where �����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsfPerson.EditText))
            If rs.BOF = False Then
                '����
                Cancel = True
                
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.EditText
                
                ShowSimpleMsg "��ǰ����ţ�" & Val(vsfPerson.EditText) & "�Ѿ����ڣ��������ظ���"
                vsfPerson.EditText = ""
                vsfPerson.TextMatrix(Row, Col) = ""
                
            End If
        End If
    End Select
End Sub


