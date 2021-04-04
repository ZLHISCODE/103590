VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatholAntibody_AntiUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ά��"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   Icon            =   "frmPatholAntibody_AntiUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   3615
      TabIndex        =   33
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.TextBox txtAlredyCount 
      Height          =   300
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtUseCount 
      Height          =   300
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   1
      Top             =   915
      Width           =   1695
   End
   Begin VB.CheckBox chkContinue 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȷ�Ϻ�������"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtApplySituation 
      Height          =   300
      Left            =   4200
      TabIndex        =   9
      Top             =   2915
      Width           =   2025
   End
   Begin VB.ComboBox cbxActionObject 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   2915
      Width           =   2025
   End
   Begin VB.ComboBox cbxLieracType 
      Height          =   300
      Left            =   4200
      TabIndex        =   7
      Top             =   2415
      Width           =   2025
   End
   Begin VB.ComboBox cbxCloneType 
      Height          =   300
      ItemData        =   "frmPatholAntibody_AntiUpdate.frx":179A
      Left            =   1080
      List            =   "frmPatholAntibody_AntiUpdate.frx":179C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2415
      Width           =   2025
   End
   Begin VB.TextBox txtMemo 
      Height          =   780
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3915
      Width           =   5145
   End
   Begin VB.CommandButton cmdNewAntibody_Cancel 
      Caption         =   "ȡ ��(&C)"
      Height          =   400
      Left            =   5040
      TabIndex        =   14
      Top             =   5235
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewAntibody_Sure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   3720
      TabIndex        =   13
      Top             =   5235
      Width           =   1215
   End
   Begin VB.TextBox txtAntibodyName 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   915
      Width           =   1815
   End
   Begin VB.ComboBox cbxValidCount 
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   1915
      Width           =   1665
   End
   Begin VB.TextBox txtRegisterDoctor 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   10
      Top             =   3415
      Width           =   2025
   End
   Begin MSComCtl2.DTPicker dtpMadeDate 
      Height          =   300
      Left            =   4200
      TabIndex        =   3
      Top             =   1415
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   99745795
      CurrentDate     =   40646.4399652778
   End
   Begin MSComCtl2.DTPicker dtpOverdueDate 
      Height          =   300
      Left            =   4200
      TabIndex        =   5
      Top             =   1915
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   99745795
      CurrentDate     =   40646.4399652778
   End
   Begin MSComCtl2.DTPicker dtpRegisterTime 
      Height          =   300
      Left            =   4200
      TabIndex        =   11
      Top             =   3415
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   99745795
      CurrentDate     =   40646.4399652778
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   32
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   31
      Top             =   2935
      Width           =   900
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ö���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   2935
      Width           =   900
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ʣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   29
      Top             =   2445
      Width           =   900
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�� ¡ �ԣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   2445
      Width           =   900
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2880
      TabIndex        =   27
      Top             =   1960
      Width           =   180
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ���˷ݣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   26
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������ƣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  ��  ע��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6360
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "    ����ȷ¼�뿹��������Ϣ���к�ɫ�Ǻű�ǵ�Ϊ��¼���ݣ�����ӹ����У��������ݲ�����ϵͳҪ��ģ�ϵͳ��������ʾ��"
      Height          =   495
      Left            =   840
      TabIndex        =   23
      Top             =   195
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   120
      Picture         =   "frmPatholAntibody_AntiUpdate.frx":179E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����˷ݣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   1465
      Width           =   900
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������ڣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   21
      Top             =   1465
      Width           =   900
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�� Ч �ڣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   1955
      Width           =   900
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������ڣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   19
      Top             =   1955
      Width           =   900
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� �ˣ�"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ǽ�ʱ�䣺"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   17
      Top             =   3425
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "frmPatholAntibody_AntiUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mufgParentGrid As ucFlexGrid
Private mblnIsSucceed As Boolean
Private mblnIsUpdate As Boolean

Property Get IsSucceed() As Boolean
    IsSucceed = mblnIsSucceed
End Property

Property Get IsUpdate() As Boolean
    IsUpdate = mblnIsUpdate
End Property

Property Let IsUpdate(value As Boolean)
    mblnIsUpdate = value
End Property

Public Function ShowAddAntibodyWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'��ʾ������������
    Dim curDate As Date
    
    ShowAddAntibodyWindow = False
    
    Set mufgParentGrid = ufgParentGrid
    
    Me.Caption = "��������"
    mblnIsUpdate = False
    
    curDate = zlDatabase.Currentdate
    
    dtpMadeDate.value = curDate
    dtpOverdueDate.value = curDate
    dtpRegisterTime.value = curDate
    txtRegisterDoctor.Text = UserInfo.����
    
    Call CloseProcessHint
    
    chkContinue.value = False
    chkContinue.Visible = True
    
    Call Me.Show(1, owner)
    
    ShowAddAntibodyWindow = mblnIsSucceed

End Function


Public Function ShowUpdateAntibodyWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'��ʾ������´���
    ShowUpdateAntibodyWindow = False
    
    Set mufgParentGrid = ufgParentGrid
        
    Me.Caption = "���¿���"
    mblnIsUpdate = True
        
    Call CloseProcessHint
    
    Call ConfigUpdateFace
    
    chkContinue.value = False
    chkContinue.Visible = False

    
    Call Me.Show(1, owner)
    
    ShowUpdateAntibodyWindow = mblnIsSucceed
End Function


Private Function GetCloneTypeIndex(ByVal strCloneValue As String) As Long
'ȡ�õ�ǰ��¡����
    GetCloneTypeIndex = 0
    
    If strCloneValue = "���¡" Then
        GetCloneTypeIndex = 1
    End If
End Function

Public Sub ConfigUpdateFace()
    
    With mufgParentGrid
        txtAntibodyName.Text = .Text(.SelectionRow, gstrAntibody_��������)
        txtUseCount.Text = .Text(.SelectionRow, gstrAntibody_ʹ���˷�)
        txtAlredyCount.Text = .Text(.SelectionRow, gstrAntibody_�����˷�)
        dtpMadeDate.value = .Text(.SelectionRow, gstrAntibody_��������)
        cbxValidCount.Text = Val(.Text(.SelectionRow, gstrAntibody_��Ч��))
        dtpOverdueDate.value = .Text(.SelectionRow, gstrAntibody_��������)
        cbxCloneType.ListIndex = GetCloneTypeIndex(.Text(.SelectionRow, gstrAntibody_��¡��))
        cbxLieracType.Text = .Text(.SelectionRow, gstrAntibody_������)
        cbxActionObject.Text = .Text(.SelectionRow, gstrAntibody_���ö���)
        txtApplySituation.Text = .Text(.SelectionRow, gstrAntibody_Ӧ�����)
        txtRegisterDoctor.Text = .Text(.SelectionRow, gstrAntibody_�Ǽ���)
        dtpRegisterTime.value = .Text(.SelectionRow, gstrAntibody_�Ǽ�ʱ��)
        txtMemo.Text = .Text(.SelectionRow, gstrAntibody_��ע)
    End With
    
    '�жϸÿ����Ƿ��ѱ�ʹ�ù�������ѱ�ʹ�ã���ĳЩ��Ϣ���ܽ��и���
    Dim strSQL As String
    Dim rsUsed As ADODB.Recordset
    
    
    strSQL = "select 1 from �����ؼ���Ϣ where ����ID=[1]"
    Set rsUsed = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow))
    
    If rsUsed.RecordCount <= 0 Then Exit Sub
    
    txtAntibodyName.Enabled = False
    txtAntibodyName.BackColor = &HE0E0E0
    
    cbxCloneType.Enabled = False
    cbxCloneType.BackColor = &HE0E0E0
    
    cbxLieracType.Enabled = False
    cbxLieracType.BackColor = &HE0E0E0
    
    cbxActionObject.Enabled = False
    cbxActionObject.BackColor = &HE0E0E0
    
End Sub

Private Sub LoadValidDate()
'������Ч��
    cbxValidCount.Clear
    
    Call cbxValidCount.AddItem("3")
    Call cbxValidCount.AddItem("6")
    Call cbxValidCount.AddItem("9")
    Call cbxValidCount.AddItem("12")
    Call cbxValidCount.AddItem("18")
    Call cbxValidCount.AddItem("24")
    Call cbxValidCount.AddItem("36")
End Sub


Private Sub LoadCloneType()
'�����¡����
    cbxCloneType.Clear
    
    Call cbxCloneType.AddItem("0-����¡��Ũ���ͣ�")
    Call cbxCloneType.AddItem("1-����¡�������ͣ�")
    Call cbxCloneType.AddItem("2-���¡��Ũ���ͣ�")
    Call cbxCloneType.AddItem("3-���¡�������ͣ�")
    
    cbxCloneType.ListIndex = 0
End Sub


Private Sub LoadLieracType()
'����������
    cbxLieracType.Clear
    
    Call cbxLieracType.AddItem("IgM")
    Call cbxLieracType.AddItem("IgG")
    Call cbxLieracType.AddItem("IgA")
    Call cbxLieracType.AddItem("IgE")
    Call cbxLieracType.AddItem("IgD")
End Sub


Private Sub LoadActionObject()
'�������ö���
    cbxActionObject.Clear
    
    Call cbxActionObject.AddItem("������")
    Call cbxActionObject.AddItem("��������")
    Call cbxActionObject.AddItem("����������")
    Call cbxActionObject.AddItem("��ϸ������")
End Sub


Private Function CheckAntibodyDataIsValid() As String
    CheckAntibodyDataIsValid = ""
    
    '���걾�����Ƿ�Ϊ��
    If Trim(txtAntibodyName.Text) = "" Then
        CheckAntibodyDataIsValid = "�������Ʋ���Ϊ�ա�"
        
        Call txtAntibodyName.SetFocus
        Exit Function
    End If
    
    '���걾�����Ƿ���ȷ¼��
    If Trim(txtUseCount.Text) = "" Or Val(txtUseCount.Text) = 0 Then
        CheckAntibodyDataIsValid = "ʹ���˷�������Ч����������Ч���֡�"
        
        Call txtUseCount.SetFocus
        Exit Function
    End If
    
    If dtpOverdueDate.value <= dtpMadeDate.value Then
        CheckAntibodyDataIsValid = "�������ڱ�������������ڡ�"
        
        Call dtpOverdueDate.SetFocus
        Exit Function
    End If
    
    
    
    '��鿹�������Ƿ��ظ�

    Dim i As Integer
    For i = 1 To mufgParentGrid.GridRows - 1
        If Not mufgParentGrid.RowState(i) = TDataRowState.Del Then
            If Not mblnIsUpdate Then
                If mufgParentGrid.Text(i, gstrAntibody_��������) = txtAntibodyName.Text Then
                    CheckAntibodyDataIsValid = "���������ظ���"
                
                    Call txtAntibodyName.SetFocus
                    Exit Function
                End If
            Else
                If Not mufgParentGrid.SelectionRow = i Then
                    If mufgParentGrid.Text(i, gstrAntibody_��������) = txtAntibodyName.Text Then
                        CheckAntibodyDataIsValid = "���������ظ���"
                    
                        Call txtAntibodyName.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    
End Function



Private Function NewAntibody(ByRef lngAntibodyId As Long) As String
'�����ݿ������������¼
On Error GoTo ErrHandle

    Dim strSQL As String
    Dim rsReture As ADODB.Recordset
    
    NewAntibody = ""
    
'    strSQL = "zl_������_����('" & txtAntibodyName.Text & "'," & Val(txtUseCount.Text) & "," & Val(txtAlredyCount.Text) & "," & _
'                                To_Date(dtpMadeDate.value) & "," & Val(cbxValidCount.Text) & "," & To_Date(dtpOverdueDate.value) & "," & _
'                                Val(cbxCloneType.Text) & ",'" & cbxActionObject.Text & "','" & cbxLieracType.Text & "','" & _
'                                txtApplySituation.Text & "','" & txtRegisterDoctor.Text & "'," & To_Date(dtpRegisterTime.value) & ",'" & _
'                                txtMemo.Text & "')"
'
'    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    
                                
    strSQL = "select zl_������_����([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13]) as ����ֵ from dual"
                                
    Set rsReture = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                txtAntibodyName.Text, _
                                Val(txtUseCount.Text), _
                                Val(txtAlredyCount.Text), _
                                CDate(dtpMadeDate.value), _
                                Val(cbxValidCount.Text), _
                                CDate(dtpOverdueDate.value), _
                                Val(cbxCloneType.Text), _
                                cbxActionObject.Text, _
                                cbxLieracType.Text, _
                                txtApplySituation.Text, _
                                txtRegisterDoctor.Text, _
                                CDate(dtpRegisterTime.value), _
                                txtMemo.Text)
                                
    If rsReture.RecordCount > 0 Then lngAntibodyId = rsReture!����ֵ
    
Exit Function
ErrHandle:
    NewAntibody = err.Description
End Function


Private Function AddAntibodyToList(lngNewAntibodyId As Long) As String
'��ӿ����¼����ʾ�б�
On Error GoTo ErrHandle
    AddAntibodyToList = ""
    
    Dim lngNewRecordIndex As Long
    
    AddAntibodyToList = ""
    
    With mufgParentGrid
        lngNewRecordIndex = .NewRow
        
        .Text(lngNewRecordIndex, gstrAntibody_����ID) = lngNewAntibodyId
        .Text(lngNewRecordIndex, gstrAntibody_��������) = txtAntibodyName.Text
        .Text(lngNewRecordIndex, gstrAntibody_ʹ���˷�) = Val(txtUseCount.Text)
        .Text(lngNewRecordIndex, gstrAntibody_�����˷�) = Val(txtAlredyCount.Text)
        .Text(lngNewRecordIndex, gstrAntibody_��������) = Format(dtpMadeDate.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrAntibody_��Ч��) = Val(cbxValidCount.Text) & "��"
        .Text(lngNewRecordIndex, gstrAntibody_��������) = Format(dtpOverdueDate.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrAntibody_��¡��) = Trim(zlStr.SubB(cbxCloneType.Text, InStr(1, cbxCloneType.Text, "-") + 1, 50))
        .Text(lngNewRecordIndex, gstrAntibody_���ö���) = cbxActionObject.Text
        .Text(lngNewRecordIndex, gstrAntibody_������) = cbxLieracType.Text
        .Text(lngNewRecordIndex, gstrAntibody_Ӧ�����) = txtApplySituation.Text
        .Text(lngNewRecordIndex, gstrAntibody_ʹ��״̬) = "ʹ����"
        .Text(lngNewRecordIndex, gstrAntibody_�Ǽ���) = txtRegisterDoctor.Text
        .Text(lngNewRecordIndex, gstrAntibody_�Ǽ�ʱ��) = dtpRegisterTime.value
        .Text(lngNewRecordIndex, gstrAntibody_��ע) = txtMemo.Text
    
    End With
     
    
Exit Function
ErrHandle:
    AddAntibodyToList = err.Description
End Function


Private Function UpdateAntibodyInfToList()
'���¿����б��еĿ�����Ϣ
On Error GoTo ErrHandle
    UpdateAntibodyInfToList = ""
    
    With mufgParentGrid
        .Text(.SelectionRow, gstrAntibody_��������) = txtAntibodyName.Text
        .Text(.SelectionRow, gstrAntibody_ʹ���˷�) = Val(txtUseCount.Text)
        .Text(.SelectionRow, gstrAntibody_�����˷�) = Val(txtAlredyCount.Text)
        .Text(.SelectionRow, gstrAntibody_��������) = Format(dtpMadeDate.value, gstrDateFormat)
        .Text(.SelectionRow, gstrAntibody_��Ч��) = Val(cbxValidCount.Text) & "��"
        .Text(.SelectionRow, gstrAntibody_��������) = Format(dtpOverdueDate.value, gstrDateFormat)
        .Text(.SelectionRow, gstrAntibody_��¡��) = Trim(zlStr.SubB(cbxCloneType.Text, InStr(1, cbxCloneType.Text, "-") + 1, 50))
        .Text(.SelectionRow, gstrAntibody_���ö���) = cbxActionObject.Text
        .Text(.SelectionRow, gstrAntibody_������) = cbxLieracType.Text
        .Text(.SelectionRow, gstrAntibody_Ӧ�����) = txtApplySituation.Text
        .Text(.SelectionRow, gstrAntibody_ʹ��״̬) = "ʹ����"
        .Text(.SelectionRow, gstrAntibody_�Ǽ���) = txtRegisterDoctor.Text
        .Text(.SelectionRow, gstrAntibody_�Ǽ�ʱ��) = dtpRegisterTime.value
        .Text(.SelectionRow, gstrAntibody_��ע) = txtMemo.Text
    End With
Exit Function
ErrHandle:
    UpdateAntibodyInfToList = err.Description
End Function



Private Function UpdateAntibody() As String
'�������ݿ��еĿ�������
On Error GoTo ErrHandle

    Dim strSQL As String
    Dim lngCurAntibodyId As Long
    
    UpdateAntibody = ""
    
    lngCurAntibodyId = mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow)
    
    strSQL = "zl_������_����(" & lngCurAntibodyId & ",'" & txtAntibodyName.Text & "'," & Val(txtUseCount.Text) & "," & Val(txtAlredyCount.Text) & "," & _
                                zlStr.To_Date(dtpMadeDate.value) & "," & Val(cbxValidCount.Text) & "," & zlStr.To_Date(dtpOverdueDate.value) & "," & _
                                Val(cbxCloneType.Text) & ",'" & cbxActionObject.Text & "','" & cbxLieracType.Text & "','" & _
                                txtApplySituation.Text & "','" & txtRegisterDoctor.Text & "'," & zlStr.To_Date(dtpRegisterTime.value) & ",'" & _
                                txtMemo.Text & "')"
                                
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
Exit Function
ErrHandle:
    UpdateAntibody = err.Description
End Function





Private Sub cbxValidCount_LostFocus()
On Error Resume Next
    If Val(cbxValidCount.Text) > 0 Then
        dtpOverdueDate.value = DateAdd("m", Val(cbxValidCount.Text), dtpMadeDate.value)
    End If
End Sub

Private Sub cmdNewAntibody_Cancel_Click()
'    mblnIsSucceed = False '��ȷ�Ϻ�������˳���ӽ��棬��ȡ����ʱ�򣬲��ܸ�ֵΪ��
    Call Me.Hide
End Sub

Private Sub cmdNewAntibody_Sure_Click()
On Error GoTo ErrHandle
    Dim strErr As String
    Dim lngNewAntibodyId As Long
    
    mblnIsSucceed = False
    
    
    strErr = CheckAntibodyDataIsValid()
    If Trim(strErr) <> "" Then
        Call ShowProcessHint(strErr)
        Exit Sub
    End If
    
    If Not mblnIsUpdate Then
        '���������¼
        strErr = NewAntibody(lngNewAntibodyId)
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        strErr = AddAntibodyToList(lngNewAntibodyId)
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        Call mufgParentGrid.LocateRow(mufgParentGrid.GridRows - 1)
    Else
        '���¿����¼
        strErr = UpdateAntibody()
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        strErr = UpdateAntibodyInfToList()
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
    End If
    
    mblnIsSucceed = True
    
    If Not CBool(chkContinue.value) Then
        Call Unload(Me)
    End If
    
    Call CloseProcessHint
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpOverdueDate_LostFocus()
On Error Resume Next
    cbxValidCount.Text = DateDiff("m", dtpMadeDate.value, dtpOverdueDate.value)
End Sub

Private Sub Form_Initialize()
    mblnIsSucceed = False
    mblnIsUpdate = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadValidDate
    Call LoadCloneType
    Call LoadLieracType
    Call LoadActionObject
    
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowProcessHint(ByVal strHint As String)
'��ʾ������Ϣ
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Private Sub CloseProcessHint()
'�رմ�����ʾ
    picShow.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub txtUseCount_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAlredyCount_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
