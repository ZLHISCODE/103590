VERSION 5.00
Begin VB.Form frmNurseFileMerge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ϲ���ӡ"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4155
   Icon            =   "frmNurseFileMerge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   4
      Top             =   1440
      Width           =   1155
   End
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -120
      TabIndex        =   3
      Top             =   1110
      Width           =   4365
   End
   Begin VB.ComboBox cbo�����ļ� 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   2625
   End
   Begin VB.Label lbl�����ļ� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ļ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl��ǰ�ļ� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�ļ�:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   810
   End
End
Attribute VB_Name = "frmNurseFileMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngFile As Long
Private mblnOK As Boolean

Public Function ShowEditor(ByVal lngFile As Long) As Boolean
    On Error Resume Next
    mlngFile = lngFile
    mblnOK = False
    Me.Show 1
    ShowEditor = mblnOK
End Function

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnTrans As Boolean
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    blnTrans = True
    gstrSQL = "ZL_���˻����ļ�_STATE(" & mlngFile & ",2,NULL," & Me.cbo�����ļ�.ItemData(Me.cbo�����ļ�.ListIndex) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�ϲ���ӡ")
    gstrSQL = "Zl_���˻����ӡ_Batchretrypage(" & Me.cbo�����ļ�.ItemData(Me.cbo�����ļ�.ListIndex) & ",'1;0')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ҳ������")
    gcnOracle.CommitTrans
    blnTrans = False
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim str��ʼʱ�� As String, str����ʱ�� As String
    Dim lngPati As Long, lngPage As Long, lngBaby As Long, lngFormat As Long
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '��ȡ�ò�����ָ���ļ���ʽ��ͬ���ļ�,�趨�ϲ���ӡ(ֻ�ܰ�ʱ����Ⱥ�˳������趨)
    
    gstrSQL = " Select ����ID,��ҳID,nvl(Ӥ��,0) Ӥ��,��ʽID,�ļ�����,��ʼʱ��,����ʱ�� From ���˻����ļ� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ�����", mlngFile)
    lngPati = NVL(rsTemp!����ID, 0)
    lngPage = NVL(rsTemp!��ҳID, 0)
    lngBaby = NVL(rsTemp!Ӥ��, 0)
    lngFormat = NVL(rsTemp!��ʽID, 0)
    str��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
    str����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    Me.lbl��ǰ�ļ�.Caption = "��ǰ�ļ���" & rsTemp!�ļ�����
    
    gstrSQL = _
            "  Select ID,�ļ�����" & vbNewLine & _
            "  From (With ���˻����ļ�_F1 As" & vbNewLine & _
            "   (Select Id, ����id, �ļ�����, ��ʼʱ��,����ʱ��" & vbNewLine & _
            "   From ���˻����ļ�" & vbNewLine & _
            "   Where ����id = [2] And ��ҳid = [3] And Nvl(Ӥ��, 0) = [4] And ��ʽid = [5])" & vbNewLine & _
            "   Select Id, �ļ�����" & vbNewLine & _
            "   From ���˻����ļ�_F1 a" & vbNewLine & _
            "   Where a.Id <> [1] And (a.��ʼʱ��>[6] OR (a.��ʼʱ��=[6] And a.����ʱ��>[7]))" & vbNewLine & _
            "   And Not Exists (Select Id From ���˻����ļ�_F1 Where a.Id = ����id)" & vbNewLine & _
            "   Order By a.��ʼʱ��)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò�����ָ���ļ���ʽ��ͬ���ļ�,�趨�ϲ���ӡ", mlngFile, lngPati, lngPage, lngBaby, lngFormat, CDate(str��ʼʱ��), CDate(str����ʱ��))
    With rsTemp
        Me.cbo�����ļ�.Clear
        Do While Not .EOF
            cbo�����ļ�.AddItem !�ļ�����
            cbo�����ļ�.ItemData(cbo�����ļ�.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    If cbo�����ļ�.ListCount = 0 Then
        MsgBox "��ǰ�ļ�֮�󲻴���ͬ��ʽ���ļ�����˲���ҪΪ��ǰ�ļ�ָ���ϲ���ӡ��", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    cbo�����ļ�.ListIndex = 0
    
End Sub

