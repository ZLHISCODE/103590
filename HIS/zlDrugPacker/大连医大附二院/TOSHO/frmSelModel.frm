VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ѡ��"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmSelModel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4830
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwModel 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4600
      _ExtentX        =   8123
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblAll 
      AutoSize        =   -1  'True
      Caption         =   "ȫѡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4680
      Width           =   390
   End
   Begin VB.Label lblUnAll 
      AutoSize        =   -1  'True
      Caption         =   "ȫ��ѡ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   840
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4680
      Width           =   585
   End
End
Attribute VB_Name = "frmSelModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strDrugModel As String
Public mcnHIS As New ADODB.Connection

Public Property Get DrugModel() As String
    DrugModel = strDrugModel
End Property

'Public Property Get ResultModal() As Boolean
'    ResultModal = blnResultModal
'End Property

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim i As Byte
    Dim strTmp As String, strInsert As String
    Dim rsTmp As New ADODB.Recordset
    Dim cmdInsert As New ADODB.Command
    Dim blnBegin As Boolean
    
    For i = 1 To lvwModel.ListItems.Count
        If lvwModel.ListItems(i).Checked = True Then
            strTmp = strTmp & "'" & Trim(lvwModel.ListItems(i).Text) & "',"
        End If
    Next
    If strTmp = "" Then Exit Sub
    strDrugModel = Left(strTmp, Len(strTmp) - 1)
    
    '����ҩƷ������Ϣ

    Screen.MousePointer = vbHourglass
    DoEvents
    '���ɽű�
    strTmp = "Select A.*, B.���� ���̱��� " _
           & "From (Select A.ҩƷid, B.����, B.����, B.���, D.ҩƷ���� ҩƷ����, A.����ϵ�� ����, E.���㵥λ ������λ," _
           & "        A.סԺ��װ ��װ����, A.סԺ��λ ��װ��λ, Nvl(A.�ϴβ���, B.����) ��������," _
           & "        zlTools.zlSpellCode(B.����) ƴ����" _
           & "      From ҩƷ��� A, �շ���ĿĿ¼ B, ҩƷ���� D, ������ĿĿ¼ E" _
           & "      Where A.ҩƷid = B.ID And A.ҩ��id = D.ҩ��id And A.ҩ��id = E.ID And B.��� In ('5', '6', '7') And" _
           & "        Nvl(B.����ʱ��, To_Date('3000-1-1', 'yyyy-mm-dd')) = To_Date('3000-1-1', 'yyyy-mm-dd') And" _
           & "        D.ҩƷ���� In (" & frmSelModel.DrugModel & ")) A, ҩƷ������ B " _
           & "Where A.�������� = B.����(+) order by cast(a.���� as int) "
    If Me.mcnHIS.State = adStateClosed Then Me.mcnHIS.Open
    rsTmp.Open strTmp, Me.mcnHIS
    If rsTmp.EOF Then
        Screen.MousePointer = vbDefault
        rsTmp.Close
        Exit Sub
    End If
    
    On Error GoTo errHand
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        gcnOutside.BeginTrans
        blnBegin = True
        strInsert = "delete dbo.atf_his_druginfo where drug_code='" & rsTmp!���� & "' and drugname='" & rsTmp!���� & "' " & Chr(13) _
                  & "insert into dbo.atf_his_druginfo (drug_code,drugname,specification,drug_type," _
                  & "dosage,dos_unit,pack_amount,pack_name,manufactory,py_code,manu_no) " & Chr(13)
        strTmp = "select '" & rsTmp!���� & "'," _
               & "'" & rsTmp!���� & "'," _
               & "'" & rsTmp!��� & "'," _
               & "'" & rsTmp!ҩƷ���� & "'," _
               & CDbl(rsTmp!����) & "," _
               & "'" & rsTmp!������λ & "'," _
               & CDbl(rsTmp!��װ����) & "," _
               & "'" & rsTmp!��װ��λ & "'," _
               & "'" & IIf(IsNull(rsTmp!��������), "", rsTmp!��������) & "'," _
               & "'" & rsTmp!ƴ���� & "'," _
               & "'" & IIf(IsNull(rsTmp!���̱���), "", rsTmp!���̱���) & "' "
        strInsert = strInsert & strTmp & Chr(13)
        '��������
        With cmdInsert
            .ActiveConnection = gcnOutside
            .CommandText = strInsert
            .Execute
        End With
        If blnBegin Then
            gcnOutside.CommitTrans
        End If
        blnBegin = False
        rsTmp.MoveNext
    Loop
    rsTmp.Close

    Screen.MousePointer = vbDefault
    MsgBox "���ͳɹ���", vbInformation, GSTR_SYSNAME
    Exit Sub

errHand:
    gcnOutside.RollbackTrans
    blnBegin = False
    Call OutPutLog("�ϴ�ҩƷ�����쳣��" & Err.Description)
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    lvwModel.View = lvwReport
'    Call InitLvwModel
End Sub

Public Sub InitLvwModel() '(ByVal cnZLHIS As ADODB.Connection)
    Dim rsTmp As New ADODB.Recordset
    Dim itmTmp As ListItem
    'Set rsTmp = zlDatabase.OpenSQLRecord("select distinct ҩƷ���� from ҩƷ����", Me.Caption)
    If Me.mcnHIS.State = adStateClosed Then Me.mcnHIS.Open
    rsTmp.Open "select distinct ҩƷ���� from ҩƷ����", Me.mcnHIS
    If Not rsTmp.EOF Then
        With lvwModel
            .ListItems.Clear
            .ColumnHeaders.Add 1, "K_Choose", "ѡ�����", lvwModel.Width - 400
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                Set itmTmp = .ListItems.Add(, , rsTmp!ҩƷ����)
                rsTmp.MoveNext
            Loop
        End With
    End If
    rsTmp.Close
End Sub

Private Sub lblAll_Click()
    Dim i As Integer
    For i = 1 To lvwModel.ListItems.Count
        lvwModel.ListItems(i).Checked = True
    Next
End Sub

Private Sub lblUnAll_Click()
    Dim i As Integer
    For i = 1 To lvwModel.ListItems.Count
        lvwModel.ListItems(i).Checked = False
    Next
End Sub

