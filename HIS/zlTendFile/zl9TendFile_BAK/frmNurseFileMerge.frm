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
    On Error GoTo errHand
    
    gstrSQL = "ZL_���˻����ļ�_STATE(" & mlngFile & ",2,NULL," & Me.cbo�����ļ�.ItemData(Me.cbo�����ļ�.ListIndex) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "ȡ���ϲ���ӡ")
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim str��ʼʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '��ȡ�ò�����ָ���ļ���ʽ��ͬ���ļ�,�趨�ϲ���ӡ(ֻ�ܰ�ʱ����Ⱥ�˳������趨)
    
    gstrSQL = " Select �ļ�����,��ʼʱ�� From ���˻����ļ� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ�����", mlngFile)
    str��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss")
    Me.lbl��ǰ�ļ�.Caption = "��ǰ�ļ���" & rsTemp!�ļ�����
    
    gstrSQL = " Select ID,�ļ����� " & _
              " From ���˻����ļ� " & _
              " Where (����ID,��ҳID,Ӥ��,��ʽID) IN " & _
              "     (Select B.����ID,B.��ҳID,B.Ӥ��,A.ID " & _
              "     From �����ļ��б� A,���˻����ļ� B " & _
              "     Where A.ID=B.��ʽID And B.ID=[1])" & _
              " And ID<>[1] And ��ʼʱ��>=to_date('" & str��ʼʱ�� & "','yyyy-MM-dd hh24:mi:ss')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò�����ָ���ļ���ʽ��ͬ���ļ�,�趨�ϲ���ӡ", mlngFile)
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
