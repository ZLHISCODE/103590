VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocPrintPatiList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ����"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   ControlBox      =   0   'False
   Icon            =   "frmDocPrintPatiList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkPrinted 
      Caption         =   "δ��ӡ"
      Height          =   255
      Left            =   7380
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox chkChoose 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7380
      TabIndex        =   2
      Top             =   502
      Width           =   1125
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7380
      TabIndex        =   1
      Top             =   45
      Width           =   1125
   End
   Begin MSComctlLib.ListView lvwlist 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   5900
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
      Icons           =   "iLsTree32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��Դ"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ա�"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "��λ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ҽ��ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "��ӡ״̬"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "PACS����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ִ�п���ID"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmDocPrintPatiList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReturn As String

Public Function Showfrm(ByVal vsList As VSFlexGrid, frmParent As Form, ByVal blnCanPrint As Boolean, _
    ByVal blnPacsReport As Boolean, ByVal lngDeptID As Long) As String
'������vsList�����б�blnCanPrint ƽ�ﱨ����Ҫ��˲��ܴ�ӡ
    Dim i As Integer, lvwItem As ListItem
    Dim iCount As Integer
    Dim lngOldDeptID As Long

    chkPrinted.value = 0
    mstrReturn = ""
    iCount = 0
    lngOldDeptID = 0
    
    lvwlist.ListItems.Clear
    For i = 1 To vsList.Rows - 1
        With vsList
            If .TextMatrix(i, GetColNum(vsList, "������")) = "�ѱ���" _
                Or .TextMatrix(i, GetColNum(vsList, "������")) = "�����" _
                Or .TextMatrix(i, GetColNum(vsList, "������")) = "�����" Then
            
                '����С�ִ�п���ID��������Ҫ���¸��ݿ���ID��ȡƽ�ﱨ������˵Ĳ���
                If GetColNum(vsList, "ִ�п���ID") <> 0 Then
                    If lngOldDeptID <> .TextMatrix(i, GetColNum(vsList, "ִ�п���ID")) Then   '����ID�ı��ˣ����¶�ȡƽ�ﱨ���ӡ�Ĳ���
                        lngOldDeptID = .TextMatrix(i, GetColNum(vsList, "ִ�п���ID"))
                        blnCanPrint = GetDeptPara(lngOldDeptID, "ƽ������˲��ܴ򱨸�") = "1"           'ƽ����Ҫ��˲��ܴ�ӡ =true
                        blnPacsReport = GetDeptPara(lngOldDeptID, "����༭��", 0) = "1" '              '����༭��
                    End If
                Else
                    lngOldDeptID = lngDeptID
                End If
                If IIf(blnCanPrint, IIf(.Cell(flexcpData, i, GetColNum(vsList, "����")) = 1, .TextMatrix(i, GetColNum(vsList, "������")) <> "", .TextMatrix(i, GetColNum(vsList, "������")) <> ""), True) Then
                    iCount = iCount + 1
                    Set lvwItem = lvwlist.ListItems.Add(, "_" & .TextMatrix(i, GetColNum(vsList, "ҽ��ID")), .TextMatrix(i, GetColNum(vsList, "����")))
                        lvwItem.SubItems(1) = .TextMatrix(i, GetColNum(vsList, "��Դ"))
                        lvwItem.SubItems(2) = .TextMatrix(i, GetColNum(vsList, "����"))
                        lvwItem.SubItems(3) = .TextMatrix(i, GetColNum(vsList, "�Ա�"))
                        lvwItem.SubItems(4) = .TextMatrix(i, GetColNum(vsList, "����"))
                        lvwItem.SubItems(5) = .TextMatrix(i, GetColNum(vsList, "ҽ������"))
                        lvwItem.SubItems(6) = .TextMatrix(i, GetColNum(vsList, "��λ����"))
                        lvwItem.SubItems(7) = .TextMatrix(i, GetColNum(vsList, "ҽ��ID"))
                        lvwItem.SubItems(8) = Nvl(.TextMatrix(i, GetColNum(vsList, "�����ӡ")), "")
                        lvwItem.SubItems(9) = IIf(blnPacsReport, 1, 0)
                        lvwItem.SubItems(10) = lngOldDeptID
                End If
            End If
        End With
    Next
    Me.Caption = "ѡ����Ҫ��ӡ��ҽ����ҽ������Ϊ��" & iCount
    Me.Show 1, frmParent
    Showfrm = mstrReturn
End Function

Private Sub chkChoose_Click()
Dim l As Integer
    If chkChoose.value = 1 Then
        chkChoose.Caption = "ȫ��(&D)"
        For l = 1 To lvwlist.ListItems.Count
            lvwlist.ListItems(l).Checked = True
        Next
    Else
        chkChoose.Caption = "ȫѡ(&A)"
        For l = 1 To lvwlist.ListItems.Count
            lvwlist.ListItems(l).Checked = False
        Next
    End If
End Sub

Private Sub chkPrinted_Click()
    Dim i As Integer
    
    For i = 1 To lvwlist.ListItems.Count
        If lvwlist.ListItems(i).SubItems(8) = "" Then lvwlist.ListItems(i).Checked = IIf(chkPrinted.value = 1, True, False)
    Next i
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '��֯����ֵ������ֵ��"ҽ��ID-�Ƿ�PACS����༭��-ִ�п���ID|ҽ��ID-�Ƿ�PACS����༭��-ִ�п���ID|..."���
    Dim l As Long
    
    For l = 1 To lvwlist.ListItems.Count
        If lvwlist.ListItems(l).Checked Then
            mstrReturn = mstrReturn & "|" & lvwlist.ListItems(l).SubItems(7) _
                         & "-" & lvwlist.ListItems(l).SubItems(9) & "-" & lvwlist.ListItems(l).SubItems(10)
        End If
    Next
    mstrReturn = Mid(mstrReturn, 2)
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        cmdOK_Click
    End If
End Sub

Private Sub lvwList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    zlControl.LvwSortColumn lvwlist, ColumnHeader.Index
End Sub
