VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugSendQueryCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   11565
   Icon            =   "frmDrugSendQueryCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   6435
      ScaleHeight     =   450
      ScaleWidth      =   2565
      TabIndex        =   36
      Top             =   4170
      Width           =   2565
      Begin VB.ComboBox cboReqDruDep 
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   60
         Width           =   2460
      End
   End
   Begin VB.PictureBox picCond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4000
      Left            =   3435
      ScaleHeight     =   4005
      ScaleWidth      =   2460
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   2460
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   630
         TabIndex        =   8
         Top             =   1770
         Width           =   1800
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H80000005&
         Caption         =   "��ҩ����ҩΪ׼"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   560
         Width           =   1560
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H80000005&
         Caption         =   "��ҽ������Ϊ׼"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   350
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.ComboBox cboҩ�� 
         Height          =   300
         Left            =   465
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   15
         Width           =   1845
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "��ѯ(&O)"
         Height          =   350
         Left            =   1200
         TabIndex        =   16
         Top             =   3600
         Width           =   1100
      End
      Begin VB.CheckBox chk��ҩ 
         Alignment       =   1  'Right Justify
         Caption         =   "��ҩ����ʱ��"
         Height          =   180
         Left            =   45
         TabIndex        =   13
         Top             =   2620
         Width           =   1380
      End
      Begin VB.CheckBox chk��Ч 
         Caption         =   "����"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   9
         Top             =   2115
         Width           =   675
      End
      Begin VB.CheckBox chk��Ч 
         Caption         =   "����"
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   10
         Top             =   2340
         Width           =   660
      End
      Begin VB.CheckBox chk��ҩ 
         Caption         =   "δ��ҩ"
         Height          =   195
         Index           =   0
         Left            =   1350
         TabIndex        =   11
         Top             =   2115
         Width           =   885
      End
      Begin VB.CheckBox chk��ҩ 
         Caption         =   "�ѷ�ҩ"
         Height          =   195
         Index           =   1
         Left            =   1350
         TabIndex        =   12
         Top             =   2340
         Width           =   885
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   630
         TabIndex        =   7
         Top             =   1440
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   465
         TabIndex        =   6
         Top             =   1095
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   465
         TabIndex        =   5
         Top             =   780
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   3
         Left            =   465
         TabIndex        =   15
         Top             =   3240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   2
         Left            =   465
         TabIndex        =   14
         Top             =   2900
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��"
         Height          =   180
         Index           =   1
         Left            =   45
         TabIndex        =   35
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label lblҩ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҩ��"
         Height          =   180
         Left            =   45
         TabIndex        =   33
         Top             =   75
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   5
         Left            =   225
         TabIndex        =   32
         Top             =   3315
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   31
         Top             =   2975
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   30
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   29
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʱ��"
         Height          =   180
         Left            =   45
         TabIndex        =   28
         Top             =   350
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   27
         Top             =   1500
         Width           =   540
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   6480
      ScaleHeight     =   3525
      ScaleWidth      =   2475
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   2475
      Begin VB.CheckBox chkPreOut 
         Caption         =   "Ԥ��Ժ(&P)"
         Height          =   195
         Left            =   0
         TabIndex        =   34
         Top             =   2970
         Width           =   1200
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   0
         Width           =   2490
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "ȫѡ"
         Height          =   330
         Left            =   1620
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   3210
         Width           =   870
      End
      Begin VB.CheckBox chkOut 
         Caption         =   "�����Ժ(&A)"
         Height          =   195
         Left            =   1140
         TabIndex        =   20
         Top             =   2970
         Width           =   1320
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   2610
         Left            =   0
         TabIndex        =   19
         Top             =   330
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   4604
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "סԺ��"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "�ѱ�"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "��Ժ����"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "��Ժ����"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "��������"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "ȫ��"
         Height          =   330
         Left            =   765
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3210
         Width           =   870
      End
   End
   Begin VB.PictureBox picWay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   8925
      ScaleHeight     =   2790
      ScaleWidth      =   2430
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   390
      Width           =   2430
      Begin VB.CommandButton cmdAllWay 
         Caption         =   "ȫѡ"
         Height          =   330
         Left            =   1575
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   2475
         Width           =   870
      End
      Begin MSComctlLib.ListView lvwWay 
         Height          =   2445
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   4313
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��ҩ;��"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton cmdNoWay 
         Caption         =   "ȫ��"
         Height          =   330
         Left            =   720
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   2475
         Width           =   870
      End
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   6600
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3285
      _Version        =   589884
      _ExtentX        =   5794
      _ExtentY        =   11642
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmDrugSendQueryCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event DoQuery(ByVal ҩ��ID As Long, ByVal Mode As Byte, ByVal DateBegin As Date, ByVal DateEnd As Date, ByVal ��ҩDateB As Date, ByVal ��ҩDateE As Date, _
    ByVal NO As String, ByVal ��ҩ�� As String, ByVal ��Ч As Integer, ByVal ״̬ As String, ByVal ����ID As Long, ByVal ����IDs As String, ByVal ��ҩ;�� As String, ByVal ��ҩ����ID As Long)

Private mMainPrivs As String 'IN
Private mlng����ID As Long 'IN
Private mlng����ID As Long 'IN
Private mblnOnePati As Boolean 'IN��������ģʽ

Private Type QUERY_COND
    DateBegin As Date
    DateEnd As Date
    ��ҩDateB As Date
    ��ҩDateE As Date
    ��ҩ;�� As String
    NO As String
    ��ҩ�� As String
    ҩ��ID As Long
    ����IDs As String
    ����ID As Long
    ��ҩ����ID As Long
    ��Ч As Integer '2-ȫ��
    ״̬ As String
End Type
Private mvQuery As QUERY_COND

Private Enum tkpItemIndex
    Item_��ѯ���� = 1
    Item_�����벡�� = 2
    Item_��ҩ;�� = 3
    Item_��ҩ���� = 4
End Enum


Public Sub InitParameter(ByVal strMainPrivs As String, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal blnOnePati As Boolean)
    mMainPrivs = strMainPrivs
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mblnOnePati = blnOnePati
End Sub

Private Sub cboUnit_Click()
'���ܣ���ȡָ����Χ�ڵĲ����б�
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str����IDs As String, lng����ID As Long
    Dim lngUnitID, lngColor As Long
        
    On Error GoTo errH
    lvwPati.ListItems.Clear
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    
    
    str����IDs = zlDatabase.GetPara("���Ͳ���", glngSys, pסԺҽ������)
    If str����IDs <> "" And InStr(str����IDs, ":") > 0 Then
        lng����ID = Val(Split(str����IDs, ":")(0))
        str����IDs = Split(str����IDs, ":")(1)
    End If
        
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng����ID, False, False, chkOut.value, chkPreOut.value)
  
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!����ID & "_" & rsTmp!��ҳID, rsTmp!����)
        objItem.SubItems(1) = Nvl(rsTmp!סԺ��)
        objItem.SubItems(2) = Nvl(rsTmp!����)
        objItem.SubItems(3) = Nvl(rsTmp!�ѱ�)
        objItem.SubItems(4) = Nvl(rsTmp!����)
        objItem.SubItems(5) = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
        objItem.SubItems(6) = Format(Nvl(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
        objItem.SubItems(7) = Nvl(rsTmp!��������)
        
        '������ɫ
        lngColor = zlDatabase.GetPatiColor(Nvl(rsTmp!��������))
        objItem.ListSubItems(1).ForeColor = lngColor
        objItem.ListSubItems(7).ForeColor = lngColor
        
        '�ϴ��Ƿ�ѡ��
        If lngUnitID = lng����ID And str����IDs <> "" Then
            If str����IDs = "ALL" _
                Or Left(str����IDs, 1) <> "-" And InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") > 0 _
                Or Left(str����IDs, 1) = "-" And InStr("," & Mid(str����IDs, 2) & ",", "," & rsTmp!����ID & ",") = 0 Then
                objItem.Checked = True
                If k = 0 Then 'Ϊ�˿�����ѡ���
                    objItem.EnsureVisible
                    objItem.Selected = True
                    k = 1
                End If
            End If
        ElseIf rsTmp!����ID = mlng����ID Then
            objItem.Checked = True 'ȱʡֻѡ��ǰ����
            objItem.EnsureVisible
            objItem.Selected = True
        End If
       
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkOut_Click()
    If Visible Then Call cboUnit_Click
End Sub

Private Sub chkPreOut_Click()
    If Visible Then Call cboUnit_Click
End Sub

Private Sub chk��ҩ_Click(index As Integer)
    If chk��ҩ(0).value = 0 And chk��ҩ(1).value = 0 Then
        chk��ҩ(index).value = 1: Exit Sub
    End If
    
    chk��ҩ.Enabled = chk��ҩ(1).value = 0
    If Not chk��ҩ.Enabled Then chk��ҩ.value = 0
End Sub

Private Sub chk��Ч_Click(index As Integer)
    If chk��Ч(0).value = 0 And chk��Ч(1).value = 0 Then
        chk��Ч(index).value = 1: Exit Sub
    End If
End Sub

Private Sub chk��ҩ_Click()
    dtpDate(2).Enabled = chk��ҩ.value = 1 And dtpDate(2).Tag = ""
    dtpDate(3).Enabled = chk��ҩ.value = 1 And dtpDate(3).Tag = ""
    
    If dtpDate(2).Enabled Then dtpDate(2).SetFocus
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

Private Sub cmdAllWay_Click()
    Call SelectLVW(lvwWay, True)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoWay_Click()
    Call SelectLVW(lvwWay, False)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdQuery_Click()
    Dim str����IDs As String, strUn����IDs As String
    Dim str��ҩIDs As String, i As Long
        
    If cboҩ��.ListIndex = -1 Then
        MsgBox "��ѡ��һ��ҩ����", vbInformation, gstrSysName
        tkpMain.Groups(Item_��ѯ����).Expanded = True: cboҩ��.SetFocus: Exit Sub
    End If
    If dtpDate(0).value >= dtpDate(1).value Then
        MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
        tkpMain.Groups(Item_��ѯ����).Expanded = True: dtpDate(0).SetFocus: Exit Sub
    End If
    If chk��ҩ.value = 1 Then
        If dtpDate(2).value >= dtpDate(3).value Then
            MsgBox "��ʼʱ��ӦС�ڽ���ʱ�䡣", vbInformation, gstrSysName
            tkpMain.Groups(Item_��ѯ����).Expanded = True: dtpDate(2).SetFocus: Exit Sub
        End If
    End If
    If cboUnit.ListIndex = -1 Then
        MsgBox "��ѡ��һ��������", vbInformation, gstrSysName
        tkpMain.Groups(Item_�����벡��).Expanded = True: cboUnit.SetFocus: Exit Sub
    End If
    
    '����
    str����IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            str����IDs = str����IDs & "," & Split(lvwPati.ListItems(i).Key, "_")(1)
        Else
            strUn����IDs = strUn����IDs & "," & Split(lvwPati.ListItems(i).Key, "_")(1)
        End If
    Next
    str����IDs = Mid(str����IDs, 2)
    strUn����IDs = Mid(strUn����IDs, 2)
    If str����IDs = "" Or (UBound(Split(str����IDs, ",")) = 0 And Val(str����IDs) = mlng����ID) Then
        str����IDs = ""
    Else
        If strUn����IDs = "" Then
            str����IDs = cboUnit.ItemData(cboUnit.ListIndex) & ":ALL"
        ElseIf UBound(Split(str����IDs, ",")) > UBound(Split(strUn����IDs, ",")) Then
            str����IDs = cboUnit.ItemData(cboUnit.ListIndex) & ":-" & strUn����IDs
        Else
            str����IDs = cboUnit.ItemData(cboUnit.ListIndex) & ":" & str����IDs
        End If
    End If
    
    '��ҩ;��
    mvQuery.��ҩ;�� = "": str��ҩIDs = ""
    For i = 1 To lvwWay.ListItems.Count
        If lvwWay.ListItems(i).Checked Then
            str��ҩIDs = str��ҩIDs & "," & Mid(lvwWay.ListItems(i).Key, 2)
            mvQuery.��ҩ;�� = mvQuery.��ҩ;�� & "," & lvwWay.ListItems(i).Text
        End If
    Next
    str��ҩIDs = Mid(str��ҩIDs, 2)
    mvQuery.��ҩ;�� = Mid(mvQuery.��ҩ;��, 2)
    If str��ҩIDs = "" Then
        MsgBox "������ѡ��һ�ָ�ҩ;����", vbInformation, gstrSysName
        tkpMain.Groups(Item_��ҩ;��).Expanded = True: lvwWay.SetFocus: Exit Sub
    End If
    If UBound(Split(str��ҩIDs, ",")) + 1 = lvwWay.ListItems.Count Then
        str��ҩIDs = "": mvQuery.��ҩ;�� = ""
    End If
        
    '���������ע�����
    '---------------------------------------------------------------
    'ҩ��
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯҩ��", cboҩ��.ItemData(cboҩ��.ListIndex), glngSys, pסԺҽ������)
    
    '��ҩ����
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯ��ҩ����", cboReqDruDep.ItemData(cboReqDruDep.ListIndex), glngSys, pסԺҽ������)
    
    'ʱ��
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯ���", DateDiff("d", dtpDate(0).value, dtpDate(1).value), glngSys, pסԺҽ������)
    If chk��ҩ.value = 1 Then
        Call zlDatabase.SetPara("��ҩ��ѯ���", DateDiff("d", dtpDate(2).value, dtpDate(3).value), glngSys, pסԺҽ������)
    End If
    
    '��Ч
    If chk��Ч(0).value = 1 And chk��Ч(1).value = 1 Then
        i = 2
    ElseIf chk��Ч(0).value = 1 Then
        i = 0
    Else
        i = 1
    End If
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯ��Ч", i, glngSys, pסԺҽ������)
    
    '״̬
    If chk��ҩ(0).value = 1 And chk��ҩ(1).value = 1 Then
        i = 2
    ElseIf chk��ҩ(0).value = 1 Then
        i = 0
    Else
        i = 1
    End If
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯ״̬", i, glngSys, pסԺҽ������)
        
    '����
    Call zlDatabase.SetPara("���Ͳ���", str����IDs, glngSys, pסԺҽ������)
    
    '������Ժ����
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯ��Ժ����", chkOut.value, glngSys, pסԺҽ������)
    '����Ԥ��Ժ����
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯԤ��Ժ����", chkPreOut.value, glngSys, pסԺҽ������)
    
    '��ҩ;��
    Call zlDatabase.SetPara("ҩ�Ʋ�ѯ��ҩ;��", str��ҩIDs, glngSys, pסԺҽ������)
    
    '�ռ�����
    '---------------------------------------------------------------------
    '����
    mvQuery.����ID = cboUnit.ItemData(cboUnit.ListIndex)

    '����
    mvQuery.����IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mvQuery.����IDs = mvQuery.����IDs & "," & Split(lvwPati.ListItems(i).Key, "_")(1)
        End If
    Next
    mvQuery.����IDs = Mid(mvQuery.����IDs, 2)

    'ʱ��
    mvQuery.DateBegin = Format(dtpDate(0).value, "yyyy-MM-dd HH:mm:00")
    mvQuery.DateEnd = Format(dtpDate(1).value, "yyyy-MM-dd HH:mm:59")
    If chk��ҩ.value = 1 Then
        mvQuery.��ҩDateB = Format(dtpDate(2).value, "yyyy-MM-dd HH:mm:00")
        mvQuery.��ҩDateE = Format(dtpDate(3).value, "yyyy-MM-dd HH:mm:59")
    Else
        mvQuery.��ҩDateB = Empty
        mvQuery.��ҩDateE = Empty
    End If
    
    'NO
    mvQuery.NO = txtNO(0).Text
    '��ҩ��
    mvQuery.��ҩ�� = Trim(txtNO(1).Text)
    
    '��Ч
    If chk��Ч(0).value = 1 And chk��Ч(1).value = 1 Then
        mvQuery.��Ч = 2
    ElseIf chk��Ч(0).value = 1 Then
        mvQuery.��Ч = 0
    ElseIf chk��Ч(1).value = 1 Then
        mvQuery.��Ч = 1
    End If

    '״̬
    mvQuery.״̬ = chk��ҩ(0).value & chk��ҩ(1).value

    'ҩ��
    mvQuery.ҩ��ID = cboҩ��.ItemData(cboҩ��.ListIndex)
    
    '��ҩ����
    mvQuery.��ҩ����ID = cboReqDruDep.ItemData(cboReqDruDep.ListIndex)
    '�����¼�
    '------------------------------------------------------------------------
    With mvQuery
        RaiseEvent DoQuery(.ҩ��ID, IIF(optDate(0).value, 0, 1), .DateBegin, .DateEnd, .��ҩDateB, .��ҩDateE, .NO, .��ҩ��, .��Ч, .״̬, .����ID, .����IDs, .��ҩ;��, .��ҩ����ID)
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdAllPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdAllWay_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdNoPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdNoWay_Click
        End If
    ElseIf KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim curDate As Date, i As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    Me.Width = tkpMain.Width: Me.Height = tkpMain.Height
    
    '����ؼ�------------------------------------------
    Call tkpMain.SetMargins(8, 8, 8, 8, 8)

    Set objGroup = tkpMain.Groups.Add(Item_��ѯ����, "��ѯ����")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picCond
    picCond.BackColor = objItem.BackColor
    optDate(0).BackColor = objItem.BackColor
    optDate(1).BackColor = objItem.BackColor
    chk��Ч(0).BackColor = objItem.BackColor
    chk��Ч(1).BackColor = objItem.BackColor
    chk��ҩ(0).BackColor = objItem.BackColor
    chk��ҩ(1).BackColor = objItem.BackColor
    chk��ҩ.BackColor = objItem.BackColor
    
    If mblnOnePati Then
        picPati.Visible = False
    Else
        Set objGroup = tkpMain.Groups.Add(Item_�����벡��, "�����벡��")
        Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
        Set objItem.Control = picPati
        picPati.BackColor = objItem.BackColor
        chkOut.BackColor = objItem.BackColor
        chkPreOut.BackColor = objItem.BackColor
    End If
    
    Set objGroup = tkpMain.Groups.Add(Item_��ҩ����, "��ҩ����")
    objGroup.Expanded = True
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picDept
    picDept.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(Item_��ҩ;��, "��ҩ;��")
    objGroup.Expanded = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picWay
    picWay.BackColor = objItem.BackColor
    
    '-------------------------------------------------
    '����ȱʡ��ѯʱ��
    i = Val(zlDatabase.GetPara("ҩ�Ʋ�ѯ���", glngSys, pסԺҽ������, "0", Array(lblDate, dtpDate(0), dtpDate(1))))
    curDate = zlDatabase.Currentdate
    dtpDate(0).value = Format(DateAdd("d", -1 * i, curDate), "yyyy-MM-dd 00:00")
    dtpDate(1).value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpDate(0).MaxDate = dtpDate(1).value
    dtpDate(1).MaxDate = dtpDate(1).value
    
    i = Val(zlDatabase.GetPara("��ҩ��ѯ���", glngSys, pסԺҽ������, "0", Array(dtpDate(2), dtpDate(3))))
    If Not dtpDate(2).Enabled Then
        dtpDate(2).Tag = "1": dtpDate(3).Tag = "1" '��ʾ�̶�������
    Else
        dtpDate(2).Enabled = False: dtpDate(3).Enabled = False '����������״̬��ʼΪ������
    End If
    curDate = zlDatabase.Currentdate
    dtpDate(2).value = Format(DateAdd("d", -1 * i, curDate), "yyyy-MM-dd 00:00")
    dtpDate(3).value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpDate(2).MaxDate = dtpDate(3).value
    dtpDate(3).MaxDate = dtpDate(3).value
    
    'ȱʡ��ѯ��Ч
    i = Val(zlDatabase.GetPara("ҩ�Ʋ�ѯ��Ч", glngSys, pסԺҽ������, "2", Array(chk��Ч(0), chk��Ч(1))))
    If i = 2 Then
        chk��Ч(0).value = 1
        chk��Ч(1).value = 1
    Else
        chk��Ч(i).value = 1
    End If
    
    'ȱʡ��ѯ״̬
    i = Val(zlDatabase.GetPara("ҩ�Ʋ�ѯ״̬", glngSys, pסԺҽ������, "2", Array(chk��ҩ(0), chk��ҩ(1))))
    If i = 2 Then
        chk��ҩ(0).value = 1
        chk��ҩ(1).value = 1
    Else
        chk��ҩ(i).value = 1
    End If
    
    'ȱʡ�Ƿ������Ժ����
    chkOut.value = Val(zlDatabase.GetPara("ҩ�Ʋ�ѯ��Ժ����", glngSys, pסԺҽ������, "0", Array(chkOut)))
    'ȱʡ�Ƿ����Ԥ��Ժ����
    chkPreOut.value = Val(zlDatabase.GetPara("ҩ�Ʋ�ѯԤ��Ժ����", glngSys, pסԺҽ������, "0", Array(chkPreOut)))
    '����/����
    'Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
    
    'ҩ��
    Call Loadҩ��
    
    '��ҩ����
    Call LoadReqDruDep
    
    '��ҩ;��
    Call Load��ҩ;��
    
End Sub

Private Function Load��ҩ;��() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem, str��ҩIDs As String
    
    On Error GoTo errH
    
    str��ҩIDs = zlDatabase.GetPara("ҩ�Ʋ�ѯ��ҩ;��", glngSys, pסԺҽ������, "", Array(lvwWay))

    strSQL = "Select ID,����,���� From ������ĿĿ¼" & _
        " Where ���='E' And �������� in ('2', '4') And ������� IN(2,3) And (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ��������, ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwWay.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����)
        
        If str��ҩIDs <> "" Then
            If InStr("," & str��ҩIDs & ",", "," & rsTmp!ID & ",") > 0 Then
                objItem.Checked = True
            End If
        Else
            objItem.Checked = True
        End If
        rsTmp.MoveNext
    Next
    Load��ҩ;�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '��������۲���
    If InStr(mMainPrivs, "ȫԺ����") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        '����Ȩ������ֱ�����ڲ���+���ڿ�����������
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And B.������� in(1,2,3) And B.��������='����'" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.����,C.����,Nvl(B.ȱʡ,0) as ȱʡ" & _
            " From �������Ҷ�Ӧ A,������Ա B,���ű� C" & _
            " Where A.����ID=C.ID And B.����ID=A.����ID And B.��ԱID=[1]" & _
            " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
            " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
    End If
    
    cboUnit.Clear
    If InStr(mMainPrivs, "ȫԺ����") > 0 Then
        cboUnit.AddItem "���в���"
        If mlng����ID = 0 Then cboUnit.ListIndex = cboUnit.NewIndex
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!���� & "-" & rsTmp!����
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng����ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Loadҩ��() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngҩ�� As Long
    
    On Error GoTo errH
    
    lngҩ�� = Val(zlDatabase.GetPara("ҩ�Ʋ�ѯҩ��", glngSys, pסԺҽ������, "0", Array(lblҩ��, cboҩ��)))

    strSQL = _
        "Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " AND B.����ID=A.ID And B.������� IN(2,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cboҩ��.AddItem rsTmp!���� & "-" & rsTmp!����
        cboҩ��.ItemData(cboҩ��.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngҩ�� Then
            cboҩ��.ListIndex = cboҩ��.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cboҩ��.ListCount > 0 And cboҩ��.ListIndex = -1 Then cboҩ��.ListIndex = 0
    Loadҩ�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadReqDruDep() As Boolean
'���ܣ�������ҩ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng���� As Long
    
    On Error GoTo errH
    
    lng���� = Val(zlDatabase.GetPara("ҩ�Ʋ�ѯ��ҩ����", glngSys, pסԺҽ������, "0", Array(cboReqDruDep)))

    strSQL = "Select a.Id, a.����, a.����" & _
        " From ���ű� A, ��������˵�� B" & _
        " Where a.Id = b.����id And b.�������� = '��ҩ����' And (a.����ʱ�� Is Null Or Trunc(a.����ʱ��) = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        " Order By ����"

    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        
    With cboReqDruDep
        .Clear
        .AddItem "���в���"
        .ItemData(.NewIndex) = 0
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            .ItemData(.NewIndex) = rsTmp!ID
            If rsTmp!ID = lng���� Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        
        If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
    End With
    LoadReqDruDep = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tkpMain.Left = 0
    tkpMain.Top = 0
    tkpMain.Width = Me.ScaleWidth
    tkpMain.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub


Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.index)
End Sub

Private Sub optDate_Click(index As Integer)
    If index = 1 And optDate(index).value Then
        chk��ҩ(0).Enabled = False
        chk��ҩ(1).value = 1 '���ֵ�ı䣬�Զ�����Click�������ع�һ�������ȹ�����
        chk��ҩ(0).value = 0 '���ֵ�ı䣬�Զ�����Click
    Else
        chk��ҩ(0).Enabled = True
    End If
End Sub

Private Sub picCond_Resize()
    On Error Resume Next
    
    cboҩ��.Width = picCond.ScaleWidth - cboҩ��.Left
    
    dtpDate(0).Width = picCond.ScaleWidth - dtpDate(0).Left
    dtpDate(1).Width = dtpDate(0).Width
    dtpDate(2).Width = dtpDate(0).Width
    dtpDate(3).Width = dtpDate(0).Width
    
    txtNO(0).Width = picCond.ScaleWidth - txtNO(0).Left
    txtNO(1).Width = picCond.ScaleWidth - txtNO(1).Left
    
    cmdQuery.Left = picCond.ScaleWidth - cmdQuery.Width - 30
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    
    cboUnit.Left = 0
    cboUnit.Width = picPati.ScaleWidth
    
    lvwPati.Left = 0
    lvwPati.Width = picPati.ScaleWidth
    
    chkOut.Left = lvwPati.Left + lvwPati.Width - chkOut.Width - 15
    chkPreOut.Left = chkOut.Left - chkPreOut.Width - 15
    cmdAllPati.Left = picPati.ScaleWidth - cmdAllPati.Width + 15
    cmdNoPati.Left = cmdAllPati.Left - cmdNoPati.Width + 15
End Sub

Private Sub picDept_Resize()
    On Error Resume Next
    
    cboReqDruDep.Left = 0
    cboReqDruDep.Width = picPati.ScaleWidth
    
End Sub

Private Sub picWay_Resize()
    On Error Resume Next

    lvwWay.Left = 0
    lvwWay.Width = picWay.ScaleWidth
    
    cmdAllWay.Left = picWay.ScaleWidth - cmdAllWay.Width + 15
    cmdNoWay.Left = cmdAllWay.Left - cmdNoWay.Width + 15
End Sub

Private Sub txtNO_GotFocus(index As Integer)
    Call zlControl.TxtSelAll(txtNO(index))
End Sub


Private Sub txtNO_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtNO_Validate(index, False)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Len(Trim(txtNO(index).Text)) > 18 And index = 1 Then
            txtNO(index).Text = Mid(Trim(txtNO(index).Text), 18)
            MsgBox "��ҩ�ų��Ȳ��ܴ���18λ��", vbInformation, gstrSysName
        End If
        If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtNO_Validate(index As Integer, Cancel As Boolean)
    If txtNO(index).Text <> "" Then
        txtNO(index).Text = IIF(index = 0, GetFullNO(txtNO(index), 14), Trim(txtNO(index).Text))
    End If
End Sub
