VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugListAsk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ѯ��������"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraRangeSelect 
      Caption         =   "��Χѡ��"
      Height          =   1875
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3945
      Begin VB.ComboBox cob�ⷿ 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2460
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   270
         Left            =   3615
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1395
         Width           =   255
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1005
         TabIndex        =   3
         Top             =   1380
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   1005
         TabIndex        =   2
         Top             =   1020
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   101122051
         CurrentDate     =   36257.9583333333
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   300
         Left            =   1005
         TabIndex        =   1
         Top             =   660
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   101122051
         CurrentDate     =   36257
      End
      Begin VB.Label lbl�ⷿ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ"
         Height          =   180
         Left            =   615
         TabIndex        =   11
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "ҩƷ����"
         Height          =   180
         Left            =   255
         TabIndex        =   10
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label lblStartDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblEndDate 
         BackStyle       =   0  'Transparent
         Caption         =   "��ֹ����"
         Height          =   180
         Left            =   255
         TabIndex        =   8
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4185
      TabIndex        =   6
      Top             =   555
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4185
      TabIndex        =   5
      Top             =   150
      Width           =   1100
   End
End
Attribute VB_Name = "frmDrugListAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnAskOk As Boolean

Public InDrugId As Long            'ҩƷid
Public InDrugName  As String       'ҩƷ����
Public InDrugStAndard As String      'ҩƷ���
Public inDeptId As Long
Public InDrugUnit As String          'ҩƷ��λ
Public intUnitLevel As Integer       '��λ����

Dim blnFirst As Boolean
Dim rsDrug As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim PrvPara   As Byte       '���� 0����ʾ��������˫��ƥ��,1: ��ʾ�������ƴ���ƥ��
Dim StrFh As String         'ǰƥ�����%"
Dim strsql As String
Dim Lng���� As Long       '���浱ǰѡ�����Ĳ���
Dim sngLeft, sngTop As Single
Dim Bln����ҩ As Boolean '��ʾ�Ƿ���в�ѯ����ҩ��Ȩ��
Dim Bln�г�ҩ As Boolean '��ʾ�Ƿ���в�ѯ�г�ҩ��Ȩ��
Dim Bln�в�ҩ As Boolean '��ʾ�Ƿ���в�ѯ�в�ҩ��Ȩ��
Dim Str���� As String



Private Sub cmdCancel_Click()
    blnAskOk = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandle
    If Me.txt����.Tag = "C" Then
        MsgBox "��ѡ��ҩƷ!", vbInformation, gstrSysName
        Me.txt����.SetFocus
        Exit Sub
    End If
    
    If InStr(gstrStockSearchPrivs, "����ҩ") <> 0 Then
        Bln����ҩ = True
    Else
        Bln����ҩ = False
    End If
    
    If InStr(gstrStockSearchPrivs, "�г�ҩ") <> 0 Then
        Bln�г�ҩ = True
    Else
        Bln�г�ҩ = False
    End If
    
    If InStr(gstrStockSearchPrivs, "�в�ҩ") <> 0 Then
        Bln�в�ҩ = True
    Else
        Bln�в�ҩ = False
    End If
    
    Str���� = "''"
    If Bln����ҩ Then Str���� = "'����ҩ'"
    If Bln�г�ҩ Then
        If Bln����ҩ Then
            Str���� = Str���� & ",'�г�ҩ'"
        Else
            Str���� = "'�г�ҩ'"
        End If
    End If
    If Bln�в�ҩ Then
        If Bln�г�ҩ Or Bln����ҩ Then
            Str���� = Str���� & ",'�в�ҩ'"
        Else
            Str���� = "'�в�ҩ'"
        End If
    End If

    Set rsTemp = New ADODB.Recordset
    strsql = "Select A.ҩƷid From ҩƷĿ¼ A,ҩƷ��Ϣ B Where A.ҩ��id=B.ҩ��id And A.ҩƷid=[1] And B.���ʷ��� In (" & Str���� & ")"
    Call SQLTest(App.Title, Me.Caption, strsql)
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "cmdOK_Click", InDrugId)
    Call SQLTest

    If rsTemp.RecordCount = 0 Then
        MsgBox "��û�в�ѯ��ҩƷ��Ȩ�ޣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus
        Exit Sub
    End If
    rsTemp.Close
    
    blnAskOk = True
    Me.Hide
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    Dim RecReturn As Recordset
    Dim intLevel As Integer
    
    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, cob�ⷿ.ItemData(cob�ⷿ.ListIndex))
 
      If RecReturn.RecordCount = 0 Then
         Unload FrmҩƷѡ����
         Exit Sub
      End If
    
'InDrugId = RecReturn!ҩƷID
'    InDrugName = RecReturn!��Ʒ��
'    InDrugStAndard = IIf(IsNull(RecReturn!���), " ", RecReturn!���)
'    InDrugUnit = IIf(IsNull(RecReturn!������λ), " ", RecReturn!������λ)
'    Me.txt����.Text = InDrugName
'    Me.txt����.Tag = InDrugId
'
    'Unload FrmҩƷѡ����
    With RecReturn
        InDrugId = !ҩƷID
        InDrugName = "[" & !ҩƷ���� & "]" & !��Ʒ��
        InDrugStAndard = IIf(IsNull(!���), " ", !���)
        Me.txt����.Text = InDrugName
        Me.txt����.Tag = InDrugId
        intLevel = frmDrugQuery.intChoose����
        
        Select Case intLevel
            Case 1
                InDrugUnit = !�ۼ۵�λ
                frmDrugList.Tag = "1"
            Case 2
                InDrugUnit = !���ﵥλ
                frmDrugList.Tag = !�����װ
            Case 3
                InDrugUnit = !ҩ�ⵥλ
                frmDrugList.Tag = !ҩ���װ
            Case 4
                InDrugUnit = !סԺ��λ
                frmDrugList.Tag = !סԺ��װ
        End Select
         
    End With
    
    
End Sub

Private Sub cob�ⷿ_Validate(Cancel As Boolean)
    Me.txt����.Tag = "C"
    txt����_Validate False
End Sub

Private Sub dtpStartDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpEndDate.Value = Me.dtpStartDate.Value
    End If
End Sub

Private Sub dtpEndDate_Change()
    If Me.dtpStartDate.Value > Me.dtpEndDate.Value Then
        Me.dtpStartDate.Value = Me.dtpEndDate.Value
    End If
End Sub

Private Sub Form_Activate()
    Dim i As Long
    If Not blnFirst Then Exit Sub
    PrvPara = Val(GetSetting(appName:="ZLSOFT", Section:="����", Key:="ƥ�䷽��", Default:="0"))
    StrFh = IIf(PrvPara = 0, "%", "")
    
    Me.txt����.Tag = InDrugId
    Me.txt����.Text = InDrugName
    Me.cob�ⷿ.Clear
    With frmDrugQuery.cob�ⷿ
         For i = 0 To .ListCount - 1
            cob�ⷿ.AddItem .List(i)
            cob�ⷿ.ItemData(cob�ⷿ.NewIndex) = .ItemData(i)
            If .ItemData(i) = inDeptId Then
                cob�ⷿ.ListIndex = cob�ⷿ.NewIndex
            End If
         Next
    End With
    blnFirst = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strsql As String
    Dim i As Long
    Me.dtpEndDate.MaxDate = Currentdate()
    Me.dtpEndDate.Value = Me.dtpEndDate.MaxDate
    Me.dtpStartDate.MaxDate = Me.dtpEndDate.Value
    
    blnFirst = True
    Me.dtpStartDate.Value = DateAdd("m", -1, Me.dtpEndDate.Value)
End Sub

Private Sub txt����_Change()
    If blnFirst Then Exit Sub
    Me.txt����.Tag = "C"
End Sub

Private Sub txt����_DblClick()
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
End Sub

Private Sub txt����_GotFocus()
    txt����.SelStart = 0
    txt����.SelLength = LenB(StrConv(txt����, vbFromUnicode))
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim intLevel As Integer
    
    
    
    If InStr(Me.txt����.Text, "'") <> 0 Then
        MsgBox "�����г����˷Ƿ��ַ�:'", vbInformation, gstrSysName
        Cancel = True
        txt����.SelStart = 0
        txt����.SelLength = LenB(StrConv(txt����, vbFromUnicode))
        Exit Sub
    End If
    
    Set rsDrug = New ADODB.Recordset
    If Me.txt����.Tag <> "C" Or Me.txt���� = "" Then Exit Sub
    
    sngLeft = Me.Left + fraRangeSelect.Left + txt����.Left
    sngTop = Me.Top + Me.Height - Me.ScaleHeight + txt����.Top + txt����.Height
    If sngTop + 4730 > Screen.Height Then
        sngTop = sngTop - txt����.Height - 4730
    End If
           
    Dim strkey As String
    
    strkey = txt����.Text
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
           
    Set rsDrug = FrmҩƷ��ѡѡ����.ShowME(Me, 1, cob�ⷿ.ItemData(cob�ⷿ.ListIndex), , , strkey, sngLeft, sngTop)
    With rsDrug
        If .RecordCount = 0 Then
            Cancel = True
            txt����.SelStart = 0
            txt����.SelLength = LenB(StrConv(txt����, vbFromUnicode))
            Exit Sub
        End If
        intLevel = frmDrugQuery.intChoose����
        
        If .RecordCount > 1 Then
            .MoveFirst
            InDrugId = !ҩƷID
            InDrugName = "[" & !ҩƷ���� & "]" & !��Ʒ��
            InDrugStAndard = IIf(IsNull(!���), " ", !���)
            Select Case intLevel
                Case 1
                    InDrugUnit = !�ۼ۵�λ
                    frmDrugList.Tag = "1"
                Case 2
                    InDrugUnit = !���ﵥλ
                    frmDrugList.Tag = !�����װ
                Case 3
                    InDrugUnit = !ҩ�ⵥλ
                    frmDrugList.Tag = !ҩ���װ
                Case 4
                    InDrugUnit = !סԺ��λ
                    frmDrugList.Tag = !סԺ��װ
            End Select
         Else
            InDrugId = !ҩƷID
            InDrugName = "[" & !ҩƷ���� & "]" & !��Ʒ��
            InDrugStAndard = IIf(IsNull(!���), " ", !���)
            Select Case intLevel
                Case 1
                    InDrugUnit = !�ۼ۵�λ
                    frmDrugList.Tag = "1"
                Case 2
                    InDrugUnit = !���ﵥλ
                    frmDrugList.Tag = !�����װ
                Case 3
                    InDrugUnit = !ҩ�ⵥλ
                    frmDrugList.Tag = !ҩ���װ
                Case 4
                    InDrugUnit = !סԺ��λ
                    frmDrugList.Tag = !סԺ��װ
            End Select
        End If
    End With
    
    Me.txt����.Text = InDrugName
    Me.txt����.Tag = InDrugId
End Sub


Private Function GetLevel(ByVal lng����id As Long) As Integer
    '�жϸò���ֻ��ҩ�������ҩ��
    Dim rsTemp As New ADODB.Recordset
    Dim intChoose���� As Integer
    Dim strsql As String
    
    On Error GoTo errHandle
    strsql = "Select * From ��������˵�� " & _
        " Where ����id=[1] And �������� IN ('��ҩ��','��ҩ��','��ҩ��','�Ƽ���','��ҩ��','��ҩ��','��ҩ��') "
    Set rsTemp = zldatabase.OpenSQLRecord(strsql, "GetLevel", lng����id)
    If Not rsTemp.EOF Then
        Select Case rsTemp!�������
            Case 0
                intChoose���� = 3
            Case 1, 3
                intChoose���� = 2
            Case 2
                intChoose���� = 4
            Case Else
                intChoose���� = 1
        End Select
    Else
        intChoose���� = 1
    End If
   
    rsTemp.Close
    
    GetLevel = intChoose����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

