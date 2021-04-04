VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmImportEPRDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ĵ���..."
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmImportEPRDemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3195
      Left            =   120
      TabIndex        =   3
      Top             =   390
      Width           =   6705
      _Version        =   589884
      _ExtentX        =   11827
      _ExtentY        =   5636
      _StockProps     =   0
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.TextBox txtFilt 
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   3690
      Width           =   3330
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5685
      TabIndex        =   5
      Top             =   4110
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "��(&O)"
      Height          =   350
      Left            =   5685
      TabIndex        =   4
      Top             =   3675
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   3375
      Top             =   4350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportEPRDemo.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportEPRDemo.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportEPRDemo.frx":0EBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblList 
      AutoSize        =   -1  'True
      Caption         =   "��ǰ��###�����÷���(&L):"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2070
   End
   Begin VB.Shape shpList 
      BorderColor     =   &H00FFC0C0&
      Height          =   3225
      Left            =   105
      Top             =   375
      Width           =   6735
   End
   Begin VB.Label lblFilt 
      AutoSize        =   -1  'True
      Caption         =   "�������(&F)"
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   3750
      Width           =   990
   End
End
Attribute VB_Name = "frmImportEPRDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngFileId As Long          '�����ļ�id
Private mlngPatient As Long         '����id���ڲ��˲����༭ʱ������ȷ������ʾ���Ƿ�����
Private mlngVisit As Long           '��ҳid��Һŵ�ID
Private mlngAdvice As Long          'ҽ��ID
Private mstrLike As String
Private mintPower As Integer    'Ȩ��
Private mblnOK As Boolean        '����

Private rptCol As ReportColumn
Private rptRcd As ReportRecord
Private rptItem As ReportRecordItem
Private rptRow As ReportRow

'################################################################################################################
'## ���ܣ�  װ��ָ��ID�ķ����б�
'##
'## ������  frmParent       :������
'##
'## ���أ�  ����ѡ�еķ���ID
'################################################################################################################
Public Function ShowMe(frmParent As Object) As Long
    Dim strCaption As String
    
    If frmParent.Name <> "frmMain" Then ShowMe = 0: Exit Function
    Err = 0: On Error Resume Next
    With frmParent.Document
        strCaption = .EPRFileInfo.����
        mlngFileId = .EPRFileInfo.ID
        mlngPatient = .EPRPatiRecInfo.����ID
        mlngVisit = .EPRPatiRecInfo.��ҳID
        mlngAdvice = .EPRPatiRecInfo.ҽ��id
    End With
    
    Err = 0: On Error GoTo 0
    Call zlGetPower
    Call FillEPRDemos
    Me.lblList.Caption = "��ǰ��" & strCaption & "�����÷���(&L):"
    Me.Show vbModal, frmParent
    If mblnOK Then ShowMe = Me.rptList.FocusedRow.Record.Item(1).Value
    Unload Me
End Function

Private Function zlGetPower() As Integer
    '���ܣ���õ�ǰ�û���ʾ�������Ȩ��
    '���أ�ʾ������Ȩ����ֵ
    Dim strPrivs As String
    strPrivs = GetPrivFunc(glngSys, 1070)
    If InStr(1, strPrivs, "ȫԺ��������") <> 0 Then
        mintPower = 0
    ElseIf InStr(1, strPrivs, "���Ҳ�������") <> 0 Then
        mintPower = 1
    ElseIf InStr(1, strPrivs, "���˲�������") <> 0 Then
        mintPower = 2
    Else
        mintPower = -1
    End If
    zlGetPower = mintPower
End Function

Private Function FillEPRDemos() As Long
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select l.Id, l.���, l.����, l.����, Nvl(l.����,'δ����') As ����,l.˵��, l.ͨ�ü�" & vbNewLine & _
            "From ��������Ŀ¼ l, Table(Cast(f_Segment_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) u" & vbNewLine & _
            "Where l.�ļ�id = [1] And Nvl(l.����, 0) = [5] And l.Id = To_Number(u.����)"
    Select Case mintPower
    Case 0
    Case 1
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

    Case Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
    End Select
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId, mlngPatient, mlngVisit, mlngAdvice, 2)
    
    Me.rptList.Records.DeleteAll
    If rsTemp.EOF Then cmdOK.Enabled = False
    Do While Not rsTemp.EOF
        Set rptRcd = Me.rptList.Records.Add()
        Set rptItem = rptRcd.AddItem(CStr(Nvl(rsTemp!ͨ�ü�, 0))): rptItem.Icon = rptItem.Value
        rptRcd.AddItem CStr(rsTemp!ID)
        rptRcd.AddItem CStr(rsTemp!����)
        rptRcd.AddItem CStr(rsTemp!���)
        rptRcd.AddItem CStr(rsTemp!����)
        rptRcd.AddItem CStr("" & rsTemp!����)
        rptRcd.AddItem CStr("" & rsTemp!˵��)
        rsTemp.MoveNext
    Loop
    Me.rptList.Populate
    If Me.rptList.Rows.Count > 1 And Me.rptList.FocusedRow Is Nothing Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(1)
    End If
    
    FillEPRDemos = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    FillEPRDemos = Me.rptList.Records.Count
End Function

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim blnNoSelected As Boolean
    If Me.rptList.FocusedRow.GroupRow Then
        blnNoSelected = True
    ElseIf Me.rptList.FocusedRow Is Nothing Then
        blnNoSelected = True
    End If
    
    If blnNoSelected Then
        MsgBox "û��ѡ�з��ģ����ܴ򿪣�", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnOK = True: Me.Hide
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    mstrLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    With Me.rptList
        Set rptCol = .Columns.Add(0, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(1, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(2, "����", 50, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(3, "���", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(4, "����", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(5, "����", 60, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(6, "˵��", 200, True): rptCol.Editable = False: rptCol.Groupable = False
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(2)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(3)
        
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.Record Is Nothing Then Exit Sub
    
    cmdOK_Click
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    cmdOK_Click
End Sub

Private Sub txtFilt_GotFocus()
    Me.txtFilt.SelStart = 0: Me.txtFilt.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtFilt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        For Each rptRcd In Me.rptList.Records
            If Trim(Me.txtFilt.Text) = "" Then
                rptRcd.Visible = True
            Else
                rptRcd.Visible = (rptRcd(5).Value Like IIf(mstrLike <> "", "*", "") & Trim(Me.txtFilt.Text) & "*")
            End If
        Next
        Me.rptList.Populate
        If Me.rptList.Rows.Count > 0 And Me.rptList.FocusedRow Is Nothing Then
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
        Call txtFilt_GotFocus
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        If KeyAscii = Asc("*") Or KeyAscii = Asc("?") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

