VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEInvoice 
   BorderStyle     =   0  'None
   Caption         =   "����Ʊ�ݹ���"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkסԺ 
      Caption         =   "סԺԤ��"
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "����Ԥ��"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   6000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmd�ͻ��� 
      Caption         =   "��"
      Height          =   280
      Left            =   4440
      TabIndex        =   12
      Top             =   705
      Width           =   280
   End
   Begin VB.TextBox txt�ͻ��� 
      Height          =   300
      Left            =   1080
      TabIndex        =   11
      Top             =   700
      Width           =   3345
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����Ʊ������"
      Height          =   300
      Left            =   7680
      TabIndex        =   10
      Top             =   6000
      Width           =   1335
   End
   Begin VB.OptionButton Option����Ʊ�� 
      Caption         =   "�����õ���Ʊ��"
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option����Ʊ�� 
      Caption         =   "���õ���Ʊ��"
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option����Ʊ�� 
      Caption         =   "���ͻ������õ���Ʊ��"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CheckBox chkƽ̨���� 
      Caption         =   "����ƽ̨����ֽ��Ʊ��"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   300
      Left            =   4800
      Picture         =   "frmEInvoice.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   700
      Width           =   300
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFҽ�� 
      Height          =   4620
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   3780
      _cx             =   6667
      _cy             =   8149
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoice.frx":0A02
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid VSF�ͻ��� 
      Height          =   4620
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   4860
      _cx             =   8572
      _cy             =   8149
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEInvoice.frx":0AA3
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblԤ�� 
      Caption         =   "����Ԥ�����"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   6030
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "�Һ�"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   300
      Width           =   615
   End
   Begin VB.Label lbl�ͻ��� 
      Alignment       =   1  'Right Justify
      Caption         =   "�ͻ���"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblҽ�� 
      Caption         =   "ҽ������"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   735
      Width           =   735
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000003&
      X1              =   6000
      X2              =   9720
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmEInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private mint���� As Integer  '1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
Private mrs������� As ADODB.Recordset
Private mstr�Һ� As String, mstr�շ� As String, mstrԤ�� As String, mstr���� As String
Private Type Para
    int���  As Integer  '0-�����֣�1-���2-סԺ
    int���õ���Ʊ��  As Integer  '0-�����õ���Ʊ�ݣ�1-���õ���Ʊ�ݣ�2-��վ�����õ���Ʊ��
    bln����Ʊ�ݹ���  As Boolean  'True-HIS����Ʊ�ݣ�False-����ƽ̨����Ʊ�ݣ�
    strҽ������  As String  '�ַ����ĸ�ʽΪ��"0:"��"1:998"��0��ʾδ���ã�1��ʾ���ã�: ���Ϊ����(��998)������Ϊ�ձ�ʾ����ҽ��������
End Type
Private mPara As Para
Private Enum Page
    Pg_�շ� = 1
    Pg_Ԥ�� = 2
    Pg_���� = 3
    Pg_�Һ� = 4
    Pg_���￨ = 5
End Enum
Private mIndex As Integer

Private Sub chk����_Click()
    If chkסԺ.value = 0 Then
        If chk����.value = 0 Then
            MsgBox "������������һ��Ԥ������!", vbInformation, gstrSysName
            chk����.value = 1
        End If
    End If
End Sub

Private Sub chkסԺ_Click()
    If chk����.value = 0 Then
        If chkסԺ.value = 0 Then
            MsgBox "������������һ��Ԥ������!", vbInformation, gstrSysName
            chkסԺ.value = 1
        End If
    End If
End Sub

Private Sub cmdDelete_Click()
    If VSF�ͻ���.Enabled = False Then Exit Sub
    If VSF�ͻ���.Row <= 0 Then Exit Sub
    If mPara.int���õ���Ʊ�� = mIndex Then
        If CheckHaveData = False Then Exit Sub
    End If
    Call Delete�ͻ���(VSF�ͻ���.Row)
    zlcontrol.ControlSetFocus VSF�ͻ���
End Sub

Private Sub cmd�ͻ���_Click()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = GetControlRect(txt�ͻ���.hWnd)
    strSQL = "Select Rownum As id,Upper(����վ) as ����վ, Upper(��;) as ��;,Upper(����) as ����  From zlClients "
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�ͻ���", 1, "", "��ѡ��ͻ���", False, False, True, vRect.Left, vRect.Top, txt�ͻ���.Height, blnCancel, False, False, "%" & Trim(txt�ͻ���.Text) & "%", "bytSize=1")
    EnableWindow Me.hWnd, True '������ShowSQLSelect����Զ���������
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.State = 0 Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    Call Add�ͻ���(NVL(rsTmp!����վ), NVL(rsTmp!����), NVL(rsTmp!��;))
End Sub

Private Sub cmd����_Click()
    '���õ��ӷ�Ʊ�豸�������ýӿ�
    Dim objEInvoice As Object
    
    If zlCreatEInvoice(objEInvoice, Me) = False Then Exit Sub
    If objEInvoice Is Nothing Then Exit Sub
    Call objEInvoice.zlEInvoiceSet(Me)
    Call objEInvoice.zlTerminate
    EnableWindow Me.hWnd, True '�����˸ð�ť����Զ���������
End Sub

Private Sub Form_GotFocus()
    If Option����Ʊ��(0).value Then
        zlcontrol.ControlSetFocus Option����Ʊ��(0)
    ElseIf Option����Ʊ��(1).value Then
        zlcontrol.ControlSetFocus Option����Ʊ��(1)
    Else
        zlcontrol.ControlSetFocus Option����Ʊ��(2)
    End If
End Sub

Private Sub Form_Load()
    Call InitPara
    Call InitData
End Sub

Private Sub InitData()
    On Error GoTo ErrHand
    Call SetEnable
    If Get�������(mrs�������) = False Then Exit Sub
    Call Load�ͻ�����Ϣ
    Call Loadҽ�����
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mint���� = 0
End Sub

Private Function Get�������(ByRef rsTmp As ADODB.Recordset) As Boolean
    '���ܣ���ȡ���еı������
    Dim strSQL As String
    
    On Error GoTo ErrHand
    Set rsTmp = New ADODB.Recordset
    strSQL = "Select ���,����,˵��,ҽԺ���� From ������� Where Nvl(�Ƿ��ֹ,0)=0 Order By ���"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption)
    Get������� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Load�ͻ�����Ϣ()
    Dim i As Integer
    Dim strSQL  As String, rsData As ADODB.Recordset

     With VSF�ͻ���
         If .Enabled Then
             .Clear 2
             strSQL = "Select b.����վ, b.����, b.��; From ����Ʊ��վ����� A, zlClients B Where a.վ�� = b.����վ And a.����=[1] order by b.����վ "
             Set rsData = zldatabase.OpenSQLRecord(strSQL, "��ȡ����Ʊ��վ�����", mint����)
             If Not rsData.EOF Then
                 .Rows = rsData.RecordCount + 1
                 For i = 1 To rsData.RecordCount
                     .TextMatrix(i, .ColIndex("�ͻ�������")) = rsData!����վ
                     .TextMatrix(i, .ColIndex("����")) = NVL(rsData!����)
                     .TextMatrix(i, .ColIndex("��;")) = NVL(rsData!��;)
                     rsData.MoveNext
                 Next
             End If
         End If
     End With
End Sub

Private Sub Loadҽ�����()
    Dim i As Integer, j As Integer
    Dim strҽ�� As String, blnҽ������ As Boolean
    Dim varTmp As Variant
    
    If mrs������� Is Nothing Then Exit Sub
    If mrs�������.RecordCount = 0 Then Exit Sub
    mrs�������.MoveFirst
    
    With VSFҽ��
        .Clear 2
        .Rows = mrs�������.RecordCount + 1
        varTmp = Split(mPara.strҽ������ & ":::", ":")
        blnҽ������ = varTmp(0) = 1: strҽ�� = varTmp(1)
        For j = 1 To mrs�������.RecordCount
            .TextMatrix(j, .ColIndex("�������")) = mrs�������!���
            If blnҽ������ And InStr("," & strҽ�� & ",", "," & NVL(mrs�������!���) & ",") > 0 Then
                .TextMatrix(j, .ColIndex("����")) = "-1"
                .TextMatrix(j, .ColIndex("ԭ����")) = "1"
            End If
            .TextMatrix(j, .ColIndex("��������")) = mrs�������!����
            mrs�������.MoveNext
        Next
    End With
End Sub

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function Save����Ʊ�ݿ���() As Boolean
    Dim strSQL As String, i As Integer, j As Integer
    Dim str�ͻ��� As String, blnTrans As Boolean
    Dim strTmp As String, intTmp As Integer
    Dim str���� As String

    On Error GoTo ErrHand

    If Option����Ʊ��(0).value Then
        strTmp = "0"
    ElseIf Option����Ʊ��(1).value Then
        strTmp = "1"
    Else
        strTmp = "2"
    End If
    strTmp = strTmp & "|" & IIF(chkƽ̨����.value = 1, "1", "0")
    str���� = ""
    With VSFҽ��
        For j = 1 To .Rows - 1
            If .TextMatrix(j, .ColIndex("����")) = "-1" Then str���� = str���� & "," & .TextMatrix(j, .ColIndex("�������"))
        Next
    End With
    str���� = Mid(str����, 2)
    If str���� = "" Then
        strTmp = strTmp & "|" & "0:"
    Else
        strTmp = strTmp & "|" & "1:" & str����
    End If
    Select Case mint����
        Case Pg_�Һ�
            zldatabase.SetPara "�Һŵ���Ʊ�ݿ���", strTmp, glngSys
        Case Pg_�շ�
            zldatabase.SetPara "�շѵ���Ʊ�ݿ���", strTmp, glngSys
        Case Pg_Ԥ��
            If chk����.value = 1 And chkסԺ.value = 1 Then
                intTmp = 0
            ElseIf chk����.value = 1 Then
                intTmp = 1
            Else
                intTmp = 2
            End If
            zldatabase.SetPara "Ԥ������Ʊ�ݿ���", intTmp & "|" & strTmp, glngSys
        Case Pg_����
            zldatabase.SetPara "���ʵ���Ʊ�ݿ���", strTmp, glngSys
        Case Pg_���￨
            zldatabase.SetPara "���￨����Ʊ�ݿ���", strTmp, glngSys
        Case Else
    End Select

    str�ͻ��� = ""
    With VSF�ͻ���
        For j = 1 To .Rows - 1
            str�ͻ��� = str�ͻ��� & "," & .TextMatrix(j, .ColIndex("�ͻ�������"))
        Next
    End With
    str�ͻ��� = Mid(str�ͻ���, 2)
    strSQL = "Zl_����Ʊ��վ�����_Update( " & mint���� & "," & IIF(str�ͻ��� = "", "NULL", "'" & str�ͻ��� & "'") & ")"
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    Save����Ʊ�ݿ��� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckƱ��վ��() As Boolean
    '����:���ͻ������õ���Ʊ��ʱ,����Ƿ�ѡ���˿ͻ���
    Dim str�ͻ��� As String, i As Integer
    
    str�ͻ��� = ""
    With VSF�ͻ���
        For i = 1 To .Rows - 1
            str�ͻ��� = str�ͻ��� & "," & .TextMatrix(i, .ColIndex("�ͻ�������"))
        Next
    End With
     str�ͻ��� = Mid(str�ͻ���, 2)
     CheckƱ��վ�� = str�ͻ��� <> ""
End Function

Private Sub Add�ͻ���(ByVal str�ͻ��� As String, ByVal str���� As String, ByVal str��; As String)
    Dim i As Integer
    If str�ͻ��� = "" Then Exit Sub
    With VSF�ͻ���
        For i = 1 To .Rows - 1
            If str�ͻ��� = .TextMatrix(i, .ColIndex("�ͻ�������")) Then Exit Sub
        Next
        If .TextMatrix(.Rows - 1, .ColIndex("�ͻ�������")) <> "" Then .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("�ͻ�������")) = str�ͻ���
        .TextMatrix(.Rows - 1, .ColIndex("����")) = str����
        .TextMatrix(.Rows - 1, .ColIndex("��;")) = str��;
    End With
End Sub

Private Sub Delete�ͻ���(ByVal intRow As Integer)
    Dim i As Integer
    If intRow = 0 Then Exit Sub
    With VSF�ͻ���
        If intRow = 1 And .Rows = 2 Then
            .TextMatrix(intRow, .ColIndex("�ͻ�������")) = ""
            .TextMatrix(intRow, .ColIndex("����")) = ""
            .TextMatrix(intRow, .ColIndex("��;")) = ""
        Else
            .RemoveItem intRow
        End If
    End With
End Sub

Private Sub InitPara()
    Dim strTmp As String, varTmp As Variant
    
    Select Case mint����
        Case Pg_�Һ�
            strTmp = zldatabase.GetPara("�Һŵ���Ʊ�ݿ���", glngSys, , "0|1|0:")
        Case Pg_�շ�
            strTmp = zldatabase.GetPara("�շѵ���Ʊ�ݿ���", glngSys, , "0|1|0:")
        Case Pg_Ԥ��
            strTmp = zldatabase.GetPara("Ԥ������Ʊ�ݿ���", glngSys, , "0|0|1|0:")
        Case Pg_����
            strTmp = zldatabase.GetPara("���ʵ���Ʊ�ݿ���", glngSys, , "0|1|0:")
        Case Pg_���￨
            strTmp = zldatabase.GetPara("���￨����Ʊ�ݿ���", glngSys, , "0|1|0:")
    End Select
    varTmp = Split(strTmp & "||||", "|")
    If mint���� = Pg_Ԥ�� Then
        mPara.int��� = varTmp(0)
        chk����.value = IIF(mPara.int��� <> 2, 1, 0)
        chkסԺ.value = IIF(mPara.int��� <> 1, 1, 0)
        mPara.int���õ���Ʊ�� = varTmp(1)
        mPara.bln����Ʊ�ݹ��� = varTmp(2) = 1
        mPara.strҽ������ = varTmp(3)
        Exit Sub
    End If
    mPara.int���õ���Ʊ�� = varTmp(0)
    mPara.bln����Ʊ�ݹ��� = varTmp(1) = 1
    mPara.strҽ������ = varTmp(2)
End Sub

Private Sub SetEnable(Optional ByVal intIndex As Integer = -1)
    Dim i As Integer, intTmp As Integer
    If intIndex = -1 Then
        intTmp = mPara.int���õ���Ʊ��
        Option����Ʊ��(0).value = intTmp = 0
        Option����Ʊ��(1).value = intTmp = 1
        Option����Ʊ��(2).value = intTmp = 2
        cmd�ͻ���.Enabled = intTmp = 2
        txt�ͻ���.Enabled = intTmp = 2
        VSF�ͻ���.Enabled = intTmp = 2
        VSFҽ��.Enabled = intTmp > 0
        chkƽ̨����.value = IIF(mPara.bln����Ʊ�ݹ���, 1, 0)
        cmdDelete.Enabled = intTmp = 2
    Else
        cmd�ͻ���.Enabled = intIndex = 2
        txt�ͻ���.Enabled = intIndex = 2
        VSF�ͻ���.Enabled = intIndex = 2
        VSFҽ��.Enabled = intIndex > 0
        cmdDelete.Enabled = intIndex = 2
    End If
    
End Sub

Private Sub Clear�ͻ�����Ϣ()
    '����:��տͻ�����Ϣ
    txt�ͻ���.Text = ""
    VSF�ͻ���.Rows = 2
    VSF�ͻ���.Clear 2
End Sub

Private Sub Clearҽ�����()
    '����:���ҽ�����
    VSFҽ��.Cell(flexcpText, 1, 0, VSFҽ��.Rows - 1, 0) = 0
End Sub

Public Sub InitMe(ByVal int���� As Integer)
    '���ܣ���ʼ������
    mint���� = int����
    Select Case int����
        Case Pg_�Һ�
            lbl.Caption = "�Һ�"
        Case Pg_�շ�
            lbl.Caption = "�շ�"
        Case Pg_Ԥ��
            lbl.Caption = "Ԥ��"
            lblԤ��.Visible = True: chk����.Visible = True: chkסԺ.Visible = True
        Case Pg_����
            lbl.Caption = "����"
        Case Pg_���￨
            lbl.Caption = "���￨"
    End Select
End Sub

Private Function zlCreatEInvoice(ByRef objEInvoice As Object, ByVal frmMain As Object) As Boolean
    Dim strExtend As String
    err = 0: On Error Resume Next
    Set objEInvoice = CreateObject("zlPublicExpense.clsPubEInvoice")
    If err <> 0 Then
        MsgBox "�����ڿ��õĵ���Ʊ�ݽӿڲ���(zlPublicExpense.clsPubEInvoice)������ϵͳ����Ա��ϵ,��ϸ�Ĵ�����ϢΪ:" & vbCrLf & err.Description, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    zlCreatEInvoice = objEInvoice.zlInitialize(frmMain, 0, gcnOracle, glngSys, glngModul, True, strExtend)
End Function

Private Sub Option����Ʊ��_Click(Index As Integer)
    If lbl.Tag = "1" Then Exit Sub
    If RemindUser(Index) = False Then
        lbl.Tag = "1"
        Option����Ʊ��(mIndex).value = True
        lbl.Tag = "": Exit Sub
    End If
    Select Case Index
        Case 2
            SetEnable (Index)
            Call Load�ͻ�����Ϣ
            Call Loadҽ�����
        Case 1
            Call SetEnable(Index)
            Call Clear�ͻ�����Ϣ
            Call Loadҽ�����
        Case Else
            Call Clear�ͻ�����Ϣ
            Call Clearҽ�����
            Call SetEnable(Index)
    End Select
    mIndex = Index
End Sub

Private Sub txt�ͻ���_KeyPress(KeyAscii As Integer)
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    If KeyAscii <> 13 Then Exit Sub
    vRect = GetControlRect(txt�ͻ���.hWnd)
    strSQL = "Select Rownum As id,Upper(����վ) as ����վ, Upper(��;) as ��;,Upper(����) as ����  From zlClients " & _
                  "Where ����վ Like Upper([1]) Or ��; Like Upper([1]) Or ���� Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(����վ)) Like Upper([1]) Or Upper(zlPinYinCode(��;)) Like Upper([1]) Or Upper(zlPinYinCode(����)) Like Upper([1]) Order By ����վ "
    Set rsTmp = zldatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�ͻ���", 1, "", "��ѡ��ͻ���", False, False, True, vRect.Left, vRect.Top, txt�ͻ���.Height, blnCancel, False, False, "%" & Trim(txt�ͻ���.Text) & "%", "bytSize=1")
    EnableWindow Me.hWnd, True '������ShowSQLSelect����Զ���������
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.State = 0 Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    Call Add�ͻ���(NVL(rsTmp!����վ), NVL(rsTmp!����), NVL(rsTmp!��;))
    
End Sub

Private Function RemindUser(ByVal intIndex As Integer) As Boolean
    '�û��л����õ���Ʊ�ݷ�ʽʱ��ʾ�û�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, i As Integer
    Dim str���� As String, intƱ�� As Integer  '1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    
    If CheckHaveData = False Then Exit Function
    If intIndex > mIndex Then RemindUser = True: Exit Function
    '������ֵ�Ƿ�ı�
    With VSF�ͻ���
        For i = 1 To .Rows - 1
            strTmp = strTmp & "," & .TextMatrix(i, .ColIndex("�ͻ�������"))
        Next
    End With
    strTmp = Mid(strTmp, 2)
    If strTmp <> "" Then
        If MsgBox("�ı䡾����Ʊ�����÷�ʽ����֮ǰ���޸Ľ�����ա�" & vbCrLf & "�Ƿ�ȷ�ϸı䣿", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            Else
                RemindUser = True: Exit Function
            End If
    End If
    
    If intIndex = 1 Then RemindUser = True: Exit Function
    
    With VSFҽ��
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) = "-1" Then strTmp = strTmp & "," & .TextMatrix(i, .ColIndex("�������"))
        Next
    End With
    strTmp = Mid(strTmp, 2)
    If strTmp <> "" Then
        If MsgBox("�ı䡾����Ʊ�����÷�ʽ����֮ǰ���޸Ľ�����ա�" & vbCrLf & "�Ƿ�ȷ�ϸı䣿", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    RemindUser = True
End Function

Private Function CheckHaveData() As Boolean
    '�û��л����õ���Ʊ�ݷ�ʽʱ����Ƿ��Ѿ���������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str���� As String, intƱ�� As Integer  '1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
    If mIndex = 0 Then CheckHaveData = True: Exit Function
    
    '����Ƿ��Ѿ�����������
    strSQL = "Select 1 From ����Ʊ��ʹ�ü�¼ Where Ʊ�� = [1] And Rownum < 2 "
    Select Case mint����
        Case 1 '�Һ�
            intƱ�� = 4
            str���� = "�Һ�"
        Case 2 '�շ�
            intƱ�� = 1
            str���� = "�շ�"
        Case 3 'Ԥ��
            intƱ�� = 2
            str���� = "Ԥ��"
        Case 4 '����
            intƱ�� = 3
            str���� = "����"
        Case 5 '���￨
            intƱ�� = 5
            str���� = "���￨"
    End Select
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, intƱ��)
    If Not rsTmp.EOF Then
        If MsgBox(str���� & "ҵ���Ѳ�������Ʊ��ʹ�ü�¼����������˲���������Ӱ�쵽" & str���� & "ҵ���Ʊ��ʹ�ü���ӡ�����п������Ʊ��ʹ�����ݵĻ��ҡ�" & vbCrLf & "�Ƿ�ȷ�ϵ���������", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        Else
            CheckHaveData = True: Exit Function
        End If
    End If
    
    CheckHaveData = True
End Function

Private Sub VSFҽ��_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then
        With VSFҽ��
        If .TextMatrix(Row, Col) = "-1" And .TextMatrix(Row, .ColIndex("ԭ����")) = "1" Then
            If CheckHaveData = False Then
                Cancel = True
            End If
        End If
        End With
    End If
End Sub

Public Function Check����Ʊ��Valid() As Boolean
    If Option����Ʊ��(2).value Then
        If CheckƱ��վ�� = False Then
            MsgBox "���ͻ������õ���Ʊ��ʱ,������������һ���ͻ���!", vbInformation, gstrSysName
            zlcontrol.ControlSetFocus cmd�ͻ���
            Exit Function
        End If
    End If
    Check����Ʊ��Valid = True
End Function
