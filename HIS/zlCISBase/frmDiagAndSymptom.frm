VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiagAndSymptom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����벡�ֶ�Ӧ"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "frmDiagAndSymptom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   2640
      Picture         =   "frmDiagAndSymptom.frx":000C
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ�����(&C)"
      Height          =   350
      Left            =   1350
      Picture         =   "frmDiagAndSymptom.frx":0156
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1290
   End
   Begin VB.CommandButton cmdDiag 
      Caption         =   "&P"
      Height          =   300
      Left            =   8460
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   870
      Width           =   285
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      Picture         =   "frmDiagAndSymptom.frx":02A0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7665
      TabIndex        =   10
      Top             =   5340
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -405
      Top             =   6090
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
            Picture         =   "frmDiagAndSymptom.frx":03EA
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagAndSymptom.frx":0984
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagAndSymptom.frx":0F1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDiag 
      Height          =   300
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   2
      Top             =   885
      Width           =   7200
   End
   Begin VB.Frame fraӦ���� 
      Caption         =   "Ӧ����(&B)"
      Height          =   1050
      Left            =   120
      TabIndex        =   5
      Top             =   4155
      Width           =   8640
      Begin VB.ComboBox cbo���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   570
         Width           =   2100
      End
      Begin VB.OptionButton opt����Ӧ�÷�Χ 
         Caption         =   "��Ӧ���ڵ�ǰ���(&1)"
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton opt����Ӧ�÷�Χ 
         Caption         =   "Ӧ����                        ���������(&2)"
         Height          =   240
         Index           =   1
         Left            =   330
         TabIndex        =   8
         Top             =   630
         Width           =   5145
      End
      Begin VB.OptionButton opt����Ӧ�÷�Χ 
         Caption         =   "Ӧ��������������������(&3)"
         Height          =   240
         Index           =   2
         Left            =   5595
         TabIndex        =   7
         Top             =   300
         Width           =   2730
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vs���� 
      Height          =   2820
      Left            =   105
      TabIndex        =   4
      Top             =   1245
      Width           =   8670
      _cx             =   15293
      _cy             =   4974
      Appearance      =   1
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDiagAndSymptom.frx":2C28
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   6570
      TabIndex        =   9
      Top             =   5340
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   330
      Picture         =   "frmDiagAndSymptom.frx":2C78
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "ָ�����(&Z)"
      Height          =   180
      Left            =   195
      TabIndex        =   1
      Top             =   930
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��ѡ��ָ���ļ�����Ϻ����øü����������Ӧ�ı��ղ��֡�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   390
      Width           =   5400
   End
End
Attribute VB_Name = "frmDiagAndSymptom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln�༭ As Boolean
Private mlng���ID As Long                  '������ĿID
Private mlng����id As Long
Private mblnChange As Boolean
Private mbln��ҽ  As Boolean
Private Sub cmdClear_Click()
    Dim lngRow As Long
    With vs����
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, .ColIndex("����")) = ""
            .Cell(flexcpData, lngRow, .ColIndex("����")) = ""
        Next
    End With
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function MulitSelect���(strkey As String) As Boolean
    '------------------------------------------------------------------------------------------
    '����:��ѡ�����Ϣ
    '����:strKey-��������ֵ
    '����:Trueѡ���������Ϣ,����:ѡ��ʧ��!
    '����:���˺�
    '����:2007/06/17
    '------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSearchKey As String, strTittle As String, lngH As Long
    Dim vRect As RECT
    Dim rsTemp As New ADODB.Recordset
 
    On Error GoTo errHandle
    gstrSql = "" & _
        "   Select Distinct decode(a.���,1,'��ҽ','��ҽ') as ���, a.Id ,a.����,a.����,a.˵��,a.���� " & _
        "   From �������Ŀ¼ a,������ϱ��� b " & _
        "   Where a.id=b.���id and b.����=1  and a.���=" & IIf(mbln��ҽ, 2, 1) & " and (a.���� like [1] or a.���� like [1] Or b.���� like [2])" & _
        " and (a.����ʱ�� Is Null Or a.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) " & _
        "   Order by  ���,����"
    
    
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    
    
    vRect = zlControl.GetControlRect(txtDiag.hWnd)
    lngH = txtDiag.Height
    strSearchKey = gstrMatch & strkey & "%"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strSearchKey, CStr(UCase(strSearchKey)))
    If blnCancel = True Then
        If txtDiag.Enabled Then txtDiag.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "������ָ���ļ������,����!"
        If txtDiag.Enabled Then txtDiag.SetFocus
        Exit Function
    End If
    txtDiag.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
    txtDiag.Tag = Nvl(rsTemp!ID)
    
    
    Call init��������(Val(txtDiag.Tag), 0)
    Call Initȱʡ����(Val(txtDiag.Tag))
    MulitSelect��� = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmdDiag_Click()

    Dim rsTemp As New ADODB.Recordset
    Dim blnCancel As Boolean
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "" & _
    "   Select Id||'_' As Id ,�ϼ�id||decode(�ϼ�id,Null,Null,'_') As �ϼ�id,����,����,0 As ĩ��,'' ˵��,'' ���� " & _
    "   From ������Ϸ��� where ���=" & IIf(mbln��ҽ, 2, 1) & " Start With �ϼ�id Is Null Connect By Prior Id=�ϼ�id " & _
    "   Union All " & _
    "   Select A.Id|| '_'||b.����ID As Id,b.����id||'_' As �ϼ�id,a.����,a.����,1 As ĩ��,a.˵��,a.���� " & _
    "   From �������Ŀ¼ A,����������� b" & _
    "   Where a.ID = b.���ID and a.���=" & IIf(mbln��ҽ, 2, 1) & "" & _
    " and (a.����ʱ�� Is Null Or a.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd')) "
    
    '-------------------------------------------------------------------------------------------------------------------------------
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    '-------------------------------------------------------------------------------------------------------------------------------
    Set rsTemp = zlDatabase.ShowSelect(Me, gstrSql, 1, IIf(mbln��ҽ, "��ҽ", "��ҽ") & "�������", True, , "ѡ��ָ���ļ������", , , , , , , blnCancel, False)
    If rsTemp Is Nothing Then Exit Sub
    If blnCancel = True Then Exit Sub
    txtDiag.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
    txtDiag.Tag = Split(Nvl(rsTemp!ID), "_")(0)
    
    Call init��������(Val(txtDiag.Tag), Val(Split(Nvl(rsTemp!ID), "_")(1)))
    Call Initȱʡ����(Val(txtDiag.Tag))
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function Initȱʡ����(ByVal lng���ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------
    '����:�������,�����Ѿ����úõĶ�Ӧ����
    '����:lng���id-���id
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/17
    '---------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    Dim strSql As String, rsTemp As New ADODB.Recordset
    strSql = "Select a.����,a.����ID,c.����||'-'||c.���� as ���� From ��ϲ��ֶ�Ӧ a ,���ղ��� C Where  a.���id=[1] And a.����id=c.Id"
    
    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng���ID)
    With vs����
        For i = 1 To .Rows - 1
            rsTemp.Filter = "����=" & .RowData(i)
            If rsTemp.EOF = True Then
                .TextMatrix(i, .ColIndex("����")) = ""
                .Cell(flexcpData, i, .ColIndex("����")) = ""
            Else
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
                .Cell(flexcpData, i, .ColIndex("����")) = Nvl(rsTemp!����ID)
            End If
        Next
    End With
    rsTemp.Close
    Initȱʡ���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function init��������(ByVal lng���ID As Long, ByVal lng����id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------
    '����:����ָ�������,���������������combox�ؼ���
    '����:lng����id-����id
    '     lng���id-���id
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/08/17
    '---------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    strSql = " Select a.Id,a.����,a.���� From ������Ϸ��� a,����������� b Where a.Id=b.����id And b.���ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng���ID)
    With rsTemp
        cbo����.Clear
        Do While Not .EOF
            cbo����.AddItem Nvl(!����) & "-" & Nvl(!����)
            cbo����.ItemData(cbo����.NewIndex) = Val(Nvl(!ID))
            If Val(Nvl(!ID)) = lng����id Then
                cbo����.ListIndex = cbo����.NewIndex
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 And cbo����.ListIndex < 0 Then cbo����.ListIndex = 0
    End With
    rsTemp.Close
    init�������� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub Load�����Ϣ()
    '------------------------------------------------------------------------------
    '����:����ָ���������Ϣ
    '����:
    '����:
    '����:���˺�
    '����:2007/08/17
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If mlng���ID = 0 Then GoTo Init:
    On Error GoTo errHandle
    gstrSql = "Select id,����,���� From �������Ŀ¼ where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlng���ID)
    If rsTemp.EOF Then
        GoTo Init:
    End If
    txtDiag.Text = "[" & Nvl(rsTemp!����) & "]" & Nvl(rsTemp!����)
    txtDiag.Tag = Nvl(rsTemp!ID)
    
Init:
    Call init��������(Val(txtDiag.Tag), mlng����id)
    Call Initȱʡ����(Val(txtDiag.Tag))
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cmdRestore_Click()
    '�ָ���ص�����
    Call init��������(Val(txtDiag.Tag), mlng����id)
    Call Initȱʡ����(Val(txtDiag.Tag))
End Sub

Private Function IsValid() As Boolean
    '------------------------------------------------------------------------------------------
    '����:��������Ĳ����Ƿ���Ч
    '����:
    '����ֵ:��Ч����True,����ΪFalse
    '------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTemp As String
    
    '���
    With vs����
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("����"))) <> "" And Val(.Cell(flexcpData, i, .ColIndex("����"))) = 0 Then
                MsgBox "������Ĳ��ֲ���ȷ,����������!", vbInformation + vbDefaultButton1, gstrSysName
                .Row = i
                If .RowIsVisible(i) = False Then
                    .TopRow = i
                End If
                .SetFocus
                Exit Function
            End If
        Next
    End With
    If opt����Ӧ�÷�Χ(1).Value = True Then
        If cbo����.ListIndex < 0 Then
            MsgBox "δѡ��ָ���ķ���,����!", vbInformation + vbDefaultButton1, gstrSysName
            If cbo����.Enabled Then cbo����.SetFocus
            Exit Function
        End If
    End If
    If Val(txtDiag.Tag) = 0 Then
        MsgBox "δѡ��ָ���ķ���,����!", vbInformation + vbDefaultButton1, gstrSysName
        If txtDiag.Enabled Then txtDiag.SetFocus
        Exit Function
    End If
    IsValid = True
End Function


Private Sub cmdSave_Click()
    '����:��֤����벡�ֵĶ�Ӧ��ϵ
    Dim n As Long, str���� As String, int����Ӧ�÷�Χ As Integer
   
    If IsValid() = False Then Exit Sub
    
    For n = 0 To opt����Ӧ�÷�Χ.UBound
         If opt����Ӧ�÷�Χ(n).Value = True Then
             int����Ӧ�÷�Χ = n
             Exit For
         End If
    Next
    str���� = ""
    With vs����
        For n = 1 To .Rows - 1
            If .RowData(n) <> 0 Then
                If Val(.Cell(flexcpData, n, .ColIndex("����"))) <> 0 Then
                    str���� = str���� & "," & .RowData(n) & "|" & Val(.Cell(flexcpData, n, .ColIndex("����")))
                End If
            End If
        Next
    End With
    If str���� <> "" Then str���� = Mid(str����, 2)
    
    'Zl_��ϲ��ֶ�Ӧ_Update
    gstrSql = "Zl_��ϲ��ֶ�Ӧ_Update("
    '  ���ID_In     In �������Ŀ¼.ID%Type,
    gstrSql = gstrSql & "" & Val(txtDiag.Tag) & ","
    '  ����id_In In �����������.����id%Type := 0, --ָ���ķ���ID
    If int����Ӧ�÷�Χ = 1 Then
        gstrSql = gstrSql & "" & cbo����.ItemData(cbo����.ListIndex) & ","
    Else
        gstrSql = gstrSql & "" & 0 & ","
    End If
    '  ����_In   In Varchar2 := Null, --����id��,����1|����id1,����2|����id2.....
    gstrSql = gstrSql & "'" & str���� & "',"
    '  Ӧ��_In   In Number := 0 --���ֵ�Ӧ�÷�Χ:0-Ӧ���ڵ�ǰ��Ŀ;1-Ӧ����ָ������;2-Ӧ������������
    gstrSql = gstrSql & "" & int����Ӧ�÷�Χ & ")"
    
    Err = 0: On Error GoTo ErrHand:
    Me.Enabled = False
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Me.Enabled = True
    MsgBox "����ɹ�!", vbInformation + vbDefaultButton1, gstrSysName
    Exit Sub
ErrHand:
    Me.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If Init����() = False Then Unload Me: Exit Sub
    Call Load�����Ϣ
    Call CtlEnableSet
    mblnChange = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub opt����Ӧ�÷�Χ_Click(Index As Integer)
    cbo����.Enabled = opt����Ӧ�÷�Χ(1).Value
End Sub

Private Sub txtDiag_Change()
    txtDiag.Tag = ""
    mblnChange = True
End Sub

Private Sub txtDiag_GotFocus()
    zlControl.TxtSelAll txtDiag
End Sub

Private Sub txtDiag_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub

    
    strTemp = UCase(Trim(Me.txtDiag.Text))
    If strTemp = "" Then mlng���ID = 0: Me.txtDiag.Tag = "": Me.txtDiag.Text = "": Exit Sub

    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    If MulitSelect���(strTemp) = False Then
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Function Init����() As Boolean
    '--------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����
    '����:
    '����:���óɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2007/08/17
    '--------------------------------------------------------------------------------------------------------------------
    Dim rs���� As New ADODB.Recordset, i As Long
    Err = 0: On Error GoTo ErrHand:
    
    gstrSql = "Select ���,���� From ������� where ҽԺ���� is not null"
    Call zlDatabase.OpenRecordset(rs����, gstrSql, Me.Caption)
    If rs����.RecordCount = 0 Then
        MsgBox "δ��װ��ص�ҽ��,����ϵͳ����Ա!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With vs����
        If rs����.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .Cell(flexcpData, 1, i) = ""
                .RowData(1) = 0
            Next
            .Editable = flexEDNone
            Exit Function
        End If
        .Rows = rs����.RecordCount + 1
        i = 1
        Do While Not rs����.EOF
            .RowData(i) = Val(Nvl(rs����!���))
            .TextMatrix(i, .ColIndex("����")) = Nvl(rs����!����)
            .TextMatrix(i, .ColIndex("����")) = ""
            .Cell(flexcpData, i, .ColIndex("����")) = ""
            i = i + 1
             rs����.MoveNext
        Loop
        If mbln�༭ = False Then
            .Editable = flexEDKbdMouse
            .ColComboList(.ColIndex("����")) = ""
        Else
            .Editable = flexEDKbdMouse
            .ColComboList(.ColIndex("����")) = "..."
        End If
    End With
    Init���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowEdit(ByVal frmMain As Object, ByVal lng���ID As Long, ByVal lng����id As Long, ByVal bln�༭ As Boolean, ByVal bln��ҽ As Boolean)
    '-------------------------------------------------------------------------------------------------
    '����:��ʾ�༭����,�������
    '����:frmMain-������
    '     lng���id=���id
    '     lng����id=Ĭ�ϵķ���id(��ҪӦ����Ӧ����ָ��������)
    '     bln�༭=�Ƿ���Ա༭
    '����:���˺�
    '����:2007/08/17
    '-------------------------------------------------------------------------------------------------
    On Error Resume Next
    mblnFirst = True
    mlng���ID = lng���ID
    mlng����id = lng����id
    mbln�༭ = bln�༭
    mbln��ҽ = bln��ҽ
    
    Me.Show 1, frmMain
End Sub
Private Sub CtlEnableSet()
    '---------------------------------------------------------------------------------------------------------------------
    '����:������ؿؼ���Enable
    '����:
    '����:���˺�
    '����:2007/08/17
    '---------------------------------------------------------------------------------------------------------------------
    txtDiag.Enabled = mbln�༭
    cmdDiag.Enabled = mbln�༭
    vs����.Editable = flexEDKbdMouse
    If mbln�༭ Then
        vs����.Editable = flexEDKbdMouse
    Else
        vs����.Editable = flexEDNone
    End If
    cmdClear.Visible = mbln�༭
    cmdRestore.Visible = mbln�༭
    cmdSave.Visible = mbln�༭
    fraӦ����.Enabled = mbln�༭
    opt����Ӧ�÷�Χ(0).Enabled = mbln�༭
    opt����Ӧ�÷�Χ(1).Enabled = mbln�༭
    opt����Ӧ�÷�Χ(2).Enabled = mbln�༭
    cbo����.Enabled = mbln�༭
End Sub


Private Sub vs����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs����
        Select Case Col
        Case .ColIndex("����")
             .ColComboList(0) = "..."
        End Select
    End With
End Sub

Private Sub vs����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs����
        Select Case Col
        Case .ColIndex("����")
             Cancel = True
        End Select
    End With

End Sub

Private Sub vs����_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case vs����.ColIndex("����")
        'ѡ����
        Call Select����(vs����.RowData(Row), "")
    Case Else
    End Select
End Sub
Private Function Select����(ByVal lng���� As Long, ByVal strkey As String)
    '---------------------------------------------------------------------------------
    '����:ѡ��ָ������Ĳ���
    '����:lng����-����
    '����:ѡ��ɹ�,����ture,���򷵻�False
    '����:���˺�
    '����:2007/08/15
    '---------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strLeft As String
    Dim blnCancel As Boolean
    
    Dim vRect As RECT

    strLeft = gstrMatch
    
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Err = 0: On Error GoTo ErrHand:
    
    Dim sngX As Single, sngY As Single
    Call CalcPosition(sngX, sngY, vs����)
     
     If strkey <> "" Then
        strkey = strLeft & strkey & "%"
        gstrSql = "" & _
          "   Select Id, ����, ����, ����, decode('0','��ͨ��','1','���Բ�','2','���ֲ�','') As ���, ����ⶥ��, �ⶥ�߽�� " & _
          "    From ���ղ��� " & _
          "    Where ���� = [1] And (���� Like [2] Or ���� Like [2] Or ���� Like [3]) " & _
          "    Order by ����"
    Else
        gstrSql = "" & _
          "   Select Id, ����, ����, ����, decode('0','��ͨ��','1','���Բ�','2','���ֲ�','') As ���, ����ⶥ��, �ⶥ�߽�� " & _
          "    From ���ղ��� " & _
          "    Where ���� = [1]" & _
          "    Order by ����"
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "���ղ���ѡ��", False, "", "", False, False, True, sngX, sngY - vs����.CellHeight, vs����.CellHeight, blnCancel, False, False, lng����, strkey, CStr(UCase(strkey)))
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "������ָ���Ĳ���,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With vs����
        .TextMatrix(.Row, .ColIndex("����")) = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
        If strkey <> "" Then
            .EditText = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
        End If
        .Cell(flexcpData, .Row, .ColIndex("����")) = Nvl(rsTemp!ID)
    End With
    Select���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vs����_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vs����_DblClick()
   Select Case vs����.Col
    Case vs����.ColIndex("����")
        'ѡ����
        
        vs����.ColComboList(vs����.ColIndex("����")) = ""
        
        
    Case Else
    End Select
End Sub

Private Sub vs����_EnterCell()
    vs����.ColComboList(vs����.ColIndex("����")) = "..."
End Sub

Private Sub vs����_GotFocus()
    With vs����
        .BackColorSel = &H8000000D
'        .GridColor = &H0&
'        .GridColorFixed = &H0&
    End With
End Sub

Private Sub vs����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    If vs����.Col = vs����.ColIndex("����") And KeyCode <> vbKeyReturn Then
       vs����.ColComboList(vs����.ColIndex("����")) = ""
    End If
End Sub

Private Sub vs����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
      Dim strkey As String
        
        With vs����
        
            Select Case Col
            Case .ColIndex("����")
                strkey = Trim(vs����.EditText)
                strkey = Replace(strkey, Chr(vbKeyReturn), "")
                strkey = Replace(strkey, Chr(10), "")
                If strkey = "" Then Exit Sub
                .Cell(flexcpData, Row, Col) = ""
                If KeyCode <> vbKeyReturn Then Exit Sub
               If Select����(.RowData(.Row), strkey) = False Then
                    'ѡ��ʧ��
                    
                End If
                .ColComboList(.ColIndex("����")) = "..."
                .Col = 1
                .SetFocus
            End Select
        End With
End Sub

Private Sub vs����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vs����_LostFocus()
    With vs����
        .BackColorSel = &H8000000C
'        .GridColor = &H808080
'        .GridColorFixed = &H808080
    End With
End Sub
Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        x = objPoint.x * 15 'objBill.Left +
        y = objPoint.y * 15 + objBill.Height '+ objBill.Top
    Else
        x = objPoint.x * 15 + objBill.CellLeft
        y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub



