VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMicrobeEdit 
   BorderStyle     =   0  'None
   Caption         =   "ϸ����Ϣ"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cbo�걾���� 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":0000
      Left            =   1035
      List            =   "frmMicrobeEdit.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2925
      Width           =   1620
   End
   Begin VB.ComboBox cboEdit 
      Height          =   300
      Index           =   1
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Tag             =   "������Ⱦɫ����"
      Top             =   2100
      Width           =   1425
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&P"
      Height          =   285
      Index           =   0
      Left            =   4980
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   17
      Top             =   2505
      Width           =   4230
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   26
      Top             =   2925
      Width           =   1485
   End
   Begin VB.ComboBox cboĬ�Ϸ��� 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":0004
      Left            =   1215
      List            =   "frmMicrobeEdit.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3330
      Width           =   1440
   End
   Begin VB.TextBox txtWHONET�� 
      Height          =   300
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1695
      Width           =   1545
   End
   Begin VB.ComboBox cboĬ��ҩ�� 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":0008
      Left            =   4095
      List            =   "frmMicrobeEdit.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3330
      Width           =   1155
   End
   Begin VB.TextBox txt��д 
      Height          =   300
      Left            =   1020
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1695
      Width           =   1425
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1020
      MaxLength       =   10
      TabIndex        =   3
      Top             =   495
      Width           =   1425
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   5
      Top             =   885
      Width           =   4245
   End
   Begin VB.TextBox txtӢ�� 
      Height          =   300
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1290
      Width           =   4245
   End
   Begin VB.ComboBox cboϸ������ 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":000C
      Left            =   1020
      List            =   "frmMicrobeEdit.frx":0013
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   4290
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1800
      Left            =   105
      TabIndex        =   24
      Top             =   3975
      Width           =   5085
      _cx             =   8969
      _cy             =   3175
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.ComboBox cboEdit 
      Height          =   300
      Index           =   0
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "ϸ�����"
      Top             =   2100
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�걾����"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   2985
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ⱦɫ"
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   2580
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϸ�����"
      Height          =   180
      Index           =   0
      Left            =   2910
      TabIndex        =   14
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ�Ͻ��"
      Height          =   180
      Left            =   2910
      TabIndex        =   25
      Top             =   2985
      Width           =   720
   End
   Begin VB.Label lblĬ�Ϸ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ�ϼ�ⷽ��"
      Height          =   180
      Left            =   90
      TabIndex        =   19
      Top             =   3390
      Width           =   1080
   End
   Begin VB.Label lblWHONET�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WHONET��"
      Height          =   180
      Left            =   2910
      TabIndex        =   10
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ӧ����ҩ��ʵ��Ŀ�������:"
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   2430
   End
   Begin VB.Label lblĬ��ҩ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ��ҩ�����"
      Height          =   180
      Left            =   2910
      TabIndex        =   21
      Top             =   3390
      Width           =   1080
   End
   Begin VB.Label lbl��д 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ����д"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ϸ������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   570
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblӢ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӣ������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lblϸ������ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ����"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   165
      Width           =   720
   End
End
Attribute VB_Name = "frmMicrobeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngGermId As Long          '��ǰ��ʾ������id

Private Enum mCol
    ID = 0: ѡ��: ����: ����: ��ע
End Enum

'���˺�:����
Private Enum mcboIndex
    idx_ϸ����� = 0
    idx_������Ⱦɫ = 1
End Enum
Private Enum mTxtIndex
    idx_�������� = 0
End Enum

Dim lngCount As Long
Private Function SelectItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strkey As String, ByVal strTable As String, ByVal strTittle As String) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single, intƥ�䷽ʽ As Integer
    
    Dim rsTemp  As ADODB.Recordset
    intƥ�䷽ʽ = Val(zlDatabase.GetPara("����ƥ��", , , True))
    
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
    
    gstrSql = "Select rownum as ID,a.* From " & strTable & " a"
    
    If strkey <> "" Then
        gstrSql = gstrSql & _
        "   Where ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  " & _
        "    "
    End If
    gstrSql = gstrSql & _
    "   order by ����"
    
    strkey = IIf(intƥ�䷽ʽ = 0, "%", "") & strkey & "%"
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSql, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strkey)
    frmMain.SetFocus
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        MsgBox "û���ҵ���������������,����!", vbDefaultButton1 + vbInformation, gstrSysName
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If

    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
            .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!����)
        End With
    Else
        If objCtl.Enabled Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!����)
        objCtl.Tag = Nvl(rsTemp!����)
        zlCommFun.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 5: .FixedCols = 0
            .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.����) = "����"
            .TextMatrix(0, mCol.����) = "����": .TextMatrix(0, mCol.��ע) = "��ע"
            .ColWidth(mCol.����) = 900: .ColWidth(mCol.����) = 2600: .ColWidth(mCol.��ע) = 1000
        End If
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.ѡ��) = 280: .TextMatrix(0, mCol.ѡ��) = ""
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.ѡ��) = 1 Then
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexUnchecked
            End If
            .TextMatrix(lngCount, mCol.ѡ��) = ""
        Next
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub InitData()
    '------------------------------------------------------------------------------
    '����:��ʼ����Ӧcombox���ݣ���ֵ
    '����:
    '����:���˺�
    '����:2008/03/18
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select ����,����,����,ȱʡ��־ From ����ϸ����� order by ����"
    zlDatabase.OpenRecordset rsTemp, gstrSql, Me.Caption
    With cboEdit(mcboIndex.idx_ϸ�����)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!����) & "." & Nvl(rsTemp!����)
            If Val(Nvl(rsTemp!ȱʡ��־)) = 1 Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
    End With
    gstrSql = "Select ����,����,����,ȱʡ��־ From ����Ⱦɫ���� order by ����"
    zlDatabase.OpenRecordset rsTemp, gstrSql, Me.Caption
    With cboEdit(mcboIndex.idx_������Ⱦɫ)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!����) & "." & Nvl(rsTemp!����)
            If Val(Nvl(rsTemp!ȱʡ��־)) = 1 Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function zlRefresh(lngGermId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    mlngGermId = lngGermId
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, ����, �������� From ����ϸ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cboϸ������.Clear
        Do While Not .EOF
            Me.cboϸ������.AddItem !���� & "-" & !��������
            Me.cboϸ������.ItemData(Me.cboϸ������.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    '�����ǰ��Ŀ����ʾ
    
    Me.txt����.Text = "": Me.txt����.Text = "": Me.txtӢ��.Text = "": Me.txt��д.Text = ""
    Me.cboϸ������.ListIndex = -1: Me.cboĬ�Ϸ���.ListIndex = -1: Me.cboĬ��ҩ��.ListIndex = -1: Me.cbo�걾����.ListIndex = -1
    '���˺����:2008/03/17
    Me.cboEdit(mcboIndex.idx_������Ⱦɫ).ListIndex = -1: Me.cboEdit(mcboIndex.idx_ϸ�����).ListIndex = -1
    Me.txtEdit(mTxtIndex.idx_��������).Text = ""
    If lngGermId = 0 Then setListFormat: zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    gstrSql = "Select ����, ������, Ӣ����, ����, ����id, Ĭ��ҩ��, Ĭ�Ϸ���, Whonet��,Ĭ�Ͻ��, ϸ�����,ϸ������ ,�����Ϸ���" & vbCrLf & _
              " From ����ϸ�� Where ID = [1]  order by ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngGermId)
    With rsTemp
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("������").DefinedSize
        Me.txtӢ��.MaxLength = .Fields("Ӣ����").DefinedSize
        Me.txt��д.MaxLength = .Fields("����").DefinedSize
        Me.txtWHONET��.MaxLength = .Fields("WHONET��").DefinedSize
        Me.txt���.MaxLength = .Fields("Ĭ�Ͻ��").DefinedSize
        If .RecordCount > 0 Then
            Me.txt����.Text = "" & !����
            Me.txt����.Text = "" & !������
            Me.txtӢ��.Text = "" & !Ӣ����
            Me.txt��д.Text = "" & !����
            Me.txtWHONET��.Text = "" & !WHONET��
            Me.txt���.Tag = "" & !Ĭ�Ͻ��
            
            For lngCount = 0 To Me.cbo�걾����.ListCount - 1
                If InStr(Me.txt���.Tag, Mid(Me.cbo�걾����.List(lngCount), InStr(Me.cbo�걾����.List(lngCount), "-") + 1)) > 0 Then
                    Me.cbo�걾����.ListIndex = lngCount
                    Exit For
                End If
            Next
            If Me.cbo�걾����.ListIndex = -1 And Me.cbo�걾����.ListCount > 0 Then Me.cbo�걾����.ListIndex = 0
            
            For lngCount = 0 To Me.cboϸ������.ListCount - 1
                If Me.cboϸ������.ItemData(lngCount) = Val("" & !����id) Then Me.cboϸ������.ListIndex = lngCount: Exit For
            Next
            
            For lngCount = 0 To Me.cboĬ��ҩ��.ListCount - 1
                If Left(Me.cboĬ��ҩ��.List(lngCount), 1) = "" & !Ĭ��ҩ�� Then Me.cboĬ��ҩ��.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cboĬ�Ϸ���.ListCount - 1
                If Mid(Me.cboĬ�Ϸ���.List(lngCount), 4) = "" & !Ĭ�Ϸ��� Then Me.cboĬ�Ϸ���.ListIndex = lngCount: Exit For
            Next
            '���˺�:2008/03/18����
            Me.txtEdit(mTxtIndex.idx_��������).Text = Nvl(rsTemp!ϸ������)
            Me.txtEdit(mTxtIndex.idx_��������).Tag = Nvl(rsTemp!ϸ������)
            With Me.cboEdit(mcboIndex.idx_������Ⱦɫ)
                For lngCount = 0 To .ListCount - 1
                    strTemp = .List(lngCount): strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
                    
                    If strTemp = Nvl(rsTemp!�����Ϸ���) Then .ListIndex = lngCount: Exit For
                Next
                If .ListIndex < 0 And Trim(Nvl(rsTemp!�����Ϸ���)) <> "" Then
                        .AddItem Nvl(rsTemp!�����Ϸ���)
                        .ListIndex = .NewIndex
                End If
            End With
            With Me.cboEdit(mcboIndex.idx_ϸ�����)
                For lngCount = 0 To .ListCount - 1
                    strTemp = .List(lngCount): strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
                    
                    If strTemp = Nvl(rsTemp!ϸ�����) Then .ListIndex = lngCount: Exit For
                Next
                If .ListIndex < 0 And Trim(Nvl(rsTemp!ϸ�����)) <> "" Then
                        .AddItem Nvl(rsTemp!ϸ�����)
                        .ListIndex = .NewIndex
                End If
            End With
            
        End If
    End With
    
    gstrSql = "Select I.ID, 1 As ѡ��, I.����, I.����, Decode(D.ȱʡ��־, 1, '��Ĭ��������', '') As ��ע" & vbNewLine & _
            "From ����ϸ�������� D, ���鿹������ I" & vbNewLine & _
            "Where D.�����ط���id = I.ID And ϸ��id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngGermId)
    Set Me.vfgList.DataSource = rsTemp
    Call setListFormat(True)
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngGermId As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngGermId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    Dim str���� As String
    Dim intLoop As Integer
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(����), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From ����ϸ��"

'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
'            Call SQLTest
        With rsTemp
            If !���� <> 0 And !���� <= Me.txt����.MaxLength Then
                If Val(!����) = 0 Then
                    For intLoop = 1 To Len(!����)
                        If IsNumeric(Mid(!����, intLoop, 1)) Then
                            str���� = str���� & Format((Val(Mid(!����, intLoop)) + 1), String(Len(Mid(!����, intLoop)), "0"))
                            Exit For
                        Else
                            str���� = str���� & Mid(!����, intLoop, 1)
                        End If
                    Next
                    Me.txt����.Text = str����
                Else
                    Me.txt����.Text = Format(Val(!����) + 1, String(!����, "0"))
                End If
            Else
                Me.txt����.Text = Format(Val(!����) + 1, String(Me.txt����.MaxLength, "0"))
            End If
        End With
        
        '��������ñ�עֵ
        Me.txt����.Text = "": Me.txtӢ��.Text = "": Me.txt��д.Text = "": Me.txt���.Tag = ""
    End If
    If Me.cboϸ������.ListIndex = -1 And Me.cboϸ������.ListCount > 0 Then Me.cboϸ������.ListIndex = 0
    If Me.cboĬ��ҩ��.ListIndex = -1 And Me.cboĬ��ҩ��.ListCount > 0 Then Me.cboĬ��ҩ��.ListIndex = 0
    If Me.cboĬ�Ϸ���.ListIndex = -1 And Me.cboĬ�Ϸ���.ListCount > 0 Then Me.cboĬ�Ϸ���.ListIndex = 0
    
    gstrSql = "Select I.ID, Decode(D.ϸ��id, Null, 0, 1) As ѡ��, I.����, I.����, Decode(D.ȱʡ��־, 1, '��Ĭ��������', '') As ��ע" & vbNewLine & _
            "From (Select * From ����ϸ�������� Where ϸ��id = [1]) D, ���鿹������ I" & vbNewLine & _
            "Where D.�����ط���id(+) = I.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngGermId)
    Set Me.vfgList.DataSource = rsTemp
    Call setListFormat(True)
    
    mlngGermId = lngGermId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "����", "�޸�"): Call Form_Resize
    Me.txt����.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngGermId)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long
    Dim strLists As String
    
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked Then
                strLists = strLists & "|" & .TextMatrix(lngCount, mCol.ID) & _
                    ";" & IIf(Trim(.TextMatrix(lngCount, mCol.��ע)) <> "", 1, 0)
            End If
        Next
    End With
'    If strLists = "" Then MsgBox "û��ѡ��ҩ�����鿹���أ�", vbInformation, gstrSysName: zlEditSave = False: Exit Function
    strLists = Mid(strLists, 2)
    
    'һ�����Լ��
    If Me.cboϸ������.ListIndex = -1 Then
        MsgBox "��ѡ��ǰϸ�����ͣ�", vbInformation, gstrSysName
        Me.cboϸ������.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt����.Text) > Val(String(Me.txt����.MaxLength, "9")) Then
        MsgBox "����̫��", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "�������������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "�������Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtӢ��.Text), vbFromUnicode)) > Me.txtӢ��.MaxLength Then
        MsgBox "Ӣ�����Ƴ��������" & Me.txtӢ��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txtӢ��.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt��д.Text), vbFromUnicode)) > Me.txt��д.MaxLength Then
        MsgBox "��д���������" & Me.txt��д.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt��д.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtWHONET��.Text), vbFromUnicode)) > Me.txtWHONET��.MaxLength Then
        MsgBox "WHONET�볬�������" & Me.txtWHONET��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txtWHONET��.SetFocus: zlEditSave = 0: Exit Function
    End If
    '���˺�2008/03/17����
    If txtEdit(mTxtIndex.idx_��������).Tag = "" And Trim(txtEdit(mTxtIndex.idx_��������).Text) <> "" Then
        MsgBox "��������ѡ�����,���飡", vbInformation, gstrSysName
        Me.txtEdit(mTxtIndex.idx_��������).SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '�����������
    If zlExistItem("����ϸ������", "ID", Val(Me.cboϸ������.ItemData(Me.cboϸ������.ListIndex)), _
                                   Me.cboϸ������.List(Me.cboϸ������.ListIndex)) = False Then
        Me.zlRefresh (mlngGermId)
        zlEditSave = 0: Exit Function
    End If
    
    If zlExistItem("ϸ����ⷽ��", "����", Mid(Me.cboĬ�Ϸ���.Text, 4), Me.cboĬ�Ϸ���.Text) = False Then
        Me.zlRefresh (mlngGermId)
        zlEditSave = 0: Exit Function
    End If
    
    
    '���ݱ��������֯
    If Me.Tag = "����" Then
        lngNewId = zlDatabase.GetNextId("����ϸ��")
    Else
        If zlExistItem("����ϸ��", "ID", mlngGermId, Trim(Me.txt����.Text)) = False Then
            zlEditSave = 0: Exit Function
        End If
        lngNewId = mlngGermId
        
    End If
    gstrSql = "'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "','" & Trim(Me.txtӢ��.Text) & "','" & Trim(Me.txt��д.Text) & "'"
    gstrSql = gstrSql & "," & Me.cboϸ������.ItemData(Me.cboϸ������.ListIndex) & ",'" & Left(Me.cboĬ��ҩ��.Text, 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cboĬ�Ϸ���.Text, 4) & "','" & Trim(Me.txtWHONET��.Text) & "'"
    
    '���˺����:
    Dim strTemp As String
    '  ϸ�����_In   In ����ϸ��.ϸ�����%Type,
    strTemp = Me.cboEdit(mcboIndex.idx_ϸ�����).Text: strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
    gstrSql = gstrSql & ",'" & strTemp & "'"
    '  ϸ������_In   In ����ϸ��.ϸ������%Type,
    gstrSql = gstrSql & ",'" & Trim(txtEdit(mTxtIndex.idx_��������).Tag) & "'"
    '  �����Ϸ���_In In ����ϸ��.�����Ϸ���%Type,
    strTemp = Me.cboEdit(mcboIndex.idx_������Ⱦɫ).Text: strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
    gstrSql = gstrSql & ",'" & strTemp & "'"
            
    '�������ļ�¼
    Call txt���_LostFocus
    
    If Me.Tag = "����" Then
        gstrSql = "Zl_����ϸ��_Edit(1," & lngNewId & "," & gstrSql & ",'" & strLists & "','" & Me.txt���.Tag & "')"
    Else
        gstrSql = "Zl_����ϸ��_Edit(2," & lngNewId & "," & gstrSql & ",'" & strLists & "','" & Me.txt���.Tag & "')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mlngGermId = lngNewId
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    zlEditSave = mlngGermId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Sub cboEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    If KeyCode = vbKeyDelete Then cboEdit(Index).ListIndex = -1
End Sub

Private Sub cbo�걾����_Click()
    Dim intLoop As Integer
    Dim strItem() As String
    Dim strTmp As String
    
    Me.txt���.Text = ""
    strTmp = Mid(cbo�걾����, InStr(cbo�걾����, "-") + 1)
    
    strItem = Split(Me.txt���.Tag, ";")
    
    For intLoop = 1 To UBound(strItem)
        If Split(strItem(intLoop), ",")(0) = strTmp Then
            Me.txt���.Text = Split(strItem(intLoop), ",")(1)
        End If
    Next
    
    If Me.txt���.Tag <> "" And InStr(Me.txt���.Tag, ",") = 0 And InStr(Me.txt���.Tag, ";") = 0 Then
        Me.txt���.Text = Me.txt���.Tag
    End If
End Sub

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
Private Sub cboĬ�Ϸ���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboĬ��ҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboϸ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
 

Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
    Case mTxtIndex.idx_��������
        If SelectItem(Me, txtEdit(mTxtIndex.idx_��������), "", "����ϸ������", "����ϸ������ѡ����") = False Then Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mlngGermId = 0
    
    With Me.cboĬ��ҩ��
        .AddItem "R-��ҩ": .AddItem "I-�н�": .AddItem "S-����"
    End With
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ����, ����, ���� From ϸ����ⷽ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cboĬ�Ϸ���.Clear
        Do While Not .EOF
            Me.cboĬ�Ϸ���.AddItem !���� & "-" & !����
            .MoveNext
        Loop
    End With
    
    gstrSql = "SELECT ����,���� FROM ���Ƽ���걾 order by ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo�걾����.Clear
        Do While Not .EOF
            Me.cbo�걾����.AddItem !���� & "-" & !����
            .MoveNext
        Loop
        If Me.cbo�걾����.ListCount > 0 Then
            Me.cbo�걾����.ListIndex = 1
        End If
    End With
    
    '------------------------------------------------------
    '���˺�:2008/03/18����
    Call InitData
    '------------------------------------------------------
    Call setListFormat
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 180
    If Me.Tag <> "" Then
        Me.BackColor = RGB(250, 250, 250)
        Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.BackColor = &H8000000F
        Me.vfgList.FocusRect = flexFocusNone
    End If
End Sub
 

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
    
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlCommFun.OpenIme (True)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mTxtIndex.idx_��������
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If SelectItem(Me, txtEdit(Index), DelInvalidChar(Trim(txtEdit(Index).Text)), "����ϸ������", "����ϸ������ѡ����") = False Then Exit Sub
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub txtWHONET��_GotFocus()
    Me.txtWHONET��.SelStart = 0: Me.txtWHONET��.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtWHONET��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���_LostFocus()
    Dim strTmp As String
    Dim intLoop As Integer
    Dim strItem() As String
    strTmp = Mid(cbo�걾����, InStr(cbo�걾����, "-") + 1)
    If Trim(txt���.Text) <> "" Then
        '����
        If InStr(txt���.Tag, ";" & strTmp & ",") > 0 Then
            strItem = Split(txt���.Tag, ";")
            Me.txt���.Tag = ""
            For intLoop = 1 To UBound(strItem)
                If Split(strItem(intLoop), ",")(0) = strTmp Then
                    Me.txt���.Tag = Me.txt���.Tag & ";" & strTmp & "," & Me.txt���.Text
                Else
                    Me.txt���.Tag = Me.txt���.Tag & ";" & strItem(intLoop)
                End If
            Next
        Else
            Me.txt���.Tag = Me.txt���.Tag & ";" & strTmp & "," & Me.txt���
        End If
    Else
        'ɾ��
        If InStr(txt���.Tag, ";" & strTmp & ",") > 0 Then
            strItem = Split(txt���.Tag, ";")
            Me.txt���.Tag = ""
            For intLoop = 1 To UBound(strItem)
                If Split(strItem(intLoop), ",")(0) = strTmp Then
                    Me.txt���.Tag = Me.txt���.Tag & ""
                Else
                    Me.txt���.Tag = Me.txt���.Tag & ";" & strItem(intLoop)
                End If
            Next
        End If
    End If
End Sub

Private Sub txt��д_GotFocus()
    Me.txt��д.SelStart = 0: Me.txt��д.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��д_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtӢ��_GotFocus()
    Me.txtӢ��.SelStart = 0: Me.txtӢ��.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtӢ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    With Me.vfgList
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Col = mCol.��ע Then
            If .Cell(flexcpChecked, .Row, mCol.ѡ��) <> flexChecked Then Exit Sub
            For lngCount = .FixedRows To .Rows - 1
                .TextMatrix(lngCount, mCol.��ע) = IIf(lngCount = .Row, "��Ĭ��������", "")
            Next
        Else
            If .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexChecked Then
                .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked
                .TextMatrix(.Row, mCol.��ע) = ""
            Else
                .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexChecked
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfgList_DblClick
End Sub

