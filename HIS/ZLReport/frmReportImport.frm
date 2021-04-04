VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReportImport 
   Caption         =   "������"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmReportImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8445
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCopyTypeSet 
      Caption         =   "����"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   4950
      Width           =   615
   End
   Begin VB.ComboBox cboCopyType 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdImportTypeSet 
      Caption         =   "����"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   4590
      Width           =   615
   End
   Begin VB.ComboBox cboImportType 
      Height          =   300
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4920
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _cx             =   14843
      _cy             =   7858
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�������ø��Ƿ�ʽ"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   4980
      Width           =   1440
   End
   Begin VB.Label lblImportType 
      AutoSize        =   -1  'True
      Caption         =   "�������õ��뷽ʽ"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   4620
      Width           =   1440
   End
End
Attribute VB_Name = "frmReportImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsReports As ADODB.Recordset
Private mrsFiles As ADODB.Recordset
Private mlngSys As Long
Private mlngGroup As Long
Private mblnAllImp As Boolean
Private mblnOK As Boolean

Public Function ShowMe(ByVal lngSys As Long, ByVal lngGroup As Long, ByVal blnAllImp As Boolean, ByVal rsReports As ADODB.Recordset, ByRef rsFiles As ADODB.Recordset, ByVal objParent As Object) As Boolean
    Set mrsReports = rsReports
    Set mrsFiles = rsFiles
    mlngSys = lngSys
    mlngGroup = lngGroup
    mblnAllImp = blnAllImp
    Call InitData
    Me.Show 1, objParent
    Set rsFiles = mrsFiles
    ShowMe = mblnOK
End Function


Public Function InitData() As Boolean
    Dim arrTmp As Variant, strInfo As String
    Dim strFilter As String
    Dim intErrType As Integer, intImpType As Integer, lngImpGroup As Long, lngRPTID As Long
    Dim strMsg As String, strOption As String, strReturn As String
    Dim i As Long, lngCount As Long
    Dim cllCover As New Collection '�����ǵı���ID,�����ų�һ�ε���һ�������θ���
    Dim blnSingle  As Boolean, strFileName As String
    Dim strFlag As String
    
    With cboImportType
        .Clear
        .AddItem "��������"
        .AddItem "���ǵ���"
        .ListIndex = 0
    End With
    
    With cboCopyType
        .Clear
        .AddItem "���帲��"
        .AddItem "����Դ����"
        .ListIndex = 0
    End With
    '��ʼ�����
    With vsf
        .Cols = 7
        .Rows = 1
        .TextMatrix(0, 0) = "������"
        .ColKey(0) = "������"
        .ColDataType(0) = flexDTString
        .ColWidth(0) = 1200
        
        .TextMatrix(0, 1) = "��������"
        .ColKey(1) = "��������"
        .ColDataType(1) = flexDTString
        .ColWidth(0) = 1200
        
        .TextMatrix(0, 2) = "��������"
        .ColKey(2) = "��������"
        .Editable = flexEDKbdMouse
        .ColComboList(2) = "��������|���ǵ���"
        .ColWidth(2) = 1200
        
        .TextMatrix(0, 3) = "��������"
        .ColKey(3) = "��������"
        .Editable = flexEDKbdMouse
        .ColComboList(3) = Space$(1) & "|���帲��|����Դ����"
        
        .TextMatrix(0, 4) = "˵��"
        .ColKey(4) = "˵��"
        .ColDataType(4) = flexDTString
        .ColWidth(4) = 1800
        
        .TextMatrix(0, 5) = "�����ʶ"
        .ColKey(5) = "�����ʶ"
        .ColDataType(5) = flexDTLong
        .ColWidth(5) = 0
        
        .TextMatrix(0, 6) = "�ļ���"
        .ColKey(6) = "�ļ���"
        .ColDataType(6) = flexDTString
        .ColWidth(6) = 0
    End With
    
    With mrsFiles
        .Filter = "": .Sort = "FilePath Desc"
        blnSingle = mrsFiles.RecordCount = 1 '�Ƿ񵥸�������
'        If blnSingle Then strFileName = mrsFiles!FileName
        Do While Not .EOF
            intErrType = 0: intImpType = 0: lngImpGroup = 0: lngRPTID = 0
            arrTmp = Split(GetReportInfo(!FilePath & ""), ";") '��ȡ�ļ���Ϣ
            If UBound(arrTmp) <> 2 Then intErrType = 4 '�ļ����
            If Val(arrTmp(2)) <> 9 And intErrType = 0 Then intErrType = 5  '�汾���
            strFileName = arrTmp(1)
            If intErrType = 0 Then
                If mlngSys = 0 Then '��ϵͳ����Ҫ�����ı����в��ܴ�����ͬ����
                    '�ǹ̶�����ȫ�������Ѿ�ȷ������Ҫ����ķ���
                    mrsReports.Filter = "����='" & arrTmp(1) & "' And ���='" & arrTmp(0) & "'" & IIF(mlngSys = 0 And mblnAllImp, " And ��ID=" & !��ID, "")
                    If mrsReports.EOF Then mrsReports.Filter = "����='" & arrTmp(1) & "'" & IIF(mlngSys = 0 And mblnAllImp, " And ��ID=" & !��ID, "")
                Else 'ϵͳ����ͨ�����ֱ�Ӳ���
                    mrsReports.Filter = "���='" & arrTmp(0) & "'"
                End If
                'ȷ��������ķ��飬������ڵ�ͬ���ģ����Ȳ���û�з���ı���
                mrsReports.Sort = "ID Desc,��ID"
                If Not mrsReports.EOF Then
                    lngRPTID = mrsReports!id: lngImpGroup = mrsReports!��ID
                    If lngRPTID = 0 Then intErrType = 1 '�ñ����Ѿ����������
                    If intErrType = 0 Then
                        On Error Resume Next
                        cllCover.Add "1", "_" & lngRPTID
                        If Err.Number <> 0 Then Err.Clear: intErrType = 2 '�ñ����Ѿ�����Ǹ���
                        On Error GoTo errH
                    End If
                    If intErrType = 0 Then intImpType = 2
                    '������Ʋ�ƥ��
                    If intErrType = 0 And (CStr(arrTmp(0)) <> mrsReports!��� & "" Or CStr(arrTmp(1)) <> mrsReports!����) Then intErrType = 6
                Else
                    If mlngSys <> 0 Then intErrType = 3  'ϵͳ�̶�������븲��ͬ������
                    If intErrType = 0 Then intImpType = 1  '��ϵͳ����û��ͬ��������������
                End If
                If mlngSys = 0 And mblnAllImp Then lngImpGroup = !��ID '�ǹ̶�������ȡԭ���ķ���
                '�ñ�����������������뻺�棬��ֹ�������
                If mrsReports.EOF And mlngSys = 0 Then mrsReports.AddNew Array("Id", "���", "����", "��iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), lngImpGroup)
            End If
            
            vsf.Rows = vsf.Rows + 1
            vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("������")) = arrTmp(0)
            vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("��������")) = arrTmp(1)
            vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("�ļ���")) = !FileName
            Select Case intErrType
            Case 2
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("˵��")) = "����""" & strFileName & """������ͬ���ݵı����޷����룡"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("�����ʶ")) = 1
            Case 3
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("˵��")) = "����""" & strFileName & """����û�п��Ը��ǵı�����޷����룡"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("�����ʶ")) = 1
            Case 4
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("˵��")) = "����""" & strFileName & """�������ݴ���������޷����룡"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("�����ʶ")) = 1
            Case 5
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("˵��")) = "����""" & strFileName & """���ڰ汾���Զ��޷����룡"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("�����ʶ")) = 1
            Case 6
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("˵��")) = "����""" & strFileName & """��Ż������븲�ǵı����������ѡ��ȷ�ϣ�"
                vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("�����ʶ")) = 1
            Case Else
                Select Case intImpType
                Case 1
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("˵��")) = "���������뱨��" & strFileName
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("��������")) = "��������"
                Case 2
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("˵��")) = "����" & strFileName & "���Ḳ��ԭ�б���"
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("��������")) = "���ǵ���"
                    vsf.TextMatrix(vsf.Rows - 1, vsf.ColIndex("��������")) = "���帲��"
                End Select
            End Select

            .Update Array("��ID", "ͬ��ID", "��������", "ErrType"), Array(lngImpGroup, lngRPTID, intImpType, intErrType)
            
            .MoveNext
        Loop
    End With
    mrsFiles.Filter = ""
errH:
End Function

Private Sub cmdCopyTypeSet_Click()
    Dim i As Integer
    Dim strCopyType As String
    Dim strFileName As String
    
    With vsf
        strCopyType = cboCopyType.Text
        mrsFiles.Filter = ""
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("��������")) = "���ǵ���" Then
                '�ж��Ƿ�������޷�����,�ж��Ƿ���ID����(�����Ƿ��鱨���������������)
                If Val(.TextMatrix(i, .ColIndex("�����ʶ"))) = 0 Then
                    .TextMatrix(i, .ColIndex("��������")) = strCopyType
                End If
            End If
        Next
    End With
End Sub

Private Sub cmdImportTypeSet_Click()
    Dim i As Integer
    Dim strImportType As String
    Dim strFileName As String
    
    With vsf
        strImportType = cboImportType.Text
        mrsFiles.Filter = ""
        For i = 1 To .Rows - 1
            '�ж��Ƿ���鱨��,���Ƿ��鱨�����������ǵ���Ϊ��������
            mrsFiles.Filter = "FileName='" & .TextMatrix(i, .ColIndex("�ļ���")) & "'"
            Select Case strImportType
            Case "��������"
                '�ж��Ƿ�������޷�����,�ж��Ƿ���ID����(�����Ƿ��鱨���������������)
                If Val(.TextMatrix(i, .ColIndex("�����ʶ"))) = 0 And Val(mrsFiles("��ID")) = 0 Then
                    .TextMatrix(i, .ColIndex("��������")) = "��������"
                    .TextMatrix(i, .ColIndex("��������")) = Space$(1)
                    .TextMatrix(i, .ColIndex("˵��")) = "���������뱨��" & .TextMatrix(i, .ColIndex("��������"))
                End If
            Case "���ǵ���"
                '�ж��Ƿ�������޷�����,�ж��Ƿ���ID����(�����Ƿ��鱨���������������)
                If Val(.TextMatrix(i, .ColIndex("�����ʶ"))) = 0 And Val(mrsFiles("ͬ��ID")) > 0 Then
                    .TextMatrix(i, .ColIndex("��������")) = "���ǵ���"
                    If .TextMatrix(i, .ColIndex("��������")) = Space$(1) Then
                        .TextMatrix(i, .ColIndex("��������")) = "���帲��"
                    End If
                    .TextMatrix(i, .ColIndex("˵��")) = "����" & .TextMatrix(i, .ColIndex("��������")) & "���Ḳ��ԭ�б���"
                End If
            End Select
        Next
    End With
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim intImpType As Integer
    Dim intCopyType As Integer
    Dim intSameID As Integer
    With mrsFiles
        .Filter = ""
        For i = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(i, vsf.ColIndex("�����ʶ"))) = 0 Then
                intImpType = IIF(vsf.TextMatrix(i, vsf.ColIndex("��������")) = "��������", 1, 2)
                
                intCopyType = IIF(vsf.TextMatrix(i, vsf.ColIndex("��������")) = "����Դ����", 1, 0)
                mrsFiles.Filter = "FileName='" & vsf.TextMatrix(i, vsf.ColIndex("�ļ���")) & "'"
                intSameID = Val(mrsFiles("ͬ��ID").Value)
                If intImpType = 1 Then
                    intSameID = 0
                End If
                mrsFiles.Update Array("��������", "ͬ��ID", "��������"), Array(intImpType, intSameID, intCopyType)
            Else
                mrsFiles.Filter = "FileName='" & vsf.TextMatrix(i, vsf.ColIndex("�ļ���")) & "'"
                mrsFiles.Delete (adAffectCurrent)
            End If
        Next
    End With
    mblnOK = True
    Unload Me
End Sub

Private Sub Command2_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = vsf.ColIndex("��������") Then
        If vsf.TextMatrix(Row, vsf.ColIndex("��������")) = "��������" Then
            vsf.TextMatrix(Row, vsf.ColIndex("��������")) = Space$(1)
            vsf.TextMatrix(Row, vsf.ColIndex("˵��")) = "���������뱨��" & vsf.TextMatrix(Row, vsf.ColIndex("��������"))
        End If
        If vsf.TextMatrix(Row, vsf.ColIndex("��������")) = "���ǵ���" Then
            If vsf.TextMatrix(Row, vsf.ColIndex("��������")) = Space$(1) Then
                vsf.TextMatrix(Row, vsf.ColIndex("��������")) = "���帲��"
                vsf.TextMatrix(Row, vsf.ColIndex("˵��")) = "����" & vsf.TextMatrix(Row, vsf.ColIndex("��������")) & "���Ḳ��ԭ�б���"
            End If
        End If
    End If
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = vsf.ColIndex("��������") Then
        '�ж��Ƿ�������޷�����
        If Val(vsf.TextMatrix(NewRow, vsf.ColIndex("�����ʶ"))) > 0 Then Cancel = True
        If vsf.TextMatrix(NewRow, vsf.ColIndex("��������")) = "��������" Then Cancel = True
    End If
    If NewCol = vsf.ColIndex("��������") Then
        '�ж��Ƿ�������޷�����
        If Val(vsf.TextMatrix(NewRow, vsf.ColIndex("�����ʶ"))) > 0 Then Cancel = True
        '�ж��жϸñ����Ƿ���鱨�������Ƿ��鱨�������������������������������룬����ֹ�༭�����ݵ�Ԫ��
        mrsFiles.Filter = ""
        mrsFiles.Filter = "FileName='" & vsf.TextMatrix(NewRow, vsf.ColIndex("�ļ���")) & "'"
        If Val(mrsFiles("��ID").Value) > 0 Then
            Cancel = True
        End If
        If Val(mrsFiles("ͬ��ID").Value) = 0 Then
            Cancel = True
        End If
    End If
End Sub

