VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm������Ŀѡ���Ԫ���� 
   AutoRedraw      =   -1  'True
   Caption         =   "ҽ����Ŀѡ��"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm������Ŀѡ���Ԫ����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7845
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   3690
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7845
      TabIndex        =   5
      Top             =   4350
      Width           =   7845
      Begin VB.CommandButton cmdRequery 
         Caption         =   "������ϸ"
         Height          =   350
         Left            =   3900
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ�б�"
         Height          =   350
         Left            =   2790
         TabIndex        =   10
         Top             =   150
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   7
         Top             =   175
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   9
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   8
         Top             =   150
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ϸ����(&F)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1590
      Width           =   45
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4050
      Left            =   3060
      TabIndex        =   3
      Top             =   270
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   7144
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2434
      EndProperty
   End
   Begin MSComctlLib.ListView lvwClass 
      Height          =   3990
      Left            =   15
      TabIndex        =   1
      Top             =   285
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7038
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   15
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ���Ԫ����.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ���Ԫ����.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ����(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ��ϸ(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   2
      Top             =   30
      Width           =   4710
   End
End
Attribute VB_Name = "frm������Ŀѡ���Ԫ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String '�������,ҽ����ĿDetailCode
Private mrsDetail As ADODB.Recordset
Private mblnOK As Boolean
Private mint���� As Integer
Private mint���õ��� As Integer '����ר�ã�0��ʾ����������1��ʾ����������ɾ������˵���Ŀ��
Private mint���� As Integer
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    '����ѡ����Ŀ����
    mstrCode = Mid(lvwDetail.SelectedItem.Key, 2)
    mblnOK = True
    Unload Me
End Sub

Public Function GetCode(strCode As String, ByVal int���� As Integer, ByVal int���� As Integer) As Boolean
'���ܣ����һ���շ���Ŀ��ҽ������
'������strCode ����Ϊ��������������
'���أ��ɹ�����True
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, objItem As ListItem
    
    mblnOK = False
    mint���� = int����
    
    On Error GoTo ErrH
    
    Set rsTmp = New ADODB.Recordset
    Set mrsDetail = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    mrsDetail.CursorLocation = adUseClient
    mint���� = int����
    
    gstrSQL = "Select ���� AS CODE,���� AS NAME From ����֧������ where ����=[1] order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", mint����)
        
    If mint���� = TYPE_�ɶ����� Then
        gstrSQL = "Select ������� as CLASSCODE,���� AS CODE,���� AS NAME ,����," & _
                  "substr(��ע,1,instr(��ע," & "'|'" & ",1,1)-1) As �������," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,1)+1,instr(��ע," & "'|'" & ",1,2)-1-instr(��ע," & "'|'" & ",1,1)) As �Ը�����," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,2)+1,instr(��ע," & "'|'" & ",1,3)-1-instr(��ע," & "'|'" & ",1,2)) As ������Ŀ," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,3)+1,instr(��ע," & "'|'" & ",1,4)-1-instr(��ע," & "'|'" & ",1,3)) As ���Ʒ�Χ," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,4)+1,instr(��ע," & "'|'" & ",1,5)-1-instr(��ע," & "'|'" & ",1,4)) As ��ע" & _
                  " from ҽ���շ���Ŀ_���� where ����=[1] and ����=[2] order by �������,����"
    Else
        gstrSQL = "Select ������� as CLASSCODE,���� AS CODE,���� AS NAME ,����," & _
                  "substr(��ע,1,instr(��ע," & "'|'" & ",1,1)-1) As �������," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,1)+1,instr(��ע," & "'|'" & ",1,2)-1-instr(��ע," & "'|'" & ",1,1)) As �Ը�����," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,2)+1,instr(��ע," & "'|'" & ",1,3)-1-instr(��ע," & "'|'" & ",1,2)) As ������Ŀ," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,3)+1,instr(��ע," & "'|'" & ",1,4)-1-instr(��ע," & "'|'" & ",1,3)) As ���Ʒ�Χ," & _
                  "substr(��ע,instr(��ע," & "'|'" & ",1,4)+1,instr(��ע," & "'|'" & ",1,5)-1-instr(��ע," & "'|'" & ",1,4)) As ��ע" & _
                  " from ҽ���շ���Ŀ where ����=[1] and ����=[2] order by �������,����"
    End If
    Set mrsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", mint����, int����)
    
    'Ϊ��ϸ���Ӷ�����ʾ����
    Dim fld As ADODB.Field
    For Each fld In mrsDetail.Fields
        If fld.Name <> "CLASSCODE" And fld.Name <> "NAME" And fld.Name <> "CODE" Then
            If fld.Name <> "��ע" Then
                lvwDetail.ColumnHeaders.Add , , fld.Name, 1000
            End If
        End If
    Next
    
    '��ʼ������
    If rsTmp.State = adStateOpen Then
        If Not rsTmp.EOF Then
            lvwClass.ListItems.Clear
            For i = 1 To rsTmp.RecordCount
                Set objItem = lvwClass.ListItems.Add(, "_" & rsTmp("CODE"), rsTmp("CODE"), , "Class")
                objItem.SubItems(1) = IIf(IsNull(rsTmp("NAME")), "", rsTmp("NAME"))
                rsTmp.MoveNext
            Next
        End If
    Else
        '�����������û�д����
        lblClass.Visible = False
        lvwClass.Visible = False
        picSplit.Visible = False
        Call lvwClass.ListItems.Add(, "_1", "1", , "Class")
    End If
    cmdRequery.Visible = True
    
    If Not mrsDetail.EOF Then
       If mstrCode <> "" Then
            '���Ҵ�����벢��λ
            mrsDetail.Filter = "CODE Like '" & UCase(mstrCode) & "%'"
            If Not mrsDetail.EOF Then
                lvwClass.ListItems("_" & mrsDetail("CLASSCODE")).Selected = True
            ElseIf lvwClass.ListItems.Count > 0 Then
                lvwClass.ListItems(1).Selected = True
            End If
            Call lvwClass_ItemClick(lvwClass.SelectedItem)
            lvwClass.SelectedItem.EnsureVisible
        Else
            If lvwClass.ListItems.Count > 0 Then
                lvwClass.ListItems(1).Selected = True
            End If
            Call lvwClass_ItemClick(lvwClass.SelectedItem)
        End If
        
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    
    frm������Ŀѡ���Ԫ����.Show 1
    '����ֵ
    If mblnOK = True Then
        strCode = mstrCode
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdPrint_Click()
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As New zlPrintLvw
    
    
    objPrint.Title.Text = "������Ŀ"
    Set objPrint.Body.objData = lvwDetail
    objPrint.UnderAppItems.Add "ҽ�����ࣺ" & lvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
    End Select

End Sub

Private Sub cmdRequery_Click()
    Dim str�������� As String
    Dim str��ע As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnȫ�� As Boolean
    Dim blnReturn As Boolean
    
    If MsgBox("���������ܻỨ�Ƚϳ���ʱ�䣬�Ƿ������" & vbCrLf & vbCrLf & "����ע�⣬������ֻ����ҽ����Ŀ��ϸ������������Ӧ��ϵ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
      With rsTemp
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Fields.Append "CLASSCODE", adVarChar, 6   '�������
        .Fields.Append "CODE", adVarChar, 20       '���
        .Fields.Append "NAME", adVarChar, 40       '����
        .Fields.Append "PY", adVarChar, 10         '����
        .Fields.Append "FYLB", adVarChar, 10       '�������
        .Fields.Append "ZFBL", adVarChar, 6       '�Ը�����
        .Fields.Append "FYXM", adVarChar, 14       '������Ŀ
        .Fields.Append "XZFW", adVarChar, 100      '���Ʒ�Χ
        .Fields.Append "BZ", adVarChar, 100        '��ע
        .Fields.Append "MEMO", adVarChar, 500      '��ע
        .Open
      End With
      
    blnȫ�� = True
    Me.Caption = "ҽ����Ŀѡ�����ڶ�ȡ���ļ��������ȡ������Ŀ��ϸ�����Ժ�......��"
    If MsgBox("�Ƿ����ԭ����ҽ����Ŀ��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) <> vbYes Then
        blnȫ�� = False
    End If
    blnReturn = ҽ����Ŀ_�ɶ��ڽ�(rsTemp)
    
    If blnReturn = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.Caption = "ҽ����Ŀѡ�����ڸ���ҽ����Ŀ......��"
    If mint���� = TYPE_�ɶ����� Then
        gcnOracle_�ɶ�����.BeginTrans
    ElseIf mint���� = TYPE_�ϳ����� Then
        gcnOracle_�ϳ�����.BeginTrans
    Else
        gcnOracle_��Ԫ����.BeginTrans
    End If
    On Error GoTo errHandle
    If blnȫ�� Then
        gstrSQL = "zl_ҽ���շ���Ŀ_Clear(" & mint���� & "," & mint���� & ")"
        If mint���� = TYPE_�ɶ����� Then
            Call ExecuteProcedure_�ɶ�����("ҽ����Ŀѡ��")
        ElseIf mint���� = TYPE_�ϳ����� Then
            Call ExecuteProcedure_�ϳ�����("ҽ����Ŀѡ��")
        Else
            Call ExecuteProcedure_��Ԫ����("ҽ����Ŀѡ��")
        End If
    End If
    
    '���±�����Ŀ
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    Do Until rsTemp.EOF
        rsTemp("FYLB") = IIf(Trim(rsTemp("FYLB")) = "", "��", rsTemp("FYLB"))
        rsTemp("ZFBL") = IIf(Trim(rsTemp("ZFBL")) = "", "��", rsTemp("ZFBL"))
        rsTemp("FYXM") = IIf(Trim(rsTemp("FYXM")) = "", "��", rsTemp("FYXM"))
        rsTemp("XZFW") = IIf(Trim(rsTemp("XZFW")) = "", "��", rsTemp("XZFW"))
        rsTemp("BZ") = IIf(Trim(rsTemp("BZ")) = "", "��", rsTemp("BZ"))
        rsTemp("MEMO") = IIf(Trim(rsTemp("MEMO")) = "", "��", rsTemp("MEMO"))
        str��ע = rsTemp("FYLB") & "|" & rsTemp("ZFBL") & "|" & rsTemp("FYXM") & "|" & rsTemp("XZFW") & "|" & rsTemp("BZ") & "|" & rsTemp("MEMO")

        '���뱣����Ŀ
        gstrSQL = "zl_ҽ���շ���Ŀ_Insert(" & mint���� & "," & mint���� & ",'" & rsTemp("CODE") & "','" & ToVarchar(rsTemp("NAME"), 40) & _
            "','" & ToVarchar(rsTemp("PY"), 10) & "','" & ToVarchar(rsTemp("CLASSCODE"), 6) & "','" & ToVarchar(str��ע, 500) & "')"
        If mint���� = TYPE_�ɶ����� Then
            ExecuteProcedure_�ɶ����� ("����ҽ����Ŀ")
        ElseIf mint���� = TYPE_�ϳ����� Then
            Call ExecuteProcedure_�ϳ�����("����ҽ����Ŀ")
        Else
            Call ExecuteProcedure_��Ԫ����("����ҽ����Ŀ")
        End If
        Me.Caption = "ҽ����Ŀѡ�����ڸ���ҽ����Ŀ���Ѳ���" & rsTemp.AbsolutePosition & "����¼��"
        rsTemp.MoveNext
    Loop
    
    If mint���� = TYPE_�ɶ����� Then
        gcnOracle_�ɶ�����.CommitTrans
    ElseIf mint���� = TYPE_�ϳ����� Then
        gcnOracle_�ϳ�����.CommitTrans
    Else
        gcnOracle_��Ԫ����.CommitTrans
    End If
    '����װ����ϸ
    mrsDetail.Requery
    Call lvwClass_ItemClick(lvwClass.SelectedItem)
    MousePointer = vbDefault
    Me.Caption = "ҽ����Ŀѡ��"
    MsgBox "������ɡ�", vbInformation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If mint���� = TYPE_�ɶ����� Then
        gcnOracle_�ɶ�����.RollbackTrans
    ElseIf mint���� = TYPE_�ϳ����� Then
        gcnOracle_�ϳ�����.RollbackTrans
    Else
        gcnOracle_��Ԫ����.RollbackTrans
    End If
    MousePointer = vbDefault
End Sub
Private Function ҽ����Ŀ_�ɶ��ڽ�(ByVal rsTemp As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��������
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Const COL_��Ŀ����   As Long = 1
    Const COL_���� As Long = 2
    Const COL_���� As Long = 3
    Const COL_����  As Long = 4
    Const COL_�������  As Long = 5
    Const COL_�Ը�����  As Long = 6
    Const COL_������Ŀ  As Long = 7
    Const COL_���Ʒ�Χ  As Long = 8
    Const COL_��ע  As Long = 9
    Const COL_����  As Long = 10
    Err = 0
    On Error GoTo errHand:
    Dim ObjExcel As Object, ObjCell As Object, strFile As String, strValue As String
    
    'ѡ��ָ���ļ�
    On Error Resume Next
    Err = 0
    With dlg
        .Filter = "EXCEL�ļ�(*.xls)|*.xls"
        .flags = cdlOFNFileMustExist Or cdlOFNLongNames
        .ShowOpen
        If Err <> 0 Then Exit Function
        strFile = .FileName
    End With
    
    '����EXCEL����
    On Error Resume Next
    Err = 0
    Set ObjExcel = CreateObject("Excel.Application")
    If Err <> 0 Then
        MsgBox "EXCELδ��ȷ��װ������ȷ��װEXCEL���İ�������У�", vbInformation, gstrSysName
        Exit Function
    End If
    
    On Error GoTo errHand:
    Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ����......��"
    Dim strCode As String
    
    'ȡEXCEL�ļ�������
    With ObjExcel
        .Workbooks.Open strFile
        
        'ȡ���е�ֵ
        Dim lngRow As Long
        lngRow = 2
        Do While True
            If .ActiveSheet.Cells(lngRow, COL_����) <> "" Then
                rsTemp.AddNew
                rsTemp("Code") = Mid(Trim(.ActiveSheet.Cells(lngRow, COL_����)), 1, 20)
                rsTemp("Name") = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_����)), 40), "'", "")
                rsTemp("PY") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_����)), 10)
                strCode = Replace(ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_��Ŀ����)), 6), "'", "")
                strCode = Decode(strCode, "ҩƷ", 0, "����", 1, "����", 1, "���", 2, "����", 3, strCode)
                rsTemp("CLASSCODE") = strCode
                If mint���� = TYPE_�ϳ����� Or TYPE_��Ԫ���� Then
                    rsTemp("FYLB") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_�������)), 10)
                    rsTemp("ZFBL") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_�Ը�����)), 6)
                    rsTemp("FYXM") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_������Ŀ)), 14)
                    rsTemp("XZFW") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_���Ʒ�Χ)), 100)
                    rsTemp("BZ") = ToVarchar(Trim(.ActiveSheet.Cells(lngRow, COL_��ע)), 100)
                End If
                
                rsTemp.Update
                Me.Caption = "ҽ����Ŀѡ�����ڴ�EXCEL�ļ�����ȡ���ݣ��ѻ�ȡ" & rsTemp.RecordCount & "����¼��"
            Else
                Exit Do
            End If
            lngRow = lngRow + 1
        Loop
    End With
    
    '�ر�EXCEL����
    ObjExcel.quit
    Set ObjExcel = Nothing
    ҽ����Ŀ_�ɶ��ڽ� = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = lvwClass.Width
    
    On Error Resume Next
    
    lvwClass.Left = 0: lvwClass.Top = lblClass.Top + lblClass.Height
    lvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = lvwClass.Top
    picSplit.Left = lvwClass.Left + lvwClass.Width
    picSplit.Height = lvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If lvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    
    lvwDetail.Top = lvwClass.Top
    lvwDetail.Left = lblDetail.Left
    lvwDetail.Width = lblDetail.Width
    lvwDetail.Height = lvwClass.Height
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwDetail_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If lvwClass.Width + x < 1000 Or lvwDetail.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        lvwClass.Width = lvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        lvwDetail.Left = lvwDetail.Left + x
        lvwDetail.Width = lvwDetail.Width - x
    End If
End Sub

Private Sub lvwdetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwDetail.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwDetail.SortOrder = lvwDescending
    Else
        lvwDetail.SortOrder = lvwAscending
    End If
    lvwDetail.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwDetail.SelectedItem Is Nothing Then lvwDetail.SelectedItem.EnsureVisible
End Sub

Private Sub lvwclass_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwClass.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwClass.SortOrder = lvwDescending
    Else
        lvwClass.SortOrder = lvwAscending
    End If
    lvwClass.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwClass.SelectedItem Is Nothing Then lvwClass.SelectedItem.EnsureVisible
End Sub

Private Sub lvwClass_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, objItem As ListItem
    Dim lngCount As Long, str�� As String, bln���⴦�� As Boolean
    Dim BLNSEL As Boolean
    Dim varPart As Variant
    
    
    Me.MousePointer = vbHourglass
    lvwDetail.ListItems.Clear
    If Item Is Nothing Then Exit Sub
    
    mrsDetail.Filter = "CLASSCODE='" & Mid(Item.Key, 2) & "'"
    If Not mrsDetail.EOF Then
        For i = 1 To mrsDetail.RecordCount
            Set objItem = lvwDetail.ListItems.Add(, "_" & mrsDetail("CODE"), mrsDetail("CODE"), , "Detail")
            objItem.SubItems(1) = IIf(IsNull(mrsDetail("NAME")), "", mrsDetail("NAME"))
            objItem.Tag = mrsDetail("CLASSCODE")
            
            '��ʾ�������
            With lvwDetail.ColumnHeaders
                For lngCount = 3 To lvwDetail.ColumnHeaders.Count
                    str�� = .Item(lngCount).Text
                    'û�н������⴦��
                    objItem.SubItems(lngCount - 1) = IIf(IsNull(mrsDetail(.Item(lngCount).Text)), "", mrsDetail(.Item(lngCount).Text))
                Next
            End With
                        
            If InStr(mrsDetail("CODE"), mstrCode) > 0 And Not BLNSEL Then
                objItem.Selected = True
                BLNSEL = True
            End If
            mrsDetail.MoveNext
        Next
        If Not BLNSEL And lvwDetail.ListItems.Count > 0 Then lvwDetail.ListItems(1).Selected = True
        lvwDetail.SelectedItem.EnsureVisible
    End If
    Call zlControl.LvwSetColWidth(lvwDetail)
    Me.MousePointer = vbDefault
End Sub

Private Sub txtFind_Change()
'���ܣ������û���������ݲ���ƥ�������
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvwDetail.ListItems.Count = 0 Then Exit Sub
    
    Set lst = lvwDetail.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '���ı�������������ƥ��
        lngSubItems = lvwDetail.ColumnHeaders.Count - 1
        For Each lst In lvwDetail.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub
Private Sub AddRecord(rsObj As ADODB.Recordset, ByVal str���� As String, ByVal str���� As String, _
str���� As String, ByVal str��ע As String, ByVal str���� As String)
    With rsObj
        .AddNew
        !CODE = str����
        !Name = Replace(str����, "'", "")
        !py = Replace(str����, "'", "")
        !Memo = Replace(str��ע, "'", "")
        !ClassCode = str����
        .Update
    End With
End Sub


