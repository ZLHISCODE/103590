VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendBodyActiveItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ŀ"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBodyActiveItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picCloumn 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   5955
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5955
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2610
         MaxLength       =   20
         TabIndex        =   5
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   2460
         Picture         =   "frmCaseTendBodyActiveItem.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "ȷ��"
         Top             =   2460
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   3000
         Picture         =   "frmCaseTendBodyActiveItem.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "ȡ��"
         Top             =   2460
         Width           =   450
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   2430
         TabIndex        =   7
         Top             =   1515
         Width           =   1095
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
         Height          =   300
         Index           =   1
         Left            =   2430
         TabIndex        =   8
         Top             =   1845
         Width           =   1095
      End
      Begin MSComctlLib.ListView lstColumnItems 
         Height          =   2490
         Left            =   60
         TabIndex        =   4
         Top             =   480
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstColumnUsed 
         Height          =   2490
         Left            =   3855
         TabIndex        =   6
         Top             =   480
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2160
         TabIndex        =   11
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ѷ�������,������ɾ��."
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Top             =   945
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ�Ļ�����Ŀ"
         Height          =   180
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   1620
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ѿ�ѡ��Ļ�����Ŀ"
         Height          =   180
         Left            =   3855
         TabIndex        =   2
         Top             =   180
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmCaseTendBodyActiveItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjVsf As Object
Private mblnChage As Boolean
Private mstrSQL As String
Private mstrActiveItem As String
Private mblnInit As Boolean
Private mlng����ȼ� As Long
Private mlngӤ�� As Long
Private mlng����ID As Long

Private Enum TYPE_Tab
    COL_tab������ = 0
    COL_tab�ַ��� = 1
    COL_tab��Ŀ��� = 2
    COL_tab��Ŀ�� = 3  '--��������λ
    COL_TabNull = 4
    COL_tab��Ŀ���� = 5 '--������λ
End Enum

Public Function ShowMe(objVsf As Object, ByVal frmParent As Form, ByVal lng����ȼ� As Long, ByVal lngӤ�� As Long, ByVal lng����ID As Long) As Boolean
    mblnChage = False
    mstrActiveItem = ""
    Set mobjVsf = objVsf
    mlng����ȼ� = lng����ȼ�
    mlngӤ�� = lngӤ��
    mlng����ID = lng����ID
    If Not BoundItems Then Exit Function
    lblNote.Visible = False
    mblnInit = True
    Me.Show 1, frmParent
    ShowMe = mblnChage
End Function

Private Sub cmdFilterCancel_Click()
    mblnChage = False
    Unload Me
End Sub

Private Sub cmdFilterOK_Click()
'
    Dim intItem As Integer, intRow As Integer, i As Integer
    Dim lngItemCode As Integer, strItemName As String
    Dim blnAdd As Boolean, blnDelete As Boolean
    Dim strPart As String
    Dim arrStr() As String
    Dim arrTmp() As String, varCode() As String
    
    arrTmp = Split(mstrActiveItem, ";")
    
    '��ӻ��Ŀ
    For intItem = 1 To lstColumnUsed.ListItems.Count
        lngItemCode = Val(lstColumnUsed.ListItems(intItem).Text)
        strItemName = lstColumnUsed.ListItems(intItem).SubItems(1)
        strPart = lstColumnUsed.ListItems(intItem).SubItems(2)
        blnAdd = True
        For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
            If Val(Split(mobjVsf.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 Then
                If lngItemCode = Val(mobjVsf.TextMatrix(intRow, COL_tab��Ŀ���)) And strItemName = mobjVsf.TextMatrix(intRow, COL_tab��Ŀ��) Then
                    blnAdd = False
                    Exit For
                End If
            End If
        Next intRow
        
        If blnAdd = True Then
            mblnChage = True
            For i = LBound(arrTmp) To UBound(arrTmp)
                varCode = Split(arrTmp(i), "'")
                If Val(varCode(2)) = lngItemCode And varCode(4) = strItemName Then
                    mobjVsf.Rows = mobjVsf.Rows + 1
                    arrStr = Split(varCode(1), ",")
                    If UBound(arrStr) > 6 Then arrStr(7) = strPart
                    varCode(1) = Join(arrStr, ",")
                    mobjVsf.TextMatrix(intRow, COL_tab������) = varCode(0)
                    mobjVsf.TextMatrix(intRow, COL_tab�ַ���) = varCode(1)
                    mobjVsf.TextMatrix(intRow, COL_tab��Ŀ���) = lngItemCode
                    mobjVsf.TextMatrix(intRow, COL_tab��Ŀ��) = strItemName
                    mobjVsf.TextMatrix(intRow, COL_TabNull) = ""
                    mobjVsf.TextMatrix(intRow, COL_tab��Ŀ����) = varCode(3)
                    '��λ������ӵ���
                    mobjVsf.Row = mobjVsf.Rows - 1: mobjVsf.Col = mobjVsf.FixedCols
                End If
            Next i
        End If
    Next intItem
    '��Ҫ�������û�а󶨹̶���Ŀ�����
    If mobjVsf.Rows > mobjVsf.FixedRows + 1 And mobjVsf.Tag = "NO" Then
        mobjVsf.Tag = ""
        Call mobjVsf.RemoveItem(mobjVsf.FixedRows)
    End If
    'ɾ�����Ŀ
    For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
        If intRow > mobjVsf.Rows - 1 Then Exit For
        If Val(Split(mobjVsf.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 Then
            lngItemCode = Val(mobjVsf.TextMatrix(intRow, COL_tab��Ŀ���))
            strItemName = mobjVsf.TextMatrix(intRow, COL_tab��Ŀ��)
            blnDelete = True
            For intItem = 1 To lstColumnUsed.ListItems.Count
                If lngItemCode = Val(lstColumnUsed.ListItems(intItem).Text) And strItemName = lstColumnUsed.ListItems(intItem).SubItems(1) Then
                    blnDelete = False
                    Exit For
                End If
            Next intItem
            
            If blnDelete = True Then
                mblnChage = True
                If mobjVsf.Rows = mobjVsf.FixedRows + 1 And intRow = mobjVsf.FixedRows Then
                    '��Ҫ�������û�а󶨹̶���Ŀ�����
                    mobjVsf.Cell(flexcpText, intRow, 0, intRow, mobjVsf.Cols - 1) = ""
                    varCode = Split("',0,0,0,0,1,0,,0'-999''", "'")
                    mobjVsf.TextMatrix(intRow, COL_tab������) = varCode(0)
                    mobjVsf.TextMatrix(intRow, COL_tab�ַ���) = varCode(1)
                    mobjVsf.TextMatrix(intRow, COL_tab��Ŀ���) = varCode(2)
                    mobjVsf.TextMatrix(intRow, COL_tab��Ŀ��) = varCode(4)
                    mobjVsf.TextMatrix(intRow, COL_TabNull) = ""
                    mobjVsf.TextMatrix(intRow, COL_tab��Ŀ����) = varCode(3)
                    mobjVsf.Tag = "NO"
                    '��λ������ӵ���
                    mobjVsf.Row = mobjVsf.Rows - 1: mobjVsf.Col = mobjVsf.FixedCols
                Else
                    Call mobjVsf.RemoveItem(intRow)
                    intRow = intRow - 1
                End If
            End If
        End If
    Next intRow
    
    Unload Me
End Sub


Private Function BoundItems() As Boolean
'---------------------------------------------------------------------
'����:��ȡ���Ŀ��Ϣ
'---------------------------------------------------------------------
    Dim lstItem As ListItem
    Dim rsActive As New ADODB.Recordset
    Dim arrActive() As String, ArrCode() As String
    Dim strSQL As String, strText As String
    Dim i As Integer, j As Integer
    Dim strItemCode As String, strֵ�� As String
    Dim intRow As Integer
    On Error GoTo Errhand
    
    If mobjVsf Is Nothing Then Exit Function
    
    For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
        If Val(Split(mobjVsf.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 Then
            strText = ""
            strText = "" & mobjVsf.TextMatrix(intRow, COL_tab��Ŀ���) & " ��Ŀ���,'" & mobjVsf.TextMatrix(intRow, COL_tab��Ŀ��) & "' ��Ŀ����,1 ��ʶ"
            strSQL = strSQL & IIf(strSQL = "", "", "UNION ALL") & " SELECT " & strText & " FROM Dual "
        End If
    Next intRow
    
    mstrSQL = "" & _
            "Select a.��Ŀ���, a.��Ŀ����,a.��λ ,a.��Ŀֵ��,a.��Ŀ����,a.��Ŀ����, a.��Ŀ����, a.��ĿС��, a.��¼Ƶ��,a.��Ժ�ײ�, a.������,a.��Ŀ��λ, a.��Ŀ��ʾ," & vbNewLine & _
            IIf(strSQL = "", "0 ��ʶ", "            Nvl(b.��ʶ, 0) ��ʶ") & vbNewLine & _
            "From (Select a.��Ŀ���, c.��λ || b.��Ŀ���� ��Ŀ����,c.��λ, b.��Ŀֵ��, b.��Ŀ����, b.��Ŀ����, b.��ĿС��," & vbNewLine & _
            "                           Nvl(a.��¼Ƶ��, 2) ��¼Ƶ��,A.��Ժ�ײ�, b.������, b.��Ŀ��ʾ,b.��Ŀ����,b.��Ŀ��λ" & vbNewLine & _
            "            From ���¼�¼��Ŀ a, ���²�λ c, �����¼��Ŀ b" & vbNewLine & _
            "            Where a.��Ŀ��� = b.��Ŀ��� And b.��Ŀ��� = c.��Ŀ���(+) And b.��Ŀ���� = 2 And Nvl(b.Ӧ�÷�ʽ, 0) = 1 And" & vbNewLine & _
            "                        b.����ȼ� >= [1] And Nvl(b.���ò���, 0) In (0, [2]) And" & vbNewLine & _
            "                        (b.���ÿ��� = 1 Or" & vbNewLine & _
            "                        (b.���ÿ��� = 2 And Exists (Select 1 From �������ÿ��� d Where d.��Ŀ��� = b.��Ŀ��� And d.����id = [3])))) a"
    If strSQL <> "" Then
        mstrSQL = mstrSQL & vbNewLine & ",(" & strSQL & ") b" & vbNewLine & _
            "Where a.��Ŀ��� = b.��Ŀ���(+) And a.��Ŀ���� = b.��Ŀ����(+)"
    End If
    mstrSQL = mstrSQL & vbNewLine & "   Order By a.��Ŀ���, a.��Ŀ����"
            
    Set rsActive = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡδ���õĻ��Ŀ", mlng����ȼ�, IIf(mlngӤ�� = 0, 1, 2), mlng����ID)
    
    If rsActive.RecordCount = 0 Then
        MsgBox "û�пɹ�ѡ��Ļ��Ŀ�����ڻ�����Ŀ����ģ���н������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������Ŀ
    txtFind.Text = ""
    lstColumnItems.ListItems.Clear
    lstColumnUsed.ListItems.Clear
    strItemCode = ""
    mstrActiveItem = ""
    
    With rsActive
        Do While Not .EOF
            strֵ�� = zlCommFun.Nvl(!��Ŀֵ��)
            If zlCommFun.Nvl(!��Ŀ����) = 0 Then
                If InStr(1, strֵ��, ";") <> 0 Then strֵ�� = Split(strֵ��, ";")(0) & "��" & Split(strֵ��, ";")(1)
            End If
            strֵ�� = Replace(Replace(Replace(strֵ��, ";", ":"), "'", ""), ",", "")
            If strItemCode = "" Then
                strItemCode = !��Ŀ��� & "'" & Nvl(!��Ŀ����)
                mstrActiveItem = zlCommFun.Nvl(!������, "2)���±����Ŀ") & "'" & strֵ�� & "," & zlCommFun.Nvl(!��Ŀ����) & "," & _
                    zlCommFun.Nvl(!��ĿС��) & "," & zlCommFun.Nvl(!��¼Ƶ��) & "," & zlCommFun.Nvl(!��Ŀ��ʾ) & "," & zlCommFun.Nvl(!��Ŀ����) & "," & _
                    zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��λ) & "," & zlCommFun.Nvl(!��Ժ�ײ�, 0) & "'" & _
                    zlCommFun.Nvl(!��Ŀ���) & "'" & Replace(zlCommFun.Nvl(!��Ŀ����) & IIf(zlCommFun.Nvl(!��Ŀ��λ, "") = "", "", "(" & !��Ŀ��λ & ")"), ";", ":") & "'" & zlCommFun.Nvl(!��Ŀ����)

            Else
                If InStr(1, "," & strItemCode & ",", "," & !��Ŀ��� & "'" & Nvl(!��Ŀ����) & ",") = 0 Then
                    strItemCode = strItemCode & "," & !��Ŀ��� & "'" & Nvl(!��Ŀ����)
                    mstrActiveItem = mstrActiveItem & ";" & zlCommFun.Nvl(!������, "2)���±����Ŀ") & "'" & strֵ�� & "," & zlCommFun.Nvl(!��Ŀ����) & "," & _
                        zlCommFun.Nvl(!��ĿС��) & "," & zlCommFun.Nvl(!��¼Ƶ��) & "," & zlCommFun.Nvl(!��Ŀ��ʾ) & "," & zlCommFun.Nvl(!��Ŀ����) & "," & _
                        zlCommFun.Nvl(!��Ŀ����) & "," & zlCommFun.Nvl(!��λ) & "," & zlCommFun.Nvl(!��Ժ�ײ�, 0) & "'" & _
                        zlCommFun.Nvl(!��Ŀ���) & "'" & Replace(zlCommFun.Nvl(!��Ŀ����) & IIf(zlCommFun.Nvl(!��Ŀ��λ, "") = "", "", "(" & !��Ŀ��λ & ")"), ";", ":") & "'" & zlCommFun.Nvl(!��Ŀ����)
                End If
            End If
            
            If !��ʶ = 1 Then
                Set lstItem = lstColumnUsed.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnUsed.ListItems.Count, !��Ŀ���)
                lstItem.SubItems(1) = zlCommFun.Nvl(!��Ŀ����)
                lstItem.SubItems(2) = zlCommFun.Nvl(!��λ)
            Else
                Set lstItem = lstColumnItems.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnItems.ListItems.Count + 100, !��Ŀ���)
                lstItem.SubItems(1) = zlCommFun.Nvl(!��Ŀ����)
                lstItem.SubItems(2) = zlCommFun.Nvl(!��λ)
            End If
            .MoveNext
        Loop
    End With
    
    BoundItems = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnItems_DblClick
End Sub

Private Sub lstColumnUsed_DblClick()
    Call cmdColumn_Click(1)
End Sub

Private Sub lstColumnUsed_ItemClick(ByVal Item As MSComctlLib.ListItem)
        '����Ƿ��������,������������ʾ�ò�������ɾ��
    If Not Item Is Nothing Then
        If CheckGridData(Val(Item.Text), Item.SubItems(1)) Then
            lblNote.Caption = Item.SubItems(1) & "�Ѿ���������,���ܽ���ɾ��."
            lblNote.Visible = True
            cmdColumn(1).Enabled = False
        Else
            lblNote.Caption = ""
            lblNote.Visible = False
            cmdColumn(1).Enabled = True
        End If
    End If
End Sub

Private Sub lstColumnUsed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnUsed_DblClick
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim lstItem As ListItem
    
    If Index = 0 Then
        'add
        If Not lstColumnItems.SelectedItem Is Nothing Then
            Set lstItem = lstColumnUsed.ListItems.Add(, lstColumnItems.SelectedItem.Key, lstColumnItems.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnItems.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnItems.SelectedItem.SubItems(2)
            lstColumnItems.ListItems.Remove lstColumnItems.SelectedItem.Index
        End If
    Else
        'del
        If Not lstColumnUsed.SelectedItem Is Nothing Then
            If CheckGridData(Val(lstColumnUsed.SelectedItem.Text), lstColumnUsed.SelectedItem.SubItems(1)) = True Then Exit Sub
            Set lstItem = lstColumnItems.ListItems.Add(, lstColumnUsed.SelectedItem.Key, lstColumnUsed.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnUsed.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnUsed.SelectedItem.SubItems(2)
            lstColumnUsed.ListItems.Remove lstColumnUsed.SelectedItem.Index
        End If
    End If
End Sub

Private Function CheckGridData(ByVal lngID As Long, ByVal strName As String) As Boolean
'-------------------------------------------------------------------
'��鵱����Ŀ�Ƿ��������,������������ɾ��
'����:lngID ��Ŀ��� strName ��Ŀ���ƣ���Ŀ����+��λ��
'-------------------------------------------------------------------
    CheckGridData = True
    Dim intRow As Integer, intCOl As Integer

    For intRow = mobjVsf.FixedRows To mobjVsf.Rows - 1
        If Val(mobjVsf.TextMatrix(intRow, COL_tab��Ŀ���)) = lngID And mobjVsf.TextMatrix(intRow, COL_tab��Ŀ����) = strName Then
            Exit For
        End If
    Next intRow
    
    If intRow > mobjVsf.Rows - 1 Then CheckGridData = False: Exit Function
    
    '�����Ŀ���Ƿ��������
    For intCOl = mobjVsf.FixedCols To Val(Split(mobjVsf.TextMatrix(intRow, COL_tab�ַ���), ",")(3)) + mobjVsf.FixedCols - 1  '��¼Ƶ��+�̶���
        If Trim(mobjVsf.TextMatrix(intRow, intCOl)) <> "" Then
            Exit Function
        End If
    Next intCOl
    
    CheckGridData = False
End Function

Private Sub txtFind_Change()
    Call txtFind_KeyDown(10, 0)
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = 100
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Static lngPreIndex As Long
    Dim strText As String
    Dim lngIndex As Long
    
    '61855:������,2013-11-07,�󶨻��Ŀ��ô����������
    strText = Trim(txtFind.Text)
    If KeyCode = 10 Or strText = "" Then
        '��Ҫ�������������ֵ
        lngPreIndex = 0
    ElseIf KeyCode = vbKeyReturn And strText <> "" Then
        If Not (lngPreIndex > 0 And lngPreIndex < lstColumnItems.ListItems.Count) Then lngPreIndex = 1
        For lngIndex = lngPreIndex To lstColumnItems.ListItems.Count
            If UCase(lstColumnItems.ListItems(lngIndex).SubItems(1)) Like UCase(strText) & "*" Then
                lstColumnItems.ListItems(lngIndex).Selected = True
                lstColumnItems.ListItems(lngIndex).EnsureVisible
                Exit For
            End If
        Next
        
        If lngIndex > lstColumnItems.ListItems.Count Then
            If lngPreIndex > 1 Then
                For lngIndex = 1 To lstColumnItems.ListItems.Count
                    If UCase(lstColumnItems.ListItems(lngIndex).SubItems(1)) Like UCase(strText) & "*" Then
                        lstColumnItems.ListItems(lngIndex).Selected = True
                        lstColumnItems.ListItems(lngIndex).EnsureVisible
                        Exit For
                    End If
                Next
            End If
            lngPreIndex = 1
        Else
            lngPreIndex = lngIndex + 1
        End If
    End If
End Sub

