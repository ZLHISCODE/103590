VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   Icon            =   "frmSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ListView lvw 
      Height          =   2850
      Left            =   2535
      TabIndex        =   2
      Top             =   555
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   5027
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
      NumItems        =   0
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   6120
      Begin VB.Image Image1 
         Height          =   240
         Left            =   165
         Picture         =   "frmSelect.frx":014A
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   540
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   2355
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3660
      Width           =   6120
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4395
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3150
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   165
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSelect.frx":06D4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   2760
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   4868
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   2625
      Left            =   2190
      TabIndex        =   8
      Top             =   615
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   4630
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmSelect.frx":082E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4905
      TabIndex        =   9
      Top             =   315
      Width           =   435
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�룺SQL���ֶ�����
Public mstrSQLList As String
Public mstrSQLTree As String
Public mstrFLDList As String
Public mstrFLDTree As String
Public mstrParName As String '��������
Public mbytDataType As Byte      '������������
Public mstrMatch As String '����ƥ�������
Public mlngSeekHwnd As Long '���ڶ�λ����λ�õĿؼ�

'����δ����ʽ���������ԭʼֵ
Public mstrOutBand As String 'ѡ��İ�ֵ,��Ӧ&B
Public mstrOutDisp As String 'ѡ�����ʾֵ,��Ӧ&D

Private intPreNode As Long
Private blnItem As Boolean
Private blnSetFlex As Boolean, blnSetLvw As Boolean
Private rsList As ADODB.Recordset
Private strList As String
Private blnSave As Boolean
Private rParent As RECT

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strDisp As String, strBand As String
    
    strDisp = GetScript(mstrFLDList, "&D") '��ʾ���ֶ���
    strBand = GetScript(mstrFLDList, "&B") '�󶨵��ֶ���
    
    If strDisp = "" Or strBand = "" Then
        MsgBox "ѡ������û�ж��������İ󶨼���ʾ�ֶ���Ŀ��", vbInformation, App.Title
        Exit Sub
    End If
    
    If strList = "lvw" Then
        If lvw.SelectedItem Is Nothing Then
            MsgBox "û��ѡ���κ����ݣ�", vbInformation, App.Title
            If tvw_s.Visible Then tvw_s.SetFocus
            Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(0) = "" Then
            MsgBox "������""" & mstrParName & """����ʾ����Ŀ""" & strDisp & """Ϊ�գ���ѡ���������ݣ�", vbInformation, App.Title
            Exit Sub
        End If
        If Split(lvw.SelectedItem.Tag, "|")(1) = "" Then
            MsgBox "������ֵ""" & mstrParName & """�󶨵���Ŀ""" & strBand & """Ϊ�գ���ѡ���������ݣ�", vbInformation, App.Title
            Exit Sub
        End If
        '���ͼ��
        Select Case mbytDataType
            Case 1
                If Not IsNumeric(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    Exit Sub
                End If
            Case 2
                If Not IsDate(Split(lvw.SelectedItem.Tag, "|")(1)) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    Exit Sub
                End If
        End Select
        
        mstrOutDisp = Split(lvw.SelectedItem.Tag, "|")(0)
        mstrOutBand = Split(lvw.SelectedItem.Tag, "|")(1)
    Else
        '���FlexGrid�ɼ�,��rsListһ��������
        If msh.TextMatrix(msh.Row, GetColNum(strDisp)) = "" Then
            MsgBox "������""" & mstrParName & """����ʾ����Ŀ""" & strDisp & """Ϊ�գ���ѡ���������ݣ�", vbInformation, App.Title
            Exit Sub
        End If
        If msh.TextMatrix(msh.Row, GetColNum(strBand)) = "" Then
            MsgBox "������ֵ""" & mstrParName & """�󶨵���Ŀ""" & strBand & """Ϊ�գ���ѡ���������ݣ�", vbInformation, App.Title
            Exit Sub
        End If
        '���ͼ��
        Select Case mbytDataType
            Case 1
                If Not IsNumeric(msh.TextMatrix(msh.Row, GetColNum(strBand))) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    Exit Sub
                End If
            Case 2
                If Not IsDate(msh.TextMatrix(msh.Row, GetColNum(strBand))) Then
                    MsgBox "��Ŀ""" & strBand & """�����ݷ�������,���ܱ�ѡ��", vbInformation, App.Title
                    Exit Sub
                End If
        End Select
        
        mstrOutDisp = msh.TextMatrix(msh.Row, GetColNum(strDisp))
        mstrOutBand = msh.TextMatrix(msh.Row, GetColNum(strBand))
    End If
    gblnOK = True
    
    On Error Resume Next
    Hide
End Sub

Private Sub Form_Activate()
    If tvw_s.Visible Then
        If Not tvw_s.SelectedItem Is Nothing Then
            If tvw_s.SelectedItem.Key = "ALL" Then
                If lvw.Visible Then
                    lvw.SetFocus
                ElseIf msh.Visible Then
                    msh.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub msh_DblClick()
    If msh.MouseRow = 0 Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub msh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim lngW As Long, i As Integer
    
    If Not InDesign Then
        glngSelProc = GetWindowLong(hwnd, GWL_WNDPROC)
        Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SelMessage)
    End If
    
    gblnOK = False
    blnSave = True
    blnSetFlex = False '�Ƿ��Ѿ��Ա��ָ����
    blnSetLvw = False
    intPreNode = 0
    
    mstrOutBand = ""
    mstrOutDisp = ""
    
    msh.Tag = mstrParName
    lvw.Tag = mstrParName
    
    Me.Caption = mstrParName & "ѡ����"
    
    mstrSQLList = Replace(mstrSQLList, "[*]", mstrMatch)
    mstrSQLTree = Replace(mstrSQLTree, "[*]", mstrMatch)
    
    If mstrSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
        If Not FillList Then blnSave = False: Unload Me: Exit Sub
    Else
        tvw_s.Visible = True
        If Not FillTree Then blnSave = False: Unload Me: Exit Sub
        If tvw_s.Nodes.Count > 0 Then
            tvw_s.Nodes(1).Selected = True
            If Not tvw_s.Nodes(1).Child Is Nothing And mstrMatch = "" Then
                tvw_s.Nodes(1).Child.Selected = True
            End If
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        End If
    End If
    
    '����ƥ���Զ�����
    If mstrMatch <> "" Then
        If rsList.RecordCount = 1 Then
            blnSave = False
            Call cmdOK_Click
            Unload Me: Exit Sub
        ElseIf rsList.RecordCount = 0 Then
            MsgBox "û���ҵ���ƥ�����Ŀ,���������룡", vbInformation, App.Title
            blnSave = False
            Call cmdCancel_Click: Exit Sub
        End If
    End If
    
    Call Form_Resize
    
    '���弰�б�ȱʡ���
    Select Case strList
        Case "lvw"
            If lvw.ColumnHeaders.Count = 1 Then
                lvw.ColumnHeaders(1).Width = 2500
                Me.Width = 3000 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
            Else
                For i = 1 To lvw.ColumnHeaders.Count
                    lngW = lngW + lvw.ColumnHeaders(i).Width
                Next
                Me.Width = lngW + 500 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
                If Me.Width < 3000 Then Me.Width = 3000
            End If
        Case "msh"
            If msh.Cols = 1 Then
                msh.ColWidth(0) = 2500
                Me.Width = 3000 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
            Else
                For i = 0 To msh.Cols - 1
                    lngW = lngW + msh.ColWidth(i)
                Next
                Me.Width = lngW + 500 + IIf(mstrSQLTree = "", 0, tvw_s.Width + pic.Width)
                If Me.Width < 3000 Then Me.Width = 3000
            End If
    End Select
    If mstrSQLTree <> "" Then
        If Me.Width < (tvw_s.Width + pic.Width) * 2.2 Then Me.Width = (tvw_s.Width + pic.Width) * 2.2
    End If
    
    RestoreWinState Me, App.ProductName, mstrParName
    
    If mstrSQLTree = "" Then
        tvw_s.Visible = False
        pic.Visible = False
    Else
        tvw_s.Visible = True
    End If
    
    '��λ
    If mlngSeekHwnd <> 0 Then
        Call Form_Resize
        GetWindowRect mlngSeekHwnd, rParent
        If rParent.Top >= Me.Height / 15 Then
            Me.Top = rParent.Bottom * 15 - Me.Height + 30
        Else
            Me.Top = (rParent.Bottom - rParent.Top) * 15 + 30
        End If
        If rParent.Left >= Me.Width / 15 Then
            Me.Left = rParent.Right * 15 - Me.Width + 30
        Else
            Me.Left = (rParent.Right - rParent.Left) * 15 + 30
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim lngTVW As Long
    lngTVW = IIf(tvw_s.Visible, tvw_s.Width + pic.Width, 0)
    
    tvw_s.Left = Me.ScaleLeft
    tvw_s.Top = picInfo.Top + picInfo.Height + 15
    tvw_s.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height - 15
    
    pic.Left = tvw_s.Left + tvw_s.Width
    pic.Top = tvw_s.Top
    pic.Height = tvw_s.Height
    
    lvw.Left = Me.ScaleLeft + lngTVW
    lvw.Top = tvw_s.Top
    lvw.Height = tvw_s.Height
    lvw.Width = Me.ScaleWidth - lngTVW
    
    msh.Left = lvw.Left
    msh.Top = lvw.Top
    msh.Width = lvw.Width
    msh.Height = lvw.Height
    
    lbl.Left = lvw.Left
    lbl.Top = lvw.Top
    lbl.Width = lvw.Width
    lbl.Height = lvw.Height
    
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrMatch = ""
    mlngSeekHwnd = 0
    If blnSave Then SaveWinState Me, App.ProductName, mstrParName
    If Not InDesign Then Call SetWindowLong(hwnd, GWL_WNDPROC, glngSelProc)
End Sub

Private Sub lvw_DblClick()
    If blnItem Then Call cmdOK_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnItem = True
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOK_Click
End Sub

Private Sub lvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnItem = False
End Sub

Private Sub msh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If msh.MouseRow = 0 Then
        msh.MousePointer = 99
    Else
        msh.MousePointer = 0
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or lvw.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        lvw.Left = lvw.Left + X
        lvw.Width = lvw.Width - X
        msh.Left = msh.Left + X
        msh.Width = msh.Width - X
        
        lbl.Left = lbl.Left + X
        lbl.Width = lbl.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Index = intPreNode Then Exit Sub
    intPreNode = Node.Index
    DoEvents
    Call FillList(Node.Tag)
End Sub

Private Function FillTree() As Boolean
'���ܣ����ݶ�������Դ���ֶ����ԣ�������������ʾ��TreeView��
'���أ������Ƿ�ɹ�(�û�����������)
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, objNode As Node
    Dim strSel As String, strRela As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSel = GetScript(mstrFLDTree, "&S")
    strRela = GetScript(mstrFLDTree, "&R")
    
    If strSel = "" Or strRela = "" Then
        MsgBox "δ��������ѡ�������ϸ�б���������ֶ���Ŀ��", vbInformation, App.Title
        Exit Function
    End If
    strSQL = RemoveNote(mstrSQLTree)
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "FillTree")
    
    tvw_s.Nodes.Clear
        
    If Not rsTmp.EOF Then
        If InStr("|" & UCase(mstrFLDTree), "|ID,") > 0 And InStr("|" & UCase(mstrFLDTree), "|�ϼ�ID,") > 0 Then
            '���������б���ʾ
            Set objNode = tvw_s.Nodes.Add(, , "ALL", "������Ŀ", 1)
            objNode.Tag = "ALL"
            objNode.Expanded = True
            
            For i = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!�ϼ�ID) Then
                    Set objNode = tvw_s.Nodes.Add("ALL", 4, "_" & rsTmp!ID, IIf(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
                Else
                    Set objNode = tvw_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, IIf(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
                End If
                objNode.Tag = IIf(IsNull(rsTmp.Fields(strRela).Value), "", rsTmp.Fields(strRela).Value)
                rsTmp.MoveNext
            Next
        Else
            '����һ���б���ʾ
            For i = 1 To rsTmp.RecordCount
                Set objNode = tvw_s.Nodes.Add(, , , IIf(IsNull(rsTmp.Fields(strSel).Value), "", rsTmp.Fields(strSel).Value), 1)
                objNode.Tag = IIf(IsNull(rsTmp.Fields(strRela).Value), "", rsTmp.Fields(strRela).Value)
                rsTmp.MoveNext
            Next
        End If
    End If
    FillTree = True
    Exit Function
errH:
    If Err.Number = 35601 Then
        MsgBox "�����������������б�����ѡ��������ʹ�ã�", vbExclamation, App.Title
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function

Private Function GetRelaSQL(ByVal strSQL As String, ByVal strFld As String, ByVal strKey As String) As String
'���ܣ����������SQL
    Dim i As Integer, strRela As String
    
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), "&R") > 0 Then
            strRela = Split(Split(strFld, "|")(i), ",")(0)
            If strKey = "" Then
                GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & " is NULL"
            Else
                Select Case Split(Split(strFld, "|")(i), ",")(1)
                    Case adNumeric, adVarNumeric
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=" & strKey
                    Case adChar, adVarChar
                        GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "='" & strKey & "'"
                    Case adDBTimeStamp
                        If Format(strKey, "hh:mm:ss") = "00:00:00" Then
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & ">=To_Date('" & Format(strKey, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS') And " & strRela & "<=To_Date('" & Format(strKey, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            GetRelaSQL = "Select * From (" & strSQL & ") A Where " & strRela & "=To_Date('" & Format(strKey, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        End If
                End Select
            End If
            Exit Function
        End If
    Next
End Function

Private Function GetScript(strFld As String, strType As String) As String
'���ܣ�����ָ�����ֶ����������ֶ���
'������strType="&S &D &B &R"
'˵����������Ψһ�������ֶ�(����ֶ�)
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If InStr(Split(strFld, "|")(i), strType) > 0 Then
            GetScript = Split(Split(strFld, "|")(i), ",")(0)
            Exit Function
        End If
    Next
End Function

Private Function HaveScript(strFld As String, StrName As String, strType As String) As Boolean
'���ܣ��ж����ֶ������У�ָ�����ֶ��Ƿ����ָ������������
'������strName=�ֶ���,strFld=�ֶ�������,strType="&S &D &B &R"
'���أ�False=δ�����ֶλ��ֶβ�����ָ������
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If Split(Split(strFld, "|")(i), ",")(0) = StrName Then
            If InStr(Split(Split(strFld, "|")(i), ",")(2), strType) > 0 Then
                HaveScript = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function FillList(Optional strKey As String, Optional blnSort As Boolean) As Boolean
'���ܣ����ݵ�ǰѡ��ķ�������޷���ʱ�����Ӧ����ϸ�б�
'������strKey=�����б��еĵ�ǰ����ֵ
'˵���������������Ķ��٣�ȷ����ListView����DataGrid
    Dim strSQL As String, i As Long, j As Integer
    Dim objItem As ListItem, strValue As String
    Dim strDisp As String, strBand As String
    
    On Error GoTo errH
    
    lvw.ListItems.Clear
    
    lvw.Visible = False
    msh.Visible = False
    strList = ""
    msh.Clear
        
    '����Ϊֻ��������
    If Not blnSort Then
        If mstrSQLTree = "" Then
            strSQL = mstrSQLList
        Else
            '��̬����ϸ���ݴ���Ϊֻ��ȡ�����ķ��ಿ��(���� Order by �Ӿ�)
            If strKey = "ALL" Then
                strSQL = mstrSQLList
            Else
                strSQL = GetRelaSQL(RemoveOrderBy(mstrSQLList), mstrFLDList, strKey)
            End If
            
            If strSQL = "" Then
                MsgBox "�������ݶ�ȡʧ�ܣ�", vbInformation, App.Title
                Exit Function
            End If
        End If
        
        Set rsList = New ADODB.Recordset
        rsList.CursorLocation = adUseClient
        Screen.MousePointer = 11
        Me.Refresh
        strSQL = RemoveNote(strSQL)
        Set rsList = zldatabase.OpenSQLRecord(strSQL, "FillList")
        
    End If
    
    If Not rsList.EOF Then
        If rsList.RecordCount <= 500 Then
            If lvw.ColumnHeaders.Count = 0 Then Call AddListCols
            
            strDisp = GetScript(mstrFLDList, "&D") '��ʾֵ��Ŀ
            strBand = GetScript(mstrFLDList, "&B") '��ֵ��Ŀ
            
            For i = 1 To rsList.RecordCount
                strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(1).Text))
                If lvw.ColumnHeaders(1).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(1).Tag)
                Set objItem = lvw.ListItems.Add(, , strValue, , 1)
                For j = 2 To lvw.ColumnHeaders.Count
                    strValue = GetValue(rsList.Fields(lvw.ColumnHeaders(j).Text))
                    If lvw.ColumnHeaders(j).Tag <> "" Then strValue = Format(strValue, lvw.ColumnHeaders(j).Tag)
                    objItem.SubItems(j - 1) = strValue
                Next
                
                '����ʾֵ����ֵ������TAG��,��Ϊ��һ����Щ�ֶλ�Ϊѡ���ֶ�
                '��ʽΪ"��ʾֵ|��ֵ"
                If strDisp <> "" Then
                    objItem.Tag = IIf(IsNull(rsList.Fields(strDisp).Value), "", rsList.Fields(strDisp).Value)
                End If
                objItem.Tag = objItem.Tag & "|"
                If strBand <> "" Then
                    objItem.Tag = objItem.Tag & IIf(IsNull(rsList.Fields(strBand).Value), "", rsList.Fields(strBand).Value)
                End If
                                
                rsList.MoveNext
            Next
            
            '�Զ������п�
            Call AutoSizeCol(lvw)
            
            If Not Visible Or Not blnSetLvw Then
                Call RestoreListViewState(lvw, App.ProductName & "\" & Me.Name & mstrParName)
                blnSetLvw = True
            End If
            lvw.Visible = True
            strList = "lvw"
        Else
            msh.Redraw = False
            msh.Clear
            Set msh.DataSource = rsList
           
            For i = 0 To msh.Cols - 1
                'ɾ������ʾ����(&S)
                If Not HaveScript(mstrFLDList, msh.TextMatrix(0, i), "&S") Then
                    msh.ColWidth(i) = 0
                Else
                    '�����ж���
                    Select Case rsList.Fields(msh.TextMatrix(0, i)).Type
                        Case adNumeric, adVarNumeric
                            If rsList.Fields(msh.TextMatrix(0, i)).NumericScale > 0 Then
                                j = rsList.Fields(msh.TextMatrix(0, i)).NumericScale
                                msh.ColAlignment(i) = 7
                            Else
                                If rsList.Fields(msh.TextMatrix(0, i)).Precision < 3 Then
                                    msh.ColAlignment(i) = 4
                                Else
                                    msh.ColAlignment(i) = 1
                                End If
                            End If
                        Case adDBTimeStamp
                            msh.ColAlignment(i) = 4
                        Case Else
                            msh.ColAlignment(i) = 1
                    End Select
                    If msh.TextMatrix(0, i) Like "*��λ*" Then msh.ColAlignment(i) = 4
                    If msh.TextMatrix(0, i) Like "*��*" Then msh.ColAlignment(i) = 4
                End If
            Next
            '�����п��
            Call SetColWidth(msh, Me)
            
            msh.Col = 0: msh.ColSel = msh.Cols - 1
            If Not Visible Or Not blnSetFlex Then
                blnSetFlex = True
                RestoreFlexState msh, App.ProductName & "\" & Me.Name & mstrParName
            End If
            msh.Redraw = True
            msh.Visible = True
            strList = "msh"
        End If
        lblInfo.Caption = "�� " & rsList.RecordCount & " ����ϸ��Ŀ."
    Else
        'û������ʱ����ʾ�յ�ListView(����ͷ)
        If lvw.ColumnHeaders.Count = 0 Then Call AddListCols
        lvw.Visible = True
        strList = "lvw"
        lblInfo.Caption = "û����ϸ��Ŀ."
    End If
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Me.Refresh
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AddListCols()
'���ܣ�����mstrFLDList�ֶ�����ֵ,ΪListView������ͷ
    Dim i As Integer, j As Integer, strFld As String
    Dim objCol As ColumnHeader
    
    For i = 0 To UBound(Split(mstrFLDList, "|"))
        strFld = Split(mstrFLDList, "|")(i)
        If strFld Like "*&S*" Then
            Set objCol = lvw.ColumnHeaders.Add(, "_" & Split(strFld, ",")(0), Split(strFld, ",")(0))
            
            objCol.Width = Me.TextWidth(Split(strFld, ",")(0) & "��")
            
            '�����ֶ������������ö���(��1ֻ�������)
            Select Case Split(strFld, ",")(1)
                Case adNumeric, adVarNumeric
                    If rsList.Fields(objCol.Text).NumericScale > 0 Then
                        j = rsList.Fields(objCol.Text).NumericScale
                        objCol.Tag = "0." & String(IIf(j > 2, 2, j), "0; ;")
                        If objCol.Index <> 1 Then objCol.Alignment = lvwColumnRight
                    ElseIf objCol.Index <> 1 Then
                        If rsList.Fields(objCol.Text).Precision < 3 Then
                            objCol.Alignment = lvwColumnCenter
                        Else
                            objCol.Alignment = lvwColumnLeft
                        End If
                    End If
                    If objCol.Text Like "*��" Then objCol.Tag = "0.000"
                    If objCol.Text Like "*��" Then objCol.Tag = "0.00"
                Case adDBTimeStamp
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
                Case Else
                    If objCol.Index <> 1 Then objCol.Alignment = lvwColumnLeft
            End Select
            If objCol.Text Like "*��λ*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
            If objCol.Text Like "*��*" And objCol.Index <> 1 Then objCol.Alignment = lvwColumnCenter
        End If
    Next
End Sub

Private Function GetValue(objFld As Field) As String
'����:�����ֶ�����ȡ���ʵ���ʾֵ
    Dim strValue As String
    Select Case objFld.Type
        Case adChar, adVarChar, adLongVarChar
            strValue = IIf(IsNull(objFld.Value), "", objFld.Value)
        Case adNumeric, adVarNumeric
            strValue = IIf(IsNull(objFld.Value), 0, objFld.Value)
        Case adDBTimeStamp
            strValue = IIf(IsNull(objFld.Value), "", objFld.Value)
            If Format(strValue, "HH:mm:ss") = "00:00:00" Then
                strValue = Format(strValue, "yyyy-MM-dd")
            Else
                strValue = Format(strValue, "yyyy-MM-dd HH:mm:ss")
            End If
        Case Else
            strValue = IIf(IsNull(objFld.Value), "", objFld.Value)
    End Select
    GetValue = strValue
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'���ܣ���������
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub msh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    If msh.MouseRow <> 0 Then Exit Sub
    
    lngCol = msh.MouseCol
    
    If Button = 1 And msh.MousePointer = 99 Then
        If msh.TextMatrix(0, lngCol) = "" Then Exit Sub
        If rsList Is Nothing Then Exit Sub
        If rsList.State = 0 Then Exit Sub
        
        Set msh.DataSource = Nothing

        rsList.Sort = msh.TextMatrix(0, lngCol) & IIf(msh.ColData(lngCol) = 0, "", " DESC")
        msh.ColData(lngCol) = (msh.ColData(lngCol) + 1) Mod 2
        
        Call FillList(, True)
    End If
End Sub
