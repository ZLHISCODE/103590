VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildPatient 
   BorderStyle     =   0  'None
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   630
      ScaleHeight     =   1935
      ScaleWidth      =   3000
      TabIndex        =   8
      Top             =   3465
      Width           =   3000
      Begin MSComctlLib.TreeView tvw 
         Height          =   1635
         Left            =   300
         TabIndex        =   9
         Top             =   210
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2884
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2625
      Index           =   0
      Left            =   750
      ScaleHeight     =   2625
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   510
      Width           =   3000
      Begin VB.CheckBox chkNCommit 
         Caption         =   "�����Ѿ���Ժ������δ�ύ����"
         Height          =   225
         Left            =   60
         TabIndex        =   13
         Top             =   660
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.ComboBox cboStatus 
         Height          =   300
         Left            =   825
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   345
         Width           =   2130
      End
      Begin VB.PictureBox picSelect 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   135
         ScaleHeight     =   435
         ScaleWidth      =   3720
         TabIndex        =   2
         Top             =   2175
         Width           =   3720
         Begin VB.CommandButton cmdStatus 
            Cancel          =   -1  'True
            Height          =   285
            Left            =   2535
            Picture         =   "frmChildPatient.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "���·�ѡ�е��ļ�"
            Top             =   90
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.Label labStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ѡ��      ����      ��"
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   105
            Width           =   1980
         End
         Begin VB.Label labNum 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   1515
            TabIndex        =   4
            Top             =   90
            Width           =   345
         End
         Begin VB.Label labSelect 
            Alignment       =   2  'Center
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   450
            TabIndex        =   3
            Top             =   90
            Width           =   570
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Left            =   150
         TabIndex        =   6
         Top             =   930
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
         Appearance      =   2
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12698049
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
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
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   825
         TabIndex        =   1
         Top             =   30
         Width           =   2130
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״̬"
         Height          =   180
         Left            =   45
         TabIndex        =   11
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         Height          =   180
         Left            =   45
         TabIndex        =   7
         Top             =   75
         Width           =   720
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChildPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Const CB_GETDROPPEDSTATE = &H157

Private mfrmMain As Object
Private mlngKey As Long
Private mlngReferKey As Long
Private mblnReading As Boolean
Private mstrSQL As String
Private mblnDataChanged As Boolean
Private mbytApplyMode   As Byte               '1-��Ժ���ύ����;3-��Ժ����;4-����׼�Ľ��Ĳ���
Private mbytMode        As Byte
Private mrsCondition    As ADODB.Recordset
Private mclsVsf      As clsVsf
Private mlng����ID      As Long
Private mlng��ҳID      As Long
Private mstrKey         As String
Private mstrSvr�������� As String
Private mblnDrop        As Boolean
Private mrsDept         As ADODB.Recordset
Private mrsData         As ADODB.Recordset
Private mstrPrivs       As String
Private mstrDepts       As String
Private blnReadUsed     As Boolean
Private mblnRead�����ṹ As Boolean

Public Event AfterDeptChanged()
Public Event StatusChanged()
Public Event DbClick()
Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event AfterDocumentChanged(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng�ύId As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)

Public Property Get Depts() As String
    Depts = mstrDepts
End Property

Public Property Let Depts(ByVal vDepts As String)
    mstrDepts = vDepts
End Property
Public Sub cboDeptRefresh(strDeptName As String)
    On Error GoTo ErrH
    If cboDept.Text <> "" Then
        cboDept.ListIndex = GetCboIndex(cboDept, zlCommFun.GetNeedName(strDeptName))
    End If
    Call cboDept_Click
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function GetCboIndex(cbo As ComboBox, strFind As String, Optional blnKeep As Boolean, Optional blnLike As Boolean) As Long
    '���ܣ����ַ�����ComboBox�в�������
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '�Ⱦ�ȷ����
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), "-") > 0 Then
            If zlCommFun.GetNeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '���ģ������
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Function VsfBody() As VSFlexGrid
    Set VsfBody = vsf
End Function

Public Property Get Title() As String
    If Not tvw.SelectedItem Is Nothing Then
        Title = tvw.SelectedItem.Text
    End If
End Property

Public Function zlColumnSelect() As Boolean
    If frmTemplateColumn.ShowColumn(mfrmMain, mclsVsf) Then
        mclsVsf.AppendRows = True
    End If
End Function

Public Function zlLocationDocument(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal byt�������� As Byte, ByVal strFileKey As String)
    Dim strKey As String
    Dim intRow As Integer
    Dim objNode As Node
    
    '1-סԺҽ��;2-סԺ����;3-������;4-�����¼;5-��ҳ��¼;6-ҽ������;7-����֤��;8-֪���ļ�
    With vsf
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("����id"))) = lng����ID And Val(.TextMatrix(intRow, .ColIndex("��ҳid"))) = lng��ҳID Then
                
                .Row = intRow
                If IsNumeric(strFileKey) Then
                    strKey = "R" & byt��������
                    If strFileKey <> "" And strFileKey <> "0,0,0" Then strKey = strKey & "K" & strFileKey
                End If
                
                On Error Resume Next
                mblnRead�����ṹ = True
                Set objNode = tvw.Nodes(strKey)
                If Not (objNode Is Nothing) Then
                    objNode.EnsureVisible
                    objNode.Selected = True
                    If mblnRead�����ṹ Then
                        mblnRead�����ṹ = False
                        Call tvw_NodeClick(objNode)
                    End If
                    zlLocationDocument = True
                End If
                Exit Function
            End If
        Next
    End With
End Function

Public Function zlLocationPatient(Optional ByVal bytApplyMode As Byte = 1, Optional ByVal strFindKey As String, Optional ByVal strLocationText As String, Optional ByVal strNo As String, _
    Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long, Optional ByVal lng����ID As Long, Optional ByVal strFindDeal As String) As Boolean
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim intCol As Integer
    Dim intRow As Integer
    Dim bytMatch As Byte
    Dim intLoop As Integer
    Dim strCols As String
    Dim i           As Integer
    
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "ID", adVarChar, 30
        .Fields.Append "����", adVarChar, 100
        .Fields.Append "�Ա�", adVarChar, 10
        .Fields.Append "��Ժ����", adVarChar, 50
        .Fields.Append "����", adVarChar, 30
        .Fields.Append "סԺ��", adVarChar, 30
        .Fields.Append "��Ժ����ID", adDouble, 50
        .Fields.Append "��Ժ����", adDBTimeStamp, 20
        .Fields.Append "��Ժ����", adDBTimeStamp, 20
        .Open
    End With
    
    With vsf
        Select Case bytApplyMode
             Case 0
                
                If lng����ID > 0 And lng��ҳID > 0 Then
                    For intRow = 1 To .Rows - 1
                        
                        If .TextMatrix(intRow, .ColIndex("No")) = strNo Then
                            If Val(.TextMatrix(intRow, .ColIndex("����id"))) = lng����ID And Val(.TextMatrix(intRow, .ColIndex("��ҳid"))) = lng��ҳID Then
                                '�ҵ��˲���
                                .Row = intRow
                                .ShowCell .Row, .Col
                                GoTo endHand
                            End If
                        End If
                    Next
                End If
            Case 1
                
                If lng����ID > 0 And lng��ҳID > 0 Then
                    For intRow = 1 To .Rows - 1
                        If .TextMatrix(intRow, .ColIndex("No")) = strNo Then
                            If Val(.TextMatrix(intRow, .ColIndex("����id"))) = lng����ID And Val(.TextMatrix(intRow, .ColIndex("��ҳid"))) = lng��ҳID Then
                                '�ҵ��˲���
                                .Row = intRow
                                .ShowCell .Row, .Col
                                GoTo endHand
                            End If
                        End If
                    Next
                End If
            Case 2
                intRow = -1
                bytMatch = 2
                intCol = mclsVsf.ColIndex(strFindKey)
                '���ҳ������ȷŵ���¼���У�����ж���ʱ�������Ի������ѡ��ȷ��
                If intCol > 0 Then
                    If strFindKey = "����" Then
                        mrsData.Filter = "���� like '%" & strLocationText & "%'"
                    ElseIf strFindKey = "סԺ��" Then
                        If IsNumeric(strLocationText) Then
                          mrsData.Filter = strFindKey & "  = '" & strLocationText & "'"
                        End If
                    Else
                        mrsData.Filter = strFindKey & "  like '%" & strLocationText & "%'"
                    End If
                    Do While Not (mrsData.EOF Or mrsData.BOF)
                        '�ҵ���,��д
                        rsData.AddNew
                        rsData("ID").Value = mrsData!ID
                        rsData("����").Value = NVL(mrsData!����)
                        rsData("סԺ��").Value = NVL(mrsData!סԺ��)
                        rsData("��Ժ����").Value = NVL(mrsData!��Ժ����)
                        rsData("�Ա�").Value = NVL(mrsData!�Ա�)
                        rsData("����").Value = NVL(mrsData!����)
                        rsData("��Ժ����ID").Value = mrsData!��Ժ����ID
                        rsData("��Ժ����").Value = NVL(mrsData!��Ժ����, 0)
                        rsData("��Ժ����").Value = NVL(mrsData!��Ժ����, 0)
                        mrsData.MoveNext
                    Loop
                    rsData.Sort = "��Ժ����"
                    If rsData.RecordCount = 0 Then Exit Function
                    If rsData.RecordCount = 1 Then
                        Set rs = rsData
                    Else
                        rsData.MoveFirst
                        strCols = "����,900,0,;�Ա�,500,0,;����,600,0,;סԺ��,1200,0,;��Ժ����,1500,0,;��Ժ����,1600,0,;��Ժ����,1600,0,"
                        If ShowPubSelect(mfrmMain, mfrmMain.txtLocation, 2, strCols, "˽��ģ��\" & App.ProductName & "\" & Me.Name & "\��λ���Ҳ���" & mbytApplyMode, "����±���ѡ�������ҵĲ���", rsData, rs, 8790, 4500, , , , True) <> 1 Then
                            Exit Function
                        End If
                    End If
                    .SetFocus
                    DoEvents
                    '����Ҳ������ˣ��ҵ�ǰ���Ҳ������п��ң����ȡ��ǰ�����ڿ��ҵ�����
                    If Not (cboDept.Text = "���п���" Or cboDept.ItemData(cboDept.ListIndex) = rs!��Ժ����ID) Then
                        If mbytApplyMode <> 3 Then
                            If Val(labNum.Caption) > 0 Then
                                If MsgBox("����ѡ���ˡ�" & labNum.Caption & "���ݲ�������ǰ����������ˢ�����ݣ�" & vbCrLf & "ȷ��ִ�иò�����", vbOKCancel + vbQuestion + vbDefaultButton2, ParamInfo.��Ʒ����) = vbCancel Then
                                    Exit Function
                                End If
                            End If
                        End If
                        For i = 0 To cboDept.ListCount - 1
                            If cboDept.ItemData(i) = rs!��Ժ����ID Then
                                cboDept.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                    intRow = mclsVsf.FindRow(rs("ID").Value, .ColIndex("ID"), bytMatch, .Row + 1)
                    If intRow = -1 Then
                        intRow = mclsVsf.FindRow(rs("ID").Value, .ColIndex("ID"), bytMatch)
                    End If
                    .Row = intRow
                    .ShowCell .Row, .Col
                    DoEvents
                    GoTo endHand
                End If
                

            Case 3
                If mbytApplyMode = 3 And cboDept.ListIndex >= 0 Then
                    If cboDept.ItemData(cboDept.ListIndex) <> lng����ID And lng����ID > 0 Then
                        zlControl.CboLocate cboDept, lng����ID, True
                    End If
                End If
                
                If lng����ID > 0 And lng��ҳID > 0 Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("����id"))) = lng����ID And Val(.TextMatrix(intRow, .ColIndex("��ҳid"))) = lng��ҳID Then
                            '�ҵ��˲���
                            .Row = intRow
                            .ShowCell .Row, .Col
                            GoTo endHand
                        End If
                    Next
                End If
            End Select
    End With
    zlLocationPatient = False
    
    Exit Function
endHand:
    With vsf
        If bytApplyMode <> 3 And .ColIndex("ѡ��") >= 0 Then
        
            Select Case strFindDeal
            Case "���Ҳ�ѡ��"
                
                .TextMatrix(.Row, .ColIndex("ѡ��")) = 1
                
            Case "���Ҳ���ѡ"
                
                .TextMatrix(.Row, .ColIndex("ѡ��")) = 0
                    
            Case "���Ҳ���ѡ"
                
                If Abs(Val(.TextMatrix(.Row, .ColIndex("ѡ��")))) = 0 Then
                    .TextMatrix(.Row, .ColIndex("ѡ��")) = 1
                Else
                    .TextMatrix(.Row, .ColIndex("ѡ��")) = 0
                End If
                
            End Select
    
        End If
    End With
    
    zlLocationPatient = True
End Function

Public Function zlInitData(ByVal frmMain As Object, ByVal bytApplyMode As Byte, Optional ByVal strPrivs As String) As Boolean
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    mbytApplyMode = bytApplyMode
    If InitControl = False Then Exit Function
    If InitData = False Then Exit Function
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then 'ʹ�ø��Ի�����
        mclsVsf.LoadStateFromString Trim(GetRegister(˽��ģ��, Me.Name, "������_20100719_" & mbytApplyMode, ""))
    End If
    zlInitData = True

End Function

Public Function zlRefreshData(Optional ByVal rsCondition As ADODB.Recordset, Optional ByVal lng����ID As Long, Optional ByVal lng��ҳID As Long) As Boolean
    Set mrsCondition = rsCondition
    mbytMode = 2
    mstrKey = ""
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    
    If ExecuteCommand("ˢ������", 0) = False Then Exit Function
    If mbytApplyMode <> 4 Then
        Call cboDept_Click
    End If
    
    zlRefreshData = True
    
End Function

Public Function zlRefreshStruct() As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ�µ�ǰ�ĵ���
    '������
    '���أ�
    '******************************************************************************************************************
     zlRefreshStruct = ExecuteCommand("��ȡ�����ṹ", "Read")
    
End Function

Public Function zlShowDocument() As Boolean
    mstrKey = ""
    Call ExecuteCommand("��ȡ�����ṹ", "Read")
    
End Function

'######################################################################################################################

Private Function GetParamRecord(ByVal strParam As String) As ADODB.Recordset
    Dim varTmp As Variant
    Dim varAry As Variant
    Dim rs As New ADODB.Recordset
    Dim intCount As Integer
    Dim intCol As Integer
    
    If strParam <> "" Then

        '���������ҷ�Χ����
        With rs
            .Fields.Append "��Աid", adBigInt
            .Fields.Append "����id", adBigInt
            .Open
        End With
        
        varTmp = Split(strParam, ";")
                
         For intCount = 0 To UBound(varTmp)

            If varTmp(intCount) <> "" Then
                varAry = Split(varTmp(intCount), ",")
                For intCol = 1 To UBound(varAry)
                    rs.AddNew
                    rs("��Աid").Value = Val(varAry(0))
                    rs("����id").Value = Val(varAry(intCol))
                Next
                
            End If
         Next
         
         If rs.RecordCount > 0 Then rs.MoveFirst
    End If
    
    Set GetParamRecord = rs
End Function
Private Function InitControl() As Boolean

    mblnReading = True

    Set mclsVsf = New clsVsf
    With mclsVsf

        Select Case mbytApplyMode
            Case 1               '��鲡��
            
                Call .Initialize(Me.Controls, vsf, True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("����״ֵ̬", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("����ת��", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("���￨��", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("����", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("���ʱ��", 0, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True, , , True)
                
                Call .AppendColumn("", 240, flexAlignCenterCenter, flexDTBoolean, , "[ѡ��]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[ͼ��]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[·��]", False)
                Call .AppendColumn("����", 810, flexAlignLeftCenter, flexDTString, , , True)

                Call .AppendColumn("סԺ��", 900, flexAlignLeftCenter, flexDTDecimal, , , True)
                
                Call .AppendColumn("����", 500, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("����ȼ�", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("סԺҽʦ", 810, flexAlignLeftCenter, flexDTDecimal, , , True)

                Call .AppendColumn("��Ժ����", 1080, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("����״̬", 840, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("�ύ��", 810, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("�ύʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("������", 810, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("��������", 600, flexAlignLeftCenter, flexDTString, , , True)
                
                Call .AppendColumn("��Ժ����ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("��������", 0, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("�������", 0, flexAlignLeftCenter, flexDTString, "", , True) '=1��� =0δ���
                
                Call .AppendColumn("��Ժ����", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("��Ժ����", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("�Ա�", 450, flexAlignLeftCenter, flexDTString, "", , True)
                
                .SysHidden(.ColIndex("ID")) = True
                .SysHidden(.ColIndex("����id")) = True
                .SysHidden(.ColIndex("��ҳid")) = True
                .SysHidden(.ColIndex("����״ֵ̬")) = True
                .SysHidden(.ColIndex("���￨��")) = True
                .SysHidden(.ColIndex("����")) = True
                .SysHidden(.ColIndex("���ʱ��")) = True
                .SysHidden(.ColIndex("����ת��")) = True
                .SysHidden(.ColIndex("��Ժ����ID")) = True
                .SysHidden(.ColIndex("��������")) = True
                .SysHidden(.ColIndex("�������")) = True
                .SysHidden(.ColIndex("�Ա�")) = True
                .SysHidden(.ColIndex("��Ժ����")) = True
                .SysHidden(.ColIndex("��Ժ����")) = True
                
                Call .InitializeEdit(True, False, False)
                Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
            Case 3                  '��Ժ����
                
                Call .Initialize(Me.Controls, vsf, True, False, frmPubResource.GetImageList(16))
                Call .ClearColumn
                
                Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("���￨��", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("����״ֵ̬", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("����ת��", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("���ʱ��", 0, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", , True, , , True)
                                
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[·��]", False)
                Call .AppendColumn("����", 840, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("�Ա�", 450, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("����", 600, flexAlignRightCenter, flexDTDecimal, "", , True)
                Call .AppendColumn("סԺ��", 900, flexAlignLeftCenter, flexDTDecimal, "", , True)
                
                Call .AppendColumn("����", 500, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("����ȼ�", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("סԺҽʦ", 810, flexAlignLeftCenter, flexDTDecimal, , , True)

                Call .AppendColumn("���״̬", 840, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("��Ժ����", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("��Ժ����", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                
                Call .AppendColumn("��Ժ����ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
    
                
                .SysHidden(.ColIndex("ID")) = True
                .SysHidden(.ColIndex("����id")) = True
                .SysHidden(.ColIndex("��ҳid")) = True
                .SysHidden(.ColIndex("����״ֵ̬")) = True
                .SysHidden(.ColIndex("���￨��")) = True
                .SysHidden(.ColIndex("���ʱ��")) = True
                .SysHidden(.ColIndex("����ת��")) = True
                .SysHidden(.ColIndex("��Ժ����ID")) = True
                
                
            Case 4                  '����׼�Ľ��Ĳ���
            
                Call .Initialize(Me.Controls, vsf, True, False, frmPubResource.GetImageList(16))
                Call .ClearColumn
                
                Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("��ҳid", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("���￨��", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("����ת��", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[·��]", False)
                Call .AppendColumn("����", 840, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("�Ա�", 450, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("סԺ��", 900, flexAlignLeftCenter, flexDTDecimal, "", , True)
                
                Call .AppendColumn("No", 900, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("������", 840, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("����ʱ��", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("��������", 600, flexAlignLeftCenter, flexDTString, "", , True)
                
                Call .AppendColumn("��Ժ����ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
                
                .SysHidden(.ColIndex("ID")) = True
                .SysHidden(.ColIndex("����id")) = True
                .SysHidden(.ColIndex("��ҳid")) = True
                .SysHidden(.ColIndex("���￨��")) = True
                .SysHidden(.ColIndex("����ת��")) = True
                .SysHidden(.ColIndex("��Ժ����ID")) = True
            
        End Select
        .AppendRows = True
    End With
    DoEvents
     
    '����ͣ������

    Dim objPane As Pane
    Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing): objPane.Title = "�����б�": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 200, DockBottomOf, objPane): objPane.Title = "���Ӳ���": objPane.Options = PaneNoCaption

    Call DockPannelInit(dkpMain)
    

    Select Case mbytApplyMode
        Case 3
            lbl.Caption = "סԺ����"
        Case Else
            lbl.Caption = "��Ժ����"
    End Select


    
    mblnReading = False
    InitControl = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function InitData(Optional ByVal btnShowDept As Boolean = False) As Boolean
Dim strTmp As String
Dim rs As New ADODB.Recordset
Dim rsParam As New ADODB.Recordset
    
    mblnReading = True
    Set mrsDept = New ADODB.Recordset
    With mrsDept
        .Fields.Append "����", adVarChar, 20
        .Fields.Append "����", adVarChar, 100
        .Fields.Append "����", adVarChar, 30
        .Open
    End With
    
    cboDept.Clear
    Select Case mbytApplyMode
        Case 4
            cboDept.AddItem "���п���"
        Case Else
            If IsPrivs(mstrPrivs, "���п���") Then cboDept.AddItem "���п���"
    End Select
    

    If mbytApplyMode = 4 Then
        Set rs = gclsPackage.GetDept("�ٴ�", , btnShowDept)
    Else
        Set rs = gclsPackage.GetDept("�ٴ�", , IsPrivs(mstrPrivs, "���п���"), btnShowDept)
    End If
    
    If rs.BOF = False Then
        
        strTmp = Trim(zlDatabase.GetPara("�����ҷ�Χ", ParamInfo.ϵͳ��, mfrmMain.ģ���))
        
        If strTmp = "" Then
            'û���������ҷ�Χ����
            Do While Not rs.EOF
                cboDept.AddItem rs("��ʾ����").Value
                cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                mstrDepts = mstrDepts & rs("ID").Value & ","
                mrsDept.AddNew
                mrsDept("����").Value = rs("����").Value
                mrsDept("����").Value = rs("����").Value
                mrsDept("����").Value = rs("����").Value & ""
                
                rs.MoveNext
            Loop
            mstrDepts = Left(mstrDepts, Len(mstrDepts) - 1)
        Else
            '���������ҷ�Χ����
            Set rsParam = GetParamRecord(strTmp)
            rsParam.Filter = ""
            rsParam.Filter = "��Աid=" & UserInfo.ID
            If rsParam.RecordCount > 0 Then
                cboDept.Clear
                
                Do While Not rs.EOF
                    rsParam.Filter = ""
                    rsParam.Filter = "��Աid=" & UserInfo.ID & " And ����id=" & Val(rs("ID").Value)
                    If rsParam.RecordCount > 0 Then
                        cboDept.AddItem rs("��ʾ����").Value
                        cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                        mstrDepts = mstrDepts & rs("ID").Value & ","
                        
                        mrsDept.AddNew
                        mrsDept("����").Value = rs("����").Value
                        mrsDept("����").Value = rs("����").Value
                        mrsDept("����").Value = rs("����").Value & ""
                    End If
                    rs.MoveNext
                Loop
                mstrDepts = Left(mstrDepts, Len(mstrDepts) - 1)
            Else

                Do While Not rs.EOF
                    cboDept.AddItem rs("��ʾ����").Value
                    cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                    mstrDepts = mstrDepts & rs("ID").Value & ","
                    mrsDept.AddNew
                    mrsDept("����").Value = rs("����").Value
                    mrsDept("����").Value = rs("����").Value
                    mrsDept("����").Value = rs("����").Value & ""
                    
                    rs.MoveNext
                Loop
                mstrDepts = Left(mstrDepts, Len(mstrDepts) - 1)
            End If
            
        End If
    End If
    
    If cboDept.ListCount = 0 Then
        cboDept.AddItem ""
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    
    cboDept.ListIndex = 0
    mstrSvr�������� = cboDept.Text

    '��Ӳ���״̬
    cboStatus.Clear
    If mbytApplyMode = 1 Then
        cboStatus.AddItem "����״̬"
        cboStatus.ItemData(cboStatus.NewIndex) = 0
        cboStatus.AddItem "�ύ����"
        cboStatus.ItemData(cboStatus.NewIndex) = 1
        cboStatus.AddItem "���մ���"
        cboStatus.ItemData(cboStatus.NewIndex) = 10
        cboStatus.AddItem "�������"
        cboStatus.ItemData(cboStatus.NewIndex) = 3
        cboStatus.AddItem "��鷴��"
        cboStatus.ItemData(cboStatus.NewIndex) = 4
        cboStatus.AddItem "�������"
        cboStatus.ItemData(cboStatus.NewIndex) = 6
        cboStatus.AddItem "���鵵"
        cboStatus.ItemData(cboStatus.NewIndex) = 5
        cboStatus.ListIndex = 0
    End If
    
    mblnReading = False
    InitData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
Dim rs As New ADODB.Recordset
Dim lngLoop As Long
Dim intRow As Integer
    
    On Error GoTo errHand
    
    mblnReading = True
    Select Case strCmd
    Case "��ȡ��Ժ����"
        If mlng����ID = 0 Then
            mclsVsf.SaveKey = vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID"))
            mclsVsf.ClearGrid
            Set rs = gclsPackage.GetAduitPatient(3, ParamRead(mrsCondition, "��鿪ʼʱ��"), ParamRead(mrsCondition, "������ʱ��"), _
                                                Val(ParamRead(mrsCondition, "���մ���")) & ";" & Val(ParamRead(mrsCondition, "�ܾ�����")) & ";" & Val(ParamRead(mrsCondition, "�������")) & ";" & Val(ParamRead(mrsCondition, "��鷴��")) & ";" & Val(ParamRead(mrsCondition, "�������")) & ";" & Val(ParamRead(mrsCondition, "�ύ����")), 0, 0, _
                                                ParamRead(mrsCondition, "��Ժ���"), Val(ParamRead(mrsCondition, "��������")), ParamRead(mrsCondition, "ҽ������"), _
                                                ParamRead(mrsCondition, "סԺҽʦ"), ParamRead(mrsCondition, "��������"), ParamRead(mrsCondition, "�������"), _
                                                ParamRead(mrsCondition, "ҩƷ��Ϣ"), ParamRead(mrsCondition, "ҽ����ʼʱ��"), ParamRead(mrsCondition, "ҽ������ʱ��") _
                                                )
            If rs.BOF = False Then
                Call mclsVsf.LoadDataSource(rs)
            End If
        Else
            
            Set rs = gclsPackage.GetAduitPatient(3, ParamRead(mrsCondition, "��鿪ʼʱ��"), ParamRead(mrsCondition, "������ʱ��"), _
                                                Val(ParamRead(mrsCondition, "���մ���")) & ";" & Val(ParamRead(mrsCondition, "�ܾ�����")) & ";" & Val(ParamRead(mrsCondition, "�������")) & ";" & Val(ParamRead(mrsCondition, "��鷴��")) & ";" & Val(ParamRead(mrsCondition, "�������")) & ";" & Val(ParamRead(mrsCondition, "�ύ����")), _
                                                mlng����ID, mlng��ҳID, ParamRead(mrsCondition, "��Ժ���"), Val(ParamRead(mrsCondition, "��������")), ParamRead(mrsCondition, "ҽ������"), _
                                                ParamRead(mrsCondition, "סԺҽʦ"), ParamRead(mrsCondition, "��������"), ParamRead(mrsCondition, "�������"), _
                                                ParamRead(mrsCondition, "ҩƷ��Ϣ"), ParamRead(mrsCondition, "ҽ����ʼʱ��"), ParamRead(mrsCondition, "ҽ������ʱ��") _
                                                )
            If rs.BOF = False Then
                intRow = 0
                For lngLoop = 1 To vsf.Rows - 1
                    If Val(vsf.TextMatrix(lngLoop, vsf.ColIndex("����id"))) = mlng����ID And Val(vsf.TextMatrix(lngLoop, vsf.ColIndex("��ҳid"))) = mlng��ҳID Then
                        intRow = lngLoop
                        Exit For
                    End If
                Next
                If intRow > 0 Then
                    '�Ѽ���
                    vsf.Row = intRow
                    Call mclsVsf.LoadGridRow(vsf.Row, rs)
                End If
            End If
        End If
        Set mrsData = rs
    Case "��ȡ�����ṹ"
        
        Dim objNode As Node
        Dim strIcon As String
        Dim strKey As String
        
        If Not (tvw.SelectedItem Is Nothing) Then strKey = tvw.SelectedItem.Key
        If InStr(strKey, "K") = 0 And strKey <> "R1" And strKey <> "R5" Then strKey = ""
        
        LockWindowUpdate tvw.hWnd
        
        tvw.Nodes.Clear
                
        With vsf
            Set rs = gclsPackage.GetCISStruct(Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), Val(.TextMatrix(.Row, .ColIndex("��Ժ����ID"))), Val(.TextMatrix(.Row, .ColIndex("����ת��"))) = 1)
        End With
                
        If rs.BOF = False Then
            
            '���棺Decode(a.ҽ��id, Null, a.����, '<'||b.ҽ������||'>' || a.���� || '(' || To_Char(b.��ʼִ��ʱ��, 'yyyy-mm-dd') || ')') As ����, Trim(To_Char(a.ID))||';'||Decode(a.ҽ��id,Null,'0',Trim(To_Char(a.ҽ��id))) As ����
            '����A.���� || '(' || B.���� || '��' || To_Char(A.��ʼ, 'yyyy-mm-dd hh24:mi') || ' �� ' ||To_Char(A.��ֹ, 'yyyy-mm-dd hh24:mi') || ')' As ����, Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(����,0)))||';'||To_Char(A.��ʼ, 'yyyy-mm-dd hh24:mi')||' �� '||To_Char(A.��ֹ, 'yyyy-mm-dd hh24:mi')||';'||Trim(To_Char(A.ID)) As ����
            
            Do While Not rs.EOF
                strIcon = zlCommFun.NVL(rs("ͼ��").Value)

                If zlCommFun.NVL(rs("�ϼ�id").Value) = "" Then
                    Set objNode = tvw.Nodes.Add(, , rs("ID").Value, rs("����").Value, strIcon, strIcon)
                    objNode.Tag = zlCommFun.NVL(rs("����").Value)
                Else
                    Set objNode = tvw.Nodes.Add(rs("�ϼ�id").Value, tvwChild, rs("ID").Value, rs("����").Value, strIcon, strIcon)
                    objNode.Tag = zlCommFun.NVL(rs("����").Value)
                End If
            
                rs.MoveNext
            Loop
        End If
        
        With vsf
            Set rs = New ADODB.Recordset
            Set rs = gclsPackage.GetEmrCISStruct(Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))))
        End With
        If Not rs Is Nothing Then
            If rs.State = ADODB.adStateOpen Then
            If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do Until rs.EOF
                Set objNode = tvw.Nodes.Add(rs!�ϼ�ID.Value, tvwChild, rs!ID.Value, rs!����.Value, rs!ͼ��.Value, rs!ͼ��.Value)
                objNode.Tag = NVL(rs!����) '�ĵ�ID[|���ĵ�ID]
                rs.MoveNext
            Loop
            End If
            End If
        End If

        If tvw.Nodes.count > 0 Then
            
            On Error Resume Next
            Err = 0
            If strKey <> "" Then tvw.Nodes(strKey).Selected = True
            On Error GoTo errHand
            
            If Err <> 0 Or strKey = "" Or tvw.SelectedItem Is Nothing Then
                tvw.Nodes(1).Selected = True
            End If
            
            If Not (tvw.SelectedItem Is Nothing) Then
                If varParam(0) = "NoRead" Then
                    If mblnRead�����ṹ Then
                        mblnRead�����ṹ = False
                        With vsf
                            If mbytApplyMode = 3 Then
                                RaiseEvent AfterDocumentChanged(0, 0, "��ҳ��¼", "", "", 0, Val(.TextMatrix(.Row, .ColIndex("����ת��"))) = 1, False)
                            Else
                                RaiseEvent AfterDocumentChanged(0, 0, "��ҳ��¼", "", "", IIf(vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")) = "", 0, vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID"))), Val(.TextMatrix(.Row, .ColIndex("����ת��"))) = 1, False)
                            End If
                        End With
                    End If
                Else
                    mblnRead�����ṹ = True
                    Call tvw_NodeClick(tvw.SelectedItem)
                End If
            End If
        Else
            With vsf
                RaiseEvent AfterDocumentChanged(0, 0, "��ҳ��¼", "", "", vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")), Val(.TextMatrix(.Row, .ColIndex("����ת��"))) = 1, False)
            End With
        End If
        
        LockWindowUpdate 0
        Call vsf_RowColChange
    Case "ˢ������"
        Select Case mbytApplyMode
            Case 1          '��Ժ����
                ExecuteCommand = ExecuteCommand("��ȡ��Ժ����")
            Case 3          '��Ժ����
                ExecuteCommand = RefreshPatientIn
            Case 4          '����׼���ĵĲ���
                ExecuteCommand = RefreshPatientBorrow
        End Select
    End Select

    ExecuteCommand = True
    GoTo endHand
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
    mblnReading = False
End Function
Private Function RefreshPatientBorrow() As Boolean
'��ȡ������Ա
Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    mclsVsf.SaveKey = vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID"))
    mclsVsf.ClearGrid
    
    Set rs = gclsPackage.GetBorrowPatient(0, ParamRead(mrsCondition, "��ʼ���ݺ�"), _
                                            ParamRead(mrsCondition, "�������ݺ�"), _
                                            ParamRead(mrsCondition, "������"), _
                                            ParamRead(mrsCondition, "��׼��"), _
                                            ParamRead(mrsCondition, "�ܾ���"), _
                                            IIf(Val(ParamRead(mrsCondition, "�µǼǵ���")) = 1, ParamRead(mrsCondition, "�Ǽǿ�ʼ����"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "�µǼǵ���")) = 1, ParamRead(mrsCondition, "�Ǽǽ�������"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "����׼����")) = 1, ParamRead(mrsCondition, "��׼��ʼ����"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "����׼����")) = 1, ParamRead(mrsCondition, "��׼��������"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "�Ѿܾ�����")) = 1, ParamRead(mrsCondition, "�ܾ���ʼ����"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "�Ѿܾ�����")) = 1, ParamRead(mrsCondition, "�ܾ���������"), ""))
    If rs.BOF = False Then
        Call mclsVsf.LoadDataSource(rs)
    End If
    Set mrsData = rs
    
    RefreshPatientBorrow = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function RefreshPatientIn() As Boolean
 '��Ժ���˺ͳ�Ժδ�ύ����
    Dim l As Long, rs As New ADODB.Recordset
    
    On Error GoTo errHand
    mclsVsf.SaveKey = vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")): mclsVsf.ClearGrid
    
    Set rs = gclsPackage.GetDeptPatient(ParamRead(mrsCondition, "��Ժ��ʼʱ��"), ParamRead(mrsCondition, "��Ժ����ʱ��"), _
                                        ParamRead(mrsCondition, "��ǰ����"), ParamRead(mrsCondition, "��Ժ���"), _
                                        Val(ParamRead(mrsCondition, "��������")), ParamRead(mrsCondition, "ҽ������"), _
                                        ParamRead(mrsCondition, "סԺҽʦ"), ParamRead(mrsCondition, "��������"), _
                                        ParamRead(mrsCondition, "�������"), ParamRead(mrsCondition, "ҩƷ��Ϣ"), _
                                        ParamRead(mrsCondition, "ҽ����ʼʱ��"), ParamRead(mrsCondition, "ҽ������ʱ��"), _
                                        chkNCommit.Value = vbChecked)
    If cboDept.ListIndex >= 0 Then
        If cboDept.Text = "���п���" Then
            rs.Filter = ""
        Else
            rs.Filter = "��Ժ����ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
        End If
    End If
                    
    If rs.BOF = False Then
        Call mclsVsf.LoadDataSource(rs)
        vsf.Row = 1
    End If
        
    If mlng����ID <> 0 Then
        For l = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(l, vsf.ColIndex("����id"))) = mlng����ID And _
                Val(vsf.TextMatrix(l, vsf.ColIndex("��ҳid"))) = mlng��ҳID Then
                vsf.Row = l: Exit For
            End If
        Next
        mclsVsf.AppendRows = True
    End If
    Set mrsData = rs
    Call vsf_RowColChange
    RefreshPatientIn = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub cboDept_Click()
    Dim rsTemp As New ADODB.Recordset
    If mblnReading Then Exit Sub
    
    mstrSvr�������� = cboDept.Text
    If mrsData Is Nothing Then
        Call ExecuteCommand("ˢ������")
    Else
        If mbytApplyMode = 3 Or mbytApplyMode = 4 Then
            '��Ժ����
             If cboDept.Text = "���п���" Then
                mrsData.Filter = ""
             Else
                mrsData.Filter = "��Ժ����ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
             End If
        Else
            If cboDept.Text = "���п���" Then
                If cboStatus.Text = "����״̬" Then
                    mrsData.Filter = "����ʱ�� = '3000-01-01'"
                Else
                    mrsData.Filter = "����ʱ��= '3000-01-01' and ����״ֵ̬='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
                End If
                mfrmMain.AllowModify = True
            Else
                If cboStatus.Text = "����״̬" Then
                    mrsData.Filter = "��Ժ����ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
                Else
                    
                    mrsData.Filter = "��Ժ����ID = '" & cboDept.ItemData(cboDept.ListIndex) & "' And ����״ֵ̬='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
                End If
            End If
        End If
        Call mclsVsf.LoadDataSource(mrsData)
        '�Ѽ���
        vsf.Row = 1
        Call mclsVsf.LoadGridRow(vsf.Row, mrsData)
        mclsVsf.AppendRows = True
    End If
    Call ExecuteCommand("��ȡ�����ṹ", "Read")
    Call vsf_RowColChange
    
    If cboDept.ListIndex <> 0 And Me.cboStatus.Visible Then
        gstrSQL = "SELECT a.id FROM ���ű� a Where  a.id=[1] and ( TO_CHAR (A.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or A.����ʱ�� is null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, cboDept.ItemData(cboDept.ListIndex))
        If rsTemp.RecordCount = 0 Then mfrmMain.AllowModify = False Else mfrmMain.AllowModify = True
    End If
    RaiseEvent AfterDeptChanged
End Sub

Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If cboDept.Locked Then Exit Sub
    cboDept.Tag = "Changed"
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cboDept.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim StrText As String
    Dim strResult As String
    
    If InStr(1, "��'|[](){}*%", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If cboDept.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        StrText = UCase(cboDept.Text)
        If cboDept.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If StrText <> cboDept.List(cboDept.ListIndex) Then Call zlControl.CboSetIndex(cboDept.hWnd, -1)
        End If
        If StrText = "" Then
            cboDept.ListIndex = 0
            cboDept.Tag = ""
        ElseIf cboDept.ListIndex = -1 Then
            intIdx = -1
            
            If mrsDept.State = adStateOpen Then
                If IsNumeric(StrText) Then                              '�����ͱ���
                    mrsDept.Filter = "���� like '" & StrText & "*'"
                ElseIf zlCommFun.IsCharAlpha(StrText) Then              '�ַ��ͼ���
                    mrsDept.Filter = "���� like '*" & StrText & "*'"
                ElseIf zlCommFun.IsCharChinese(StrText) Then            '����
                    mrsDept.Filter = "���� like '*" & StrText & "*'"
                Else                                                    '���֧������N001,���������ZYK01����
                    mrsDept.Filter = "(���� like '" & StrText & "*') OR (���� like '*" & StrText & "*')"
                End If
                If mrsDept.RecordCount > 0 Then
                    mrsDept.MoveFirst
                    strResult = mrsDept("����").Value     'ֻȡ��һ��
                End If
            End If

            If mrsDept.State = adStateOpen Then mrsDept.Filter = ""
                        
            If strResult <> "" Then
                For i = 0 To cboDept.ListCount - 1
                    If zlCommFun.GetNeedName(cboDept.List(i)) = strResult Then
                        cboDept.ListIndex = i
                        cboDept.Tag = ""
                        Exit For
                    End If
                Next
            Else    '����12����ֻ��1201,1202��;����ZY,��ֻ��ZYK,ZYH��
                For i = 0 To cboDept.ListCount - 1
                    If UCase(cboDept.List(i)) Like StrText & "*" Then
                        If intIdx = -1 Then
                            cboDept.ListIndex = i
                            cboDept.Tag = ""
                        End If
                        
                        intIdx = i
                    End If
                Next
                
            End If
            
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cboDept_Click
'            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cboDept.ListIndex = -1 Then
            cboDept.Text = ""
        Else
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cboDept_Click
            ElseIf intIdx <> cboDept.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cboDept.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cboDept_Click
            End If
        End If
    End If
End Sub

Private Sub cboDept_LostFocus()
    If cboDept.Tag = "Changed" Then
        cboDept.Text = mstrSvr��������
    End If
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    If cboDept.Text <> "" Then
        If GetCboIndex(cboDept, zlCommFun.GetNeedName(cboDept.Text)) = -1 Then cboDept.ListIndex = -1: cboDept.Text = ""
    End If
    If cboDept.Text = "" Then Call cboDept_KeyPress(vbKeyReturn)
    
    If cboDept.ListIndex = -1 Then Cancel = True
End Sub

Private Sub cboStatus_Click()
    If mblnReading Then Exit Sub
    
    If mrsData Is Nothing Then
        Call ExecuteCommand("ˢ������")
    Else
        If cboStatus.Text = "����״̬" Then
            If cboDept.Text = "���п���" Then
                mrsData.Filter = ""
            Else
                mrsData.Filter = "��Ժ����ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
            End If
        Else
            If cboDept.Text = "���п���" Then
                mrsData.Filter = "����״ֵ̬='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
            Else
                mrsData.Filter = "��Ժ����ID = '" & cboDept.ItemData(cboDept.ListIndex) & "' And ����״ֵ̬='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
            End If
        End If
        Call mclsVsf.LoadDataSource(mrsData)
        '�Ѽ���
        vsf.Row = 1
        Call mclsVsf.LoadGridRow(vsf.Row, mrsData)
        mclsVsf.AppendRows = True
    End If
    Call ExecuteCommand("��ȡ�����ṹ", "Read")
    Call vsf_RowColChange
    RaiseEvent StatusChanged
End Sub

Private Sub chkNCommit_Click()
    Call ExecuteCommand("ˢ������")
End Sub

Private Sub cmdStatus_Click()
    Call FileBatPrint
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Initialize()
    On Error GoTo ErrH
    DoEvents
    frmPubResource.Hide     '����һ��ͼ�괰��
    Set tvw.ImageList = frmPubResource.ils16
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub Form_Load()
    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 2, 100, 100, Me.ScaleWidth, 300)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not mclsVsf Is Nothing Then Call SetRegister(˽��ģ��, Me.Name, "������_20100719_" & mbytApplyMode, mclsVsf.SaveStateToString)
    Set mfrmMain = Nothing
    Set mrsCondition = Nothing
    Set mclsVsf = Nothing
    Set mrsDept = Nothing
    Set mrsData = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
        Case 0
            cboDept.Move cboDept.Left, cboDept.Top, picPane(Index).Width - cboDept.Left - 15
            If mbytApplyMode = 1 Then '��Ժ����
                cboStatus.Move cboStatus.Left, cboDept.Height + 30, picPane(Index).Width - cboDept.Left - 15
                vsf.Move 0, cboStatus.Top + cboStatus.Height + 30, picPane(Index).Width, picPane(Index).Height - (cboStatus.Top + cboStatus.Height + 30) - 400
                lblStatus.Visible = True
                cboStatus.Visible = True
                chkNCommit.Visible = False
            ElseIf mbytApplyMode = 3 Then '��Ժ
                chkNCommit.Move cboDept.Left, cboDept.Height + 30, picPane(Index).Width - chkNCommit.Left - 15
                vsf.Move 0, chkNCommit.Top + chkNCommit.Height + 30, picPane(Index).Width, picPane(Index).Height - (chkNCommit.Top + chkNCommit.Height + 30) - 400
                lblStatus.Visible = False
                cboStatus.Visible = False
                chkNCommit.Visible = True
            Else '���Ĳ���
                vsf.Move 0, cboDept.Top + cboDept.Height + 30, picPane(Index).Width, picPane(Index).Height - (cboDept.Top + cboDept.Height + 30) - 400
                lblStatus.Visible = False
                cboStatus.Visible = False
                chkNCommit.Visible = False
            End If
            picSelect.Move vsf.Left, vsf.Top + vsf.Height, vsf.Width, 400
            mclsVsf.AppendRows = True
            labStatus.Width = vsf.Width
            cmdStatus.Move picPane(0).Width - cmdStatus.Width - 30
        Case 1
            tvw.Move 15, 15, picPane(Index).Width - 15, picPane(Index).Height - 15
    End Select
End Sub
Public Sub FileBatPrint()
    Dim strObject As String
    Dim strParam As String
    Dim strTmp As String
    
    If ObjPtr(tvw.SelectedItem) <= 0 Then Exit Sub
    With tvw.SelectedItem
        If .Parent Is Nothing Then
            Select Case .Key
            Case "R5"
                strObject = "��ҳ��¼"
            Case "R1"
                strObject = "סԺҽ��"
            Case "R9"
                strObject = "�ٴ�·��"
            Case Else
                blnReadUsed = False
                Exit Sub
            End Select
        Else
            strParam = .Tag
            Select Case .Parent.Key
            Case "R2"
                strObject = "סԺ����"
            Case "R3"
                strObject = "������"
            Case "R4"
                strObject = "�����¼"
            Case "R6"
                strObject = "ҽ������"
            Case "R7"
                strObject = "����֤��"
            Case "R8"
                strObject = "֪���ļ�"
            End Select
        End If
    End With
    If blnReadUsed Then Exit Sub
    blnReadUsed = True
    With vsf
        strTmp = tvw.SelectedItem.Key & "," & Val(.TextMatrix(.Row, .ColIndex("����id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("��ҳid")))
        RaiseEvent AfterDocumentChanged(Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), strObject, strParam, .TextMatrix(.Row, .ColIndex("����")) & " -> " & .Text, Val(.TextMatrix(.Row, .ColIndex("ID"))), Val(.TextMatrix(.Row, .ColIndex("����ת��"))) = 1, True)
    End With
    blnReadUsed = False
    
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strObject As String
    Dim strParam As String
    Dim strTmp As String

    If Node.Parent Is Nothing Then
        Select Case Node.Key
        Case "R5"
            strObject = "��ҳ��¼"
        Case "R1"
            strObject = "סԺҽ��"
        Case "R9"
            strObject = "�ٴ�·��"
        Case Else
            Select Case Node.Key
            Case "R2"
                strObject = "סԺ����"
            Case "R3"
                strObject = "������"
            Case "R4"
                strObject = "�����¼"
            Case "R6"
                strObject = "ҽ������"
            Case "R7"
                strObject = "����֤��"
            Case "R8"
                strObject = "֪���ļ�"
            End Select
        End Select
    Else
        strParam = Node.Tag
        Select Case Node.Parent.Key
        Case "R2"
            strObject = "סԺ����"
        Case "R3"
            strObject = "������"
        Case "R4"
            strObject = "�����¼"
        Case "R6"
            strObject = "ҽ������"
        Case "R7"
            strObject = "����֤��"
        Case "R8"
            strObject = "֪���ļ�"
        End Select
    End If
    
    With vsf
        tvw.Tag = Node.Key
        strTmp = Node.Key & "," & Val(.TextMatrix(.Row, .ColIndex("����id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("��ҳid")))

        mstrKey = strTmp
        RaiseEvent AfterDocumentChanged(Val(.TextMatrix(.Row, .ColIndex("����id"))), Val(.TextMatrix(.Row, .ColIndex("��ҳid"))), strObject, strParam, .TextMatrix(.Row, .ColIndex("����")) & " -> " & Node.Text, Val(.TextMatrix(.Row, .ColIndex("ID"))), Val(.TextMatrix(.Row, .ColIndex("����ת��"))) = 1, False)
    End With
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    RaiseEvent AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    
    If OldRow <> NewRow Then
        Call ExecuteCommand("��ȡ�����ṹ", "NoRead")
    End If
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterSort(ByVal Col As Long, Order As Integer)
    Call mclsVsf.RestoreRow(mclsVsf.SaveKey)
    vsf.ShowCell vsf.Row, vsf.Col
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsf.ColIndex("ѡ��") <> Col Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub vsf_RowColChange()
On Error Resume Next
    If mbytApplyMode = 3 Or mbytApplyMode = 4 Then
        labSelect.Visible = False
        labNum.Visible = False
        With vsf
            If .Rows = 1 Then
                labStatus.Caption = ""
            Else
                If .ColIndex("����") <> -1 Then
                    labStatus.Caption = ChkStrUniCode("������" & .TextMatrix(.Row, .ColIndex("����")) & "                    ", 20) & "   סԺ�ţ�" & .TextMatrix(.Row, .ColIndex("סԺ��"))
                End If
            End If
        End With
    End If
End Sub

Private Sub vsf_BeforeSort(ByVal Col As Long, Order As Integer)
    mclsVsf.SaveKey = Val(vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")))
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick()
    Call mclsVsf.DbClick
    Call ExecuteCommand("��ȡ�����ṹ", "Read")
    RaiseEvent DbClick
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick
    End If
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SendLMouseButton(vsf.hWnd, x, y)
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub


