VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIconManage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ͼ�����"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   Icon            =   "frmIconManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��ͼ��"
      Height          =   350
      Left            =   1320
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdLoaclIcon 
      Caption         =   "���ͼ��"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   4320
      TabIndex        =   4
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   350
      Left            =   3120
      TabIndex        =   3
      Top             =   3840
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   2640
      TabIndex        =   5
      Top             =   0
      Width           =   2775
      Begin MSComDlg.CommonDialog dlgIcon 
         Left            =   1800
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgIcon 
         Height          =   855
         Left            =   1080
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblView 
         AutoSize        =   -1  'True
         Caption         =   "ͼ��Ԥ����"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfIconName 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _cx             =   4260
      _cy             =   6376
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
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
   Begin VB.Image imgNoCheck 
      Height          =   255
      Left            =   0
      Picture         =   "frmIconManage.frx":058A
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Left            =   0
      Picture         =   "frmIconManage.frx":08FC
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmIconManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnIsOK As Boolean
Private mstrUsedIcon As String
Private mblnIsDelete As Boolean
Private mstrIconName As String    'ͼ������
Private Const M_NUM_ѡ�� = 0
Private Const M_NUM_ͼ������ = 1
Private Const M_NUM_ͼ��Ԥ�� = 2


Private Sub cmdCancel_Click()
    On Error GoTo errHandle
    
    mblnIsOK = False
    Unload Me
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdDelete_Click()
'ɾ��ͼ��
    Dim strSql As String
    Dim strName As String
    
    On Error GoTo errHandle
    
    If Not IsSelectionRow(vsfIconName) Then Exit Sub
    
    strName = vsfIconName.TextMatrix(vsfIconName.Row, M_NUM_ͼ������)
    If Len(strName) = 0 Then
        Exit Sub
    End If
    
    If InStr(UCase(Trim(mstrUsedIcon)), UCase(Trim(strName))) > 0 Then
        If MsgBox("��ͼ������ʹ�ã��Ƿ�ɾ����", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If
    
    strSql = "Zl_Ӱ���ѯ_ɾ��ͼ��('" & vsfIconName.TextMatrix(vsfIconName.Row, M_NUM_ͼ������) & "')"
    Call ExecuteCmd(strSql, "ɾ��ͼƬ")
    vsfIconName.RemoveItem (vsfIconName.Row)
    If vsfIconName.Rows < 2 Then cmdDelete.Enabled = False
    Call ShowIcon
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Function UsedIcon() As String
    Dim strSql As String
    Dim strIcon As String
    Dim objSqlScheme As clsSqlScheme
    Dim rsData As ADODB.Recordset
    Dim strSchemeText As String
    Dim j As Long
    Dim m As Long
    
    strSql = "select id, ��������,'' as ��������" & _
            " from Ӱ���ѯ���� order by id"
    Set rsData = ExecuteSql(strSql, "��������")
    
    If rsData.RecordCount <= 0 Then
        MsgBox Me, "û�п����ڵ��������ݣ����鷽�����á�", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    rsData.MoveFirst
    While Not rsData.EOF
        strSchemeText = ""
        strSchemeText = ReadSchemeXml(rsData.Fields!Id, "")

        Set objSqlScheme = New clsSqlScheme
        Call objSqlScheme.OpenScheme(strSchemeText)
        For j = 1 To objSqlScheme.ShowCfgCount
            For m = 1 To objSqlScheme.ShowCfg(j).RowRelationCount
                If Len(Trim(objSqlScheme.ShowCfg(j).RowRelation(m).Icon)) > 0 Then
                    If InStr(UCase(strIcon), UCase("[" & Trim(objSqlScheme.ShowCfg(j).RowRelation(m).Icon)) & "]") = 0 Then
                        strIcon = strIcon & ",[" & objSqlScheme.ShowCfg(j).RowRelation(m).Icon & "]"
                    End If
                End If
            Next
            If Len(Trim(objSqlScheme.ShowCfg(j).Icon)) > 0 Then
                If InStr(UCase(strIcon), UCase("[" & Trim(objSqlScheme.ShowCfg(j).Icon)) & "]") = 0 Then
                    strIcon = strIcon & ",[" & objSqlScheme.ShowCfg(j).Icon & "]"
                End If
            End If
        Next
        Call rsData.MoveNext
    Wend
    UsedIcon = Mid(strIcon, 2)
End Function

Private Sub cmdLoaclIcon_Click()
    Dim arrName() As String
    Dim strName As String
    Dim strFile As String
    
    On Error GoTo errHandle
    
    dlgIcon.Filter = "(*.ico)|*.ico|(*.*)|*.*"
    dlgIcon.DefaultExt = "*.ico   "

    dlgIcon.ShowOpen
    strFile = dlgIcon.FileName
    If Len(strFile) = 0 Then Exit Sub
    
    arrName = Split(dlgIcon.FileName, "\")
    strName = arrName(UBound(arrName))
    arrName = Split(strName, ".")
    strName = Replace(strName, "." & arrName(UBound(arrName)), "")
    
    imgIcon.Picture = LoadPicture(strFile)

    Call NewIcon(strFile, strName)
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub CmdOK_Click()
    On Error GoTo errHandle
    
    mblnIsOK = True
    Me.Hide
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Call GridInit("ѡ��|ͼ������|ͼ��Ԥ��", vsfIconName)
    Call InitIconList
    mstrUsedIcon = UsedIcon
    
    cmdDelete.Enabled = vsfIconName.Rows > 1
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Public Sub UnloadMe()
    Unload Me
End Sub

Private Sub NewIcon(strFileRoad As String, strIconName As String)
'����ͼ��
    Dim strSql As String
    Dim strName As String
    Dim i As Long
    
    If Len(strFileRoad) = 0 Then Exit Sub
    If Len(Dir(strFileRoad)) = 0 Then Exit Sub
    If Len(Trim(strIconName)) = 0 Then Exit Sub
    
    For i = 1 To vsfIconName.Rows - 1
        If UCase(Trim(vsfIconName.TextMatrix(i, M_NUM_ͼ������))) = UCase(Trim(strIconName)) Then
            MsgBox "�Ѵ��ڸ����Ƶ�ͼ��,����", vbInformation, Me.Caption
            Exit Sub
        End If
    Next
    
    strSql = "Zl_Ӱ���ѯ_����ͼ��('" & strIconName & "','1')"
    Call ExecuteCmd(strSql, "����ͼ��")
    Call zlBlobSave(strIconName, strFileRoad)
    Call NewRow(vsfIconName)
    
    With vsfIconName
        
        .Cell(flexcpPicture, .Row, M_NUM_ѡ��) = imgNoCheck.Picture
        .Cell(flexcpData, .Row, M_NUM_ѡ��) = 0
        .Cell(flexcpPictureAlignment, .Row, M_NUM_ѡ��) = flexPicAlignCenterCenter
        .TextMatrix(.Row, M_NUM_ͼ������) = strIconName
        If .Rows > 1 Then cmdDelete.Enabled = True
    End With
    Call ShowIcon
End Sub

Private Function IsDBA() As Boolean
On Error GoTo errH
    Dim strSql As String
    Dim rsTmp As Recordset
    
    strSql = "select ������ from ZLSystems where ��� = 100 and ���� = 'ҽԺϵͳ��׼��'"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ȡ������")
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    If UCase(getUser(gcnOracle.ConnectionString)) = UCase(rsTmp("������")) Then
        IsDBA = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub InitIconList()
'����ͼ���б�
    Dim rsIcon As Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "select '' as ѡ��,��Դ���� as ͼ������,'' as ͼ��Ԥ�� from Ӱ���ѯ��Դ where ��Դ���� = [1]"
    Set rsIcon = ExecuteSql(strSql, "��ѯͼ��", 1)
    
    If rsIcon.RecordCount <= 0 Then Exit Sub
    
    With vsfIconName
        Set .DataSource = rsIcon
        
        .ColHidden(M_NUM_ͼ��Ԥ��) = True
        If mblnIsDelete Then
            .ColHidden(M_NUM_ѡ��) = True
        Else
            .Cell(flexcpPicture, 1, M_NUM_ѡ��, .Rows - 1, M_NUM_ѡ��) = imgNoCheck.Picture
            .Cell(flexcpData, 1, M_NUM_ѡ��, .Rows - 1, M_NUM_ѡ��) = 0
            .Cell(flexcpPictureAlignment, 1, M_NUM_ѡ��, .Rows - 1, M_NUM_ѡ��) = flexPicAlignCenterCenter
            .ColWidth(M_NUM_ѡ��) = 480
        End If
        
        If Len(mstrIconName) > 0 Then
            For i = 1 To vsfIconName.Rows - 1
                If .TextMatrix(i, M_NUM_ͼ������) = mstrIconName Then
                    .Row = i
                    .Cell(flexcpPicture, i, M_NUM_ѡ��) = imgCheck.Picture
                    .Cell(flexcpData, i, M_NUM_ѡ��) = 1
                    .Cell(flexcpPictureAlignment, i, M_NUM_ѡ��) = flexPicAlignCenterCenter
                    Exit For
                End If
            Next
        End If
    End With
    Call ShowIcon
End Sub


Private Sub vsfIconName_Click()
    Dim blnCheck As Boolean
    
    On Error GoTo errHandle
    
    If mblnIsDelete Then Exit Sub
    
    With vsfIconName
        If .Cell(flexcpData, .Row, M_NUM_ѡ��) = 0 Then
            blnCheck = False
        Else
            blnCheck = True
        End If
        .Cell(flexcpPicture, 1, M_NUM_ѡ��, .Rows - 1, M_NUM_ѡ��) = imgNoCheck.Picture
        .Cell(flexcpData, 1, M_NUM_ѡ��, .Rows - 1, M_NUM_ѡ��) = 0
        
        If Not blnCheck Then
            .Cell(flexcpPicture, .Row, M_NUM_ѡ��) = imgCheck.Picture
            .Cell(flexcpData, .Row, M_NUM_ѡ��) = 1
        End If
        
        If .Cell(flexcpData, .Row, M_NUM_ѡ��) = 1 Then
            Call ShowIcon
        Else
            imgIcon.Picture = Nothing
        End If
    End With
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub
'
Private Sub vsfIconName_RowColChange()
    On Error GoTo errHandle
    
    If mblnIsDelete Then
        Call ShowIcon
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub ShowIcon()
'Ԥ��ͼ��
    Dim strFile As String
    
    Set imgIcon.Picture = Nothing
    If vsfIconName.Row < 1 Then Exit Sub
    
    
    If vsfIconName.Cell(flexcpPicture, vsfIconName.Row, M_NUM_ͼ��Ԥ��) Is Nothing Then
        strFile = zlBlobRead(vsfIconName.TextMatrix(vsfIconName.Row, M_NUM_ͼ������))
        If Len(strFile) = 0 Then Exit Sub
        
        If Len(Dir(strFile)) > 0 Then
            imgIcon.Picture = LoadPicture(strFile)
            vsfIconName.Cell(flexcpPicture, vsfIconName.Row, M_NUM_ͼ��Ԥ��) = imgIcon.Picture
            Kill strFile
        End If
    Else
        Set imgIcon.Picture = vsfIconName.Cell(flexcpPicture, vsfIconName.Row, M_NUM_ͼ��Ԥ��)
    End If
End Sub

Public Function ShowIconWindow(ByRef strIconName As String, owner As Object, Optional lngDelete As Long) As Object
    mblnIsOK = False
    mstrIconName = strIconName
    
    If lngDelete <> 1 And IsDBA Then
        mblnIsDelete = True
    Else
        cmdDelete.Visible = False
        mblnIsDelete = False
    End If
    
    Call Me.Show(1, owner)
    
    If mblnIsOK Then
        If vsfIconName.Cell(flexcpData, vsfIconName.Row, M_NUM_ѡ��) = 1 Then
            strIconName = Trim(vsfIconName.TextMatrix(vsfIconName.Row, M_NUM_ͼ������))
        Else
            strIconName = ""
        End If
        
        Set ShowIconWindow = imgIcon.Picture
    Else
        Set ShowIconWindow = Nothing
    End If
End Function


Private Function getUser(strTmp As String) As String
    Dim arrTmp() As String
    
    arrTmp = Split(strTmp, "User ID=")
    If UBound(arrTmp) > 0 Then
        getUser = Split(arrTmp(1), ";")(0)
    End If
    
End Function
