VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmUserQueryReleation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�û����ò�ѯ����"
   ClientHeight    =   6690
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserQueryReleation.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�Ĭ��(&R)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   1665
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
      Height          =   4455
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   11175
      _cx             =   19711
      _cy             =   7858
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      BackColorSel    =   13082765
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
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
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   6120
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ ��(&Q)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10245
      TabIndex        =   4
      Top             =   6120
      Width           =   1185
   End
   Begin VB.ComboBox cbxUser 
      Appearance      =   0  'Flat
      Height          =   312
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   220
      Width           =   3012
   End
   Begin VB.ComboBox cbxDepart 
      Appearance      =   0  'Flat
      Height          =   312
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   216
      Width           =   2892
   End
   Begin VB.Image imgNoCheck 
      Height          =   255
      Left            =   0
      Picture         =   "frmUserQueryReleation.frx":000C
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   1800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   255
      Left            =   0
      Picture         =   "frmUserQueryReleation.frx":037E
      Stretch         =   -1  'True
      Tag             =   "1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label labNote 
      BackColor       =   &H00DDF8FB&
      Caption         =   "����˵����"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   5145
      Width           =   11175
   End
   Begin VB.Label Label2 
      Caption         =   "�������ƣ�"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "�û�����:"
      Height          =   252
      Left            =   4800
      TabIndex        =   0
      Top             =   240
      Width           =   972
   End
End
Attribute VB_Name = "frmUserQueryReleation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TColDef
    cdName = 0          '��������
    cdDeptId = 1        '��������
    cdUserDef = 2       '�û�Ĭ��
    cdCommonUse = 3     '�Ƿ���
    cdDefLoadStation = 4 'Ĭ�ϼ���վ��
    cdOnlyStation = 5   '��վ����ʾ
    cdSchemeDescript = 6 '��������
End Enum

Private mlngModuleNo As Long
Private mlngUserId As Long
Private mlngDeptId As Long
Private mblnIsLoading As Boolean

Private mblnIsOK As Boolean


Public Function ShowUserScheme(owner As Object, ByVal lngModuleNo As Long, _
    Optional ByVal lngUserId As Long = 0, Optional ByVal lngDeptId As Long = 0) As Boolean
    mblnIsOK = False
    
    ShowUserScheme = False
    mlngModuleNo = lngModuleNo
    mlngUserId = lngUserId
    
    mlngDeptId = lngDeptId
    If lngUserId = 0 Then mlngDeptId = 0
    
    Me.Show 1, owner
    
    ShowUserScheme = mblnIsOK
End Function

Private Sub LoadDepartInfo()
'���������Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    cbxDepart.Clear
    
    '�û�IDΪ��ʱ��˵���ǴӲ�ѯ���ô����е��õĳ��÷�������
    If mlngUserId <> 0 Then
        cbxDepart.BackColor = &H8000000F
        cbxDepart.Enabled = False
        Exit Sub
    Else
        cbxDepart.BackColor = vbWhite
        cbxDepart.Enabled = True
    End If
    
    strSql = "Select ID,���� From ���ű� A, ��������˵�� B where A.ID=B.����ID And B.��������='���' Order By ����"
    Set rsData = ExecuteSql(strSql, "��ѯ������Ϣ")
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        
        cbxDepart.AddItem NVL(rsData!����)
        cbxDepart.ItemData(cbxDepart.ListCount - 1) = Val(NVL(rsData!Id))
        
        Call rsData.MoveNext
    Wend
    
    cbxDepart.AddItem ""
    
    cbxDepart.ListIndex = 0
End Sub

Private Sub LoadUserInfo()
'�����û���Ϣ
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngUserId As Long
    Dim lngIndex As Long
    Dim blnIsQueryCurUser As Boolean
    
    cbxUser.Clear
    
    If mlngUserId <= 0 Then
        cbxUser.BackColor = vbWhite
        cbxUser.Enabled = True
        
        If cbxDepart.Text = "" Then
            vsfGrid.Rows = 1
            Exit Sub
        End If
        
        
        strSql = "Select ID, ����,�û��� From ��Ա�� A, ������Ա B, �ϻ���Ա�� C Where A.ID=B.��ԱID  And A.ID=C.��ԱID And B.����ID=[1] Order By ����"
        Set rsData = ExecuteSql(strSql, "��ѯ��Ա��Ϣ", cbxDepart.ItemData(cbxDepart.ListIndex))
        
        If rsData.RecordCount <= 0 Then
            vsfGrid.Rows = 1
            Exit Sub
        End If
    Else
        cbxUser.BackColor = &H8000000F
        cbxUser.Enabled = False
        
        strSql = "Select Id, ����,'��ǰ�û�' as �û��� From ��Ա�� Where ID=[1]"
        Set rsData = ExecuteSql(strSql, "��ѯ��ǰ��Ա��Ϣ", mlngUserId)
        
        If rsData Is Nothing Then
            vsfGrid.Rows = 1
            Exit Sub
        End If
        
        If rsData.RecordCount <= 0 Then
            vsfGrid.Rows = 1
            Exit Sub
        End If
    End If
        
    While Not rsData.EOF
        lngUserId = Val(NVL(rsData!Id))
        
        cbxUser.AddItem NVL(rsData!�û���) & "-" & NVL(rsData!����)
        cbxUser.ItemData(cbxUser.ListCount - 1) = lngUserId
        
        If lngUserId = mlngUserId Then
            lngIndex = cbxUser.ListCount - 1
        End If
        
        Call rsData.MoveNext
    Wend
    
    cbxUser.ListIndex = lngIndex
End Sub

Public Sub LoadSchemeConfig()
On Error GoTo errH
'�����û���������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim i As Long
    Dim blnIsUser As Boolean
    Dim strDeptIds As String
    Dim blnIsDeptDef As Boolean
    
    vsfGrid.Rows = 1
    
    If cbxUser.Text = "" Then
        vsfGrid.Rows = 1
        Exit Sub
    End If
    
    strDeptIds = ""
    If mlngUserId <> 0 Then
        If mlngDeptId = 0 Then
            strSql = "SELECT ��ԱID,WM_CONCAT(����ID) as ��������  FROM ������Ա  where ��ԱID=[1] GROUP BY ��ԱID"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ�û���Ӧ����", mlngUserId)
            If rsData.RecordCount > 0 Then
                strDeptIds = ",0," & NVL(rsData!��������) & ","
            End If
        Else
            strDeptIds = ",0," & mlngDeptId & ","
        End If
    Else
        If mlngDeptId <> 0 Then
            strDeptIds = ",0," & mlngDeptId & ","
        End If
    End If
    
    strSql = "Select A.ID, A.��������,A.��������, B.�û�ID, " & vbCrLf & _
            "   case when B.�û�ID Is Null then A.�Ƿ�Ĭ�� else decode(B.�Ƿ�Ĭ��, null,B.�Ƿ�Ĭ��,B.�Ƿ�Ĭ��+1) End As �Ƿ�Ĭ��, " & vbCrLf & _
            "   case when B.�û�ID Is Null then A.�Ƿ��� else decode(B.�Ƿ���, null,B.�Ƿ���,B.�Ƿ���+1) End As �Ƿ���, " & vbCrLf & _
            "   B.Ĭ�ϼ���վ��, B.����վ��, A.����˵�� " & vbCrLf & _
            " From Ӱ���ѯ���� A, Ӱ���ѯ���� B " & vbCrLf & _
            " Where A.ID=B.��ѯ����ID(+) And A.ʹ��״̬=1 And A.����ģ��=[1] And B.�û�ID(+)=[2] " & IIf(strDeptIds <> "", " and Instr([3], ',' || nvl(A.��������, 0) || ',' ) > 0 ", " ") & " order by �������"
              

    Set rsData = ExecuteSql(strSql, "�������з���", mlngModuleNo, Val(cbxUser.ItemData(cbxUser.ListIndex)), strDeptIds)
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    rsData.Filter = "�û�ID=" & Val(cbxUser.ItemData(cbxUser.ListIndex))
    
    '�ж��Ƿ�����û�����������
    blnIsUser = IIf(rsData.RecordCount <= 0, False, True)
 
    rsData.Filter = ""

    
    vsfGrid.Rows = rsData.RecordCount + 1
    vsfGrid.ColHidden(cdDeptId) = True
    
    blnIsDeptDef = False
    i = 1
    While Not rsData.EOF
        vsfGrid.RowData(i) = NVL(rsData!Id)
        vsfGrid.Cell(flexcpText, i, cdDeptId) = Val(NVL(rsData!��������))
        
        vsfGrid.Cell(flexcpText, i, cdName) = NVL(rsData!��������)
        
        If Val(NVL(rsData!�Ƿ�Ĭ��)) > IIf(blnIsUser, 1, 0) Then
            vsfGrid.Cell(flexcpData, i, cdUserDef) = 1
            vsfGrid.Cell(flexcpPicture, i, cdUserDef) = imgCheck.Picture
            
            If Val(NVL(rsData!��������)) <> 0 Then
                blnIsDeptDef = True
            End If
        Else
            vsfGrid.Cell(flexcpData, i, cdUserDef) = 0
            vsfGrid.Cell(flexcpPicture, i, cdUserDef) = imgNoCheck.Picture
        End If
        
        If Val(NVL(rsData!�Ƿ���)) > IIf(blnIsUser, 1, 0) Then
            vsfGrid.Cell(flexcpData, i, cdCommonUse) = 1
            vsfGrid.Cell(flexcpPicture, i, cdCommonUse) = imgCheck.Picture
        Else
            vsfGrid.Cell(flexcpData, i, cdCommonUse) = 0
            vsfGrid.Cell(flexcpPicture, i, cdCommonUse) = imgNoCheck.Picture
        End If
                
        vsfGrid.Cell(flexcpText, i, cdDefLoadStation) = NVL(rsData!Ĭ�ϼ���վ��)
        vsfGrid.Cell(flexcpText, i, cdOnlyStation) = NVL(rsData!����վ��)
        vsfGrid.Cell(flexcpText, i, cdSchemeDescript) = NVL(rsData!����˵��)
        
        i = i + 1
        
        Call rsData.MoveNext
    Wend
    
    If blnIsDeptDef And blnIsUser = False Then
        '���û�п������õ�Ĭ�Ϸ�����ѡ
        For i = 1 To vsfGrid.Rows - 1
            If Val(vsfGrid.TextMatrix(i, cdDeptId)) = 0 Then
                vsfGrid.Cell(flexcpData, i, cdUserDef) = 0
                vsfGrid.Cell(flexcpPicture, i, cdUserDef) = imgNoCheck.Picture
            End If
        Next
    End If
    
    vsfGrid.Cell(flexcpBackColor, 1, cdName, i - 1, cdName) = &HDDF8FB
    vsfGrid.Cell(flexcpPictureAlignment, 1, cdUserDef, i - 1, cdUserDef) = flexPicAlignCenterCenter
    vsfGrid.Cell(flexcpPictureAlignment, 1, cdCommonUse, i - 1, cdCommonUse) = flexPicAlignCenterCenter
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbxDepart_Change()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    If cbxDepart.ListIndex >= 0 Then
        mlngUserId = 0
        mlngDeptId = Val(cbxDepart.ItemData(cbxDepart.ListIndex))
    End If
    
    Call LoadUserInfo
    Call LoadStationInfos
    
'    Call LoadSchemeConfig
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxDepart_Click()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    If cbxDepart.ListIndex >= 0 Then
        mlngUserId = 0
        mlngDeptId = Val(cbxDepart.ItemData(cbxDepart.ListIndex))
    End If
    
    Call LoadUserInfo
    Call LoadStationInfos
    
'    Call LoadSchemeConfig
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxUser_Change()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    If cbxUser.ListIndex >= 0 Then
        mlngUserId = Val(cbxUser.ItemData(cbxUser.ListIndex))
    End If
    
    Call LoadSchemeConfig
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cbxUser_Click()
On Error GoTo errHandle
    If mblnIsLoading Then Exit Sub
    
    If cbxUser.ListIndex >= 0 Then
        mlngUserId = Val(cbxUser.ItemData(cbxUser.ListIndex))
    End If
    
    Call LoadSchemeConfig
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub SaveConfig(ByVal lngUserId As Long)
    Dim i As Long
    Dim blnIsDefault As Boolean
    Dim blnIsCommonUse As Boolean
    Dim strStationDef As String
    Dim strOnlyStation As String
    Dim strSql As String
    Dim blnIsStartTrans As Boolean
    
    strSql = "zl_Ӱ���ѯ_�������(" & lngUserId & ")"
    Call ExecuteCmd(strSql, "����û���ѯ����")
    
    On Error GoTo errHandle:
    
    blnIsStartTrans = False
    For i = 1 To vsfGrid.Rows - 1
        blnIsDefault = IIf(vsfGrid.Cell(flexcpData, i, cdUserDef) = 1, True, False)
        blnIsCommonUse = IIf(vsfGrid.Cell(flexcpData, i, cdCommonUse) = 1, True, False)
        strOnlyStation = vsfGrid.Cell(flexcpText, i, cdOnlyStation)
        strStationDef = UCase(vsfGrid.Cell(flexcpText, i, cdDefLoadStation))
        
        If blnIsDefault Or blnIsCommonUse Or Trim(strOnlyStation) <> "" Then
            If blnIsStartTrans = False Then
                gcnOracle.BeginTrans
                blnIsStartTrans = True
            End If
            
            strSql = "zl_Ӱ���ѯ_���¹���(" & lngUserId & "," & Val(vsfGrid.RowData(i)) & "," & _
                                            IIf(blnIsDefault, 1, 0) & "," & IIf(blnIsCommonUse, 1, 0) & ",'" & strStationDef & "','" & _
                                            strOnlyStation & "')"
            Call ExecuteCmd(strSql, "�û���ѯ����")
        End If
    Next i
    
    If blnIsStartTrans Then gcnOracle.CommitTrans
Exit Sub
errHandle:
    If blnIsStartTrans Then gcnOracle.RollbackTrans
    Debug.Print "SaveConfig Err:" & Err.Description
    Err.Raise -1, "SaveConfig", "[SaveConfig]�����û��������ô���>>" & Err.Description
    Resume
End Sub

Private Function Validate() As Boolean
    Dim i As Long
    Dim j As Long
    Dim strDefLoadStationCfg As String
    Dim strOnlyStation As String
    Dim strStations() As String
     
    Validate = True
    
    For i = 1 To vsfGrid.Rows - 1
        strDefLoadStationCfg = UCase(vsfGrid.TextMatrix(i, cdDefLoadStation))
        strOnlyStation = UCase(vsfGrid.TextMatrix(i, cdOnlyStation))
        
        If Trim(strOnlyStation) <> "" And Trim(strDefLoadStationCfg) <> "" Then
            If strDefLoadStationCfg <> strOnlyStation Then
                Validate = False
                MsgBox "���ý�վ����ʾ��Ĭ�ϼ���վ��Ӧ�ͽ�վ����ʾ������ͬ.", vbOKOnly, Me.Caption
                
                vsfGrid.Row = i
                vsfGrid.Col = cdDefLoadStation
                vsfGrid.EditCell
                        
                Exit Function
            End If
        End If
        
        If Trim(strDefLoadStationCfg) <> "" Then
            strStations = Split(strDefLoadStationCfg, ";")
            For j = 0 To UBound(strStations)
                If Trim(strStations(j)) <> "" Then
                    If HasStation(strStations(j), i + 1) Then
                        Validate = False
                        MsgBox "Ĭ�ϼ���վ�� [" & strStations(j) & "] ֻ�ܶ�Ӧһ�ַ���.", vbOKOnly, Me.Caption
                        
                        vsfGrid.Row = i
                        vsfGrid.Col = cdDefLoadStation
                        vsfGrid.EditCell
                        
                        Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

Private Function HasStation(ByVal strStationName As String, ByVal lngStartRow As Long) As Boolean
    Dim i As Long
    
    HasStation = False
    For i = lngStartRow To vsfGrid.Rows - 1
        If InStr(UCase(vsfGrid.TextMatrix(i, cdDefLoadStation)), strStationName) > 0 Then
            HasStation = True
        End If
    Next
End Function

Private Sub cmdRestore_Click()
    Dim strSql As String
    Dim lngUserId As Long
    
    lngUserId = 0
    If cbxUser.ListIndex >= 0 Then lngUserId = Val(cbxUser.ItemData(cbxUser.ListIndex))
  
    If lngUserId <= 0 Then
        MsgBox "��ѡ����Ҫ�ָ�Ĭ�����õ��û���", vbOKOnly, Me.Caption
    End If
    
    strSql = "zl_Ӱ���ѯ_�������(" & lngUserId & ")"
    Call ExecuteCmd(strSql, "����û���ѯ����")
    
    mlngUserId = lngUserId
    Call LoadSchemeConfig
End Sub

Private Sub cmdSure_Click()
'ȷ�ϴ���
On Error GoTo errHandle
    If Validate = False Then Exit Sub
    
    Call SaveConfig(Val(cbxUser.ItemData(cbxUser.ListIndex)))
    mblnIsOK = True
    
    MsgBox "�������óɹ�,���ý����´μ���ʱ��Ч��", vbOKOnly, Me.Caption
    
    Unload Me
Exit Sub
errHandle:
    MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub Form_Load()
    mblnIsLoading = True
    
    Call InitFace
    Call InitList
    
    Call LoadDepartInfo
    Call LoadUserInfo
    Call LoadStationInfos
    
    Call LoadSchemeConfig
    
    mblnIsLoading = False
End Sub

Private Sub InitFace()
    If mlngUserId = 0 Then
        vsfGrid.Top = 600
        vsfGrid.Height = 4455
    Else
        vsfGrid.Top = 120
        vsfGrid.Height = 4935
    End If
End Sub

Private Function GetStationCfgString(ByVal strDepartName As String) As String
    Dim strResult As String
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strCurStationName As String
    
    strCurStationName = UCase(StationName)
    
    strResult = " |" & strCurStationName
    GetStationCfgString = strResult
    
    strSql = "Select ����վ From ZlClients Where ����=[1] or ���� Is Null Order By ����վ"
    
    Set rsData = ExecuteSql(strSql, "��ѯվ��", strDepartName)
    
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount <= 0 Then Exit Function
    
    While Not rsData.EOF
        If NVL(rsData!����վ) <> strCurStationName Then
            If strResult <> "" Then strResult = strResult & "|"
            strResult = strResult & "|" & NVL(rsData!����վ)
        End If
        
        Call rsData.MoveNext
    Wend
    
    GetStationCfgString = strResult
End Function

Private Sub LoadStationInfos()
    vsfGrid.ColComboList(cdOnlyStation) = GetStationCfgString(cbxDepart.Text)
End Sub

Private Sub InitList()
    vsfGrid.Cols = 7
    
    vsfGrid.Cell(flexcpText, 0, cdName) = "��������"
    vsfGrid.Cell(flexcpText, 0, cdDeptId) = "��������"
    vsfGrid.Cell(flexcpText, 0, cdUserDef) = "�û�Ĭ��"
    vsfGrid.Cell(flexcpText, 0, cdCommonUse) = "�Ƿ���"
    vsfGrid.Cell(flexcpText, 0, cdDefLoadStation) = "Ĭ�ϼ���վ��"
    vsfGrid.Cell(flexcpText, 0, cdOnlyStation) = "��վ����ʾ"
    vsfGrid.Cell(flexcpText, 0, cdSchemeDescript) = "����˵��"
    
    vsfGrid.ColWidth(cdDefLoadStation) = 2600
    vsfGrid.ColHidden(cdDeptId) = True
    vsfGrid.ColHidden(cdSchemeDescript) = True
    
    
    
    vsfGrid.ColWidth(0) = 4000
End Sub

Private Sub vsfGrid_Click()
On Error GoTo errHandle
    Dim i As Long
    
    If vsfGrid.RowSel < 1 Then Exit Sub
    
    Select Case vsfGrid.ColSel
        Case cdUserDef  '�Ƿ�Ĭ���д���
            If vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdUserDef) = 1 Then
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdUserDef) = 0
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdUserDef) = imgNoCheck.Picture
            Else
                For i = 1 To vsfGrid.Rows - 1
                    vsfGrid.Cell(flexcpData, i, cdUserDef) = 0
                    vsfGrid.Cell(flexcpPicture, i, cdUserDef) = imgNoCheck.Picture
                Next i
                
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdUserDef) = imgCheck.Picture
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdUserDef) = 1
            End If
        Case cdCommonUse  '�Ƿ����д���
            If vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdCommonUse) = 1 Then
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdCommonUse) = 0
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdCommonUse) = imgNoCheck.Picture
            Else
                vsfGrid.Cell(flexcpData, vsfGrid.RowSel, cdCommonUse) = 1
                vsfGrid.Cell(flexcpPicture, vsfGrid.RowSel, cdCommonUse) = imgCheck.Picture
            End If
    End Select
    
    
Exit Sub
errHandle:
    Debug.Print "vsfGrid_DblClick Err:" & Err.Description
End Sub


 
Private Sub vsfGrid_SelChange()
On Error GoTo errHandle
    labNote.Caption = "����˵����" & vsfGrid.Cell(flexcpText, vsfGrid.RowSel, cdSchemeDescript)
Exit Sub
errHandle:
    Debug.Print "" & Err.Description
End Sub

Private Sub vsfGrid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> cdDefLoadStation And Col <> cdOnlyStation Then Cancel = True
End Sub
