VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpsStationArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6570
   Icon            =   "frmOpsStationArrange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame3 
      Caption         =   "��������"
      Height          =   2055
      Left            =   30
      TabIndex        =   10
      Top             =   3375
      Width           =   5145
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   2
         Left            =   4620
         TabIndex        =   23
         Text            =   "1"
         Top             =   315
         Width           =   390
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   660
         Width           =   1845
      End
      Begin VB.ListBox lst 
         Height          =   900
         Left            =   1455
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   1035
         Width           =   3555
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   1845
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ⱦ����(&3)"
         Height          =   195
         Index           =   3
         Left            =   3330
         TabIndex        =   15
         Top             =   720
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ⱦ����(&2)"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   1065
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Caption         =   "��̨����(&F)"
         Height          =   195
         Index           =   0
         Left            =   3330
         TabIndex        =   13
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������(&X)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����̶�(&J)"
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ա(&R)"
      Height          =   2370
      Left            =   30
      TabIndex        =   8
      Top             =   945
      Width           =   5145
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1995
         Left            =   105
         TabIndex        =   9
         Top             =   255
         Width           =   4950
         _cx             =   8731
         _cy             =   3519
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
         BackColorSel    =   16772055
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5355
      TabIndex        =   18
      Top             =   1350
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5355
      TabIndex        =   16
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5355
      TabIndex        =   17
      Top             =   570
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   945
      Left            =   45
      TabIndex        =   19
      Top             =   -45
      Width           =   5115
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   525
         Width           =   3510
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   1
         Left            =   4635
         Picture         =   "frmOpsStationArrange.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "��ѡ����ݼ���F3"
         Top             =   495
         Width           =   345
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   300
         Index           =   0
         Left            =   3945
         TabIndex        =   2
         Text            =   "1"
         Top             =   180
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   1125
         TabIndex        =   1
         Top             =   165
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   106692611
         CurrentDate     =   38083
      End
      Begin MSComCtl2.UpDown udp 
         Height          =   300
         Left            =   4350
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   165
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt(0)"
         BuddyDispid     =   196610
         BuddyIndex      =   0
         OrigLeft        =   4395
         OrigTop         =   165
         OrigRight       =   4635
         OrigBottom      =   465
         Max             =   12
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chk 
         Caption         =   "ʱ��(&H)"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   24
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��(&T)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�� �� ��(&M)"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   5
         Top             =   600
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Сʱ"
         Height          =   180
         Left            =   4650
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmOpsStationArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'���������弶��������

Private mblnReading As Boolean
Private mblnDataChanged As Boolean
Private mblnOK As Boolean
Private mlngKey As Long
Private mlngDeptKey As Long
Private mfrmMain As Form
Private mstrPrivs As String
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Form, Optional lngKey As Long = 0, Optional lngDeptKey As Long = 0, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ��򿪱༭����������ݵ��������޸Ĳ���
    '������
    '���أ�
    '******************************************************************************************************************
    mlngKey = lngKey
    mlngDeptKey = lngDeptKey
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    Call ExecuteCommand("��ȡ����")
    
    DataChanged = False
    
    Me.Show 1, mfrmMain
    
    ShowEdit = mblnOK
    
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ����������޸ĵ����ݽ��кϷ���У��
    '���أ�У��Ϸ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset

    If chk(1).Value = 1 Then
        If Val(txt(0).Text) < 1 Or Val(txt(0).Text) > 12 Then
            ShowSimpleMsg "����ʱ���������1Сʱ��С��12Сʱ��"
            
            zlControl.TxtSelAll txt(0)
            txt(0).SetFocus
            Exit Function
        End If
    End If
    
    gstrSQL = "SELECT 1 FROM ҽ��ִ�з��� WHERE ִ�м�=[1] AND ����id=[2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txt(1).Text, mlngDeptKey)
    
    If rs.BOF Then
        ShowSimpleMsg "������������һ�������ڵ������䣡"
        zlControl.TxtSelAll txt(1)
        txt(1).SetFocus
        Exit Function
    End If
    
    '�������ʱ��������ʱ��Ĺ�ϵ
    gstrSQL = "SELECT ����ʱ�� FROM ����ҽ����¼ WHERE ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        If Format(dtp.Value, "yyyy-MM-dd HH:mm") < Format(rs("����ʱ��").Value, "yyyy-MM-dd HH:mm") Then
            
            If MsgBox("������ʼʱ��(" & Format(dtp.Value, "yyyy-MM-dd HH:mm") & ")��������ʱ��(" & Format(rs("����ʱ��").Value, "yyyy-MM-dd HH:mm") & ")" & vbCrLf & "�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                dtp.SetFocus
                Exit Function
            End If
            
        End If
    End If
    
    '���һ�������Ƿ���ͬһʱ���������������
    gstrSQL = "SELECT 1 FROM ����������¼ B " & _
                "WHERE B.����״̬ In (2,3) AND  " & _
                       "B.ҽ��id <> [3] AND  " & _
                       "(B.����id, NVL(B.��ҳid,0)) IN (SELECT ����id, NVL(��ҳid,0) FROM ����ҽ����¼ WHERE ID = [3]) AND  " & _
                       "((B.������ʼʱ�� BETWEEN [1] AND [2]) OR  " & _
                       "(B.��������ʱ�� BETWEEN [1] AND [2]))"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(DateAdd("h", Val(txt(0).Text), dtp.Value), "YYYY-MM-DD HH:MM:SS")), mlngKey)
    If rs.BOF = False Then
        ShowSimpleMsg "��ǰ���˲���ͬʱ���ж���������"
        dtp.SetFocus
        Exit Function
    End If
    
    ValidData = True
    
End Function

Private Function SaveData() As Boolean
    '******************************************************************************************************************
    '���ܣ����������޸ĺ�����ݽ��б���/���´���
    '���������ز�lngKey����ʾ���¼�¼�Ĺؼ���
    '���أ�����ɹ�����True�����򷵻�False
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim rsSQL As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim str��Ⱦ���� As String
    Dim blnTrans As Boolean
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    With vsf
        For lngLoop = 1 To .Rows - 1
            If .TextMatrix(lngLoop, .ColIndex("��λ")) <> "" And .TextMatrix(lngLoop, .ColIndex("����")) <> "" Then
                strTmp = strTmp & ";" & Val(.RowData(lngLoop)) & "," & .TextMatrix(lngLoop, .ColIndex("��λ")) & "," & .TextMatrix(lngLoop, .ColIndex("����")) & "," & .TextMatrix(lngLoop, .ColIndex("���"))
            End If
        Next
    End With
    
    If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    
    For lngLoop = 0 To lst.ListCount - 1
        If lst.Selected(lngLoop) Then
            str��Ⱦ���� = str��Ⱦ���� & ";" & lst.List(lngLoop)
        End If
    Next
    If str��Ⱦ���� <> "" Then str��Ⱦ���� = Mid(str��Ⱦ����, 2)
    
    gstrSQL = "zl_����������¼_Arrange(" & mlngKey & "," & _
                                        "To_Date('" & Format(dtp.Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                        IIf(chk(1).Value = 0, "Null", "TO_DATE('" & Format(DateAdd("h", Val(txt(0).Text), dtp.Value), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')") & ",'" & _
                                        txt(1).Text & "'," & _
                                        mlngDeptKey & ",'" & _
                                        strTmp & "',2,'" & cbo(0).Text & "'," & Val(txt(2).Text) & ",'" & zlCommFun.GetNeedName(cbo(1).Text) & "'," & chk(2).Value & ",'" & str��Ⱦ���� & "'," & chk(3).Value & ")"
    Call SQLRecordAdd(rsSQL, gstrSQL)
    
    gstrSQL = "Zl_����������¼_Updateadvice(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, gstrSQL)
    
    '����SQLִ��,�ɾ���ִ��ʱ����
    '------------------------------------------------------------------------------------------------------------------
    If mfrmMain.NotAutoCharge = False Then
        gstrSQL = "Select b.ҽ��id From ����ҽ����¼ a,����������¼ b Where b.ID=[1] And b.ҽ��id=a.ID And a.ҽ��״̬ Not In (4,8)"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            Call SQLRecordAdd(rsSQL, "", 0, 1, Val(rs("ҽ��id").Value))
        End If
    End If
            
    '��ʼִ��SQL,���ύ�����ݿ���
    '------------------------------------------------------------------------------------------------------------------
    If rsSQL.RecordCount > 0 Then
        
        rsSQL.MoveFirst
        
        blnTrans = True
        gcnOracle.BeginTrans

        For lngLoop = 1 To rsSQL.RecordCount
            
            If Val(rsSQL("Trans").Value) = 1 And blnTrans = False Then
                blnTrans = True
                gcnOracle.BeginTrans
            End If
            
            '���������Ŀ�ķ���,
            If Val(rsSQL("Custom").Value) = 1 Then
                If CreateOrderCharge(Val(rsSQL("Parameter").Value), mstrPrivs) = False Then
                    blnTrans = False
                    gcnOracle.RollbackTrans
                    Exit Function
                End If
            Else
                If Val(rsSQL("Trans").Value) = 2 Then
                    If blnTrans Then
                        gcnOracle.CommitTrans
                        blnTrans = False
                        SaveData = True
                    End If
                Else
                    gstrSQL = CStr(rsSQL("SQL").Value)
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            End If
            rsSQL.MoveNext
        Next
        
        If blnTrans Then
            gcnOracle.CommitTrans
            blnTrans = False
            SaveData = True
        End If
    End If
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)

            Call .AppendColumn("��λ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���", 900, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .InitializeEdit(True, True, True)
            
            Call .InitializeEditColumn(.ColIndex("��λ"), True, vbVsfEditCombox)
            Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
            
            .IndicatorCol = 0
            Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            
            .AppendRows = True
        End With
        txt(1).BackColor = COLOR.��ɫ

    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        '���������̶�
        '--------------------------------------------------------------------------------------------------------------
        With cbo(0)
            .Clear
            .AddItem ""
            strSQL = "Select ����,����,����,ȱʡ��־ From ���������̶�"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("����").Value & "-" & rs("����").Value
                    If rs("ȱʡ��־").Value = 1 Then .ListIndex = .NewIndex
                    rs.MoveNext
                Loop
            End If
            If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
        End With
            

        '�������ʷ���
        '--------------------------------------------------------------------------------------------------------------
        With cbo(1)
            .Clear
            .AddItem ""
            strSQL = "Select ����,����,����,ȱʡ��־ From �������ʷ���"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("����").Value & "-" & rs("����").Value
                    If rs("ȱʡ��־").Value = 1 Then .ListIndex = .NewIndex
                    rs.MoveNext
                Loop
            End If
            If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
        End With
        
        
        '������Ⱦ����
        '--------------------------------------------------------------------------------------------------------------
        With lst
            .Clear
            strSQL = "Select ����,����,���� From ������Ⱦ����"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    .AddItem rs("����").Value
                    rs.MoveNext
                Loop
            End If
        End With
        
        '������λ
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT ����||'-'||���� As ���� FROM ������λ Order by ����"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
        Call mclsVsf.InitializeEditColumn(mclsVsf.ColIndex("��λ"), True, vbVsfEditCombox, vsf.BuildComboList(rs, "����", "����"))
        
        
        dtp.Value = Format(zlDatabase.Currentdate + 1, dtp.CustomFormat)
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
        
        mblnReading = True
        
        
        mblnReading = False
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        gstrSQL = "SELECT a.�����̶�,a.��̨����,a.��������,a.��Ⱦ����,a.��Ⱦ����,a.��Ⱦ����,a.������ʼʱ��,Ceil((a.��������ʱ��-a.������ʼʱ��)*24) As Сʱ,a.������ FROM ����������¼ a Where a.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
        If rs.BOF = False Then
            
            zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("�����̶�").Value)
            txt(2).Text = zlCommFun.NVL(rs("��̨����").Value, 0)
            chk(0).Value = IIf(Val(txt(2).Text) > 0, 1, 0)
            
            Call zlControl.CboLocate(cbo(1), zlCommFun.NVL(rs("��������").Value))

            chk(2).Value = zlCommFun.NVL(rs("��Ⱦ����").Value, 0)
            If chk(2).Value = 1 Then
                strTmp = ";" & zlCommFun.NVL(rs("��Ⱦ����").Value) & ";"
                For intLoop = 0 To lst.ListCount - 1
                    If InStr(strTmp, ";" & lst.List(intLoop) & ";") > 0 Then
                        lst.Selected(intLoop) = True
                    End If
                Next
            End If
            
            chk(3).Value = zlCommFun.NVL(rs("��Ⱦ����").Value, 0)
            
            txt(1).Text = zlCommFun.NVL(rs("������").Value)
            dtp.Value = Format(zlCommFun.NVL(rs("������ʼʱ��").Value, zlDatabase.Currentdate), dtp.CustomFormat)
            txt(0).Text = zlCommFun.NVL(rs("������").Value, 1)
            
        End If
        
        
        '��ȡ�Ѱ��ŵ�������Ա
        '--------------------------------------------------------------------------------------------------------------
        mclsVsf.ClearGrid
        gstrSQL = "Select A.��Աid As ID,a.��λ,B.���,a.���� From ����������Ա a,��Ա�� b,������λ c Where c.����=a.��λ And Nvl(a.�ڼ�,1)=1 And a.��¼id=[1] And a.��Աid=b.ID(+) order by c.����"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
        
    End Select

    ExecuteCommand = True

    Exit Function
    
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub chk_Click(Index As Integer)
    If Index = 1 Then
        txt(0).Visible = (chk(Index).Value = 1)
        udp.Visible = txt(0).Visible
        Label2.Visible = txt(0).Visible
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 1      '����ִ�м�
        
        gstrSQL = "Select RowNum As ID,ִ�м�,Decode(b.������,Null,'����',Decode(b.����״̬,2,'Ԥ��',3,'����')) As ״̬" & vbNewLine & _
                    "From ҽ��ִ�з��� a," & vbNewLine & _
                    "     (" & vbNewLine & _
                    "      Select ������,Max(����״̬) As ����״̬" & vbNewLine & _
                    "      From ����������¼" & vbNewLine & _
                    "      Where Not (��������ʱ��<[2] OR ������ʼʱ��>[3]) AND ������id=[1] AND ����״̬ In (2,3) Group By ������" & vbNewLine & _
                    "     ) b" & vbNewLine & _
                    "Where a.����id=[1]" & vbNewLine & _
                    "      And a.ִ�м�=b.������(+)"
                        
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngDeptKey, CDate(Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(DateAdd("h", Val(txt(0).Text), dtp.Value))))
 
        If ShowPubSelect(Me, txt(1), 2, "ִ�м�,2100,0,;״̬,900,0,", Me.Name & "\����ִ�м�ѡ��", "����±���ѡ��һ������ִ�м�", rsData, rs, 3600, 4200) = 1 Then
            txt(1).Text = zlCommFun.NVL(rs("ִ�м�").Value)
            DataChanged = True
        End If

    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If ValidData = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    DataChanged = False
    
    Unload Me
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf
        If .TextMatrix(Row, .ColIndex("��λ")) = "" And .TextMatrix(Row, .ColIndex("����")) = "" Then
            Cancel = True
        End If
    End With
End Sub

Private Sub txt_Change(Index As Integer)
    If mblnReading Then Exit Sub
    
    DataChanged = True

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub
    
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
    
    With vsf
        Select Case Col
        Case .ColIndex("��λ")
            If .ComboIndex > -1 Then
                .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.ComboItem(.ComboIndex))
            End If
        End Select
    End With
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    Dim strTmp As String
    
    With vsf
        If Col = .ColIndex("����") Then

            strTmp = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("��λ")))
            
            gstrSQL = "Select �Ƿ�Ψһ,�Ƿ�ҽ��,�Ƿ�ʿ From ������λ Where ����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTmp)
            If rs.BOF = False Then
                If zlCommFun.NVL(rs("�Ƿ�ҽ��").Value, 0) = 1 Then strTmp = "ҽ��"
                If zlCommFun.NVL(rs("�Ƿ�ʿ").Value, 0) = 1 Then strTmp = "��ʿ"
            Else
                strTmp = "ҽ��"
            End If
                
            gstrSQL = GetPublicSQL(SQL.��Ա����ѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strTmp, mlngDeptKey, mlngKey)
            bytRet = ShowPubSelect(Me, vsf, 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,;״̬,900,0,", Me.Name & "\��Ա����ѡ��", "����±���ѡ��һ��������Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        
            If bytRet = 1 Then
            
'                If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                    ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
'                    Exit Sub
'                End If
                       
                .EditText = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                DataChanged = True
    
            End If
            
        End If
    End With
End Sub

Private Sub vsf_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    
    With vsf
        Select Case Col
        Case .ColIndex("��λ")
            
            Call mclsVsf.ComboLocation(Row, Col)

        End Select
    End With
End Sub

Private Sub vsf_DblClick()
    Call mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
    
    
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    Dim bytRet As Byte
    Dim strDoctor As String
    
    With vsf
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("����") Then
            
                If InStr(.EditText, "'") > 0 Then
                    KeyCode = 0
                    .EditText = ""
                    Exit Sub
                End If
    
                strTmp = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("��λ")))
                
                gstrSQL = "Select �Ƿ�Ψһ,�Ƿ�ҽ��,�Ƿ�ʿ From ������λ Where ����=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTmp)
                If rs.BOF = False Then
                    If zlCommFun.NVL(rs("�Ƿ�ҽ��").Value, 0) = 1 Then strDoctor = "ҽ��"
                    If zlCommFun.NVL(rs("�Ƿ�ʿ").Value, 0) = 1 Then strDoctor = "��ʿ"
                Else
                    strDoctor = "ҽ��"
                End If
            
                strText = UCase(.EditText)
                bytMode = GetApplyMode(strText)
                
                strText = strText & "%"
                strTmp = IIf(ParamInfo.��Ŀ����ƥ�䷽ʽ = 1, strText, "%" & strText)
                
                gstrSQL = GetPublicSQL(SQL.��Ա���Ź���, bytMode)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strDoctor, mlngDeptKey, mlngKey, strText, strTmp)
    
                If ShowPubSelect(Me, vsf, 2, "���,1200,0,;����,1200,0,;����,900,0,;����,1200,0,;״̬,900,0,", Me.Name & "\��Ա���Ź���", "����±���ѡ��һ����Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

'                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
'                        ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
'                        Exit Sub
'                    End If
                           
                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                Else
                    .Cell(flexcpData, Row, Col) = .EditText
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    DataChanged = True
                End If

            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    Call mclsVsf.KeyPress(KeyAscii)
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf.MouseRow, vsf.MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub


