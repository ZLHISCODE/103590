VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPhysicalSel 
   BorderStyle     =   0  'None
   Caption         =   "1"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2475
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraMK 
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         ItemData        =   "frmPhysicalSel.frx":0000
         Left            =   1200
         List            =   "frmPhysicalSel.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         ItemData        =   "frmPhysicalSel.frx":001A
         Left            =   1200
         List            =   "frmPhysicalSel.frx":0024
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "2015-02-01"
         Top             =   795
         Width           =   1455
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         ItemData        =   "frmPhysicalSel.frx":0034
         Left            =   3960
         List            =   "frmPhysicalSel.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   4
         Left            =   3000
         TabIndex        =   17
         ToolTipText     =   "���ڸ�ʽ:YYYY-MM-DD"
         Top             =   840
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����\����"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   840
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ι�"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   14
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   3960
         TabIndex        =   13
         ToolTipText     =   "���￪ʼ����"
         Top             =   1080
         Width           =   90
      End
   End
   Begin VB.Timer tmrThis 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2490
      Index           =   1
      Left            =   9120
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   0
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   2490
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   8775
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   0
      Width           =   8655
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8760
      TabIndex        =   6
      Top             =   1875
      Width           =   8760
      Begin VB.Timer tmrAir 
         Interval        =   1000
         Left            =   2040
         Top             =   120
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "�˳�(&Q)"
         Height          =   350
         Left            =   7080
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.ListBox lstItem 
      Height          =   1740
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VSFlex8Ctl.VSFlexGrid vsSymptom 
      Height          =   1740
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   6615
      _cx             =   11668
      _cy             =   3069
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
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   11250603
      GridColorFixed  =   0
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
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   400
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPhysicalSel.frx":004E
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
      WordWrap        =   -1  'True
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
      BackColorFrozen =   16777215
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   2160
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line lin 
      Index           =   7
      X1              =   5760
      X2              =   6435
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   5880
      X2              =   6555
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   5880
      X2              =   6555
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   5880
      X2              =   6555
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   5880
      X2              =   6555
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   5880
      X2              =   6555
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   5760
      X2              =   6435
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   5880
      X2              =   6555
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmPhysicalSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mstrPhysical As String
'Private mobjAir As clsAirBubble zl9ComLib.clsAirBubble
Private mobjAir As Object
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mbytSex As Integer    '�Ա� 0-��,1-Ů
Private mbyt���� As Byte      '1-����༭��2-סԺ�༭
Private mstrDelOrder As String   '��¼ɾ��֢״��¼���:���1�����2��...
Private mlngNum As Long  '��¼ʱ���޸�λ��
Private mIntWaitTime As Integer   '��¼�����ӳ�ʱ�䣬���ڵ�������ʱ���˵ĵ�һ��������Picture,�������ݲ����Զ��ӳ�
'֢״�к�
Private Enum COL֢״�к�
    COL_��� = 0
    COL_״̬ = 1
    col_֢״ = 2
    col_��ʼ���� = 3
    col_�������� = 4
    COL_ҽ�� = 5
    COL_���� = 6
End Enum

Private Enum CBO_IX
    Ix_�ι� = 0
    IX_���� = 1
    IX_���� = 2
    IX_���� = 3
End Enum

Public Sub ShowMe(ByRef objMap As clsPassMap)
'����:
'����:
'      byt����-1-����༭��2-סԺ�༭
    Dim objPoint As RECT
    
    mlng����ID = objMap.PassPati.lng����ID
    mlng��ҳID = IIf(objMap.lngModel = PM_����༭, objMap.PassPati.lng�Һ�ID, objMap.PassPati.lng��ҳID)
    mbytSex = IIf(objMap.PassPati.str�Ա� = "Ů", 1, 0)
    mbyt���� = IIf(objMap.lngModel = PM_����༭, 1, 2)
    
    Call GetWindowRect(objMap.objCmdBar.hWnd, objPoint)
    If gbytPass = DT Or gbytPass = YWS Or gbytPass = ZL Then
        Me.Width = 2040
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY + objMap.objCmdBar.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX - Me.Width + objMap.objCmdBar.Width
    ElseIf gbytPass = TYT Then
        Me.Width = 8760
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY + objMap.objCmdBar.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX - Me.Width + objMap.objCmdBar.Width
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        Me.Width = 6150: Me.Height = 2100
        Me.Top = objPoint.Top * Screen.TwipsPerPixelY + objMap.objCmdBar.Height
        Me.Left = objPoint.Left * Screen.TwipsPerPixelX - Me.Width + objMap.objCmdBar.Width
        
        lstItem.Visible = False
        vsSymptom.Visible = False
        fraMK.Visible = True
    End If
    Me.Show 1, objMap.frmMain
End Sub

Private Sub LoadDict()
'����:���ز���������ֵ�����
'����:bytSex:0-��,1-Ů
    Dim strSQL As String, i As Long
    Dim strList As String
    Dim rsDict As ADODB.Recordset
    If gbytPass <> ZL Then
        strSQL = "Select ���� From ��������� Order by ����"
    Else
        strList = "����,����,�ι��ܲ�ȫ,���ظι��ܲ�ȫ,�����ܲ�ȫ,���������ܲ�ȫ"
        strSQL = "Select /*+cardinality(A,10)*/ a.Column_Value As ���� From Table(f_Str2list([1])) A"
    End If
    On Error GoTo errH
    Set rsDict = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strList)
    lstItem.Clear
    With rsDict
        For i = 1 To .RecordCount
            If !���� = "�и�" Or !���� = "������" Or !���� = "����" Or !���� = "����" Then
                If mbytSex = 1 Then lstItem.AddItem !����
            Else
                lstItem.AddItem !����
            End If
            .MoveNext
        Next
    End With
    Exit Sub

errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub LoadLists()
'����:���ز��˵Ĳ��������
'����:bytSex:0-��,1-Ů
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim lngTmp As Long
    Dim strTmp As String

    If gbytPass = YWS Or gbytPass = DT Or gbytPass = TYT Or gbytPass = ZL Then
       Call LoadDict
    
       If mbyt���� = 1 Then
           lngTmp = Val(zlDatabase.GetPara(21, glngSys))
           strSQL = "Select ���������" & vbNewLine & _
                   "From ���˹Һż�¼" & vbNewLine & _
                   "Where ����id = [1] And �Ǽ�ʱ�� > Trunc(Sysdate-[2]) And ��������� Is Not Null And Rownum = 1"
       Else
           strSQL = "Select ��Ϣֵ As ���������" & vbNewLine & _
                   "From ������ҳ�ӱ� Where ����id = [1] And ��ҳid = [2] And ��Ϣ�� = '���������'"
       End If
       
       On Error GoTo errH
       Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, IIf(mbyt���� = 1, lngTmp, mlng��ҳID))
       If rsTmp.RecordCount > 0 Then
           For i = 0 To lstItem.ListCount - 1
               lstItem.Selected(i) = InStr(1, "," & rsTmp!��������� & ",", "," & lstItem.List(i) & ",") > 0
           Next
       End If
       
       mstrPhysical = GetLists
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        '��������
        txtDate.Text = ""
        cboStatus(Ix_�ι�).Clear
        cboStatus(IX_����).Clear
        cboStatus(IX_����).Clear
        '-1-��ȷ����Ĭ�ϣ�,0-�޸���,1-�ι��ܲ�ȫ,2-��ȸ���,3-�жȸ���,4-�ضȸ���"
        strTmp = ",�޸���,�ι��ܲ�ȫ,��ȸ���,�жȸ���,�ضȸ���"
        For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
            cboStatus(Ix_�ι�).AddItem Split(strTmp, ",")(i)
        Next
        cboStatus(Ix_�ι�).ListIndex = -1 'ȱʡ
        
        '-1-��ȷ����Ĭ�ϣ���0-�����𺦣�1-�����ܲ�ȫ��2-������𺦣�3-�ж����𺦣�4-�ض�����
        strTmp = ",������,�����ܲ�ȫ,�������,�ж�����,�ض�����"
        For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
            cboStatus(IX_����).AddItem Split(strTmp, ",")(i)
        Next
        cboStatus(IX_����).ListIndex = -1
        
        If mbytSex = 1 Then 'Ů�Բ���
            '����\����
            strTmp = ",����,����"
            For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
                cboStatus(IX_����).AddItem Split(strTmp, ",")(i)
            Next
            cboStatus(IX_����).ListIndex = -1
            
            cboStatus(IX_����).Enabled = True
            cboStatus(IX_����).BackColor = &H80000005
        Else
            cboStatus(IX_����).Enabled = False
            cboStatus(IX_����).BackColor = &H8000000C
        End If
        '����ȱʡ����
        txtDate.Enabled = False
        txtDate.BackColor = &H8000000C
        
        strTmp = Get���˲��������(mlng����ID, IIf(mbyt���� = 1, 0, mlng��ҳID))
        mstrPhysical = strTmp
        If strTmp <> "" Then
            For i = LBound(Split(strTmp, ",")) To UBound(Split(strTmp, ","))
                If InStr(",�޸���,�ι��ܲ�ȫ,��ȸ���,�жȸ���,�ضȸ���,", "," & Split(strTmp, ",")(i) & ",") > 0 Then
                    zlControl.CboLocate cboStatus(Ix_�ι�), Split(strTmp, ",")(i)
                ElseIf InStr(",������,�����ܲ�ȫ,�������,�ж�����,�ض�����,", "," & Split(strTmp, ",")(i) & ",") > 0 Then
                    zlControl.CboLocate cboStatus(IX_����), Split(strTmp, ",")(i)
                ElseIf InStr(",����,����,", "," & Split(strTmp, ",")(i) & ",") > 0 Then
                    zlControl.CboLocate cboStatus(IX_����), Split(strTmp, ",")(i)
                ElseIf InStr("," & Split(strTmp, ",")(i), ",��������|") > 0 Then '��������|2015-02-01
                    strTmp = Split(strTmp, ",")(i)
                    txtDate.Text = Split(strTmp, "|")(1)
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetLists() As String
'���ܣ���ȡѡ��Ĳ���������ַ������Զ��ŷָ�
    Dim i As Long, strRetu As String
    
    For i = 0 To lstItem.ListCount - 1
        If lstItem.Selected(i) Then strRetu = strRetu & "," & lstItem.List(i)
    Next
    
    If strRetu <> "" Then GetLists = Mid(strRetu, 2)
End Function

Private Sub cmdQuit_Click()
    '���
    If CheckCell Then Exit Sub
    '��������
    Call SaveData
    
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyQ = KeyCode And Shift = vbCtrlMask Then
        Call cmdQuit_Click
    End If
End Sub

Private Sub Form_Load()
    
    Call LoadLists
    If gbytPass = TYT Then
        '��ʼ��֢״��
        Call InitSymptom
        '��������
        Call LoadSyptoms
        'Set mobjAir = New gobjComLib.clsAirBubble
        Set mobjAir = CreateObject("zl9ComLib.clsAirBubble")
    End If
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    '����߿�����
    Call InitFormBorder
    If gbytPass = DT Or gbytPass = YWS Then
        lstItem.Top = fraBorder(0).Height + 80
        lstItem.Left = fraBorder(3).Width + 80
        vsSymptom.Visible = False
    ElseIf gbytPass = TYT Then
        lstItem.Top = fraBorder(0).Height + 80
        lstItem.Left = fraBorder(3).Width + 80
        vsSymptom.Top = fraBorder(0).Height + 80
        vsSymptom.Left = fraBorder(3).Width + 80 + lstItem.Width + 80
    ElseIf gbytPass = MK And gstrVersion = "4.0" Then
        fraMK.Top = fraBorder(0).Height + 80
        fraMK.Left = fraBorder(3).Width + 80
    End If
    cmdQuit.Left = picBottom.Width - cmdQuit.Width - 200
    
End Sub

Private Sub SaveData()
    Dim strTmp As String
    Dim bytFunc As Byte
    Dim arrSQL As Variant
    Dim curDate As Date
    Dim i As Long
    arrSQL = Array()
    
    If gbytPass = MK And gstrVersion = "4.0" Then
        strTmp = IIf(cboStatus(Ix_�ι�).Tag <> "", "," & cboStatus(Ix_�ι�).Tag, "")
        strTmp = strTmp & IIf(cboStatus(IX_����).Tag <> "", "," & cboStatus(IX_����).Tag, "")
        strTmp = strTmp & IIf(cboStatus(IX_����).Tag <> "", "," & cboStatus(IX_����).Tag, "")
        strTmp = strTmp & IIf(txtDate.Tag <> "", "," & txtDate.Tag, "")
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    Else
        strTmp = GetLists
    End If
    
    If strTmp <> mstrPhysical Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        If mbyt���� = 1 Then    '1-����༭
            arrSQL(UBound(arrSQL)) = "Zl_���˲��������_Insert(" & mlng����ID & ",0," & mlng��ҳID & ",'" & strTmp & "')"   '��ʱmlng��ҳID������Һ�ID
        Else    '2-סԺ�༭
            arrSQL(UBound(arrSQL)) = "Zl_���˲��������_Insert(" & mlng����ID & "," & mlng��ҳID & ",Null,'" & strTmp & "')"
        End If
    End If

    If gbytPass = TYT Then
        '��װɾ��sql
        If mstrDelOrder <> "" Then
            For i = 0 To UBound(Split(mstrDelOrder, ",")) - 1    '���һ����ȡ
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����֢״��¼_Update(3," & mlng����ID & "," & mlng��ҳID & "," & Split(mstrDelOrder, ",")(i) & ")"
            Next
        End If
        curDate = zlDatabase.Currentdate
        With vsSymptom
            For i = .FixedRows To .Rows - 2  '���һ�пհ�
                bytFunc = Val(.TextMatrix(i, COL_״̬))
                If bytFunc = 2 Then  '���� ����ڹ�����ȡ���ֵ����
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����֢״��¼_Update(1," & mlng����ID & "," & mlng��ҳID & "," & Val(.TextMatrix(i, COL_���)) & " ,'" & _
                                             .RowData(i) & "','" & .TextMatrix(i, col_֢״) & "',To_Date('" & .TextMatrix(i, col_��ʼ����) & "','YYYY-MM-DD HH24:MI:SS')," & _
                                             "To_date('" & .TextMatrix(i, col_��������) & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.���� & _
                                             "',To_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'))"
                ElseIf bytFunc = 3 Then   '�޸�
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_����֢״��¼_Update(2," & mlng����ID & "," & mlng��ҳID & "," & Val(.TextMatrix(i, COL_���)) & " ,'" & _
                                             .RowData(i) & "','" & .TextMatrix(i, col_֢״) & "',To_Date('" & .TextMatrix(i, col_��ʼ����) & "','YYYY-MM-DD HH24:MI:SS')," & _
                                             "To_date('" & .TextMatrix(i, col_��������) & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.���� & _
                                             "',To_date('" & Format(curDate, "yyyy-mm-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'))"
                End If
            Next

        End With
    End If

    On Error GoTo errH
    '����ִ��ɾ�������޸ģ���β����� ������Ϊ���������
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����״̬")
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitSymptom()
'����: ��ʼ������֢״��¼��ͷ
    Dim strCol As String, arrHead As Variant
    Dim i As Long
    '״̬: 0-δ���,1-ԭʼ��2-������3-�޸�
    strCol = "���;״̬;֢״,2000,4;��ʼ����,1300,4;��������,1300,4;ҽ��,1000,4;����,50,1"
    arrHead = Split(strCol, ";")
    With vsSymptom
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows + 1    'ȱʡ��ʾһ�пհ�

        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)

            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrDelOrder = ""
    mstrPhysical = ""
    If Not mobjAir Is Nothing Then
        mobjAir.CloseAirBubble
        Set mobjAir = Nothing
    End If
End Sub

Private Sub tmrAir_Timer()
'����:�������ݵ�ʱ������mIntWaitTime=3
    If mIntWaitTime = 0 Then Exit Sub
    mIntWaitTime = mIntWaitTime - 1
    If mIntWaitTime = 1 Then
        If Not mobjAir Is Nothing Then
            mobjAir.CloseAirBubble
        End If
        mIntWaitTime = 0
    End If
End Sub

Private Sub tmrThis_Timer()
    Dim lngTmp As Long
    With vsSymptom
        If .Col = col_��ʼ���� Or .Col = col_�������� Then
            lngTmp = .EditSelStart
            If .EditSelText = "-" Then
                Call Vs_EditSelChange(lngTmp - 1)    'ѡ����"-"
            ElseIf lngTmp = 0 Or lngTmp = 5 Or lngTmp = 8 Then
                mlngNum = 0
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strDate As String

    With vsSymptom
        If .TextMatrix(Row, col_֢״) <> "" Then
            '����Ĭ��ֵ
            If .TextMatrix(Row, col_��ʼ����) = "" And .TextMatrix(Row, col_��������) = "" Then
                strDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
                .TextMatrix(Row, col_��ʼ����) = strDate
                .TextMatrix(Row, col_��������) = strDate
            End If
            If .TextMatrix(Row, col_��ʼ����) <> "" And .TextMatrix(Row, col_��������) <> "" _
               And (.TextMatrix(Row, COL_ҽ��) = "" Or .Cell(flexcpData, Row, COL_ҽ��) <> UserInfo.����) Then
                .TextMatrix(Row, COL_ҽ��) = UserInfo.����
                .Cell(flexcpAlignment, Row, COL_ҽ��) = flexAlignLeftCenter
            End If
        End If
        '״̬����
        If .TextMatrix(Row, COL_״̬) = "1" Then
            If .TextMatrix(Row, Col) <> .Cell(flexcpData, Row, Col) Then
                .TextMatrix(Row, COL_״̬) = "3"   '3-�޸�
            End If
        ElseIf .TextMatrix(Row, COL_״̬) = "" And .TextMatrix(Row, COL_ҽ��) <> "" Then  'ҽ��¼�����һ������¼�����
            .TextMatrix(Row, COL_״̬) = "2"   '2-����
        End If

    End With
End Sub

Private Sub vsSymptom_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsSymptom

        If col_֢״ = NewCol Then
            .ColComboList(col_֢״) = "..."
            .FocusRect = flexFocusLight
        Else
            .ColComboList(col_֢״) = ""
            .FocusRect = flexFocusLight
        End If

        If .TextMatrix(.Row, col_֢״) <> "" And .TextMatrix(.Rows - 1, col_֢״) <> "" Then
            .Rows = .Rows + 1
        End If
    End With
End Sub

Private Sub vsSymptom_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With vsSymptom
        '��꿿��
        If Row > .FixedRows Then
            .Cell(flexcpAlignment, Row, Col) = flexAlignLeftCenter
        End If
        'ҽ���������в��ɱ༭
        If COL_ҽ�� = Col Or COL_���� = Col Then
            Cancel = True
        End If
        
        If col_��ʼ���� = Col Or Col = col_�������� Then
            If .TextMatrix(Row, col_֢״) = "" Then
                Cancel = True '����֢״��Ϊ�յ����������༭
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSymptom As String

    If col_֢״ = Col Then
        On Error Resume Next
        strSymptom = gobjPass.inputDiagside
        If Err.Number <> 0 Then
            MsgBox "̫Ԫͨ�ӿڵ���ʧ�ܣ������Ƿ���������", vbInformation + vbOKOnly, Me.Caption
        End If
        Err.Clear: On Error GoTo 0
        If strSymptom <> "" Then
            vsSymptom.RowData(Row) = Val(Split(strSymptom, ";")(0))
            vsSymptom.TextMatrix(Row, Col) = Split(strSymptom, ";")(1)
            Call vsSymptom_AfterEdit(Row, Col)
        End If

    End If
End Sub

Private Sub vsSymptom_Click()
    Dim i As Long

    With vsSymptom
        If .Col = COL_���� And Not .Cell(flexcpPicture, .Row, .Col) Is Nothing Then
            If .Rows - 1 = .FixedRows Then
                .Cell(flexcpText, .Row, col_֢״, .Row, COL_����) = ""
            Else
                If MsgBox("ȷ��Ҫɾ��֢״��" & .TextMatrix(.Row, col_֢״) & "����", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                '����ɾ����SQL
                If Val(.TextMatrix(.Row, COL_״̬)) = 1 Or Val(.TextMatrix(.Row, COL_״̬)) = 3 Then
                    mstrDelOrder = mstrDelOrder & Val(.TextMatrix(.Row, COL_���)) & ","
                End If
                'ɾ����ʾ��
                .RemoveItem (.Row)
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsSymptom
        If KeyCode = vbKeyF1 Then
            If .Col = col_֢״ Then
            .ColComboList(vsSymptom.Col) = ""
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col_֢״) <> "" Then
                If MsgBox("ȷ��Ҫɾ��֢״��" & .TextMatrix(.Row, col_֢״) & "����", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                '����ɾ����SQL
                If Val(.TextMatrix(.Row, COL_״̬)) = 1 Or Val(.TextMatrix(.Row, COL_״̬)) = 3 Then
                    mstrDelOrder = mstrDelOrder & Val(.TextMatrix(.Row, COL_���)) & ","
                End If
                'ɾ����ʾ��
                .RemoveItem (.Row)
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then 'delete����del������һ��
        Call vsSymptom_KeyPressEdit(Row, Col, vbKeyDelete)
    End If
End Sub

Private Sub vsSymptom_KeyPress(KeyAscii As Integer)

    With vsSymptom
        If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
            KeyAscii = 0
            If .Col <> COL_ҽ�� Then
                .TextMatrix(.Row, .Col) = ""
            End If
        ElseIf KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            Call EnterNextCell
            If .Row = .Rows - 1 And .TextMatrix(.Row, col_֢״) = "" And .Col >= col_�������� Then
                cmdQuit.SetFocus
            End If
        End If
    End With
End Sub

Private Sub vsSymptom_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strChr As String
    Dim lngTmp As Long

    With vsSymptom

        If KeyAscii = vbKeyBack Then
            If col_֢״ = Col And .ColComboList(col_֢״) = "" Then
                .EditText = ""
            End If

            If Col = col_��ʼ���� Or Col = col_�������� Then
                If .EditText <> "" Then
                    If Len(.EditText) = .EditSelStart Then    '��������
                        .EditText = Left(.EditText, .EditSelStart - 1)
                    ElseIf Len(.EditText) > .EditSelStart And .EditSelLength = 0 Then    '������м�
                        lngTmp = .EditSelStart
                        If lngTmp <> 0 Then
                            .EditText = Mid(.EditText, 1, lngTmp - 1) & Mid(.EditText, lngTmp)
                            .EditSelStart = lngTmp
                        End If
                        Exit Sub
                    ElseIf Len(.EditText) = .EditSelLength Then    'ȫѡ��
                        .EditText = ""
                    ElseIf .EditSelText <> "-" And .EditSelLength <> 0 Then
                        If .EditSelLength = 4 Then
                            .EditText = "2000" & Mid(.EditText, 5)
                            lngTmp = 4
                        ElseIf .EditSelStart <= 7 Then
                            .EditText = Left(.EditText, 5) & "01" & Right(.EditText, 3)
                            lngTmp = 7
                        Else
                            .EditText = Left(.EditText, 8) & "01"
                            lngTmp = 8
                        End If
                        Call Vs_EditSelChange(lngTmp)
                    End If
                End If
            End If
        ElseIf KeyAscii = vbKeyReturn Then
            KeyAscii = 0

            Call EnterNextCell: Exit Sub

        ElseIf KeyAscii = vbKeyDelete Then
            KeyAscii = 0
            .EditText = Mid(.EditText, 1, .EditSelStart)
            .EditSelStart = Len(.EditText)
            Exit Sub
        End If

        If Col = col_��ʼ���� Or Col = col_�������� Then
            'ֻ����������
            strChr = Chr(KeyAscii)

            If InStr("0123456789-", strChr) = 0 Then KeyAscii = 0: Exit Sub
            If .EditSelStart < 10 And Len(.EditText) = .EditSelStart Then
                '����
                '���
                If Len(.EditText) = 0 And InStr("123", strChr) = 0 Then KeyAscii = 0: Exit Sub

                '�·�
                If (.EditSelStart = 4 Or .EditSelStart = 5) And InStr("01", strChr) = 0 Then KeyAscii = 0: Exit Sub
                If .EditSelStart = 6 Then
                    If (Val(Right(.EditText, 1)) = 0 And Val(strChr) = 0) Or (Val(Right(.EditText, 1)) = 1 And Val(strChr) > 2) Then
                        KeyAscii = 0: Exit Sub
                    End If
                End If
                '����
                If (.EditSelStart = 7 Or .EditSelStart = 8) And InStr("0123", strChr) = 0 Then KeyAscii = 0: Exit Sub
                If .EditSelStart = 9 Then
                    If (Val(Right(.EditText, 1)) = 0 And Val(strChr) = 0) Or (Val(Right(.EditText, 1)) = 3 And Val(strChr) > 1) Then
                        KeyAscii = 0: Exit Sub
                    End If
                End If
                '�Զ���ӷָ���
                If .EditSelStart = 4 Or .EditSelStart = 7 Then
                    .EditText = .EditText & "-"
                End If
            ElseIf .EditSelStart < Len(.EditText) And .EditSelLength = 0 And Len(.EditText) < 10 Then    '�м����
                KeyAscii = 0
                lngTmp = .EditSelStart

                If lngTmp = 4 Or lngTmp = 7 Then
                    .EditText = Mid(.EditText, 1, lngTmp) & "-" & strChr & Mid(.EditText, lngTmp + 1)
                    .EditSelStart = lngTmp + 2
                Else
                    .EditText = Mid(.EditText, 1, lngTmp) & strChr & Mid(.EditText, lngTmp + 1)
                    .EditSelStart = lngTmp + 1
                End If
                Exit Sub
            ElseIf Len(.EditText) >= 10 Or .EditSelStart < Len(.EditText) Then
                '�޸�
                strChr = Chr(KeyAscii)
                KeyAscii = 0

                If .EditSelStart <= 4 Then
                    '���
                    mlngNum = mlngNum + 1
                    If mlngNum = 1 And InStr("123", strChr) = 0 Then mlngNum = mlngNum - 1: Exit Sub
                    .EditText = Left(.EditText, mlngNum - 1) & strChr & Mid(.EditText, mlngNum + 1)
                    .EditSelStart = mlngNum
                    .EditSelLength = 4 - mlngNum
                    If mlngNum = 4 Then Call Vs_EditSelChange(5)
                ElseIf .EditSelStart >= 5 And .EditSelStart <= 7 Then
                    '�·�
                    mlngNum = mlngNum + 1
                    If mlngNum = 1 And InStr("01", strChr) = 0 Then mlngNum = mlngNum - 1: Exit Sub
                    If mlngNum = 2 Then
                        If Val(Mid(.EditText, 6, 1)) = 0 And Val(strChr) = 0 Then
                            mlngNum = mlngNum - 1: Exit Sub  '�·���С��01
                        ElseIf Val(Mid(.EditText, 6, 1)) = 1 And Val(strChr) > 2 Then
                            mlngNum = mlngNum - 1: Exit Sub     '�·����12
                        End If
                    End If
                    .EditText = Left(.EditText, mlngNum + 4) & strChr & Mid(.EditText, mlngNum + 6)
                    .EditSelStart = 5 + mlngNum
                    .EditSelLength = 2 - mlngNum
                    If mlngNum = 2 Then Call Vs_EditSelChange(8)
                Else
                    '����
                    mlngNum = mlngNum + 1
                    If mlngNum = 1 And InStr("0123", strChr) = 0 Then mlngNum = mlngNum - 1: Exit Sub
                    If mlngNum = 2 Then
                        If Val(Mid(.EditText, 9, 1)) = 0 And Val(strChr) = 0 Then
                            mlngNum = mlngNum - 1: Exit Sub  '������С��01
                        ElseIf Val(Mid(.EditText, 9, 1)) = 3 And Val(strChr) > 1 Then
                            mlngNum = mlngNum - 1: Exit Sub     '�������31
                        End If
                    End If
                    .EditText = Left(.EditText, mlngNum + 7) & strChr & Right(.EditText, 2 - mlngNum)
                    .EditSelStart = 8 + mlngNum
                    .EditSelLength = 2 - mlngNum
                    If mlngNum = 2 Then Call Vs_EditSelChange(4)
                End If

            End If
        End If
    End With
End Sub

Private Sub EnterNextCell()
    With vsSymptom
        If .Col >= col_�������� Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = col_֢״
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Sub vsSymptom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsSymptom
        If .Row >= .FixedRows And .Row <= .Rows - 2 Then
            '���ɾ����ť
            Set .Cell(flexcpPicture, .FixedRows, COL_����, .Rows - 1, COL_����) = Nothing
            '��ʾɾ����ť
            Set .Cell(flexcpPicture, .Row, COL_����) = imgButtonDel.Picture
        End If
        If .Col = col_֢״ Then
            .ToolTipText = "��F1������¼��֢״"
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub vsSymptom_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsSymptom
        If Col = col_֢״ Then
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        ElseIf Col = col_��ʼ���� Or Col = col_�������� Then
            tmrThis.Enabled = True
            .EditSelStart = 0
            .EditSelLength = 4
        End If
    End With
End Sub

Private Function ValidateDate(ByRef Row As Long, ByRef Col As Long) As Boolean
    Dim curDate As Date

    With vsSymptom    '���ڸ�ʽ���
        If Col = col_��ʼ���� Or Col = col_�������� Then
            If Not IsDate(.TextMatrix(Row, Col)) Then '��������ʾ
                Call mobjAir.OpenTransparentAirBubble(picBottom, "��������ڸ�ʽ����ȷ�����ڴ���", 1, 1, 80, &H99CCFF, vbRed, , 1, , , ����, True)
                mIntWaitTime = 3: ValidateDate = True
                Exit Function
            Else  '������ʾ
                If .TextMatrix(Row, Col) <> "" Then
                    curDate = zlDatabase.Currentdate
                    curDate = Format(curDate, "yyyy-mm-dd")
                    If CDate(.TextMatrix(Row, Col)) > curDate Then
                        Call mobjAir.OpenTransparentAirBubble(picBottom, "����������ڲ��ܴ��ڵ�ǰ���ڡ���ǰ���ڣ�" & curDate & "��", 1, 1, 80, &H99CCFF, vbRed, , 1, , , ����, True)
                        mIntWaitTime = 3: ValidateDate = True
                        Exit Function
                    End If
                    '��ʼ����<��������
                    If Col = col_�������� Then
                        If CDate(.TextMatrix(Row, col_��ʼ����)) > CDate(.TextMatrix(Row, Col)) Then
                            Call mobjAir.OpenTransparentAirBubble(picBottom, "��ʼ���ڲ��ܴ��ڽ�������", 1, 1, 80, &H99CCFF, vbRed, , 1, , , ����, True)
                            mIntWaitTime = 3: ValidateDate = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Sub LoadSyptoms()
'����:���ز���֢״
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim lngRow As Long

    strSQL = "Select ����,���,����,��ʼ����,��������,��¼�� From ����֢״��¼ Where ����id = [1] And ��ҳid = [2]"
    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    With vsSymptom
        .Redraw = flexRDNone
        .Rows = 2  'Ĭ����ʾһ��
        For i = 1 To rsTmp.RecordCount
            lngRow = .Rows - 1
            .RowData(lngRow) = rsTmp!���� & ""
            .TextMatrix(lngRow, col_֢״) = rsTmp!���� & ""
            .Cell(flexcpData, lngRow, col_֢״) = rsTmp!���� & ""
            .TextMatrix(lngRow, COL_���) = rsTmp!��� & ""
            .TextMatrix(lngRow, col_��ʼ����) = Format(rsTmp!��ʼ���� & "", "yyyy-mm-dd")
            .Cell(flexcpData, lngRow, col_��ʼ����) = Format(rsTmp!��ʼ���� & "", "yyyy-mm-dd")
            .TextMatrix(lngRow, col_��������) = Format(rsTmp!�������� & "", "yyyy-mm-dd")
            .Cell(flexcpData, lngRow, col_��������) = Format(rsTmp!�������� & "", "yyyy-mm-dd")
            .TextMatrix(lngRow, COL_ҽ��) = rsTmp!��¼�� & ""
            .Cell(flexcpData, lngRow, COL_ҽ��) = rsTmp!��¼�� & ""
            .TextMatrix(lngRow, COL_״̬) = "1"    '1-ԭʼ

            .Rows = lngRow + 2
            rsTmp.MoveNext
        Next
        .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter  '��Ԫ���������ж���
        .Redraw = flexRDDirect

    End With

    Exit Sub
errH:
    If ErrCenter() Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckCell() As Boolean
'����:���������ݲ���Ϊ��,������ڵ�Ԫ����ȷ�ԡ�
    Dim i As Long, j As Long
    If gbytPass <> TYT Then Exit Function
    With vsSymptom
        For i = .FixedRows To .Rows - 2
            For j = col_֢״ To COL_ҽ��
                If .TextMatrix(i, j) = "" Then
                    MsgBox "֢״������д������������д���������˳�", vbInformation + vbOKOnly, gstrSysName
                    '��λ��Ԫ��
                    .Row = i: .Col = j
                    .EditCell
                    CheckCell = True
                    Exit Function
                End If
                If j = col_��ʼ���� Or j = col_�������� Then
                    If ValidateDate(i, j) Then
                        .Row = i: .Col = j
                        .EditCell
                        CheckCell = True
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
End Function

Private Sub InitFormBorder()
    Dim i As Long
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2

    '�߿�����
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
End Sub

Private Sub Vs_EditSelChange(ByVal lngSelNum As Long)
'���û��л�����ʱ�򴥷�
    With vsSymptom
        If lngSelNum <= 4 Then
            .EditSelStart = 0
            .EditSelLength = 4
        ElseIf lngSelNum <= 7 Then
            .EditSelStart = 5
            .EditSelLength = 2
        ElseIf lngSelNum <= 10 Then
            .EditSelStart = 8
            .EditSelLength = 2
        End If
        mlngNum = 0
    End With
End Sub

Private Sub cboStatus_Click(Index As Integer)
    If Index = IX_���� Then
        If cboStatus(Index).Text = "����" Then
            txtDate.Enabled = True
            txtDate.BackColor = &H80000005
            If Me.Visible Then txtDate.SetFocus
        Else
            txtDate.Enabled = False
            txtDate.BackColor = &H8000000C
            txtDate.Text = "": txtDate.Tag = ""
        End If
    End If
    cboStatus(Index).Tag = cboStatus(Index).Text
End Sub

Private Sub cboStatus_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub txtDate_GotFocus()
    Call zlControl.TxtSelAll(txtDate)
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then
        cmdQuit.SetFocus
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtDate.Text <> "" Then
            txtDate.Text = zlStr.FullDate(txtDate.Text)
        End If
        zlCommFun.PressKey vbKeyTab
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    Dim datCurr As Date
    
    If Trim(txtDate.Text) <> "" Then
        If Not IsDate(txtDate) Then
            MsgBox "��������:" & txtDate.Text & "������Ч�����ڸ�ʽ:YYYY-MM-dd��", vbInformation + vbOKOnly, gstrSysName
            Cancel = True
        Else
            datCurr = zlDatabase.Currentdate
            If DateDiff("D", CDate(txtDate), datCurr) < 0 Then
                MsgBox "�������ڳ�����ǰʱ�䣡", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
            Else
                txtDate.Text = Format(txtDate.Text, "YYYY-MM-DD")
                txtDate.Tag = "��������|" & txtDate.Text
            End If
        End If
    Else
        MsgBox "�������ڲ���Ϊ�գ�", vbInformation + vbOKOnly, gstrSysName
        Cancel = True
    End If
    If Cancel Then
        Call txtDate_GotFocus
    End If
End Sub
